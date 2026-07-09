# -*- coding: utf-8 -*-
"""
TDnet適時開示監視スクリプト
監視リスト(watchlist.csv)の銘柄に適時開示が出たら、
好材料/悪材料を判定してWindows通知+Gmailで知らせます。

使い方:
  python kaiji_monitor.py          … 常駐監視(start_monitor.batから起動)
  python kaiji_monitor.py --test   … 通知テスト(Windows通知とメールを試験送信)
  python kaiji_monitor.py --once   … 1回だけ巡回して終了(クラウド実行・動作確認用)

外部ライブラリ不要(Python標準機能のみで動作)。
データ源: Yanoshin TDnet Web API (https://webapi.yanoshin.jp/tdnet/)
"""
import argparse
import configparser
import csv
import datetime
import io
import json
import os
import smtplib
import subprocess
import sys
import time
import traceback
import urllib.request
from email.header import Header
from email.mime.text import MIMEText

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.ini")
WATCHLIST_PATH = os.path.join(BASE_DIR, "watchlist.csv")
SEEN_PATH = os.path.join(BASE_DIR, "seen_ids.txt")
LOG_PATH = os.path.join(BASE_DIR, "alerts_log.csv")
ERROR_PATH = os.path.join(BASE_DIR, "errors.log")
API_URL = "https://webapi.yanoshin.jp/webapi/tdnet/list/recent.json?limit={limit}"

# コンソールの文字化け対策
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


# ---------- 設定・監視リスト読み込み ----------

def read_text_flexible(path):
    """UTF-8(BOM付き含む)→Shift-JISの順に試して読む(Excel保存形式の違いを吸収)"""
    for enc in ("utf-8-sig", "cp932"):
        try:
            with open(path, encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    raise RuntimeError(f"{path} の文字コードを判別できません")


def load_config():
    cfg = configparser.ConfigParser()
    cfg.read_string(read_text_flexible(CONFIG_PATH))
    return cfg


def parse_keywords(cfg, key):
    raw = cfg.get("keywords", key, fallback="")
    return [w.strip() for w in raw.split(",") if w.strip()]


def load_watchlist():
    """watchlist.csv -> {証券コード4桁: (銘柄名, セクター)}"""
    watch = {}
    with io.StringIO(read_text_flexible(WATCHLIST_PATH)) as f:
        reader = csv.reader(f)
        next(reader, None)  # ヘッダ行を飛ばす
        for row in reader:
            if len(row) >= 2 and row[0].strip():
                code = row[0].strip()
                name = row[1].strip()
                sector = row[2].strip() if len(row) >= 3 else ""
                watch[code] = (name, sector)
    return watch


# ---------- 既知IDの管理(同じ開示を二度通知しない) ----------

def load_seen():
    if not os.path.exists(SEEN_PATH):
        return set()
    with open(SEEN_PATH, encoding="utf-8") as f:
        return set(line.strip() for line in f if line.strip())


def save_seen(seen):
    # 肥大化防止のため直近2万件だけ保持
    ids = list(seen)[-20000:]
    with open(SEEN_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(ids))


# ---------- データ取得 ----------

def fetch_recent(limit):
    """TDnet開示の最新一覧を取得して[{...}, ...]で返す"""
    url = API_URL.format(limit=limit)
    req = urllib.request.Request(
        url, headers={"User-Agent": "kaiji-monitor/1.0 (personal-use)"}
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        data = json.loads(r.read().decode("utf-8"))
    items = []
    for wrapper in data.get("items", []):
        item = wrapper.get("Tdnet", {})
        if item:
            items.append(item)
    return items


# ---------- 判定 ----------

def classify(title, neg_kw, pos_kw):
    for k in neg_kw:
        if k in title:
            return "悪材料?"
    for k in pos_kw:
        if k in title:
            return "好材料?"
    return "その他"


def is_muted(title, mute_kw):
    return any(k in title for k in mute_kw)


# ---------- 通知 ----------

def toast(title, body):
    """Windowsのトースト通知(PowerShell経由・追加ソフト不要)"""
    if os.name != "nt":
        return  # クラウド(Linux)実行時はスキップ
    ps = r"""
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null
$xml = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
$texts = $xml.GetElementsByTagName("text")
$texts.Item(0).AppendChild($xml.CreateTextNode($env:TOAST_TITLE)) | Out-Null
$texts.Item(1).AppendChild($xml.CreateTextNode($env:TOAST_BODY)) | Out-Null
$toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("開示監視").Show($toast)
"""
    env = os.environ.copy()
    env["TOAST_TITLE"] = title[:120]
    env["TOAST_BODY"] = body[:200]
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
        env=env, capture_output=True, timeout=30,
    )


def send_mail(cfg, subject, body):
    # クラウド実行時は環境変数(GitHub Secrets)が優先される
    addr = os.environ.get("GMAIL_ADDRESS") or cfg.get("mail", "gmail_address", fallback="")
    pw = os.environ.get("GMAIL_APP_PASSWORD") or cfg.get("mail", "app_password", fallback="")
    to = os.environ.get("MAIL_TO") or cfg.get("mail", "mail_to", fallback=addr)
    addr, pw, to = addr.strip(), pw.replace(" ", "").strip(), to.strip()
    if not addr or not pw or "ここに" in pw:
        print("  [メール] アプリパスワード未設定のため送信をスキップ")
        return
    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = addr
    msg["To"] = to
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=30) as s:
        s.login(addr, pw)
        s.send_message(msg)


def notify(cfg, judge, code, name, sector, pubdate, title, pdf_url):
    head = f"【{judge}】{code} {name}"
    body = f"{pubdate}\n{title}\n({sector})"
    print(f"◆通知: {head} | {title}")
    if cfg.getboolean("notify", "windows_toast", fallback=True):
        try:
            toast(head, title)
        except Exception as e:
            print(f"  [Windows通知エラー] {e}")
    if cfg.getboolean("notify", "email", fallback=False):
        try:
            send_mail(cfg, head, body + "\n\nPDF: " + pdf_url)
        except Exception as e:
            print(f"  [メール送信エラー] {e}")


# ---------- 記録 ----------

def append_log(row):
    """alerts_log.csv に追記(Excelでそのまま開ける形式)"""
    new_file = not os.path.exists(LOG_PATH)
    with open(LOG_PATH, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        if new_file:
            w.writerow(["記録日時", "開示日時", "コード", "銘柄名", "セクター",
                        "判定", "タイトル", "PDFリンク", "通知"])
        w.writerow(row)


def log_error(e):
    with open(ERROR_PATH, "a", encoding="utf-8") as f:
        f.write(f"----- {datetime.datetime.now()} -----\n")
        f.write(traceback.format_exc() + "\n")
    print(f"  [エラー] {e} (詳細は errors.log)")


# ---------- 巡回1回分の処理 ----------

def check_once(cfg, watch, seen, first_run):
    neg_kw = parse_keywords(cfg, "negative")
    pos_kw = parse_keywords(cfg, "positive")
    mute_kw = parse_keywords(cfg, "mute")
    notify_all = cfg.getboolean("notify", "notify_all", fallback=True)
    limit = cfg.getint("monitor", "fetch_limit", fallback=500)

    items = fetch_recent(limit)
    hits = 0
    for item in items:
        item_id = str(item.get("id", ""))
        if not item_id or item_id in seen:
            continue
        seen.add(item_id)
        code5 = str(item.get("company_code", ""))
        code4 = code5[:4]
        if code4 not in watch:
            continue

        name, sector = watch[code4]
        title = item.get("title", "")
        pubdate = item.get("pubdate", "")
        pdf_url = item.get("document_url", "") or ""
        judge = classify(title, neg_kw, pos_kw)
        muted = is_muted(title, mute_kw)

        if first_run:
            # 初回起動時は過去分を通知せず既読化のみ
            continue

        do_notify = (not muted) and (notify_all or judge != "その他")
        append_log([
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            pubdate, code4, name, sector, judge, title, pdf_url,
            "通知" if do_notify else "記録のみ",
        ])
        hits += 1
        if do_notify:
            notify(cfg, judge, code4, name, sector, pubdate, title, pdf_url)

    save_seen(seen)
    return hits


# ---------- 監視時間帯の判定 ----------

def in_active_window(cfg, now):
    if now.weekday() >= 5 and not cfg.getboolean("monitor", "weekend", fallback=False):
        return False
    start = cfg.get("monitor", "active_start", fallback="07:00")
    end = cfg.get("monitor", "active_end", fallback="20:00")
    t = now.strftime("%H:%M")
    return start <= t <= end


# ---------- メイン ----------

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--test", action="store_true", help="通知テストのみ実行")
    parser.add_argument("--once", action="store_true", help="1回だけ巡回して終了")
    args = parser.parse_args()

    cfg = load_config()
    watch = load_watchlist()
    print("=" * 60)
    print(" TDnet開示監視スクリプト")
    print(f" 監視銘柄: {len(watch)}銘柄 / 巡回間隔: "
          f"{cfg.getint('monitor', 'interval_sec', fallback=300)}秒")
    print(" 停止するにはこのウィンドウを閉じてください")
    print("=" * 60)

    if args.test:
        print("通知テストを実行します…")
        try:
            toast("【テスト】開示監視", "Windows通知は正常に動作しています")
            print("  Windows通知: OK(画面右下を確認してください)")
        except Exception as e:
            print(f"  Windows通知: 失敗 {e}")
        try:
            send_mail(cfg, "【テスト】開示監視",
                      "メール通知は正常に動作しています。")
            print("  メール: 送信しました(受信箱を確認してください)")
        except Exception as e:
            print(f"  メール: 失敗 {e}")
            print("  → config.ini の gmail_address / app_password を確認してください")
        return

    seen = load_seen()
    first_run = len(seen) == 0
    if first_run:
        print("初回起動: 現時点までの開示を既読として登録します(通知は次の新着から)")

    interval = cfg.getint("monitor", "interval_sec", fallback=300)
    interval = max(interval, 180)  # 無料APIへの配慮で下限3分

    while True:
        now = datetime.datetime.now()
        try:
            if args.once or in_active_window(cfg, now):
                hits = check_once(cfg, watch, seen, first_run)
                first_run = False
                print(f"[{now.strftime('%m/%d %H:%M:%S')}] 巡回完了 "
                      f"(監視銘柄の新着: {hits}件)")
            else:
                print(f"[{now.strftime('%m/%d %H:%M:%S')}] 監視時間外のため待機中")
        except Exception as e:
            log_error(e)
        if args.once:
            break
        time.sleep(interval)


if __name__ == "__main__":
    main()
