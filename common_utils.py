import os
import io
import datetime
import logging
import smtplib
from email.mime.text import MIMEText

import pandas as pd
import dropbox

logging.basicConfig(level=logging.INFO)

# ── Dropbox ───────────────────────────────────────────────────────
def get_dropbox_client() -> dropbox.Dropbox:
    """環境変数から直接トークンを読み込む"""
    return dropbox.Dropbox(
        app_key              = os.environ["DROPBOX_APP_KEY"],
        app_secret           = os.environ["DROPBOX_APP_SECRET"],
        oauth2_refresh_token = os.environ["DROPBOX_REFRESH_TOKEN"],
    )


def download_excel(path: str) -> bytes | None:
    try:
        _, res = get_dropbox_client().files_download(path)
        logging.info("✅  Dropbox から Excel を取得")
        return res.content
    except Exception as e:
        logging.error(f"❌ Dropbox ダウンロード失敗: {e}")
        return None

# ── メール送信 ────────────────────────────────────────────────────
def send_email(subject: str, body: str):
    smtp_server   = os.environ["SMTP_SERVER"]           # 例: smtp.gmail.com
    smtp_port     = int(os.environ.get("SMTP_PORT", 587))
    smtp_user     = os.environ["SMTP_USER"]             # 送信元アドレス
    smtp_password = os.environ["SMTP_PASSWORD"]         # アプリパスワード
    recipients    = os.environ["EMAIL_RECIPIENTS"].split(",")

    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = subject
    msg["From"]    = smtp_user
    msg["To"]      = ", ".join(recipients)

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as s:
            s.starttls()
            s.login(smtp_user, smtp_password)
            s.send_message(msg)
        logging.info("✅  メール送信完了 → %s", recipients)
    except Exception as e:
        logging.error(f"❌ メール送信エラー: {e}")
        raise

# ── アラート判定（共通）───────────────────────────────────────────
def should_alert(due: datetime.date, alert_days: int) -> bool:
    """
    通知判定:
      • 指定日前 (alert_days) のみ通知
      • 1～2日遅延のものは⚠️通知
      • 3日以上遅延したものは無視
    """
    today = datetime.date.today()
    delta = (due - today).days  # 未来 = 正、過去 = 負

    if delta == alert_days:
        return True   # 指定日前ぴったり
    if -2 <= delta < 0:
        return True   # 遅延1～2日
    return False
