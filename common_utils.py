# common_utils.py
# ---------------------------------------------------------------------------
# 共有ユーティリティ：
#   • Dropbox から Excel を取得（download_excel）
#   • 行スキップ判定：セルの背景色 or 文字色が白以外なら除外（rows_to_skip_by_color）
#   • アラート判定（should_alert）
#   • SMTP 経由でメール送信（send_email）
# ---------------------------------------------------------------------------
import os
import io
import datetime
import logging
import smtplib
from email.mime.text import MIMEText
from typing import Set

import dropbox
from openpyxl import load_workbook

IS_DRY_RUN = os.getenv("DRY_RUN", "0") == "1"   # ★追加

logging.basicConfig(level=logging.INFO)

# ──────────────────────────────────────────────────────────────────────
# Dropbox
# ──────────────────────────────────────────────────────────────────────
def get_dropbox_client() -> dropbox.Dropbox:
    """環境変数からリフレッシュトークンを読み込んで Dropbox クライアントを返す"""
    return dropbox.Dropbox(
        app_key              = os.environ["DROPBOX_APP_KEY"],
        app_secret           = os.environ["DROPBOX_APP_SECRET"],
        oauth2_refresh_token = os.environ["DROPBOX_REFRESH_TOKEN"],
        timeout              = int(os.getenv("DROPBOX_TIMEOUT", 900)),  # ★追加
    )


def download_excel(path: str) -> bytes | None:
    """
    Dropbox から指定パスのファイルをダウンロードして raw bytes を返す。
    失敗したら None を返す。
    """
    try:
        _, res = get_dropbox_client().files_download(path)
        logging.info("✅  Dropbox から Excel を取得: %s", path)
        return res.content
    except Exception as e:
        logging.error("❌ Dropbox ダウンロード失敗: %s", e)
        return None


# ------------------------------------------------------------------
#  ❖ これだけで OK
#     - 背景色が #F7DFDF（ARGB でも RGB でも可）のセルを持つ行だけ除外
#     - 文字色は判定しない
# ------------------------------------------------------------------
SKIP_BG_HEX = {"f7dfdf"}        # 除外したい 6 桁 RGB を列挙

def _is_skip_color(argb: str | None) -> bool:
    """openpyxl の ARGB 8桁 or RGB 6桁を受け取り、対象色なら True"""
    if argb is None:
        return False
    return argb[-6:].upper() in SKIP_BG_HEX      # 下 6 桁で比較

def rows_to_skip_by_color(raw_bytes: bytes, sheet_name: str,
                          target_col: int,
                          first_data_row_excel: int = 8) -> set[int]:
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(raw_bytes), data_only=True)
    ws = wb[sheet_name]

    skip = set()
    for df_row_idx, row in enumerate(ws.iter_rows(min_row=first_data_row_excel), 0):
        cell = row[target_col]
        bg_rgb = getattr(cell.fill.fgColor, "rgb", None)
        if _is_skip_color(bg_rgb):
            skip.add(df_row_idx)
    return skip


# ──────────────────────────────────────────────────────────────────────
# SMTP メール送信
# ──────────────────────────────────────────────────────────────────────
def send_email(subject: str, body: str):
    """
    TEXT メールを SMTP で送信。
    DRY_RUN=1 の場合はログ出力だけでスキップ。
    """

    if IS_DRY_RUN:                                 # ★追加
        logging.info("🟡 DRY‑RUN → メール送信スキップ: %s", subject)
        return

    smtp_server   = os.environ["SMTP_SERVER"]
    smtp_port     = int(os.environ.get("SMTP_PORT", 587))
    smtp_user     = os.environ["SMTP_USER"]
    smtp_password = os.environ["SMTP_PASSWORD"]
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
        logging.info("✅ メール送信完了 → %s", recipients)
    except Exception as e:
        logging.error("❌ メール送信エラー: %s", e)
        raise



# ──────────────────────────────────────────────────────────────────────
# アラート判定（共通ロジック）
# ──────────────────────────────────────────────────────────────────────
def should_alert(due: datetime.date, alert_days: int) -> bool:
    """
    通知判定:
      • 指定日前 (alert_days) のみ通知
      • 1～2日遅延のものは ⚠️ 通知
      • 3日以上遅延したものは無視
    """
    today = datetime.date.today()
    delta = (due - today).days  # 未来 = 正、過去 = 負

    if delta == alert_days:
        return True   # 指定日前ぴったり
    if -2 <= delta < 0:
        return True   # 遅延 1～2 日
    return False
