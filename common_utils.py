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


# ──────────────────────────────────────────────────────────────────────
# 行スキップ判定（色付きセルを除外）
# ──────────────────────────────────────────────────────────────────────
WHITE_CODES = {"00000000", "FFFFFFFF", None}   # XLSX で “白” とみなす RGB 値


def rows_to_skip_by_color(
    raw_bytes: bytes,
    sheet_name: str,
    target_col: int,
    first_data_row_excel: int = 8,
) -> Set[int]:
    """
    指定列 (0‑index) のセルが「白以外の背景色」または「白以外の文字色」
    になっている行番号（DataFrame index 相当）をセットで返す。

    Parameters
    ----------
    raw_bytes : bytes
        Dropbox から取得した XLSX 生データ
    sheet_name : str
        対象のシート名
    target_col : int
        チェックする列の 0‑index
    first_data_row_excel : int, default 8
        Excel の何行目から DataFrame 0 行目が始まるか（デフォルト: 行見出し 7 行 +1）

    Returns
    -------
    set[int]
        スキップすべき DataFrame 行 index
    """
    wb = load_workbook(io.BytesIO(raw_bytes), data_only=True)
    ws = wb[sheet_name]

    skip: Set[int] = set()
    # openpyxl は 1‑index 行番号。enumerate start=0 で DF index に合わせる
    for df_row, row in enumerate(
        ws.iter_rows(min_row=first_data_row_excel), start=0
    ):
        cell = row[target_col]

        # 背景色
        fill = cell.fill
        bg_rgb = getattr(fill.fgColor, "rgb", None)
        custom_bg = fill.patternType and bg_rgb not in WHITE_CODES

        # 文字色
        font = cell.font
        fg_rgb = getattr(font.color, "rgb", None)
        custom_fg = fg_rgb not in WHITE_CODES

        if custom_bg or custom_fg:
            skip.add(df_row)

    logging.debug("rows_to_skip_by_color → %s", skip)
    return skip


# ──────────────────────────────────────────────────────────────────────
# SMTP メール送信
# ──────────────────────────────────────────────────────────────────────
def send_email(subject: str, body: str):
    """
    TEXT メールを SMTP で送信。
    必要な環境変数:
        • SMTP_SERVER, SMTP_PORT(省略可), SMTP_USER, SMTP_PASSWORD
        • EMAIL_RECIPIENTS (カンマ区切り)
    """
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
