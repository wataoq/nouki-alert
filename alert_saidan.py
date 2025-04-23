# alert_saidan.py  （裁断 7 日前）
import io
import datetime
import logging
import pandas as pd
from common_utils import download_excel, send_email, should_alert

# ========= 個別設定 =========
ALERT_DAYS = 7        # 裁断 7 日前通知
ALERT_NAME = "裁断納期"
FILE_PATH  = "/生産部/工場予定表(2025)_新レイアウト.xlsx"
SHEET_NAME = "25AW"
# 0-index の列番号
COL_BRAND  = 3   # D列: ブランド
COL_PERSON = 2   # C列: 担当者名
COL_ITEM   = 4   # E列: 品番
COL_DUE    = 16  # Q列: 裁断日
# ============================


def fetch_items() -> list[dict]:
    raw = download_excel(FILE_PATH)
    if not raw:
        return []

    df = pd.read_excel(io.BytesIO(raw), sheet_name=SHEET_NAME, header=None)
    df = df.iloc[7:, [COL_BRAND, COL_PERSON, COL_ITEM, COL_DUE]]
    df.columns = ["brand", "person", "item", "due"]
    df["due"] = pd.to_datetime(df["due"], errors="coerce").dt.date

    today = datetime.date.today()
    rows = []
    for _, r in df.dropna(subset=["due"]).iterrows():
        if should_alert(r["due"], ALERT_DAYS):
            rows.append({
                "brand":  str(r["brand"]).strip() or "不明",
                "person": str(r["person"]).strip() or "不明",
                "item":   str(r["item"]).strip() or "不明",
                "due":    r["due"],
                "delta":  (r["due"] - today).days,
            })
    return rows


def build_body(rows: list[dict]) -> str:
    # ヘッダー: アラート種別のみ
    body_lines = [f"【{ALERT_NAME}アラート】", ""]

    if not rows:
        return "\n".join(body_lines + ["該当する品番はありません。"])

    # 担当者→ブランド→アイテム のネスト
    data: dict[str, dict[str, list[str]]] = {}
    for r in rows:
        person = r["person"]
        brand = r["brand"]
        d = r["delta"]
        due_str = r["due"].strftime("%Y-%m-%d")
        prefix = "⚠️ " if d < 0 else "• "
        when = f"出荷日超過 {abs(d)} 日" if d < 0 else f"出荷まで {d} 日"
        line = f"{prefix}品番: {r['item']} — {when} ({due_str})"

        data.setdefault(person, {}).setdefault(brand, []).append(line)

    # 本文組み立て
    for person, brands in data.items():
        body_lines.append(f"【担当: {person}】")
        for brand, items in brands.items():
            body_lines.append(f"*〔{brand}〕*")
            body_lines.extend(items)
            body_lines.append("")  # ブランド毎に空行
        body_lines.append("")      # 担当者毎に空行

    return "\n".join(body_lines)


def run():
    rows = fetch_items()
    body = build_body(rows)
    send_email(f"[{ALERT_NAME}アラート]", body)
    return body  # ログ確認用


if __name__ == "__main__":
    run()