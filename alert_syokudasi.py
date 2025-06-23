import os, io, datetime, logging, pandas as pd
from common_utils import (
    download_excel, send_email, should_alert, rows_to_skip_by_color,
)

RECIPIENT_KEY = "EMAIL_SEISAN"
ALERT_DAYS, ALERT_NAME = 3, "生産職出し納期"
FILE_PATH  = "/生産部/工場予定表(2025)_新レイアウト.xlsx"
SHEET_NAME = "25AW"

# 0‑index
COL_BRAND, COL_PERSON, COL_ITEM = 3, 2, 4
COL_CHECK, COL_DUE = 5, 15  # F, P


def fetch_items() -> list[dict]:
    raw = download_excel(FILE_PATH)
    if not raw:
        return []

    df = pd.read_excel(io.BytesIO(raw), sheet_name=SHEET_NAME, header=None)
    df = df.iloc[7:, [COL_BRAND, COL_PERSON, COL_ITEM, COL_CHECK, COL_DUE]]
    df.columns = ["brand", "person", "item", "check", "due"]
    df["due"] = pd.to_datetime(df["due"], errors="coerce").dt.date

    # ----- 色付きセルを除外 -----------------------------------
    skip = rows_to_skip_by_color(raw, SHEET_NAME, COL_DUE)
    df = df.reset_index(drop=True).loc[~df.index.isin(skip)]
    # -----------------------------------------------------------

    truthy = {"true", "1", "yes", "y", "✓"}
    df["priority"] = (
        df["check"].astype(str).str.strip().str.lower().isin(truthy).astype(int)
    )
    df.sort_values(["item", "priority"], ascending=[True, False], inplace=True)
    df = df.drop_duplicates(subset="item", keep="first")

    today = datetime.date.today()
    rows = []
    for _, r in df.dropna(subset=["due"]).iterrows():
        if should_alert(r["due"], ALERT_DAYS):
            rows.append(
                {
                    "brand": str(r["brand"]).strip() or "不明",
                    "person": str(r["person"]).strip() or "不明",
                    "item": str(r["item"]).strip() or "不明",
                    "due": r["due"],
                    "delta": (r["due"] - today).days,
                }
            )
    return rows


def build_body(rows: list[dict]) -> str:
    header = [f"【{ALERT_NAME}アラート】", ""]

    if not rows:
        return "\n".join(header + ["該当する品番はありません。"])

    tree: dict[str, dict[str, list[str]]] = {}
    for r in rows:
        d = r["delta"]
        prefix = "⚠️ " if d < 0 else "• "
        when = f"出荷日超過 {abs(d)} 日" if d < 0 else f"出荷まで {d} 日"
        line = f"{prefix}品番: {r['item']} — {when} ({r['due']:%Y-%m-%d})"
        tree.setdefault(r["person"], {}).setdefault(r["brand"], []).append(line)

    body = header[:]
    for person, brands in tree.items():
        body.append(f"【担当: {person}】")
        for brand, items in brands.items():
            body.append(f"*〔{brand}〕*")
            body.extend(items)
            body.append("")
        body.append("")
    return "\n".join(body)


def run():
    if RECIPIENT_KEY in os.environ:
        os.environ["EMAIL_RECIPIENTS"] = os.environ[RECIPIENT_KEY]

    rows = fetch_items()
    if not rows:
        logging.info("該当する品番がないため、メールを送信しません。")
        return

    body = build_body(rows)
    send_email(f"[{ALERT_NAME}アラート]", body)


if __name__ == "__main__":
    run()