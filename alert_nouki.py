import io, datetime, logging, pandas as pd
from common_utils import download_excel, send_email, should_alert

ALERT_DAYS  = 7
ALERT_NAME  = "量産納期"
FILE_PATH   = "/生産部/工場予定表(2025)_新レイアウト.xlsx"
SHEET_NAME  = "25AW"
COL_BRAND   = 3   # D列
COL_PERSON  = 2   # C列
COL_ITEM    = 4   # E列
COL_DUE     = 22  # W列

def fetch_items() -> list[dict]:
    raw = download_excel(FILE_PATH)
    if not raw:
        return []
    df = pd.read_excel(io.BytesIO(raw), sheet_name=SHEET_NAME, header=None)
    df = df.iloc[7:, [COL_BRAND, COL_PERSON, COL_ITEM, COL_DUE]]
    df.columns = ["brand", "person", "item", "due"]
    df["due"]  = pd.to_datetime(df["due"], errors="coerce").dt.date

    today = datetime.date.today()
    return [
        {"brand":str(r["brand"]).strip() or "不明",
         "person":str(r["person"]).strip() or "不明",
         "item":str(r["item"]).strip() or "不明",
         "due":r["due"],
         "delta":(r["due"]-today).days}
        for _, r in df.dropna(subset=["due"]).iterrows()
        if should_alert(r["due"], ALERT_DAYS)
    ]

def build_body(rows:list[dict]) -> str:
    body=["【"+ALERT_NAME+"アラート】",""]
    if not rows:
        return "\n".join(body+["該当する品番はありません。"])

    tree:dict[str,dict[str,list[str]]]={}
    for r in rows:
        delta=r["delta"]
        prefix="⚠️ " if delta<0 else "• "
        when  =f"出荷日超過 {abs(delta)} 日" if delta<0 else f"出荷まで {delta} 日"
        line  =f"{prefix}品番: {r['item']} — {when} ({r['due']:%Y-%m-%d})"
        tree.setdefault(r["person"],{}).setdefault(r["brand"],[]).append(line)

    for person,brands in tree.items():
        body.append(f"【担当: {person}】")
        for brand,items in brands.items():
            body.append(f"*〔{brand}〕*")
            body.extend(items)
            body.append("")
        body.append("")
    return "\n".join(body)

def run():
    rows=fetch_items()
    send_email(f"[{ALERT_NAME}アラート]",build_body(rows))

if __name__=="__main__":
    run()
