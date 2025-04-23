import io, datetime, logging, pandas as pd
from common_utils import download_excel, send_email, should_alert

ALERT_DAYS, ALERT_NAME = 7, "縫製納期"
FILE_PATH="/生産部/工場予定表(2025)_新レイアウト.xlsx"; SHEET_NAME="25AW"
COL_BRAND,COL_PERSON,COL_ITEM,COL_DUE = 3,2,4,17   # D,C,E,R

def fetch_items()->list[dict]:
    raw=download_excel(FILE_PATH);  today=datetime.date.today()
    if not raw: return []
    df=pd.read_excel(io.BytesIO(raw),sheet_name=SHEET_NAME,header=None)
    df=df.iloc[7:,[COL_BRAND,COL_PERSON,COL_ITEM,COL_DUE]]
    df.columns=["brand","person","item","due"]; df["due"]=pd.to_datetime(df["due"],errors="coerce").dt.date
    return[{"brand":str(r.brand).strip()or"不明","person":str(r.person).strip()or"不明","item":str(r.item).strip()or"不明","due":r.due,"delta":(r.due-today).days}
           for _,r in df.dropna(subset=["due"]).iterrows() if should_alert(r.due,ALERT_DAYS)]

def build_body(rows:list[dict])->str:
    body=["【"+ALERT_NAME+"アラート】",""];  tree={}
    if not rows: return"\n".join(body+["該当する品番はありません。"])
    for r in rows:
        d,prefix=str(abs(r["delta"])),"⚠️ " if r["delta"]<0 else "• "
        when=f"出荷日超過 {d} 日" if r["delta"]<0 else f"出荷まで {d} 日"
        line=f"{prefix}品番: {r['item']} — {when} ({r['due']:%Y-%m-%d})"
        tree.setdefault(r["person"],{}).setdefault(r["brand"],[]).append(line)
    for person,brands in tree.items():
        body.append(f"【担当: {person}】")
        for brand,items in brands.items():
            body.append(f"*〔{brand}〕*"); body.extend(items); body.append("")
        body.append("")
    return"\n".join(body)

def run():
    send_email(f"[{ALERT_NAME}アラート]",build_body(fetch_items()))

if __name__=="__main__": run()
