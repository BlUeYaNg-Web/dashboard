"""
每週日 13:05 台灣時間 (05:05 UTC) 由 GitHub Actions 執行：
1. 從 dualtaipei.datazen.info 抓取 11 個建案（主要來源）
2. 從 Google Sheet 補充土城區資料（交叉比對，含和耀美家雅居）
3. 本週新增判斷：以「資料來源是否本週才出現」為準，與交易日期無關
   - 上週 unit_keys 有記錄的戶 → 不算新增
   - 本週才出現的戶（即使成交在幾個月前）→ 算本週新增 1
4. 更新 results/pptx_update.json（含 unit_keys，供下週比對）
"""

import json, re, sys, urllib.request, urllib.parse
from collections import defaultdict
from datetime import datetime, timezone
from html.parser import HTMLParser
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

RESULTS_DIR = Path(__file__).parent.parent / "results"
OUTPUT_PATH = RESULTS_DIR / "pptx_update.json"
BASE_URL = "https://dualtaipei.datazen.info"
GS_URL = (
    "https://docs.google.com/spreadsheets/d/e/"
    "2PACX-1vRvMScWNP5idpnmDApHAQvZIxMixZC4Y5-dYA1OjTf5sltG-83C_jF5EHKb2O3gcYr7gglpam6RWFai"
    "/pubhtml/sheet?headers=false&gid=1957863253"
)
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "application/json, */*",
    "Referer": "https://dualtaipei.datazen.info/",
    "Origin": "https://dualtaipei.datazen.info",
}
CASES = [
    ("丰藏","家泰嘉潤丰藏","newtaipei"),("台信琢岳","台信琢岳","newtaipei"),
    ("紅布朗花園","紅布朗花園","newtaipei"),("寶佳淳青","寶佳淳青","newtaipei"),
    ("松陽馥麗","松陽馥麗","newtaipei"),("新濠岳","新濠岳","newtaipei"),
    ("朗沐","朗沐","newtaipei"),("合康雙匯","合康雙匯","newtaipei"),
    ("迴東騰","迴東騰","newtaipei"),("天好運3","天好運3","newtaipei"),
    ("若水秧翠","若水秧翠","newtaipei"),
]
GS_CASE_MAP = {
    "家泰嘉潤丰藏":"丰藏","台信琢岳":"台信琢岳","紅布朗花園":"紅布朗花園",
    "寶佳淳青":"寶佳淳青","松陽馥麗":"松陽馥麗","新濠岳":"新濠岳",
    "朗沐":"朗沐","合康雙匯":"合康雙匯","迴東騰":"迴東騰",
    "天好運3":"天好運3","若水秧翠":"若水秧翠","和耀美家 雅居":"和耀美家雅居",
}
ALL_DISPLAY_NAMES = [
    "丰藏","台信琢岳","紅布朗花園","寶佳淳青","松陽馥麗","新濠岳",
    "朗沐","合康雙匯","迴東騰","天好運3","若水秧翠","和耀美家雅居",
]

def fetch_json(url):
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.loads(r.read().decode("utf-8"))

def roc_to_iso(s):
    s = str(s).strip()
    if len(s)==7 and s.isdigit():
        return f"{int(s[:3])+1911}-{int(s[3:5]):02d}-{int(s[5:7]):02d}T00:00:00+00:00"
    return None

def parse_price_float(p):
    m = re.search(r"([\d.]+)", str(p))
    return float(m.group(1)) if m else None

def normalize_unit_key(unit_str):
    m = re.match(r"^([A-Za-z]+\d*)-0*(\d+)F?$", unit_str.strip())
    return f"{m.group(1).upper()}-{int(m.group(2))}" if m else unit_str.strip()

def gs_unit_to_std(s):
    s = s.strip()
    if "/" not in s: return s
    parts = s.split("/")
    if len(parts)<2: return s
    u, f = parts[0].strip(), parts[1].strip()
    if re.match(r"^[A-Za-z]+\d*-0*\d+F?$", f): return f
    fn = re.sub(r"^0*(\d+)F?$", r"\1F", f, flags=re.IGNORECASE)
    return f"{u}-{fn}" if u else fn

def parse_dt(t):
    try: return datetime.fromisoformat(t.get("date","").replace("Z","+00:00"))
    except: return datetime.min.replace(tzinfo=timezone.utc)

class TableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows=[]; self.current_row=[]; self.in_td=False; self.current_text=""
    def handle_starttag(self,tag,attrs):
        if tag in("td","th"): self.in_td=True; self.current_text=""
        elif tag=="tr": self.current_row=[]
    def handle_endtag(self,tag):
        if tag in("td","th"): self.current_row.append(self.current_text.strip()); self.in_td=False
        elif tag=="tr":
            if any(c.strip() for c in self.current_row): self.rows.append(self.current_row[:])
    def handle_data(self,data):
        if self.in_td: self.current_text+=data

def fetch_dualtaipei():
    print("=== dualtaipei 抓取 ===")
    result={}
    for key,full_name,city in CASES:
        print(f"  {key}...",end="",flush=True)
        encoded=urllib.parse.quote(full_name)
        base=f"{BASE_URL}/api/{city}/project/{encoded}"
        try:
            monthly=fetch_json(f"{base}?type=monthly").get("monthly",[])
            td=fetch_json(f"{base}?type=transactions")
            txns=td.get("transactions",[])
            total=td.get("pagination",{}).get("total",len(txns))
            offset=len(txns)
            while offset<total:
                more=fetch_json(f"{base}?type=transactions&limit=50&offset={offset}")
                b=more.get("transactions",[])
                if not b: break
                txns.extend(b); offset+=len(b)
            result[key]={"monthly":monthly,"transactions":txns}
            print(f" OK {len(txns)}筆")
        except Exception as e:
            print(f" FAIL:{e}"); result[key]={"monthly":[],"transactions":[]}
    return result

def fetch_google_sheet():
    print("\n=== Google Sheet 抓取 ===")
    h={**HEADERS,"Accept":"text/html,*/*"}
    req=urllib.request.Request(GS_URL,headers=h)
    with urllib.request.urlopen(req,timeout=30) as r:
        html=r.read().decode("utf-8",errors="replace")
    p=TableParser(); p.feed(html)
    ct=defaultdict(list)
    for row in p.rows:
        if len(row)<10 or not any("土城" in c for c in row): continue
        ck=None; nc=None
        for ci in [4,3]:
            if ci>=len(row): continue
            cv=row[ci].strip()
            for gn,k in GS_CASE_MAP.items():
                if cv==gn.strip(): ck=k; nc=ci; break
            if ck: break
        if not ck: continue
        off=nc-3
        iso=roc_to_iso(row[1].strip())
        unit=gs_unit_to_std(row[4+off].strip())
        price=parse_price_float(row[5+off].strip())
        if not iso or not unit or price is None: continue
        ct[ck].append({"date":iso,"unit":unit,"unitPrice":price,"_source":"google_sheet"})
    total=sum(len(v) for v in ct.values())
    print(f"  OK {total}筆（{len(ct)}個案名）")
    return ct

def cross_validate(dt_data,gs_data):
    print("\n=== 交叉比對 ===")
    merged={}
    for key,di in dt_data.items():
        txns=list(di["transactions"])
        ex={normalize_unit_key(t.get("unit","")) for t in txns}
        gs=gs_data.get(key,[])
        added=0
        for t in gs:
            uk=normalize_unit_key(t["unit"])
            if uk not in ex: txns.append(t); ex.add(uk); added+=1
        print(f"  [{key}] GS補充{added}筆" if added else f"  [{key}] 一致")
        merged[key]={"monthly":di["monthly"],"transactions":txns}
    for key,gs in gs_data.items():
        if key not in merged:
            merged[key]={"monthly":[],"transactions":list(gs)}
            print(f"  [{key}] 僅GS，{len(gs)}筆")
    return merged

def load_prev_unit_keys():
    try:
        prev=json.loads(OUTPUT_PATH.read_text(encoding="utf-8"))
        result={c["name"]:set(c.get("unit_keys",[])) for c in prev.get("cases",[])}
        print(f"上週unit_keys載入（{len(result)}個案名）")
        return result
    except:
        print("無上週unit_keys，本週新增全以差值計")
        return {}

def calc_stats(display,txns,prev_map):
    if not txns: return 0,0,None,None,None,[]
    current_total=len(txns)
    prev_units=prev_map.get(display,set())
    uk_list=sorted({normalize_unit_key(t.get("unit","")) for t in txns})
    # 本週新增 = 這週有、上週沒有（不論交易日期）
    weekly_new=sum(1 for uk in uk_list if uk not in prev_units)
    s=sorted(txns,key=parse_dt)
    latest=s[-1]
    lp=round(latest.get("unitPrice") or 0,2) or None
    la=round(latest.get("area") or 0,1) or None
    now=datetime.now(timezone.utc)
    wd=(now-parse_dt(latest)).days//7 if weekly_new==0 else 0
    return weekly_new,current_total,lp,la,wd,uk_list

def main():
    ns=datetime.now().strftime("%Y/%m/%d %H:%M")
    print(f"[{ns}] 開始（GitHub Actions 雲端模式）...")
    prev_map=load_prev_unit_keys()
    dt_data=fetch_dualtaipei()
    try: gs_data=fetch_google_sheet()
    except Exception as e: print(f"  GS失敗:{e}"); gs_data={}
    merged=cross_validate(dt_data,gs_data)
    print("\n=== 統計結果 ===")
    cases=[]
    for d in ALL_DISPLAY_NAMES:
        info=merged.get(d,{"transactions":[]})
        wn,tr,lp,la,wd,uk=calc_stats(d,info["transactions"],prev_map)
        parts=[f"本週新增{wn}戶",f"共{tr}筆"]
        if lp: parts.append(f"最新{lp}萬")
        if wd: parts.append(f"乾旱{wd}週")
        print(f"  {d}: {', '.join(parts)}")
        cases.append({"name":d,"weekly_new":wn,"total_records":tr,
                      "latest_price":lp,"latest_area":la,"weeks_dry":wd,"unit_keys":uk})
    out={"run_at":ns,"output_file":"GitHub Actions 雲端自動更新（無PPT）","cases":cases}
    OUTPUT_PATH.write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding="utf-8")
    print(f"\nOK {OUTPUT_PATH}")

if __name__=="__main__":
    main()
