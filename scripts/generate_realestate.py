"""
每週日 13:05 台灣時間 (05:05 UTC) 由 GitHub Actions 執行：
1. 從 dualtaipei.datazen.info 抓取 11 個建案（主要來源）
2. 從 Google Sheet 補充土城區資料（交叉比對，含和耀美家雅居）
3. 計算本週新增（與上週 pptx_update.json 比對）
4. 更新 results/pptx_update.json（GitHub Actions 自動 git push）

不需要電腦開機、不需要 Claude 在線，GitHub Actions 全自動執行。
PPT 檔案更新仍需電腦開機（由 Windows 工作排程器負責）。
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
    ("丰藏",       "家泰嘉潤丰藏", "newtaipei"),
    ("台信琢岳",   "台信琢岳",     "newtaipei"),
    ("紅布朗花園", "紅布朗花園",   "newtaipei"),
    ("寶佳淳青",   "寶佳淳青",     "newtaipei"),
    ("松陽馥麗",   "松陽馥麗",     "newtaipei"),
    ("新濠岳",     "新濠岳",       "newtaipei"),
    ("朗沐",       "朗沐",         "newtaipei"),
    ("合康雙匯",   "合康雙匯",     "newtaipei"),
    ("迴東騰",     "迴東騰",       "newtaipei"),
    ("天好運3",    "天好運3",      "newtaipei"),
    ("若水秧翠",   "若水秧翠",     "newtaipei"),
]

GS_CASE_MAP = {
    "家泰嘉潤丰藏":  "丰藏",
    "台信琢岳":      "台信琢岳",
    "紅布朗花園":    "紅布朗花園",
    "寶佳淳青":      "寶佳淳青",
    "松陽馥麗":      "松陽馥麗",
    "新濠岳":        "新濠岳",
    "朗沐":          "朗沐",
    "合康雙匯":      "合康雙匯",
    "迴東騰":        "迴東騰",
    "天好運3":       "天好運3",
    "若水秧翠":      "若水秧翠",
    "和耀美家 雅居": "和耀美家雅居",
}

ALL_DISPLAY_NAMES = [
    "丰藏", "台信琢岳", "紅布朗花園", "寶佳淳青", "松陽馥麗", "新濠岳",
    "朗沐", "合康雙匯", "迴東騰", "天好運3", "若水秧翠", "和耀美家雅居",
]


def fetch_json(url):
    req = urllib.request.Request(url, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.loads(r.read().decode("utf-8"))


def roc_to_iso(roc_str):
    s = str(roc_str).strip()
    if len(s) == 7 and s.isdigit():
        year = int(s[:3]) + 1911
        month = int(s[3:5])
        day = int(s[5:7])
        return f"{year}-{month:02d}-{day:02d}T00:00:00+00:00"
    return None


def parse_price_float(price_str):
    m = re.search(r"([\d.]+)", str(price_str))
    return float(m.group(1)) if m else None


def normalize_unit_key(unit_str):
    m = re.match(r"^([A-Za-z]+\d*)-0*(\d+)F?$", unit_str.strip())
    if m:
        return (m.group(1).upper(), int(m.group(2)))
    return (unit_str.strip(), 0)


def gs_unit_to_std(floor_str):
    s = floor_str.strip()
    if "/" not in s:
        return s
    parts = s.split("/")
    if len(parts) < 2:
        return s
    unit_part = parts[0].strip()
    floor_part = parts[1].strip()
    if re.match(r"^[A-Za-z]+\d*-0*\d+F?$", floor_part):
        return floor_part
    floor_num = re.sub(r"^0*(\d+)F?$", r"\1F", floor_part, flags=re.IGNORECASE)
    if unit_part:
        return f"{unit_part}-{floor_num}"
    return floor_num


def parse_dt(t):
    try:
        return datetime.fromisoformat(t.get("date", "").replace("Z", "+00:00"))
    except Exception:
        return datetime.min.replace(tzinfo=timezone.utc)


class TableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows = []
        self.current_row = []
        self.in_td = False
        self.current_text = ""

    def handle_starttag(self, tag, attrs):
        if tag in ("td", "th"):
            self.in_td = True
            self.current_text = ""
        elif tag == "tr":
            self.current_row = []

    def handle_endtag(self, tag):
        if tag in ("td", "th"):
            self.current_row.append(self.current_text.strip())
            self.in_td = False
        elif tag == "tr":
            if any(c.strip() for c in self.current_row):
                self.rows.append(self.current_row[:])

    def handle_data(self, data):
        if self.in_td:
            self.current_text += data


def fetch_dualtaipei():
    print("=== dualtaipei 抓取 ===")
    result = {}
    for key, full_name, city in CASES:
        print(f"  {key}（{full_name}）...", end="", flush=True)
        encoded = urllib.parse.quote(full_name)
        base = f"{BASE_URL}/api/{city}/project/{encoded}"
        try:
            monthly = fetch_json(f"{base}?type=monthly").get("monthly", [])
            txn_data = fetch_json(f"{base}?type=transactions")
            transactions = txn_data.get("transactions", [])
            total = txn_data.get("pagination", {}).get("total", len(transactions))
            offset = len(transactions)
            while offset < total:
                more = fetch_json(f"{base}?type=transactions&limit=50&offset={offset}")
                batch = more.get("transactions", [])
                if not batch:
                    break
                transactions.extend(batch)
                offset += len(batch)
            result[key] = {"monthly": monthly, "transactions": transactions}
            print(f" OK {len(transactions)} 筆")
        except Exception as e:
            print(f" FAIL: {e}")
            result[key] = {"monthly": [], "transactions": []}
    return result


def fetch_google_sheet():
    print("\n=== Google Sheet 抓取 ===")
    gs_headers = {**HEADERS, "Accept": "text/html,*/*"}
    req = urllib.request.Request(GS_URL, headers=gs_headers)
    with urllib.request.urlopen(req, timeout=30) as r:
        html = r.read().decode("utf-8", errors="replace")
    parser = TableParser()
    parser.feed(html)
    case_txns = defaultdict(list)
    for row in parser.rows:
        if len(row) < 10:
            continue
        if not any("土城" in c for c in row):
            continue
        case_key = None
        name_col = None
        for col_idx in [4, 3]:
            if col_idx >= len(row):
                continue
            cell_val = row[col_idx].strip()
            for gs_name, key in GS_CASE_MAP.items():
                if cell_val == gs_name.strip():
                    case_key = key
                    name_col = col_idx
                    break
            if case_key:
                break
        if not case_key:
            continue
        offset = name_col - 3
        iso_date = roc_to_iso(row[1].strip())
        unit = gs_unit_to_std(row[4 + offset].strip())
        price = parse_price_float(row[5 + offset].strip())
        if not iso_date or not unit or price is None:
            continue
        case_txns[case_key].append({"date": iso_date, "unit": unit,
                                     "unitPrice": price, "_source": "google_sheet"})
    total_rows = sum(len(v) for v in case_txns.values())
    print(f"  OK {total_rows} 筆（{len(case_txns)} 個案名）")
    return case_txns


def cross_validate(dt_data, gs_data):
    """交叉比對：GS 有而 dualtaipei 沒有的戶，補充進去"""
    print("\n=== 交叉比對 ===")
    merged = {}
    for key, dt_info in dt_data.items():
        transactions = list(dt_info["transactions"])
        existing_units = {normalize_unit_key(t.get("unit", "")) for t in transactions}
        gs_txns = gs_data.get(key, [])
        added = 0
        for t in gs_txns:
            uk = normalize_unit_key(t["unit"])
            if uk not in existing_units:
                transactions.append(t)
                existing_units.add(uk)
                added += 1
        if added:
            print(f"  [{key}] GS 補充 {added} 筆（dualtaipei 沒有）")
        elif gs_txns:
            print(f"  [{key}] 兩來源一致")
        else:
            print(f"  [{key}] GS 無資料")
        merged[key] = {"monthly": dt_info["monthly"], "transactions": transactions}
    for key, gs_txns in gs_data.items():
        if key not in merged:
            merged[key] = {"monthly": [], "transactions": list(gs_txns)}
            print(f"  [{key}] 僅 GS 來源，{len(gs_txns)} 筆")
    return merged


def calc_stats(transactions, prev_total):
    if not transactions:
        return 0, 0, None, None, None
    current_total = len(transactions)
    weekly_new = max(0, current_total - prev_total) if prev_total is not None else 0
    sorted_txns = sorted(transactions, key=parse_dt)
    latest = sorted_txns[-1]
    latest_price = round(latest.get("unitPrice") or 0, 2) or None
    latest_area = round(latest.get("area") or 0, 1) or None
    now = datetime.now(timezone.utc)
    days_ago = (now - parse_dt(latest)).days
    weeks_dry = days_ago // 7 if weekly_new == 0 else 0
    return weekly_new, current_total, latest_price, latest_area, weeks_dry


def main():
    now_str = datetime.now().strftime("%Y/%m/%d %H:%M")
    print(f"[{now_str}] 開始更新實價登錄資料（GitHub Actions 雲端模式）...")

    prev_totals = {}
    try:
        prev_json = json.loads(OUTPUT_PATH.read_text(encoding="utf-8"))
        for c in prev_json.get("cases", []):
            prev_totals[c["name"]] = c.get("total_records", 0)
        nonzero = {k: v for k, v in prev_totals.items() if v}
        if nonzero:
            print(f"上週有資料：{nonzero}")
    except Exception:
        print("無上週資料，本週新增顯示 0")

    dt_data = fetch_dualtaipei()

    try:
        gs_data = fetch_google_sheet()
    except Exception as e:
        print(f"  Google Sheet 失敗：{e}（改用純 dualtaipei）")
        gs_data = {}

    merged = cross_validate(dt_data, gs_data)

    print("\n=== 統計結果 ===")
    cases = []
    for display in ALL_DISPLAY_NAMES:
        info = merged.get(display, {"transactions": []})
        txns = info["transactions"]
        prev_total = prev_totals.get(display, 0)
        weekly_new, total_records, latest_price, latest_area, weeks_dry = calc_stats(txns, prev_total)
        parts = [f"本週新增 {weekly_new} 戶", f"共 {total_records} 筆"]
        if latest_price:
            parts.append(f"最新 {latest_price} 萬")
        if weeks_dry:
            parts.append(f"乾旱 {weeks_dry} 週")
        print(f"  {display}: {', '.join(parts)}")
        cases.append({"name": display, "weekly_new": weekly_new,
                       "total_records": total_records, "latest_price": latest_price,
                       "latest_area": latest_area, "weeks_dry": weeks_dry})

    output = {"run_at": now_str, "output_file": "GitHub Actions 雲端自動更新（無PPT）",
              "cases": cases}
    OUTPUT_PATH.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nOK 已寫出 {OUTPUT_PATH}")


if __name__ == "__main__":
    main()