#!/usr/bin/env python3
"""
generate_chips.py
量化數據：TWSE / TAIFEX 真實 API
質化分析：GitHub Models AI（summary / causality / advice）
非交易日（週末/假日）：全部改用 AI 估算
"""
import json
import os
import time
import requests
from datetime import datetime
from openai import OpenAI

TWSE    = "https://openapi.twse.com.tw/v1"
TAIFEX  = "https://openapi.taifex.com.tw/v1"
MAX_RETRIES = 3
HEADERS = {
    "Accept": "application/json",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
}


# ── 工具 ─────────────────────────────────────────────────────────────────────

def sf(val, default=0.0):
    try:
        return float(str(val).replace(",", "").replace("+", "").strip())
    except Exception:
        return default


def find(row, *keywords):
    """在 row 的 key 中找包含任一 keyword 的欄位值"""
    for kw in keywords:
        for k, v in row.items():
            if kw in k and str(v) not in ("", "--", "N/A", "nan"):
                return v
    return None


def safe_get(url):
    for i in range(MAX_RETRIES):
        try:
            r = requests.get(url, headers=HEADERS, timeout=15)
            r.raise_for_status()
            d = r.json()
            if isinstance(d, list) and d:
                return d
        except Exception as e:
            print(f"  [API] {url.split('/')[-1]} 失敗({i+1}): {e}")
            if i < MAX_RETRIES - 1:
                time.sleep(2 ** i)
    return None


def call_api(client, prompt, max_tokens=2048):
    for i in range(MAX_RETRIES):
        try:
            res = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=max_tokens,
            )
            return res.choices[0].message.content.strip()
        except Exception as e:
            print(f"  [AI] 失敗({i+1}): {e}")
            if i < MAX_RETRIES - 1:
                time.sleep(2 ** i)
    return None


def parse_json(text):
    if not text:
        return None
    t = text.strip()
    if t.startswith("```"):
        t = t.split("```", 1)[1]
        if t.startswith("json"):
            t = t[4:]
    if t.endswith("```"):
        t = t[:-3]
    try:
        return json.loads(t.strip())
    except Exception:
        return None


# ── TWSE / TAIFEX 資料抓取 ────────────────────────────────────────────────────

def get_taiex():
    """大盤加權指數：收盤、漲跌點、漲跌%、成交金額(億)"""
    data = safe_get(f"{TWSE}/exchangeReport/MI_INDEX")
    if not data:
        return None
    for row in data:
        name = str(find(row, "名稱") or "")
        if "加權" in name and "電" not in name and "OTC" not in name:
            close = sf(find(row, "收盤"))
            chg   = sf(find(row, "漲跌點"))
            pct   = sf(find(row, "百分比", "漲跌(%)"))
            vol   = sf(find(row, "成交金額"))
            if close > 0:
                return {"close": close, "change": chg, "pct": pct, "vol": vol}
    return None


def get_institutional():
    """三大法人當日買賣超（億元）— BFI82U 單位為百萬元"""
    data = safe_get(f"{TWSE}/fund/BFI82U")
    if not data:
        return None
    res = {}
    for row in data:
        name = str(find(row, "單位名稱", "名稱") or "")
        raw  = sf(find(row, "買賣差額", "差額", "淨") or 0)
        bil  = round(raw / 100, 1)          # 百萬 → 億
        if "外資" in name and "自營" not in name:
            res["foreign"] = bil
        elif "投信" in name:
            res["investment_trust"] = bil
        elif "自營" in name:
            res["dealer"] = bil
    return res if res else None


def get_top_stocks():
    """外資個股買賣超 top5 買 / top5 賣（張）— TWT38U"""
    data = safe_get(f"{TWSE}/fund/TWT38U")
    if not data:
        return [], []

    def net(row):
        return sf(find(row, "買賣超股數", "買賣差股數") or 0)

    srt = sorted(data, key=net, reverse=True)

    def fmt(row, sign):
        code = str(find(row, "股票代號", "代號") or "")
        name = str(find(row, "股票名稱", "名稱") or code)
        lots = abs(int(net(row))) // 1000   # 股 → 張
        return {"stock": f"{name} {code}", "shares": f"{sign}{lots:,}", "who": "外資"}

    return (
        [fmt(r, "+") for r in srt[:5]],
        [fmt(r, "-") for r in srt[-5:][::-1]],
    )


def get_margin():
    """融資融券整體餘額 — MI_MARGN"""
    data = safe_get(f"{TWSE}/exchangeReport/MI_MARGN")
    if not data:
        return None
    row = data[-1]
    for r in data:
        if "上市" in str(find(r, "種類") or ""):
            row = r
            break
    # 融資餘額(千元) → 億；融券餘額(千股) → 萬張
    m_bal = round(sf(find(row, "融資餘額") or 0) / 100_000)
    m_chg = round(sf(find(row, "融資增減") or 0) / 100_000)
    s_bal = round(sf(find(row, "融券餘額") or 0) / 10_000, 1)
    s_chg = round(sf(find(row, "融券增減") or 0) / 10_000, 1)
    return {"m_bal": m_bal, "m_chg": m_chg, "s_bal": s_bal, "s_chg": s_chg}


def get_futures_net():
    """外資台指期淨部位（口）— TAIFEX"""
    data = safe_get(f"{TAIFEX}/DailyForeignInstitutionalInvestorsFuturesAndOptions")
    if not data:
        return None
    for row in data:
        code  = str(find(row, "ContractCode", "商品代號") or "")
        ident = str(find(row, "Identity", "身份別") or "")
        if "TX" in code and "外資" in ident:
            return int(sf(find(row, "NetPosition", "淨部位") or 0))
    return None


# ── 主流程 ────────────────────────────────────────────────────────────────────

def generate_chips():
    today = datetime.now().strftime("%Y/%m/%d")
    client = OpenAI(
        base_url="https://models.inference.ai.azure.com",
        api_key=os.environ["GITHUB_TOKEN"],
    )

    print(f"[{today}] 抓取台股籌碼...")

    taiex             = get_taiex()
    inst              = get_institutional()
    top_buy, top_sell = get_top_stocks()
    margin            = get_margin()
    fut_net           = get_futures_net()

    has_real = bool(taiex)
    print(f"  大盤:{'✓' if taiex else '✗'}  法人:{'✓' if inst else '✗'}  "
          f"個股:{'✓' if top_buy else '✗'}  融資券:{'✓' if margin else '✗'}  "
          f"期貨:{'✓' if fut_net is not None else '✗'}")

    if has_real:
        close = taiex["close"]
        chg   = taiex["change"]
        pct   = taiex["pct"]
        vol   = round(taiex["vol"])
        sign  = "+" if chg >= 0 else ""
        taiex_str = f"{close:.0f} ({sign}{chg:.0f}點 / {sign}{pct:.2f}%)"

        foreign = inst.get("foreign", 0) if inst else 0
        trust   = inst.get("investment_trust", 0) if inst else 0
        dealer  = inst.get("dealer", 0) if inst else 0

        m_bal_s = f"{margin['m_bal']}" if margin else "N/A"
        m_chg_s = f"{margin['m_chg']:+}" if margin else "N/A"
        s_bal_s = f"{margin['s_bal']}" if margin else "N/A"
        s_chg_s = f"{margin['s_chg']:+}" if margin else "N/A"

        fut_val = fut_net if fut_net is not None else 0
        fut_dir = "多單" if fut_val >= 0 else "空單"

        quant = (
            f"加權指數：{taiex_str}\n"
            f"成交金額：{vol}億元\n"
            f"外資買賣超：{foreign:+.1f}億\n"
            f"投信買賣超：{trust:+.1f}億\n"
            f"自營商買賣超：{dealer:+.1f}億\n"
            f"融資餘額：{m_bal_s}億（{m_chg_s}億）\n"
            f"融券餘額：{s_bal_s}萬張（{s_chg_s}萬張）\n"
            f"外資台指期淨部位：{fut_val:+,}口（{fut_dir}）\n"
            f"外資大買：{', '.join(s['stock'] for s in top_buy[:3]) if top_buy else 'N/A'}\n"
            f"外資大賣：{', '.join(s['stock'] for s in top_sell[:3]) if top_sell else 'N/A'}"
        )

        prompt = f"""以下是 {today} 台股真實籌碼數據：
{quant}

請根據以上數據輸出 JSON（不加說明）：
{{
  "date": "{today}",
  "taiex": "{taiex_str}",
  "volume": "{vol}",
  "sentiment": "市場情緒一句話",
  "foreign": {foreign},
  "investment_trust": {trust},
  "dealer": {dealer},
  "foreign_streak": "根據外資數值推斷連續買超或賣超描述",
  "trust_streak": "投信連續買超描述",
  "margin": {{"balance": "{m_bal_s}", "change": "{m_chg_s}", "usage_rate": "使用率%"}},
  "short": {{"balance": "{s_bal_s}", "change": "{s_chg_s}"}},
  "futures": {{
    "foreign_net": {fut_val},
    "direction": "{fut_dir}",
    "put_call_ratio": PCR估算值,
    "pc_bias": "偏多/偏空/中性"
  }},
  "top_buy": {json.dumps(top_buy, ensure_ascii=False)},
  "top_sell": {json.dumps(top_sell, ensure_ascii=False)},
  "summary": "根據真實數據的盤勢摘要3-5句",
  "causality": ["【因】事件 → 【果】影響"],
  "investment_advice": [{{"stock":"名稱 代號","verdict":"謹慎/觀察/可留意","reason":"原因","risk":"風險"}}],
  "recommendations": [{{"stock":"名稱 代號","reason":"推薦原因","risk":"注意事項"}}],
  "disclaimer": "⚠️ 以上僅為籌碼面分析參考，不構成投資建議，投資人須自行判斷風險。"
}}
investment_advice 3-5檔，recommendations 3檔，causality 3條。"""

    else:
        prompt = f"""今天是 {today}（台灣時間），今日為非交易日。請以最近交易日為基準，生成台股籌碼報告。
直接輸出 JSON，不加說明：
{{
  "date": "{today}",
  "taiex": "指數（漲跌點/漲跌%）",
  "volume": "成交量億元",
  "sentiment": "情緒",
  "foreign": 外資億元,
  "investment_trust": 投信億元,
  "dealer": 自營商億元,
  "foreign_streak": "連續描述",
  "trust_streak": "連續描述",
  "margin": {{"balance": "億元", "change": "變動", "usage_rate": "使用率%"}},
  "short": {{"balance": "萬張", "change": "變動"}},
  "futures": {{"foreign_net": 口數, "direction": "多/空單", "put_call_ratio": PCR, "pc_bias": "描述"}},
  "top_buy": [{{"stock": "名稱 代號", "shares": "+張數", "who": "外資/投信"}}],
  "top_sell": [{{"stock": "名稱 代號", "shares": "-張數", "who": "外資/自營商"}}],
  "summary": "摘要3-5句",
  "causality": ["【因】→【果】"],
  "investment_advice": [{{"stock":"名稱 代號","verdict":"謹慎/觀察/可留意","reason":"原因","risk":"風險"}}],
  "recommendations": [{{"stock":"名稱 代號","reason":"推薦","risk":"風險"}}],
  "disclaimer": "⚠️ 以上僅為籌碼面分析參考，不構成投資建議，投資人須自行判斷風險。"
}}
top_buy/top_sell 各5檔，investment_advice 3-5檔，recommendations 3檔，causality 3條。"""

    text = call_api(client, prompt, max_tokens=3000)
    data = parse_json(text)
    if not data:
        text = call_api(client, prompt, max_tokens=3000)
        data = parse_json(text)
    if not data:
        print("  ✗ 無法生成報告")
        return

    # 強制用真實量化數值覆蓋（防止 AI 修改）
    if has_real:
        data.update({
            "date": today, "taiex": taiex_str, "volume": str(vol),
            "foreign": foreign, "investment_trust": trust, "dealer": dealer,
            "top_buy": top_buy, "top_sell": top_sell,
        })
        if margin:
            data["margin"].update({"balance": m_bal_s, "change": m_chg_s})
            data["short"].update({"balance": s_bal_s, "change": s_chg_s})
        if fut_net is not None:
            data["futures"].update({"foreign_net": fut_val, "direction": fut_dir})

    os.makedirs("results", exist_ok=True)
    with open("results/chips_report.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    src = "TWSE 真實數據" if has_real else "AI 估算（非交易日）"
    print(f"  chips_report.json 已更新（{src}）")


if __name__ == "__main__":
    generate_chips()
