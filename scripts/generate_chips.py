#!/usr/bin/env python3
"""
generate_chips.py
量化數據：TWSE（主要）+ Yahoo Finance（交叉驗證）
質化分析：GitHub Models AI（summary / causality / advice）
非交易日：全部跳過，不更新

修正 2026-05-31：
  - get_taiex() Yahoo 備援時，補抓 TWSE FMTQIK 取得成交量（原本寫死 vol=0）
  - get_margin() 失敗時，data.update 步驟補寫 "N/A" 防止 AI 生成 0
"""
import json
import os
import time
import requests
from datetime import datetime
from openai import OpenAI

TWSE_API = "https://openapi.twse.com.tw/v1"
TWSE_RWD = "https://www.twse.com.tw/rwd/zh"
TAIFEX   = "https://openapi.taifex.com.tw/v1"
MAX_RETRIES = 3

HEADERS = {
    "Accept": "application/json",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
}


def sf(val, default=0.0):
    try:
        return float(str(val).replace(",", "").replace("+", "").strip())
    except Exception:
        return default


def find(row, *keywords):
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


def get_raw_old_api(url):
    for i in range(MAX_RETRIES):
        try:
            r = requests.get(url, headers=HEADERS, timeout=15)
            r.raise_for_status()
            d = r.json()
            if d.get("stat") == "OK" and d.get("data"):
                return d
            else:
                print(f"  [API] {url.split('?')[0].split('/')[-1]} 無資料 stat={d.get('stat')}")
        except Exception as e:
            print(f"  [API] {url.split('?')[0].split('/')[-1]} 失敗({i+1}): {e}")
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


def _get_vol_from_fmtqik():
    """備援：從 TWSE FMTQIK 取得最新一日成交金額（億元）"""
    try:
        r = requests.get(
            f"{TWSE_RWD}/afterTrading/FMTQIK?response=json",
            headers=HEADERS, timeout=12,
        )
        d = r.json()
        if d.get("stat") == "OK" and d.get("data"):
            fields = d.get("fields", [])
            last_row = d["data"][-1]
            vol_idx = next(
                (i for i, f in enumerate(fields) if "成交金額" in str(f)), None
            )
            if vol_idx is None:
                vol_idx = 2
            if vol_idx < len(last_row):
                raw_val = sf(last_row[vol_idx])
                if raw_val > 1e10:
                    vol = round(raw_val / 1e8)
                elif raw_val > 1e7:
                    vol = round(raw_val / 1e5)
                else:
                    vol = 0
                if vol > 0:
                    print(f"  [VOL] FMTQIK 補成交量: {vol}億 (raw={raw_val})")
                    return vol
    except Exception as e:
        print(f"  [VOL] FMTQIK 失敗: {e}")
    return 0


def get_taiex():
    """
    大盤加權指數：TWSE openapi 優先，失敗時 fallback Yahoo Finance。
    ★ 修正：Yahoo 備援時補抓 TWSE FMTQIK 取得成交量，不再寫死 vol=0。
    """
    data = safe_get(f"{TWSE_API}/exchangeReport/MI_INDEX")
    if data:
        for row in data:
            name = str(find(row, "名稱") or "")
            if "加權" in name and "電" not in name and "OTC" not in name:
                close = sf(find(row, "收盤"))
                chg   = sf(find(row, "漲跌點"))
                pct   = sf(find(row, "百分比", "漲跌(%)"))
                vol   = sf(find(row, "成交金額"))
                if close > 0:
                    return {"close": close, "change": chg, "pct": pct, "vol": vol, "source": "TWSE"}

    url = "https://query1.finance.yahoo.com/v8/finance/chart/%5ETWII?interval=1d&range=5d"
    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        d = r.json()
        result = d["chart"]["result"][0]
        closes = result["indicators"]["quote"][0]["close"]
        valid = [c for c in closes if c is not None]
        if len(valid) >= 2:
            curr = round(valid[-1], 2)
            prev = round(valid[-2], 2)
            chg  = round(curr - prev, 2)
            pct  = round(chg / prev * 100, 2)

            # ★ 嘗試從 Yahoo volume 欄取成交量
            vol = 0
            try:
                vols = result["indicators"]["quote"][0].get("volume", [])
                valid_vols = [v for v in vols if v is not None]
                if valid_vols and valid_vols[-1]:
                    raw_vol = valid_vols[-1]
                    if raw_vol > 1e10:
                        vol = round(raw_vol / 1e8)
                    elif raw_vol > 1e7:
                        vol = round(raw_vol / 1e5)
            except Exception:
                pass

            # ★ Yahoo volume 仍為 0，改用 TWSE FMTQIK 補抓
            if vol == 0:
                vol = _get_vol_from_fmtqik()

            print(f"  [TAIEX] TWSE 無資料，Yahoo 備援: {curr:.2f}，成交量: {vol}億")
            return {"close": curr, "change": chg, "pct": pct, "vol": vol, "source": "Yahoo"}
    except Exception as e:
        print(f"  [TAIEX] Yahoo 備援失敗: {e}")
    return None


def get_institutional():
    d = get_raw_old_api(f"{TWSE_RWD}/fund/BFI82U?response=json")
    if not d:
        return None, None
    api_date = d.get("date", "")
    fields = d["fields"]
    rows = [dict(zip(fields, row)) for row in d["data"]]
    res = {"foreign": 0.0, "investment_trust": 0.0, "dealer": 0.0}
    for row in rows:
        name = str(row.get("單位名稱", ""))
        raw  = sf(row.get("買賣差額", 0))
        bil  = round(raw / 1e8, 1)
        if "合計" in name:
            continue
        if "外資" in name and "不含" in name:
            res["foreign"] = bil
        elif "投信" in name:
            res["investment_trust"] = bil
        elif "自營商" in name and "外資" not in name:
            res["dealer"] = round(res["dealer"] + bil, 1)
    return res, api_date


def get_top_stocks():
    d = get_raw_old_api(f"{TWSE_RWD}/fund/TWT38U?response=json")
    if not d:
        return [], []

    parsed = []
    for row in d["data"]:
        try:
            code = str(row[1]).strip()
            name = str(row[2]).strip()
            net_shares = sf(row[11])
            if code and name:
                parsed.append({"code": code, "name": name, "net": net_shares})
        except Exception:
            continue

    if not parsed:
        return [], []

    srt = sorted(parsed, key=lambda x: x["net"], reverse=True)

    def fmt(p, sign):
        lots = abs(int(p["net"])) // 1000
        return {"stock": f"{p['name']} {p['code']}", "shares": f"{sign}{lots:,}", "who": "外資"}

    buy  = [fmt(p, "+") for p in srt[:5]           if p["net"] > 0]
    sell = [fmt(p, "-") for p in reversed(srt[-5:]) if p["net"] < 0]
    return buy, sell


def get_margin():
    data = safe_get(f"{TWSE_API}/exchangeReport/MI_MARGN")
    if not data:
        return None
    row = data[-1]
    for r in data:
        if "上市" in str(find(r, "種類") or ""):
            row = r
            break
    m_bal = round(sf(find(row, "融資餘額") or 0) / 100_000)
    m_chg = round(sf(find(row, "融資增減") or 0) / 100_000)
    s_bal = round(sf(find(row, "融券餘額") or 0) / 10_000, 1)
    s_chg = round(sf(find(row, "融券增減") or 0) / 10_000, 1)
    return {"m_bal": m_bal, "m_chg": m_chg, "s_bal": s_bal, "s_chg": s_chg}


def get_futures_net():
    data = safe_get(f"{TAIFEX}/DailyForeignInstitutionalInvestorsFuturesAndOptions")
    if not data:
        return None
    for row in data:
        code  = str(find(row, "ContractCode", "商品代號") or "")
        ident = str(find(row, "Identity", "身份別") or "")
        if "TX" in code and "外資" in ident:
            return int(sf(find(row, "NetPosition", "淨部位") or 0))
    return None


def generate_chips():
    today = datetime.now().strftime("%Y/%m/%d")
    client = OpenAI(
        base_url="https://models.inference.ai.azure.com",
        api_key=os.environ["GITHUB_TOKEN"],
    )

    print(f"[{today}] 抓取台股籌碼...")

    taiex             = get_taiex()
    inst, api_date    = get_institutional()
    top_buy, top_sell = get_top_stocks()
    margin            = get_margin()
    fut_net           = get_futures_net()

    has_real = bool(taiex)

    if api_date and len(api_date) == 8:
        data_date = f"{api_date[:4]}/{api_date[4:6]}/{api_date[6:]}"
    else:
        data_date = today

    cross_val_note = ""
    if taiex and taiex.get("source") == "Yahoo":
        cross_val_note = "（資料來源：Yahoo Finance 備援）"

    print(f"  大盤:{'✓' if taiex else '✗'}  法人:{'✓' if inst else '✗'}  "
          f"個股:{'✓' if top_buy else '✗'}  融資券:{'✓' if margin else '✗'}  "
          f"期貨:{'✓' if fut_net is not None else '✗'}")

    if not has_real:
        print("  大盤指數無法取得（TWSE + Yahoo 均失敗），不更新。")
        return

    close = taiex["close"]
    chg   = taiex["change"]
    pct   = taiex["pct"]
    vol   = round(taiex["vol"])
    sign  = "+" if chg >= 0 else ""
    taiex_str = f"{close:,.2f}（{sign}{chg:.0f} / {sign}{pct:.2f}%）{cross_val_note}"

    foreign = inst.get("foreign", 0) if inst else 0
    trust   = inst.get("investment_trust", 0) if inst else 0
    dealer  = inst.get("dealer", 0) if inst else 0

    m_bal_s  = f"{margin['m_bal']}" if margin else "N/A"
    m_chg_s  = f"{margin['m_chg']:+}" if margin else "N/A"
    m_rate_s = "N/A"
    s_bal_s  = f"{margin['s_bal']}" if margin else "N/A"
    s_chg_s  = f"{margin['s_chg']:+}" if margin else "N/A"

    fut_val = fut_net if fut_net is not None else 0
    fut_dir = "多單" if fut_val >= 0 else "空單"
    fut_str = (f"外資台指期淨部位：{fut_val:+,}口（{fut_dir}）"
               if fut_net is not None else "外資台指期：資料暫無")

    vol_display = str(vol) if vol > 0 else "N/A"

    quant = (
        f"加權指數：{taiex_str}\n"
        f"成交金額：{vol_display}億元\n"
        f"外資買賣超：{foreign:+.1f}億\n"
        f"投信買賣超：{trust:+.1f}億\n"
        f"自營商買賣超：{dealer:+.1f}億\n"
        f"融資餘額：{m_bal_s}億（{m_chg_s}億）\n"
        f"融券餘額：{s_bal_s}萬張（{s_chg_s}萬張）\n"
        f"{fut_str}\n"
        f"外資大買：{', '.join(s['stock'] for s in top_buy[:3]) if top_buy else 'N/A'}\n"
        f"外資大賣：{', '.join(s['stock'] for s in top_sell[:3]) if top_sell else 'N/A'}"
    )

    prompt = f"""以下是 {data_date} 台股真實籌碼數據：
{quant}

請根據以上數據輸出 JSON（不加說明）：
{{
  "date": "{data_date}",
  "taiex": "{taiex_str}",
  "volume": "{vol_display}",
  "sentiment": "市場情緒一句話",
  "foreign": {foreign},
  "investment_trust": {trust},
  "dealer": {dealer},
  "foreign_streak": "根據外資數值推斷連續買超或賣超描述",
  "trust_streak": "投信連續買超描述",
  "margin": {{"balance": "{m_bal_s}", "change": "{m_chg_s}", "usage_rate": "{m_rate_s}"}},
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

    text = call_api(client, prompt, max_tokens=3000)
    data = parse_json(text)
    if not data:
        text = call_api(client, prompt, max_tokens=3000)
        data = parse_json(text)
    if not data:
        print("  ✗ AI 無法生成報告")
        return

    # 強制用真實量化數值覆蓋 AI 輸出
    data.update({
        "date": data_date,
        "taiex": taiex_str,
        "volume": vol_display,
        "foreign": foreign,
        "investment_trust": trust,
        "dealer": dealer,
        "top_buy": top_buy,
        "top_sell": top_sell,
    })
    # margin / short：API 成功時覆蓋；失敗時強制填 N/A，杜絕 AI 生成假 0
    if margin:
        data["margin"].update({"balance": m_bal_s, "change": m_chg_s})
        data["short"].update({"balance": s_bal_s, "change": s_chg_s})
    else:
        data["margin"] = {"balance": "N/A", "change": "N/A", "usage_rate": "N/A"}
        data["short"]  = {"balance": "N/A", "change": "N/A"}

    if fut_net is not None:
        data["futures"].update({"foreign_net": fut_val, "direction": fut_dir})

    os.makedirs("results", exist_ok=True)
    with open("results/chips_report.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    src = "TWSE 真實數據" if taiex.get("source") == "TWSE" else "Yahoo Finance 備援"
    print(f"  chips_report.json 已更新（{data_date}，{src}）")
    print(f"  volume={data['volume']}  margin={data['margin']['balance']}  "
          f"short={data['short']['balance']}")


if __name__ == "__main__":
    generate_chips()
