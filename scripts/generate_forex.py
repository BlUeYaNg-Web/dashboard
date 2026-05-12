#!/usr/bin/env python3
import anthropic
import json
import os
import requests
from datetime import datetime, timedelta

HISTORY_FILE = "results/forex_history.json"

def fetch_real_rates():
    """Fetch live rates from free API (no key required)."""
    try:
        r = requests.get("https://open.er-api.com/v6/latest/USD", timeout=10)
        data = r.json()
        if data.get("result") == "success":
            raw = data["rates"]
            usd_twd = raw.get("TWD", 32.5)
            cny_twd = usd_twd / raw.get("CNY", 7.2)
            usd_cny = raw.get("CNY", 7.2)
            return {"USD_TWD": usd_twd, "CNY_TWD": cny_twd, "USD_CNY": usd_cny}
    except Exception:
        pass
    return None

def update_history(usd_twd, cny_twd, usd_cny):
    today = datetime.now().strftime("%Y/%m/%d")
    try:
        with open(HISTORY_FILE, encoding="utf-8") as f:
            hist = json.load(f)
    except Exception:
        hist = {"entries": []}
    entries = hist.get("entries", [])
    if entries and entries[-1].get("date") == today:
        entries[-1] = {"date": today, "USD_TWD": usd_twd, "CNY_TWD": cny_twd, "USD_CNY": usd_cny}
    else:
        entries.append({"date": today, "USD_TWD": usd_twd, "CNY_TWD": cny_twd, "USD_CNY": usd_cny})
    hist["entries"] = entries[-60:]  # keep 60 days
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(hist, f, ensure_ascii=False, indent=2)

def generate_forex():
    today = datetime.now().strftime("%Y/%m/%d")
    fetched_at = datetime.now().strftime("%Y/%m/%d %H:%M")

    real = fetch_real_rates()
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    if real:
        rate_info = (
            f"USD/TWD={real['USD_TWD']:.2f}, CNY/TWD={real['CNY_TWD']:.2f}, USD/CNY={real['USD_CNY']:.2f}"
        )
        prompt = f"""今天是 {today}，台灣銀行即期匯率：{rate_info}

請直接輸出以下格式 JSON，change/change_pct 請合理估算，不加任何說明：
{{
  "fetched_at": "{fetched_at}",
  "rates": {{
    "USD_TWD": {{"rate": {real['USD_TWD']:.2f}, "rate_buy": 買入價, "rate_sell": 賣出價, "change": 變動, "change_pct": 變動%}},
    "CNY_TWD": {{"rate": {real['CNY_TWD']:.2f}, "rate_buy": 買入價, "rate_sell": 賣出價, "change": 變動, "change_pct": 變動%}},
    "USD_CNY": {{"rate": {real['USD_CNY']:.2f}, "change": 變動, "change_pct": 變動%}}
  }},
  "news": {{
    "台幣（TWD）": [{{"title":"...", "summary":"...", "source":"...", "credibility":"⭐⭐⭐⭐"}}],
    "美元（USD）": [{{"title":"...", "summary":"...", "source":"...", "credibility":"⭐⭐⭐⭐"}}],
    "人民幣（CNY）": [{{"title":"...", "summary":"...", "source":"...", "credibility":"⭐⭐⭐"}}]
  }}
}}"""
    else:
        prompt = f"""今天是 {today}，請生成台灣銀行即期匯率報告。
請直接輸出以下格式 JSON，不加任何說明：
{{
  "fetched_at": "{fetched_at}",
  "rates": {{
    "USD_TWD": {{"rate": 數字, "rate_buy": 買入, "rate_sell": 賣出, "change": 變動, "change_pct": 變動%}},
    "CNY_TWD": {{"rate": 數字, "rate_buy": 買入, "rate_sell": 賣出, "change": 變動, "change_pct": 變動%}},
    "USD_CNY": {{"rate": 數字, "change": 變動, "change_pct": 變動%}}
  }},
  "news": {{
    "台幣（TWD）": [{{"title":"...", "summary":"...", "source":"...", "credibility":"⭐⭐⭐⭐"}}],
    "美元（USD）": [{{"title":"...", "summary":"...", "source":"...", "credibility":"⭐⭐⭐⭐"}}],
    "人民幣（CNY）": [{{"title":"...", "summary":"...", "source":"...", "credibility":"⭐⭐⭐"}}]
  }}
}}"""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        messages=[{"role": "user", "content": prompt}]
    )

    text = message.content[0].text.strip()
    if text.startswith("```"):
        text = text.split("```", 1)[1]
        if text.startswith("json"):
            text = text[4:]
    if text.endswith("```"):
        text = text[:-3]

    data = json.loads(text.strip())

    os.makedirs("results", exist_ok=True)
    with open("results/forex_report.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    r = data.get("rates", {})
    usd_twd = r.get("USD_TWD", {}).get("rate", 0)
    cny_twd = r.get("CNY_TWD", {}).get("rate", 0)
    usd_cny = r.get("USD_CNY", {}).get("rate", 0)
    update_history(usd_twd, cny_twd, usd_cny)

    source = "live API + Claude" if real else "Claude AI"
    print(f"forex_report.json + forex_history.json updated: {today} (source: {source})")

if __name__ == "__main__":
    generate_forex()
