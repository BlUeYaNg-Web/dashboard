#!/usr/bin/env python3
import json
import os
import time
import requests
from datetime import datetime
from openai import OpenAI

HISTORY_FILE = "results/forex_history.json"
MAX_RETRIES = 3

def fetch_real_rates():
    try:
        r = requests.get("https://open.er-api.com/v6/latest/USD", timeout=10)
        data = r.json()
        if data.get("result") == "success":
            raw = data["rates"]
            usd_twd = raw.get("TWD", 32.5)
            return {
                "USD_TWD": usd_twd,
                "CNY_TWD": usd_twd / raw.get("CNY", 7.2),
                "USD_CNY": raw.get("CNY", 7.2),
            }
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
    hist["entries"] = entries[-60:]
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(hist, f, ensure_ascii=False, indent=2)

def call_api(client, prompt, max_tokens=2048):
    for attempt in range(MAX_RETRIES):
        try:
            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[{"role": "user", "content": prompt}],
                max_tokens=max_tokens,
            )
            return response.choices[0].message.content.strip()
        except Exception as e:
            print(f"API error (attempt {attempt + 1}/{MAX_RETRIES}): {e}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(2 ** attempt)
    raise RuntimeError("All API retries exhausted")

def parse_json(text):
    if text.startswith("```"):
        text = text.split("```", 1)[1]
        if text.startswith("json"):
            text = text[4:]
    if text.endswith("```"):
        text = text[:-3]
    return json.loads(text.strip())

def generate_forex():
    today = datetime.now().strftime("%Y/%m/%d")
    fetched_at = datetime.now().strftime("%Y/%m/%d %H:%M")

    real = fetch_real_rates()
    client = OpenAI(
        base_url="https://models.inference.ai.azure.com",
        api_key=os.environ["GITHUB_TOKEN"],
    )

    if real:
        rate_info = f"USD/TWD={real['USD_TWD']:.2f}, CNY/TWD={real['CNY_TWD']:.2f}, USD/CNY={real['USD_CNY']:.2f}"
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

    for attempt in range(MAX_RETRIES):
        try:
            text = call_api(client, prompt)
            data = parse_json(text)
            break
        except (json.JSONDecodeError, RuntimeError) as e:
            print(f"Failed (attempt {attempt + 1}/{MAX_RETRIES}): {e}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(2 ** attempt)
            else:
                raise

    os.makedirs("results", exist_ok=True)
    with open("results/forex_report.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    r = data.get("rates", {})
    update_history(
        r.get("USD_TWD", {}).get("rate", 0),
        r.get("CNY_TWD", {}).get("rate", 0),
        r.get("USD_CNY", {}).get("rate", 0),
    )

    print(f"forex_report.json updated: {today} (source: {'live API' if real else 'AI'})")

if __name__ == "__main__":
    generate_forex()
