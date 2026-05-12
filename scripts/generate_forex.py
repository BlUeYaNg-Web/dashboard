#!/usr/bin/env python3
import anthropic
import json
import os
import requests
from datetime import datetime

CURRENCIES = [
    ("USD", "美元"),
    ("JPY", "日幣（百元）"),
    ("CNY", "人民幣"),
    ("EUR", "歐元"),
]

def fetch_real_rates():
    """Fetch live rates from free API (no key required)."""
    try:
        r = requests.get("https://open.er-api.com/v6/latest/USD", timeout=10)
        data = r.json()
        if data.get("result") == "success":
            raw = data["rates"]
            twd = raw.get("TWD", 32.5)
            return {
                "USD": twd,
                "JPY": twd / raw["JPY"] * 100,   # TWD per 100 JPY
                "CNY": twd / raw["CNY"],
                "EUR": twd / raw["EUR"],
            }
    except Exception:
        pass
    return None

def generate_forex():
    today = datetime.now().strftime("%Y/%m/%d")
    updated_at = datetime.now().strftime("%Y/%m/%d %H:%M")

    real = fetch_real_rates()
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

    if real:
        real_block = "\n".join(
            f"- {name}：1 TWD = {real[code]:.4f}" for code, name in CURRENCIES
        )
        prompt = f"""今天是 {today}，台灣銀行即期匯率如下：
{real_block}

請直接輸出以下格式的 JSON，change 與 change_pct 請根據近期走勢合理估算（小數），不要加任何說明：
{{
  "updated_at": "{updated_at}",
  "rates": [
    {{"currency": "USD", "name": "美元",      "rate": {real['USD']:.2f}, "change": 估算變動, "change_pct": 估算變動%}},
    {{"currency": "JPY", "name": "日幣（百元）","rate": {real['JPY']:.2f}, "change": 估算變動, "change_pct": 估算變動%}},
    {{"currency": "CNY", "name": "人民幣",    "rate": {real['CNY']:.2f}, "change": 估算變動, "change_pct": 估算變動%}},
    {{"currency": "EUR", "name": "歐元",      "rate": {real['EUR']:.2f}, "change": 估算變動, "change_pct": 估算變動%}}
  ],
  "summary": "台幣今日走勢一句話",
  "impact": "外幣走強/弱對台灣購屋族的意義（一句話，提到土城新成屋保值）"
}}"""
    else:
        prompt = f"""今天是 {today}，請生成台灣銀行即期匯率資訊。
請直接輸出以下格式的 JSON，不要加任何說明：
{{
  "updated_at": "{updated_at}",
  "rates": [
    {{"currency": "USD", "name": "美元",      "rate": 數字, "change": 變動, "change_pct": 變動%}},
    {{"currency": "JPY", "name": "日幣（百元）","rate": 數字, "change": 變動, "change_pct": 變動%}},
    {{"currency": "CNY", "name": "人民幣",    "rate": 數字, "change": 變動, "change_pct": 變動%}},
    {{"currency": "EUR", "name": "歐元",      "rate": 數字, "change": 變動, "change_pct": 變動%}}
  ],
  "summary": "台幣今日走勢一句話",
  "impact": "外幣走強/弱對台灣購屋族的意義（一句話，提到土城新成屋保值）"
}}"""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
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
    with open("results/forex.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    source = "live API" if real else "Claude AI"
    print(f"forex.json updated: {today} (source: {source})")

if __name__ == "__main__":
    generate_forex()
