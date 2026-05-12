#!/usr/bin/env python3
import json
import os
from datetime import datetime
from openai import OpenAI

def generate_chips():
    client = OpenAI(
        base_url="https://models.inference.ai.azure.com",
        api_key=os.environ["GITHUB_TOKEN"],
    )
    today = datetime.now().strftime("%Y/%m/%d")

    prompt = f"""今天是 {today}（台灣時間）。
你是台灣股市籌碼分析師，請生成今日台股籌碼面報告。

請直接輸出符合以下格式的 JSON，不要加任何說明文字：

{{
  "date": "{today}",
  "taiex": "指數（漲跌點 / 漲跌%）",
  "volume": "成交量（億元，純數字）",
  "sentiment": "市場情緒一句話描述",
  "foreign": 外資買賣超億元（數字，賣超為負）,
  "investment_trust": 投信買賣超億元（數字）,
  "dealer": 自營商買賣超億元（數字，賣超為負）,
  "foreign_streak": "外資連續買超或賣超描述",
  "trust_streak": "投信連續買超描述",
  "margin": {{
    "balance": "融資餘額億元",
    "change": "變動（含+/-）",
    "usage_rate": "使用率（含%）"
  }},
  "short": {{
    "balance": "融券餘額萬張",
    "change": "變動（含+/-）"
  }},
  "futures": {{
    "foreign_net": 外資台指期淨部位口數（整數，空單為負）,
    "direction": "多單或空單",
    "put_call_ratio": PCR數值（小數）,
    "pc_bias": "略偏多/略偏空/中性等描述"
  }},
  "top_buy": [
    {{"stock": "公司名 代號", "shares": "+張數", "who": "外資/投信/外資＋投信"}}
  ],
  "top_sell": [
    {{"stock": "公司名 代號", "shares": "-張數", "who": "外資/自營商"}}
  ],
  "summary": "當日盤勢完整摘要3-5句話",
  "causality": [
    "【因】事件 → 【果】影響"
  ],
  "investment_advice": [
    {{"stock": "公司名 代號", "verdict": "謹慎/觀察/可留意", "reason": "原因", "risk": "風險"}}
  ],
  "recommendations": [
    {{"stock": "公司名 代號", "reason": "推薦原因", "risk": "注意事項"}}
  ],
  "disclaimer": "⚠️ 以上僅為籌碼面分析參考，不構成投資建議，投資人須自行判斷風險。"
}}

要求：
1. top_buy 和 top_sell 各 5 檔，選台灣知名上市公司
2. causality 寫 3 條因果關係
3. investment_advice 寫 3-5 檔
4. recommendations 寫 3 檔
5. 數據要合理，符合台股當前走勢
6. 週末（六日）不開盤，若今天是六日請以前一個交易日為準"""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        max_tokens=4096,
    )

    text = response.choices[0].message.content.strip()
    if text.startswith("```"):
        text = text.split("```", 1)[1]
        if text.startswith("json"):
            text = text[4:]
    if text.endswith("```"):
        text = text[:-3]

    data = json.loads(text.strip())

    os.makedirs("results", exist_ok=True)
    with open("results/chips_report.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"chips_report.json updated: {today}")

if __name__ == "__main__":
    generate_chips()
