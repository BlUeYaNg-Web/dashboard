#!/usr/bin/env python3
import json
import os
from datetime import datetime
from openai import OpenAI

def generate_news():
    client = OpenAI(
        base_url="https://models.inference.ai.azure.com",
        api_key=os.environ["GITHUB_TOKEN"],
    )
    today = datetime.now().strftime("%Y/%m/%d")

    prompt = f"""今天是 {today}。你是台灣房地產建設公司的業務分析師，每天早上為銷售團隊準備財經新聞摘要。
請生成今日財經新聞摘要，用於銷售「土城區新成屋」。請直接輸出符合以下格式的 JSON，不要加任何說明文字：

{{
  "fetched_at": "{today} 09:00",
  "total": 13,
  "categories": {{
    "台灣財經新聞": [
      {{
        "title": "新聞標題",
        "summary": "2-3句客觀摘要",
        "source": "媒體來源",
        "credibility": "⭐⭐⭐⭐⭐",
        "analysis": "市場影響分析",
        "sales_pitch": "連結到土城新成屋的銷售說詞"
      }}
    ],
    "房地產新聞": [
      {{
        "title": "新聞標題",
        "summary": "2-3句客觀摘要",
        "source": "媒體來源",
        "credibility": "⭐⭐⭐⭐",
        "analysis": "市場影響分析",
        "sales_pitch": "連結到土城新成屋的銷售說詞"
      }}
    ],
    "中國財經經濟新聞": [
      {{
        "title": "新聞標題",
        "summary": "2-3句客觀摘要",
        "source": "媒體來源",
        "credibility": "⭐⭐⭐⭐",
        "analysis": "市場影響分析"
      }}
    ],
    "國際重大新聞": [
      {{
        "title": "新聞標題",
        "summary": "2-3句客觀摘要",
        "source": "媒體來源",
        "credibility": "⭐⭐⭐⭐",
        "analysis": "市場影響分析",
        "sales_pitch": "連結到土城新成屋的銷售說詞"
      }}
    ]
  }}
}}

每個類別至少 2-4 則新聞，內容要：
1. 符合當前時事（今天 {today}）
2. 專業、具體、有數據
3. sales_pitch 要自然連結財經新聞到「現在是買土城新成屋的好時機」
4. 台灣財經、房地產、國際 3 個類別都要有 sales_pitch，中國新聞不用"""

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
    with open("results/daily_news.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"daily_news.json updated: {today}")

if __name__ == "__main__":
    generate_news()
