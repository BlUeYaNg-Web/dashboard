#!/usr/bin/env python3
import json, os, time, requests
from datetime import datetime, timezone, timedelta
from xml.etree import ElementTree as ET
from openai import OpenAI

MAX_RETRIES = 3
TAIWAN_TZ = timezone(timedelta(hours=8))
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

RSS_FEEDS = {
    "台灣財經新聞": [
        ("Yahoo財經",      "https://tw.news.yahoo.com/rss/finance"),
        ("TechNews科技新報", "https://technews.tw/feed/"),
        ("自由時報財經",    "https://news.ltn.com.tw/rss/business.xml"),
        ("Google新聞-台灣財經", "https://news.google.com/rss/search?q=%E5%8F%B0%E7%81%A3+%E7%B6%93%E6%BF%9F+%E6%99%AF%E6%B0%A3&hl=zh-TW&gl=TW&ceid=TW:zh-Hant"),
    ],
    "房地產新聞": [
        ("住展房屋網",     "https://www.myhousing.com.tw/feed"),
        ("Google新聞-房地產", "https://news.google.com/rss/search?q=%E5%8F%B0%E7%81%A3+%E6%88%BF%E5%9C%B0%E7%94%A2+%E6%88%BF%E5%B8%82&hl=zh-TW&gl=TW&ceid=TW:zh-Hant"),
        ("Google新聞-房價",   "https://news.google.com/rss/search?q=%E6%88%BF%E5%83%B9+%E6%96%B0%E6%88%90%E5%B1%8B+%E8%B2%B7%E6%88%BF&hl=zh-TW&gl=TW&ceid=TW:zh-Hant"),
    ],
    "中國財經經濟新聞": [
        ("BBC中文",        "https://feeds.bbci.co.uk/zhongwen/trad/rss.xml"),
    ],
    "國際重大新聞": [
        ("Reuters Business", "https://feeds.reuters.com/reuters/businessNews"),
        ("Reuters World",    "https://feeds.reuters.com/reuters/worldNews"),
        ("BBC World",        "https://feeds.bbci.co.uk/news/world/rss.xml"),
        ("BBC Business",     "https://feeds.bbci.co.uk/news/business/rss.xml"),
    ],
}

# 各類別是否需要 sales_pitch
NEEDS_PITCH = {
    "台灣財經新聞": True,
    "房地產新聞": True,
    "中國財經經濟新聞": False,
    "國際重大新聞": True,
}


def fetch_rss(source_name, url, max_items=15):
    try:
        r = requests.get(url, headers=HEADERS, timeout=12)
        r.raise_for_status()
        root = ET.fromstring(r.content)

        items = []
        for item in root.findall(".//item")[:max_items]:
            title = (item.findtext("title") or "").strip()
            link = (item.findtext("link") or "").strip()
            if not link:
                link_el = item.find("link")
                if link_el is not None:
                    link = (link_el.text or "").strip()
            if title and link:
                items.append({"title": title, "url": link, "source": source_name})

        if not items:
            ns = {"a": "http://www.w3.org/2005/Atom"}
            for entry in root.findall(".//a:entry", ns)[:max_items]:
                title = (entry.findtext("a:title", "", ns) or "").strip()
                link_el = entry.find("a:link", ns)
                link = (link_el.get("href", "") if link_el is not None else "").strip()
                if title and link:
                    items.append({"title": title, "url": link, "source": source_name})

        print(f"  {source_name}: {len(items)} 篇")
        return items
    except Exception as e:
        print(f"  {source_name} 失敗: {e}")
        return []


def call_api(client, prompt, max_tokens=4096):
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


def build_article_list(articles):
    lines = []
    for i, a in enumerate(articles, 1):
        lines.append(f"{i}. [{a['source']}] {a['title']}")
        lines.append(f"   URL: {a['url']}")
    return "\n".join(lines)


def select_articles(client, today, category, articles, needs_pitch):
    """針對單一類別，呼叫 AI 選出 5 篇並生成摘要"""
    if not articles:
        return []

    article_text = build_article_list(articles[:10])

    pitch_field = (
        '\n        "sales_pitch": "口語化銷售說詞（1-2句，連結土城青埔特區新成屋，像朋友建議）",'
        if needs_pitch else ""
    )
    pitch_rule = "6. sales_pitch 口語化，像朋友在聊天，不要書面廣告腔" if needs_pitch else ""

    prompt = f"""今天是 {today}，以下是【{category}】的真實新聞列表：

{article_text}

請選出對「土城區新成屋銷售團隊」最有參考價值的 5 則（不足 5 則則全選），輸出 JSON 陣列，直接輸出不加說明：

[
  {{
    "title": "使用原文標題，不要修改",
    "summary": "2-3句中文客觀摘要",
    "source": "媒體名稱",
    "url": "直接複製上方原始URL，不可修改或自行生成",
    "credibility": "⭐⭐⭐⭐",
    "analysis": "對房市或景氣的影響分析（2句）",{pitch_field}
  }}
]

規則：
1. url 必須直接複製上方的原始網址
2. 必須選滿 5 則（不足則全選）
3. {pitch_rule}"""

    for attempt in range(MAX_RETRIES):
        try:
            text = call_api(client, prompt, max_tokens=3000)
            result = parse_json(text)
            if isinstance(result, list):
                return result
        except (json.JSONDecodeError, RuntimeError) as e:
            print(f"  [{category}] 失敗({attempt+1}): {e}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(2 ** attempt)
    return []


def generate_news():
    client = OpenAI(
        base_url="https://models.inference.ai.azure.com",
        api_key=os.environ["GITHUB_TOKEN"],
    )
    today = datetime.now(TAIWAN_TZ).strftime("%Y/%m/%d")
    fetched_at = datetime.now(TAIWAN_TZ).strftime("%Y/%m/%d %H:%M")

    # 抓取各類 RSS
    all_articles = {}
    for category, sources in RSS_FEEDS.items():
        print(f"\n抓取 {category}...")
        articles = []
        for source_name, url in sources:
            articles.extend(fetch_rss(source_name, url))
        all_articles[category] = articles

    # 每類別各自呼叫 AI
    categories_result = {}
    for category, articles in all_articles.items():
        print(f"\n處理 {category}（{len(articles)} 篇輸入）...")
        needs_pitch = NEEDS_PITCH.get(category, True)
        selected = select_articles(client, today, category, articles, needs_pitch)
        categories_result[category] = selected
        print(f"  -> 選出 {len(selected)} 篇")

    total = sum(len(v) for v in categories_result.values())
    data = {
        "fetched_at": fetched_at,
        "total": total,
        "categories": categories_result,
    }

    os.makedirs("results", exist_ok=True)
    with open("results/daily_news.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"\ndaily_news.json updated: {today} ({total} 篇，含真實 URL)")


if __name__ == "__main__":
    generate_news()
