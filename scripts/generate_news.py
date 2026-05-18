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
    ],
    "房地產新聞": [
        ("住展房屋網",     "https://www.myhousing.com.tw/feed"),
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


def fetch_rss(source_name, url, max_items=15):
    try:
        r = requests.get(url, headers=HEADERS, timeout=12)
        r.raise_for_status()
        root = ET.fromstring(r.content)

        items = []
        # RSS 2.0
        for item in root.findall(".//item")[:max_items]:
            title = (item.findtext("title") or "").strip()
            # <link> in RSS 2.0 is sometimes a text node, sometimes after a CDATA
            link = (item.findtext("link") or "").strip()
            if not link:
                link_el = item.find("link")
                if link_el is not None:
                    link = (link_el.text or "").strip()
            desc = (item.findtext("description") or "").strip()[:200]
            if title and link:
                items.append({"title": title, "url": link, "desc": desc, "source": source_name})

        # Atom
        if not items:
            ns = {"a": "http://www.w3.org/2005/Atom"}
            for entry in root.findall(".//a:entry", ns)[:max_items]:
                title = (entry.findtext("a:title", "", ns) or "").strip()
                link_el = entry.find("a:link", ns)
                link = (link_el.get("href", "") if link_el is not None else "").strip()
                desc = (entry.findtext("a:summary", "", ns) or "").strip()[:200]
                if title and link:
                    items.append({"title": title, "url": link, "desc": desc, "source": source_name})

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
        if a['desc']:
            lines.append(f"   摘要: {a['desc']}")
    return "\n".join(lines)


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
        all_articles[category] = articles[:20]

    # 組建提示詞
    sections = []
    for category, articles in all_articles.items():
        if articles:
            sections.append(f"## {category}\n{build_article_list(articles)}")

    articles_text = "\n\n".join(sections) if sections else "（RSS 抓取失敗，無文章）"

    prompt = f"""今天是 {today}。以下是從各大媒體 RSS 抓取的【真實新聞】列表：

{articles_text}

請從上方真實新聞中，為「土城區新成屋」銷售團隊選出最重要的新聞，輸出以下格式 JSON，直接輸出不加說明：

{{
  "fetched_at": "{fetched_at}",
  "total": 0,
  "categories": {{
    "台灣財經新聞": [
      {{
        "title": "使用原文標題，不要修改",
        "summary": "2-3句客觀摘要",
        "source": "媒體名稱",
        "url": "直接複製上方列表中的原始URL，不可修改或自行生成",
        "credibility": "⭐⭐⭐⭐⭐",
        "analysis": "對房市或景氣的影響分析",
        "sales_pitch": "口語化銷售說詞（1-2句，連結土城青埔特區新成屋優勢，像朋友建議而非廣告）"
      }}
    ],
    "房地產新聞": [
      {{
        "title": "...", "summary": "...", "source": "...", "url": "...",
        "credibility": "...", "analysis": "...", "sales_pitch": "..."
      }}
    ],
    "中國財經經濟新聞": [
      {{
        "title": "...", "summary": "...", "source": "...", "url": "...",
        "credibility": "...", "analysis": "..."
      }}
    ],
    "國際重大新聞": [
      {{
        "title": "...", "summary": "...", "source": "...", "url": "...",
        "credibility": "...", "analysis": "...", "sales_pitch": "..."
      }}
    ]
  }}
}}

重要規則：
1. url 必須直接複製上方提供的原始網址，絕對不可自行生成、猜測或修改
2. 若某類別沒有文章，填入空陣列 []
3. 每類選 2-4 則
4. 中國財經新聞不需要 sales_pitch
5. sales_pitch 口語化，像朋友在聊天，不要書面廣告腔
6. total 填入所有類別文章總數"""

    for attempt in range(MAX_RETRIES):
        try:
            text = call_api(client, prompt, max_tokens=4000)
            data = parse_json(text)
            break
        except (json.JSONDecodeError, RuntimeError) as e:
            print(f"Failed (attempt {attempt + 1}/{MAX_RETRIES}): {e}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(2 ** attempt)
            else:
                raise

    data["total"] = sum(len(v) for v in data["categories"].values())

    os.makedirs("results", exist_ok=True)
    with open("results/daily_news.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"\ndaily_news.json updated: {today} ({data['total']} 篇，含真實 URL)")


if __name__ == "__main__":
    generate_news()
