#!/usr/bin/env python3
"""
update_realprice.py
每週日 13:00 在本機執行：
1. 從 Google Sheets 抓取最新實價登錄資料
2. 從 dualtaipei.datazen.info 交叉比對
3. 更新 PPTX（本週新增、月份統計、成交價格）
4. 儲存新檔名（含當週日期）
5. 更新 results/pptx_update.json 並 git push
"""

import csv
import json
import os
import re
import subprocess
import sys
import glob
from datetime import datetime, timedelta
from io import StringIO

import requests
from pptx import Presentation

# ===== 路徑設定 =====
SHEETS_ID  = "1obKJ7DT__0MeclJGJc7cNzQ8eYLz6T2kX5JdcrVsINs"
SHEETS_GID = "337874303"
# 試三種 URL 格式（不需 API key，但試算表需設為「知道連結的人可以查看」）
SHEETS_CSV_URLS = [
    f"https://docs.google.com/spreadsheets/d/{SHEETS_ID}/gviz/tq?tqx=out:csv&gid={SHEETS_GID}",
    f"https://docs.google.com/spreadsheets/d/{SHEETS_ID}/export?format=csv&gid={SHEETS_GID}",
    f"https://docs.google.com/spreadsheets/d/{SHEETS_ID}/export?format=csv&id={SHEETS_ID}&gid={SHEETS_GID}",
]
DATAZEN_URL = "https://dualtaipei.datazen.info/"

PPTX_DIR = r"I:\我的雲端硬碟\家泰嘉潤丰藏JANU\03.建設公司\01.業主回報\週報"
PPTX_BASENAME = "土城區實價登錄明細"

REPO_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_PATH = os.path.join(REPO_DIR, "results", "pptx_update.json")

# ===== 建案對應 =====
# (Google Sheets 關鍵字, PPTX 顯示名稱)
BUILDINGS = [
    ("丰藏",    "丰藏"),
    ("台信琢岳", "台信琢岳"),
    ("紅布朗",  "紅布朗花園"),
    ("寶佳淳青", "寶佳淳青"),
    ("松陽馥麗", "松陽馥麗"),
    ("新濠岳",  "新濠岳"),
    ("朗沐",    "朗沐"),
    ("合康雙匯", "合康雙匯"),
    ("和耀美家", "和耀美家雅居"),
    ("迴東騰",  "迴東騰"),
    ("天好運",  "天好運"),
    ("若水秧翠", "若水秧翠"),
]

MONTH_CN = {
    "一月": 1, "二月": 2, "三月": 3, "四月": 4,
    "五月": 5, "六月": 6, "七月": 7, "八月": 8,
    "九月": 9, "十月": 10, "十一月": 11, "十二月": 12,
}


def roc_to_date(roc_str):
    """民國日期字串 YYYMMDD → datetime，失敗回傳 None"""
    try:
        s = str(int(float(str(roc_str).strip())))
        if len(s) == 7:
            return datetime(int(s[:3]) + 1911, int(s[3:5]), int(s[5:7]))
    except Exception:
        pass
    return None


def parse_price(price_str):
    """'87.38萬' → 87.38，失敗回傳 None"""
    m = re.match(r"([\d.]+)萬", str(price_str).strip())
    return round(float(m.group(1)), 2) if m else None


def fetch_sheets(urls):
    """
    嘗試多個 URL 格式下載 Google Sheets CSV。
    試算表必須設為「知道連結的人可以查看」才能免登入存取。
    跳過前 2 列（metadata + 欄位標題）。
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept-Language": "zh-TW,zh;q=0.9",
        "Accept": "text/csv,text/plain,*/*",
    }
    last_err = None
    for url in urls:
        try:
            r = requests.get(url, headers=headers, timeout=30, allow_redirects=True)
            if r.status_code == 200:
                content = r.text.strip()
                if content and len(content) > 100:
                    r.encoding = "utf-8-sig"
                    rows = list(csv.reader(StringIO(r.text)))
                    # gviz 格式第 0 列 = 欄位標題；export 格式第 0-1 列可能含 metadata
                    # 找到實際資料起始列（第一個有 7 位數字的民國日期欄位）
                    data_rows = []
                    for row in rows:
                        if len(row) >= 2:
                            val = row[1].strip()
                            if re.match(r"^\d{7}$", val):
                                data_rows.append(row)
                            elif data_rows:  # 已開始收集資料，遇到非日期列繼續
                                data_rows.append(row)
                    if data_rows:
                        print(f"  成功：{url.split('?')[0].split('/')[-1]}")
                        return data_rows
                    # fallback：跳過前 2 列
                    rows = list(csv.reader(StringIO(r.text)))
                    return rows[2:] if len(rows) > 2 else []
            last_err = f"HTTP {r.status_code}"
        except Exception as e:
            last_err = str(e)
        print(f"  嘗試失敗（{last_err}）：{url[:70]}...")
    raise RuntimeError(
        f"無法取得 Google Sheets 資料：{last_err}\n"
        "請確認試算表已設為「知道連結的人可以查看」\n"
        "（Google Sheets → 共用 → 變更為知道連結的人）"
    )


def fetch_datazen():
    """
    從 dualtaipei.datazen.info 抓取交叉比對資料。
    回傳 dict: { 建案關鍵字: { '月': count, ... } }
    若無法取得則回傳空 dict（不中斷主流程）。
    """
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        r = requests.get(DATAZEN_URL, headers=headers, timeout=15)
        r.raise_for_status()
        # TODO: 解析 datazen 頁面結構（依實際 HTML 調整）
        # 目前先回傳空 dict，由 Google Sheets 資料為主
        print("  [datazen] 取得資料（待解析）")
    except Exception as e:
        print(f"  [datazen] 無法取得：{e}（改以 Google Sheets 為主）")
    return {}


def get_building_stats(rows, keyword, prev_total):
    """
    計算建案統計資料。
    本週新增 = 本週檔案總筆數 - 上週檔案總筆數（prev_total）

    欄位索引（0-based）：
      0=執行日期, 1=登錄日期, 2=所在區域, 3=建案名稱,
      4=戶別樓層, 5=平均單價, 9=坪數（若有）
    """
    monthly = {}
    all_txs = []

    for row in rows:
        if len(row) < 6:
            continue
        name = row[3].strip() if len(row) > 3 else ""
        if keyword not in name:
            continue

        reg_date = roc_to_date(row[1]) if len(row) > 1 else None
        if not reg_date:
            continue

        price = parse_price(row[5]) if len(row) > 5 else None
        floor_unit = row[4].strip() if len(row) > 4 else ""
        area_raw = row[9].strip() if len(row) > 9 else ""
        area = None
        area_m = re.match(r"([\d.]+)", area_raw)
        if area_m:
            area = round(float(area_m.group(1)), 2)

        monthly[reg_date.month] = monthly.get(reg_date.month, 0) + 1
        all_txs.append({"date": reg_date, "floor_unit": floor_unit,
                         "price": price, "area": area})

    current_total = len(all_txs)

    # 本週新增 = 本週總筆數 - 上週總筆數
    weekly_new = max(0, current_total - prev_total) if prev_total is not None else 0

    # 本週新增的交易（取最新的 weekly_new 筆，即尾端新增部分）
    all_txs_sorted = sorted(all_txs, key=lambda x: x["date"])
    weekly_txs = all_txs_sorted[-weekly_new:] if weekly_new > 0 else []

    # 最新成交
    latest = all_txs_sorted[-1] if all_txs_sorted else {}
    latest_price = latest.get("price")
    latest_area = latest.get("area")

    # 乾旱週數
    if weekly_new > 0:
        weeks_dry = 0
    elif all_txs_sorted:
        last_date = all_txs_sorted[-1]["date"]
        weeks_dry = max(0, (datetime.now() - last_date).days // 7)
    else:
        weeks_dry = None

    return {
        "weekly_new": weekly_new,
        "current_total": current_total,
        "monthly": monthly,
        "latest_price": latest_price,
        "latest_area": latest_area,
        "weeks_dry": weeks_dry,
        "weekly_txs": weekly_txs,
    }


def update_weekly_text(slide, weekly_new):
    """
    更新投影片上「本週新增 N 戶」。
    原始 PPTX 的文字分 4 個 run：
      run0: '建案名 '  run1: '本週新增 '  run2: 數字  run3: ' 戶'
    只更新 run2 的數字文字與顏色。
    有新增 → 紅色；無新增 → 還原為黑色。
    """
    from pptx.dml.color import RGBColor
    from pptx.oxml.ns import qn
    RED   = RGBColor(0xFF, 0x00, 0x00)
    BLACK = RGBColor(0x00, 0x00, 0x00)

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        full_text = shape.text_frame.text
        if "本週新增" not in full_text or "戶" not in full_text:
            continue
        for para in shape.text_frame.paragraphs:
            runs = para.runs
            # 找「本週新增」run 的下一個 run（即數字 run）
            for i, run in enumerate(runs):
                if "本週新增" in run.text and i + 1 < len(runs):
                    num_run = runs[i + 1]
                    num_run.text = str(weekly_new)
                    num_run.font.color.rgb = RED if weekly_new > 0 else BLACK
                    break
            else:
                # fallback：整段文字在同一個 run
                for run in runs:
                    if "本週新增" in run.text and "戶" in run.text:
                        run.text = re.sub(r"本週新增\s*\d+\s*戶",
                                          f"本週新增 {weekly_new} 戶", run.text)
                        run.font.color.rgb = RED if weekly_new > 0 else BLACK


def update_monthly_shapes(slide, monthly):
    """更新投影片上矩形的月份統計"""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        for cn, num in MONTH_CN.items():
            if text.startswith(cn):
                count = monthly.get(num, 0)
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        run.text = re.sub(
                            rf"({cn}[：:]\s*)\d+",
                            rf"\g<1>{count}",
                            run.text,
                        )
                break


def _get_table_font_template(table):
    """
    從表格中找第一個有內容的價格格子，取出字體設定作為範本。
    回傳 (font_name, font_size, bold, para_alignment)
    """
    from pptx.enum.text import PP_ALIGN
    font_name = "微軟正黑體"
    font_size = None
    bold = None
    alignment = PP_ALIGN.CENTER

    for r, row in enumerate(table.rows):
        if r < 2:
            continue  # 跳過標題列
        for c, cell in enumerate(row.cells):
            if c == 0:
                continue  # 跳過樓層欄
            text = cell.text.strip()
            if "萬" not in text:
                continue
            tf = cell.text_frame
            if tf.paragraphs:
                alignment = tf.paragraphs[0].alignment or PP_ALIGN.CENTER
                for run in tf.paragraphs[0].runs:
                    font = run.font
                    font_name = font.name or "微軟正黑體"
                    font_size = font.size  # EMU 單位，None 表示繼承
                    bold = font.bold
                    return font_name, font_size, bold, alignment
    return font_name, font_size, bold, alignment


def _write_cell(cell, text, font_name, font_size, bold, alignment, color_rgb=None):
    """
    清空格子並以指定格式寫入文字。
    color_rgb: RGBColor 物件，None 表示繼承原色（不強制設定）。
    """
    from pptx.dml.color import RGBColor

    tf = cell.text_frame
    tf.word_wrap = False

    # 清除所有段落，只保留第一個
    for para in tf.paragraphs[1:]:
        p_elem = para._p
        p_elem.getparent().remove(p_elem)

    para = tf.paragraphs[0]
    para.alignment = alignment

    # 清除現有 runs
    for run in para.runs:
        run._r.getparent().remove(run._r)

    run = para.add_run()
    run.text = text
    run.font.name = font_name
    if font_size is not None:
        run.font.size = font_size
    if bold is not None:
        run.font.bold = bold
    if color_rgb is not None:
        run.font.color.rgb = color_rgb


def update_price_table(slide, weekly_txs):
    """
    根據本週新交易更新成交價格矩陣。
    本週新增的格子字體改為紅色，方便開啟後一眼辨識。
    戶別樓層格式範例：丙/F02-06F/六樓
      → 解析樓層 '6F'、戶別前綴 'F02'
    """
    for shape in slide.shapes:
        if shape.shape_type != 19:  # 19 = TABLE
            continue
        if shape.name in ("表格 10", "表格 7"):  # 跳過建案資訊表
            continue

        table = shape.table
        col_headers = [c.text.strip() for c in table.rows[0].cells]
        row_headers = [table.rows[r].cells[0].text.strip()
                       for r in range(len(table.rows))]

        # 從此表格取得字體範本（確保與同表格其他格子一致）
        from pptx.dml.color import RGBColor
        RED = RGBColor(0xFF, 0x00, 0x00)
        font_name, font_size, bold, alignment = _get_table_font_template(table)

        for tx in weekly_txs:
            floor_unit = tx.get("floor_unit", "")
            price = tx.get("price")
            if not price:
                continue

            # 解析樓層：取數字部分，去除前導零（06 → 6）
            floor_m = re.search(r"(\d+)[Ff]", floor_unit)
            if not floor_m:
                continue
            floor_label = f"{int(floor_m.group(1))}F"

            # 解析戶別：取第二段（以 / 分隔）的字母前綴
            parts = floor_unit.split("/")
            unit_prefix = ""
            if len(parts) >= 2:
                unit_m = re.match(r"([A-Za-z]+\d*)", parts[1].strip())
                if unit_m:
                    unit_prefix = unit_m.group(1).upper()

            # 在表格中找對應位置（完全匹配優先，再試前綴匹配）
            row_idx = next((i for i, h in enumerate(row_headers)
                            if h == floor_label), None)
            col_idx = next(
                (i for i, h in enumerate(col_headers) if h.upper() == unit_prefix),
                next(
                    (i for i, h in enumerate(col_headers)
                     if h.upper().startswith(unit_prefix) and unit_prefix),
                    None,
                ),
            )

            if row_idx is not None and col_idx is not None:
                cell = table.rows[row_idx].cells[col_idx]
                price_text = f"{price}萬"
                _write_cell(cell, price_text, font_name, font_size, bold, alignment,
                            color_rgb=RED)
                print(f"    更新 {floor_label}/{unit_prefix} → {price_text} [紅色]")
            else:
                print(f"    找不到對應格子：{floor_label}/{unit_prefix}")


def find_latest_pptx(pptx_dir, basename):
    """找資料夾中最新的對應 PPTX"""
    pattern = os.path.join(pptx_dir, f"{basename}*.pptx")
    files = glob.glob(pattern)
    if not files:
        raise FileNotFoundError(f"找不到 PPTX：{pattern}")
    return max(files, key=os.path.getmtime)


def main():
    today = datetime.now()
    date_tag = today.strftime("%y%m%d")       # e.g., 260516
    run_at = today.strftime("%Y/%m/%d %H:%M")

    print(f"[{run_at}] 開始更新實價登錄...")

    # 1. Google Sheets 資料
    print("下載 Google Sheets 資料...")
    try:
        rows = fetch_sheets(SHEETS_CSV_URLS)
        print(f"  取得 {len(rows)} 筆")
    except Exception as e:
        print(f"  失敗：{e}")
        sys.exit(1)

    # 2. Datazen 交叉比對（不中斷）
    print("取得 datazen 資料...")
    datazen = fetch_datazen()

    # 3. 讀取 PPTX
    print(f"尋找 PPTX 於：{PPTX_DIR}")
    try:
        src_path = find_latest_pptx(PPTX_DIR, PPTX_BASENAME)
        print(f"  來源：{os.path.basename(src_path)}")
    except FileNotFoundError as e:
        print(f"  {e}")
        sys.exit(1)

    prs = Presentation(src_path)

    # 4. 讀取上週各建案筆數（存於 pptx_update.json 的 prev_totals）
    prev_totals = {}
    try:
        with open(JSON_PATH, encoding="utf-8") as f:
            prev_json = json.load(f)
        for c in prev_json.get("cases", []):
            prev_totals[c["name"]] = c.get("total_records", None)
        print(f"  讀取上週筆數：{prev_totals}")
    except Exception:
        print("  無上週資料，本週新增將顯示 0")

    # 5. 逐建案計算 + 更新 PPTX
    cases = []
    for keyword, display in BUILDINGS:
        prev_total = prev_totals.get(display)
        stats = get_building_stats(rows, keyword, prev_total)
        print(f"  {display}: 本週新增 {stats['weekly_new']} 戶 "
              f"(上週 {prev_total} → 本週 {stats['current_total']}) | "
              f"最新 {stats['latest_price']} 萬 | 乾旱 {stats['weeks_dry']} 週")

        # 找所有含此建案名稱的投影片並更新
        for slide in prs.slides:
            slide_text = " ".join(
                s.text_frame.text for s in slide.shapes if s.has_text_frame
            )
            if display in slide_text or keyword in slide_text:
                update_weekly_text(slide, stats["weekly_new"])
                update_monthly_shapes(slide, stats["monthly"])
                if stats["weekly_txs"]:
                    update_price_table(slide, stats["weekly_txs"])

        cases.append({
            "name": display,
            "weekly_new": stats["weekly_new"],
            "total_records": stats["current_total"],   # 存本週總筆數供下週比對
            "latest_price": stats["latest_price"],
            "latest_area": stats["latest_area"],
            "weeks_dry": stats["weeks_dry"],
        })

    # 6. 儲存新檔名
    new_name = f"{PPTX_BASENAME}{date_tag}自動更新.pptx"
    new_path = os.path.join(PPTX_DIR, new_name)
    prs.save(new_path)
    print(f"PPTX 已儲存：{new_name}")

    # 7. 更新 pptx_update.json
    json_data = {
        "run_at": run_at,
        "output_file": new_path,
        "cases": cases,
    }
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)
    print("pptx_update.json 已更新")

    # 7. Git push
    try:
        subprocess.run(["git", "-C", REPO_DIR, "add", "results/pptx_update.json"],
                       check=True)
        subprocess.run(["git", "-C", REPO_DIR, "commit", "-m",
                         f"Update 實價登錄 {run_at}"], check=True)
        subprocess.run(["git", "-C", REPO_DIR, "push"], check=True)
        print("Git push 完成 → 網頁自動更新")
    except subprocess.CalledProcessError as e:
        print(f"Git 操作失敗：{e}")


if __name__ == "__main__":
    main()
