#!/usr/bin/env python3
"""
update_realprice.py
每週日 13:00 在本機執行：
1. 從 dualtaipei.datazen.info/newtaipei 抓取各建案累積交易筆數
2. 與上週 pptx_update.json 比對，計算本週新增
3. 更新 PPTX（本週新增 N 戶）
4. 儲存新檔名（含當週日期）
5. 更新 results/pptx_update.json 並 git push
"""

import json
import os
import re
import subprocess
import sys
import glob
from datetime import datetime
from urllib.parse import quote

import requests
from pptx import Presentation

# ===== 路徑設定 =====
DATAZEN_BASE = "https://dualtaipei.datazen.info"
DATAZEN_CITY = "newtaipei"

PPTX_DIR = r"I:\我的雲端硬碟\家泰嘉潤丰藏JANU\03.建設公司\01.業主回報\週報"
PPTX_BASENAME = "土城區實價登錄明細"

REPO_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JSON_PATH = os.path.join(REPO_DIR, "results", "pptx_update.json")

# ===== 建案對應 =====
# (datazen 建案名稱, JSON/PPTX 顯示名稱)
# datazen 名稱為 None 表示尚未上 datazen，保留上週數值
BUILDINGS = [
    ("家泰嘉潤丰藏", "丰藏"),
    ("台信琢岳",     "台信琢岳"),
    ("紅布朗花園",   "紅布朗花園"),
    ("寶佳淳青",     "寶佳淳青"),
    ("松陽馥麗",     "松陽馥麗"),
    ("新濠岳",       "新濠岳"),
    ("朗沐",         "朗沐"),
    ("合康雙匯",     "合康雙匯"),
    (None,           "合耀美家雅居"),   # 尚未上 datazen
    ("迴東騰",       "迴東騰"),
    (None,           "天好運"),         # 尚未上 datazen
    ("若水秧翠",     "若水秧翠"),
]

MONTH_CN = {
    "一月": 1, "二月": 2, "三月": 3, "四月": 4,
    "五月": 5, "六月": 6, "七月": 7, "八月": 8,
    "九月": 9, "十月": 10, "十一月": 11, "十二月": 12,
}


def get_datazen_build_id() -> str:
    """從 datazen 首頁的 __NEXT_DATA__ 取得 Next.js buildId"""
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    r = requests.get(f"{DATAZEN_BASE}/{DATAZEN_CITY}", headers=headers, timeout=20)
    r.raise_for_status()
    m = re.search(r'"buildId":"([^"]+)"', r.text)
    if not m:
        raise ValueError("無法從 datazen 取得 buildId")
    return m.group(1)


def fetch_datazen_project(build_id: str, name: str) -> dict | None:
    """
    抓取特定建案的累積交易資料。
    回傳 dict 或 None（建案不存在）。
    """
    url = (f"{DATAZEN_BASE}/_next/data/{build_id}"
           f"/{DATAZEN_CITY}/project/{quote(name)}.json")
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    r = requests.get(url, headers=headers, timeout=20)
    r.raise_for_status()
    project = r.json().get("pageProps", {}).get("project")
    if not project or not project.get("name"):
        return None
    return {
        "transactionCount": project.get("transactionCount", 0),
        "avgUnitPrice":     project.get("avgUnitPrice"),
        "avgArea":          project.get("avgArea"),
        "lastTransaction":  project.get("lastTransaction"),
    }


def calc_stats(datazen_data, prev_total, prev_weeks_dry,
               prev_latest_price, prev_latest_area):
    """
    根據 datazen 資料與上週數值計算本週統計。
    datazen_data 為 None 時（建案尚未上 datazen）保留上週數值。
    """
    if datazen_data is None:
        # 建案尚未上 datazen，維持上週數值並累加乾旱週數
        return {
            "weekly_new":    0,
            "current_total": prev_total or 0,
            "latest_price":  prev_latest_price,
            "latest_area":   prev_latest_area,
            "weeks_dry":     (prev_weeks_dry or 0) + 1,
        }

    current_total = datazen_data["transactionCount"]

    # 防呆：筆數下降視為資料異常，保留上週
    if prev_total and prev_total > 0 and current_total < prev_total:
        print(f"    ⚠ 資料異常（{prev_total} → {current_total}），保留上週筆數")
        current_total = prev_total
        weekly_new = 0
    else:
        weekly_new = max(0, current_total - (prev_total or 0))

    weeks_dry = 0 if weekly_new > 0 else (prev_weeks_dry or 0) + 1

    return {
        "weekly_new":    weekly_new,
        "current_total": current_total,
        "latest_price":  datazen_data["avgUnitPrice"],
        "latest_area":   datazen_data["avgArea"],
        "weeks_dry":     weeks_dry,
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

    # 1. 取得 datazen buildId
    print("取得 datazen buildId...")
    try:
        build_id = get_datazen_build_id()
        print(f"  buildId: {build_id}")
    except Exception as e:
        print(f"  失敗：{e}")
        sys.exit(1)

    # 2. 讀取 PPTX
    print(f"尋找 PPTX 於：{PPTX_DIR}")
    try:
        src_path = find_latest_pptx(PPTX_DIR, PPTX_BASENAME)
        print(f"  來源：{os.path.basename(src_path)}")
    except FileNotFoundError as e:
        print(f"  {e}")
        sys.exit(1)

    prs = Presentation(src_path)

    # 4. 讀取上週資料（total_records / weeks_dry / latest_price / latest_area）
    prev_totals = {}
    prev_weeks_dry_map = {}
    prev_latest_price_map = {}
    prev_latest_area_map = {}
    try:
        with open(JSON_PATH, encoding="utf-8") as f:
            prev_json = json.load(f)
        for c in prev_json.get("cases", []):
            n = c["name"]
            prev_totals[n] = c.get("total_records")
            prev_weeks_dry_map[n] = c.get("weeks_dry")
            prev_latest_price_map[n] = c.get("latest_price")
            prev_latest_area_map[n] = c.get("latest_area")
        print(f"  讀取上週筆數：{prev_totals}")
    except Exception:
        print("  無上週資料，本週新增將顯示 0")

    # 5. 逐建案抓 datazen + 更新 PPTX
    cases = []
    for datazen_name, display in BUILDINGS:
        prev_total = prev_totals.get(display)

        # 從 datazen 取資料
        datazen_data = None
        if datazen_name:
            try:
                datazen_data = fetch_datazen_project(build_id, datazen_name)
                if datazen_data is None:
                    print(f"  {display}: datazen 查無此建案")
            except Exception as e:
                print(f"  {display}: datazen 取得失敗 ({e})，保留上週數值")

        stats = calc_stats(
            datazen_data,
            prev_total,
            prev_weeks_dry=prev_weeks_dry_map.get(display),
            prev_latest_price=prev_latest_price_map.get(display),
            prev_latest_area=prev_latest_area_map.get(display),
        )
        print(f"  {display}: 本週新增 {stats['weekly_new']} 戶 "
              f"(上週 {prev_total} → 本週 {stats['current_total']}) | "
              f"均價 {stats['latest_price']} 萬/坪 | 乾旱 {stats['weeks_dry']} 週")

        # 更新 PPTX「本週新增 N 戶」
        for slide in prs.slides:
            slide_text = " ".join(
                s.text_frame.text for s in slide.shapes if s.has_text_frame
            )
            if display in slide_text or (datazen_name and datazen_name in slide_text):
                update_weekly_text(slide, stats["weekly_new"])

        cases.append({
            "name":          display,
            "weekly_new":    stats["weekly_new"],
            "total_records": stats["current_total"],
            "latest_price":  stats["latest_price"],
            "latest_area":   stats["latest_area"],
            "weeks_dry":     stats["weeks_dry"],
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
        subprocess.run(["git", "-C", REPO_DIR, "push", "origin", "HEAD:main"], check=True)
        print("Git push 完成 → 網頁自動更新")
    except subprocess.CalledProcessError as e:
        print(f"Git 操作失敗：{e}")


if __name__ == "__main__":
    main()
