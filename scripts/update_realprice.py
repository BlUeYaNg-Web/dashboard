#!/usr/bin/env python3
"""
每週日 13:00 自動執行（Windows Task Scheduler）
從 Google Sheets 讀取實價登錄資料，更新 pptx_update.json 並 git push
"""
import json
import os
import subprocess
import sys
from datetime import datetime

import requests

# ── 設定區 ──────────────────────────────────────────────────────────────────
# 請將 Google Sheets 共享連結換成 CSV 匯出格式：
# 原始連結：https://docs.google.com/spreadsheets/d/試算表ID/edit
# CSV 連結：https://docs.google.com/spreadsheets/d/試算表ID/export?format=csv&gid=工作表ID
GOOGLE_SHEETS_CSV_URL = os.environ.get(
    "SHEETS_URL",
    "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/export?format=csv&gid=0"
)

PPTX_DIR = r"I:\我的雲端硬碟\家泰嘉潤丰藏JANU\03.建設公司\01.業主回報\週報"
RESULT_FILE = "results/pptx_update.json"
# ──────────────────────────────────────────────────────────────────────────────


def fetch_sheets_data(url: str) -> list[dict]:
    """下載 Google Sheets CSV，回傳 list of dict（每行一筆）"""
    resp = requests.get(url, timeout=30)
    resp.raise_for_status()
    resp.encoding = "utf-8-sig"

    lines = resp.text.splitlines()
    if not lines:
        raise ValueError("Google Sheets 回傳空資料")

    headers = [h.strip() for h in lines[0].split(",")]
    rows = []
    for line in lines[1:]:
        if not line.strip():
            continue
        values = [v.strip() for v in line.split(",")]
        row = dict(zip(headers, values))
        rows.append(row)
    return rows


def parse_float(val) -> float | None:
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


def parse_int(val) -> int | None:
    try:
        return int(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


def build_cases(rows: list[dict]) -> list[dict]:
    """
    將 Sheets 資料轉換成 pptx_update.json 所需格式。

    預期 CSV 欄位（可依實際 Sheet 調整）：
      案名, 本週新增, 最新單價(萬/坪), 最新面積(坪), 乾旱週數
    """
    cases = []
    for row in rows:
        name = row.get("案名") or row.get("name", "").strip()
        if not name:
            continue
        cases.append({
            "name": name,
            "weekly_new": parse_int(row.get("本週新增") or row.get("weekly_new")),
            "latest_price": parse_float(row.get("最新單價(萬/坪)") or row.get("latest_price")),
            "latest_area": parse_float(row.get("最新面積(坪)") or row.get("latest_area")),
            "weeks_dry": parse_int(row.get("乾旱週數") or row.get("weeks_dry")),
        })
    return cases


def find_latest_pptx() -> str | None:
    """找到 PPTX_DIR 裡最新的 .pptx 檔"""
    if not os.path.isdir(PPTX_DIR):
        return None
    files = [
        os.path.join(PPTX_DIR, f)
        for f in os.listdir(PPTX_DIR)
        if f.lower().endswith(".pptx")
    ]
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def git_push():
    """add → commit → push"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    repo_root = os.path.dirname(script_dir)

    subprocess.run(["git", "add", RESULT_FILE], cwd=repo_root, check=True)

    status = subprocess.run(
        ["git", "diff", "--staged", "--quiet"],
        cwd=repo_root,
    )
    if status.returncode == 0:
        print("沒有變更，不需要 commit。")
        return

    now_str = datetime.now().strftime("%Y/%m/%d %H:%M")
    subprocess.run(
        ["git", "commit", "-m", f"Update results/pptx_update.json {now_str}"],
        cwd=repo_root,
        check=True,
    )
    subprocess.run(["git", "push"], cwd=repo_root, check=True)
    print("已成功推送到 GitHub。")


def main():
    now = datetime.now()
    now_str = now.strftime("%Y/%m/%d %H:%M")
    print(f"[{now_str}] 開始更新實價登錄資料...")

    # 1. 抓 Google Sheets 資料
    try:
        rows = fetch_sheets_data(GOOGLE_SHEETS_CSV_URL)
        print(f"  → 讀取到 {len(rows)} 筆資料")
    except Exception as e:
        print(f"  ✗ 無法讀取 Google Sheets：{e}")
        print("  請確認 GOOGLE_SHEETS_CSV_URL 設定正確，且試算表已設為「任何人可檢視」")
        sys.exit(1)

    # 2. 轉換格式
    cases = build_cases(rows)
    if not cases:
        print("  ✗ 資料格式不符，無法解析案名。請確認 CSV 欄位名稱。")
        sys.exit(1)

    # 3. 找最新 PPTX
    pptx_path = find_latest_pptx() or ""
    if pptx_path:
        print(f"  → 最新 PPTX：{os.path.basename(pptx_path)}")
    else:
        print(f"  ! 找不到 PPTX（{PPTX_DIR}），僅更新 JSON。")

    # 4. 寫入 JSON
    os.makedirs("results", exist_ok=True)
    data = {
        "run_at": now_str,
        "output_file": pptx_path,
        "cases": cases,
    }
    with open(RESULT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"  → {RESULT_FILE} 已更新")

    # 5. Git push
    try:
        git_push()
    except subprocess.CalledProcessError as e:
        print(f"  ✗ git push 失敗：{e}")
        print("  請確認已設定 git 遠端，且有推送權限。")
        sys.exit(1)

    print("完成！")


if __name__ == "__main__":
    main()
