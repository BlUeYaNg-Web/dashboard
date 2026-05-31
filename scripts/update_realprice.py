#!/usr/bin/env python3
"""
全量同步實價登錄 → PPTX
- Google Sheet 每一筆都對照 PPTX 格子
- 格子空白 or 值不同 → 填入，標紅色
- 格子已有正確值 → 不動
- 本週最新登錄（2026/05/01 之後）優先標紅，其餘補填也標紅（讓你一眼看出差異）
"""
import sys, csv, json, re, os, glob, base64, urllib.request
import requests
from io import StringIO
from datetime import datetime
from pptx import Presentation
from pptx.dml.color import RGBColor
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

SHEET_ID  = "1obKJ7DT__0MeclJGJc7cNzQ8eYLz6T2kX5JdcrVsINs"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv"
PPTX_DIR  = r"I:\我的雲端硬碟\家泰嘉潤丰藏JANU\03.建設公司\01.業主回報\週報"
PPTX_BASE = "土城區實價登錄明細"
cfg   = json.loads(open(r'C:\Users\dxb71\.github_dashboard.json').read())
TOKEN = cfg['token']; REPO = cfg['repo']

BUILDINGS = [
    ("丰藏","丰藏"),("台信琢岳","台信琢岳"),("紅布朗","紅布朗花園"),
    ("寶佳淳青","寶佳淳青"),("松陽馥麗","松陽馥麗"),("新濠岳","新濠岳"),
    ("朗沐","朗沐"),("合康雙匯","合康雙匯"),("和耀美家","和耀美家雅居"),
    ("迴東騰","迴東騰"),("天好運","天好運3"),("若水秧翠","若水秧翠"),
]
MONTH_CN = {"一月":1,"二月":2,"三月":3,"四月":4,"五月":5,"六月":6,
            "七月":7,"八月":8,"九月":9,"十月":10,"十一月":11,"十二月":12}
CN_TO_INT = {'一':1,'二':2,'三':3,'四':4,'五':5,'六':6,'七':7,'八':8,'九':9,
             '十':10,'十一':11,'十二':12,'十三':13,'十四':14,'十五':15,
             '十六':16,'十七':17,'十八':18,'十九':19,'二十':20}
RED   = RGBColor(0xFF,0x00,0x00)
BLACK = RGBColor(0x00,0x00,0x00)

def roc_to_date(s):
    try:
        s = str(int(float(str(s).strip())))
        if len(s)==7: return datetime(int(s[:3])+1911,int(s[3:5]),int(s[5:7]))
    except: pass
    return None

def parse_price(s):
    m = re.match(r'([\d.]+)萬', str(s).strip())
    return round(float(m.group(1)),2) if m else None

def cell_price(text):
    """取格子現有數值（float），用來做數值比對而非字串比對"""
    m = re.match(r'([\d.]+)萬', str(text).strip())
    return round(float(m.group(1)),2) if m else None

def cn_floor(s):
    s = s.replace('樓','').strip()
    n = CN_TO_INT.get(s)
    if n: return f"{n}F"
    m = re.match(r'(\d+)',s)
    return f"{int(m.group(1))}F" if m else None

def parse_floor_unit(fu, col_hdrs):
    parts = [p.strip() for p in fu.split('/')]
    if len(parts) < 3: return None, None
    p0,p1,p2 = parts[0],parts[1],parts[2]
    floor = cn_floor(p2)
    if not floor:
        fm = re.search(r'(\d+)[Ff]',p1)
        if fm: floor = f"{int(fm.group(1))}F"
        else: return None, None
    col_up = [h.upper() for h in col_hdrs]
    def try_u(c):
        if c and c.upper() in col_up: return col_hdrs[col_up.index(c.upper())]
        return None
    cands = []
    if p0:
        cands.append(p0)
        num_m = re.match(r'^(\d+)',p1)
        if num_m: cands.append(p0+num_m.group(1))
    alpha_m = re.match(r'([A-Za-z]+\d*)',p1)
    if alpha_m and not p1[0].isdigit(): cands.append(alpha_m.group(1))
    for c in cands:
        u = try_u(c)
        if u: return floor, u
    return floor, None

def _get_font(table):
    from pptx.enum.text import PP_ALIGN
    fn,fs,bold,align = '微軟正黑體',None,None,PP_ALIGN.CENTER
    for r,row in enumerate(table.rows):
        if r<2: continue
        for c,cell in enumerate(row.cells):
            if c==0 or '萬' not in cell.text: continue
            tf=cell.text_frame
            if tf.paragraphs:
                align=tf.paragraphs[0].alignment or align
                for run in tf.paragraphs[0].runs:
                    return run.font.name or fn,run.font.size,run.font.bold,align
    return fn,fs,bold,align

def write_cell(cell, text, fn, fs, bold, align, color):
    tf=cell.text_frame; tf.word_wrap=False
    for para in tf.paragraphs[1:]: para._p.getparent().remove(para._p)
    para=tf.paragraphs[0]; para.alignment=align
    for run in para.runs: run._r.getparent().remove(run._r)
    run=para.add_run(); run.text=text
    run.font.name=fn
    if fs: run.font.size=fs
    if bold is not None: run.font.bold=bold
    run.font.color.rgb=color

def sync_to_slide(slide, all_txs, new_cutoff):
    """
    new_cutoff: datetime，登錄日期 >= 此日的視為「本週新增」→ 強制標紅
    舊資料：格子空白才補填（標紅）；格子已有正確值則不動
    """
    filled = 0
    new_count = 0
    for shape in slide.shapes:
        if shape.shape_type!=19: continue
        if shape.name in ('表格 10','表格 7'): continue
        table=shape.table
        col_hdrs=[c.text.strip() for c in table.rows[0].cells]
        row_hdrs=[table.rows[r].cells[0].text.strip() for r in range(len(table.rows))]
        fn,fs,bold,align=_get_font(table)
        for tx in all_txs:
            fu=tx.get('floor_unit',''); price=tx.get('price')
            if not price: continue
            fl,unit=parse_floor_unit(fu,col_hdrs)
            if not fl or not unit: continue
            ri=next((i for i,h in enumerate(row_hdrs) if h==fl),None)
            ci=next((i for i,h in enumerate(col_hdrs) if h==unit),None)
            if ri is None or ci is None: continue
            cell=table.rows[ri].cells[ci]
            cur=cell.text.strip()
            want=f"{price}萬"
            cur_p=cell_price(cur)     # 格子現值（float 或 None）
            is_new=(tx.get('date') and tx['date']>=new_cutoff)

            if is_new:
                # 本月新增 → 一律強制標紅（不管格子已有黑字還是空白）
                write_cell(cell,want,fn,fs,bold,align,RED)
                status='本月新增[補填]' if not cur else f'本月新增[已有{cur}→標紅]'
                print(f"    ★ {status} [{fl}][{unit}] → {want}")
                filled+=1
                new_count+=1
            elif cur_p==price:
                continue  # 舊資料且格子已有正確值 → 不動
            elif not cur:
                # 舊資料但格子空白 → 補填（標紅，提醒缺漏）
                write_cell(cell,want,fn,fs,bold,align,RED)
                print(f"    ★ 舊資料補填 [{fl}][{unit}] → {want}")
                filled+=1
            # 舊資料但值不同（可能是來源差異）→ 不動，避免誤改
    return filled, new_count

def update_weekly_text(slide, n):
    for shape in slide.shapes:
        if not shape.has_text_frame: continue
        if '本週新增' not in shape.text_frame.text: continue
        for para in shape.text_frame.paragraphs:
            runs=para.runs
            for i,run in enumerate(runs):
                if '本週新增' in run.text and i+1<len(runs):
                    nr=runs[i+1]
                    nr.text=str(n)
                    nr.font.color.rgb=RED if n>0 else BLACK
                    break
            else:
                for run in runs:
                    if '本週新增' in run.text and '戶' in run.text:
                        run.text=re.sub(r'本週新增\s*\d+\s*戶',f'本週新增 {n} 戶',run.text)
                        run.font.color.rgb=RED if n>0 else BLACK

def update_monthly(slide, monthly):
    for shape in slide.shapes:
        if not shape.has_text_frame: continue
        text=shape.text_frame.text.strip()
        for cn,num in MONTH_CN.items():
            if text.startswith(cn):
                count=monthly.get(num,0)
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        run.text=re.sub(rf'({cn}[：:]\s*)\d+',rf'\g<1>{count}',run.text)
                break

def push_json(data, sha):
    url=f'https://api.github.com/repos/{REPO}/contents/results/pptx_update.json'
    b64=base64.b64encode(json.dumps(data,ensure_ascii=False,indent=2).encode()).decode()
    payload=json.dumps({'message':f'Update 實價登錄全量同步 {data["run_at"]}','content':b64,'sha':sha}).encode()
    req=urllib.request.Request(url,data=payload,method='PUT',
        headers={'Authorization':f'token {TOKEN}','Content-Type':'application/json'})
    with urllib.request.urlopen(req) as resp: return json.load(resp)['commit']['sha'][:8]

def main():
    today=datetime.now(); run_at=today.strftime('%Y/%m/%d %H:%M'); date_tag=today.strftime('%y%m%d')
    print(f"[{run_at}] 全量同步實價登錄")

    # 讀 GitHub JSON SHA
    url=f'https://api.github.com/repos/{REPO}/contents/results/pptx_update.json'
    req=urllib.request.Request(url,headers={'Authorization':f'token {TOKEN}'})
    with urllib.request.urlopen(req) as resp: d=json.load(resp)
    json_sha=d['sha']

    # Google Sheet
    print("下載 Google Sheet...")
    r=requests.get(SHEET_URL,headers={'User-Agent':'Mozilla/5.0'},timeout=30)
    r.encoding='utf-8-sig'
    data=list(csv.reader(StringIO(r.text)))[2:]
    print(f"  {len(data)} 筆")

    # 底稿：用最新手動版（非自動更新）
    manual=[f for f in sorted(glob.glob(os.path.join(PPTX_DIR,f'{PPTX_BASE}*.pptx'))) if '自動更新' not in f]
    src=manual[-1] if manual else sorted(glob.glob(os.path.join(PPTX_DIR,f'{PPTX_BASE}*.pptx')))[-1]
    print(f"底稿：{os.path.basename(src)}")
    prs=Presentation(src)

    # 本月第一天作為「新增」基準
    new_cutoff=datetime(today.year,today.month,1)
    print(f"本月新增基準：>= {new_cutoff.strftime('%Y/%m/%d')}")

    cases=[]; grand_total=0; grand_new=0
    for kw,display in BUILDINGS:
        hits=[row for row in data if len(row)>5 and kw in row[3]]
        all_txs=[]
        monthly={}
        for row in hits:
            d2=roc_to_date(row[0])   # row[0]=登錄日期
            if not d2: continue
            p=parse_price(row[5])
            all_txs.append({'date':d2,'price':p,'floor_unit':row[4] if len(row)>4 else ''})
            monthly[d2.month]=monthly.get(d2.month,0)+1
        all_txs.sort(key=lambda x:x['date'])
        # 本月新增筆數（用於本週新增戶數顯示）
        weekly_new_txs=sum(1 for t in all_txs if t['date']>=new_cutoff)
        latest=all_txs[-1] if all_txs else {}
        if all_txs:
            weeks_dry=max(0,(today-all_txs[-1]['date']).days//7)
        else:
            weeks_dry=None

        # 找出此建案的所有相關 slides（含緊接在後、只有月份沒有建案名的延續頁）
        all_bld_names=[b[1] for b in BUILDINGS]+[b[0] for b in BUILDINGS]
        bld_slides=[]
        for si,slide in enumerate(prs.slides):
            st=' '.join(s.text_frame.text for s in slide.shapes if s.has_text_frame)
            if display in st or kw in st:
                bld_slides.append((si,slide))
                # 往後掃：緊接的頁面若無任何建案名稱 且有表格 → 延續頁
                for sj in range(si+1,min(si+3,len(prs.slides))):
                    nxt=prs.slides[sj]
                    nxt_st=' '.join(s.text_frame.text for s in nxt.shapes if s.has_text_frame)
                    if any(nm in nxt_st for nm in all_bld_names): break
                    if any(s.shape_type==19 for s in nxt.shapes):
                        bld_slides.append((sj,nxt))

        # 同步到 PPTX
        bld_filled=0; bld_new=0
        for _,slide in bld_slides:
            cnt,nc=sync_to_slide(slide,all_txs,new_cutoff)
            update_weekly_text(slide,weekly_new_txs)
            update_monthly(slide,monthly)
            bld_filled+=cnt; bld_new+=nc

        if bld_new>0:
            print(f"  ★ {display}: 本月新增 {bld_new} 格[紅色]，補填舊 {bld_filled-bld_new} 格")
        elif bld_filled>0:
            print(f"  ★ {display}: 補填舊資料 {bld_filled} 格[紅色]")
        else:
            print(f"  {display}: 無更動（{len(all_txs)}筆，本月新增{weekly_new_txs}筆）")
        grand_total+=bld_filled; grand_new+=bld_new

        cases.append({'name':display,'weekly_new':weekly_new_txs,
                      'total_records':len(all_txs),
                      'latest_price':latest.get('price'),'latest_area':None,
                      'weeks_dry':weeks_dry})

    # 存 PPTX（先存本機，再複製到雲端硬碟）
    new_name=f'{PPTX_BASE}{date_tag}自動更新.pptx'
    local_path=os.path.join(r'C:\Users\dxb71', new_name)
    prs.save(local_path)
    print(f"\nPPTX 已存本機：{local_path}（共補填/更新 {grand_total} 格）")
    # 複製到 I:\ 雲端硬碟
    import shutil
    new_path=os.path.join(PPTX_DIR, new_name)
    try:
        shutil.copy2(local_path, new_path)
        print(f"PPTX 已複製到雲端：{new_path}")
    except Exception as e:
        print(f"⚠️ 複製到雲端失敗：{e}（本機版本仍可用）")
        new_path=local_path

    # 推送 JSON
    jd={'run_at':run_at,'output_file':new_path,'cases':cases}
    commit=push_json(jd,json_sha)
    print(f"pptx_update.json 推送 GitHub (commit: {commit})")
    print(f"\n=== 完成 === 本月新增標紅 {grand_new} 格，補填舊資料 {grand_total-grand_new} 格")

if __name__=='__main__':
    main()
