"""
engine.py — Deserve Dashboard core builder
Takes raw Excel bytes → returns Excel dashboard bytes
"""

import io, warnings
import pandas as pd
import numpy as np

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gc

warnings.filterwarnings('ignore')

# ─── STYLE CONSTANTS ─────────────────────────────────────────────────────────
T  = Side(style='thin',   color='D0D0D0')
M  = Side(style='medium', color='595959')
BT = Border(left=T, right=T, top=T, bottom=T)

def F(h): return PatternFill("solid", fgColor=h)
def A(h="center", wrap=False): return Alignment(horizontal=h, vertical="center", wrap_text=wrap)

def W(ws, r, c, val, bg=None, fg="000000", b=False, sz=9, fmt=None,
      h="center", wrap=False, border=BT, italic=False, merge_to=None):
    cc = ws.cell(row=r, column=c)
    cc.value = val
    cc.font = Font(name="Arial", bold=b, size=sz, color=fg, italic=italic)
    cc.alignment = Alignment(horizontal=h, vertical="center", wrap_text=wrap)
    cc.border = border
    if bg:  cc.fill = F(bg)
    if fmt: cc.number_format = fmt
    if merge_to:
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=merge_to)
    return cc

MTH = {1:'ม.ค.',2:'ก.พ.',3:'มี.ค.',4:'เม.ย.',5:'พ.ค.',
       6:'มิ.ย.',7:'ก.ค.',8:'ส.ค.',9:'ก.ย.',10:'ต.ค.',11:'พ.ย.',12:'ธ.ค.'}

STORE_CAT_STYLE = {
    '🌟 STAR':       {'bg':'FFF0A0','fg':'7B5E00'},
    '📈 GROWTH':     {'bg':'C6EFCE','fg':'375623'},
    '🔄 RECOVERING': {'bg':'DEEAF1','fg':'1F4E79'},
    '📉 DECLINING':  {'bg':'FCE4D6','fg':'843C0C'},
    '⚠️ WARNING':    {'bg':'FFD7D7','fg':'C00000'},
    '😴 STABLE':     {'bg':'F2F2F2','fg':'595959'},
    '🔵 SPORADIC':   {'bg':'EBF3FB','fg':'1F4E79'},
    '⚡ INACTIVE':   {'bg':'F5F5F5','fg':'AAAAAA'},
}
PAIR_STYLE = {
    'GROWING':  ('C6EFCE','375623','📈 กำลังโต'),
    'STABLE':   ('F2F2F2','595959','➡️ คงที่'),
    'DECLINING':('FCE4D6','843C0C','📉 ลดลง'),
    'DROPPED':  ('FFD7D7','C00000','⛔ หยุดซื้อ'),
    'STOPPED':  ('FFE0B2','7B3F00','🔴 หายไป'),
    'SPORADIC': ('EBF3FB','1F4E79','💤 ไม่สม่ำเสมอ'),
    'ONE_TIME': ('FAFAFA','AAAAAA','1️⃣ ครั้งเดียว'),
}

# ─── DATA LOADING ────────────────────────────────────────────────────────────
def load_data(file_bytes):
    buf = io.BytesIO(file_bytes)
    df  = pd.read_excel(buf, sheet_name='ข้อมูลเดือน 1-5', header=1)
    buf.seek(0)
    c1  = pd.read_excel(buf, sheet_name='เงื่อนไข1', header=None)
    buf.seek(0)
    c2  = pd.read_excel(buf, sheet_name='เงื่อนไข2', header=None)

    df  = df[df['สถานะรายการ'] == 'สำเร็จ'].copy()
    for col in ['ชื่อลูกค้า','รหัสลูกค้า','รหัสสินค้า','ชื่อสินค้า','หมวดหมู่']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    df['date']    = pd.to_datetime(df['วันที่ทำรายการ'], dayfirst=True, errors='coerce')
    df['month']   = df['date'].dt.month
    df['ราคารวม'] = pd.to_numeric(df['ราคารวม'], errors='coerce').fillna(0)
    df['จำนวน']   = pd.to_numeric(df['จำนวน'],   errors='coerce').fillna(0)

    # Cond1: name → store
    n2s = {}; cur = None
    for _, r_ in c1.iterrows():
        s = str(r_[0]).strip() if pd.notna(r_[0]) else ''
        c = str(r_[2]).strip() if pd.notna(r_[2]) else ''
        if s and s not in ('เงื่อนไข1','ชื่อร้าน','nan'): cur = s
        if c and c not in ('ชื่อลูกค้า','nan') and cur: n2s[c] = cur
    df['ชื่อร้าน'] = df['ชื่อลูกค้า'].map(n2s).fillna(df['ชื่อลูกค้า'])

    # Cond2: normalize รหัสลูกค้า
    code_map = {}
    for _, r_ in c2.iterrows():
        c1v = str(r_[2]).strip() if pd.notna(r_[2]) else ''
        c2v = str(r_[3]).strip() if pd.notna(r_[3]) else ''
        if c2v and c2v != 'nan' and c1v and c1v != 'nan': code_map[c2v] = c1v
    df['รหัสลูกค้า_norm'] = df['รหัสลูกค้า'].map(lambda x: code_map.get(x, x))

    # Product grouping
    df['prod_grp'] = df['รหัสสินค้า'].str[:6]
    prod_names = {}; prod_cat = {}
    for grp, gdf in df.groupby('prod_grp'):
        nm = gdf['ชื่อสินค้า'].value_counts()
        for n in nm.index:
            if not any(ch.isdigit() for ch in str(n)[-8:]): prod_names[grp]=n; break
        if grp not in prod_names: prod_names[grp] = nm.index[0]
        prod_cat[grp] = gdf['หมวดหมู่'].mode()[0]

    return df, prod_names, prod_cat

# ─── PIVOT HELPERS ───────────────────────────────────────────────────────────
def make_pivot(df, group_col, months):
    pm = df.groupby([group_col,'month'])['ราคารวม'].sum().unstack(fill_value=0)
    for m in months:
        if m not in pm.columns: pm[m] = 0
    return pm[months]

def _linregress(x, y):
    n = len(x); xm, ym = x.mean(), y.mean()
    ss_xy = ((x-xm)*(y-ym)).sum(); ss_xx = ((x-xm)**2).sum()
    sl = ss_xy/ss_xx if ss_xx != 0 else 0.0; ic = ym - sl*xm
    ss_tot = ((y-ym)**2).sum()
    r2 = max(0.0, 1 - ((y-(sl*x+ic))**2).sum()/ss_tot) if ss_tot != 0 else 0.0
    return sl, ic, r2

def fc_linear(vals, months_act, months_fc):
    x = np.array(months_act, dtype=float)
    sl, ic, r2 = _linregress(x, vals)
    fc = {m: max(0.0, sl*m + ic) for m in months_fc}
    return sl, r2, fc

def wma3(vals):
    return float(max(0, np.average(vals[-3:], weights=[1,2,3])))

# ─── SHEET BUILDERS ──────────────────────────────────────────────────────────

def build_sheet1(wb, df, grand_total, months):
    ws = wb.create_sheet("1_ภาพรวมยอดขาย")
    ws.sheet_properties.tabColor = "1F4E79"
    W(ws,1,1,"Deserve Dashboard – ภาพรวมยอดขาย",
      bg="1F4E79",fg="FFFFFF",b=True,sz=14,h="left",merge_to=6)
    ws.row_dimensions[1].height = 26

    labels = ["ยอดขายรวม (บาท)","จำนวนรายการ","จำนวนชื่อร้าน","จำนวนรหัสลูกค้า"]
    values = [grand_total,
              df['รายการ'].dropna().nunique() if 'รายการ' in df.columns else len(df),
              df['ชื่อร้าน'].nunique(),
              df['รหัสลูกค้า_norm'].nunique()]
    fmts   = ['#,##0.00','#,##0','#,##0','#,##0']
    for i,(lbl,val,fmt) in enumerate(zip(labels,values,fmts)):
        W(ws,3,i+1,lbl,bg="2E75B6",fg="FFFFFF",b=True,sz=9)
        W(ws,4,i+1,val,bg="DEEAF1",b=True,sz=13,fg="1F4E79",fmt=fmt)
    ws.row_dimensions[4].height = 26

    hdrs = ["เดือน","ยอดขาย (บาท)","% ของยอดรวม","จำนวนร้าน","จำนวนรหัส","เพิ่ม/ลด vs เดือนก่อน"]
    for ci,h in enumerate(hdrs,1):
        W(ws,6,ci,h,bg="1F4E79",fg="FFFFFF",b=True,sz=9)
    ws.row_dimensions[6].height = 18

    prev_sales = None
    for ri,m in enumerate(months,7):
        mdf = df[df['month']==m]
        sales = mdf['ราคารวม'].sum()
        pct   = sales/grand_total if grand_total else 0
        chg   = (sales-prev_sales)/prev_sales if prev_sales else None
        alt   = "F2F2F2" if ri%2==0 else None
        W(ws,ri,1,MTH[m],bg=alt)
        W(ws,ri,2,sales,bg=alt,fmt='#,##0.00')
        W(ws,ri,3,pct,bg=alt,fmt='0.0%')
        W(ws,ri,4,mdf['ชื่อร้าน'].nunique(),bg=alt)
        W(ws,ri,5,mdf['รหัสลูกค้า_norm'].nunique(),bg=alt)
        if chg is not None:
            cbg = "E2EFDA" if chg>=0 else "FFD7D7"
            W(ws,ri,6,chg,bg=cbg,fmt='+0.0%;-0.0%;0.0%')
        prev_sales = sales

    tr = 7+len(months)
    W(ws,tr,1,"รวม",bg="BDD7EE",b=True)
    W(ws,tr,2,grand_total,bg="2E75B6",fg="FFFFFF",b=True,fmt='#,##0.00')
    W(ws,tr,3,1.0,bg="BDD7EE",b=True,fmt='0.0%')

    ws.column_dimensions['A'].width=10; ws.column_dimensions['B'].width=18
    ws.column_dimensions['C'].width=14; ws.column_dimensions['D'].width=14
    ws.column_dimensions['E'].width=16; ws.column_dimensions['F'].width=20


def build_sheet2(wb, df, grand_total, months):
    ws = wb.create_sheet("2_Ranking_ร้านค้า")
    ws.sheet_properties.tabColor = "2E75B6"
    W(ws,1,1,"Ranking ยอดขายรายเดือน – ชื่อร้าน",
      bg="1F4E79",fg="FFFFFF",b=True,sz=12,h="left",merge_to=4+len(months))

    sm = make_pivot(df,'ชื่อร้าน',months)
    sm['รวม'] = sm.sum(axis=1)
    sm['เฉลี่ย/เดือน'] = sm[months].mean(axis=1)
    sm['%ยอดรวม'] = sm['รวม']/grand_total
    sm = sm.sort_values('รวม',ascending=False)

    hdrs = ['อันดับ','ชื่อร้าน']+[MTH[m] for m in months]+['รวม (บาท)','เฉลี่ย/เดือน','%ยอดรวม','สถานะ']
    for ci,h in enumerate(hdrs,1):
        W(ws,3,ci,h,bg="1F4E79",fg="FFFFFF",b=True,sz=8,wrap=True)
    ws.row_dimensions[3].height = 30

    min_m = sm[months]
    sm['_drops'] = min_m.apply(
        lambda r_: sum(1 for i in range(1,len(r_)) if r_.iloc[i]<r_.iloc[i-1] and r_.iloc[i-1]>0), axis=1)

    for ri_off,(store,row) in enumerate(sm.iterrows(),1):
        ri = ri_off+3
        ws.row_dimensions[ri].height = 15
        alt = "F0F4FF" if ri_off%2==0 else None
        W(ws,ri,1,ri_off,bg=alt,b=(ri_off<=10))
        W(ws,ri,2,store,bg=alt,h="left",b=(ri_off<=5),sz=9)
        for ci,m in enumerate(months,3):
            v = row[m]
            W(ws,ri,ci,v,bg="FFF0F0" if v==0 else alt,fmt='#,##0',sz=8)
        W(ws,ri,3+len(months),row['รวม'],bg="BDD7EE",b=True,fmt='#,##0.00')
        W(ws,ri,4+len(months),row['เฉลี่ย/เดือน'],bg=alt,fmt='#,##0',sz=8)
        W(ws,ri,5+len(months),row['%ยอดรวม'],bg=alt,fmt='0.00%',sz=8)
        drops = row['_drops']
        sbg = "FFD7D7" if drops>=2 else ("FFF2CC" if drops==1 else "E2EFDA")
        slbl = "⚠️ เฝ้าระวัง" if drops>=2 else ("📊 ปกติ" if drops==1 else "📈 ดี")
        W(ws,ri,6+len(months),slbl,bg=sbg,b=True,sz=8)

    tr = len(sm)+4
    W(ws,tr,2,"รวมทั้งหมด",bg="2E75B6",fg="FFFFFF",b=True,h="left")
    for ci,m in enumerate(months,3):
        W(ws,tr,ci,df[df['month']==m]['ราคารวม'].sum(),bg="2E75B6",fg="FFFFFF",b=True,fmt='#,##0')
    W(ws,tr,3+len(months),grand_total,bg="1F4E79",fg="FFFFFF",b=True,fmt='#,##0.00')

    ws.column_dimensions['A'].width=6; ws.column_dimensions['B'].width=28
    for i in range(3,3+len(months)+5): ws.column_dimensions[gc(i)].width=12
    ws.freeze_panes="C4"


def build_sheet3(wb, df, grand_total, months):
    ws = wb.create_sheet("3_Ranking_รหัสลูกค้า")
    ws.sheet_properties.tabColor = "375623"
    W(ws,1,1,"Ranking ยอดขายรายเดือน – รหัสลูกค้า",
      bg="1F4E79",fg="FFFFFF",b=True,sz=12,h="left",merge_to=5+len(months))

    cm = make_pivot(df,'รหัสลูกค้า_norm',months)
    cm['รวม'] = cm.sum(axis=1)
    cm['เฉลี่ย/เดือน'] = cm[months].mean(axis=1)
    cm['%ยอดรวม'] = cm['รวม']/grand_total
    code_store = df.groupby('รหัสลูกค้า_norm')['ชื่อร้าน'].agg(lambda x: x.mode()[0])
    cm['ชื่อร้าน'] = cm.index.map(code_store)
    cm = cm.sort_values('รวม',ascending=False)

    hdrs = ['อันดับ','ชื่อร้าน','รหัสลูกค้า']+[MTH[m] for m in months]+['รวม','เฉลี่ย','%ยอดรวม']
    for ci,h in enumerate(hdrs,1):
        W(ws,3,ci,h,bg="1F4E79",fg="FFFFFF",b=True,sz=8,wrap=True)
    ws.row_dimensions[3].height=30

    for ri_off,(code,row) in enumerate(cm.iterrows(),1):
        ri=ri_off+3; alt="F0F4FF" if ri_off%2==0 else None
        ws.row_dimensions[ri].height=14
        W(ws,ri,1,ri_off,bg=alt,b=(ri_off<=10))
        W(ws,ri,2,row.get('ชื่อร้าน',''),bg=alt,h="left",sz=8)
        W(ws,ri,3,code,bg=alt,h="left",sz=8)
        for ci,m in enumerate(months,4):
            v=row[m]; W(ws,ri,ci,v,bg="FFF0F0" if v==0 else alt,fmt='#,##0',sz=8)
        W(ws,ri,4+len(months),row['รวม'],bg="BDD7EE",b=True,fmt='#,##0')
        W(ws,ri,5+len(months),row['เฉลี่ย/เดือน'],bg=alt,fmt='#,##0',sz=8)
        W(ws,ri,6+len(months),row['%ยอดรวม'],bg=alt,fmt='0.00%',sz=8)

    ws.column_dimensions['A'].width=6; ws.column_dimensions['B'].width=26
    ws.column_dimensions['C'].width=30
    for i in range(4,4+len(months)+4): ws.column_dimensions[gc(i)].width=12
    ws.freeze_panes="D4"


def build_sheet4(wb, df, grand_total, months):
    ws = wb.create_sheet("4_ยอดขายฟาร์ม")
    ws.sheet_properties.tabColor = "375623"
    farm_df = df[df['รหัสลูกค้า'].str.startswith('ฟาร์ม')].copy()
    farm_total = farm_df['ราคารวม'].sum()
    num_farms  = farm_df['รหัสลูกค้า'].nunique()

    W(ws,1,1,"ยอดขายลูกค้าฟาร์ม",bg="375623",fg="FFFFFF",b=True,sz=13,h="left",
      merge_to=4+len(months))
    W(ws,2,1,f"จำนวน {num_farms} ฟาร์ม | ยอดรวม {farm_total:,.2f} บาท | {farm_total/grand_total:.2%} ของยอดทั้งหมด",
      bg="E2EFDA",fg="375623",b=True,sz=10,h="left",merge_to=4+len(months))

    fm = make_pivot(farm_df,'รหัสลูกค้า',months)
    fm['รวม']       = fm.sum(axis=1)
    fm['เฉลี่ย']   = fm[months].mean(axis=1)
    fm['%ยอดรวม']  = fm['รวม']/grand_total
    fm['เจ้าของ']  = fm.index.map(farm_df.groupby('รหัสลูกค้า')['ชื่อลูกค้า'].first())
    fm = fm.sort_values('รวม',ascending=False)

    hdrs = ['อันดับ','ฟาร์ม (รหัสลูกค้า)','เจ้าของ']+[MTH[m] for m in months]+['รวม','เฉลี่ย','%']
    for ci,h in enumerate(hdrs,1):
        W(ws,4,ci,h,bg="375623",fg="FFFFFF",b=True,sz=8)
    ws.row_dimensions[4].height=22

    for ri_off,(fc,row) in enumerate(fm.iterrows(),1):
        ri=ri_off+4; alt="F0FFF0" if ri_off%2==0 else None
        ws.row_dimensions[ri].height=15
        W(ws,ri,1,ri_off,bg=alt,b=True)
        W(ws,ri,2,fc,bg=alt,h="left",b=True,sz=9)
        W(ws,ri,3,row.get('เจ้าของ',''),bg=alt,h="left",sz=8)
        for ci,m in enumerate(months,4):
            v=row[m]; W(ws,ri,ci,v,bg="FFF0F0" if v==0 else alt,fmt='#,##0',sz=8)
        W(ws,ri,4+len(months),row['รวม'],bg="A9D18E",b=True,fmt='#,##0.00')
        W(ws,ri,5+len(months),row['เฉลี่ย'],bg=alt,fmt='#,##0',sz=8)
        W(ws,ri,6+len(months),row['%ยอดรวม'],bg=alt,fmt='0.00%',sz=8)

    ws.column_dimensions['A'].width=6; ws.column_dimensions['B'].width=34
    ws.column_dimensions['C'].width=26
    for i in range(4,4+len(months)+4): ws.column_dimensions[gc(i)].width=12


def build_sheet5(wb, df, grand_total, months,
                 plc_target, plc_deadline, recco_target, recco_deadline):
    ws = wb.create_sheet("5_เป้าหมาย_PLC_Recco")
    ws.sheet_properties.tabColor = "C00000"
    W(ws,1,1,"ติดตามเป้าหมายยอดขาย",bg="C00000",fg="FFFFFF",b=True,sz=13,h="left",merge_to=12)

    plc_names  = ['บริษัท วีว่าพรีเมี่ยม เพ็ท สโตร์ จำกัด',
                  'บริษัท วีว่า เพ็ทสโตร์ จำกัด','บริษัท วีว่า เพ็ท สโตร์ จำกัด']
    plc_df   = df[df['ชื่อลูกค้า'].isin(plc_names)]
    plc_act  = plc_df['ราคารวม'].sum()
    recco_df = df[df['ชื่อลูกค้า']=='บริษัท เรคโค เพ็ท จำกัด']
    recco_act= recco_df['ราคารวม'].sum()

    def block(start_col, name, target, actual, deadline, store_df):
        c = start_col
        W(ws,3,c,name,bg="2E75B6",fg="FFFFFF",b=True,sz=11,merge_to=c+5)
        rows = [("เป้าหมาย",target,'#,##0'),
                ("ยอดปัจจุบัน",actual,'#,##0.00'),
                ("ยอดที่เหลือ",max(0,target-actual),'#,##0.00'),
                ("% บรรลุเป้า",actual/target if target else 0,'0.0%'),
                ("กำหนดเวลา",deadline,None)]
        for ri,(lbl,val,fmt) in enumerate(rows,4):
            W(ws,ri,c,lbl,bg="DEEAF1",b=True,sz=9,merge_to=c+1)
            pct = actual/target if target else 0
            vbg = "C6EFCE" if pct>=1 else ("FFF2CC" if pct>=0.7 else "FFD7D7") if lbl=="% บรรลุเป้า" else "FFFFFF"
            W(ws,ri,c+2,val,bg=vbg,b=(lbl=="% บรรลุเป้า"),sz=11,fmt=fmt,merge_to=c+5)

        W(ws,9,c,"เดือน",bg="2E75B6",fg="FFFFFF",b=True,sz=8)
        W(ws,9,c+1,"ยอดขาย",bg="2E75B6",fg="FFFFFF",b=True,sz=8)
        W(ws,9,c+2,"% ของเป้า",bg="2E75B6",fg="FFFFFF",b=True,sz=8)
        for ri2,m in enumerate(months,10):
            mv = store_df[store_df['month']==m]['ราคารวม'].sum()
            alt2 = "F0F4FF" if ri2%2==0 else None
            W(ws,ri2,c,MTH[m],bg=alt2,sz=8)
            W(ws,ri2,c+1,mv,bg=alt2,fmt='#,##0.00',sz=8)
            W(ws,ri2,c+2,mv/target if target else 0,bg=alt2,fmt='0.0%',sz=8)
        sr = 10+len(months)
        W(ws,sr,c,"สะสม",bg="BDD7EE",b=True,sz=9)
        W(ws,sr,c+1,actual,bg="BDD7EE",b=True,fmt='#,##0.00')
        W(ws,sr,c+2,actual/target if target else 0,bg="BDD7EE",b=True,fmt='0.0%')

    block(1, "Pet Lover Centre",         plc_target,   plc_act,   plc_deadline,   plc_df)
    block(8, "บริษัท เรคโค เพ็ท จำกัด", recco_target, recco_act, recco_deadline, recco_df)

    for i in range(1,15): ws.column_dimensions[gc(i)].width=14


def build_sheet6(wb, df, grand_total):
    ws = wb.create_sheet("6_ภาพรวมสินค้า")
    ws.sheet_properties.tabColor = "7030A0"
    W(ws,1,1,"ภาพรวมยอดขายสินค้า – แยกหมวดหมู่",
      bg="3B1F78",fg="FFFFFF",b=True,sz=12,h="left",merge_to=5)
    cats = df.groupby('หมวดหมู่').agg(
        ยอดขาย=('ราคารวม','sum'),จำนวนชิ้น=('จำนวน','sum'),
        จำนวน_SKU=('prod_grp','nunique')).reset_index().sort_values('ยอดขาย',ascending=False)
    cats['%'] = cats['ยอดขาย']/grand_total
    CAT_CLR = {'Dog Food':'D6E4F7','Cat Food':'E8D5F5','Supplement':'D9F0E0',
               'Healthy Snack':'FFF2CC','Deserve Life':'FFE6D9','RAW MATERIAL':'F2F2F2'}
    hdrs=['หมวดหมู่','ยอดขาย (บาท)','% ของยอดรวม','จำนวนชิ้น','จำนวน SKU']
    for ci,h in enumerate(hdrs,1):
        W(ws,3,ci,h,bg="3B1F78",fg="FFFFFF",b=True,sz=9)
    for ri_off,row in enumerate(cats.itertuples(),1):
        ri=ri_off+3; cbg=CAT_CLR.get(row.หมวดหมู่,'F2F2F2')
        W(ws,ri,1,row.หมวดหมู่,bg=cbg,h="left",b=True,sz=9)
        W(ws,ri,2,row.ยอดขาย,bg="F2F2F2" if ri_off%2==0 else None,fmt='#,##0.00')
        W(ws,ri,3,row._5,fmt='0.00%')
        W(ws,ri,4,row.จำนวนชิ้น,fmt='#,##0')
        W(ws,ri,5,row.จำนวน_SKU)
    tr=len(cats)+4
    W(ws,tr,1,"รวม",bg="BDD7EE",b=True,h="left")
    W(ws,tr,2,grand_total,bg="2E75B6",fg="FFFFFF",b=True,fmt='#,##0.00')
    W(ws,tr,3,1.0,bg="BDD7EE",b=True,fmt='0.0%')
    ws.column_dimensions['A'].width=18; ws.column_dimensions['B'].width=16
    ws.column_dimensions['C'].width=14; ws.column_dimensions['D'].width=13; ws.column_dimensions['E'].width=11


def build_sheet7(wb, df, grand_total, months, prod_names, prod_cat):
    ws = wb.create_sheet("7_Ranking_สินค้า")
    ws.sheet_properties.tabColor = "7030A0"
    W(ws,1,1,"Ranking ยอดขายสินค้า (กลุ่มรหัส 6 ตัว)",
      bg="3B1F78",fg="FFFFFF",b=True,sz=12,h="left",merge_to=4+len(months)+3)
    pm = make_pivot(df,'prod_grp',months)
    pm['รวม']=pm.sum(axis=1); pm['เฉลี่ย']=pm[months].mean(axis=1)
    pm['%']=pm['รวม']/grand_total
    pm['ชื่อ']=pm.index.map(prod_names); pm['หมวด']=pm.index.map(prod_cat)
    pm=pm.sort_values('รวม',ascending=False)
    CAT_CLR={'Dog Food':'D6E4F7','Cat Food':'E8D5F5','Supplement':'D9F0E0',
             'Healthy Snack':'FFF2CC','Deserve Life':'FFE6D9','RAW MATERIAL':'F2F2F2'}
    hdrs=['อันดับ','รหัส','ชื่อสินค้า','หมวดหมู่']+[MTH[m] for m in months]+['รวม','เฉลี่ย','%']
    for ci,h in enumerate(hdrs,1):
        W(ws,3,ci,h,bg="3B1F78",fg="FFFFFF",b=True,sz=8,wrap=True)
    ws.row_dimensions[3].height=30
    for ri_off,(code,row) in enumerate(pm.iterrows(),1):
        ri=ri_off+3; alt="F5F0FF" if ri_off%2==0 else None
        ws.row_dimensions[ri].height=14
        W(ws,ri,1,ri_off,bg=alt,b=(ri_off<=10))
        W(ws,ri,2,code,bg=alt,sz=8)
        W(ws,ri,3,row['ชื่อ'],bg=alt,h="left",sz=8,wrap=True)
        W(ws,ri,4,row['หมวด'],bg=CAT_CLR.get(row['หมวด'],'F2F2F2'),sz=8)
        for ci,m in enumerate(months,5):
            v=row[m]; W(ws,ri,ci,v,bg="FFF0F0" if v==0 else alt,fmt='#,##0',sz=8)
        W(ws,ri,5+len(months),row['รวม'],bg="E8D5F5",b=True,fmt='#,##0.00')
        W(ws,ri,6+len(months),row['เฉลี่ย'],bg=alt,fmt='#,##0',sz=8)
        W(ws,ri,7+len(months),row['%'],bg=alt,fmt='0.00%',sz=8)
    ws.column_dimensions['A'].width=6; ws.column_dimensions['B'].width=9
    ws.column_dimensions['C'].width=34; ws.column_dimensions['D'].width=13
    for i in range(5,5+len(months)+4): ws.column_dimensions[gc(i)].width=11
    ws.freeze_panes="E4"


def build_sheets_89(wb, df, grand_total, months, prod_names, prod_cat):
    """Sheet 8 & 9: สินค้า × Top10 ร้าน/รหัส"""
    for sheet_num, group_col, sname, tab_color in [
        (8, 'ชื่อร้าน',        '8_สินค้า×Top10ร้าน',  'ED7D31'),
        (9, 'รหัสลูกค้า_norm', '9_สินค้า×Top10รหัส',  'FF0000'),
    ]:
        ws = wb.create_sheet(sname)
        ws.sheet_properties.tabColor = tab_color
        W(ws,1,1,f"Ranking สินค้า – Top10 {group_col}",
          bg="1F4E79",fg="FFFFFF",b=True,sz=12,h="left",merge_to=3+len(months)+2)

        # Top10
        top10 = (df.groupby(group_col)['ราคารวม'].sum()
                   .sort_values(ascending=False).head(10).index.tolist())
        hbg   = "ED7D31" if sheet_num==8 else "C00000"

        ri=3
        for rank_i, grp_val in enumerate(top10,1):
            grp_total = df[df[group_col]==grp_val]['ราคารวม'].sum()
            W(ws,ri,1,f"#{rank_i} {grp_val}  (รวม {grp_total:,.0f} บาท)",
              bg=hbg,fg="FFFFFF",b=True,sz=10,h="left",merge_to=3+len(months)+2)
            ri+=1
            for ci,h in enumerate(['รหัส','ชื่อสินค้า','หมวดหมู่']+[MTH[m] for m in months]+['รวม','%'],1):
                W(ws,ri,ci,h,bg="FCE4D6" if sheet_num==8 else "FCE9E9",b=True,sz=8)
            ri+=1

            sub = df[df[group_col]==grp_val]
            pm  = make_pivot(sub,'prod_grp',months)
            pm['รวม']=pm.sum(axis=1)
            pm=pm.sort_values('รวม',ascending=False)

            for ri_off2,(pcode,prow) in enumerate(pm.iterrows(),1):
                alt2="F9F5FF" if ri_off2%2==0 else None
                ws.row_dimensions[ri].height=13
                W(ws,ri,1,pcode,bg=alt2,sz=8)
                W(ws,ri,2,prod_names.get(pcode,pcode),bg=alt2,h="left",sz=8,wrap=True)
                W(ws,ri,3,prod_cat.get(pcode,''),bg=alt2,sz=8)
                for ci,m in enumerate(months,4):
                    v=prow.get(m,0); W(ws,ri,ci,v,bg="FFF0F0" if v==0 else alt2,fmt='#,##0',sz=8)
                W(ws,ri,4+len(months),prow['รวม'],bg="BDD7EE",b=True,fmt='#,##0')
                W(ws,ri,5+len(months),prow['รวม']/grp_total if grp_total else 0,
                  bg=alt2,fmt='0.0%',sz=8)
                ri+=1
            ri+=1

        ws.column_dimensions['A'].width=10; ws.column_dimensions['B'].width=34
        ws.column_dimensions['C'].width=13
        for i in range(4,4+len(months)+3): ws.column_dimensions[gc(i)].width=11
        ws.freeze_panes="D2"


def build_sheet10(wb, df, grand_total, months):
    ACT=months; FCI=list(range(max(months)+1,13))
    if not FCI: FCI=list(range(6,13))

    ws = wb.create_sheet("10_Forecast_ร้านค้า")
    ws.sheet_properties.tabColor = "1F4E79"
    W(ws,1,1,f"📊 Forecast & วิเคราะห์ร้านค้า  |  ยอดจริง {MTH[min(ACT)]}–{MTH[max(ACT)]} + คาดการณ์ {MTH[min(FCI)]}–{MTH[max(FCI)]} 2569",
      bg="1F4E79",fg="FFFFFF",b=True,sz=12,h="left",merge_to=3+len(ACT)+1+len(FCI)+1+len(FCI)+1+3)

    # Legend row 2
    for ci,(cat,st) in enumerate(STORE_CAT_STYLE.items(),1):
        W(ws,2,ci,cat,bg=st['bg'],fg=st['fg'],b=True,sz=8)
    ws.row_dimensions[2].height=16

    sm = make_pivot(df,'ชื่อร้าน',ACT)
    store_rows=[]
    for store,row in sm.iterrows():
        v=row.values.astype(float)
        sl,r2,fc_lin=fc_linear(v,ACT,FCI)
        w3=wma3(v)
        fc_ma={m:w3 for m in FCI}
        active=sum(1 for x in v if x>0)
        peak_m=ACT[int(np.argmax(v))]; peak_v=float(np.max(v)); cur=float(v[-1])
        r2a=np.mean(v[len(v)//2:]); r2b=max(np.mean(v[:len(v)//2]),1)
        if   sl>8000 and r2>0.5:               cat='🌟 STAR'
        elif sl>2000 and r2>0.3:               cat='📈 GROWTH'
        elif r2a>r2b*1.2 and active>=3:        cat='🔄 RECOVERING'
        elif peak_v>0 and cur<peak_v*.55 and active>=3: cat='📉 DECLINING'
        elif sl<-3000 and active>=3:           cat='⚠️ WARNING'
        elif active<=2:                        cat='⚡ INACTIVE'
        elif active==3:                        cat='🔵 SPORADIC'
        else:                                  cat='😴 STABLE'

        action_map={
            '🌟 STAR':"ขยาย SKU / volume discount",
            '📈 GROWTH':"ติดตาม / เสนอ bundle",
            '🔄 RECOVERING':"โปร + ติดตาม 2 เดือน",
            '📉 DECLINING':f"ยอดตกจาก peak {MTH[peak_m]} → โทรหา/เสนอโปรฯ",
            '⚠️ WARNING':"เร่งด่วน! ตรวจสอบสาเหตุ",
            '⚡ INACTIVE':f"ไม่มียอดใน M{max(ACT)-1}-M{max(ACT)} → follow up",
            '🔵 SPORADIC':"ซื้อไม่สม่ำเสมอ → สร้าง habit",
            '😴 STABLE':"รักษาฐาน / เสนอสินค้าใหม่",
        }
        store_rows.append({
            'ชื่อร้าน':store,'cat':cat,'action':action_map.get(cat,''),
            **{f'M{m}':v[i] for i,m in enumerate(ACT)},
            'total':v.sum(),'pct':v.sum()/grand_total,
            'active':active,'peak_m':peak_m,'peak_v':peak_v,'cur':cur,
            'vs_peak':(cur-peak_v)/peak_v if peak_v>0 else 0,
            'slope':sl,'r2':r2,'wma3':w3,
            'fc_lin':fc_lin,'fc_ma':fc_ma,
            'fc_lin_sum':sum(fc_lin.values()),'fc_ma_sum':sum(fc_ma.values()),
        })

    CAT_ORDER=['🌟 STAR','📈 GROWTH','🔄 RECOVERING','📉 DECLINING','⚠️ WARNING','🔵 SPORADIC','😴 STABLE','⚡ INACTIVE']
    stores=pd.DataFrame(store_rows)
    stores['_so']=stores['cat'].map({v:i for i,v in enumerate(CAT_ORDER)})
    stores=stores.sort_values(['_so','total'],ascending=[True,False])

    # Summary row 3
    W(ws,3,1,"📌",bg="2E75B6",fg="FFFFFF",b=True)
    ci2=2
    for cat in CAT_ORDER:
        sub=stores[stores['cat']==cat]
        if len(sub)>0:
            st2=STORE_CAT_STYLE.get(cat,{'bg':'F2F2F2','fg':'000000'})
            W(ws,3,ci2,f"{cat}: {len(sub)} ร้าน / {sub['total'].sum():,.0f}฿",
              bg=st2['bg'],fg=st2['fg'],b=True,sz=8,h="left"); ci2+=1
    ws.row_dimensions[3].height=16

    # Headers row 4
    ws.row_dimensions[4].height=36
    HDRS=[('#','1F4E79'),('ชื่อร้าน','1F4E79'),('สถานะ','1F4E79')]
    for m in ACT: HDRS.append((f"{MTH[m]}\nจริง","2E75B6"))
    HDRS.append(('รวมจริง','1F4E79'))
    HDRS+=[('Peak\nเดือน','595959'),('Peak\nมูลค่า','595959'),('vs Peak','595959')]
    for m in FCI: HDRS.append((f"🔵{MTH[m]}\nLinear","375623"))
    HDRS.append(('FC รวม\nLinear','1E5C1E'))
    for m in FCI: HDRS.append((f"🟠{MTH[m]}\nWMA","C55A11"))
    HDRS.append(('FC รวม\nWMA','7B2D00'))
    HDRS+=[('slope/\nเดือน','595959'),('%ยอดรวม','595959'),('📋 แนะนำ','1F4E79')]
    for ci,(h,bg) in enumerate(HDRS,1):
        W(ws,4,ci,h,bg=bg,fg="FFFFFF",b=True,sz=8,wrap=True)

    for ri_off,(_,row) in enumerate(stores.iterrows(),1):
        r=ri_off+4; ws.row_dimensions[r].height=15
        st2=STORE_CAT_STYLE.get(row['cat'],{'bg':None,'fg':'000000'})
        alt="F7FBFF" if ri_off%2==0 else None; ci=1
        W(ws,r,ci,ri_off,bg=alt,b=(ri_off<=10)); ci+=1
        W(ws,r,ci,row['ชื่อร้าน'],bg=alt,h="left",b=(row['cat'] in ['🌟 STAR','📈 GROWTH']),sz=9); ci+=1
        W(ws,r,ci,row['cat'],bg=st2['bg'],fg=st2['fg'],b=True,sz=8); ci+=1
        for m in ACT:
            v=row[f'M{m}']
            mbg="FFF0A0" if m==row['peak_m'] and v>0 else ("FFF0F0" if v==0 else alt)
            W(ws,r,ci,v,bg=mbg,fmt='#,##0',sz=8); ci+=1
        W(ws,r,ci,row['total'],bg="BDD7EE",b=True,fmt='#,##0.00'); ci+=1
        W(ws,r,ci,MTH.get(row['peak_m'],''),bg=alt,sz=8); ci+=1
        W(ws,r,ci,row['peak_v'],bg=alt,fmt='#,##0',sz=8); ci+=1
        vp=row['vs_peak']; vbg="FFD7D7" if vp<-0.4 else("FFF2CC" if vp<-0.15 else("E2EFDA" if vp>=0 else alt))
        W(ws,r,ci,vp,bg=vbg,fmt='+0%;-0%;0%',b=(abs(vp)>0.3),sz=8); ci+=1
        for m in FCI:
            v=row['fc_lin'][m]; W(ws,r,ci,v,bg="E2EFDA" if v>0 else "FFF0F0",fmt='#,##0',sz=8,italic=True); ci+=1
        W(ws,r,ci,row['fc_lin_sum'],bg="A9D18E",b=True,fmt='#,##0.00'); ci+=1
        for m in FCI:
            W(ws,r,ci,row['fc_ma'][m],bg="FFF2E3",fmt='#,##0',sz=8,italic=True); ci+=1
        W(ws,r,ci,row['fc_ma_sum'],bg="F4B183",b=True,fmt='#,##0.00'); ci+=1
        W(ws,r,ci,row['slope'],bg="E2EFDA" if row['slope']>0 else "FFD7D7",fmt='+#,##0;-#,##0;0',sz=8); ci+=1
        W(ws,r,ci,row['pct'],bg=alt,fmt='0.00%',sz=8); ci+=1
        W(ws,r,ci,row['action'],bg=alt,h="left",sz=8,wrap=True); ci+=1

    # Total row
    tr=len(stores)+5
    W(ws,tr,1,'',bg="1F4E79"); W(ws,tr,2,'รวม',bg="1F4E79",fg="FFFFFF",b=True,h="left")
    W(ws,tr,3,'',bg="1F4E79")
    for ci,m in enumerate(ACT,4):
        W(ws,tr,ci,df[df['month']==m]['ราคารวม'].sum(),bg="2E75B6",fg="FFFFFF",b=True,fmt='#,##0')
    W(ws,tr,4+len(ACT),grand_total,bg="1F4E79",fg="FFFFFF",b=True,fmt='#,##0.00')
    for i in range(3): W(ws,tr,5+len(ACT)+i,'',bg="1F4E79")
    fct=5+len(ACT)+3
    for i,m in enumerate(FCI):
        W(ws,tr,fct+i,stores.apply(lambda r_:r_['fc_lin'][m],axis=1).sum(),bg="375623",fg="FFFFFF",b=True,fmt='#,##0')
    W(ws,tr,fct+len(FCI),stores['fc_lin_sum'].sum(),bg="1E5C1E",fg="FFFFFF",b=True,fmt='#,##0.00')
    mat=fct+len(FCI)+1
    for i,m in enumerate(FCI):
        W(ws,tr,mat+i,stores.apply(lambda r_:r_['fc_ma'][m],axis=1).sum(),bg="C55A11",fg="FFFFFF",b=True,fmt='#,##0')
    W(ws,tr,mat+len(FCI),stores['fc_ma_sum'].sum(),bg="7B2D00",fg="FFFFFF",b=True,fmt='#,##0.00')

    ws.column_dimensions['A'].width=5; ws.column_dimensions['B'].width=30
    ws.column_dimensions['C'].width=18
    for i in range(4,4+len(ACT)+4): ws.column_dimensions[gc(i)].width=11
    for i in range(4+len(ACT)+4,4+len(ACT)+4+len(FCI)*2+2): ws.column_dimensions[gc(i)].width=9
    lc=4+len(ACT)+4+len(FCI)*2+2
    ws.column_dimensions[gc(lc)].width=9; ws.column_dimensions[gc(lc+1)].width=8
    ws.column_dimensions[gc(lc+2)].width=30
    ws.freeze_panes=f"{gc(4+len(ACT))}5"


def build_sheet11(wb, df, grand_total, months, prod_names, prod_cat):
    ws = wb.create_sheet("11_Forecast_สินค้า")
    ws.sheet_properties.tabColor = "7030A0"
    W(ws,1,1,"📦 วิเคราะห์สินค้า × ร้านค้า  |  กำลังโต / เคยดีแล้วลด / หยุดซื้อ",
      bg="3B1F78",fg="FFFFFF",b=True,sz=12,h="left",merge_to=18)

    for ci,(k,(bg,fg,lbl)) in enumerate(PAIR_STYLE.items(),1):
        W(ws,2,ci,lbl,bg=bg,fg=fg,sz=8,b=True)
    ws.row_dimensions[2].height=15
    W(ws,3,1,"💡 header row = ยอดรวมสินค้าทุกร้าน  |  แถวย่อย = แต่ละร้านที่ซื้อสินค้านี้",
      bg="F3EEFF",fg="3B1F78",sz=8,h="left",merge_to=18)
    ws.row_dimensions[3].height=14

    P_HDRS=[('#','3B1F78'),('รหัส','3B1F78'),('ชื่อสินค้า','3B1F78'),
            ('หมวดหมู่','3B1F78'),('ชื่อร้าน','3B1F78'),('สถานะ','3B1F78')]
    for m in months: P_HDRS.append((f"{MTH[m]}\nจริง","2E75B6"))
    P_HDRS+=[('รวม','1F4E79'),('%สินค้า','595959'),('Peak','595959'),
             ('Early\nAvg','595959'),('Late\nAvg','595959'),('Trend\n%','595959'),('💬 Insight','3B1F78')]
    for ci,(h,bg) in enumerate(P_HDRS,1):
        W(ws,4,ci,h,bg=bg,fg="FFFFFF",b=True,sz=8,wrap=True)
    ws.row_dimensions[4].height=36

    # Cross analysis
    cross = df.groupby(['prod_grp','ชื่อร้าน','month'])['ราคารวม'].sum().unstack(fill_value=0)
    for m in months:
        if m not in cross.columns: cross[m]=0
    cross=cross[months]; cross['total']=cross.sum(axis=1)
    cross['early_avg']=cross[months[:3]].mean(axis=1)
    cross['late_avg'] =cross[months[-2:]].mean(axis=1)
    cross['trend_pct']=(cross['late_avg']-cross['early_avg'])/cross['early_avg'].replace(0,np.nan)
    cross['active']=(cross[months]>0).sum(axis=1)
    cross['peak_m']=cross[months].idxmax(axis=1)
    cross['peak_v']=cross[months].max(axis=1)
    cross['last_v']=cross[months[-1]]

    def cpair(row):
        tp=row['trend_pct']; ea=row['early_avg']; la=row['late_avg']
        if row['last_v']==0 and ea>5000: return 'STOPPED'
        if tp<-0.8 and ea>5000:         return 'DROPPED'
        if la>3000 and tp>0.3:           return 'GROWING'
        if row['active']<=1:             return 'ONE_TIME'
        if row['active']==2:             return 'SPORADIC'
        if tp<-0.2:                      return 'DECLINING'
        if tp>0.1:                       return 'GROWING'
        return 'STABLE'

    cross['pair_status']=cross.apply(cpair,axis=1)
    cross=cross.reset_index(); cross.columns.name=None
    cross['prod_name']=cross['prod_grp'].map(prod_names)
    cross=cross[cross['total']>1000].copy()

    # Product totals
    prod_tot=df.groupby(['prod_grp','month'])['ราคารวม'].sum().unstack(fill_value=0)
    for m in months:
        if m not in prod_tot.columns: prod_tot[m]=0
    prod_tot=prod_tot[months]; prod_tot['total']=prod_tot.sum(axis=1)
    prod_tot['early']=prod_tot[months[:3]].mean(axis=1)
    prod_tot['late'] =prod_tot[months[-2:]].mean(axis=1)
    prod_tot['trend']=(prod_tot['late']-prod_tot['early'])/prod_tot['early'].replace(0,np.nan)
    prod_tot=prod_tot.sort_values('total',ascending=False)

    CAT_CLR={'Dog Food':'D6E4F7','Cat Food':'E8D5F5','Supplement':'D9F0E0',
             'Healthy Snack':'FFF2CC','Deserve Life':'FFE6D9','RAW MATERIAL':'F2F2F2'}
    SORT_PS={'GROWING':0,'STABLE':1,'DECLINING':2,'DROPPED':3,'STOPPED':4,'SPORADIC':5,'ONE_TIME':6}

    ri=5; rank=0
    for pcode in prod_tot.index:
        if pcode not in prod_names: continue
        pname=prod_names[pcode]; pcat=prod_cat.get(pcode,'')
        prows=cross[cross['prod_grp']==pcode].copy()
        if len(prows)==0: continue
        ptotal=prod_tot.loc[pcode,'total']
        pearly=prod_tot.loc[pcode,'early']; plate=prod_tot.loc[pcode,'late']
        ptrend=prod_tot.loc[pcode,'trend']
        rank+=1
        pbg={'📈':'E2EFDA','📉':'FCE4D6'}.get(('📈' if ptrend>0.2 else '📉' if ptrend<-0.3 else ''), 'DEEAF1')
        ptlbl='📈 GROWTH' if ptrend>0.2 else ('📉 DECLINING' if ptrend<-0.3 else '➡️ STABLE')

        ws.row_dimensions[ri].height=17
        W(ws,ri,1,rank,bg=pbg,b=True,sz=9)
        W(ws,ri,2,pcode,bg=pbg,b=True,sz=9)
        W(ws,ri,3,pname,bg=pbg,b=True,sz=9,h="left",wrap=True)
        W(ws,ri,4,pcat,bg=CAT_CLR.get(pcat,'F2F2F2'),sz=8,b=True)
        W(ws,ri,5,f"รวม {len(prows)} ร้าน",bg=pbg,sz=8,italic=True)
        W(ws,ri,6,ptlbl,bg=pbg,b=True,sz=8)
        for ci2,m in enumerate(months,7):
            v=prod_tot.loc[pcode,m] if m in prod_tot.columns else 0
            W(ws,ri,ci2,v,bg=pbg,b=True,fmt='#,##0',sz=9)
        W(ws,ri,7+len(months),ptotal,bg=pbg,b=True,fmt='#,##0.00',sz=9)
        W(ws,ri,8+len(months),ptotal/grand_total,bg=pbg,fmt='0.00%',sz=8)
        peak_m2=prod_tot.loc[pcode,months].idxmax()
        W(ws,ri,9+len(months),MTH.get(peak_m2,''),bg=pbg,sz=8)
        W(ws,ri,10+len(months),pearly,bg=pbg,fmt='#,##0',sz=8)
        W(ws,ri,11+len(months),plate,bg=pbg,fmt='#,##0',sz=8)
        tbg="E2EFDA" if ptrend>0 else "FCE4D6"
        W(ws,ri,12+len(months),ptrend if not np.isnan(ptrend) else 0,bg=tbg,fmt='+0%;-0%;0%',b=True,sz=8)
        ng=len(prows[prows['pair_status']=='GROWING'])
        ns=len(prows[prows['pair_status'].isin(['STOPPED','DROPPED'])])
        ins=f"{ng} ร้านกำลังโต" + (f"  |  {ns} ร้านหยุดซื้อ" if ns>0 else "")
        W(ws,ri,13+len(months),ins,bg=pbg,h="left",sz=8)
        ri+=1

        prows=prows.copy()
        prows['_ps']=prows['pair_status'].map(SORT_PS).fillna(9)
        prows=prows.sort_values(['_ps','total'],ascending=[True,False])

        for _,pr in prows.iterrows():
            ws.row_dimensions[ri].height=14
            ps=pr['pair_status']; pb2,pf2,plbl2=PAIR_STYLE.get(ps,('F2F2F2','000000',''))
            alt2="FDFBFF" if ri%2==0 else None
            for ci2 in [1,2,3,4]: W(ws,ri,ci2,'',bg=alt2)
            W(ws,ri,5,pr['ชื่อร้าน'],bg=alt2,h="left",sz=8)
            W(ws,ri,6,plbl2,bg=pb2,fg=pf2,b=True,sz=8)
            for ci2,m in enumerate(months,7):
                v=pr.get(m,0)
                mbg="FFF0A0" if m==pr['peak_m'] and v>0 else("FFF0F0" if v==0 else alt2)
                W(ws,ri,ci2,v,bg=mbg,fmt='#,##0',sz=8)
            W(ws,ri,7+len(months),pr['total'],bg=alt2,b=(ps=='GROWING'),fmt='#,##0',sz=8)
            pcp=pr['total']/ptotal if ptotal>0 else 0
            W(ws,ri,8+len(months),pcp,bg=alt2,fmt='0.0%',sz=8)
            W(ws,ri,9+len(months),MTH.get(pr['peak_m'],''),bg=alt2,sz=8)
            W(ws,ri,10+len(months),pr['early_avg'],bg=alt2,fmt='#,##0',sz=8)
            W(ws,ri,11+len(months),pr['late_avg'],bg="FFF0F0" if pr['late_avg']==0 else alt2,fmt='#,##0',sz=8)
            tp2=pr['trend_pct'] if not(isinstance(pr['trend_pct'],float) and np.isnan(pr['trend_pct'])) else 0
            W(ws,ri,12+len(months),tp2,bg="E2EFDA" if tp2>0.1 else("FFD7D7" if tp2<-0.3 else alt2),fmt='+0%;-0%;0%',sz=8)
            if ps=='GROWING':    ins2=f"↑ avg {pr['late_avg']:,.0f}฿/เดือน → ดัน SKU เพิ่ม"
            elif ps=='STOPPED':  ins2=f"⛔ เคย avg {pr['early_avg']:,.0f}฿ → ไม่มียอดใน {MTH[max(months)]} → โทรหาด่วน"
            elif ps=='DROPPED':  ins2=f"⚠️ ลดจาก {pr['peak_v']:,.0f}฿ → ตรวจสอบเหตุผล"
            elif ps=='DECLINING': ins2=f"↓ trend {tp2:.0%} → เสนอโปรฯ"
            elif ps=='SPORADIC':  ins2="ซื้อสลับ → สร้างความสม่ำเสมอ"
            elif ps=='ONE_TIME':  ins2="ซื้อครั้งเดียว → follow up"
            else:                 ins2="คงที่ → รักษาฐาน"
            W(ws,ri,13+len(months),ins2,bg=alt2,h="left",sz=8,wrap=True)
            ri+=1
        ri+=1

    ws.column_dimensions['A'].width=4; ws.column_dimensions['B'].width=8
    ws.column_dimensions['C'].width=34; ws.column_dimensions['D'].width=13
    ws.column_dimensions['E'].width=28; ws.column_dimensions['F'].width=14
    for i in range(7,7+len(months)): ws.column_dimensions[gc(i)].width=10
    for i in range(7+len(months),7+len(months)+7): ws.column_dimensions[gc(i)].width=11
    ws.column_dimensions[gc(7+len(months)+7)].width=32
    ws.freeze_panes="D5"


def build_sheet12(wb, df, grand_total, months, monthly_targets):
    """Sheet 12: เป้าหมายยอดขายรายเดือน vs ยอดจริง"""
    ws = wb.create_sheet("12_เป้าหมายรายเดือน")
    ws.sheet_properties.tabColor = "FF0000"

    MTH_FULL = {1:'ม.ค.',2:'ก.พ.',3:'มี.ค.',4:'เม.ย.',5:'พ.ค.',
                6:'มิ.ย.',7:'ก.ค.',8:'ส.ค.',9:'ก.ย.',10:'ต.ค.',11:'พ.ย.',12:'ธ.ค.'}

    all_months = sorted(set(list(months) + list(monthly_targets.keys())))

    # Title
    W(ws,1,1,"🎯 เป้าหมายยอดขายรายเดือน vs ยอดจริง",
      bg="C00000",fg="FFFFFF",b=True,sz=13,h="left",merge_to=8)
    ws.row_dimensions[1].height = 26

    # ── Summary KPI row ──
    actual_in_target_months = sum(
        df[df['month']==m]['ราคารวม'].sum() for m in monthly_targets
    )
    total_target = sum(monthly_targets.values())
    overall_ach  = actual_in_target_months / total_target if total_target else 0
    gap          = actual_in_target_months - total_target

    kpis = [
        ("เป้าหมายรวม", total_target, '#,##0'),
        ("ยอดจริงรวม",  actual_in_target_months, '#,##0.00'),
        ("% Achievement", overall_ach, '0.0%'),
        ("Gap (จริง - เป้า)", gap, '+#,##0;-#,##0;0'),
    ]
    for i,(lbl,val,fmt) in enumerate(kpis):
        vbg = "C6EFCE" if (i==2 and overall_ach>=1) else \
              "FFF2CC" if (i==2 and overall_ach>=0.8) else \
              "FFD7D7" if (i==2) else \
              "E2EFDA" if (i==3 and gap>=0) else \
              "FFD7D7" if (i==3) else "DEEAF1"
        W(ws,3,i*2+1,lbl,bg="2E75B6",fg="FFFFFF",b=True,sz=9)
        W(ws,4,i*2+1,val,bg=vbg,b=True,sz=12,fg="1F4E79",fmt=fmt,merge_to=i*2+2)
    ws.row_dimensions[4].height = 26

    # ── Monthly detail table ──
    hdrs = ["เดือน","เป้าหมาย (บาท)","ยอดจริง (บาท)",
            "% Achievement","Gap (บาท)","สถานะ",
            "ร้านค้าที่ active","% vs ยอดรวมปี"]
    for ci,h in enumerate(hdrs,1):
        W(ws,6,ci,h,bg="C00000",fg="FFFFFF",b=True,sz=9,wrap=True)
    ws.row_dimensions[6].height = 28

    for ri_off,m in enumerate(all_months,1):
        ri = ri_off + 6
        ws.row_dimensions[ri].height = 16
        actual = df[df['month']==m]['ราคารวม'].sum()
        target = monthly_targets.get(m, None)
        alt    = "F9F9F9" if ri_off%2==0 else None

        W(ws,ri,1,MTH_FULL.get(m,''),bg=alt,b=True,sz=10)
        # เป้า
        if target:
            W(ws,ri,2,target,bg=alt,fmt='#,##0',sz=9)
        else:
            W(ws,ri,2,"—",bg="F5F5F5",fg="AAAAAA",sz=9,italic=True)

        # ยอดจริง — แสดงเฉพาะเดือนที่มีข้อมูล
        if m in months:
            W(ws,ri,3,actual,bg=alt,fmt='#,##0.00',sz=9)
        else:
            W(ws,ri,3,"(ยังไม่มีข้อมูล)",bg="F5F5F5",fg="AAAAAA",sz=9,italic=True)

        # % Achievement, gap, status — เฉพาะเดือนที่มีทั้งเป้าและข้อมูลจริง
        if target and m in months:
            ach = actual / target
            gap_m = actual - target
            if   ach >= 1.0:  status = "✅ ถึงเป้า";    sbg = "C6EFCE"; sfg = "375623"
            elif ach >= 0.9:  status = "🟡 ใกล้เป้า";   sbg = "FFF2CC"; sfg = "7B5E00"
            elif ach >= 0.7:  status = "🟠 ต่ำกว่าเป้า"; sbg = "FCE4D6"; sfg = "843C0C"
            else:             status = "🔴 ต่ำมาก";     sbg = "FFD7D7"; sfg = "C00000"
            W(ws,ri,4,ach,  bg="C6EFCE" if ach>=1 else("FFF2CC" if ach>=0.9 else("FCE4D6" if ach>=0.7 else "FFD7D7")),
              fmt='0.0%',b=True,sz=9)
            W(ws,ri,5,gap_m,bg="E2EFDA" if gap_m>=0 else "FFD7D7",
              fmt='+#,##0;-#,##0;0',sz=9)
            W(ws,ri,6,status,bg=sbg,fg=sfg,b=True,sz=9)
        else:
            for ci in [4,5,6]:
                W(ws,ri,ci,"—",bg="F5F5F5",fg="AAAAAA",sz=9,italic=True)

        # ร้านค้าที่ active
        n_stores = df[df['month']==m]['ชื่อร้าน'].nunique() if m in months else 0
        W(ws,ri,7,n_stores if m in months else "—",bg=alt,sz=9)
        # % vs ยอดรวมปี
        pct_yr = actual/grand_total if m in months and grand_total else 0
        W(ws,ri,8,pct_yr if m in months else "—",
          bg=alt,fmt='0.0%' if m in months else None,sz=9)

    # ── Store breakdown: ยอดแต่ละร้านเทียบกับเป้าเดือน ──
    last_data_row = 6 + len(all_months) + 2
    W(ws,last_data_row,1,"📊 ยอดขายแต่ละร้าน vs เป้าหมายรายเดือน",
      bg="2E75B6",fg="FFFFFF",b=True,sz=10,h="left",
      merge_to=2+len([m for m in all_months if m in monthly_targets]))
    ws.row_dimensions[last_data_row].height = 18

    sub_months = [m for m in all_months if m in monthly_targets]
    hdr2 = ["ชื่อร้าน"] + [f"{MTH_FULL[m]}\nจริง" for m in sub_months] + \
           [f"{MTH_FULL[m]}\nเป้า" for m in sub_months] + ["รวมจริง","% ของเป้า"]
    for ci,h in enumerate(hdr2,1):
        W(ws,last_data_row+1,ci,h,bg="2E75B6",fg="FFFFFF",b=True,sz=8,wrap=True)
    ws.row_dimensions[last_data_row+1].height=30

    sm = df.groupby(['ชื่อร้าน','month'])['ราคารวม'].sum().unstack(fill_value=0)
    for m in sub_months:
        if m not in sm.columns: sm[m] = 0
    sm['total'] = sm[sub_months].sum(axis=1)
    sm = sm.sort_values('total',ascending=False)
    total_tgt_sub = sum(monthly_targets[m] for m in sub_months)

    for ri_off2,(store,row) in enumerate(sm.iterrows(),1):
        ri2 = last_data_row + 1 + ri_off2
        ws.row_dimensions[ri2].height = 14
        alt2 = "F0F4FF" if ri_off2%2==0 else None
        W(ws,ri2,1,store,bg=alt2,h="left",sz=8,b=(ri_off2<=10))
        for ci,m in enumerate(sub_months,2):
            v=row.get(m,0)
            tgt=monthly_targets[m]
            pct=v/tgt if tgt else 0
            cbg="C6EFCE" if pct>=1 else("FFF2CC" if pct>=0.8 else("FFD7D7" if v>0 and pct<0.5 else alt2))
            W(ws,ri2,ci,v,bg=cbg,fmt='#,##0',sz=8)
        for ci,m in enumerate(sub_months,2+len(sub_months)):
            W(ws,ri2,ci,monthly_targets[m],bg="F5F5F5",fmt='#,##0',sz=8,italic=True)
        W(ws,ri2,2+len(sub_months)*2,row['total'],bg="BDD7EE",b=True,fmt='#,##0')
        pct_s = row['total']/total_tgt_sub if total_tgt_sub else 0
        pbg = "C6EFCE" if pct_s>=1 else("FFF2CC" if pct_s>=0.8 else "FCE4D6")
        W(ws,ri2,3+len(sub_months)*2,pct_s,bg=pbg,fmt='0.0%',b=True)

    # Widths
    ws.column_dimensions['A'].width=26
    for i in range(2,3+len(sub_months)*2+2): ws.column_dimensions[gc(i)].width=12
    ws.freeze_panes="B7"


# ─── MAIN ENTRY POINT ────────────────────────────────────────────────────────
def build_dashboard(file_bytes, plc_target=1_250_000, plc_deadline="30 มิถุนายน 2569",
                    recco_target=1_000_000, recco_deadline="31 ธันวาคม 2569",
                    monthly_targets=None, progress_cb=None):
    def upd(pct, msg):
        if progress_cb: progress_cb(pct, msg)

    upd(10, "📥 โหลดข้อมูลและเงื่อนไข...")
    df, prod_names, prod_cat = load_data(file_bytes)
    months = sorted(df['month'].dropna().unique().astype(int))
    grand_total = df['ราคารวม'].sum()

    wb = Workbook()
    wb.remove(wb.active)

    upd(20, "📊 สร้างภาพรวมยอดขาย...")
    build_sheet1(wb, df, grand_total, months)

    upd(30, "🏪 สร้าง Ranking ร้านค้า...")
    build_sheet2(wb, df, grand_total, months)

    upd(38, "🔑 สร้าง Ranking รหัสลูกค้า...")
    build_sheet3(wb, df, grand_total, months)

    upd(44, "🌾 สร้างข้อมูลฟาร์ม...")
    build_sheet4(wb, df, grand_total, months)

    upd(50, "🎯 คำนวณเป้าหมาย PLC & Recco...")
    build_sheet5(wb, df, grand_total, months, plc_target, plc_deadline, recco_target, recco_deadline)

    upd(56, "📦 ภาพรวมสินค้า...")
    build_sheet6(wb, df, grand_total)

    upd(62, "📋 Ranking สินค้า...")
    build_sheet7(wb, df, grand_total, months, prod_names, prod_cat)

    upd(70, "🔗 สินค้า × Top10...")
    build_sheets_89(wb, df, grand_total, months, prod_names, prod_cat)

    upd(80, "🔮 Forecast ร้านค้า...")
    build_sheet10(wb, df, grand_total, months)

    upd(92, "📦 Forecast สินค้า × ร้าน...")
    build_sheet11(wb, df, grand_total, months, prod_names, prod_cat)

    if monthly_targets:
        upd(97, "🎯 สร้างเป้าหมายรายเดือน...")
        build_sheet12(wb, df, grand_total, months, monthly_targets)

    upd(98, "💾 บันทึกไฟล์...")
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()
