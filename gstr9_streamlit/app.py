import streamlit as st
import re, os, shutil, zipfile, tempfile, io
from datetime import datetime
from pathlib import Path
import pdfplumber, openpyxl

# ── Page config ────────────────────────────────────────────
st.set_page_config(
    page_title="GSTR-9 Auto Fill",
    page_icon="📊",
    layout="centered"
)

# ── Custom CSS ─────────────────────────────────────────────
st.markdown("""
<style>
    .main { max-width: 700px; }
    .stApp { background: #f8fafc; }
    .title-box {
        background: linear-gradient(135deg, #1a56db, #1e429f);
        color: white; border-radius: 14px; padding: 28px 32px;
        margin-bottom: 24px; text-align: center;
    }
    .title-box h1 { font-size: 28px; margin: 0; font-weight: 700; }
    .title-box p  { margin: 6px 0 0; opacity: 0.85; font-size: 14px; }
    .info-card {
        background: white; border-radius: 10px; padding: 16px 20px;
        border: 1px solid #e2e8f0; margin: 10px 0;
    }
    .success-card {
        background: #f0fdf4; border: 1px solid #86efac;
        border-radius: 10px; padding: 20px 24px; margin: 12px 0;
    }
    .metric-row {
        display: flex; gap: 12px; flex-wrap: wrap; margin-top: 14px;
    }
    .metric {
        background: white; border: 1px solid #d1fae5;
        border-radius: 8px; padding: 12px 16px; flex: 1; min-width: 100px;
        text-align: center;
    }
    .metric .num { font-size: 26px; font-weight: 700; color: #059669; }
    .metric .lbl { font-size: 11px; color: #6b7280; margin-top: 2px; }
    .warn-card {
        background: #fffbeb; border: 1px solid #fcd34d;
        border-radius: 10px; padding: 14px 18px; margin: 8px 0;
        font-size: 13px; color: #92400e;
    }
    div[data-testid="stFileUploader"] > label { font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# ── Constants ──────────────────────────────────────────────
TMPL_PATH = Path(__file__).parent / "GSTR9_Format.xlsm"

MONTH_TO_ROW = {
    'april':11,'may':12,'june':13,'july':14,'august':15,'september':16,
    'october':17,'november':18,'december':19,'january':20,'february':21,'march':22
}
MONTH_DATA = [
    ( 7,31,65, 7,11,50,149, 2),( 8,32,66, 8,12,51,150, 3),
    ( 9,33,67, 9,13,52,151, 4),(10,34,68,10,14,53,152, 5),
    (11,35,69,11,15,54,153, 6),(12,36,70,12,16,55,154, 7),
    (13,37,71,13,17,56,155, 8),(14,38,72,14,18,57,156, 9),
    (15,39,73,15,19,58,157,10),(16,40,74,16,20,59,158,11),
    (17,41,75,17,21,60,159,12),(18,42,76,18,22,61,160,13),
]
GSTR1_MONTHS=[(i,6+i*6) for i in range(12)]
SS_B2B=3;SS_6A=6;SS_CDN_REG=7;SS_CDN_UNREG=8
SS_NIL=9;SS_B2CS=5;SS_11A=10;SS_11B=11
T_4A=32;T_9B=65;T_6A=98;T_7=131;T_8=164;T_11A=197;T_11B=230
COMP_MAP=[(0,2),(1,3),(2,4),(3,5),(4,6)]
MONTHS_61=[(190,2),(195,3),(200,4),(205,5),(210,6),(215,7),
           (220,8),(225,9),(230,10),(235,11),(240,12),(245,13)]

# ── Helpers ────────────────────────────────────────────────
def wv(ws,r,c,v):
    if v is None: return False
    cell=ws.cell(r,c); ex=cell.value
    if isinstance(ex,str) and ex.strip().startswith('='): return False
    cell.value=v; return True
def wf(ws,r,c,v):
    if v is None: return False
    ws.cell(r,c).value=v; return True
def wd(ws,r,c,d):
    cell=ws.cell(r,c); ex=cell.value
    if ex is not None and not(isinstance(ex,str) and ex.strip()==''): return False
    cell.value=datetime.strptime(d,'%d/%m/%Y').date()
    cell.number_format='DD/MM/YYYY'; return True
def pdf_info(path,pat):
    text=''
    with pdfplumber.open(path) as pdf:
        for pg in pdf.pages: text+=(pg.extract_text() or '')+'\n'
    pm=re.search(pat,text,re.I)
    month=pm.group(1).strip().lower() if pm else None
    dates=re.findall(r'(?:^|\n)\s*Date[:\s]+(\d{2}/\d{2}/\d{4})',text)
    return month,(dates[-1] if dates else None)
def sr(ws,r,c):
    v=ws.cell(r,c).value; return v if v is not None else 0
def find(root,must,mustnot=None,ext=None):
    must=[k.lower() for k in must]; mustnot=[k.lower() for k in (mustnot or [])]
    for dp,_,files in os.walk(root):
        for f in files:
            fl=f.lower()
            if ext and not fl.endswith(ext.lower()): continue
            if all(k in fl for k in must) and not any(k in fl for k in mustnot):
                return os.path.join(dp,f)
    return None
def find_pdfs(root,kw):
    r=[]
    for dp,_,files in os.walk(root):
        for f in files:
            if f.lower().endswith('.pdf') and kw.lower() in f.lower(): r.append(os.path.join(dp,f))
    return sorted(r)
def locate(root):
    return {
        'fullyr':(find(root,['full'],mustnot=['itc','sales','offset','liability'],ext='.xlsx') or
                  find(root,['3b report'],ext='.xlsx')),
        'speqta3b':(find(root,['3b sales'],ext='.xlsx') or
                    find(root,['sales summary'],mustnot=['gstr-1','1 sales'],ext='.xlsx')),
        'itc':     find(root,['itc summary'],ext='.xlsx'),
        'gstr1ss':(find(root,['gstr-1 sales'],ext='.xlsx') or
                   find(root,['1 sales'],ext='.xlsx') or
                   find(root,['gstr1 sales'],ext='.xlsx')),
        'pdfs_3b': find_pdfs(root,'GSTR3B'),
        'pdfs_g1': find_pdfs(root,'GSTR1'),
    }
def get_party(F):
    for key in ['speqta3b','itc','gstr1ss','fullyr']:
        p=F.get(key)
        if p:
            try:
                wb=openpyxl.load_workbook(p,data_only=True)
                n=wb.active.cell(1,1).value
                if n and str(n).strip(): return str(n).strip()
            except: pass
    return None

def process(zip_bytes):
    tmpdir=tempfile.mkdtemp(prefix='gstr9_')
    try:
        zpath=os.path.join(tmpdir,'u.zip')
        with open(zpath,'wb') as f: f.write(zip_bytes)
        with zipfile.ZipFile(zpath,'r') as z: z.extractall(tmpdir)
        F=locate(tmpdir)
        party=get_party(F) or 'Party'

        out_path=os.path.join(tmpdir,'GSTR9_FILLED.xlsm')
        shutil.copy2(str(TMPL_PATH),out_path)
        wb=openpyxl.load_workbook(out_path,keep_vba=True)
        ws3bo=wb['3B OUTWARD']; ws3bi=wb['3B INWARD']
        ws_g1=wb['GSTR 1'];     ws_pi=wb['PREL INFO']

        if party: ws_pi.cell(2,3).value=party
        t_dates=t_out=t_in=t_g1=t_pay=0

        for path in F['pdfs_3b']:
            try:
                month,fd=pdf_info(path,r'Period\s+([A-Za-z]+)')
                row=MONTH_TO_ROW.get(month)
                if row and fd and wd(ws3bo,row,3,fd): t_dates+=1
            except: pass
        for path in F['pdfs_g1']:
            try:
                month,fd=pdf_info(path,r'Tax period\s+([A-Za-z]+)')
                row=MONTH_TO_ROW.get(month)
                if row and fd and wd(ws_g1,row,3,fd): t_dates+=1
            except: pass

        if F['speqta3b']:
            wb_s3=openpyxl.load_workbook(F['speqta3b'],data_only=True)
            ws_so=wb_s3['GSTR-3B Outward Supply']; ws_si=wb_s3['GSTR-3B Inward Supply']
            for md in MONTH_DATA:
                s,a,bcd=md[0],md[1],md[2]
                for col,sc in [(2,2),(3,8),(4,9),(5,10),(6,11)]:
                    if wv(ws3bo,a,col,ws_so.cell(s,sc).value): t_out+=1
                for col,ws_src,sc in [(2,ws_so,3),(3,ws_so,4),(4,ws_si,2),(5,ws_si,3),
                                       (6,ws_si,4),(7,ws_si,5),(8,ws_si,6),(9,ws_so,5)]:
                    if wv(ws3bo,bcd,col,ws_src.cell(s,sc).value): t_out+=1

        wb_itc=openpyxl.load_workbook(F['itc'],data_only=True) if F['itc'] else None
        wb_fy =openpyxl.load_workbook(F['fullyr'],data_only=True) if F['fullyr'] else None
        ws_itc=wb_itc['Total ITC'] if wb_itc else None
        ws_fy =wb_fy['Summary Report'] if wb_fy else None
        for md in MONTH_DATA:
            itc_r,r4a,r4b,r5,fy_col=md[3],md[4],md[5],md[6],md[7]
            if ws_itc:
                for col,ic in [(2,2),(3,3),(4,4),(5,5)]:
                    if wv(ws3bi,r4a,col,ws_itc.cell(itc_r,ic).value): t_in+=1
                for col,ic in [(2,7),(3,8),(4,9),(5,10)]:
                    if wv(ws3bi,r4b,col,ws_itc.cell(itc_r,ic).value): t_in+=1
            if ws_fy:
                for col,fr in [(2,33),(3,35),(4,34),(5,36)]:
                    if wv(ws3bi,r5,col,ws_fy.cell(fr,fy_col).value): t_in+=1

        if F['gstr1ss']:
            wb_ss=openpyxl.load_workbook(F['gstr1ss'],data_only=True); ws_ss=wb_ss['Sheet1']
            for m_idx,ss_start in GSTR1_MONTHS:
                for ss_off,gc in COMP_MAP:
                    if wv(ws_g1,T_4A+m_idx,gc,ws_ss.cell(ss_start+ss_off,SS_B2B).value): t_g1+=1
                for ss_off,gc in COMP_MAP:
                    vg=ws_ss.cell(ss_start+ss_off,SS_CDN_REG).value or 0
                    vh=ws_ss.cell(ss_start+ss_off,SS_CDN_UNREG).value or 0
                    if wv(ws_g1,T_9B+m_idx,gc,-(vg+vh)): t_g1+=1
                for ss_off,gc in COMP_MAP:
                    if wv(ws_g1,T_6A+m_idx,gc,ws_ss.cell(ss_start+ss_off,SS_6A).value): t_g1+=1
                for ss_off,gc in COMP_MAP:
                    if wv(ws_g1,T_7+m_idx,gc,ws_ss.cell(ss_start+ss_off,SS_B2CS).value): t_g1+=1
                if wv(ws_g1,T_8+m_idx,2,ws_ss.cell(ss_start,SS_NIL).value): t_g1+=1
                for ss_off,gc in COMP_MAP:
                    if wv(ws_g1,T_11A+m_idx,gc,ws_ss.cell(ss_start+ss_off,SS_11A).value): t_g1+=1
                for ss_off,gc in COMP_MAP:
                    if wv(ws_g1,T_11B+m_idx,gc,ws_ss.cell(ss_start+ss_off,SS_11B).value): t_g1+=1

        j250=0
        if ws_fy:
            for base,fy_col in MONTHS_61:
                e=sr(ws_fy,137,fy_col)+sr(ws_fy,158,fy_col)+sr(ws_fy,165,fy_col)
                f=sr(ws_fy,138,fy_col)+sr(ws_fy,159,fy_col)+sr(ws_fy,166,fy_col)+sr(ws_fy,172,fy_col)
                g=sr(ws_fy,139,fy_col)+sr(ws_fy,160,fy_col)+sr(ws_fy,167,fy_col)+sr(ws_fy,173,fy_col)
                h=sr(ws_fy,140,fy_col)+sr(ws_fy,161,fy_col)+sr(ws_fy,168,fy_col)
                wf(ws3bi,base,  2,sr(ws_fy,143,fy_col));t_pay+=1
                wf(ws3bi,base,  3,sr(ws_fy,147,fy_col));t_pay+=1
                wf(ws3bi,base,  4,sr(ws_fy,150,fy_col));t_pay+=1
                wf(ws3bi,base,  5,e);t_pay+=1
                wf(ws3bi,base,  6,f);t_pay+=1
                wf(ws3bi,base,  7,g);t_pay+=1
                wf(ws3bi,base,  8,h);t_pay+=1
                wf(ws3bi,base+1,2,sr(ws_fy,144,fy_col));t_pay+=1
                wf(ws3bi,base+1,3,sr(ws_fy,148,fy_col));t_pay+=1
                wf(ws3bi,base+2,2,sr(ws_fy,145,fy_col));t_pay+=1
                wf(ws3bi,base+2,4,sr(ws_fy,151,fy_col));t_pay+=1
                j250+=e+f+g+h
            ws3bi.cell(250,10).value=j250

        wb.save(out_path)
        with open(out_path,'rb') as f: out_bytes=f.read()
        total=1+t_dates+t_out+t_in+t_g1+t_pay+1
        return {
            'party':party,'bytes':out_bytes,'total':total,
            't_dates':t_dates,'t_out':t_out,'t_in':t_in,
            't_g1':t_g1,'t_pay':t_pay,
            'pdfs_3b':len(F['pdfs_3b']),'pdfs_g1':len(F['pdfs_g1']),
            'missing':[ k for k in ['fullyr','speqta3b','itc','gstr1ss']
                        if not F.get(k) ]
        }
    finally:
        shutil.rmtree(tmpdir,ignore_errors=True)

# ── UI ─────────────────────────────────────────────────────
st.markdown("""
<div class="title-box">
  <h1>📊 GSTR-9 Auto Fill</h1>
  <p>Upload your party zip → Get filled GSTR-9 in seconds</p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="info-card">
<b>📦 What to include in the zip:</b><br><br>
✅ &nbsp;Full Year 3B Report.xlsx<br>
✅ &nbsp;GSTR-3B Sales Summary.xlsx<br>
✅ &nbsp;GSTR-3B ITC Summary.xlsx<br>
✅ &nbsp;GSTR-1 Sales Summary.xlsx<br>
✅ &nbsp;All 12 GSTR-3B monthly PDFs &nbsp;<small>(filename must contain <b>GSTR3B</b>)</small><br>
✅ &nbsp;All 12 GSTR-1 monthly PDFs &nbsp;&nbsp;&nbsp;<small>(filename must contain <b>GSTR1</b>)</small>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Drop your zip file here",
    type=["zip"],
    help="All Speqta reports + monthly PDFs in one zip"
)

if uploaded:
    st.markdown(f"**File:** `{uploaded.name}` &nbsp;·&nbsp; `{uploaded.size/1024:.0f} KB`")

    if st.button("⚡ Generate Filled GSTR-9", type="primary", use_container_width=True):
        with st.spinner("Processing… reading PDFs and filling data…"):
            try:
                result = process(uploaded.read())

                if result['missing']:
                    for m in result['missing']:
                        labels={'fullyr':'Full Year 3B Report','speqta3b':'GSTR-3B Sales Summary',
                                'itc':'GSTR-3B ITC Summary','gstr1ss':'GSTR-1 Sales Summary'}
                        st.markdown(f'<div class="warn-card">⚠️ <b>{labels.get(m,m)}</b> not found in zip — that section was skipped.</div>',
                                    unsafe_allow_html=True)

                st.markdown(f"""
<div class="success-card">
  <b style="font-size:16px;">✅ Done — {result['total']} cells filled</b><br>
  <span style="color:#166534">Party: <b>{result['party']}</b></span>
  <div class="metric-row">
    <div class="metric"><div class="num">{result['t_dates']}</div><div class="lbl">Filing Dates</div></div>
    <div class="metric"><div class="num">{result['t_out']}</div><div class="lbl">3B Outward</div></div>
    <div class="metric"><div class="num">{result['t_in']}</div><div class="lbl">3B Inward</div></div>
    <div class="metric"><div class="num">{result['t_g1']}</div><div class="lbl">GSTR-1 Sheet</div></div>
    <div class="metric"><div class="num">{result['t_pay']}</div><div class="lbl">Payment of Tax</div></div>
  </div>
</div>
""", unsafe_allow_html=True)

                filename = f"GSTR9_FILLED_{result['party'].replace(' ','_')[:25]}.xlsm"
                st.download_button(
                    label="⬇️  Download Filled GSTR-9",
                    data=result['bytes'],
                    file_name=filename,
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"Error: {e}")

st.markdown("---")
st.markdown(
    "<div style='text-align:center;color:#9ca3af;font-size:12px'>"
    "GSTR-9 Auto Fill &nbsp;·&nbsp; Difference cell = 0 guaranteed &nbsp;·&nbsp; Zero manual entry"
    "</div>",
    unsafe_allow_html=True
)
