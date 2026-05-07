"""
Fragrance Wholesale Price Analysis — Streamlit Web App
Upload wholesaler files once, then upload any offer file to get instant analysis.
"""

import streamlit as st
import pandas as pd
import os, sys, io, glob, tempfile, shutil, subprocess
from pathlib import Path

st.set_page_config(
    page_title="Fragrance Price Intelligence",
    page_icon="🧴",
    layout="wide"
)

# ─── Paths ────────────────────────────────────────────────────────────────────
BASE_DIR  = Path(tempfile.gettempdir()) / "fragrance_tool"
DB_DIR    = BASE_DIR / "wholesalers"
DB_DIR.mkdir(parents=True, exist_ok=True)

# ─── Styles ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-title { font-size: 2rem; font-weight: 700; color: #1A3A5C; margin-bottom: 0; }
    .sub-title  { font-size: 1rem; color: #64748B; margin-bottom: 1.5rem; }
    .stat-box   { background: #F4F7FB; border-radius: 8px; padding: 1rem; text-align: center; border: 1px solid #E2E8F0; }
    .stat-num   { font-size: 1.8rem; font-weight: 700; color: #1A3A5C; }
    .stat-label { font-size: 0.8rem; color: #64748B; }
    .verdict-good { background: #E8F5E9; border-left: 4px solid #1B5E20; padding: 1rem; border-radius: 4px; }
    .verdict-bad  { background: #FEE2E2; border-left: 4px solid #C00000; padding: 1rem; border-radius: 4px; }
    .verdict-ok   { background: #FFF3E0; border-left: 4px solid #E65100; padding: 1rem; border-radius: 4px; }
</style>
""", unsafe_allow_html=True)

# ─── Header ───────────────────────────────────────────────────────────────────
st.markdown('<div class="main-title">🧴 Fragrance Price Intelligence</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Instant wholesale benchmarking across MTZ · Nandansons · PCA · GE · PTC</div>', unsafe_allow_html=True)

# ─── Sidebar: Wholesaler file upload ─────────────────────────────────────────
with st.sidebar:
    st.header("📦 Wholesaler Databases")
    st.caption("Upload once — they stay loaded for the session.")

    st.subheader("MTZ")
    mtz_files = st.file_uploader("MTZpricelist files", type=['xlsx'], accept_multiple_files=True, key='mtz')

    st.subheader("Nandansons")
    nan_files = st.file_uploader("Nandansons files (.xlsx or .xls)", type=['xlsx','xls'], accept_multiple_files=True, key='nan')

    st.subheader("PCA")
    pca_files = st.file_uploader("PCA price list (.xls or .xlsx)", type=['xlsx','xls'], accept_multiple_files=True, key='pca')

    st.subheader("GE")
    ge_files = st.file_uploader("GE wholesale files (.xls or .xlsx)", type=['xlsx','xls'], accept_multiple_files=True, key='ge')

    st.subheader("PTC")
    ptc_files = st.file_uploader("PTC files (.xlsx)", type=['xlsx'], accept_multiple_files=True, key='ptc')

    if st.button("💾 Save Wholesaler Files", type="primary", use_container_width=True):
        saved = 0
        for group_name, files in [('MTZ', mtz_files), ('NAN', nan_files), ('PCA', pca_files),
                                    ('GE', ge_files), ('PTC', ptc_files)]:
            for f in (files or []):
                dest = DB_DIR / f"{group_name}_{f.name}"
                dest.write_bytes(f.read())
                saved += 1
        st.success(f"✅ Saved {saved} files")

    # Show loaded files
    loaded = list(DB_DIR.glob("*.xlsx")) + list(DB_DIR.glob("*.xls"))
    if loaded:
        st.caption(f"📁 {len(loaded)} wholesaler files loaded")

# ─── Main: Offer file + settings ─────────────────────────────────────────────
col1, col2 = st.columns([2, 1])

with col1:
    st.header("📂 Upload Offer File")
    offer_file = st.file_uploader("Upload any supplier price list", type=['xlsx','xls','csv'], key='offer')

with col2:
    st.header("⚙️ Settings")
    eur_convert = st.checkbox("Convert EUR → USD", value=False)
    if eur_convert:
        rate = st.number_input("Exchange rate", value=1.16, step=0.01)
    else:
        rate = 1.0
    make_package = st.checkbox("Generate Package Price report", value=True)

# ─── Helper: convert .xls to .xlsx ───────────────────────────────────────────
def convert_xls(src_path: Path) -> Path:
    """Convert .xls to .xlsx using openpyxl via xlrd"""
    if src_path.suffix.lower() == '.xlsx':
        return src_path
    try:
        import xlrd
        from openpyxl import Workbook as OWB
        wb_xls = xlrd.open_workbook(str(src_path))
        wb_new = OWB()
        wb_new.remove(wb_new.active)
        for sheet in wb_xls.sheets():
            ws = wb_new.create_sheet(title=sheet.name)
            for row in range(sheet.nrows):
                for col in range(sheet.ncols):
                    ws.cell(row+1, col+1, sheet.cell_value(row, col))
        out = src_path.with_suffix('.xlsx')
        wb_new.save(str(out))
        return out
    except Exception:
        return src_path

# ─── Helper: normalize UPC ───────────────────────────────────────────────────
def norm_upc(v):
    try:
        s = str(v).strip().replace('.0','')
        return str(int(float(s))) if s.replace('.','').replace('e+','').replace('E+','').isdigit() else s
    except:
        return str(v).strip()

# ─── Load wholesalers ─────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_all_wholesalers(file_list_hash):
    """Load all wholesaler files from DB_DIR into {name: {upc: price}}"""
    dbs = {'MTZ': {}, 'NAN': {}, 'PCA': {}, 'GE': {}, 'PTC': {}}
    raw_descs = {'MTZ': [], 'NAN': [], 'PCA': [], 'GE': [], 'PTC': []}

    for f in DB_DIR.glob("*"):
        prefix = f.name.split('_')[0].upper()
        if prefix not in dbs:
            continue
        fpath = convert_xls(f)

        try:
            # MTZ
            if prefix == 'MTZ':
                raw = pd.read_excel(fpath, sheet_name='Price', header=None)
                date = str(raw.iloc[2, 2])[:10]
                df = pd.read_excel(fpath, sheet_name='Price', header=4)
                upc   = df.iloc[:,3].apply(norm_upc)
                price = pd.to_numeric(df.iloc[:,4], errors='coerce')
                desc  = df.iloc[:,2].astype(str)
                for u, p, d in zip(upc, price, desc):
                    if pd.notna(p) and len(u) >= 8:
                        if u not in dbs['MTZ']: dbs['MTZ'][u] = p
                        raw_descs['MTZ'].append((d.upper(), u, p))

            # NAN
            elif prefix == 'NAN':
                df = pd.read_excel(fpath, header=0)
                upc   = df.iloc[:,0].apply(norm_upc)
                ncols = df.shape[1]
                pcol  = 7 if ncols >= 8 else 5
                price = pd.to_numeric(df.iloc[:,pcol], errors='coerce')
                desc  = df.iloc[:,1].astype(str) if df.shape[1] > 1 else pd.Series([''] * len(df))
                for u, p, d in zip(upc, price, desc):
                    if pd.notna(p) and len(u) >= 8:
                        if u not in dbs['NAN']: dbs['NAN'][u] = p
                        raw_descs['NAN'].append((d.upper(), u, p))

            # PCA
            elif prefix == 'PCA':
                df = pd.read_excel(fpath, header=0)
                upc   = df.iloc[:,6].apply(norm_upc)
                price = pd.to_numeric(df.iloc[:,8], errors='coerce')
                desc  = df.iloc[:,2].astype(str)
                for u, p, d in zip(upc, price, desc):
                    if pd.notna(p) and len(u) >= 8:
                        if u not in dbs['PCA']: dbs['PCA'][u] = p
                        raw_descs['PCA'].append((d.upper(), u, p))

            # GE
            elif prefix == 'GE':
                df = pd.read_excel(fpath, header=0)
                upc   = df.iloc[:,1].apply(norm_upc)
                price = pd.to_numeric(df.iloc[:,4], errors='coerce')
                desc  = df.iloc[:,3].astype(str) if df.shape[1] > 3 else pd.Series([''] * len(df))
                for u, p, d in zip(upc, price, desc):
                    if pd.notna(p) and len(u) >= 8:
                        if u not in dbs['GE']: dbs['GE'][u] = p
                        raw_descs['GE'].append((d.upper(), u, p))

            # PTC
            elif prefix == 'PTC':
                for hrow in [11, 8, 7, 6]:
                    try:
                        sheet = 'Price List' if 'Price List' in pd.ExcelFile(fpath).sheet_names else 'Sheet1'
                        df = pd.read_excel(fpath, sheet_name=sheet, header=hrow)
                        upc   = df.iloc[:,1].apply(norm_upc)
                        price = pd.to_numeric(df.iloc[:,4], errors='coerce')
                        desc  = df.iloc[:,0].astype(str)
                        good  = [(u, p, d) for u, p, d in zip(upc, price, desc) if pd.notna(p) and len(u) >= 8]
                        if len(good) > 5:
                            for u, p, d in good:
                                if u not in dbs['PTC']: dbs['PTC'][u] = p
                                raw_descs['PTC'].append((d.upper(), u, p))
                            break
                    except: continue

        except Exception as e:
            st.warning(f"⚠️ Could not load {f.name}: {e}")

    return dbs, raw_descs

# ─── Keyword fallback match ───────────────────────────────────────────────────
SKIP = {'SPRAY','TESTER','LADIES','EAU','EDT','EDP','FOR','THE','AND','WITH',
        '100','200','50','75','90','WOMEN','PARFUM','MEN','WOMAN','ML','OZ','PC'}

def kw_match(desc: str, raw_descs: dict):
    words = [w for w in desc.upper().split() if w not in SKIP and len(w) > 2][:3]
    if not words: return {}
    results = {}
    for ws_name, records in raw_descs.items():
        for (d, upc, price) in records:
            if all(w in d for w in words):
                if ws_name not in results:
                    results[ws_name] = (upc, price, d[:50])
                break
    return results

# ─── Analyse offer ────────────────────────────────────────────────────────────
def analyse_offer(offer_df, dbs, raw_descs, eur_rate=1.0):
    rows = []
    for _, item in offer_df.iterrows():
        upc   = norm_upc(item.get('UPC', ''))
        price = float(item.get('My_Price', 0) or 0) * eur_rate
        desc  = str(item.get('Description', ''))
        qty   = float(item.get('QTY', 1) or 1)

        ws_prices = {}
        for ws_name, db in dbs.items():
            if upc in db:
                ws_prices[ws_name] = db[upc]

        # keyword fallback
        if not ws_prices and desc and desc != 'nan':
            matches = kw_match(desc, raw_descs)
            for ws_name, (matched_upc, matched_price, matched_desc) in matches.items():
                ws_prices[ws_name] = matched_price

        avg = sum(ws_prices.values()) / len(ws_prices) if ws_prices else None
        disc = ((price / avg) - 1) * 100 if avg and price > 0 else None

        rows.append({
            'UPC': upc, 'Description': desc, 'QTY': qty,
            'My_Price': price, 'Avg_Price': avg,
            'Disc_Pct': disc, 'Suppliers_Found': len(ws_prices),
            **{f'{ws}_Price': ws_prices.get(ws) for ws in dbs.keys()}
        })
    return pd.DataFrame(rows)

# ─── Package pricing ──────────────────────────────────────────────────────────
def build_package_price(analysis_df):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    matched   = analysis_df[analysis_df['Suppliers_Found'] > 0].copy()
    unmatched = analysis_df[analysis_df['Suppliers_Found'] == 0].copy()
    if matched.empty: return None

    matched['QTY'] = matched['QTY'].fillna(1)
    total_offer = (matched['My_Price'] * matched['QTY']).sum()
    total_avg   = (matched['Avg_Price'] * matched['QTY']).sum()
    if total_avg == 0: return None

    k            = total_offer / total_avg
    disc_pct     = (k - 1) * 100
    meets        = disc_pct <= -30
    min_pkg      = total_avg * 0.70
    gap          = total_offer - min_pkg

    matched['New_Price']    = (matched['Avg_Price'] * k).round(2)
    matched['New_Pct']      = ((matched['New_Price'] / matched['Avg_Price'] - 1) * 100).round(1)
    matched['Old_Price']    = matched['My_Price']
    matched['Price_Change'] = (matched['New_Price'] - matched['Old_Price']).round(2)
    matched['T30'] = (matched['Avg_Price'] * 0.70).round(2)
    matched['T35'] = (matched['Avg_Price'] * 0.65).round(2)
    matched['T40'] = (matched['Avg_Price'] * 0.60).round(2)

    new_total = (matched['New_Price'] * matched['QTY']).sum()
    t30_total = (matched['T30'] * matched['QTY']).sum()
    t35_total = (matched['T35'] * matched['QTY']).sum()
    t40_total = (matched['T40'] * matched['QTY']).sum()

    # Styles
    F_DARK  = PatternFill('solid', start_color='1A3A5C')
    F_DGRN  = PatternFill('solid', start_color='1B5E20')
    F_C00   = PatternFill('solid', start_color='C00000')
    F_GRN   = PatternFill('solid', start_color='C6EFCE')
    F_RED   = PatternFill('solid', start_color='FFC7CE')
    F_PURP  = PatternFill('solid', start_color='EDE0F5')
    F_TEAL  = PatternFill('solid', start_color='E0F2F1')
    F_ALT   = PatternFill('solid', start_color='F4F7FB')
    F_LGRY  = PatternFill('solid', start_color='F8F8F8')
    F_AMBR  = PatternFill('solid', start_color='FFF3E0')
    F_BBLUE = PatternFill('solid', start_color='E3F2FD')
    F_BGRN2 = PatternFill('solid', start_color='E8F5E9')
    F_ORG   = PatternFill('solid', start_color='E65100')
    F_SLATE = PatternFill('solid', start_color='546E7A')
    F_BLUE  = PatternFill('solid', start_color='1565C0')

    def fw(sz=10, color='FFFFFF', bold=True): return Font(name='Calibri', bold=bold, size=sz, color=color)
    def fr(sz=10, color='000000'): return Font(name='Calibri', bold=False, size=sz, color=color)
    def fb(sz=10, color='000000'): return Font(name='Calibri', bold=True, size=sz, color=color)

    _t = Side(style='thin', color='CCCCCC')
    _m = Side(style='medium', color='888888')
    BDR  = Border(left=_t, right=_t, top=_t, bottom=_t)
    MBDR = Border(left=_m, right=_m, top=_m, bottom=_m)
    CTR  = Alignment(horizontal='center', vertical='center', wrap_text=True)
    LFT  = Alignment(horizontal='left',   vertical='center', wrap_text=True)
    RGT  = Alignment(horizontal='right',  vertical='center')

    wb = Workbook()
    ws = wb.active
    ws.title = 'Package Pricing'

    def cel(row, col, val, font=None, fill=None, align=None, border=None, fmt=None):
        c = ws.cell(row=row, column=col, value=val)
        if font:   c.font   = font
        if fill:   c.fill   = fill
        if align:  c.alignment = align
        if border: c.border = border
        if fmt:    c.number_format = fmt
        return c

    for i, w in enumerate([18, 44, 8, 12, 12, 13, 13, 11, 7, 22, 13, 13, 13], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Title
    ws.merge_cells('A1:M1')
    ws.cell(1,1,'PACKAGE PRICING — Equal Competitiveness Model').font = Font(name='Calibri', bold=True, size=14, color='1A3A5C')
    ws.cell(1,1).alignment = CTR
    ws.row_dimensions[1].height = 30

    ws.merge_cells('A2:M2')
    ws.cell(2,1,'Prices redistributed so every item sits at the same % discount vs industry average, keeping total package value identical.').font = Font(name='Calibri', size=9, italic=True, color='666666')
    ws.cell(2,1).alignment = LFT
    ws.row_dimensions[2].height = 16

    # Banner
    ws.merge_cells('A4:M4')
    if meets:
        banner = f'✅  PACKAGE MEETS TARGET  |  Discount: {disc_pct:.1f}%  (target: ≥−30%)'
        bfill  = F_DGRN
    else:
        banner = f'⚠️  PACKAGE BELOW TARGET  |  Current discount: {disc_pct:.1f}%  |  Need to reduce by ${gap:,.0f}  →  Min: ${min_pkg:,.0f}'
        bfill  = F_C00
    ws.cell(4,1,banner).font  = fw(11,'FFFFFF')
    ws.cell(4,1).fill         = bfill
    ws.cell(4,1).alignment    = CTR
    ws.row_dimensions[4].height = 22

    # Summary
    ws.merge_cells('A5:M5')
    ws.cell(5,1,'PACKAGE SUMMARY').font = fw(10,'FFFFFF')
    ws.cell(5,1).fill = F_DARK; ws.cell(5,1).alignment = LFT
    ws.row_dimensions[5].height = 18

    stats = [
        ('Matched items repriced', len(matched), '', 'Uniform discount vs market', f'{disc_pct:+.2f}%', ''),
        ('Unmatched (price unchanged)', len(unmatched), '', 'Target: meets −30%?', '✅ YES' if meets else '❌ NO', ''),
        ('Original total (matched)', round(total_offer,2), '$#,##0.00', 'Min pkg for −30%', round(min_pkg,2), '$#,##0.00'),
        ('New total (matched)', round(new_total,2), '$#,##0.00', 'Gap to target', round(gap,2), '$#,##0.00'),
    ]
    for i, (ll, lv, lf, rl, rv, rf) in enumerate(stats):
        r = 6+i
        cel(r,1,ll,fr(),None,LFT); c2=cel(r,2,lv,fb(),None,RGT)
        if lf: c2.number_format=lf
        cel(r,4,rl,fr(),None,LFT); c5=cel(r,5,rv,fb(),None,RGT)
        if rf: c5.number_format=rf
        ws.row_dimensions[r].height = 16

    ws.merge_cells('A11:M11')
    ws.cell(11,1,'Why 30%?  US Tariff: ~15%  |  Shipping: ~5%  |  Client margin: ~10–15%  →  Total buffer: ~30–35%').font = Font(name='Calibri', size=9, italic=True, color='444444')
    ws.cell(11,1).fill = F_TEAL; ws.cell(11,1).alignment = LFT
    ws.row_dimensions[11].height = 16

    ws.merge_cells('A13:M13')
    ws.cell(13,1,'REPRICED ITEMS').font = fw(10,'FFFFFF')
    ws.cell(13,1).fill = F_DGRN; ws.cell(13,1).alignment = LFT
    ws.row_dimensions[13].height = 18

    # Headers — 13 columns
    hdrs = ['EAN/UPC','Description','QTY','Original $','Ind. Avg $','New Price $',
            'Change $/pc','Disc. vs Avg','# Src','Notes',
            'Target −30%\n(Minimum)','Target −35%\n(Good)','Target −40%\n(Sharp)']
    tier_fills = {11: F_ORG, 12: F_BLUE, 13: F_DGRN}
    for col, h in enumerate(hdrs, 1):
        cel(14, col, h, fw(10,'FFFFFF'), tier_fills.get(col, F_DARK), CTR, BDR)
    ws.row_dimensions[14].height = 32

    for i, (_, row) in enumerate(matched.iterrows()):
        r    = 15 + i
        fill = F_ALT if i % 2 == 0 else F_LGRY
        chg  = row['Price_Change']

        cel(r,1,  row.get('UPC',''),         fr(10), fill,  LFT, BDR)
        cel(r,2,  row.get('Description',''), fr(10), fill,  LFT, BDR)
        cel(r,3,  row.get('QTY',1),          fr(10), fill,  RGT, BDR, '#,##0')
        cel(r,4,  row['Old_Price'],           fr(10), fill,  RGT, BDR, '$#,##0.00')
        cel(r,5,  row['Avg_Price'],           fr(10), fill,  RGT, BDR, '$#,##0.00')
        cel(r,6,  row['New_Price'],           fb(10), F_GRN, RGT, BDR, '$#,##0.00')
        chg_fill = F_GRN if chg > 0 else (F_RED if chg < 0 else fill)
        chg_font = fb(10,'1B5E20') if chg > 0 else (fb(10,'C00000') if chg < 0 else fr(10))
        cel(r,7,  chg, chg_font, chg_fill,   RGT, BDR, '+$#,##0.00;-$#,##0.00;$-')
        cel(r,8,  row['New_Pct']/100, fb(10,'1B5E20'), F_GRN, CTR, BDR, '+0.0%;-0.0%;0.0%')
        cel(r,9,  int(row['Suppliers_Found']), fr(10), fill, CTR, BDR)
        if abs(chg) < 0.01:   note = 'Unchanged'
        elif chg > 0:         note = f'↑ +${chg:.2f}/pc'
        else:                 note = f'↓ −${abs(chg):.2f}/pc'
        cel(r,10, note, Font(name='Calibri', size=9, italic=True, color='666666'), fill, LFT, BDR)
        cel(r,11, row['T30'], fb(10,'E65100'), F_AMBR,  RGT, BDR, '$#,##0.00')
        cel(r,12, row['T35'], fb(10,'1565C0'), F_BBLUE, RGT, BDR, '$#,##0.00')
        cel(r,13, row['T40'], fb(10,'1B5E20'), F_BGRN2, RGT, BDR, '$#,##0.00')
        ws.row_dimensions[r].height = 18

    # Totals
    r_tot = 15 + len(matched)
    cel(r_tot,1,'TOTAL',fb(10),F_PURP,LFT,MBDR)
    cel(r_tot,2,'',fr(10),F_PURP,LFT,MBDR)
    cel(r_tot,3,int(matched['QTY'].sum()),fb(10),F_PURP,RGT,MBDR,'#,##0')
    cel(r_tot,4,round(total_offer,2),fb(10),F_PURP,RGT,MBDR,'$#,##0.00')
    cel(r_tot,5,round(total_avg,2),fb(10),F_PURP,RGT,MBDR,'$#,##0.00')
    cel(r_tot,6,round(new_total,2),fb(10),F_PURP,RGT,MBDR,'$#,##0.00')
    cel(r_tot,7,round(new_total-total_offer,2),fb(10),F_PURP,RGT,MBDR,'+$#,##0.00;-$#,##0.00;$-')
    cel(r_tot,8,f'{disc_pct:+.2f}%',fb(10),F_GRN if meets else F_RED,CTR,MBDR)
    cel(r_tot,9,'',fr(10),F_PURP,CTR,MBDR)
    cel(r_tot,10,'',fr(10),F_PURP,LFT,MBDR)
    cel(r_tot,11,round(t30_total,2),fw(10,'FFFFFF'),F_ORG, RGT,MBDR,'$#,##0.00')
    cel(r_tot,12,round(t35_total,2),fw(10,'FFFFFF'),F_BLUE,RGT,MBDR,'$#,##0.00')
    cel(r_tot,13,round(t40_total,2),fw(10,'FFFFFF'),F_DGRN,RGT,MBDR,'$#,##0.00')
    ws.row_dimensions[r_tot].height = 22

    # Callout
    r_call = r_tot + 2
    ws.merge_cells(f'A{r_call}:M{r_call}')
    if meets:
        call_txt  = f'✅  PACKAGE MEETS −30% TARGET  |  Discount: {disc_pct:.1f}%  |  Total: ${new_total:,.0f}'
        call_fill = F_DGRN
    else:
        call_txt  = f'📌  TO REACH −30%: Package value must be ≤ ${min_pkg:,.0f}  (current: ${new_total:,.0f}  |  reduce by ${gap:,.0f})'
        call_fill = F_ORG
    ws.cell(r_call,1,call_txt).font  = fw(10,'FFFFFF')
    ws.cell(r_call,1).fill           = call_fill
    ws.cell(r_call,1).alignment      = LFT
    ws.row_dimensions[r_call].height = 20

    # Unmatched
    if not unmatched.empty:
        r = r_call + 2
        ws.merge_cells(f'A{r}:M{r}')
        ws.cell(r,1,f'UNMATCHED ITEMS — {len(unmatched)} items (prices unchanged)').font = fw(10,'FFFFFF')
        ws.cell(r,1).fill = F_SLATE; ws.cell(r,1).alignment = LFT
        ws.row_dimensions[r].height = 18; r += 1
        for col, h in enumerate(['EAN/UPC','Description','QTY','Price (unchanged)'],1):
            cel(r,col,h,fw(10,'FFFFFF'),F_DARK,CTR,BDR)
        ws.row_dimensions[r].height = 20
        for j, (_,row_) in enumerate(unmatched.iterrows()):
            r += 1
            fill = F_ALT if j%2==0 else F_LGRY
            cel(r,1,row_.get('UPC',''),fr(10),fill,LFT,BDR)
            cel(r,2,row_.get('Description',''),fr(10),fill,LFT,BDR)
            cel(r,3,row_.get('QTY',''),fr(10),fill,RGT,BDR,'#,##0')
            cel(r,4,row_.get('My_Price',''),fr(10),fill,RGT,BDR,'$#,##0.00')
            ws.row_dimensions[r].height = 18

    ws.freeze_panes = 'A3'
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), disc_pct, meets, gap, min_pkg

# ─── Run analysis ─────────────────────────────────────────────────────────────
if offer_file and st.button("🚀 Run Price Analysis", type="primary", use_container_width=True):

    loaded_files = list(DB_DIR.glob("*"))
    if not loaded_files:
        st.error("❌ No wholesaler files loaded. Please upload wholesaler files in the sidebar first.")
        st.stop()

    with st.spinner("Loading wholesaler databases..."):
        file_hash = str(sorted([f.name for f in loaded_files]))
        dbs, raw_descs = load_all_wholesalers(file_hash)
        total_upcs = sum(len(v) for v in dbs.values())

    # Show DB stats
    st.subheader("📊 Database Loaded")
    cols = st.columns(5)
    for i, (ws_name, db) in enumerate(dbs.items()):
        with cols[i]:
            st.markdown(f'<div class="stat-box"><div class="stat-num">{len(db):,}</div><div class="stat-label">{ws_name} UPCs</div></div>', unsafe_allow_html=True)
    st.markdown("")

    # Load offer
    with st.spinner("Analysing offer..."):
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        tmp.write(offer_file.read())
        tmp.close()

        # Try to read offer file
        try:
            raw = pd.read_excel(tmp.name, header=None)
            header_row = 0
            for i, row in raw.iterrows():
                vals = [str(v).lower() for v in row.values if pd.notna(v)]
                if any('upc' in v or 'ean' in v or 'price' in v for v in vals):
                    header_row = i; break
            offer_df = pd.read_excel(tmp.name, header=header_row)
            offer_df.columns = [str(c).strip() for c in offer_df.columns]

            # Auto-detect columns
            def find_col(df, kws):
                for i, c in enumerate(df.columns):
                    if any(k in str(c).lower() for k in kws): return i
                return None

            upc_i   = find_col(offer_df, ['upc','ean','barcode'])
            price_i = find_col(offer_df, ['price','cost','prix'])
            desc_i  = find_col(offer_df, ['desc','name','product','designation'])
            qty_i   = find_col(offer_df, ['qty','quantity','stock','units'])

            clean = pd.DataFrame()
            if upc_i is not None:
                clean['UPC'] = offer_df.iloc[:,upc_i].apply(norm_upc)
            if price_i is not None:
                clean['My_Price'] = pd.to_numeric(offer_df.iloc[:,price_i], errors='coerce') * rate
            if desc_i is not None:
                clean['Description'] = offer_df.iloc[:,desc_i].astype(str)
            if qty_i is not None:
                clean['QTY'] = pd.to_numeric(offer_df.iloc[:,qty_i], errors='coerce').fillna(1)

            if 'UPC' in clean.columns:
                clean = clean[~clean['UPC'].str.lower().isin(['upc','ean','nan',''])]
                clean = clean.dropna(subset=['UPC'])

        except Exception as e:
            st.error(f"❌ Could not read offer file: {e}")
            st.stop()

        analysis = analyse_offer(clean, dbs, raw_descs, 1.0)  # rate already applied

    # Results
    matched   = analysis[analysis['Suppliers_Found'] > 0]
    unmatched = analysis[analysis['Suppliers_Found'] == 0]

    st.subheader("📈 Results")
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f'<div class="stat-box"><div class="stat-num">{len(analysis)}</div><div class="stat-label">Total Items</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="stat-box"><div class="stat-num">{len(matched)}</div><div class="stat-label">Matched</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="stat-box"><div class="stat-num">{len(unmatched)}</div><div class="stat-label">Not Found</div></div>', unsafe_allow_html=True)

    if not matched.empty and 'My_Price' in matched.columns and 'Avg_Price' in matched.columns:
        total_val = (matched['My_Price'] * matched.get('QTY', 1)).sum()
        total_avg = (matched['Avg_Price'] * matched.get('QTY', 1)).sum()
        pkg_disc  = (total_val / total_avg - 1) * 100 if total_avg else 0
        with c4:
            color = '#1B5E20' if pkg_disc <= -30 else ('#E65100' if pkg_disc <= 0 else '#C00000')
            st.markdown(f'<div class="stat-box"><div class="stat-num" style="color:{color}">{pkg_disc:+.1f}%</div><div class="stat-label">vs Market</div></div>', unsafe_allow_html=True)

        st.markdown("")
        if pkg_disc <= -40:
            st.markdown('<div class="verdict-good">🟢 <strong>SHARP — Excellent deal!</strong> Package is −40%+ below market. Strong buy.</div>', unsafe_allow_html=True)
        elif pkg_disc <= -35:
            st.markdown('<div class="verdict-good">🔵 <strong>GOOD — Solid deal.</strong> Package is −35–40% below market. Worth buying.</div>', unsafe_allow_html=True)
        elif pkg_disc <= -30:
            st.markdown('<div class="verdict-ok">🟠 <strong>MINIMUM — Viable.</strong> Package just clears the 30% threshold. Proceed with caution.</div>', unsafe_allow_html=True)
        elif pkg_disc <= 0:
            st.markdown('<div class="verdict-bad">⚠️ <strong>BELOW TARGET.</strong> Package is below market but not enough. Counter or negotiate down.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="verdict-bad">🔴 <strong>ABOVE MARKET.</strong> Offer price exceeds what wholesalers charge. Hard pass.</div>', unsafe_allow_html=True)

    # Preview table
    st.markdown("### Item Breakdown")
    display_cols = ['UPC','Description','QTY','My_Price','Avg_Price','Disc_Pct','Suppliers_Found']
    display_df = analysis[[c for c in display_cols if c in analysis.columns]].copy()
    if 'Disc_Pct' in display_df:
        display_df['Disc_Pct'] = display_df['Disc_Pct'].apply(lambda x: f'{x:+.1f}%' if pd.notna(x) else '—')
    if 'My_Price' in display_df:
        display_df['My_Price'] = display_df['My_Price'].apply(lambda x: f'${x:.2f}' if pd.notna(x) else '—')
    if 'Avg_Price' in display_df:
        display_df['Avg_Price'] = display_df['Avg_Price'].apply(lambda x: f'${x:.2f}' if pd.notna(x) else '—')
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    # Download buttons
    st.markdown("### 📥 Download Reports")
    dc1, dc2 = st.columns(2)

    # Analysis Excel
    from openpyxl import Workbook as OWB
    buf_analysis = io.BytesIO()
    with pd.ExcelWriter(buf_analysis, engine='openpyxl') as writer:
        analysis.to_excel(writer, sheet_name='Full Analysis', index=False)
        matched.to_excel(writer, sheet_name='Matched Items', index=False)
        unmatched.to_excel(writer, sheet_name='Not Found', index=False)

    with dc1:
        st.download_button(
            "📋 Download Price Analysis",
            data=buf_analysis.getvalue(),
            file_name=f"PriceAnalysis_{offer_file.name.replace('.xlsx','')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    if make_package and not matched.empty:
        result = build_package_price(analysis)
        if result:
            pkg_bytes, disc, meets_target, gap_amt, min_val = result
            with dc2:
                st.download_button(
                    "📦 Download Package Price",
                    data=pkg_bytes,
                    file_name=f"PackagePrice_{offer_file.name.replace('.xlsx','')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    os.unlink(tmp.name)
