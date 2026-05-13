"""
Fragrance Wholesale Price Intelligence — Streamlit Web App
Uses master_price_tool.py as the backend for 100% identical output.
"""

import streamlit as st
import pandas as pd
import os, sys, io, tempfile, shutil, subprocess, importlib.util, traceback
from pathlib import Path

st.set_page_config(page_title="Fragrance Price Intelligence", page_icon="🧴", layout="wide")

UPLOAD_DIR = Path("/mnt/user-data/uploads")
CLAUDE_DIR = Path("/home/claude")
OUTPUT_DIR = Path("/mnt/user-data/outputs")
for d in [UPLOAD_DIR, CLAUDE_DIR, OUTPUT_DIR]:
    d.mkdir(parents=True, exist_ok=True)

st.markdown("""<style>
.main-title{font-size:2rem;font-weight:700;color:#1A3A5C;}
.sub-title{font-size:1rem;color:#64748B;margin-bottom:1.5rem;}
.stat-box{background:#F4F7FB;border-radius:8px;padding:1rem;text-align:center;border:1px solid #E2E8F0;}
.stat-num{font-size:1.8rem;font-weight:700;color:#1A3A5C;}
.stat-label{font-size:0.8rem;color:#64748B;}
.verdict-sharp{background:#E8F5E9;border-left:4px solid #1B5E20;padding:1rem;border-radius:4px;}
.verdict-good{background:#E3F2FD;border-left:4px solid #1565C0;padding:1rem;border-radius:4px;}
.verdict-min{background:#FFF3E0;border-left:4px solid #E65100;padding:1rem;border-radius:4px;}
.verdict-bad{background:#FEE2E2;border-left:4px solid #C00000;padding:1rem;border-radius:4px;}
</style>""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🧴 Fragrance Price Intelligence</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">MTZ · Nandansons · PCA · GE · PTC · 13,500+ UPCs · Instant deal scoring</div>', unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📦 Wholesaler Databases")
    st.caption("Upload once per session. Files are saved automatically.")

    ws_uploads = {}
    for grp, label, exts in [
        ("MTZ",        "MTZ files (.xlsx)",           ["xlsx"]),
        ("Nandansons", "Nandansons files (.xlsx/.xls)",["xlsx","xls"]),
        ("PCA",        "PCA file (.xls/.xlsx)",        ["xls","xlsx"]),
        ("GE",         "GE files (.xls/.xlsx)",        ["xls","xlsx"]),
        ("PTC",        "PTC files (.xlsx)",            ["xlsx"]),
    ]:
        st.subheader(grp)
        ws_uploads[grp] = st.file_uploader(label, type=exts, accept_multiple_files=True, key=f"ws_{grp}") or []

    if st.button("💾 Save & Convert Files", type="primary", use_container_width=True):
        saved = 0
        for grp, files in ws_uploads.items():
            for f in files:
                dest = UPLOAD_DIR / f.name
                dest.write_bytes(f.getvalue())
                saved += 1
                if f.name.endswith(".xls"):
                    subprocess.run(
                        ["libreoffice","--headless","--convert-to","xlsx",
                         str(dest),"--outdir", str(CLAUDE_DIR)],
                        capture_output=True, timeout=60
                    )
        st.success(f"✅ {saved} files saved!")

    st.markdown("---")
    st.caption("**Files detected:**")
    counts = {
        "MTZ":        len(list(UPLOAD_DIR.glob("MTZpricelist*.xlsx"))),
        "Nandansons": len(list(UPLOAD_DIR.glob("Nandansons_*.xlsx"))),
        "PCA":        len(list(CLAUDE_DIR.glob("PRICE-LIST*.xlsx"))),
        "GE":         len(list(CLAUDE_DIR.glob("GE_*.xlsx"))) + len(list(CLAUDE_DIR.glob("WHOLESALE_*.xlsx"))) + len(list(CLAUDE_DIR.glob("Ge_*.xlsx"))),
        "PTC":        len(list(UPLOAD_DIR.glob("PTC_*.xlsx"))),
    }
    for name, cnt in counts.items():
        st.caption(f"{'✅' if cnt > 0 else '❌'} {name}: {cnt} file(s)")
    total_files = sum(counts.values())

    st.markdown("---")
    st.subheader("📤 Upload Tool")
    tool_file = st.file_uploader("master_price_tool.py", type=["py"], key="tool")
    if tool_file:
        (CLAUDE_DIR / "master_price_tool.py").write_bytes(tool_file.getvalue())
        (OUTPUT_DIR / "master_price_tool.py").write_bytes(tool_file.getvalue())
        st.success("✅ Tool saved!")

# ── Main ──────────────────────────────────────────────────────────────────────
c1, c2, c3 = st.columns([3, 1, 1])
with c1:
    offer_file = st.file_uploader("📂 Upload Offer File (.xlsx)", type=["xlsx","xls"], key="offer")
with c2:
    eur_mode = st.checkbox("EUR → USD", value=False)
    eur_rate = st.number_input("Rate", value=1.16, step=0.01, format="%.2f") if eur_mode else 1.0
with c3:
    do_package = st.checkbox("Package Price", value=True)

if offer_file and st.button("🚀 Run Price Analysis", type="primary", use_container_width=True):

    tool_path = CLAUDE_DIR / "master_price_tool.py"
    if not tool_path.exists():
        alt = OUTPUT_DIR / "master_price_tool.py"
        if alt.exists(): shutil.copy(alt, tool_path)
        else:
            st.error("❌ master_price_tool.py not found — upload it in the sidebar.")
            st.stop()

    if total_files == 0:
        st.error("❌ No wholesaler files found — upload them in the sidebar first.")
        st.stop()

    # Save offer file
    offer_tmp = CLAUDE_DIR / f"tmp_offer_{offer_file.name}"
    offer_tmp.write_bytes(offer_file.getvalue())

    # EUR conversion
    if eur_mode and eur_rate != 1.0:
        try:
            df_raw = pd.read_excel(str(offer_tmp), header=None)
            hrow = 0
            for i, row in df_raw.iterrows():
                vals = [str(v).lower() for v in row.values if pd.notna(v)]
                if any(k in ' '.join(vals) for k in ['price','upc','ean']): hrow=i; break
            df = pd.read_excel(str(offer_tmp), header=hrow)
            df.columns = [str(c).strip() for c in df.columns]
            for c in df.columns:
                if any(k in str(c).lower() for k in ['price','cost','prix']):
                    df[c] = pd.to_numeric(df[c], errors='coerce') * eur_rate
                    break
            df.to_excel(str(offer_tmp), index=False)
            st.info(f"✅ EUR × {eur_rate} = USD applied")
        except Exception as e:
            st.warning(f"⚠️ EUR conversion skipped: {e}")

    output_path = OUTPUT_DIR / f"PriceAnalysis_{offer_file.name}"
    pkg_path    = OUTPUT_DIR / f"PriceAnalysis_{offer_file.name.replace('.xlsx','')}_PackagePrice.xlsx"

    with st.spinner("🔍 Running price analysis across 5 wholesalers..."):
        try:
            if str(CLAUDE_DIR) not in sys.path:
                sys.path.insert(0, str(CLAUDE_DIR))
            # Remove cached module to reload fresh
            if "master_price_tool" in sys.modules:
                del sys.modules["master_price_tool"]
            spec = importlib.util.spec_from_file_location("master_price_tool", str(tool_path))
            tool = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(tool)
            result = tool.run(str(offer_tmp), str(output_path), package_price=do_package)
            success = True
        except Exception as e:
            st.error(f"❌ Error: {e}")
            st.code(traceback.format_exc())
            success = False

    if success:
        st.markdown("---")
        st.subheader("✅ Analysis Complete!")

        # Summary stats
        try:
            wb_r = pd.read_excel(str(output_path), sheet_name=None)
            sheets = list(wb_r.keys())
            matched_sheet   = [s for s in sheets if 'match' in s.lower() and 'not' not in s.lower()]
            notfound_sheet  = [s for s in sheets if 'not' in s.lower() and 'found' in s.lower()]
            n_matched  = len(wb_r[matched_sheet[0]])   - 4 if matched_sheet  else '?'
            n_notfound = len(wb_r[notfound_sheet[0]])  - 4 if notfound_sheet else '?'

            cc1, cc2, cc3 = st.columns(3)
            with cc1: st.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#1B5E20">{n_matched}</div><div class="stat-label">Matched</div></div>', unsafe_allow_html=True)
            with cc2: st.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#C00000">{n_notfound}</div><div class="stat-label">Not Found</div></div>', unsafe_allow_html=True)
            with cc3:
                if pkg_path.exists():
                    try:
                        df_b = pd.read_excel(str(pkg_path), sheet_name='Package Pricing', header=None)
                        banner = str(df_b.iloc[3, 0]) if len(df_b) > 3 else ''
                        if '✅' in banner and 'Discount:' in banner:
                            disc = banner.split('Discount:')[1].split('%')[0].strip()
                            st.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#1B5E20">{disc}%</div><div class="stat-label">Package Discount</div></div>', unsafe_allow_html=True)
                        elif '⚠️' in banner:
                            disc = banner.split('discount:')[1].split('%')[0].strip() if 'discount:' in banner.lower() else '?'
                            st.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#C00000">{disc}%</div><div class="stat-label">Package Discount</div></div>', unsafe_allow_html=True)
                    except: pass
        except: pass

        # Verdict
        if pkg_path.exists():
            try:
                df_b = pd.read_excel(str(pkg_path), sheet_name='Package Pricing', header=None)
                banner = str(df_b.iloc[3, 0]) if len(df_b) > 3 else ''
                st.markdown("")
                if '✅' in banner:
                    disc = banner.split('Discount:')[1].split('%')[0].strip() + '%' if 'Discount:' in banner else ''
                    if float(disc.replace('%','').replace('+','').replace('−','-').replace('−','-')) <= -40:
                        st.markdown(f'<div class="verdict-sharp">🟢 <strong>SHARP — Excellent deal!</strong> {disc} below market. Strong buy.</div>', unsafe_allow_html=True)
                    elif float(disc.replace('%','').replace('+','').replace('−','-').replace('−','-')) <= -35:
                        st.markdown(f'<div class="verdict-good">🔵 <strong>GOOD — Solid deal.</strong> {disc} below market. Worth taking.</div>', unsafe_allow_html=True)
                    else:
                        st.markdown(f'<div class="verdict-min">🟠 <strong>MINIMUM — Viable.</strong> {disc} below market. Proceed with caution.</div>', unsafe_allow_html=True)
                elif '⚠️' in banner:
                    st.markdown(f'<div class="verdict-bad">⚠️ <strong>BELOW TARGET</strong> — Package does not meet −30% minimum. Counter or negotiate.</div>', unsafe_allow_html=True)
            except: pass

        # Downloads
        st.markdown("### 📥 Download Reports")
        dl1, dl2 = st.columns(2)
        with dl1:
            if output_path.exists():
                st.download_button(
                    "📋 Price Analysis (.xlsx)",
                    data=output_path.read_bytes(),
                    file_name=f"PriceAnalysis_{offer_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        with dl2:
            if do_package and pkg_path.exists():
                st.download_button(
                    "📦 Package Price (.xlsx)",
                    data=pkg_path.read_bytes(),
                    file_name=f"PackagePrice_{offer_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    if offer_tmp.exists():
        offer_tmp.unlink()

st.markdown("---")
st.caption("🟢 −40%+ Sharp  ·  🔵 −35% Good  ·  🟠 −30% Minimum  ·  ⚠️ Below target  ·  🔴 Above market")
