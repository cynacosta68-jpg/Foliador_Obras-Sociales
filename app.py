import streamlit as st
import fitz
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import numpy as np
import re
import io
import math
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
try:
    from scipy import ndimage
    HAS_SCIPY = True
except ImportError:
    HAS_SCIPY = False
import tempfile, os, base64, time

st.set_page_config(page_title="Foliador de Obras Sociales", page_icon="📋", layout="centered", initial_sidebar_state="collapsed")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;1,9..40,300&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif!important}
.stApp{background:linear-gradient(160deg,#060e1a 0%,#0a1628 40%,#0f1f3d 100%)}
.main-header{text-align:center;padding:2.5rem 0 1rem}
.main-header h1{font-family:'DM Sans',sans-serif!important;font-weight:600;font-size:2rem;color:#e2e8f0;letter-spacing:-.5px;margin-bottom:.25rem}
.main-header p{font-family:'DM Sans',sans-serif!important;font-weight:300;font-size:.95rem;color:#64748b;letter-spacing:.3px}
.accent-line{width:48px;height:2px;background:#3b82f6;margin:.75rem auto;border-radius:2px}
.card{background:rgba(17,29,51,.6);border:1px solid rgba(59,130,246,.08);border-radius:12px;padding:1.75rem;margin-bottom:1.25rem;backdrop-filter:blur(8px)}
.card-title{font-family:'DM Sans',sans-serif!important;font-weight:500;font-size:.8rem;color:#60a5fa;text-transform:uppercase;letter-spacing:1.5px;margin-bottom:1rem}
[data-testid="stFileUploader"]{background:rgba(15,23,42,.5);border:1px dashed rgba(59,130,246,.2);border-radius:10px;padding:1rem}
[data-testid="stFileUploader"]:hover{border-color:rgba(59,130,246,.4)}
.stButton>button{font-family:'DM Sans',sans-serif!important;font-weight:500;background:linear-gradient(135deg,#2563eb,#3b82f6)!important;color:#fff!important;border:none!important;border-radius:8px!important;padding:.6rem 2rem!important;letter-spacing:.3px;transition:all .2s}
.stButton>button:hover{background:linear-gradient(135deg,#1d4ed8,#2563eb)!important;box-shadow:0 4px 20px rgba(59,130,246,.3)!important}
.stDownloadButton>button{font-family:'DM Sans',sans-serif!important;font-weight:500;background:linear-gradient(135deg,#059669,#10b981)!important;color:#fff!important;border:none!important;border-radius:8px!important;padding:.6rem 2rem!important;width:100%}
.stDownloadButton>button:hover{background:linear-gradient(135deg,#047857,#059669)!important;box-shadow:0 4px 20px rgba(16,185,129,.3)!important}
.stat-row{display:flex;gap:1rem;margin:1rem 0;flex-wrap:wrap}
.stat-box{flex:1;min-width:100px;background:rgba(15,23,42,.5);border:1px solid rgba(59,130,246,.1);border-radius:10px;padding:1.1rem;text-align:center}
.stat-number{font-family:'DM Sans',sans-serif!important;font-size:1.8rem;font-weight:600;color:#60a5fa;line-height:1}
.stat-label{font-family:'DM Sans',sans-serif!important;font-size:.7rem;color:#64748b;text-transform:uppercase;letter-spacing:1px;margin-top:.4rem}
.code-card{background:rgba(15,23,42,.5);border:1px solid rgba(59,130,246,.12);border-radius:8px;padding:.6rem;text-align:center;margin-bottom:.5rem}
.code-card img{border-radius:4px;width:100%;background:#fff;padding:4px}
.page-num{font-family:'DM Sans',sans-serif;font-size:.7rem;color:#94a3b8;margin-top:.3rem;letter-spacing:.5px}
.ocr-tag{display:inline-block;background:rgba(59,130,246,.15);color:#60a5fa;font-family:'DM Sans',sans-serif;font-size:.75rem;font-weight:500;padding:.15rem .5rem;border-radius:4px;margin-top:.25rem}
.no-tag{display:inline-block;background:rgba(251,191,36,.1);color:#fbbf24;font-family:'DM Sans',sans-serif;font-size:.7rem;padding:.15rem .5rem;border-radius:4px;margin-top:.25rem}
.match-ok{display:inline-block;background:rgba(16,185,129,.12);color:#34d399;font-family:'DM Sans',sans-serif;font-size:.65rem;font-weight:500;padding:.12rem .4rem;border-radius:3px;margin-top:.2rem}
.match-fail{display:inline-block;background:rgba(239,68,68,.12);color:#f87171;font-family:'DM Sans',sans-serif;font-size:.65rem;font-weight:500;padding:.12rem .4rem;border-radius:3px;margin-top:.2rem}
[data-testid="stTextInput"] input{font-family:'DM Sans',sans-serif!important;background:rgba(15,23,42,.6)!important;border:1px solid rgba(59,130,246,.15)!important;border-radius:8px!important;color:#e2e8f0!important}
.stAlert{border-radius:10px!important}
hr{border-color:rgba(59,130,246,.08)!important}
#MainMenu{visibility:hidden}footer{visibility:hidden}
.info-text{font-family:'DM Sans',sans-serif;font-size:.82rem;color:#94a3b8;margin-bottom:1rem;line-height:1.5}
.step-badge{display:inline-block;background:rgba(59,130,246,.12);color:#60a5fa;font-family:'DM Sans',sans-serif;font-size:.65rem;font-weight:600;padding:.2rem .6rem;border-radius:20px;letter-spacing:1px;text-transform:uppercase;margin-bottom:.5rem}
</style>
""", unsafe_allow_html=True)

CODE_RE = re.compile(r'[CcĆć({[\|0Oo]\s*(\d{1,5})')

# ═══════════════════════════════════════════════════════════════
# BASE DE USUARIOS
# ═══════════════════════════════════════════════════════════════

@st.cache_data
def load_base(file_bytes):
    """Load user database and build lookup dict keyed by Matricula."""
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = df.columns.str.strip()
    lookup = {}
    for _, row in df.iterrows():
        mat = str(row.get('Matricula', '')).strip()
        if mat:
            lookup[mat] = {
                'Nombre': str(row.get('Nombre', '')).strip(),
                'CUIT': row.get('CUIT', ''),
                'Responsabilidad Fiscal': str(row.get('Responsabilidad Fiscal', '')).strip(),
                'Especialidad': str(row.get('Especialidad', '')).strip(),
                'Arancel': str(row.get('Arancel', '')).strip(),
            }
    return lookup, len(df)


def lookup_profesional(code, base_lookup):
    """Look up a code in the base. Returns dict with user info or None."""
    if not base_lookup or not code:
        return None
    code_clean = code.strip()
    if code_clean in base_lookup:
        return base_lookup[code_clean]
    # Try without leading zeros, etc
    for k, v in base_lookup.items():
        if k.replace(' ', '') == code_clean.replace(' ', ''):
            return v
    return None


# ═══════════════════════════════════════════════════════════════
# PDF SCAN ENGINE
# ═══════════════════════════════════════════════════════════════

def phase1_density(pdf_path, total, progress_cb=None):
    doc = fitz.open(pdf_path)
    densities = []
    for i in range(total):
        page = doc[i]
        pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        crop = img.crop((0, 0, int(img.width * 0.25), int(img.height * 0.06)))
        arr = np.array(crop.convert('L'))
        densities.append(float(np.mean(arr < 100)))
        if progress_cb and i % 100 == 0:
            progress_cb(0.25 * i / total, f"Fase 1 · Densidad · Pág {i+1:,}/{total:,}")
    doc.close()
    median = float(np.median(densities)) if densities else 0
    threshold = max(median * 3, 0.02)
    return [i for i, d in enumerate(densities) if d > threshold]


def isolate_code_image(gray_arr):
    binary = (gray_arr < 140).astype(np.uint8) * 255
    bin_img = Image.fromarray(binary)
    e = bin_img
    for _ in range(2):
        e = e.filter(ImageFilter.MinFilter(3))
    d = e
    for _ in range(5):
        d = d.filter(ImageFilter.MaxFilter(3))
    mask = np.array(d) > 128
    cleaned = np.where(mask, gray_arr, 255).astype(np.uint8)
    if HAS_SCIPY:
        bw = (cleaned < 180).astype(np.uint8)
        labeled, n = ndimage.label(bw)
        if n > 1:
            sizes = ndimage.sum(bw, labeled, range(1, n + 1))
            keep = [i + 1 for i, s in enumerate(sizes) if s > max(sizes) * 0.08]
            cleaned = np.where(np.isin(labeled, keep), cleaned, 255).astype(np.uint8)
    dark = cleaned < 180
    if not dark.any():
        return None
    rows, cols = np.any(dark, axis=1), np.any(dark, axis=0)
    r0, r1 = np.where(rows)[0][[0, -1]]
    c0, c1 = np.where(cols)[0][[0, -1]]
    p = 15
    cropped = cleaned[max(0, r0-p):r1+p, max(0, c0-p):c1+p]
    b = 30
    padded = np.full((cropped.shape[0]+2*b, cropped.shape[1]+2*b), 255, dtype=np.uint8)
    padded[b:b+cropped.shape[0], b:b+cropped.shape[1]] = cropped
    return padded


def _arr_to_b64(arr):
    if arr is None: return None
    buf = io.BytesIO()
    Image.fromarray(arr).save(buf, format='JPEG', quality=70)
    return base64.b64encode(buf.getvalue()).decode()


def _process_candidate(pdf_path, pg):
    doc = fitz.open(pdf_path)
    page = doc[pg]
    mat = fitz.Matrix(300/72, 300/72)
    pix = page.get_pixmap(matrix=mat)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    # Text extraction first
    h_lim, w_lim = page.rect.height * 0.10, page.rect.width * 0.30
    for blk in page.get_text("blocks"):
        if blk[1] < h_lim and blk[0] < w_lim:
            txt = blk[4] if len(blk) > 4 else ""
            m = CODE_RE.findall(txt)
            if m:
                doc.close()
                crop = img.crop((0, 0, int(img.width*0.28), int(img.height*0.065)))
                iso = isolate_code_image(np.array(crop.convert('L')))
                return pg, f"C{m[0]}", "text", _arr_to_b64(iso)
    doc.close()
    crop = img.crop((0, 0, int(img.width*0.28), int(img.height*0.065)))
    gray = np.array(crop.convert('L'))
    iso = isolate_code_image(gray)
    img_b64 = _arr_to_b64(iso)
    if iso is None:
        return pg, None, None, img_b64
    final = np.where(iso < 140, 0, 255).astype(np.uint8)
    pil = Image.fromarray(final)
    big = pil.resize((pil.width*4, pil.height*4), Image.LANCZOS)
    results = []
    for psm in [6, 7, 8, 13]:
        for extra in ['', ' -c tessedit_char_whitelist=Cc0123456789']:
            try:
                text = pytesseract.image_to_string(big, config=f'--psm {psm} --oem 3{extra}')
                for m in CODE_RE.findall(text):
                    results.append(f"C{m}")
            except: pass
    code = Counter(results).most_common(1)[0][0] if results else None
    return pg, code, "ocr" if code else None, img_b64


def full_scan(pdf_path, total, progress_cb=None):
    if progress_cb: progress_cb(0.0, "Fase 1 · Detección por densidad de píxeles...")
    candidates = phase1_density(pdf_path, total, progress_cb)
    if progress_cb: progress_cb(0.3, f"Fase 1 · {len(candidates)} candidatas de {total:,} páginas")
    results, images = {}, {}
    if not candidates:
        return results, images, candidates
    workers = min(4, os.cpu_count() or 2)
    done = 0
    if progress_cb: progress_cb(0.35, f"Fase 2 · Aislando códigos ({len(candidates)} págs, {workers} hilos)...")
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futures = {ex.submit(_process_candidate, pdf_path, pg): pg for pg in candidates}
        for f in as_completed(futures):
            pg, code, method, img_b64 = f.result()
            if code: results[pg] = (code, method)
            if img_b64: images[pg] = img_b64
            done += 1
            if progress_cb and done % 5 == 0:
                progress_cb(0.35 + 0.6*done/len(candidates), f"Fase 2 · {done}/{len(candidates)}")
    if progress_cb: progress_cb(1.0, "Completado")
    return results, images, candidates


# ═══════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════

def build_ranges(code_pages, total):
    entries = sorted(code_pages.items())
    return [(code, pg+1, entries[i+1][0] if i+1 < len(entries) else total)
            for i, (pg, (code, _)) in enumerate(entries)]


def create_excel(rows, include_extra=True):
    wb = Workbook()
    ws = wb.active
    ws.title = "Foliado"
    hf = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
    hfill = PatternFill('solid', fgColor='1a3a6b')
    ha = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cf = Font(name='Calibri', size=11)
    ac = Alignment(horizontal='center', vertical='center')
    al = Alignment(horizontal='left', vertical='center')
    bdr = Border(left=Side(style='thin',color='2c5282'), right=Side(style='thin',color='2c5282'),
                 top=Side(style='thin',color='2c5282'), bottom=Side(style='thin',color='2c5282'))
    alt = PatternFill('solid', fgColor='edf2f7')

    if include_extra:
        headers = ['Cuenta','Página desde','Página hasta','Profesional','CUIT','Especialidad','Resp. Fiscal','Arancel']
        widths  = [12, 14, 14, 36, 18, 28, 22, 10]
        aligns  = [ac, ac, ac, al, ac, al, al, ac]
    else:
        headers = ['Cuenta','Página desde','Página hasta','Profesional']
        widths  = [16, 16, 16, 42]
        aligns  = [ac, ac, ac, al]

    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font, c.fill, c.alignment, c.border = hf, hfill, ha, bdr
        ws.column_dimensions[get_column_letter(ci)].width = w

    for ri, rd in enumerate(rows, 2):
        vals = (list(rd) + [''] * len(headers))[:len(headers)]
        for ci, (v, a) in enumerate(zip(vals, aligns), 1):
            c = ws.cell(row=ri, column=ci, value=v)
            c.font, c.alignment, c.border = cf, a, bdr
            if ri % 2 == 0: c.fill = alt

    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows)+1}"
    ws.freeze_panes = "A2"
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf


# ═══════════════════════════════════════════════════════════════
# UI
# ═══════════════════════════════════════════════════════════════

st.markdown("""
<div class="main-header">
    <h1>📋 Foliador de Obras Sociales</h1>
    <div class="accent-line"></div>
    <p>Extrae códigos de cuenta, cruza con la base de profesionales y genera el Excel</p>
</div>
""", unsafe_allow_html=True)

# ── PASO 1: Cargar archivos ──
st.markdown('<div class="card"><span class="step-badge">Paso 1</span><div class="card-title">Cargar documentos</div>', unsafe_allow_html=True)

col_pdf, col_base = st.columns(2)
with col_pdf:
    st.markdown('<p class="info-text">📄 <b>PDF escaneado</b></p>', unsafe_allow_html=True)
    uploaded_pdf = st.file_uploader("PDF", type=["pdf"], label_visibility="collapsed", key="pdf")
with col_base:
    st.markdown('<p class="info-text">📊 <b>Base de usuarios</b> (.xlsx)</p>', unsafe_allow_html=True)
    uploaded_base = st.file_uploader("Base", type=["xlsx", "xls"], label_visibility="collapsed", key="base")

st.markdown('</div>', unsafe_allow_html=True)

# Load base
base_lookup = {}
base_count = 0
if uploaded_base:
    base_lookup, base_count = load_base(uploaded_base.read())
    uploaded_base.seek(0)

if uploaded_pdf:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_pdf.read()); tmp_path = tmp.name
    doc = fitz.open(tmp_path); total = len(doc); doc.close()

    stats_html = f'<div class="stat-row"><div class="stat-box"><div class="stat-number">{total:,}</div><div class="stat-label">Páginas PDF</div></div><div class="stat-box"><div class="stat-number">{uploaded_pdf.size/1024/1024:.1f} MB</div><div class="stat-label">Tamaño</div></div>'
    if base_lookup:
        stats_html += f'<div class="stat-box"><div class="stat-number">{base_count:,}</div><div class="stat-label">Profesionales en base</div></div>'
    stats_html += '</div>'
    st.markdown(stats_html, unsafe_allow_html=True)

    # ── PASO 2: Escaneo ──
    st.markdown('<div class="card"><span class="step-badge">Paso 2</span><div class="card-title">Escaneo inteligente</div>', unsafe_allow_html=True)
    st.markdown('<p class="info-text"><b>Fase 1</b> — Detecta tinta en el margen (instantáneo) · <b>Fase 2</b> — Aísla el código y lo lee por OCR · <b>Cruce</b> — Busca cada código en la base de usuarios</p>', unsafe_allow_html=True)

    if st.button("🔍  Escanear y cruzar", use_container_width=True):
        bar = st.progress(0)
        t0 = time.time()
        results, images, candidates = full_scan(tmp_path, total, lambda p, m: bar.progress(min(p, 1.0), text=m))
        elapsed = time.time() - t0
        bar.empty()

        # Do the lookup
        matches_count = 0
        matched_info = {}
        for pg, (code, method) in results.items():
            info = lookup_profesional(code, base_lookup)
            if info:
                matched_info[pg] = info
                matches_count += 1

        st.session_state.update({
            'results': results, 'images': images, 'candidates': candidates,
            'matched_info': matched_info, 'scanned': True, 'pdf_path': tmp_path,
            'elapsed': elapsed
        })

        st.markdown(f'<div class="stat-row"><div class="stat-box"><div class="stat-number">{len(candidates)}</div><div class="stat-label">Candidatas</div></div><div class="stat-box"><div class="stat-number">{len(results)}</div><div class="stat-label">Códigos leídos</div></div><div class="stat-box"><div class="stat-number">{matches_count}/{len(results)}</div><div class="stat-label">Cruzados en base</div></div><div class="stat-box"><div class="stat-number">{elapsed:.1f}s</div><div class="stat-label">Tiempo</div></div></div>', unsafe_allow_html=True)

        not_matched = len(results) - matches_count
        if matches_count > 0 and not_matched == 0:
            st.success(f"**{matches_count}** códigos detectados y **todos cruzados** con la base de profesionales.")
        elif matches_count > 0:
            st.success(f"**{matches_count}** cruzados con la base. **{not_matched}** sin match (verificá el código en la tabla).")
        elif results:
            st.warning("Códigos detectados pero ninguno cruzó con la base. ¿Cargaste la base correcta?")
        else:
            st.warning("No se detectaron códigos. Ingresalos manualmente.")

    # Thumbnails
    if st.session_state.get('scanned'):
        results = st.session_state['results']
        images = st.session_state['images']
        candidates = st.session_state['candidates']
        matched_info = st.session_state.get('matched_info', {})

        if candidates:
            st.markdown("**Código aislado del margen** + cruce con base:")
            per_page = 12
            total_groups = math.ceil(len(candidates) / per_page)
            group = st.number_input("Grupo", 1, total_groups, 1) if total_groups > 1 else 1
            vis = candidates[(group-1)*per_page : group*per_page]

            for rs in range(0, len(vis), 4):
                cols = st.columns(4)
                for j, col in enumerate(cols):
                    if rs+j >= len(vis): break
                    pg = vis[rs+j]
                    with col:
                        b64 = images.get(pg)
                        # OCR tag
                        if pg in results:
                            code = results[pg][0]
                            tag = f'<span class="ocr-tag">{code}</span>'
                        else:
                            tag = '<span class="no-tag">no leído</span>'
                        # Match tag
                        if pg in matched_info:
                            nombre = matched_info[pg]['Nombre']
                            nombre_short = nombre[:25] + '…' if len(nombre) > 25 else nombre
                            match_tag = f'<span class="match-ok">✓ {nombre_short}</span>'
                        elif pg in results:
                            match_tag = '<span class="match-fail">✗ sin match</span>'
                        else:
                            match_tag = ''

                        if b64:
                            st.markdown(f'<div class="code-card"><img src="data:image/jpeg;base64,{b64}"/><div class="page-num">Pág. {pg+1}</div>{tag}<br>{match_tag}</div>', unsafe_allow_html=True)
                        else:
                            st.markdown(f'<div class="code-card"><div class="page-num">Pág. {pg+1}</div>{tag}<br>{match_tag}</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # ── PASO 3: Editar ──
    if st.session_state.get('scanned'):
        results = st.session_state['results']
        matched_info = st.session_state.get('matched_info', {})

        st.markdown('<div class="card"><span class="step-badge">Paso 3</span><div class="card-title">Confirmar y editar</div>', unsafe_allow_html=True)
        st.markdown('<p class="info-text">Los datos del profesional se completan automáticamente desde la base. Corregí el código si el OCR no acertó — al descargar se vuelve a cruzar.</p>', unsafe_allow_html=True)

        has_base = bool(base_lookup)

        if results:
            ranges = build_ranges(results, total)
            init = []
            for code, s, e in ranges:
                info = lookup_profesional(code, base_lookup)
                row = {
                    "Cuenta": code,
                    "Página desde": s,
                    "Página hasta": e,
                    "Profesional": info['Nombre'] if info else "",
                }
                if has_base:
                    row["CUIT"] = str(int(info['CUIT'])) if info and info.get('CUIT') and pd.notna(info['CUIT']) else ""
                    row["Especialidad"] = info['Especialidad'] if info else ""
                    row["Resp. Fiscal"] = info['Responsabilidad Fiscal'] if info else ""
                    row["Arancel"] = info['Arancel'] if info else ""
                init.append(row)
        else:
            init = [{"Cuenta": "", "Página desde": 1, "Página hasta": total, "Profesional": ""}]
            if has_base:
                init[0].update({"CUIT": "", "Especialidad": "", "Resp. Fiscal": "", "Arancel": ""})

        df = pd.DataFrame(init)

        col_config = {
            "Cuenta": st.column_config.TextColumn("Cuenta", width="small"),
            "Página desde": st.column_config.NumberColumn("Pág. desde", width="small", min_value=1, max_value=total),
            "Página hasta": st.column_config.NumberColumn("Pág. hasta", width="small", min_value=1, max_value=total),
            "Profesional": st.column_config.TextColumn("Profesional", width="medium"),
        }
        if has_base:
            col_config.update({
                "CUIT": st.column_config.TextColumn("CUIT", width="small"),
                "Especialidad": st.column_config.TextColumn("Especialidad", width="medium"),
                "Resp. Fiscal": st.column_config.TextColumn("Resp. Fiscal", width="small"),
                "Arancel": st.column_config.TextColumn("Arancel", width="small"),
            })

        edited = st.data_editor(df, use_container_width=True, num_rows="dynamic", hide_index=True, column_config=col_config)
        st.session_state['edited_df'] = edited
        st.markdown('</div>', unsafe_allow_html=True)

        # ── PASO 4: Descargar ──
        st.markdown('<div class="card"><span class="step-badge">Paso 4</span><div class="card-title">Descargar Excel</div>', unsafe_allow_html=True)

        # Re-lookup on download (in case user corrected codes)
        final_df = st.session_state['edited_df'].copy()
        if has_base:
            for idx, row in final_df.iterrows():
                code = str(row.get('Cuenta', '')).strip()
                if code and (not row.get('Profesional') or str(row.get('Profesional', '')).strip() == ''):
                    info = lookup_profesional(code, base_lookup)
                    if info:
                        final_df.at[idx, 'Profesional'] = info['Nombre']
                        final_df.at[idx, 'CUIT'] = str(int(info['CUIT'])) if pd.notna(info.get('CUIT')) else ""
                        final_df.at[idx, 'Especialidad'] = info['Especialidad']
                        final_df.at[idx, 'Resp. Fiscal'] = info['Responsabilidad Fiscal']
                        final_df.at[idx, 'Arancel'] = info['Arancel']

        excel = create_excel(final_df.values.tolist(), include_extra=has_base)
        fname = uploaded_pdf.name.replace(".pdf","").replace(".PDF","")
        st.download_button("⬇  Descargar Excel", excel, f"Foliado_{fname}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div style="text-align:center;padding:3rem 0 1.5rem;border-top:1px solid rgba(59,130,246,.06);margin-top:2rem"><p style="font-size:.72rem;color:#475569;letter-spacing:.5px;font-family:\'DM Sans\',sans-serif">FOLIADOR DE OBRAS SOCIALES · v3.0 · Detección + aislamiento + cruce con base de profesionales</p></div>', unsafe_allow_html=True)
