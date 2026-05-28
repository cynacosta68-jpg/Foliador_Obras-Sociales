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
.stat-row{display:flex;gap:1rem;margin:1rem 0}
.stat-box{flex:1;background:rgba(15,23,42,.5);border:1px solid rgba(59,130,246,.1);border-radius:10px;padding:1.1rem;text-align:center}
.stat-number{font-family:'DM Sans',sans-serif!important;font-size:1.8rem;font-weight:600;color:#60a5fa;line-height:1}
.stat-label{font-family:'DM Sans',sans-serif!important;font-size:.7rem;color:#64748b;text-transform:uppercase;letter-spacing:1px;margin-top:.4rem}
.code-card{background:rgba(15,23,42,.5);border:1px solid rgba(59,130,246,.12);border-radius:8px;padding:.6rem;text-align:center;margin-bottom:.5rem}
.code-card img{border-radius:4px;width:100%;background:#fff;padding:4px}
.page-num{font-family:'DM Sans',sans-serif;font-size:.7rem;color:#94a3b8;margin-top:.3rem;letter-spacing:.5px}
.ocr-tag{display:inline-block;background:rgba(59,130,246,.15);color:#60a5fa;font-family:'DM Sans',sans-serif;font-size:.75rem;font-weight:500;padding:.15rem .5rem;border-radius:4px;margin-top:.25rem}
.no-tag{display:inline-block;background:rgba(251,191,36,.1);color:#fbbf24;font-family:'DM Sans',sans-serif;font-size:.7rem;padding:.15rem .5rem;border-radius:4px;margin-top:.25rem}
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
# ENGINE
# ═══════════════════════════════════════════════════════════════

def phase1_density(pdf_path, total, progress_cb=None):
    """Pixel density scan at 72 DPI — identifies pages with ink in top-left margin."""
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
    """Morphological filtering: remove body text, keep only large handwritten code.
    Returns cleaned grayscale numpy array or None."""
    binary = (gray_arr < 140).astype(np.uint8) * 255
    bin_img = Image.fromarray(binary)
    # Light erosion kills small text
    e = bin_img
    for _ in range(2):
        e = e.filter(ImageFilter.MinFilter(3))
    # Dilate restores large strokes
    d = e
    for _ in range(5):
        d = d.filter(ImageFilter.MaxFilter(3))
    mask = np.array(d) > 128
    cleaned = np.where(mask, gray_arr, 255).astype(np.uint8)

    # Remove isolated blobs via connected components
    if HAS_SCIPY:
        bw = (cleaned < 180).astype(np.uint8)
        labeled, n = ndimage.label(bw)
        if n > 1:
            sizes = ndimage.sum(bw, labeled, range(1, n + 1))
            keep = [i + 1 for i, s in enumerate(sizes) if s > max(sizes) * 0.08]
            cleaned = np.where(np.isin(labeled, keep), cleaned, 255).astype(np.uint8)

    # Auto-crop
    dark = cleaned < 180
    if not dark.any():
        return None
    rows, cols = np.any(dark, axis=1), np.any(dark, axis=0)
    r0, r1 = np.where(rows)[0][[0, -1]]
    c0, c1 = np.where(cols)[0][[0, -1]]
    p = 15
    cropped = cleaned[max(0, r0 - p):r1 + p, max(0, c0 - p):c1 + p]

    # White border
    b = 30
    padded = np.full((cropped.shape[0] + 2 * b, cropped.shape[1] + 2 * b), 255, dtype=np.uint8)
    padded[b:b + cropped.shape[0], b:b + cropped.shape[1]] = cropped
    return padded


def _process_candidate(pdf_path, pg):
    """Process one candidate page: isolate code image + OCR."""
    doc = fitz.open(pdf_path)
    page = doc[pg]
    mat = fitz.Matrix(300 / 72, 300 / 72)
    pix = page.get_pixmap(matrix=mat)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    # Also try text extraction first
    h_lim, w_lim = page.rect.height * 0.10, page.rect.width * 0.30
    for blk in page.get_text("blocks"):
        if blk[1] < h_lim and blk[0] < w_lim:
            txt = blk[4] if len(blk) > 4 else ""
            m = CODE_RE.findall(txt)
            if m:
                doc.close()
                # Still generate the code image for preview
                crop = img.crop((0, 0, int(img.width * 0.28), int(img.height * 0.065)))
                gray = np.array(crop.convert('L'))
                iso = isolate_code_image(gray)
                img_b64 = _arr_to_b64(iso) if iso is not None else None
                return pg, f"C{m[0]}", "text", img_b64

    doc.close()

    # Crop margin only: top 6.5%, left 28%
    crop = img.crop((0, 0, int(img.width * 0.28), int(img.height * 0.065)))
    gray = np.array(crop.convert('L'))
    iso = isolate_code_image(gray)
    img_b64 = _arr_to_b64(iso) if iso is not None else None

    if iso is None:
        return pg, None, None, img_b64

    # OCR on isolated image scaled 4x
    final = np.where(iso < 140, 0, 255).astype(np.uint8)
    pil = Image.fromarray(final)
    big = pil.resize((pil.width * 4, pil.height * 4), Image.LANCZOS)

    results = []
    for psm in [6, 7, 8, 13]:
        for extra in ['', ' -c tessedit_char_whitelist=Cc0123456789']:
            try:
                text = pytesseract.image_to_string(big, config=f'--psm {psm} --oem 3{extra}')
                for m in CODE_RE.findall(text):
                    results.append(f"C{m}")
            except Exception:
                pass

    code = Counter(results).most_common(1)[0][0] if results else None
    return pg, code, "ocr" if code else None, img_b64


def _arr_to_b64(arr):
    """Convert numpy array to base64 JPEG."""
    if arr is None:
        return None
    img = Image.fromarray(arr)
    buf = io.BytesIO()
    img.save(buf, format='JPEG', quality=70)
    return base64.b64encode(buf.getvalue()).decode()


def full_scan(pdf_path, total, progress_cb=None):
    """Full pipeline: density scan → isolate + OCR candidates."""
    # Phase 1
    if progress_cb:
        progress_cb(0.0, "Fase 1 · Detección por densidad de píxeles...")
    candidates = phase1_density(pdf_path, total, progress_cb)
    if progress_cb:
        progress_cb(0.3, f"Fase 1 · {len(candidates)} candidatas de {total:,} páginas")

    # Phase 2: parallel isolation + OCR on candidates only
    results = {}
    images = {}
    if not candidates:
        return results, images, candidates

    workers = min(4, os.cpu_count() or 2)
    done = 0
    if progress_cb:
        progress_cb(0.35, f"Fase 2 · Aislando códigos de {len(candidates)} páginas ({workers} hilos)...")

    with ThreadPoolExecutor(max_workers=workers) as ex:
        futures = {ex.submit(_process_candidate, pdf_path, pg): pg for pg in candidates}
        for f in as_completed(futures):
            pg, code, method, img_b64 = f.result()
            if code:
                results[pg] = (code, method)
            if img_b64:
                images[pg] = img_b64
            done += 1
            if progress_cb and done % 5 == 0:
                progress_cb(0.35 + 0.6 * done / len(candidates),
                            f"Fase 2 · {done}/{len(candidates)} procesadas")

    if progress_cb:
        progress_cb(1.0, "Completado")
    return results, images, candidates


# ═══════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════

def extract_professional(doc, start_page):
    page = doc[start_page]
    text = page.get_text("text")
    m = re.search(r'por\s+([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+){1,5})', text)
    if m: return m.group(1).strip()
    m = re.search(r'\(\*\)\s*([A-ZÁÉÍÓÚÑ][^,\n]+)', text)
    if m: return m.group(1).strip()
    return ""

def build_ranges(code_pages, total):
    entries = sorted(code_pages.items())
    return [(code, pg+1, entries[i+1][0] if i+1 < len(entries) else total)
            for i, (pg, (code, _)) in enumerate(entries)]

def create_excel(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Foliado"
    hf = Font(name='Calibri', bold=True, size=11, color='FFFFFF')
    hfill = PatternFill('solid', fgColor='1a3a6b')
    ha = Alignment(horizontal='center', vertical='center')
    cf = Font(name='Calibri', size=11)
    ac, al = Alignment(horizontal='center', vertical='center'), Alignment(horizontal='left', vertical='center')
    bdr = Border(left=Side(style='thin',color='2c5282'), right=Side(style='thin',color='2c5282'),
                 top=Side(style='thin',color='2c5282'), bottom=Side(style='thin',color='2c5282'))
    alt = PatternFill('solid', fgColor='edf2f7')
    for ci,(h,w) in enumerate(zip(['Cuenta','Página desde','Página hasta','Profesional'],[16,16,16,42]),1):
        c=ws.cell(row=1,column=ci,value=h); c.font,c.fill,c.alignment,c.border=hf,hfill,ha,bdr
        ws.column_dimensions[get_column_letter(ci)].width=w
    for ri,rd in enumerate(rows,2):
        vals=(list(rd)+["","","",""])[:4]
        for ci,(v,a) in enumerate(zip(vals,[ac,ac,ac,al]),1):
            c=ws.cell(row=ri,column=ci,value=v); c.font,c.alignment,c.border=cf,a,bdr
            if ri%2==0: c.fill=alt
    ws.auto_filter.ref=f"A1:D{len(rows)+1}"; ws.freeze_panes="A2"
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# ═══════════════════════════════════════════════════════════════
# UI
# ═══════════════════════════════════════════════════════════════

st.markdown("""
<div class="main-header">
    <h1>📋 Foliador de Obras Sociales</h1>
    <div class="accent-line"></div>
    <p>Extrae códigos de cuenta y rangos de páginas desde PDFs escaneados</p>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="card"><span class="step-badge">Paso 1</span><div class="card-title">Cargar documento</div>', unsafe_allow_html=True)
uploaded = st.file_uploader("PDF", type=["pdf"], label_visibility="collapsed", help="Optimizado para +2000 páginas")
st.markdown('</div>', unsafe_allow_html=True)

if uploaded:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded.read()); tmp_path = tmp.name
    doc = fitz.open(tmp_path); total = len(doc); doc.close()

    st.markdown(f'<div class="stat-row"><div class="stat-box"><div class="stat-number">{total:,}</div><div class="stat-label">Páginas</div></div><div class="stat-box"><div class="stat-number">{uploaded.size/1024/1024:.1f} MB</div><div class="stat-label">Tamaño</div></div></div>', unsafe_allow_html=True)

    st.markdown('<div class="card"><span class="step-badge">Paso 2</span><div class="card-title">Escaneo inteligente</div>', unsafe_allow_html=True)
    st.markdown('<p class="info-text"><b>Fase 1</b> — Detecta tinta en el margen (instantáneo) · <b>Fase 2</b> — Aísla el código manuscrito y lo lee por OCR (solo candidatas, paralelo)</p>', unsafe_allow_html=True)

    if st.button("🔍  Escanear documento", use_container_width=True):
        bar = st.progress(0)
        t0 = time.time()
        results, images, candidates = full_scan(tmp_path, total, lambda p, m: bar.progress(min(p, 1.0), text=m))
        elapsed = time.time() - t0
        bar.empty()
        st.session_state.update({'results': results, 'images': images, 'candidates': candidates, 'scanned': True, 'pdf_path': tmp_path})

        st.markdown(f'<div class="stat-row"><div class="stat-box"><div class="stat-number">{len(candidates)}</div><div class="stat-label">Candidatas</div></div><div class="stat-box"><div class="stat-number">{len(results)}</div><div class="stat-label">Códigos leídos</div></div><div class="stat-box"><div class="stat-number">{elapsed:.1f}s</div><div class="stat-label">Tiempo</div></div><div class="stat-box"><div class="stat-number">{total/elapsed:.0f}/s</div><div class="stat-label">Pág/seg</div></div></div>', unsafe_allow_html=True)
        if results:
            st.success(f"**{len(results)}** códigos detectados en **{len(candidates)}** candidatas. Verificá mirando las imágenes aisladas abajo.")
        elif candidates:
            st.warning(f"Se encontraron **{len(candidates)}** páginas con tinta pero el OCR no pudo leer los códigos. Ingresalos manualmente mirando las imágenes.")
        else:
            st.warning("No se detectaron páginas con código. Usá la tabla para ingresarlos manualmente.")

    if st.session_state.get('scanned'):
        results = st.session_state['results']
        images = st.session_state['images']
        candidates = st.session_state['candidates']

        if candidates:
            st.markdown("**Código aislado del margen** (solo la tinta grande, sin texto del cuerpo):")
            per_page = 12
            total_groups = math.ceil(len(candidates) / per_page)
            group = st.number_input("Grupo", 1, total_groups, 1) if total_groups > 1 else 1
            vis = candidates[(group-1)*per_page : group*per_page]

            for rs in range(0, len(vis), 4):
                cols = st.columns(4)
                for j, col in enumerate(cols):
                    if rs + j >= len(vis): break
                    pg = vis[rs + j]
                    with col:
                        b64 = images.get(pg)
                        tag = f'<span class="ocr-tag">{results[pg][0]}</span>' if pg in results else '<span class="no-tag">no leído</span>'
                        if b64:
                            st.markdown(f'<div class="code-card"><img src="data:image/jpeg;base64,{b64}"/><div class="page-num">Pág. {pg+1}</div>{tag}</div>', unsafe_allow_html=True)
                        else:
                            st.markdown(f'<div class="code-card"><div class="page-num">Pág. {pg+1}</div>{tag}</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state.get('scanned'):
        results = st.session_state['results']
        st.markdown('<div class="card"><span class="step-badge">Paso 3</span><div class="card-title">Confirmar y editar</div>', unsafe_allow_html=True)
        st.markdown('<p class="info-text">Mirá las imágenes aisladas de arriba y corregí los códigos en la tabla. Cada código va hasta la página anterior al siguiente.</p>', unsafe_allow_html=True)

        if results:
            ranges = build_ranges(results, total)
            doc3 = fitz.open(st.session_state.get('pdf_path', tmp_path))
            init = [{"Cuenta": c, "Página desde": s, "Página hasta": e, "Profesional": extract_professional(doc3, s-1)} for c,s,e in ranges]
            doc3.close()
        else:
            init = [{"Cuenta": "", "Página desde": 1, "Página hasta": total, "Profesional": ""}]

        edited = st.data_editor(pd.DataFrame(init), use_container_width=True, num_rows="dynamic", hide_index=True,
            column_config={"Cuenta": st.column_config.TextColumn("Cuenta", width="small"),
                           "Página desde": st.column_config.NumberColumn("Pág. desde", width="small", min_value=1, max_value=total),
                           "Página hasta": st.column_config.NumberColumn("Pág. hasta", width="small", min_value=1, max_value=total),
                           "Profesional": st.column_config.TextColumn("Profesional", width="large")})
        st.session_state['edited_df'] = edited
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="card"><span class="step-badge">Paso 4</span><div class="card-title">Descargar Excel</div>', unsafe_allow_html=True)
        excel = create_excel(st.session_state['edited_df'].values.tolist())
        fname = uploaded.name.replace(".pdf","").replace(".PDF","")
        st.download_button("⬇  Descargar Excel", excel, f"Foliado_{fname}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div style="text-align:center;padding:3rem 0 1.5rem;border-top:1px solid rgba(59,130,246,.06);margin-top:2rem"><p style="font-size:.72rem;color:#475569;letter-spacing:.5px;font-family:\'DM Sans\',sans-serif">FOLIADOR DE OBRAS SOCIALES · v2.0 · Detección por densidad + aislamiento morfológico + OCR paralelo</p></div>', unsafe_allow_html=True)
