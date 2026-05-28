# 📋 Foliador de Obras Sociales

Aplicación web que extrae códigos de cuenta (`C` + número) del margen superior izquierdo de PDFs escaneados y genera un Excel con los rangos de páginas.

## Rendimiento — Motor de 3 fases

| Fase | Qué hace | Velocidad | 2000 págs |
|------|----------|-----------|-----------|
| **1. Densidad de píxeles** | Detecta cuáles páginas tienen tinta en el margen | ~47ms/pág | ~94s |
| **2a. Texto embebido** | Lee texto nativo del PDF (si existe) | ~3ms/pág | instantáneo |
| **2b. OCR paralelo** | Tesseract solo en candidatas (~5% del total) | ~2s/pág × 4 hilos | ~50s |
| **Total** | | | **~2.5 min** |

> La Fase 1 escanea **todas** las páginas buscando tinta oscura en el margen superior izquierdo. Solo las páginas candidatas (~5%) pasan a OCR. Esto hace que el tiempo sea prácticamente independiente del tamaño del documento.

## Flujo

1. **Cargar** PDF (hasta 500 MB)
2. **Escanear** — detección automática en 3 fases con miniaturas de candidatas
3. **Editar** — tabla interactiva para corregir códigos y completar profesional
4. **Descargar** — Excel formateado

## Deploy en Streamlit Cloud

1. Fork este repositorio
2. Ir a [share.streamlit.io](https://share.streamlit.io)
3. Conectar el repo → seleccionar `app.py` → **Deploy**

## Instalación local

```bash
git clone https://github.com/tu-usuario/foliador-obras-sociales.git
cd foliador-obras-sociales
sudo apt-get install tesseract-ocr tesseract-ocr-spa
pip install -r requirements.txt
streamlit run app.py
```

## Tecnologías

- **Streamlit** — UI web
- **PyMuPDF** — Renderizado PDF + extracción de texto
- **NumPy** — Análisis de densidad de píxeles (Fase 1)
- **Tesseract OCR** — Reconocimiento de texto (Fase 2b)
- **openpyxl** — Generación Excel
- **concurrent.futures** — Paralelismo en OCR

---
**v2.0** · Motor de 3 fases · Detección por densidad + OCR paralelo
