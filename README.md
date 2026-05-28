# 📋 Foliador de Obras Sociales

Extrae códigos de cuenta del margen superior izquierdo de PDFs escaneados, cruza con la base de profesionales y genera un Excel completo.

## Flujo

1. **Cargar** PDF + Base de usuarios (.xlsx)
2. **Escanear** — detecta páginas con código, aísla el manuscrito, lee por OCR y cruza contra la base
3. **Editar** — tabla con datos del profesional auto-completados (Nombre, CUIT, Especialidad, Resp. Fiscal, Arancel)
4. **Descargar** — Excel con toda la información

## Cruce con Base de Usuarios

La app busca cada código detectado (ej: `C1098`) en la columna `Matricula` de la base. Si lo encuentra, completa automáticamente: Nombre, CUIT, Especialidad, Responsabilidad Fiscal y Arancel.

## Rendimiento

| Páginas | Fase 1 (densidad) | Fase 2 (OCR candidatas) | Total |
|---------|-------------------|------------------------|-------|
| 100     | ~5s               | ~5s                    | ~10s  |
| 2000    | ~94s              | ~50s                   | ~2.5 min |

## Deploy en Streamlit Cloud

1. Fork → [share.streamlit.io](https://share.streamlit.io) → Deploy

## Local

```bash
sudo apt-get install tesseract-ocr tesseract-ocr-spa
pip install -r requirements.txt
streamlit run app.py
```

---
**v3.0** · Detección + aislamiento + cruce con base de profesionales
