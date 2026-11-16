# 📋 Foliador de Obras Sociales

Lee la cuenta manuscrita (`C+NÚMERO`) de las **hojas separadoras** de PDFs escaneados, cruza con la base de profesionales y genera un Excel completo.

## Flujo

1. **Cargar** PDF + Base de usuarios (.xlsx)
2. **Escanear** — detecta las hojas separadoras (hoja en blanco con la cuenta trazada a mano con marcador, ej: `C924`), lee la cuenta y cruza contra la base
3. **Editar** — tabla con datos del profesional auto-completados (Nombre, CUIT, Especialidad, Resp. Fiscal, Arancel)
4. **Descargar** — Excel con toda la información

## Criterio de extracción

Cada cuenta viene precedida por una **hoja en blanco** con la cuenta `C+NÚMERO` escrita a mano con marcador grueso. Puede tener otra palabra manuscrita debajo (ej: `CEDIG`), que se ignora.

- **Fase 1** — se marca como separadora toda página casi vacía cuya tinta es mayormente trazo grueso (las planillas impresas se descartan, aunque tengan anotaciones en el margen).
- **Fase 2** — se recorta el manuscrito y se lee la cuenta con Claude Vision (recomendado, requiere `ANTHROPIC_API_KEY`) o Tesseract como respaldo (poco confiable con manuscrita: revisar siempre las miniaturas).
- Cada cuenta abarca desde su hoja separadora hasta la página anterior a la siguiente separadora.

## Cruce con Base de Usuarios

La app busca cada código detectado (ej: `C1098`) en la columna `Matricula` de la base. Si lo encuentra, completa automáticamente: Nombre, CUIT, Especialidad, Responsabilidad Fiscal y Arancel.

## Rendimiento

| Páginas | Fase 1 (separadoras) | Fase 2 (lectura) | Total |
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
**v5.0** · Lectura de hojas separadoras manuscritas + cruce con base de profesionales
