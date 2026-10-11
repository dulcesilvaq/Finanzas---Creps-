# Generador de códigos CREP

App web (Streamlit): subes el reporte de flujo de INTEGRA, presionas **Generar** y descargas un Excel con
las hojas `Flujo soles`, `Flujo dólares`, `Código libre soles` y `Código libre dólares`.

## Archivos
- `app.py`: la interfaz (subir archivo, botón Generar, descargar).
- `crep.py`: la lógica (filtros + generación del código CREP).
- `requirements.txt`: librerías necesarias.

## Reglas
- Filtros: Moneda (R) soles/dólares · Saldo Pendiente Dolar (U) distinto de 0 · Fecha Venc. Actual (K) desde 01/01/2023.
- Código: misma lógica del Apps Script original. Convenio y fecha de proceso por moneda se editan al inicio de `crep.py`.

## Ejecutar en tu PC
```
pip install -r requirements.txt
streamlit run app.py
```

> No subas a este repositorio reportes con datos de clientes. El archivo se procesa en memoria y no se guarda.
