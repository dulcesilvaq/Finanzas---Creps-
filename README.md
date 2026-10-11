# Generador de códigos CREP

App web (Streamlit) con dos pasos:

1. **Generar Excel de códigos:** subes el reporte de flujo de INTEGRA y descargas un Excel con las hojas
   `Flujo soles`, `Flujo dólares`, `Código libre soles` y `Código libre dólares`.
2. **Generar archivos para el BCP:** subes ese Excel ya corregido y descargas `crep_s_DDMM.txt` (soles)
   y `crep_d_DDMM.txt` (dólares), con la cabecera (fecha, n° de registros, monto total) y el detalle.

## Archivos
- `app.py`: la interfaz (dos pestañas).
- `crep.py`: filtros y generación del código CREP (paso 1).
- `bcp_txt.py`: cabecera, validaciones y armado de los .txt (paso 2).
- `requirements.txt`: librerías necesarias.

## Reglas
- Filtros del paso 1: Moneda (R) soles/dólares · Saldo Pendiente Dolar (U) distinto de 0 · Fecha Venc. Actual (K) desde 01/01/2023.
- Convenios, fechas de proceso y nombres de cabecera se editan al inicio de `crep.py` y `bcp_txt.py`.
- Cabecera del .txt: `CC` + cuenta + `C` + nombre + fecha + n° de registros + monto total (centavos) + `A` + espacios + `000`.

## Ejecutar en tu PC
```
pip install -r requirements.txt
streamlit run app.py
```

> No subas a este repositorio reportes ni Excel con datos de clientes. Los archivos se procesan en memoria y no se guardan.
