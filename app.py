import streamlit as st

import crep

st.set_page_config(page_title="Generador de códigos CREP", page_icon="🧾", layout="wide")

st.title("🧾 Generador de códigos CREP")
st.write(
    "Sube el reporte de flujo de INTEGRA, presiona **Generar** y descarga el Excel con las hojas "
    "**Código libre soles** y **Código libre dólares**."
)

with st.expander("¿Qué hace exactamente?"):
    st.markdown(
        """
**1. Filtros** (se aplican al reporte, sin cambiar el orden de las filas)
- **Moneda (R):** soles para el flujo de soles; dólares para el de dólares (acepta *dolares*, *Dolares*, *Dólares*).
- **Saldo Pendiente Dolar (U):** se quitan las filas con 0 exacto.
- **Fecha Venc. Actual (K):** del 01/01/2023 en adelante.

**2. Códigos** (misma lógica del Apps Script)
- A = Código Matrícula · B = nombre limpio para BCP · C = cuota · D/E/F = día/mes/año (con ajuste de duplicados) · G = Monto Cuota Actual (L) · H = código CREP.
        """
    )

archivo = st.file_uploader("Reporte de INTEGRA (.xlsx)", type=["xlsx"])
h_como_valor = st.checkbox(
    "Guardar la columna H como texto fijo (en vez de fórmula)",
    value=False,
    help="Por defecto H queda como fórmula, igual que en Google Sheets. Actívalo si necesitas el valor ya calculado.",
)

if st.button("Generar", type="primary", disabled=archivo is None):
    try:
        with st.spinner("Procesando el reporte (puede tardar unos segundos)..."):
            st.session_state["resultado"] = crep.procesar(archivo.getvalue(), h_como_valor)
            st.session_state["nombre_origen"] = archivo.name
    except Exception as e:  # mensaje claro en vez de un error técnico
        st.session_state.pop("resultado", None)
        st.error(f"No se pudo procesar el archivo: {e}")

res = st.session_state.get("resultado")
if res:
    st.success("Listo. Tu Excel está generado.")

    c1, c2, c3 = st.columns(3)
    c1.metric("Filas en el reporte", f"{res['n_reporte']:,}")
    c2.metric("Filas SOLES", f"{res['n_soles']:,}")
    c3.metric("Filas DÓLARES", f"{res['n_dolares']:,}")

    for aviso in res["avisos"]:
        st.warning(aviso)

    st.download_button(
        "⬇️ Descargar Excel",
        data=res["excel"],
        file_name="CodigosCREP.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    tab_s, tab_d, tab_m = st.tabs(["Vista previa soles", "Vista previa dólares", "Monedas del reporte"])
    with tab_s:
        st.caption("Primeras 30 filas. El Excel contiene todas.")
        st.dataframe(res["vista_soles"], use_container_width=True, hide_index=True)
    with tab_d:
        st.caption("Primeras 30 filas. El Excel contiene todas.")
        st.dataframe(res["vista_dolares"], use_container_width=True, hide_index=True)
    with tab_m:
        st.caption("Para comprobar que no se quedó fuera ninguna escritura de soles o dólares.")
        st.dataframe(res["monedas"], use_container_width=True, hide_index=True)
