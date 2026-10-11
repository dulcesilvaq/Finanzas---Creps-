import streamlit as st

import bcp_txt
import crep

st.set_page_config(page_title="Generador de códigos CREP", page_icon="🧾", layout="wide")

st.title("🧾 Generador de códigos CREP")

tab1, tab2 = st.tabs(["1️⃣ Generar Excel de códigos", "2️⃣ Generar archivos para el BCP"])

# =====================================================================
# PESTAÑA 1: reporte de INTEGRA  ->  Excel con códigos CREP
# =====================================================================
with tab1:
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

    archivo = st.file_uploader("Reporte de INTEGRA (.xlsx)", type=["xlsx"], key="reporte")
    h_como_valor = st.checkbox(
        "Guardar la columna H como texto fijo (en vez de fórmula)",
        value=False,
        help="Por defecto H queda como fórmula, igual que en Google Sheets. Actívalo si necesitas el valor ya calculado.",
    )

    if st.button("Generar Excel", type="primary", disabled=archivo is None, key="btn_excel"):
        try:
            with st.spinner("Procesando el reporte (puede tardar unos segundos)..."):
                st.session_state["resultado"] = crep.procesar(archivo.getvalue(), h_como_valor)
        except Exception as e:
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
            key="dl_excel",
        )

        t_s, t_d, t_m = st.tabs(["Vista previa soles", "Vista previa dólares", "Monedas del reporte"])
        with t_s:
            st.caption("Primeras 30 filas. El Excel contiene todas.")
            st.dataframe(res["vista_soles"], use_container_width=True, hide_index=True)
        with t_d:
            st.caption("Primeras 30 filas. El Excel contiene todas.")
            st.dataframe(res["vista_dolares"], use_container_width=True, hide_index=True)
        with t_m:
            st.caption("Para comprobar que no se quedó fuera ninguna escritura de soles o dólares.")
            st.dataframe(res["monedas"], use_container_width=True, hide_index=True)

# =====================================================================
# PESTAÑA 2: Excel corregido  ->  .txt para subir al BCP
# =====================================================================
with tab2:
    st.write(
        "Cuando tu Excel ya esté corregido, súbelo aquí. La app toma el código de la columna **H** de "
        "**Código libre soles** y **Código libre dólares**, arma la cabecera y te da los dos archivos para el BCP."
    )

    with st.expander("¿Qué contiene cada archivo?"):
        st.markdown(
            """
- **Primera línea (cabecera):** fecha, número de registros y monto total (suma de las cuotas, en centavos).
- **Resto de líneas:** el código CREP de cada fila de tu Excel, en el mismo orden.
- **Nombres:** `crep_s_DDMM.txt` (soles) y `crep_d_DDMM.txt` (dólares), con el día y mes de la fecha elegida.
- Antes de generar, la app revisa cada línea (largo, convenio, fechas, montos, caracteres) y avisa si algo está mal.
            """
        )

    archivo_bcp = st.file_uploader("Excel corregido (.xlsx)", type=["xlsx"], key="excel_corregido")
    fecha = st.date_input(
        "Fecha de la cabecera",
        value=bcp_txt.hoy_lima().date(),
        format="DD/MM/YYYY",
        help="Por defecto es la fecha de hoy (hora de Perú).",
    )

    if st.button("Generar archivos BCP", type="primary", disabled=archivo_bcp is None, key="btn_bcp"):
        try:
            with st.spinner("Armando los archivos..."):
                contenido = archivo_bcp.getvalue()
                salida = {}
                for tipo in ("S", "D"):
                    datos, resumen, avisos, problemas = bcp_txt.generar_txt(contenido, tipo, fecha)
                    salida[tipo] = {"datos": datos, "resumen": resumen, "avisos": avisos, "problemas": problemas}
                st.session_state["bcp"] = salida
        except Exception as e:
            st.session_state.pop("bcp", None)
            st.error(f"No se pudo procesar el archivo: {e}")

    bcp = st.session_state.get("bcp")
    if bcp:
        hay_problemas = any(v["problemas"] for v in bcp.values())
        if hay_problemas:
            st.error("Hay líneas con problemas. Corrígelas en el Excel antes de subir los archivos al BCP.")
        else:
            st.success("Listo. Los archivos pasaron todas las validaciones.")

        etiquetas = {"S": "SOLES", "D": "DÓLARES"}
        col_s, col_d = st.columns(2)
        for col, tipo in ((col_s, "S"), (col_d, "D")):
            v = bcp[tipo]
            r = v["resumen"]
            with col:
                st.subheader(etiquetas[tipo])
                st.metric("Registros", f"{r['registros']:,}")
                simbolo = "S/" if tipo == "S" else "$"
                st.metric("Monto total", f"{simbolo} {r['total']:,.2f}")
                st.caption(f"Fecha en cabecera: {r['fecha']}")
                if r["modificadas_a_mano"]:
                    st.info(f"{r['modificadas_a_mano']} códigos de H son distintos a lo que darían las columnas A–G "
                            "(modificaciones manuales). Se respetó lo que está en H.")
                for aviso in v["avisos"]:
                    st.warning(aviso)
                if v["problemas"]:
                    st.error(f"{len(v['problemas'])} líneas con problemas (se muestran las primeras 20):")
                    for p in v["problemas"][:20]:
                        st.write(f"• {p}")
                st.download_button(
                    f"⬇️ Descargar {r['archivo']}",
                    data=v["datos"],
                    file_name=r["archivo"],
                    mime="text/plain",
                    type="primary",
                    key=f"dl_{tipo}",
                )
