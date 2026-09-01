import io
import traceback

import streamlit as st

st.set_page_config(page_title="Daily Check Combinado", layout="centered")
st.title("Daily Check Combinado")
st.caption("Generador de programa diario")
st.write("Carga PROGRAMA PROCESADO y DAILY CHECK PLANTILLA para generar el archivo combinado.")

programa = st.file_uploader(
    "1. PROGRAMA PROCESADO",
    type=["xlsx"],
    accept_multiple_files=False,
    key="programa",
)
plantilla = st.file_uploader(
    "2. DAILY CHECK PLANTILLA",
    type=["xlsx"],
    accept_multiple_files=False,
    key="plantilla",
)

if "resultado" not in st.session_state:
    st.session_state.resultado = None

if programa is None or plantilla is None:
    st.info("Selecciona los dos archivos para habilitar el procesamiento.")
else:
    if st.button("Generar archivo combinado", type="primary", use_container_width=True):
        try:
            # Importacion diferida: la interfaz siempre debe mostrarse aunque falle el procesador.
            from processor import generate_combined

            with st.spinner("Procesando archivos..."):
                st.session_state.resultado = generate_combined(
                    io.BytesIO(programa.getvalue()),
                    io.BytesIO(plantilla.getvalue()),
                )
            st.success("Archivo generado correctamente.")
        except Exception as exc:
            st.session_state.resultado = None
            st.error(f"No fue posible generar el archivo: {exc}")
            with st.expander("Detalle tecnico"):
                st.code(traceback.format_exc())

if st.session_state.resultado:
    st.download_button(
        "Descargar DAILY CHECK COMBINADO.xlsx",
        data=st.session_state.resultado,
        file_name="DAILY CHECK COMBINADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
