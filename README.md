# Daily Check Combinado

## Archivos en la raiz

Sube todos los archivos de este paquete directamente a la raiz del repositorio, sin crear carpetas.

## Streamlit Community Cloud

- Branch: `main`
- Main file path: `streamlit_app.py`

## Ejecucion local

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

La aplicacion conserva la plantilla, agrega tareas coincidentes del programa, mantiene trabajos STORAGE existentes, agrega trabajos STORAGE adicionales cuando coincide la matricula y clasifica el resto como TRANSIT CHECK o TRANSIT CHECK / RON.
