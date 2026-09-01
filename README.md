# Daily Check Combinado

## Archivos del repositorio

Sube estos archivos directamente a la raiz del repositorio:

- `streamlit_app.py`
- `app.py`
- `processor.py`
- `requirements.txt`
- `test_processor.py`

## Configuracion en Streamlit Community Cloud

- Branch: `main`
- Main file path: `streamlit_app.py`
- Python: `3.12`, seleccionado en Advanced settings

No se incluye `runtime.txt`. La version de Python debe seleccionarse desde la configuracion de despliegue de Streamlit.

## Ejecucion local

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```
