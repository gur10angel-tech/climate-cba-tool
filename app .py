# app .py — Streamlit Cloud entry point redirect
# Streamlit Cloud is configured to launch this file.
# All real application code lives in app.py.
import runpy
runpy.run_path('app.py', run_name='__main__')
