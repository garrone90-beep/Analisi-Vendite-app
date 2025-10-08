# 🍷 Dashboard Vendite — Streamlit (con link pubblici Dropbox)

Questa app mostra **sempre** i link pubblici Dropbox ai file Excel (pagina di anteprima, `dl=0`) e usa automaticamente i link **raw** (`dl=1`) per leggere i dati.

## Setup
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Dove cambiare i link
- Nel codice: variabili `PUBLIC_DROPBOX_URL_DETTAGLIO` e `PUBLIC_DROPBOX_URL_INGROSSO` (mostrate in sidebar).
- L’app usa i corrispondenti link **raw** (con `dl=1`) per caricare i dati.
- In alternativa, puoi impostare `URL_DETTAGLIO` e `URL_INGROSSO` in `.streamlit/secrets.toml`.
