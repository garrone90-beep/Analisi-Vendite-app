# 🍷 Dashboard Vendite — Streamlit (compatibile con i file Excel Dropbox v4/v5)

Questa app:
- Mostra i **link pubblici Dropbox** in sidebar (dl=0) e usa automaticamente **dl=1** per leggere i dati.
- Se non trova un foglio riga-level (es. `Dati`), **legge le TABELLE** dal foglio **`Dashboard`** (v4/v5).
- Riproduce KPI per anno, YoY, Top Produttori/Tipologie, Volumi; cutoff come nei file.

## Avvio
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Configurazione link
I link pubblici sono nel codice (`PUBLIC_DROPBOX_URL_DETTAGLIO` / `PUBLIC_DROPBOX_URL_INGROSSO`).
L’app usa i corrispondenti link **raw** (`dl=1`) per caricare i dati. Puoi anche impostare:
- `URL_DETTAGLIO` e `URL_INGROSSO` in `.streamlit/secrets.toml`, oppure
- incollarli nei campi della sidebar.

## Note
Se i nomi o la struttura dei fogli cambiano, l’app cercherà:
1) dati riga-level (colonne: `DataTurno`, `Prezzo`, `PrezzoTot`, `Qta`, `DettaglioIngrosso`), altrimenti
2) le tabelle nella pagina `Dashboard` con intestazioni riconoscibili (Anno/Fatturato, Mese + anni, Produttore + Totale, Tipologia + Totale, Anno + Bottiglie_Totali).
