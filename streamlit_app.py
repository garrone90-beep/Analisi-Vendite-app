# streamlit_app.py
# -------------------------------------------------------------
# Dashboard Vendite Enoteca (Streamlit)
# - KPI con variazione % YoY per TUTTE le colonne
# - Toggle rapido in pagina: Canale (Dettaglio/Ingrosso) e Visualizzazione (Pari periodo/Anno completo)
# - Ordine: Metriche anno selezionato -> Tabella KPI -> Grafici
# - Link pubblici Dropbox visibili in sidebar (no toggle lì)
# - Tutti i grafici a colonne (barre raggruppate)
# -------------------------------------------------------------

import io
import requests
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="Dashboard Vendite Enoteca", page_icon="🍷", layout="wide")

# --- Link pubblici Dropbox (pagina di anteprima) ---
PUBLIC_DROPBOX_URL_DETTAGLIO_PARIPER = "https://www.dropbox.com/scl/fi/ajoax5i3gthl1kg3ewe3t/AnalisiVendite_Dettaglio_Dashboard_v4.xlsx?rlkey=nfj3kg0s3n6yp206honxfhr1j&dl=0"
PUBLIC_DROPBOX_URL_DETTAGLIO_FULL   = "https://www.dropbox.com/scl/fi/a1fdmmos6zbd835w9iyx9/AnalisiVendite_Dettaglio_Dashboard_v3.xlsx?rlkey=9eht1v11lx0cpx2tzzarvxrsh&dl=0"
PUBLIC_DROPBOX_URL_INGROSSO_PARIPER = "https://www.dropbox.com/scl/fi/608xyutty9zyufb57ddde/AnalisiVendite_Ingrosso_Dashboard_v5.xlsx?rlkey=0c8haj4csbyuf75o3yeuc2c82&dl=0"

def to_direct_url(url: str) -> str:
    if not url:
        return url
    if "github.com" in url and "/blob/" in url:
        return url.replace("github.com", "raw.githubusercontent.com").replace("/blob/", "/")
    if "dropbox.com" in url and "dl=0" in url:
        return url.replace("dl=0", "dl=1")
    return url

URL_DETTAGLIO_PARIPER = to_direct_url(PUBLIC_DROPBOX_URL_DETTAGLIO_PARIPER)
URL_DETTAGLIO_FULL    = to_direct_url(PUBLIC_DROPBOX_URL_DETTAGLIO_FULL)
URL_INGROSSO_PARIPER  = to_direct_url(PUBLIC_DROPBOX_URL_INGROSSO_PARIPER)

@st.cache_data(show_spinner=True, ttl=600)
def load_excel(url: str) -> dict:
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    bio = io.BytesIO(r.content)
    xl = pd.ExcelFile(bio, engine="openpyxl")
    return {name: xl.parse(name, header=None) for name in xl.sheet_names}

# -------------------- Parser tabelle dal foglio "Dashboard" --------------------
def _find_header_row(df: pd.DataFrame, required_cols: list[str]):
    for i in range(len(df)):
        row_vals = [str(x).strip() for x in df.iloc[i].tolist()]
        pos = {}
        ok = True
        for col in required_cols:
            if col in row_vals:
                pos[col] = row_vals.index(col)
            else:
                ok = False; break
        if ok:
            return i, pos
    raise ValueError(f"Header non trovato per colonne: {required_cols}")

def _extract_table(df: pd.DataFrame, header_row: int, start_col: int, n_cols: int) -> pd.DataFrame:
    headers = [str(x).strip() for x in df.iloc[header_row, start_col:start_col+n_cols].tolist()]
    out_rows = []
    r = header_row + 1
    while r < len(df):
        row = df.iloc[r, start_col:start_col+n_cols].tolist()
        if all((x is None) or (str(x).strip() in ("", "None", "nan")) for x in row):
            break
        out_rows.append(row)
        r += 1
    return pd.DataFrame(out_rows, columns=headers)

def parse_dashboard_tables(sheets: dict) -> dict:
    dashboard_name = None
    for name in sheets.keys():
        if name.strip().lower() == "dashboard":
            dashboard_name = name; break
    if dashboard_name is None:
        raise ValueError("Foglio 'Dashboard' non trovato nel file.")
    df = sheets[dashboard_name].copy()

    # KPI per Anno
    kpi_headers = ["Anno","Fatturato_Netto","Num_Vendite","Prezzo_Medio_Articolo","Fatturato_Medio_Mensile","Margine_Stimato"]
    kpi_row, kpi_pos = _find_header_row(df, kpi_headers)
    kpi_start = min(kpi_pos.values())
    kpi = _extract_table(df, kpi_row, kpi_start, len(kpi_headers))
    for c in kpi_headers[1:]:
        kpi[c] = pd.to_numeric(kpi[c], errors="coerce")
    # Aggiungi variazione % YoY per TUTTE le colonne numeriche (eccetto 'Anno')
    kpi = kpi.sort_values("Anno")
    for col in [c for c in kpi.columns if c != "Anno"]:
        kpi[f"{col}_YoY%"] = kpi[col].pct_change() * 100

    # Fatturato mensile YoY
    def find_monthly_after(row_after: int):
        for i in range(row_after+1, len(df)):
            row_vals = [str(x).strip() for x in df.iloc[i].tolist()]
            if "Mese" in row_vals:
                c0 = row_vals.index("Mese")
                n = 1; j = c0 + 1
                while j < df.shape[1]:
                    val = df.iloc[i, j]
                    if val is None or str(val).strip() in ("", "None", "nan"):
                        break
                    n += 1; j += 1
                return i, c0, n
        raise ValueError("Tabella 'Mese + anni' non trovata.")
    rev_row, rev_col, rev_n = find_monthly_after(kpi_row)
    rev = _extract_table(df, rev_row, rev_col, rev_n)
    for c in rev.columns:
        if c != "Mese":
            rev[c] = pd.to_numeric(rev[c], errors="coerce")

    # Top Produttori
    def find_index_totale_after(row_after: int, key_header: str):
        for i in range(row_after+1, len(df)):
            row_vals = [str(x).strip() for x in df.iloc[i].tolist()]
            if key_header in row_vals and "Totale" in row_vals:
                c_key = row_vals.index(key_header)
                c_tot = row_vals.index("Totale")
                return i, c_key, (c_tot - c_key + 1)
        raise ValueError(f"Tabella '{key_header}' + 'Totale' non trovata.")
    prod_row, prod_col, prod_n = find_index_totale_after(rev_row, "Produttore_Descrizione")
    prod = _extract_table(df, prod_row, prod_col, prod_n)
    for c in prod.columns:
        if c not in ("Produttore_Descrizione",):
            prod[c] = pd.to_numeric(prod[c], errors="coerce")

    # Top Tipologie
    tip_row, tip_col, tip_n = find_index_totale_after(prod_row, "TipologiaVino_Descrizione")
    tip = _extract_table(df, tip_row, tip_col, tip_n)
    for c in tip.columns:
        if c not in ("TipologiaVino_Descrizione",):
            tip[c] = pd.to_numeric(tip[c], errors="coerce")

    # Volumi mensili
    qty_row, qty_col, qty_n = find_monthly_after(tip_row)
    qty = _extract_table(df, qty_row, qty_col, qty_n)
    for c in qty.columns:
        if c != "Mese":
            qty[c] = pd.to_numeric(qty[c], errors="coerce")

    # Volumi totali per anno
    qtyyear_headers = ["Anno","Bottiglie_Totali"]
    qtyyear_row, qtyyear_pos = _find_header_row(df, qtyyear_headers)
    qtyyear_start = min(qtyyear_pos.values())
    qty_year = _extract_table(df, qtyyear_row, qtyyear_start, len(qtyyear_headers))
    qty_year["Bottiglie_Totali"] = pd.to_numeric(qty_year["Bottiglie_Totali"], errors="coerce")

    # Cutoff (testo presente in alto)
    cutoff_text = ""
    for i in range(0, min(6, len(df))):
        row_txt = " ".join([str(x) for x in df.iloc[i].tolist() if not pd.isna(x)])
        if "Periodo confrontato" in row_txt or "Periodo" in row_txt:
            cutoff_text = row_txt.strip(); break

    return {"kpi": kpi, "rev": rev, "prod": prod, "tip": tip, "qty": qty, "qty_year": qty_year, "cutoff_text": cutoff_text}

# -------------------- UI --------------------
st.title("Dashboard Vendite Enoteca")

# --- TOGGLE RAPIDI IN PAGINA ---
left, right = st.columns([1,1])
with left:
    try:
        canale = st.segmented_control("Canale", ["Dettaglio","Ingrosso"], selection="Dettaglio")
    except Exception:
        canale = st.radio("Canale", ["Dettaglio","Ingrosso"], horizontal=True, index=0)
with right:
    try:
        visual = st.segmented_control("Visualizzazione", ["Pari periodo","Anno completo"], selection="Pari periodo")
    except Exception:
        visual = st.radio("Visualizzazione", ["Pari periodo","Anno completo"], horizontal=True, index=0)

# Sorgenti: sidebar solo per link
with st.sidebar:
    st.header("Sorgenti dati (Dropbox)")
    st.markdown("**Link pubblici:**")
    st.markdown(f"- Dettaglio (Pari periodo): [link pubblico]({PUBLIC_DROPBOX_URL_DETTAGLIO_PARIPER})")
    st.markdown(f"- Dettaglio (Anno completo): [link pubblico]({PUBLIC_DROPBOX_URL_DETTAGLIO_FULL})")
    st.markdown(f"- Ingrosso (Pari periodo): [link pubblico]({PUBLIC_DROPBOX_URL_INGROSSO_PARIPER})")
    ingr_full_override = st.text_input("Ingrosso (Anno completo) — URL opzionale (raw o dl=1)", "")
    st.caption("Se non imposti l'URL per Ingrosso 'Anno completo', verrà usato il file 'Pari periodo'.")

# Scelta URL
if canale == "Dettaglio":
    url = URL_DETTAGLIO_PARIPER if visual == "Pari periodo" else URL_DETTAGLIO_FULL
else:
    if visual == "Pari periodo":
        url = URL_INGROSSO_PARIPER
    else:
        url = to_direct_url(ingr_full_override.strip()) if ingr_full_override.strip() else URL_INGROSSO_PARIPER

# Caricamento & parsing
try:
    sheets = load_excel(url)
    parsed = parse_dashboard_tables(sheets)
except Exception as e:
    st.error("Errore nel caricamento o parsing: " + str(e))
    st.stop()

# Badge e periodo
st.markdown(f"**Vista:** {canale} • **{visual}**")
cutoff_text = parsed.get("cutoff_text","")
if cutoff_text:
    st.caption(cutoff_text)

# ====== 1) METRICHE DELL'ANNO SELEZIONATO (IN PRIMO PIANO) ======
kpi = parsed["kpi"]
if not kpi.empty and "Anno" in kpi.columns:
    anni = kpi["Anno"].tolist()
    anno_sel = st.select_slider("Anno selezionato", options=anni, value=anni[-1] if len(anni)>0 else None)
    row = kpi.loc[kpi["Anno"]==anno_sel].iloc[0].to_dict()

    c1, c2, c3, c4, c5 = st.columns(5)
    def eur(x): 
        try: return f"€ {float(x):,.2f}"
        except: return str(x)
    with c1:
        delta = row.get("Fatturato_Netto_YoY%", np.nan)
        st.metric("Fatturato netto", eur(row.get("Fatturato_Netto",0)), None if (delta is None or np.isnan(delta)) else f"{delta:+.1f}%")
    with c2:
        delta = row.get("Num_Vendite_YoY%", np.nan)
        num = row.get("Num_Vendite", None)
        st.metric("N° vendite", "-" if num is None or pd.isna(num) else int(num), None if (delta is None or np.isnan(delta)) else f"{delta:+.1f}%")
    with c3:
        delta = row.get("Prezzo_Medio_Articolo_YoY%", np.nan)
        st.metric("Prezzo medio articolo", eur(row.get("Prezzo_Medio_Articolo",0)), None if (delta is None or np.isnan(delta)) else f"{delta:+.1f}%")
    with c4:
        delta = row.get("Fatturato_Medio_Mensile_YoY%", np.nan)
        st.metric("Fatturato medio mensile", eur(row.get("Fatturato_Medio_Mensile",0)), None if (delta is None or np.isnan(delta)) else f"{delta:+.1f}%")
    with c5:
        delta = row.get("Margine_Stimato_YoY%", np.nan)
        st.metric("Margine stimato (40%)", eur(row.get("Margine_Stimato",0)), None if (delta is None or np.isnan(delta)) else f"{delta:+.1f}%")
else:
    st.info("Nessun KPI disponibile.")

st.divider()

# ====== 2) TABELLA KPI COMPLETA ======
if not kpi.empty:
    st.subheader("KPI per Anno (con variazioni YoY%)")
    fmt = {c:"€{:,.2f}" for c in ["Fatturato_Netto","Prezzo_Medio_Articolo","Fatturato_Medio_Mensile","Margine_Stimato"] if c in kpi.columns}
    for c in [col for col in kpi.columns if col.endswith("_YoY%")]:
        fmt[c] = "{:.1f}%"
    st.dataframe(kpi.style.format(fmt), use_container_width=True)

st.divider()

# ====== 3) GRAFICI ======
# Fatturato mensile (barre raggruppate)
rev = parsed["rev"]
if not rev.empty:
    st.subheader("Fatturato mensile (€/Anno) — barre raggruppate")
    rev_long = rev.melt(id_vars="Mese", var_name="Anno", value_name="Valore")
    fig = px.bar(rev_long, x="Mese", y="Valore", color="Anno", barmode="group")
    fig.update_layout(legend_title_text="Anno", xaxis_title="Mese", yaxis_title="€", height=480)
    st.plotly_chart(fig, use_container_width=True)

st.divider()

# Top Produttori (barre raggruppate)
prod = parsed["prod"]
if not prod.empty and "Produttore_Descrizione" in prod.columns:
    st.subheader("Top 10 Produttori — € per Anno (barre raggruppate)")
    prod_long = prod.melt(id_vars="Produttore_Descrizione", var_name="Anno", value_name="Valore")
    prod_long = prod_long[prod_long["Anno"]!="Totale"]
    figp = px.bar(prod_long, x="Produttore_Descrizione", y="Valore", color="Anno", barmode="group")
    figp.update_layout(xaxis_title="Produttore", yaxis_title="€", height=500)
    st.plotly_chart(figp, use_container_width=True)

st.divider()

# Top Tipologie (barre raggruppate)
tip = parsed["tip"]
if not tip.empty and "TipologiaVino_Descrizione" in tip.columns:
    st.subheader("Top 8 Tipologie — € per Anno (barre raggruppate)")
    tip_long = tip.melt(id_vars="TipologiaVino_Descrizione", var_name="Anno", value_name="Valore")
    tip_long = tip_long[tip_long["Anno"]!="Totale"]
    figt = px.bar(tip_long, x="TipologiaVino_Descrizione", y="Valore", color="Anno", barmode="group")
    figt.update_layout(xaxis_title="Tipologia", yaxis_title="€", height=500)
    st.plotly_chart(figt, use_container_width=True)

st.divider()

# Volumi mensili (barre raggruppate)
qty = parsed["qty"]
if not qty.empty:
    st.subheader("Volumi mensili (bottiglie/Anno) — barre raggruppate")
    qty_long = qty.melt(id_vars="Mese", var_name="Anno", value_name="Bottiglie")
    figq = px.bar(qty_long, x="Mese", y="Bottiglie", color="Anno", barmode="group")
    figq.update_layout(legend_title_text="Anno", xaxis_title="Mese", yaxis_title="Bottiglie", height=480)
    st.plotly_chart(figq, use_container_width=True)

# Volumi totali per Anno — TABELLA
qty_year = parsed["qty_year"]
if not qty_year.empty:
    st.subheader("Volumi totali per Anno (bottiglie)")
    st.dataframe(qty_year, use_container_width=True)

st.caption("Tutti i valori economici sono IVA esclusa. Per l'Ingrosso sono escluse le vendite dei produttori 'Cantine Garrone' e 'Cantine Isidoro'.")
