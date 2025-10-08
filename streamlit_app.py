# streamlit_app.py
# -------------------------------------------------------------
# Dashboard Vendite (Streamlit) – compatibile con i tuoi file Excel su Dropbox
# Novità:
#  - Mostra link pubblici Dropbox (dl=0) e usa link raw (dl=1) per leggere
#  - Se non trova dati riga-level, legge le TABELLE dal foglio "Dashboard" (v4/v5)
#  - Messaggi di errore corretti (mostra il dettaglio reale dell'eccezione)
# -------------------------------------------------------------

import io
import re
import requests
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="Dashboard Vendite", page_icon="🍷", layout="wide")

# --- Link pubblici Dropbox (pagina di anteprima) ---
PUBLIC_DROPBOX_URL_DETTAGLIO = "https://www.dropbox.com/scl/fi/ajoax5i3gthl1kg3ewe3t/AnalisiVendite_Dettaglio_Dashboard_v4.xlsx?rlkey=nfj3kg0s3n6yp206honxfhr1j&dl=0"
PUBLIC_DROPBOX_URL_INGROSSO  = "https://www.dropbox.com/scl/fi/608xyutty9zyufb57ddde/AnalisiVendite_Ingrosso_Dashboard_v5.xlsx?rlkey=0c8haj4csbyuf75o3yeuc2c82&dl=0"

# --- Link diretti per il download (dl=1) ---
def to_direct_url(url: str) -> str:
    if not url:
        return url
    if "github.com" in url and "/blob/" in url:
        return url.replace("github.com", "raw.githubusercontent.com").replace("/blob/", "/")
    if "dropbox.com" in url and "dl=0" in url:
        return url.replace("dl=0", "dl=1")
    return url

URL_DETTAGLIO = st.secrets.get("URL_DETTAGLIO", to_direct_url(PUBLIC_DROPBOX_URL_DETTAGLIO))
URL_INGROSSO  = st.secrets.get("URL_INGROSSO",  to_direct_url(PUBLIC_DROPBOX_URL_INGROSSO))

@st.cache_data(show_spinner=True, ttl=600)
def load_excel(url: str) -> dict:
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    bio = io.BytesIO(r.content)
    xl = pd.ExcelFile(bio, engine="openpyxl")
    return {name: xl.parse(name, header=None) for name in xl.sheet_names}

# -------------------- Parser tabelle dal foglio "Dashboard" --------------------
def _find_header_row(df: pd.DataFrame, required_cols: list[str]) -> tuple[int, dict]:
    """
    Cerca una riga che contenga tutte le colonne richieste, ovunque sulla riga.
    Restituisce (row_idx, col_pos_map).
    """
    for i in range(len(df)):
        row_vals = [str(x).strip() for x in df.iloc[i].tolist()]
        positions = {}
        ok = True
        for col in required_cols:
            if col in row_vals:
                positions[col] = row_vals.index(col)
            else:
                ok = False
                break
        if ok:
            return i, positions
    raise ValueError(f"Header non trovato per colonne: {required_cols}")

def _extract_table(df: pd.DataFrame, header_row: int, start_col: int, n_cols: int) -> pd.DataFrame:
    headers = [str(x).strip() for x in df.iloc[header_row, start_col:start_col+n_cols].tolist()]
    # Righe successive fino a riga vuota
    out_rows = []
    r = header_row + 1
    while r < len(df):
        row = df.iloc[r, start_col:start_col+n_cols].tolist()
        if all((x is None) or (str(x).strip() in ("", "None", "nan")) for x in row):
            break
        out_rows.append(row)
        r += 1
    table = pd.DataFrame(out_rows, columns=headers)
    return table

def parse_dashboard_tables(sheets: dict) -> dict:
    # trova il foglio "Dashboard"
    dashboard_name = None
    for name in sheets.keys():
        if name.strip().lower() == "dashboard":
            dashboard_name = name
            break
    if dashboard_name is None:
        raise ValueError("Foglio 'Dashboard' non trovato nel file.")
    df = sheets[dashboard_name].copy()

    # 1) Tabella 'Totale per Anno' (KPI)
    kpi_headers = ["Anno","Fatturato_Netto","Num_Vendite","Prezzo_Medio_Articolo","Fatturato_Medio_Mensile","Margine_Stimato"]
    kpi_row, kpi_pos = _find_header_row(df, kpi_headers)
    kpi_start_col = min(kpi_pos.values())
    kpi_table = _extract_table(df, kpi_row, kpi_start_col, len(kpi_headers))
    # cast numeri
    for c in kpi_headers[1:]:
        kpi_table[c] = pd.to_numeric(kpi_table[c], errors="coerce")

    # 2) Fatturato mensile YoY (cerca 'Mese' come prima colonna)
    #    Prende la PRIMA tabella con colonna 'Mese' dopo la tabella KPI
    def find_monthly_after(row_after: int) -> tuple[int, int, int]:
        for i in range(row_after+1, len(df)):
            row_vals = [str(x).strip() for x in df.iloc[i].tolist()]
            if "Mese" in row_vals:
                c0 = row_vals.index("Mese")
                # conta quante colonne piene (Mese + anni)
                n = 1
                j = c0 + 1
                while j < df.shape[1]:
                    val = df.iloc[i, j]
                    if val is None or str(val).strip() in ("", "None", "nan"):
                        break
                    n += 1
                    j += 1
                return i, c0, n
        raise ValueError("Tabella mensile 'Mese + anni' non trovata.")

    rev_row, rev_col, rev_n = find_monthly_after(kpi_row)
    rev_table = _extract_table(df, rev_row, rev_col, rev_n)
    # tipicamente 'Mese' + anni (colonne anni in int)
    for c in rev_table.columns:
        if c != "Mese":
            rev_table[c] = pd.to_numeric(rev_table[c], errors="coerce")

    # 3) Dopo la prima tabella mensile, cerco 'Top 10 Produttori'
    #    Header con 'Produttore_Descrizione' ... 'Totale'
    prod_headers = ["Produttore_Descrizione","Totale"]
    def find_index_totale_after(row_after: int, key_header: str):
        for i in range(row_after+1, len(df)):
            row_vals = [str(x).strip() for x in df.iloc[i].tolist()]
            if key_header in row_vals and "Totale" in row_vals:
                c_key = row_vals.index(key_header)
                # numero colonne fino a 'Totale' incluso
                c_tot = row_vals.index("Totale")
                n = (c_tot - c_key) + 1
                return i, c_key, n
        raise ValueError(f"Tabella con '{key_header}' e 'Totale' non trovata.")

    prod_row, prod_col, prod_n = find_index_totale_after(rev_row, "Produttore_Descrizione")
    prod_table = _extract_table(df, prod_row, prod_col, prod_n)
    # cast numeric per anni + Totale
    for c in prod_table.columns:
        if c not in ("Produttore_Descrizione",):
            prod_table[c] = pd.to_numeric(prod_table[c], errors="coerce")

    # 4) Tipologie
    tip_row, tip_col, tip_n = find_index_totale_after(prod_row, "TipologiaVino_Descrizione")
    tip_table = _extract_table(df, tip_row, tip_col, tip_n)
    for c in tip_table.columns:
        if c not in ("TipologiaVino_Descrizione",):
            tip_table[c] = pd.to_numeric(tip_table[c], errors="coerce")

    # 5) Volumi mensili (seconda tabella 'Mese' dopo tipologie)
    qty_row, qty_col, qty_n = find_monthly_after(tip_row)
    qty_table = _extract_table(df, qty_row, qty_col, qty_n)
    for c in qty_table.columns:
        if c != "Mese":
            qty_table[c] = pd.to_numeric(qty_table[c], errors="coerce")

    # 6) Volumi totali per anno (cerca 'Anno' + 'Bottiglie_Totali')
    qtyyear_headers = ["Anno","Bottiglie_Totali"]
    qtyyear_row, qtyyear_pos = _find_header_row(df, qtyyear_headers)
    qtyyear_start = min(qtyyear_pos.values())
    qtyyear_table = _extract_table(df, qtyyear_row, qtyyear_start, len(qtyyear_headers))
    qtyyear_table["Bottiglie_Totali"] = pd.to_numeric(qtyyear_table["Bottiglie_Totali"], errors="coerce")

    # 7) Cutoff: prova a leggere una cella che contenga 'Periodo confrontato:' nelle prime righe
    cutoff_text = ""
    for i in range(0, min(6, len(df))):
        row_txt = " ".join([str(x) for x in df.iloc[i].tolist() if not pd.isna(x)])
        if "Periodo confrontato" in row_txt:
            cutoff_text = row_txt.strip()
            break

    return {
        "mode": "aggregated",
        "kpi": kpi_table,
        "rev": rev_table,
        "prod": prod_table,
        "tip": tip_table,
        "qty": qty_table,
        "qty_year": qtyyear_table,
        "cutoff_text": cutoff_text
    }

# -------------------- Pipeline riga-level (se disponibile) --------------------
PRODUTTORI_DA_ESCLUDERE_INGROSSO = {"Cantine Garrone","Cantine Isidoro"}

def try_row_level(sheets: dict, canale: str):
    # Cerca un foglio con intestazioni standard (header alla riga 0)
    candidate = None
    for name, tab in sheets.items():
        # ricarico questo sheet con header=0 per tentare
        # (sheets[name] ha header=None)
        df = tab.copy()
        # prova: prima riga come header reale
        df2 = df.copy()
        df2.columns = df2.iloc[0].astype(str).tolist()
        df2 = df2.iloc[1:].reset_index(drop=True)
        cols = {c.strip().lower() for c in df2.columns.astype(str)}
        if {"dataturno","prezzo","prezzotot","qta"}.issubset(cols):
            candidate = df2
            break
    if candidate is None:
        raise ValueError("Dati riga-level non trovati.")
    df = candidate.rename(columns={
        "dataturno":"DataTurno",
        "prezzo":"Prezzo",
        "prezzotot":"PrezzoTot",
        "qta":"Qta",
        "dettaglioingrosso":"DettaglioIngrosso",
        "produttore_descrizione":"Produttore_Descrizione",
        "tipologiavino_descrizione":"TipologiaVino_Descrizione",
    })
    df["DataTurno"] = pd.to_datetime(df["DataTurno"], errors="coerce")
    if canale == "Dettaglio":
        df = df[df.get("DettaglioIngrosso", 0).fillna(0).astype(int) == 0]
        df["Prezzo_Netto"] = pd.to_numeric(df["Prezzo"], errors="coerce")/1.22
        df["PrezzoTot_Netto"] = pd.to_numeric(df["PrezzoTot"], errors="coerce")/1.22
    else:
        df = df[df.get("DettaglioIngrosso", 1).fillna(1).astype(int) == 1]
        if "Produttore_Descrizione" in df.columns:
            df = df[~df["Produttore_Descrizione"].isin(PRODUTTORI_DA_ESCLUDERE_INGROSSO)]
        df["Prezzo_Netto"] = pd.to_numeric(df["Prezzo"], errors="coerce")
        df["PrezzoTot_Netto"] = pd.to_numeric(df["PrezzoTot"], errors="coerce")

    # Escludi alimentari
    if "TipologiaVino_Descrizione" in df.columns:
        tip = df["TipologiaVino_Descrizione"].astype(str).str.upper()
        mask_exclude = tip.str.contains("PRODOTTO|ALIMENT|FOOD|GASTRON|SNACK|CONSERVE|DOLCI")
        df = df.loc[~mask_exclude].copy()

    df["Anno"] = df["DataTurno"].dt.year
    df["MeseNum"] = df["DataTurno"].dt.month
    df["GiornoNum"] = df["DataTurno"].dt.day

    cutoff = df["DataTurno"].max()
    if pd.isna(cutoff):
        cutoff_text = ""
    else:
        m, d = cutoff.month, cutoff.day
        mask = (df["MeseNum"] < m) | ((df["MeseNum"] == m) & (df["GiornoNum"] <= d))
        df = df.loc[mask].copy()
        cutoff_text = f"Periodo confrontato: 1 Gennaio – {cutoff.strftime('%d %B %Y')} (incluso)"

    # KPI per anno
    kpiy = df.groupby("Anno").agg(
        Fatturato_Netto=("PrezzoTot_Netto","sum"),
        Num_Vendite=("Anno","count"),
        Prezzo_Medio_Articolo=("Prezzo_Netto","mean"),
    ).reset_index()
    monthly = df.groupby([df["DataTurno"].dt.to_period("M"),"Anno"])["PrezzoTot_Netto"].sum().reset_index()
    avg_m = monthly.groupby("Anno")["PrezzoTot_Netto"].mean().reset_index().rename(columns={"PrezzoTot_Netto":"Fatturato_Medio_Mensile"})
    kpiy = kpiy.merge(avg_m, on="Anno", how="left")
    kpiy["Margine_Stimato"] = kpiy["Fatturato_Netto"]*(0.4/1.4)

    # Rev mensile
    piv_rev = df.pivot_table(index="MeseNum", columns="Anno", values="PrezzoTot_Netto", aggfunc="sum").fillna(0.0)
    piv_rev.insert(0, "Mese", [datetime(2000, m, 1).strftime("%b") for m in piv_rev.index])
    piv_rev = piv_rev.reset_index(drop=True)

    # Prod top 10
    if "Produttore_Descrizione" in df.columns:
        top_prod = df.groupby("Produttore_Descrizione")["PrezzoTot_Netto"].sum().sort_values(ascending=False).head(10).index
        prod = df[df["Produttore_Descrizione"].isin(top_prod)].pivot_table(index="Produttore_Descrizione", columns="Anno", values="PrezzoTot_Netto", aggfunc="sum").fillna(0.0)
        prod["Totale"] = prod.sum(axis=1)
        prod = prod.sort_values("Totale", ascending=False).reset_index()
    else:
        prod = pd.DataFrame(columns=["Produttore_Descrizione"])

    # Tip top 8
    if "TipologiaVino_Descrizione" in df.columns:
        top_tip = df.groupby("TipologiaVino_Descrizione")["PrezzoTot_Netto"].sum().sort_values(ascending=False).head(8).index
        tip = df[df["TipologiaVino_Descrizione"].isin(top_tip)].pivot_table(index="TipologiaVino_Descrizione", columns="Anno", values="PrezzoTot_Netto", aggfunc="sum").fillna(0.0)
        tip["Totale"] = tip.sum(axis=1)
        tip = tip.sort_values("Totale", ascending=False).reset_index()
    else:
        tip = pd.DataFrame(columns=["TipologiaVino_Descrizione"])

    # Qty mensile + per anno
    piv_qty = df.pivot_table(index="MeseNum", columns="Anno", values="Qta", aggfunc="sum").fillna(0.0)
    piv_qty.insert(0, "Mese", [datetime(2000, m, 1).strftime("%b") for m in piv_qty.index])
    piv_qty = piv_qty.reset_index(drop=True)

    qty_year = df.groupby("Anno")["Qta"].sum().reset_index().rename(columns={"Qta":"Bottiglie_Totali"})

    return {
        "mode": "row",
        "kpi": kpiy, "rev": piv_rev, "prod": prod, "tip": tip, "qty": piv_qty, "qty_year": qty_year,
        "cutoff_text": cutoff_text
    }

# -------------------- UI --------------------
st.title("🍷 Dashboard Vendite — Dettaglio & Ingrosso")

with st.sidebar:
    st.header("Sorgenti dati")
    st.markdown("**Link pubblici Dropbox (pagina):**")
    st.markdown(f"- Dettaglio: [link pubblico]({PUBLIC_DROPBOX_URL_DETTAGLIO})")
    st.markdown(f"- Ingrosso: [link pubblico]({PUBLIC_DROPBOX_URL_INGROSSO})")
    st.caption("L'app usa automaticamente i link diretti (dl=1) per leggere i dati.")
    dett_link = st.text_input("URL Excel Dettaglio (raw o dl=1)", to_direct_url(PUBLIC_DROPBOX_URL_DETTAGLIO))
    ingr_link = st.text_input("URL Excel Ingrosso (raw o dl=1)", to_direct_url(PUBLIC_DROPBOX_URL_INGROSSO))
    canale = st.radio("Canale", ["Dettaglio","Ingrosso"], horizontal=True)

url = dett_link if canale=="Dettaglio" else ingr_link

try:
    sheets = load_excel(url)
    # Tentativo 1: riga-level
    try:
        parsed = try_row_level(sheets, canale)
    except Exception:
        # Tentativo 2: tabelle 'Dashboard'
        parsed = parse_dashboard_tables(sheets)
except Exception as e:
    st.error("Errore nel caricamento o parsing: " + str(e))
    st.stop()

cutoff_text = parsed.get("cutoff_text","")
if cutoff_text:
    st.markdown(f"**{cutoff_text}**")

# KPI per Anno
kpiy = parsed["kpi"]
c1, c2 = st.columns([3,2])
with c1:
    st.subheader("KPI per Anno")
    if not kpiy.empty:
        format_map = {c:"€{:,.2f}" for c in ["Fatturato_Netto","Prezzo_Medio_Articolo","Fatturato_Medio_Mensile","Margine_Stimato"] if c in kpiy.columns}
        st.dataframe(kpiy.style.format(format_map), use_container_width=True)
    else:
        st.info("Nessun KPI disponibile.")
with c2:
    anni = kpiy["Anno"].tolist() if "Anno" in kpiy.columns and not kpiy.empty else []
    anno_sel = st.selectbox("Anno per dettaglio KPI", anni, index=len(anni)-1 if anni else 0)
    if anni:
        row = kpiy.loc[kpiy["Anno"]==anno_sel].iloc[0].to_dict()
        def fmt_eur(x): 
            try: return f"€ {float(x):,.2f}"
            except: return str(x)
        st.metric("Fatturato netto", fmt_eur(row.get("Fatturato_Netto",0)))
        st.metric("N° vendite", int(row.get("Num_Vendite",0)) if pd.notna(row.get("Num_Vendite",np.nan)) else "-")
        st.metric("Prezzo medio articolo", fmt_eur(row.get("Prezzo_Medio_Articolo",0)))
        st.metric("Fatturato medio mensile", fmt_eur(row.get("Fatturato_Medio_Mensile",0)))
        st.metric("Margine stimato (40%)", fmt_eur(row.get("Margine_Stimato",0)))

st.divider()

# Fatturato mensile YoY
piv_rev = parsed["rev"]
if not piv_rev.empty:
    st.subheader("Fatturato mensile (€/Anno)")
    anni_cols = [c for c in piv_rev.columns if c not in ("Mese",)]
    fig = go.Figure()
    for a in anni_cols:
        fig.add_trace(go.Scatter(x=piv_rev["Mese"], y=piv_rev[a], mode="lines+markers", name=str(a)))
    fig.update_layout(legend_title_text="Anno", xaxis_title="Mese", yaxis_title="€", height=420)
    st.plotly_chart(fig, use_container_width=True)

st.divider()

# Top Produttori
prod = parsed["prod"]
if not prod.empty and "Produttore_Descrizione" in prod.columns:
    st.subheader("Top 10 Produttori — € per Anno (barre raggruppate)")
    melt = prod.melt(id_vars="Produttore_Descrizione", var_name="Anno", value_name="Valore")
    melt = melt[melt["Anno"]!="Totale"]
    figp = px.bar(melt, x="Produttore_Descrizione", y="Valore", color="Anno", barmode="group")
    figp.update_layout(xaxis_title="Produttore", yaxis_title="€", height=480)
    st.plotly_chart(figp, use_container_width=True)

st.divider()

# Top Tipologie
tip = parsed["tip"]
if not tip.empty and "TipologiaVino_Descrizione" in tip.columns:
    st.subheader("Top 8 Tipologie — € per Anno (barre raggruppate)")
    meltt = tip.melt(id_vars="TipologiaVino_Descrizione", var_name="Anno", value_name="Valore")
    meltt = meltt[meltt["Anno"]!="Totale"]
    figt = px.bar(meltt, x="TipologiaVino_Descrizione", y="Valore", color="Anno", barmode="group")
    figt.update_layout(xaxis_title="Tipologia", yaxis_title="€", height=480)
    st.plotly_chart(figt, use_container_width=True)

st.divider()

# Volumi mensili YoY
piv_qty = parsed["qty"]
if not piv_qty.empty:
    st.subheader("Volumi mensili (bottiglie/Anno)")
    anni_cols = [c for c in piv_qty.columns if c not in ("Mese",)]
    figq = go.Figure()
    for a in anni_cols:
        figq.add_trace(go.Scatter(x=piv_qty["Mese"], y=piv_qty[a], mode="lines+markers", name=str(a)))
    figq.update_layout(legend_title_text="Anno", xaxis_title="Mese", yaxis_title="Bottiglie", height=420)
    st.plotly_chart(figq, use_container_width=True)

# Volumi totali per anno
qty_year = parsed["qty_year"]
if not qty_year.empty:
    st.subheader("Volumi totali per Anno (bottiglie)")
    st.dataframe(qty_year, use_container_width=True)

st.caption("Tutti i valori economici sono IVA esclusa. Per l'Ingrosso sono escluse le vendite dei produttori 'Cantine Garrone' e 'Cantine Isidoro'.")
