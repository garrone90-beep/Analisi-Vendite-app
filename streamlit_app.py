# streamlit_app.py
# -------------------------------------------------------------
# Dashboard web (Streamlit) — con link pubblici Dropbox ai file originali
# -------------------------------------------------------------

import io
import requests
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from datetime import datetime

st.set_page_config(page_title="Dashboard Vendite", page_icon="🍷", layout="wide")

# --- Link pubblici Dropbox (landing page) ---
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

# Sorgenti per il caricamento: secrets o input sidebar
URL_DETTAGLIO = st.secrets.get("URL_DETTAGLIO", to_direct_url(PUBLIC_DROPBOX_URL_DETTAGLIO))
URL_INGROSSO  = st.secrets.get("URL_INGROSSO",  to_direct_url(PUBLIC_DROPBOX_URL_INGROSSO))

@st.cache_data(show_spinner=True, ttl=600)
def load_excel(url: str) -> dict:
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    bio = io.BytesIO(r.content)
    xl = pd.ExcelFile(bio)
    return {name: xl.parse(name) for name in xl.sheet_names}

def normalize_rows(sheets: dict, canale: str) -> pd.DataFrame:
    # cerca 'Dati' oppure un foglio con colonne chiave
    df = None
    for name, tab in sheets.items():
        if name.strip().lower() == "dati":
            df = tab.copy(); break
    if df is None:
        for name, tab in sheets.items():
            cols = {str(c).strip().lower() for c in tab.columns}
            if {"dataturno","prezzo","prezzotot","qta"}.issubset(cols):
                df = tab.copy(); break
    if df is None or df.empty:
        raise ValueError("Foglio dati riga-level non trovato (es. 'Dati').")

    # rename robusto
    rename_map = {} 
    for col in df.columns:
        c = str(col).strip().lower()
        if c == "ordine": rename_map[col] = "Ordine"
        if c.startswith("produttore"): rename_map[col] = "Produttore_Descrizione"
        if "tipologiavino" in c or c.startswith("tipologia"): rename_map[col] = "TipologiaVino_Descrizione"
        if c in ("qta","quantita","quantità"): rename_map[col] = "Qta"
        if c == "prezzo": rename_map[col] = "Prezzo"
        if c in ("prezzotot","prezzo_tot","prezzo_totale"): rename_map[col] = "PrezzoTot"
        if c == "dataturno": rename_map[col] = "DataTurno"
        if c == "dettaglioingrosso": rename_map[col] = "DettaglioIngrosso"
        if c == "denominazione": rename_map[col] = "Denominazione"
    df = df.rename(columns=rename_map)
    if "DataTurno" not in df.columns:
        raise ValueError("Colonna 'DataTurno' mancante.")
    df["DataTurno"] = pd.to_datetime(df["DataTurno"], errors="coerce")

    # escludi alimentari
    if "TipologiaVino_Descrizione" in df.columns:
        tip = df["TipologiaVino_Descrizione"].astype(str).str.upper()
        mask_exclude = tip.str.contains("PRODOTTO|ALIMENT|FOOD|GASTRON|SNACK|CONSERVE|DOLCI")
        df = df.loc[~mask_exclude].copy()

    # canale + IVA
    if canale == "Dettaglio":
        df = df[df.get("DettaglioIngrosso",0).fillna(0).astype(int) == 0]
        df["Prezzo_Netto"] = df["Prezzo"]/1.22
        df["PrezzoTot_Netto"] = df["PrezzoTot"]/1.22
    else:
        df = df[df.get("DettaglioIngrosso",1).fillna(1).astype(int) == 1]
        if "Produttore_Descrizione" in df.columns:
            df = df[~df["Produttore_Descrizione"].isin({"Cantine Garrone","Cantine Isidoro"})]
        df["Prezzo_Netto"] = df["Prezzo"]
        df["PrezzoTot_Netto"] = df["PrezzoTot"]

    if "Qta" not in df.columns:
        df["Qta"] = 1

    df["Anno"] = df["DataTurno"].dt.year
    df["MeseNum"] = df["DataTurno"].dt.month
    df["GiornoNum"] = df["DataTurno"].dt.day
    df["Mese"] = df["DataTurno"].dt.to_period("M").astype(str)
    return df

def apply_cutoff(df: pd.DataFrame):
    cutoff = df["DataTurno"].max()
    if pd.isna(cutoff):
        return df, cutoff
    m, d = cutoff.month, cutoff.day
    mask = (df["MeseNum"] < m) | ((df["MeseNum"] == m) & (df["GiornoNum"] <= d))
    return df.loc[mask].copy(), cutoff

def kpi_by_year(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Anno","Fatturato_Netto","Num_Vendite","Prezzo_Medio_Articolo","Fatturato_Medio_Mensile","Margine_Stimato"])
    out = df.groupby("Anno").agg(
        Fatturato_Netto=("PrezzoTot_Netto","sum"),
        Num_Vendite=("Anno","count"),
        Prezzo_Medio_Articolo=("Prezzo_Netto","mean"),
    ).reset_index()
    monthly = df.groupby([df["DataTurno"].dt.to_period("M"),"Anno"])["PrezzoTot_Netto"].sum().reset_index()
    avg_m = monthly.groupby("Anno")["PrezzoTot_Netto"].mean().reset_index().rename(columns={"PrezzoTot_Netto":"Fatturato_Medio_Mensile"})
    out = out.merge(avg_m, on="Anno", how="left")
    out["Margine_Stimato"] = out["Fatturato_Netto"]*(0.4/1.4)
    return out.sort_values("Anno")

def yoy_monthly(df: pd.DataFrame, value_col: str) -> pd.DataFrame:
    if df.empty: return pd.DataFrame()
    piv = df.pivot_table(index="MeseNum", columns="Anno", values=value_col, aggfunc="sum").fillna(0.0)
    piv.insert(0,"Mese",[datetime(2000,m,1).strftime("%b") for m in piv.index])
    return piv.reset_index(drop=True)

def topN_yoy(df: pd.DataFrame, index_col: str, value_col: str, n: int) -> pd.DataFrame:
    if df.empty or index_col not in df.columns: return pd.DataFrame(columns=[index_col])
    top_index = df.groupby(index_col)[value_col].sum().sort_values(ascending=False).head(n).index
    base = df[df[index_col].isin(top_index)]
    piv = base.pivot_table(index=index_col, columns="Anno", values=value_col, aggfunc="sum").fillna(0.0)
    piv["Totale"] = piv.sum(axis=1)
    return piv.sort_values("Totale", ascending=False).reset_index()

# ---------------- UI ----------------
st.title("🍷 Dashboard Vendite — Dettaglio & Ingrosso")

with st.sidebar:
    st.header("Sorgenti dati")
    st.markdown("**File pubblici Dropbox:**")
    st.markdown(f"- Dettaglio: [link pubblico]({{PUBLIC_DROPBOX_URL_DETTAGLIO}})")
    st.markdown(f"- Ingrosso: [link pubblico]({{PUBLIC_DROPBOX_URL_INGROSSO}})")
    st.caption("I link sopra aprono la pagina pubblica Dropbox (dl=0). L'app usa i link diretti (dl=1) per leggere i dati.")
    dett_link = st.text_input("URL Excel Dettaglio (raw)", to_direct_url(PUBLIC_DROPBOX_URL_DETTAGLIO))
    ingr_link = st.text_input("URL Excel Ingrosso (raw)", to_direct_url(PUBLIC_DROPBOX_URL_INGROSSO))
    canale = st.radio("Canale", ["Dettaglio","Ingrosso"], horizontal=True)

url = dett_link if canale=="Dettaglio" else ingr_link

try:
    sheets = load_excel(url)
    rows_raw = normalize_rows(sheets, canale)
    rows, cutoff = apply_cutoff(rows_raw)
except Exception as e:
    st.error(f"Errore nel caricamento o parsing: {{e}}")
    st.stop()

periodo_txt = f"Periodo confrontato: 1 Gennaio – {{cutoff.strftime('%d %B %Y')}} (incluso)" if isinstance(cutoff, pd.Timestamp) else "Periodo non disponibile"
st.markdown(f"**{{periodo_txt}}**")
if canale == "Ingrosso":
    st.caption("Nota: esclusi i produttori 'Cantine Garrone' e 'Cantine Isidoro'.")

# KPI
kpiy = kpi_by_year(rows)
c1, c2 = st.columns([3,2])
with c1:
    st.subheader("KPI per Anno (fino al cutoff)")
    st.dataframe(kpiy.style.format({
        "Fatturato_Netto": "€{:,.2f}",
        "Prezzo_Medio_Articolo": "€{:,.2f}",
        "Fatturato_Medio_Mensile": "€{:,.2f}",
        "Margine_Stimato": "€{:,.2f}",
    }), use_container_width=True)
with c2:
    anni = kpiy["Anno"].tolist() if not kpiy.empty else []
    anno_sel = st.selectbox("Anno per dettaglio KPI", anni, index=len(anni)-1 if anni else 0)
    if anni:
        r = kpiy.loc[kpiy["Anno"]==anno_sel].iloc[0]
        st.metric("Fatturato netto", f"€ {r['Fatturato_Netto']:,.2f}")
        st.metric("N° vendite", int(r['Num_Vendite']))
        st.metric("Prezzo medio articolo", f"€ {r['Prezzo_Medio_Articolo']:,.2f}")
        st.metric("Fatturato medio mensile", f"€ {r['Fatturato_Medio_Mensile']:,.2f}")
        st.metric("Margine stimato (40%)", f"€ {r['Margine_Stimato']:,.2f}")

st.divider()

# Fatturato mensile YoY
piv_rev = yoy_monthly(rows, "PrezzoTot_Netto")
if not piv_rev.empty:
    st.subheader("Fatturato mensile (€/Anno) — fino al cutoff")
    anni_cols = [c for c in piv_rev.columns if c not in ("Mese",)]
    fig = go.Figure()
    for a in anni_cols:
        fig.add_trace(go.Scatter(x=piv_rev["Mese"], y=piv_rev[a], mode="lines+markers", name=str(a)))
    fig.update_layout(legend_title_text="Anno", xaxis_title="Mese", yaxis_title="€", height=420)
    st.plotly_chart(fig, use_container_width=True)

st.divider()

# Top Produttori
if "Produttore_Descrizione" in rows.columns:
    prod_yoy = topN_yoy(rows, "Produttore_Descrizione", "PrezzoTot_Netto", 10)
    if not prod_yoy.empty:
        st.subheader("Top 10 Produttori — € per Anno (barre raggruppate)")
        data_m = prod_yoy.melt(id_vars="Produttore_Descrizione", var_name="Anno", value_name="Valore")
        data_m = data_m[data_m["Anno"]!="Totale"]
        figp = px.bar(data_m, x="Produttore_Descrizione", y="Valore", color="Anno", barmode="group")
        figp.update_layout(xaxis_title="Produttore", yaxis_title="€", height=480)
        st.plotly_chart(figp, use_container_width=True)

st.divider()

# Top Tipologie
if "TipologiaVino_Descrizione" in rows.columns:
    tip_yoy = topN_yoy(rows, "TipologiaVino_Descrizione", "PrezzoTot_Netto", 8)
    if not tip_yoy.empty:
        st.subheader("Top 8 Tipologie — € per Anno (barre raggruppate)")
        data_t = tip_yoy.melt(id_vars="TipologiaVino_Descrizione", var_name="Anno", value_name="Valore")
        data_t = data_t[data_t["Anno"]!="Totale"]
        figt = px.bar(data_t, x="TipologiaVino_Descrizione", y="Valore", color="Anno", barmode="group")
        figt.update_layout(xaxis_title="Tipologia", yaxis_title="€", height=480)
        st.plotly_chart(figt, use_container_width=True)

st.divider()

# Volumi mensili YoY
piv_qty = yoy_monthly(rows, "Qta")
if not piv_qty.empty:
    st.subheader("Volumi mensili (bottiglie/Anno) — fino al cutoff")
    anni_cols = [c for c in piv_qty.columns if c not in ("Mese",)]
    figq = go.Figure()
    for a in anni_cols:
        figq.add_trace(go.Scatter(x=piv_qty["Mese"], y=piv_qty[a], mode="lines+markers", name=str(a)))
    figq.update_layout(legend_title_text="Anno", xaxis_title="Mese", yaxis_title="Bottiglie", height=420)
    st.plotly_chart(figq, use_container_width=True)

# Volumi totali per anno
if not rows.empty:
    st.subheader("Volumi totali per Anno (bottiglie) — fino al cutoff")
    qty_year = rows.groupby("Anno")["Qta"].sum().reset_index().rename(columns={"Qta":"Bottiglie_Totali"})
    st.dataframe(qty_year, use_container_width=True)

st.caption("Tutti i valori economici sono IVA esclusa. Per l'Ingrosso sono escluse le vendite dei produttori 'Cantine Garrone' e 'Cantine Isidoro'.")
