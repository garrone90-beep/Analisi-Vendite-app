# streamlit_app.py
import io, requests, numpy as np, pandas as pd, plotly.express as px, streamlit as st
from datetime import datetime

st.set_page_config(page_title="Dashboard Vendite Enoteca", page_icon="🍷", layout="wide")

# ——— Link pubblici Dropbox (pagina di anteprima)
PUBLIC_DROPBOX_URL_DETTAGLIO_PARIPER = "https://www.dropbox.com/scl/fi/ajoax5i3gthl1kg3ewe3t/AnalisiVendite_Dettaglio_Dashboard_v4.xlsx?rlkey=nfj3kg0s3n6yp206honxfhr1j&dl=0"
PUBLIC_DROPBOX_URL_DETTAGLIO_FULL   = "https://www.dropbox.com/scl/fi/a1fdmmos6zbd835w9iyx9/AnalisiVendite_Dettaglio_Dashboard_v3.xlsx?rlkey=9eht1v11lx0cpx2tzzarvxrsh&dl=0"
PUBLIC_DROPBOX_URL_INGROSSO_PARIPER = "https://www.dropbox.com/scl/fi/608xyutty9zyufb57ddde/AnalisiVendite_Ingrosso_Dashboard_v5.xlsx?rlkey=0c8haj4csbyuf75o3yeuc2c82&dl=0"
PUBLIC_DROPBOX_URL_INGROSSO_FULL    = "https://www.dropbox.com/scl/fi/ntd2m5ife5yz0gfosh2a0/AnalisiVendite_Ingrosso_Dashboard_v3.xlsx?rlkey=tlpvseq07dh838sqt57caeq4p&dl=0"

def to_direct_url(u: str) -> str:
    if not u: return u
    if "dropbox.com" in u and "dl=0" in u: return u.replace("dl=0","dl=1")
    if "github.com" in u and "/blob/" in u: return u.replace("github.com","raw.githubusercontent.com").replace("/blob/","/")
    return u

URL_DETTAGLIO_PARIPER = to_direct_url(PUBLIC_DROPBOX_URL_DETTAGLIO_PARIPER)
URL_DETTAGLIO_FULL    = to_direct_url(PUBLIC_DROPBOX_URL_DETTAGLIO_FULL)
URL_INGROSSO_PARIPER  = to_direct_url(PUBLIC_DROPBOX_URL_INGROSSO_PARIPER)
URL_INGROSSO_FULL     = to_direct_url(PUBLIC_DROPBOX_URL_INGROSSO_FULL)

@st.cache_data(show_spinner=True, ttl=600)
def load_excel(url: str) -> dict:
    r = requests.get(url, timeout=60); r.raise_for_status()
    xl = pd.ExcelFile(io.BytesIO(r.content), engine="openpyxl")
    return {name: xl.parse(name, header=None) for name in xl.sheet_names}

# ——— Parser “Dashboard”
def _find_header_row(df, cols):
    for i in range(len(df)):
        row = [str(x).strip() for x in df.iloc[i].tolist()]
        if all(c in row for c in cols):
            return i, {c: row.index(c) for c in cols}
    raise ValueError("Header non trovato")

def _extract_table(df, header_row, start_col, n_cols):
    headers = [str(x).strip() for x in df.iloc[header_row, start_col:start_col+n_cols].tolist()]
    out, r = [], header_row+1
    while r < len(df):
        row = df.iloc[r, start_col:start_col+n_cols].tolist()
        if all((x is None) or (str(x).strip() in ("","None","nan")) for x in row): break
        out.append(row); r += 1
    return pd.DataFrame(out, columns=headers)

def parse_dashboard_tables(sheets: dict)->dict:
    name=[n for n in sheets if n.strip().lower()=="dashboard"]
    if not name: raise ValueError("Foglio 'Dashboard' non trovato")
    df=sheets[name[0]].copy()

    # KPI per Anno
    kcols=["Anno","Fatturato_Netto","Num_Vendite","Prezzo_Medio_Articolo","Fatturato_Medio_Mensile","Margine_Stimato"]
    krow,kpos=_find_header_row(df,kcols); kstart=min(kpos.values())
    kpi=_extract_table(df,krow,kstart,len(kcols))
    for c in kcols[1:]: kpi[c]=pd.to_numeric(kpi[c], errors="coerce")
    kpi=kpi.sort_values("Anno")
    for c in [c for c in kpi.columns if c!="Anno"]:
        kpi[f"{c}_YoY%"]=kpi[c].pct_change()*100

    # Fatturato mensile
    def find_monthly_after(r0):
        for i in range(r0+1,len(df)):
            row=[str(x).strip() for x in df.iloc[i].tolist()]
            if "Mese" in row:
                c0=row.index("Mese"); n=1; j=c0+1
                while j<df.shape[1] and (df.iloc[i,j] is not None) and (str(df.iloc[i,j]).strip() not in ("","None","nan")):
                    n+=1; j+=1
                return i,c0,n
        raise ValueError("Tabella 'Mese + anni' non trovata.")
    rrow,rcol,rn = find_monthly_after(krow)
    rev=_extract_table(df,rrow,rcol,rn); 
    for c in rev.columns:
        if c!="Mese": rev[c]=pd.to_numeric(rev[c], errors="coerce")

    # Top Produttori
    def find_idx_tot_after(r0,key):
        for i in range(r0+1,len(df)):
            row=[str(x).strip() for x in df.iloc[i].tolist()]
            if key in row and "Totale" in row:
                ck=row.index(key); ct=row.index("Totale"); return i,ck,(ct-ck+1)
        raise ValueError("Tabella non trovata")
    prow,pcol,pn = find_idx_tot_after(rrow,"Produttore_Descrizione")
    prod=_extract_table(df,prow,pcol,pn)
    for c in prod.columns:
        if c!="Produttore_Descrizione": prod[c]=pd.to_numeric(prod[c], errors="coerce")

    # Top Tipologie
    trow,tcol,tn = find_idx_tot_after(prow,"TipologiaVino_Descrizione")
    tip=_extract_table(df,trow,tcol,tn)
    for c in tip.columns:
        if c!="TipologiaVino_Descrizione": tip[c]=pd.to_numeric(tip[c], errors="coerce")

    # Volumi mensili
    qrow,qcol,qn = find_monthly_after(trow)
    qty=_extract_table(df,qrow,qcol,qn)
    for c in qty.columns:
        if c!="Mese": qty[c]=pd.to_numeric(qty[c], errors="coerce")

    # Volumi totali per anno
    ycols=["Anno","Bottiglie_Totali"]; yrow,ypos=_find_header_row(df,ycols); ystart=min(ypos.values())
    qty_year=_extract_table(df,yrow,ystart,len(ycols))
    qty_year["Bottiglie_Totali"]=pd.to_numeric(qty_year["Bottiglie_Totali"], errors="coerce")

    cutoff=""
    for i in range(min(6,len(df))):
        row=" ".join([str(x) for x in df.iloc[i].tolist() if pd.notna(x)])
        if "Periodo" in row: cutoff=row.strip(); break

    return {"kpi":kpi,"rev":rev,"prod":prod,"tip":tip,"qty":qty,"qty_year":qty_year,"cutoff_text":cutoff}

# ——— UI: toggle rapidi
left,right=st.columns([1,1])
with left:
    try: canale=st.segmented_control("Canale",["Dettaglio","Ingrosso"],selection="Dettaglio")
    except Exception: canale=st.radio("Canale",["Dettaglio","Ingrosso"],horizontal=True,index=0)
with right:
    try: visual=st.segmented_control("Visualizzazione",["Pari periodo","Anno completo"],selection="Pari periodo")
    except Exception: visual=st.radio("Visualizzazione",["Pari periodo","Anno completo"],horizontal=True,index=0)

# Link (solo info) in sidebar
with st.sidebar:
    st.header("Sorgenti dati (Dropbox)")
    st.markdown("**Link pubblici:**")
    st.markdown(f"- Dettaglio (Pari periodo): [link pubblico]({PUBLIC_DROPBOX_URL_DETTAGLIO_PARIPER})")
    st.markdown(f"- Dettaglio (Anno completo): [link pubblico]({PUBLIC_DROPBOX_URL_DETTAGLIO_FULL})")
    st.markdown(f"- Ingrosso (Pari periodo): [link pubblico]({PUBLIC_DROPBOX_URL_INGROSSO_PARIPER})")
    st.markdown(f"- Ingrosso (Anno completo): [link pubblico]({PUBLIC_DROPBOX_URL_INGROSSO_FULL})")

# Scelta URL
url = URL_DETTAGLIO_PARIPER if (canale=="Dettaglio" and visual=="Pari periodo") else \
      URL_DETTAGLIO_FULL    if (canale=="Dettaglio" and visual=="Anno completo") else \
      URL_INGROSSO_PARIPER  if (canale=="Ingrosso" and visual=="Pari periodo") else \
      URL_INGROSSO_FULL

# Caricamento
try:
    sheets=load_excel(url); parsed=parse_dashboard_tables(sheets)
except Exception as e:
    st.error("Errore nel caricamento o parsing: "+str(e)); st.stop()

# Titolo e badge
st.title("Dashboard Vendite Enoteca")
st.markdown(f"**Vista:** {canale} • **{visual}**")
if parsed.get("cutoff_text",""): st.caption(parsed["cutoff_text"])

# 1) Metriche anno selezionato
kpi=parsed["kpi"]
if not kpi.empty and "Anno" in kpi.columns:
    anni=kpi["Anno"].tolist(); anno_sel=st.select_slider("Anno selezionato", options=anni, value=anni[-1])
    row=kpi.loc[kpi["Anno"]==anno_sel].iloc[0].to_dict()
    def eur(x): 
        try: return f"€ {float(x):,.2f}"
        except: return str(x)
    c1,c2,c3,c4,c5=st.columns(5)
    with c1:
        d=row.get("Fatturato_Netto_YoY%",np.nan); st.metric("Fatturato netto", eur(row.get("Fatturato_Netto",0)), None if np.isnan(d) else f"{d:+.1f}%")
    with c2:
        d=row.get("Num_Vendite_YoY%",np.nan); v=row.get("Num_Vendite",None); st.metric("N° vendite","-" if v is None or pd.isna(v) else int(v), None if np.isnan(d) else f"{d:+.1f}%")
    with c3:
        d=row.get("Prezzo_Medio_Articolo_YoY%",np.nan); st.metric("Prezzo medio articolo", eur(row.get("Prezzo_Medio_Articolo",0)), None if np.isnan(d) else f"{d:+.1f}%")
    with c4:
        d=row.get("Fatturato_Medio_Mensile_YoY%",np.nan); st.metric("Fatturato medio mensile", eur(row.get("Fatturato_Medio_Mensile",0)), None if np.isnan(d) else f"{d:+.1f}%")
    with c5:
        d=row.get("Margine_Stimato_YoY%",np.nan); st.metric("Margine stimato (40%)", eur(row.get("Margine_Stimato",0)), None if np.isnan(d) else f"{d:+.1f}%")

st.divider()

# 2) Tabella KPI
if not kpi.empty:
    st.subheader("KPI per Anno (con YoY%)")
    fmt={c:"€{:,.2f}" for c in ["Fatturato_Netto","Prezzo_Medio_Articolo","Fatturato_Medio_Mensile","Margine_Stimato"] if c in kpi.columns}
    for c in [col for col in kpi.columns if col.endswith("_YoY%")]: fmt[c] = "{:.1f}%"
    st.dataframe(kpi.style.format(fmt), use_container_width=True)

st.divider()

# 3) Grafici a colonne
rev=parsed["rev"]
if not rev.empty:
    st.subheader("Fatturato mensile (€/Anno) — barre raggruppate")
    rev_long=rev.melt(id_vars="Mese", var_name="Anno", value_name="Valore")
    fig=px.bar(rev_long, x="Mese", y="Valore", color="Anno", barmode="group"); fig.update_layout(legend_title_text="Anno", xaxis_title="Mese", yaxis_title="€", height=480)
    st.plotly_chart(fig, use_container_width=True)

st.divider()
prod=parsed["prod"]
if not prod.empty and "Produttore_Descrizione" in prod.columns:
    st.subheader("Top 10 Produttori — € per Anno (barre raggruppate)")
    prod_long=prod.melt(id_vars="Produttore_Descrizione", var_name="Anno", value_name="Valore"); prod_long=prod_long[prod_long["Anno"]!="Totale"]
    figp=px.bar(prod_long, x="Produttore_Descrizione", y="Valore", color="Anno", barmode="group"); figp.update_layout(xaxis_title="Produttore", yaxis_title="€", height=500)
    st.plotly_chart(figp, use_container_width=True)

st.divider()
tip=parsed["tip"]
if not tip.empty and "TipologiaVino_Descrizione" in tip.columns:
    st.subheader("Top 8 Tipologie — € per Anno (barre raggruppate)")
    tip_long=tip.melt(id_vars="TipologiaVino_Descrizione", var_name="Anno", value_name="Valore"); tip_long=tip_long[tip_long["Anno"]!="Totale"]
    figt=px.bar(tip_long, x="TipologiaVino_Descrizione", y="Valore", color="Anno", barmode="group"); figt.update_layout(xaxis_title="Tipologia", yaxis_title="€", height=500)
    st.plotly_chart(figt, use_container_width=True)

st.divider()
qty=parsed["qty"]
if not qty.empty:
    st.subheader("Volumi mensili (bottiglie/Anno) — barre raggruppate")
    qty_long=qty.melt(id_vars="Mese", var_name="Anno", value_name="Bottiglie")
    figq=px.bar(qty_long, x="Mese", y="Bottiglie", color="Anno", barmode="group"); figq.update_layout(legend_title_text="Anno", xaxis_title="Mese", yaxis_title="Bottiglie", height=480)
    st.plotly_chart(figq, use_container_width=True)

qty_year=parsed["qty_year"]
if not qty_year.empty:
    st.subheader("Volumi totali per Anno (bottiglie)")
    st.dataframe(qty_year, use_container_width=True)
