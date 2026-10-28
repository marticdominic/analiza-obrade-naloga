import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(
    page_title="Analiza obrade naloga - Komercijala",
    page_icon="📦",
    layout="wide"
)

st.title("📊 Analiza obrade naloga komercijala")
st.markdown("Učitaj svoju Excel datoteku i prati vrijeme obrade naloga od početnih statusa do odobrene isporuke.")

# Učitavanje datoteke preko sučelja
uploaded_file = st.file_uploader("Učitaj Excel datoteku (Analiza obrade naloga komercijala.xlsx)", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("👆 Molimo učitajte Excel datoteku kako bi se prikazala analiza.")
    st.stop()

@st.cache_data
def load_data(file):
    df = pd.read_excel(file, sheet_name="Analiza_naloga_po_statusima")
    
    # Pretvorba datumskih polja
    date_cols = [
        'Datum', 'Novo', 'Webshop - Novo', 'Poslana ponuda', 
        'Rezervacija', 'Prioritetna rezervacija', 
        'Čeka potpunu raspoloživost', 'Odobrena isporuka', 
        'Backorder – snimljena djelomična isporuka'
    ]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
            
    return df

try:
    df = load_data(uploaded_file)
except Exception as e:
    st.error(f"Greška pri čitanju Excel datoteke: {e}")
    st.stop()

# Funkcija za računanje radnih dana (ponedjeljak - petak)
def calc_business_days(start_date, end_date):
    if pd.isna(start_date) or pd.isna(end_date):
        return np.nan
    bdays = np.busday_count(start_date.date(), end_date.date())
    return int(bdays)

# ---- PRIPREMA PODATAKA I IZRAČUNI ----
def get_start_status_date(row):
    if pd.notna(row['Novo']):
        return row['Novo'], 'Novo'
    elif pd.notna(row['Backorder – snimljena djelomična isporuka']):
        return row['Backorder – snimljena djelomična isporuka'], 'Backorder'
    elif pd.notna(row['Webshop - Novo']):
        return row['Webshop - Novo'], 'Webshop - Novo'
    else:
        return pd.NaT, 'Nepoznato'

start_dates_info = df.apply(get_start_status_date, axis=1)
df['Početni_Datum'] = [x[0] for x in start_dates_info]
df['Početni_Status_Tip'] = [x[1] for x in start_dates_info]

# Izračuni dana prema pravilima
df['Dana_Do_Odobrenja_Ukupno'] = df.apply(
    lambda row: calc_business_days(row['Početni_Datum'], row['Odobrena isporuka']), axis=1
)

df['Dana_Rezervacija_Do_Odobrenja'] = df.apply(
    lambda row: calc_business_days(row['Rezervacija'], row['Odobrena isporuka']), axis=1
)

df['Dana_Prioritetna_Do_Odobrenja'] = df.apply(
    lambda row: calc_business_days(row['Prioritetna rezervacija'], row['Odobrena isporuka']), axis=1
)

df['Dana_CekaRaspolozivost_Do_Odobrenja'] = df.apply(
    lambda row: calc_business_days(row['Čeka potpunu raspoloživost'], row['Odobrena isporuka']), axis=1
)

# ---- KATEGORIZACIJA PO VREMENSKIM RASPONIMA ----
def categorize_days(days):
    if pd.isna(days):
        return 'Nema podataka'
    elif days == 1:
        return '1 dan'
    elif days == 2:
        return '2 dana'
    elif 2 < days <= 5:
        return '2 - 5 dana'
    elif 5 < days <= 10:
        return '5 - 10 dana'
    else:
        return 'Više od 10 dana'

df['Kategorija_Dana'] = df['Dana_Do_Odobrenja_Ukupno'].apply(categorize_days)

# ---- NOVI IZVJEŠTAJ: OD (NOVO / WEBSHOP NOVO / BACKORDER) DO ODOBRENJA ----
def get_earliest_start_extended(row):
    dates = []
    if pd.notna(row['Novo']):
        dates.append(row['Novo'])
    if pd.notna(row['Webshop - Novo']):
        dates.append(row['Webshop - Novo'])
    if pd.notna(row['Backorder – snimljena djelomična isporuka']):
        dates.append(row['Backorder – snimljena djelomična isporuka'])
    return min(dates) if dates else pd.NaT

df['Earliest_Start_Ext'] = df.apply(get_earliest_start_extended, axis=1)
df['Dana_Ext_Do_Odobrenja'] = df.apply(
    lambda row: calc_business_days(row['Earliest_Start_Ext'], row['Odobrena isporuka']), axis=1
)
df['Kategorija_Dana_Ext'] = df['Dana_Ext_Do_Odobrenja'].apply(categorize_days)

# ---- NOVI IZVJEŠTAJ: PUTANJA NALOGA (DIREKTNO VS MEĐUSTATUSI) ----
def classify_order_path(row):
    if pd.isna(row['Novo']) or pd.isna(row['Odobrena isporuka']):
        return 'Nije primjenjivo'
    
    has_rez = pd.notna(row['Rezervacija'])
    has_prio = pd.notna(row['Prioritetna rezervacija'])
    has_ceka = pd.notna(row['Čeka potpunu raspoloživost'])
    
    if not has_rez and not has_prio and not has_ceka:
        return 'Direktno (Novo -> Odobreno)'
    else:
        return 'Preko međustatusa (Rezervacija / Čeka raspoloživost)'

df['Putanja_Naloga'] = df.apply(classify_order_path, axis=1)

# ---- SIDEBAR FILTRI ----
st.sidebar.header("🔍 Filtri")

skladista = ['Sva skladišta'] + sorted(df['Skladište'].dropna().unique().tolist())
selected_skladiste = st.sidebar.selectbox("Skladište", skladista)

if selected_skladiste != 'Sva skladišta':
    df_filtered = df[df['Skladište'] == selected_skladiste]
else:
    df_filtered = df.copy()

# ---- GLAVNI PRIKAZ - OPERATIVNI KPI KARTICE ----
col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    avg_total = df_filtered['Dana_Do_Odobrenja_Ukupno'].mean()
    st.metric("Prosjek (Početak -> Odobrenje)", f"{avg_total:.1f} rad. dana" if pd.notna(avg_total) else "N/A")

with col2:
    median_total = df_filtered['Dana_Do_Odobrenja_Ukupno'].median()
    st.metric("Medijan (Srednja vrijednost)", f"{int(median_total)} rad. dana" if pd.notna(median_total) else "N/A")

with col3:
    valid_days_count = df_filtered['Dana_Do_Odobrenja_Ukupno'].dropna()
    pct_within_2 = (valid_days_count <= 2).mean() * 100 if len(valid_days_count) > 0 else 0
    st.metric("Riješeno unutar 2 dana", f"{pct_within_2:.1f}%")

with col4:
    avg_rez = df_filtered['Dana_Rezervacija_Do_Odobrenja'].mean()
    st.metric("Prosjek (Rezervacija)", f"{avg_rez:.1f} rad. dana" if pd.notna(avg_rez) else "N/A")

with col5:
    avg_ceka = df_filtered['Dana_CekaRaspolozivost_Do_Odobrenja'].mean()
    st.metric("Prosjek (Čeka raspol.)", f"{avg_ceka:.1f} rad. dana" if pd.notna(avg_ceka) else "N/A")

st.markdown("---")

# ---- GRAFOVI I ANALIZA ----
col_a, col_b = st.columns(2)

with col_a:
    st.subheader("📌 Distribucija početnih statusa")
    status_counts = df_filtered['Početni_Status_Tip'].value_counts().reset_index()
    status_counts.columns = ['Status', 'Broj naloga']
    fig_pie = px.pie(status_counts, names='Status', values='Broj naloga', hole=0.4)
    st.plotly_chart(fig_pie, use_container_width=True)

with col_b:
    st.subheader("📈 Prosječni broj dana po statusima")
    metrics_avg = pd.DataFrame({
        'Faza': ['Početak -> Odobrenje', 'Rezervacija -> Odobrenje', 'Prioritetna -> Odobrenje', 'Čeka raspoloživost -> Odobrenje'],
        'Prosjek radnih dana': [
            df_filtered['Dana_Do_Odobrenja_Ukupno'].mean(),
            df_filtered['Dana_Rezervacija_Do_Odobrenja'].mean(),
            df_filtered['Dana_Prioritetna_Do_Odobrenja'].mean(),
            df_filtered['Dana_CekaRaspolozivost_Do_Odobrenja'].mean()
        ]
    })
    fig_bar = px.bar(metrics_avg, x='Faza', y='Prosjek radnih dana', text_auto='.1f', color='Faza')
    st.plotly_chart(fig_bar, use_container_width=True)

# ---- NOVI DODATAK: IZVJEŠTAJ O PUTANJI NALOGA (DIREKTNO VS MEĐUSTATUSI) ----
st.subheader("🔀 Izvještaj: Udio direktnih naloga (Novo -> Odobreno) vs. Međustatusi")

path_subset = df_filtered[df_filtered['Putanja_Naloga'] != 'Nije primjenjivo']
path_counts = path_subset['Putanja_Naloga'].value_counts().reset_index()
path_counts.columns = ['Putanja', 'Broj naloga']
total_path = len(path_subset)
path_counts['Postotak (%)'] = (path_counts['Broj naloga'] / total_path * 100).round(2) if total_path > 0 else 0

col_p1, col_p2 = st.columns([1, 2])

with col_p1:
    st.dataframe(path_counts, use_container_width=True, hide_index=True)

with col_p2:
    fig_path = px.pie(
        path_counts, 
        names='Putanja', 
        values='Broj naloga', 
        hole=0.4,
        title="Udio naloga po putanji obrade"
    )
    st.plotly_chart(fig_path, use_container_width=True)

# ---- IZVJEŠTAJ S POSTOCIMA: OD POČETNOG STATUSA DO ODOBRENE ISPORUKE ----
st.subheader("⏱️ Izvještaj: Vrijeme obrade od početnog statusa do odobrene isporuke (s udjelima)")

order_categories = ['1 dan', '2 dana', '2 - 5 dana', '5 - 10 dana', 'Više od 10 dana', 'Nema podataka']

cat_counts = df_filtered['Kategorija_Dana'].value_counts().reindex(order_categories).fillna(0).reset_index()
cat_counts.columns = ['Vremenski raspon', 'Broj naloga']
total_naloga = len(df_filtered)
cat_counts['Postotak (%)'] = (cat_counts['Broj naloga'] / total_naloga * 100).round(2) if total_naloga > 0 else 0

col_c1, col_c2 = st.columns([1, 2])

with col_c1:
    st.dataframe(cat_counts, use_container_width=True, hide_index=True)

with col_c2:
    fig_hist = px.bar(
        cat_counts[cat_counts['Vremenski raspon'] != 'Nema podataka'], 
        x='Vremenski raspon', 
        y='Broj naloga', 
        text=cat_counts[cat_counts['Vremenski raspon'] != 'Nema podataka']['Postotak (%)'].apply(lambda x: f"{x}%"),
        color='Vremenski raspon',
        title="Broj naloga i udio (%) po rasponu radnih dana do odobrenja"
    )
    fig_hist.update_traces(textposition='outside')
    st.plotly_chart(fig_hist, use_container_width=True)

# ---- NOVI IZVJEŠTAJ: OD (NOVO / WEBSHOP NOVO / BACKORDER) DO ODOBRENJA ----
st.subheader("⏱️ Izvještaj: Vrijeme obrade od (Novo / Webshop Novo / Backorder) do Odobrenja")

ext_counts = df_filtered['Kategorija_Dana_Ext'].value_counts().reindex(order_categories).fillna(0).reset_index()
ext_counts.columns = ['Vremenski raspon', 'Broj naloga']
total_ext = len(df_filtered)
ext_counts['Postotak (%)'] = (ext_counts['Broj naloga'] / total_ext * 100).round(2) if total_ext > 0 else 0

col_e1, col_e2 = st.columns([1, 2])

with col_e1:
    st.dataframe(ext_counts, use_container_width=True, hide_index=True)

with col_e2:
    fig_ext = px.bar(
        ext_counts[ext_counts['Vremenski raspon'] != 'Nema podataka'], 
        x='Vremenski raspon', 
        y='Broj naloga', 
        text=ext_counts[ext_counts['Vremenski raspon'] != 'Nema podataka']['Postotak (%)'].apply(lambda x: f"{x}%"),
        color='Vremenski raspon',
        title="Udio (%) po rasponu (Novo / Webshop Novo / Backorder -> Odobrenje)"
    )
    fig_ext.update_traces(textposition='outside')
    st.plotly_chart(fig_ext, use_container_width=True)

# ---- TABLICA S REZULTATIMA ----
st.subheader("📋 Detaljni pregled naloga i izračunatih dana")
st.dataframe(df_filtered[[
    'Broj naloga', 'Naziv partnera', 'Skladište', 'Početni_Status_Tip', 
    'Putanja_Naloga', 'Dana_Do_Odobrenja_Ukupno', 'Kategorija_Dana', 
    'Dana_Ext_Do_Odobrenja', 'Kategorija_Dana_Ext',
    'Dana_Rezervacija_Do_Odobrenja', 'Dana_Prioritetna_Do_Odobrenja', 
    'Dana_CekaRaspolozivost_Do_Odobrenja'
]], use_container_width=True)
