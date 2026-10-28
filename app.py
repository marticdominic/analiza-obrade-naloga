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

# ---- SIDEBAR FILTRI ----
st.sidebar.header("🔍 Filtri")

skladista = ['Sva skladišta'] + sorted(df['Skladište'].dropna().unique().tolist())
selected_skladiste = st.sidebar.selectbox("Skladište", skladista)

if selected_skladiste != 'Sva skladišta':
    df_filtered = df[df['Skladište'] == selected_skladiste]
else:
    df_filtered = df.copy()

# ---- GLAVNI PRIKAZ - KPI KARTICE ----
col1, col2, col3, col4 = st.columns(4)

with col1:
    avg_total = df_filtered['Dana_Do_Odobrenja_Ukupno'].mean()
    st.metric("Prosjek (Početak -> Odobrenje)", f"{avg_total:.1f} radnih dana" if pd.notna(avg_total) else "N/A")

with col2:
    avg_rez = df_filtered['Dana_Rezervacija_Do_Odobrenja'].mean()
    st.metric("Prosjek (Rezervacija -> Odobrenje)", f"{avg_rez:.1f} radnih dana" if pd.notna(avg_rez) else "N/A")

with col3:
    avg_prio = df_filtered['Dana_Prioritetna_Do_Odobrenja'].mean()
    st.metric("Prosjek (Prioritetna -> Odobrenje)", f"{avg_prio:.1f} radnih dana" if pd.notna(avg_prio) else "N/A")

with col4:
    avg_ceka = df_filtered['Dana_CekaRaspolozivost_Do_Odobrenja'].mean()
    st.metric("Prosjek (Čeka raspoloživost -> Odobrenje)", f"{avg_ceka:.1f} radnih dana" if pd.notna(avg_ceka) else "N/A")

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

# ---- IZVJEŠTAJ S POSTOCIMA: OD POČETNOG STATUSA DO ODOBRENE ISPORUKE ----
st.subheader("⏱️ Izvještaj: Vrijeme obrade od početnog statusa do odobrene isporuke (s udjelima)")

order_categories = ['1 dan', '2 dana', '2 - 5 dana', '5 - 10 dana', 'Više od 10 dana', 'Nema podataka']

# Izračun broja naloga i postotaka
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

# ---- TABLICA S REZULTATIMA ----
st.subheader("📋 Detaljni pregled naloga i izračunatih dana")
st.dataframe(df_filtered[[
    'Broj naloga', 'Naziv partnera', 'Skladište', 'Početni_Status_Tip', 
    'Dana_Do_Odobrenja_Ukupno', 'Kategorija_Dana', 'Dana_Rezervacija_Do_Odobrenja', 
    'Dana_Prioritetna_Do_Odobrenja', 'Dana_CekaRaspolozivost_Do_Odobrenja'
]], use_container_width=True)
