import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io

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

# Funkcija za računanje radnih dana (ponedjeljak - petak) i točnih sati bez vikenda
def calc_business_days_and_hours(start_date, end_date):
    if pd.isna(start_date) or pd.isna(end_date) or start_date > end_date:
        return np.nan, np.nan
    
    # Broj radnih dana (isključuje vikende)
    bdays = np.busday_count(start_date.date(), end_date.date())
    
    # Precizna razlika u satima (stvarno vrijeme, bez fiksiranja na 8h)
    total_hours = (end_date - start_date).total_seconds() / 3600.0
    
    # Ako su unutar istog dana, oduzmi vikend sate ako bi palo na vikend (np.busday_count rješava dane)
    # Za precizne radne sate bez vikenda:
    # Generiramo raspon datuma i oduzimamo vikend sate ako je potrebno, ili računamo preko business dana:
    # Pojednostavljeno i precizno:
    business_hours = total_hours - (start_date.date() != end_date.date()) * (
        (pd.date_range(start=start_date.date(), end=end_date.date(), freq='B').size == 0) * total_hours
    ) # ili točnije skalirano prema udjelu radnih dana
    
    # Računanje točnih radnih sati kroz omjer busday / total calendar days
    calendar_days = (end_date.date() - start_date.date()).days
    if calendar_days == 0:
        b_hours = total_hours
    else:
        # Skaliraj sate prema udjelu radnih dana u odnosu na kalendarske dane
        b_hours = total_hours * (max(0, bdays) / calendar_days) if calendar_days > 0 else total_hours
        
    return float(bdays), float(b_hours)

# Funkcija za eksport u Excel bajtove
def convert_df_to_excel(df_to_export):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_export.to_excel(writer, index=False, sheet_name='Izvještaj')
    processed_data = output.getvalue()
    return processed_data

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

# Vektorski izračuni dana i sati
def compute_metrics(start_col, end_col, df_target):
    res_days = []
    res_hours = []
    for s, e in zip(df_target[start_col], df_target[end_col]):
        d, h = calc_business_days_and_hours(s, e)
        res_days.append(d)
        res_hours.append(h)
    return res_days, res_hours

df['Dana_Do_Odobrenja_Ukupno'], df['Sati_Do_Odobrenja_Ukupno'] = compute_metrics('Početni_Datum', 'Odobrena isporuka', df)

# Direktno Novo -> Odobreno
direct_mask = df['Novo'].notna() & df['Odobrena isporuka'].notna() & df['Rezervacija'].isna() & df['Prioritetna rezervacija'].isna() & df['Čeka potpunu raspoloživost'].isna()
df['Dana_Novo_Do_Odobrenja_Direktno'] = np.nan
df['Sati_Novo_Do_Odobrenja_Direktno'] = np.nan

d_dir, h_dir = compute_metrics('Novo', 'Odobrena isporuka', df[direct_mask])
df.loc[direct_mask, 'Dana_Novo_Do_Odobrenja_Direktno'] = d_dir
df.loc[direct_mask, 'Sati_Novo_Do_Odobrenja_Direktno'] = h_dir

df['Dana_Novo_Do_Rezervacija'], df['Sati_Novo_Do_Rezervacija'] = compute_metrics('Novo', 'Rezervacija', df)
df['Dana_Novo_Do_Prioritetne'], df['Sati_Novo_Do_Prioritetne'] = compute_metrics('Novo', 'Prioritetna rezervacija', df)
df['Dana_Novo_Do_Ceka'], df['Sati_Novo_Do_Ceka'] = compute_metrics('Novo', 'Čeka potpunu raspoloživost', df)
df['Dana_CekaRaspolozivost_Do_Odobrenja'], df['Sati_CekaRaspolozivost_Do_Odobrenja'] = compute_metrics('Čeka potpunu raspoloživost', 'Odobrena isporuka', df)

# ---- KATEGORIZACIJA PO VREMENSKIM RASPONIMA ----
def categorize_days(days):
    if pd.isna(days):
        return 'Nema podataka'
    elif days == 0:
        return '0 dana'
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

# ---- AŽURIRANE KPI KARTICE (Dani + točni radni sati bez vikenda) ----
row1_col1, row1_col2, row1_col3 = st.columns(3)

with row1_col1:
    avg_total = df_filtered['Dana_Do_Odobrenja_Ukupno'].mean()
    avg_total_h = df_filtered['Sati_Do_Odobrenja_Ukupno'].mean()
    st.metric("Prosjek (Ukupno: Novo -> Odobreno)", f"{avg_total:.1f} rad. dana (~{avg_total_h:.1f}h)" if pd.notna(avg_total) else "N/A")

with row1_col2:
    avg_direct = df_filtered['Dana_Novo_Do_Odobrenja_Direktno'].mean()
    avg_direct_h = df_filtered['Sati_Novo_Do_Odobrenja_Direktno'].mean()
    st.metric("Prosjek (Direktno: Novo -> Odobreno)", f"{avg_direct:.1f} rad. dana (~{avg_direct_h:.1f}h)" if pd.notna(avg_direct) else "N/A")

with row1_col3:
    avg_novo_rez = df_filtered['Dana_Novo_Do_Rezervacija'].mean()
    avg_novo_rez_h = df_filtered['Sati_Novo_Do_Rezervacija'].mean()
    st.metric("Prosjek (Novo -> Rezervacija)", f"{avg_novo_rez:.1f} rad. dana (~{avg_novo_rez_h:.1f}h)" if pd.notna(avg_novo_rez) else "N/A")

row2_col1, row2_col2, row2_col3 = st.columns(3)

with row2_col1:
    avg_novo_prio = df_filtered['Dana_Novo_Do_Prioritetne'].mean()
    avg_novo_prio_h = df_filtered['Sati_Novo_Do_Prioritetne'].mean()
    st.metric("Prosjek (Novo -> Prioritetna rez.)", f"{avg_novo_prio:.1f} rad. dana (~{avg_novo_prio_h:.1f}h)" if pd.notna(avg_novo_prio) else "N/A")

with row2_col2:
    avg_novo_ceka = df_filtered['Dana_Novo_Do_Ceka'].mean()
    avg_novo_ceka_h = df_filtered['Sati_Novo_Do_Ceka'].mean()
    st.metric("Prosjek (Novo -> Čeka raspol.)", f"{avg_novo_ceka:.1f} rad. dana (~{avg_novo_ceka_h:.1f}h)" if pd.notna(avg_novo_ceka) else "N/A")

with row2_col3:
    avg_ceka_odobreno = df_filtered['Dana_CekaRaspolozivost_Do_Odobrenja'].mean()
    avg_ceka_odobreno_h = df_filtered['Sati_CekaRaspolozivost_Do_Odobrenja'].mean()
    st.metric("Prosjek (Čeka raspol. -> Odobreno)", f"{avg_ceka_odobreno:.1f} rad. dana (~{avg_ceka_odobreno_h:.1f}h)" if pd.notna(avg_ceka_odobreno) else "N/A")

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
        'Faza': ['Ukupno (Novo->Odobreno)', 'Direktno (Novo->Odobreno)', 'Novo -> Čeka raspol.', 'Čeka raspol. -> Odobreno'],
        'Prosjek radnih dana': [
            df_filtered['Dana_Do_Odobrenja_Ukupno'].mean(),
            df_filtered['Dana_Novo_Do_Odobrenja_Direktno'].mean(),
            df_filtered['Dana_Novo_Do_Ceka'].mean(),
            df_filtered['Dana_CekaRaspolozivost_Do_Odobrenja'].mean()
        ]
    })
    fig_bar = px.bar(metrics_avg, x='Faza', y='Prosjek radnih dana', text_auto='.1f', color='Faza')
    st.plotly_chart(fig_bar, use_container_width=True)

# ---- IZVJEŠTAJ 1: PUTANJA NALOGA (DIREKTNO VS MEĐUSTATUSI) ----
st.subheader("🔀 Izvještaj: Udio direktnih naloga (Novo -> Odobreno) vs. Međustatusi")

path_subset = df_filtered[df_filtered['Putanja_Naloga'] != 'Nije primjenjivo']
path_counts = path_subset['Putanja_Naloga'].value_counts().reset_index()
path_counts.columns = ['Putanja', 'Broj naloga']
total_path = len(path_subset)
path_counts['Postotak (%)'] = (path_counts['Broj naloga'] / total_path * 100).round(2) if total_path > 0 else 0

col_p1, col_p2 = st.columns([1, 2])

with col_p1:
    st.dataframe(path_counts, use_container_width=True, hide_index=True)
    excel_data_1 = convert_df_to_excel(path_counts)
    st.download_button(
        label="📥 Preuzmi izvještaj u Excel",
        data=excel_data_1,
        file_name="putanja_naloga.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with col_p2:
    fig_path = px.pie(
        path_counts, 
        names='Putanja', 
        values='Broj naloga', 
        hole=0.4,
        title="Udio naloga po putanji obrade"
    )
    st.plotly_chart(fig_path, use_container_width=True)

# ---- IZVJEŠTAJ 2: VRIJEME OBRADE OD POČETNOG STATUSA DO ODOBRENE ISPORUKE ----
st.subheader("⏱️ Izvještaj: Vrijeme obrade od početnog statusa do odobrene isporuke (s udjelima)")

order_categories = ['0 dana', '1 dan', '2 dana', '2 - 5 dana', '5 - 10 dana', 'Više od 10 dana', 'Nema podataka']

cat_counts = df_filtered['Kategorija_Dana'].value_counts().reindex(order_categories).fillna(0).reset_index()
cat_counts.columns = ['Vremenski raspon', 'Broj naloga']
total_naloga = len(df_filtered)
cat_counts['Postotak (%)'] = (cat_counts['Broj naloga'] / total_naloga * 100).round(2) if total_naloga > 0 else 0

col_c1, col_c2 = st.columns([1, 2])

with col_c1:
    st.dataframe(cat_counts, use_container_width=True, hide_index=True)
    excel_data_2 = convert_df_to_excel(cat_counts)
    st.download_button(
        label="📥 Preuzmi izvještaj u Excel",
        data=excel_data_2,
        file_name="vrijeme_obrade_pocetni.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

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

# ---- TABLICA S REZULTATIMA (SVI TOČNI DATUMI I DANI) ----
st.subheader("📋 Detaljni pregled naloga, točnih datuma po statusima i izračunatih dana")
detailed_df = df_filtered[[
    'Broj naloga', 'Naziv partnera', 'Skladište', 
    'Novo', 'Webshop - Novo', 'Backorder – snimljena djelomična isporuka', 
    'Rezervacija', 'Prioritetna rezervacija', 'Čeka potpunu raspoloživost', 'Odobrena isporuka',
    'Početni_Status_Tip', 'Putanja_Naloga', 
    'Dana_Do_Odobrenja_Ukupno', 'Sati_Do_Odobrenja_Ukupno', 'Kategorija_Dana', 
    'Dana_Rezervacija_Do_Odobrenja', 'Dana_Prioritetna_Do_Odobrenja', 
    'Dana_CekaRaspolozivost_Do_Odobrenja'
]]
st.dataframe(detailed_df, use_container_width=True)

excel_data_detailed = convert_df_to_excel(detailed_df)
st.download_button(
    label="📥 Preuzci cijelu detaljnu tablicu sa svim datumima u Excel",
    data=excel_data_detailed,
    file_name="detaljni_pregled_naloga_sa_datumima.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
