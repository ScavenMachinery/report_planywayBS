import requests
import pip
pip.main(["install", "openpyxl"])
import streamlit as st
import pandas as pd
import plotly.express as px

# Imposta la configurazione della pagina
st.set_page_config(
    layout="wide",
    page_title='REPORT PLANYWAY'
)

st.title("ANALISI OPERATIVITA'")
st.markdown("_source.PL v.1.0_")

# Carica il file Excel
@st.cache_data
def load_data(file):
    data = pd.read_excel(file)
    return data

uploaded_file = st.sidebar.file_uploader("Carica planyway aggiornato")

if uploaded_file is None:
    st.info("Carica un file tramite il menu laterale")
    st.stop()

df = load_data(uploaded_file)

# Seleziona solo le colonne desiderate
desired_columns = ["Board", "List", "Card", "Member", "Date", "StartTime", "EndTime", "DurationHours"]
df = df[desired_columns]

# Assicurati che la colonna "Date" sia interpretata come una data
df['Date'] = pd.to_datetime(df['Date'])

# Estrai il mese e formattalo come "gg-mmmm-aaaa"
df['Mese'] = df['Date'].dt.strftime('%B')

# Puoi anche rimuovere l'orario dalla colonna "Date" se non ti serve più
df['Date'] = df['Date'].dt.date

# Sostituisci i valori vuoti nella colonna "Member" con "LOST"
df['Member'].fillna('LOST', inplace=True)

# Ora hai un DataFrame con la colonna "Mese", la colonna "Member" pulita, e il formato desiderato per la data

# Puoi anche rimuovere le righe con valori mancanti, se necessario
df.dropna(inplace=True)

with st.expander("Preview Table"):
    st.dataframe(df)


# Calcola la somma totale di "DurationHours"
total_duration = df['DurationHours'].sum()



# Calcola il totale in minuti e giorni
total_duration_minutes = round(total_duration * 60,2)
total_duration_days = round(total_duration / 24,2)

# Ora hai calcolato i totali in minuti e giorni
colA, colB, colC = st.columns(3)

with colB:
    st.metric(label='ORE',value=total_duration)
with colA:
    st.metric("MINUTI", total_duration_minutes)
with colC:
    st.metric("GIORNI", total_duration_days)



# Creazione della selezione per il tipo di analisi
tipo_analisi = st.sidebar.radio("Seleziona il tipo di analisi:", ["ANALISI TEAM :sunglasses:", "ANALISI LAVORAZIONI :necktie:"])

# Blocco per l'ANALISI TEAM
if tipo_analisi == "ANALISI TEAM :sunglasses:":
    st.subheader("_ANALISI TEAM_", divider="orange")

    # Calcola le somme di "DurationHours" basate su ogni valore unico di "Member"
    member_duration = df.groupby('Member')['DurationHours'].sum()

    # Distribuisci le KPI su righe con 4 colonne per riga
    col1, col2, col3, col4 = st.columns(4)
    kpi_columns = st.columns(4)

    for i, (member, duration) in enumerate(member_duration.items()):
        kpi_columns[i % 4].metric(member, duration)

    # Crea una selezione a radio per il tipo di visualizzazione
    visualizzazione = st.radio("Seleziona il tipo di visualizzazione:", ["Bar Charts 📊", "Pie Charts 🥧"])
    if visualizzazione == "Bar Charts 📊":
        # Distribuisci i grafici a barre per ogni "Member" due a due in una colonna sotto l'altra
        members = list(member_duration.keys())
        for i in range(0, len(members), 2):
            col1, col2 = st.columns(2)
            member1 = members[i]
            member_df1 = df[df['Member'] == member1]
            fig1 = px.histogram(member_df1, x='Board', y='DurationHours', title=f'KPI per {member1}')
            fig1.update_layout(xaxis_title="Board", yaxis_title="DurationHours")
            col1.plotly_chart(fig1)

            if i + 1 < len(members):
                member2 = members[i + 1]
                member_df2 = df[df['Member'] == member2]
                fig2 = px.bar(member_df2, x='Board', y='DurationHours', title=f'KPI per {member2}')
                fig2.update_layout(xaxis_title="Board", yaxis_title="DurationHours")
                col2.plotly_chart(fig2)
    elif visualizzazione == "Pie Charts 🥧":
        # Distribuisci i grafici a torta impostati come % per ogni Member
        members = list(member_duration.keys())
        for i in range(0, len(members), 2):
            col1, col2 = st.columns(2)
            member1 = members[i]
            member_df1 = df[df['Member'] == member1]
            fig1 = px.pie(member_df1, names='Board', values='DurationHours', title=f'KPI per {member1} (%)')
            col1.plotly_chart(fig1)

            if i + 1 < len(members):
                member2 = members[i + 1]
                member_df2 = df[df['Member'] == member2]
                fig2 = px.pie(member_df2, names='Board', values='DurationHours', title=f'KPI per {member2} (%)')
                col2.plotly_chart(fig2)

# Blocco per l'ANALISI LAVORAZIONI
elif tipo_analisi == "ANALISI LAVORAZIONI :necktie:":
    st.subheader("_ANALISI LAVORAZIONI_", divider="orange")
    
    # Selezione per il livello di analisi
    livello_analisi = st.radio("Seleziona il livello di analisi:", ["BOARD LEVEL", "LIST LEVEL", "CARD LEVEL"])
    
    if livello_analisi == "BOARD LEVEL":
        st.write("_ANALISI BOARD LEVEL_")
    
        # Raggruppa il DataFrame per la colonna "Board" e calcola la somma di "DurationHours" per ciascun valore unico
        board_level_data = df.groupby('Board')['DurationHours'].sum().reset_index()

        # Crea un grafico a barre
        fig = px.bar(board_level_data, x='Board', y='DurationHours', title='Analisi BOARD LEVEL')
        fig.update_layout(xaxis_title="Board", yaxis_title="Total DurationHours")
    
        # Visualizza il grafico
        st.plotly_chart(fig, use_container_width=True)
        
    # Visualizza la preview della tabella
        with st.expander("Anteprima dei dati utilizzati per il grafico"):
            st.dataframe(board_level_data)
   # Codice per l'analisi a livello di "LIST LEVEL" con grafico a torta in percentuale
    # Codice per l'analisi a livello di "LIST LEVEL" con filtro per il numero di voci da visualizzare nei grafici
    elif livello_analisi == "LIST LEVEL":
        st.write("_ANALISI LIST LEVEL_")

        # Raggruppa il DataFrame per la colonna "List" e calcola la somma di "DurationHours" per ciascun valore unico
        list_level_data = df.groupby('List')['DurationHours'].sum().reset_index()

        # Ordina il DataFrame in base a "DurationHours" in ordine decrescente
        list_level_data = list_level_data.sort_values(by='DurationHours', ascending=False)

        # Aggiungi un filtro per il numero di voci da visualizzare nei grafici
        num_entries = st.slider("Seleziona il numero di voci di List da visualizzare nei grafici", 1, len(list_level_data), 20)

        # Filtra le prime "num_entries" voci di "List"
        top_list = list_level_data.head(num_entries)

        # Calcola la percentuale per il grafico a torta
        top_list['Percentage'] = (top_list['DurationHours'] / top_list['DurationHours'].sum()) * 100

        # Crea un grafico a barre
        fig_bar = px.bar(top_list, x='List', y='DurationHours', title=f'Analisi LIST LEVEL - Top {num_entries} List (Bar Chart)')
        fig_bar.update_layout(xaxis_title="List", yaxis_title="Total DurationHours")

        # Visualizza il grafico a barre
        st.plotly_chart(fig_bar, use_container_width=True)

        # Crea il grafico a torta
        fig_pie = px.pie(top_list, names='List', values='Percentage', title=f'Analisi LIST LEVEL - Top {num_entries} List (Pie Chart in %)')

        # Visualizza il grafico a torta
        st.plotly_chart(fig_pie, use_container_width=True)
    # Codice per l'analisi a livello di "CARD LEVEL"
    elif livello_analisi == "CARD LEVEL":
        st.write("_ANALISI CARD LEVEL_")

        # Aggiungi un filtro per la selezione di "List" ordinata per "DurationHours" decrescente
        selected_list = st.selectbox("Seleziona una List", df.groupby('List')['DurationHours'].sum().reset_index().sort_values(by='DurationHours', ascending=False)['List'].tolist())

        # Filtra il DataFrame in base alla List selezionata
        filtered_df = df[df['List'] == selected_list]

        # Crea un grafico a barre con "Card" in X e "DurationHours" in Y
        fig_bar = px.bar(filtered_df, x='Card', y='DurationHours', title=f'Analisi CARD LEVEL - List: {selected_list} (Bar Chart)')
        fig_bar.update_layout(xaxis_title="Card", yaxis_title="DurationHours")

        # Visualizza il grafico a barre principale
        st.plotly_chart(fig_bar, use_container_width=True)

        # Crea un secondo grafico a barre con "Member" in X, "DurationHours" in Y, e colori diversi per ogni "Member"
        member_bar_data = filtered_df.groupby('Member')['DurationHours'].sum().reset_index()
        fig_member_bar = px.bar(member_bar_data, x='Member', y='DurationHours', title=f'Analisi CARD LEVEL - List: {selected_list} (Member Bar Chart)', color='Member')
        fig_member_bar.update_layout(xaxis_title="Member", yaxis_title="DurationHours")

        # Visualizza il secondo grafico a barre
        st.plotly_chart(fig_member_bar, use_container_width=True)
       
