import streamlit as st
import pandas as pd
from io import BytesIO

# Funzione per creare il file Excel con formule
def create_excel(data_corrispettivi, data_cassa, negozio):
    # Lista per la colonna 'Descrizione' del foglio "Corrispettivi"
    descrizione_corrispettivi = [
        f"Negozio {negozio}", "data", "Nr azzeramento fiscale", "", 
        "Scontrini annullati allegare", "", 
        "CORRISPETTIVI", "", 
        "Incassi POS", "Incasso POS Corner", "", 
        "Incassi contanti", "Pay by Link", 
        "Corrispettivi giorno incassati", "", 
        "FATTURE", "Incasso FATTURE POS", 
        "incasso FATTURE CONTANTI"
    ]
    
    # Lista per la colonna 'Valore' del foglio "Corrispettivi"
    valore_corrispettivi = [
        None,  # Per allineare la lunghezza delle liste
        data_corrispettivi['data'], 
        data_corrispettivi['nr_azzeramento'], 
        None, None, None, None,
        data_corrispettivi['incassi_pos'], 
        data_corrispettivi['incasso_pos_corner'], 
        None, data_corrispettivi['incassi_contanti'], 
        data_corrispettivi['pay_by_link'], 
        "=SUM(B9:B13)",  # Formula per Corrispettivi giorno incassati
        None, None, data_corrispettivi['incasso_fatture_pos'], 
        data_corrispettivi['incasso_fatture_contanti']
    ]

    # Verifica che le liste abbiano la stessa lunghezza
    if len(descrizione_corrispettivi) != len(valore_corrispettivi):
        raise ValueError("Le liste 'descrizione_corrispettivi' e 'valore_corrispettivi' devono avere la stessa lunghezza.")

    # Crea un DataFrame per il foglio "Corrispettivi"
    df_corrispettivi = pd.DataFrame({
        'Descrizione': descrizione_corrispettivi,
        'Valore': valore_corrispettivi
    })

    # Lista per la colonna 'Descrizione' del foglio "Cassa"
    descrizione_cassa = [
        "Saldo GIORNO PRECEDENTE", "", "Entrate", 
        "TOTALE incassi contanti", "", "", "", 
        "Uscite Extra", data_cassa['uscita1_descr'], data_cassa['uscita2_descr'], 
        data_cassa['uscita3_descr'], "", "", "", 
        "Saldo CASSA GIORNATA ODIERNA"
    ]

    # Lista per la colonna 'Valore' del foglio "Cassa"
    valore_cassa = [
        data_cassa['saldo_precedente'], None, None, 
        "=Corrispettivi!B11 + Corrispettivi!B17",  # Formula aggiornata per Totale Incassi Contanti
        None, None, None, 
        None, data_cassa['uscita1_valore'], data_cassa['uscita2_valore'], 
        data_cassa['uscita3_valore'], None, None, None,
        "=B1 + SUM(B4:B7) - SUM(B9:B11)"  # Formula aggiornata per Saldo Cassa Giornata Odierna
    ]

    # Verifica che le liste abbiano la stessa lunghezza
    if len(descrizione_cassa) != len(valore_cassa):
        raise ValueError("Le liste 'descrizione_cassa' e 'valore_cassa' devono avere la stessa lunghezza.")

    # Crea un DataFrame per il foglio "Cassa"
    df_cassa = pd.DataFrame({
        'Descrizione': descrizione_cassa,
        'Valore': valore_cassa
    })

    # Scrittura in Excel
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    df_corrispettivi.to_excel(writer, sheet_name='Corrispettivi', index=False, header=False)
    df_cassa.to_excel(writer, sheet_name='Cassa', index=False, header=False)
    
    writer.close()  # Usa close() invece di save()
    processed_data = output.getvalue()
    return processed_data

# Interfaccia Streamlit
st.title("Generatore di Excel Corrispettivi Cassa")

st.header("Seleziona il negozio")
negozio = st.radio("Seleziona il negozio", ["Cagliari", "Porto Cervo", "Castel Maggiore"])

st.header("Inserisci i dati per il foglio Corrispettivi")
data_corrispettivi = {
    'data': st.date_input("Data"),
    'nr_azzeramento': st.text_input("Numero Azzeramento Fiscale"),
    'incassi_pos': st.number_input("Incassi POS", 0.0),
    'incasso_pos_corner': st.number_input("Incasso POS Corner", 0.0),
    'incassi_contanti': st.number_input("Incassi Contanti", 0.0),
    'pay_by_link': st.number_input("Pay by Link", 0.0),
    'incasso_fatture_pos': st.number_input("Incasso Fatture POS", 0.0),
    'incasso_fatture_contanti': st.number_input("Incasso Fatture Contanti", 0.0)
}

st.header("Inserisci i dati per il foglio Cassa")
data_cassa = {
    'saldo_precedente': st.number_input("Saldo Giorno Precedente", 0.0),
    'uscita1_descr': st.text_input("Descrizione Uscita 1", "Fogli A4"),
    'uscita1_valore': st.number_input("Importo Uscita 1", 0.0),
    'uscita2_descr': st.text_input("Descrizione Uscita 2", ""),
    'uscita2_valore': st.number_input("Importo Uscita 2", 0.0),
    'uscita3_descr': st.text_input("Descrizione Uscita 3", ""),
    'uscita3_valore': st.number_input("Importo Uscita 3", 0.0)
}

# Pulsante per scaricare il file Excel
if st.button("Genera e Scarica il file Excel"):
    excel_data = create_excel(data_corrispettivi, data_cassa, negozio)
    file_name = f"{negozio}_{data_corrispettivi['data']}_corrispettivi_cassa.xlsx"
    st.download_button(label="Download Excel", data=excel_data, file_name=file_name)
