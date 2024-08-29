import streamlit as st
import pandas as pd
from io import BytesIO

# Funzione per creare il file Excel
def create_excel(data_corrispettivi, data_cassa):
    # Crea un DataFrame per ogni foglio
    df_corrispettivi = pd.DataFrame({
        '': [
            "data", "Nr azzeramento fiscale", "", 
            "Scontrini annullati allegare", "", 
            "CORRISPETTIVI", "", 
            "Incassi POS", "Incasso POS Corner", "", 
            "Incassi contanti", "Pay by Link", 
            "Corrispettivi giorno incassati", "", 
            "FATTURE", "Incasso FATTURE POS", 
            "incasso FATTURE CONTANTI"
        ],
        'Valore': [
            data_corrispettivi['data'], 
            data_corrispettivi['nr_azzeramento'], 
            None, None, None, None, None,
            data_corrispettivi['incassi_pos'], 
            data_corrispettivi['incasso_pos_corner'], 
            None, data_corrispettivi['incassi_contanti'], 
            data_corrispettivi['pay_by_link'], 
            data_corrispettivi['corrispettivi_giorno'], 
            None, None, data_corrispettivi['incasso_fatture_pos'], 
            data_corrispettivi['incasso_fatture_contanti']
        ]
    })

    df_cassa = pd.DataFrame({
        'Descrizione': [
            "Saldo GIORNO PRECEDENTE", "", "Entrate", 
            "TOTALE incassi contanti", "", "", "", 
            "Uscite (pagamenti vari)", "Fogli A4", "", "", "", 
            "Saldo CASSA GIORNATA ODIERNA"
        ],
        'Valore': [
            data_cassa['saldo_precedente'], None, None, 
            data_cassa['totale_incassi'], None, None, None,
            None, data_cassa['uscite_fogli_a4'], None, 
            None, None, data_cassa['saldo_odierno']
        ]
    })

    # Scrittura in Excel
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    df_corrispettivi.to_excel(writer, sheet_name='Corrispettivi', index=False, header=False)
    df_cassa.to_excel(writer, sheet_name='Cassa', index=False, header=False)
    
    writer.save()
    processed_data = output.getvalue()
    return processed_data

# Interfaccia Streamlit
st.title("Generatore di Excel Corrispettivi Cassa")

st.header("Inserisci i dati per il foglio Corrispettivi")
data_corrispettivi = {
    'data': st.date_input("Data"),
    'nr_azzeramento': st.text_input("Numero Azzeramento Fiscale"),
    'incassi_pos': st.number_input("Incassi POS", 0.0),
    'incasso_pos_corner': st.number_input("Incasso POS Corner", 0.0),
    'incassi_contanti': st.number_input("Incassi Contanti", 0.0),
    'pay_by_link': st.number_input("Pay by Link", 0.0),
    'corrispettivi_giorno': st.number_input("Corrispettivi Giorno Incassati", 0.0),
    'incasso_fatture_pos': st.number_input("Incasso Fatture POS", 0.0),
    'incasso_fatture_contanti': st.number_input("Incasso Fatture Contanti", 0.0)
}

st.header("Inserisci i dati per il foglio Cassa")
data_cassa = {
    'saldo_precedente': st.number_input("Saldo Giorno Precedente", 0.0),
    'totale_incassi': st.number_input("Totale Incassi Contanti", 0.0),
    'uscite_fogli_a4': st.number_input("Uscite Fogli A4", 0.0),
    'saldo_odierno': st.number_input("Saldo Cassa Giornata Odierna", 0.0)
}

# Pulsante per scaricare il file Excel
if st.button("Genera e Scarica il file Excel"):
    excel_data = create_excel(data_corrispettivi, data_cassa)
    st.download_button(label="Download Excel", data=excel_data, file_name="corrispettivi_cassa.xlsx")

