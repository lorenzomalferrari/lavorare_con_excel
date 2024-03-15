import pandas as pd

# Creazione dei dataframe per ogni foglio
df_persone = pd.DataFrame({
    'id': range(1, 11),
    'nome': ['Nome{}'.format(i) for i in range(1, 11)],
    'cognome': ['Cognome{}'.format(i) for i in range(1, 11)],
    'data_di_nascita': pd.date_range(start='1980-01-01', periods=10),
    'nazionalità': ['Nazionalità{}'.format(i) for i in range(1, 11)]
})

df_servizi = pd.DataFrame({
    'id': range(1, 21),
    'nome_servizio': ['Servizio{}'.format(i) for i in range(1, 21)],
    'costo_ora': [10 * i for i in range(1, 21)],
    'note': ['Note{}'.format(i) for i in range(1, 21)]
})

df_fatture = pd.DataFrame({
    'id': range(1, 41),
    'codice_fattura': ['Fattura{}'.format(i) for i in range(1, 41)],
    'anno_fattura': [2024] * 40, # Anno fisso per esempio
    'descrizione_lavoro': ['Descrizione{}'.format(i) for i in range(1, 41)],
    'id_servizio': [i % 20 + 1 for i in range(40)],
    'id_persona': [i % 10 + 1 for i in range(40)]
})

# Creazione del file Excel
with pd.ExcelWriter('dati.xlsx') as writer:
    df_persone.to_excel(writer, sheet_name='Persone', index=False)
    df_servizi.to_excel(writer, sheet_name='Servizi', index=False)
    df_fatture.to_excel(writer, sheet_name='Fatture', index=False)
