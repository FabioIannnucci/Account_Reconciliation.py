import pandas as pd

#Funzione di riconciliazione 
def riconcilia_conti(file_estratto_conto, file_mastrino, colonna_data, colonna_importo):
    """
    Funzione per riconciliare i conti tra estratto conto e mastrino, con tolleranza sulla data.
    
    Args:
        file_estratto_conto (str): Percorso al file Excel dell'estratto conto.
        file_mastrino (str): Percorso al file Excel del mastrino.
        colonna_data (str): Nome della colonna con le date nei file.
        colonna_importo (str): Nome della colonna con gli importi nei file.

    Returns:
        pd.DataFrame: DataFrame unico con movimenti matchati e non matchati.
    """
    # Carica i file Excel
    estratto_conto = pd.read_excel(file_estratto_conto)
    mastrino = pd.read_excel(file_mastrino)
    
    # Verifica la presenza delle colonne richieste
    for colonna in [colonna_data, colonna_importo]:
        if colonna not in estratto_conto.columns or colonna not in mastrino.columns:
            raise ValueError(f"Colonna '{colonna}' non trovata in uno dei file.")
    
    # Converti le colonne data in formato datetime
    estratto_conto[colonna_data] = pd.to_datetime(estratto_conto[colonna_data], errors='coerce')
    mastrino[colonna_data] = pd.to_datetime(mastrino[colonna_data], errors='coerce')
    
    # Creiamo liste per i risultati
    risultati = []
    
    # Tolleranza di 3 giorni per il matching
    for idx, row in estratto_conto.iterrows():
        mask = (
            (mastrino[colonna_importo] == row[colonna_importo]) &
            (mastrino[colonna_data].between(row[colonna_data] - pd.Timedelta(days=3),
                                            row[colonna_data] + pd.Timedelta(days=3)))
        )
        match = mastrino[mask]
        
        if not match.empty:
            # Match trovato
            risultati.append({
                'Data': row[colonna_data],
                'Importo': row[colonna_importo],
                'Origine': 'Matchato',
                'Dettaglio': f"Match con mastrino in data {match.iloc[0][colonna_data]}"
            })
            mastrino = mastrino.drop(match.index)
        else:
            # Nessun match trovato per estratto conto
            risultati.append({
                'Data': row[colonna_data],
                'Importo': row[colonna_importo],
                'Origine': 'Non trovato nel mastrino',
                'Dettaglio': ''
            })
    
    # Aggiungi i movimenti residui dal mastrino
    for idx, row in mastrino.iterrows():
        risultati.append({
            'Data': row[colonna_data],
            'Importo': row[colonna_importo],
            'Origine': 'Non trovato nell’estratto conto',
            'Dettaglio': ''
        })
    
    # Crea un DataFrame dai risultati
    risultati_df = pd.DataFrame(risultati)
    
    # Ordina per Data e Importo
    risultati_df = risultati_df.sort_values(by=['Data', 'Importo']).reset_index(drop=True)
    
    return risultati_df

# Esempio di utilizzo
file_estratto_conto = 'C:/Users/iannu/Downloads/conto_corrente_clean.xlsx'
file_mastrino = 'C:/Users/iannu/Desktop/RICONCILIAZIONE CONTO/mastrino_clean.xlsx'
colonna_data = 'Data'  # Modifica con il nome effettivo della colonna data
colonna_importo = 'Importo'  # Modifica con il nome effettivo della colonna importo

# Genera il file di riconciliazione
try:
    risultato = riconcilia_conti(file_estratto_conto, file_mastrino, colonna_data, colonna_importo)
    
    # Salva il risultato in un unico file
    risultato.to_excel('riconciliazione_completa.xlsx', index=False)
    print("Riconciliazione completata! File salvato come 'riconciliazione_completa.xlsx'.")
except Exception as e:
    print(f"Errore durante la riconciliazione: {e}")
import pandas as pd

def riconcilia_conti(file_estratto_conto, file_mastrino, colonna_data, colonna_importo):
    """
    Funzione per riconciliare i conti tra estratto conto e mastrino, con tolleranza sulla data.
    
    Args:
        file_estratto_conto (str): Percorso al file Excel dell'estratto conto.
        file_mastrino (str): Percorso al file Excel del mastrino.
        colonna_data (str): Nome della colonna con le date nei file.
        colonna_importo (str): Nome della colonna con gli importi nei file.

    Returns:
        pd.DataFrame: DataFrame unico con movimenti matchati e non matchati.
    """
    # Carica i file Excel
    estratto_conto = pd.read_excel(file_estratto_conto)
    mastrino = pd.read_excel(file_mastrino)
    
    # Verifica la presenza delle colonne richieste
    for colonna in [colonna_data, colonna_importo]:
        if colonna not in estratto_conto.columns or colonna not in mastrino.columns:
            raise ValueError(f"Colonna '{colonna}' non trovata in uno dei file.")
    
    # Converti le colonne data in formato datetime
    estratto_conto[colonna_data] = pd.to_datetime(estratto_conto[colonna_data], errors='coerce')
    mastrino[colonna_data] = pd.to_datetime(mastrino[colonna_data], errors='coerce')
    
    # Creiamo liste per i risultati
    risultati = []
    
    # Tolleranza di 3 giorni per il matching
    for idx, row in estratto_conto.iterrows():
        mask = (
            (mastrino[colonna_importo] == row[colonna_importo]) &
            (mastrino[colonna_data].between(row[colonna_data] - pd.Timedelta(days=3),
                                            row[colonna_data] + pd.Timedelta(days=3)))
        )
        match = mastrino[mask]
        
        if not match.empty:
            # Match trovato
            risultati.append({
                'Data': row[colonna_data],
                'Importo': row[colonna_importo],
                'Origine': 'Matchato',
                'Dettaglio': f"Match con mastrino in data {match.iloc[0][colonna_data]}"
            })
            mastrino = mastrino.drop(match.index)
        else:
            # Nessun match trovato per estratto conto
            risultati.append({
                'Data': row[colonna_data],
                'Importo': row[colonna_importo],
                'Origine': 'Non trovato nel mastrino',
                'Dettaglio': ''
            })
    
    # Aggiungi i movimenti residui dal mastrino
    for idx, row in mastrino.iterrows():
        risultati.append({
            'Data': row[colonna_data],
            'Importo': row[colonna_importo],
            'Origine': 'Non trovato nell’estratto conto',
            'Dettaglio': ''
        })
    
    # Crea un DataFrame dai risultati
    risultati_df = pd.DataFrame(risultati)
    
    # Ordina per Data e Importo
    risultati_df = risultati_df.sort_values(by=['Data', 'Importo']).reset_index(drop=True)
    
    return risultati_df

# Esempio di utilizzo
file_estratto_conto = 'C:/Users/iannu/Downloads/conto_corrente_clean.xlsx'
file_mastrino = 'C:/Users/iannu/Desktop/RICONCILIAZIONE CONTO/mastrino_clean.xlsx'
colonna_data = 'Data'  # Modifica con il nome effettivo della colonna data
colonna_importo = 'Importo'  # Modifica con il nome effettivo della colonna importo

# Genera il file di riconciliazione
try:
    risultato = riconcilia_conti(file_estratto_conto, file_mastrino, colonna_data, colonna_importo)
    
    # Salva il risultato in un unico file
    risultato.to_excel('riconciliazione_completa.xlsx', index=False)
    print("Riconciliazione completata! File salvato come 'riconciliazione_completa.xlsx'.")
except Exception as e:
    print(f"Errore durante la riconciliazione: {e}")

