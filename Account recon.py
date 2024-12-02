from itertools import combinations
import pandas as pd

def riconcilia_conti(file_estratto_conto, file_mastrino, colonna_data, colonna_importo, colonna_descrizione):
    """
    Funzione per riconciliare i conti tra estratto conto e mastrino, includendo check SPE/SPC e codice 016/066.
    """
    # Carica i file Excel
    estratto_conto = pd.read_excel(file_estratto_conto, engine='openpyxl')
    mastrino = pd.read_excel(file_mastrino, engine='openpyxl')
    
    # Forza i nomi delle colonne a essere coerenti
    estratto_conto.columns = ['DATA', 'DESCRIZIONE', 'IMPORTO']
    mastrino.columns = ['DATA', 'IMPORTO', 'DESCRIZIONE']
    
    # Converti le colonne data in formato datetime
    estratto_conto['DATA'] = pd.to_datetime(estratto_conto['DATA'], errors='coerce', dayfirst=True)
    mastrino['DATA'] = pd.to_datetime(mastrino['DATA'], errors='coerce', dayfirst=True)

    # Creiamo liste per i risultati
    risultati = []
    voci_esaminate = set()  # Set per tracciare le righe già gestite

    # Matching esatto (senza tolleranza di giorni)
    for idx, row in estratto_conto.iterrows():
        if idx in voci_esaminate:  # Salta la voce se è già stata elaborata
            continue

        mask = (
            (mastrino['IMPORTO'] == row['IMPORTO']) &
            (mastrino['DATA'] == row['DATA'])
        )
        match = mastrino[mask]
        
        if not match.empty:
            # Match trovato
            risultati.append({
                'Data': row['DATA'],
                'Importo': row['IMPORTO'],
                'Origine': 'Matchato',
                'Descrizione_Estratto': row['DESCRIZIONE'],
                'Descrizione_Mastrino': match.iloc[0]['DESCRIZIONE'],
                'Dettaglio': f"Match con mastrino in data {match.iloc[0]['DATA']}"
            })
            mastrino = mastrino.drop(match.index)
            voci_esaminate.add(idx)  # Segna la riga come gestita
        else:
            # Nessun match trovato per estratto conto
            risultati.append({
                'Data': row['DATA'],
                'Importo': row['IMPORTO'],
                'Origine': 'Non trovato nel mastrino',
                'Descrizione_Estratto': row['DESCRIZIONE'],
                'Descrizione_Mastrino': None,
                'Dettaglio': ''
            })
            voci_esaminate.add(idx)  # Segna la riga come gestita

    # Movimenti residui dal mastrino
    for idx, row in mastrino.iterrows():
        risultati.append({
            'Data': row['DATA'],
            'Importo': row['IMPORTO'],
            'Origine': 'Non trovato nell’estratto conto',
            'Descrizione_Estratto': None,
            'Descrizione_Mastrino': row['DESCRIZIONE'],
            'Dettaglio': ''
        })
    
    # Crea un DataFrame dai risultati
    risultati_df = pd.DataFrame(risultati)
    
    # Ordina per Data e Importo
    risultati_df = risultati_df.sort_values(by=['Data', 'Importo']).reset_index(drop=True)
    
    # Integra il controllo SPE/SPC
    risultati_df = integrate_spe_spc_check(
        estratto_conto=estratto_conto,
        mastrino=mastrino,
        risultati_df=risultati_df,
        colonna_data='DATA',
        colonna_importo='IMPORTO',
        colonna_descrizione='DESCRIZIONE'
    )
    
    return risultati_df


def integrate_spe_spc_check(estratto_conto, mastrino, risultati_df, colonna_data, colonna_importo, colonna_descrizione):
    """
    Controlla dinamicamente le voci SPE/SPC nel mastrino contro le combinazioni di movimenti 016/066 nell'estratto conto.
    """
    # Filtra le voci SPE/SPC dal mastrino
    mastrino_spe_spc = mastrino[mastrino[colonna_descrizione].str.contains('SPE|SPC', na=False)]

    # Filtra le voci con codice 016/066 dall'estratto conto
    estratto_conto_016_066 = estratto_conto[estratto_conto[colonna_descrizione].str.contains('016|066', na=False)]

    # Lista per i risultati del controllo SPE/SPC
    spe_spc_results = []

    # Itera sulle voci SPE/SPC del mastrino
    for _, riga_mastrino in mastrino_spe_spc.iterrows():
        data_mastrino = riga_mastrino[colonna_data]
        importo_mastrino = riga_mastrino[colonna_importo]

        # Trova tutti i movimenti 016/066 dello stesso giorno nell'estratto conto
        movimenti_giorno = estratto_conto_016_066[estratto_conto_016_066[colonna_data] == data_mastrino]

        if not movimenti_giorno.empty:
            # Estrai gli importi dei movimenti 016/066
            importi = movimenti_giorno[colonna_importo].tolist()
            combinazione_trovata = False

            # Prova tutte le combinazioni possibili
            for i in range(1, len(importi) + 1):  # Lunghezza delle combinazioni
                for combinazione in combinations(importi, i):
                    if sum(combinazione) == importo_mastrino:
                        # Match trovato
                        spe_spc_results.append({
                            'Data': data_mastrino,
                            'Importo_Mastrino': importo_mastrino,
                            'Importo_EC_Combinazione': combinazione,
                            'Descrizione_Mastrino': riga_mastrino[colonna_descrizione],
                            'Esito': 'Match SPE/SPC'
                        })
                        combinazione_trovata = True
                        break
                if combinazione_trovata:
                    break

            if not combinazione_trovata:
                # Nessuna combinazione trovata
                spe_spc_results.append({
                    'Data': data_mastrino,
                    'Importo_Mastrino': importo_mastrino,
                    'Importo_EC_Combinazione': None,
                    'Descrizione_Mastrino': riga_mastrino[colonna_descrizione],
                    'Esito': 'Mismatch: Nessuna combinazione corrisponde'
                })
        else:
            # Nessun movimento 016/066 per quella data
            spe_spc_results.append({
                'Data': data_mastrino,
                'Importo_Mastrino': importo_mastrino,
                'Importo_EC_Combinazione': None,
                'Descrizione_Mastrino': riga_mastrino[colonna_descrizione],
                'Esito': 'Mismatch: Nessun movimento 016/066 trovato'
            })

    # Converte i risultati in DataFrame
    spe_spc_df = pd.DataFrame(spe_spc_results)

    # Concatena i risultati al DataFrame principale
    risultati_df = pd.concat([risultati_df, spe_spc_df], ignore_index=True)

    # Ordina per Data e Importo per chiarezza
    risultati_df = risultati_df.sort_values(by=['Data', 'Importo'], ascending=[True, True]).reset_index(drop=True)

    return risultati_df


# Esempio di utilizzo
file_estratto_conto = 'Insert_bank_File'
file_mastrino = 'Insert_ledger_file'

try:
    risultato = riconcilia_conti(
        file_estratto_conto=file_estratto_conto,
        file_mastrino=file_mastrino,
        colonna_data='DATA',
        colonna_importo='IMPORTO',
        colonna_descrizione='DESCRIZIONE'
    )
    
    # Salva i risultati in un file Excel
    risultato.to_excel('riconciliazione_completa_con_spe_spc.xlsx', index=False)
    print("Riconciliazione completata! File salvato come 'riconciliazione_completa_con_spe_spc.xlsx'.")
except Exception as e:
    print(f"Errore durante la riconciliazione: {e}")
