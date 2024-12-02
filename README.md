# Riconciliazione Estratto Conto e Mastrino

Questo progetto Python è stato progettato per automatizzare il processo di riconciliazione tra un estratto conto bancario e il mastrino contabile. La funzione include anche un controllo avanzato per voci specifiche come **SPE** (commissioni su bonifici/operazioni bancarie) e **SPC** (commissioni POS), confrontandole con i movimenti codificati come **016** e **066** nell'estratto conto.

---

## Caratteristiche

- **Riconciliazione automatizzata**:
  - Confronta i movimenti tra estratto conto e mastrino sulla base di:
    - **Data** (match esatto, senza tolleranza di giorni).
    - **Importo**.
    - **Descrizione**.
  
- **Controllo dinamico SPE/SPC**:
  - Verifica i movimenti **SPE/SPC** del mastrino confrontandoli con la **sommatoria** di movimenti **016/066** dell'estratto conto per la stessa data.
  - Utilizza combinazioni dinamiche per individuare eventuali match.

- **Output chiaro**:
  - Genera un file Excel contenente:
    - Movimenti matchati.
    - Movimenti non trovati (sia in mastrino che nell'estratto conto).
    - Risultati del controllo SPE/SPC.

---

## Requisiti

### Librerie Python
- `pandas`: per la manipolazione dei dati.
- `openpyxl`: per la lettura e scrittura di file Excel.
- `itertools`: per gestire combinazioni dinamiche.

### Installazione delle dipendenze
Per installare le librerie richieste, eseguire:
## Input richiesto

### File richiesti
1. **Estratto Conto** (formato `.xlsx`):
   - Deve contenere tre colonne:
     - `DATA`: Data del movimento (formato `DD/MM/YYYY` o `YYYY-MM-DD`).
     - `DESCRIZIONE`: Descrizione del movimento.
     - `IMPORTO`: Importo del movimento (positivo/negativo).

2. **Mastrino** (formato `.xlsx`):
   - Deve contenere tre colonne:
     - `DATA`: Data del movimento (formato `DD/MM/YYYY` o `YYYY-MM-DD`).
     - `IMPORTO`: Importo del movimento.
     - `DESCRIZIONE`: Descrizione del movimento.

### Nota
I nomi delle colonne dei file vengono standardizzati automaticamente dal programma. Assicurarsi che i file abbiano colonne coerenti con quelle indicate sopra.

---

## Come usare il codice

1. **Posizionare i file**:
   - Salvare i file di estratto conto e mastrino nella stessa directory del codice.

2. **Modificare i percorsi dei file**:
   - Specificare i percorsi dei file nei seguenti parametri all'interno del codice:
     ```python
     file_estratto_conto = 'Percorso_del_file_estratto_conto.xlsx'
     file_mastrino = 'Percorso_del_file_mastrino.xlsx'
     ```

3. **Eseguire il codice**:
   - Lanciare lo script Python:
     ```bash
     python riconciliazione.py
     ```

4. **Output**:
   - Il file risultante verrà salvato nella stessa directory con il nome `riconciliazione_completa_con_spe_spc.xlsx`.

---

## Struttura dell'Output

Il file Excel generato contiene:
1. **Movimenti matchati**:
   - Data, Importo, e Dettagli del match.

2. **Movimenti non matchati**:
   - Movimenti presenti nell'estratto conto ma non nel mastrino (e viceversa).

3. **Esito del controllo SPE/SPC**:
   - Viene indicato se è stato trovato un match o un mismatch per le voci SPE/SPC rispetto ai movimenti 016/066.

---

## Debug

Il codice include stampe di debug per:
- Colonne lette dai file.
- Prime righe dei dati caricati.
- Risultati intermedi delle operazioni di riconciliazione e del controllo SPE/SPC.

---



