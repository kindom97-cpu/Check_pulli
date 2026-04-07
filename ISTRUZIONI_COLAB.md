# Come eseguire lo script su Google Colab

Segui questi passi — ci vogliono circa 2 minuti di setup e poi 20-25 min per scaricare i dati.

## Step 1 — Apri Colab
Vai su: https://colab.research.google.com → clic su **"Nuovo notebook"**

## Step 2 — Installa le librerie (cella 1)
```python
!pip install yfinance openpyxl requests -q
```

## Step 3 — Carica il file Excel (cella 2)
```python
from google.colab import files
uploaded = files.upload()   # seleziona "ISIN IPO.xlsx" dal tuo PC
import shutil
shutil.copy("ISIN IPO.xlsx", "ISIN_IPO.xlsx")
```

## Step 4 — Esegui lo script (cella 3)
Incolla tutto il contenuto di `scarica_tutto.py` in una nuova cella ed esegui.

## Step 5 — Scarica i risultati (cella 4)
```python
from google.colab import files
files.download("dataset_finale.xlsx")
```

---

## Note importanti
- **EV/EBITDA**: yfinance non ha dati storici precisi per l'anno IPO+2; il valore scaricato è il più recente disponibile (attuale). Se hai ancora accesso anche parziale a Refinitiv, usa `scarica_dati_eikon.py` per quel campo.
- **ESG**: yfinance scarica i punteggi ESG attuali (non storici). Per ESG storico serve Refinitiv/Bloomberg.
- **Tempo stimato**: ~20-25 minuti per 189 aziende.
- **Rate limit OpenFIGI**: Il piano gratuito è 25 richieste/minuto; lo script gestisce automaticamente i retry.
