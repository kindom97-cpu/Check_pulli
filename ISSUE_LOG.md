# FlowCheck — Issue Log

Documento di tracciatura delle anomalie rilevate, analisi delle cause radice e
relative risoluzioni applicate al tool di confronto CSV AS-IS vs TO-BE.

---

## Versione / Data rilascio

| Versione | Data       | Commit     |
|----------|------------|------------|
| 1.0      | 2026-03-31 | `313c5b7`  |
| 1.1      | 2026-03-31 | `bf337fc`  |
| 1.2      | 2026-03-31 | `c3aed5a`  |
| 1.3      | 2026-03-31 | `d60ade4`  |

---

## Elenco Issue

---

### ISSUE-001 — `ValueError: Cannot index with multidimensional key`

| Campo          | Dettaglio |
|----------------|-----------|
| **Stato**      | Risolto |
| **Severità**   | Alta — blocco completo dell'elaborazione |
| **File**       | `compare_plz3a.py` |
| **Commit fix** | `3f22729`, `8a11865` |

**Descrizione**
Durante la fase di confronto tra file CSV veniva sollevato un `ValueError` che
impediva la generazione di qualunque output Excel.

**Causa radice**
I file CSV contenevano colonne con nomi duplicati. L'operazione
`df.apply(lambda c: c.str.strip().ne("").any())` restituiva una `Series` con
indice duplicato; la successiva indicizzazione `df.loc[:, series]` (e in
seguito anche `df.loc[:, series.values]`) interpretava l'array risultante come
multidimensionale, sollevando l'eccezione.

**Soluzione**
1. Aggiunta funzione `_dedup_columns()` che rinomina i duplicati in `COL`,
   `COL_1`, `COL_2`, … prima di qualsiasi operazione.
2. Sostituzione dell'indicizzazione con una list comprehension che non usa
   `.loc`:
   ```python
   non_empty = [c for c in df.columns if df[c].str.strip().ne("").any()]
   df = df[non_empty] if non_empty else df.iloc[:, :0]
   ```

---

### ISSUE-002 — `UnicodeEncodeError` su console Windows

| Campo          | Dettaglio |
|----------------|-----------|
| **Stato**      | Risolto |
| **Severità**   | Media — errore visivo, non bloccante per l'output |
| **File**       | `compare_plz3a.py` |
| **Commit fix** | `f1d02b0` |

**Descrizione**
Lanciando lo script su console Windows con encoding `cp1252`, i caratteri
Unicode usati come simboli di stato (`✓`, `✗`, `⚠`, `×`) causavano un
`UnicodeEncodeError` a runtime.

**Causa radice**
Il codec `cp1252` (default Windows) non include i codepoint U+2713, U+2717,
U+26A0, U+00D7.

**Soluzione**
Sostituzione dei simboli Unicode con equivalenti ASCII puri:

| Simbolo originale | Sostituto ASCII |
|-------------------|-----------------|
| `✓`               | `[OK]`          |
| `✗`               | `[ERRORE]`      |
| `⚠`               | `[ATTENZIONE]`  |
| `×`               | `x`             |

---

### ISSUE-003 — Differenza strutturale falsa (colonne "SOLO TO-BE")

| Campo          | Dettaglio |
|----------------|-----------|
| **Stato**      | Risolto |
| **Severità**   | Alta — risultato errato presentato al cliente |
| **File**       | `compare_plz3a.py` → `flowcheck_engine.py` |
| **Esempio**    | File `AVTBXREG`: 21 colonne segnalate come "SOLO TO-BE" |
| **Commit fix** | `0305fbc` |

**Descrizione**
Alcuni file venivano segnalati come strutturalmente diversi (colonne presenti
solo nel TO-BE) nonostante l'intestazione fosse identica in entrambi i file.

**Causa radice**
La logica di pulizia post-lettura rimuoveva le colonne che risultavano
interamente vuote nel DataFrame corrente:
```python
non_empty = [c for c in df.columns if df[c].str.strip().ne("").any()]
df = df[non_empty] if non_empty else df.iloc[:, :0]
```
Nel file AS-IS alcune colonne nominali erano vuote (nessun dato), quindi
venivano eliminate (55 colonne mantenute). Nel TO-BE le stesse colonne avevano
dati (76 colonne mantenute). Il confronto rilevava 21 colonne "solo TO-BE"
pur essendo la struttura identica.

**Soluzione**
Rimozione solo delle colonne con **nome vuoto** (artefatti del separatore
finale), mai per contenuto:
```python
df = df[[c for c in df.columns if c != ""]]
```

---

### ISSUE-004 — Nomi file con desinenza comune sovrascrivono i report

| Campo          | Dettaglio |
|----------------|-----------|
| **Stato**      | Risolto |
| **Severità**   | Alta — Excel sovrascritta, dati persi |
| **File**       | `flowcheck_engine.py` — funzione `_stem_key()` |
| **Esempio**    | `ABK001FW_DANNI`, `ABBINAMENTO_PLZ_DANNI`, `STORNI_DANNI` → chiave `DANNI` |
| **Commit fix** | `c3aed5a` |

**Descrizione**
Più file con nomi diversi ma stessa desinenza (es. `_DANNI`) venivano abbinati
alla stessa chiave di matching, causando la sovrascrittura dei report Excel
precedenti.

**Causa radice**
`_stem_key()` estraeva solo l'**ultimo token** del nome file dopo aver rimosso
timestamp e separatori. File come `ABK001FW_DANNI`, `ABBINAMENTO_PLZ_DANNI` e
`STORNI_DANNI` producevano tutti la chiave `DANNI`.

**Soluzione**
La funzione ora **rimuove i prefissi tecnici noti** (`DW`, `D`/`M`, `PLZxx`)
e conserva l'intera parte rimanente come chiave:

| Nome file                                  | Chiave precedente | Chiave corretta       |
|--------------------------------------------|-------------------|-----------------------|
| `DW.D.PLZ3A.ABBINAMENTO_PLZ_DANNI.*.csv`   | `DANNI`           | `ABBINAMENTO_PLZ_DANNI` |
| `DW.D.PLZAA.ABK001FW_DANNI.csv`            | `DANNI`           | `ABK001FW_DANNI`      |
| `DW.D.PLZBA.STORNI_DANNI.csv`              | `DANNI`           | `STORNI_DANNI`        |
| `DW.D.PLZHA.AVTBCODI.*.csv`                | `AVTBCODI`        | `AVTBCODI` (invariato)|

---

### ISSUE-005 — Falsi positivi da spazi e caratteri non-breaking

| Campo          | Dettaglio |
|----------------|-----------|
| **Stato**      | Risolto |
| **Severità**   | Media — differenze segnalate che non esistono semanticamente |
| **File**       | `flowcheck_engine.py` |
| **Esempio**    | File `AVTBCODI`: campi uguali marcati come diversi per spazi residui |
| **Commit fix** | `bf337fc` |

**Descrizione**
Campi identici nel contenuto venivano marcati come diversi nel report Excel
a causa di spazi iniziali/finali o caratteri di spaziatura non standard.

**Causa radice**
Il metodo `str.strip()` di pandas rimuove gli spazi ASCII (U+0020) ma **non**
lo spazio non-breaking (U+00A0, `\xa0`) presente in alcuni export da sistemi
aziendali. Il confronto `!=` avveniva poi sui valori non normalizzati.

**Soluzione**
Introdotta `_clean_str_series()` applicata in due fasi:

1. **In lettura** (`read_csv`, `read_csv_from_zip`): strip + rimozione `\xa0` +
   collasso spazi multipli interni in uno singolo.
2. **In fase di confronto** (`compare_dataframes`): normalizzazione prima di
   ogni `!=`, mantenendo i valori originali nei fogli diff dell'Excel.

```python
def _clean_str_series(s):
    return (s.str.strip()
             .str.replace("\xa0", " ", regex=False)
             .str.replace(r"[ \t]+", " ", regex=True)
             .str.strip())
```

---

### ISSUE-006 — Separatori diversi tra AS-IS e TO-BE / colonne artefatto da separatore multi-char

| Campo          | Dettaglio |
|----------------|-----------|
| **Stato**      | Risolto |
| **Severità**   | Alta — lettura errata del file, confronto impossibile |
| **File**       | `flowcheck_engine.py` |
| **Esempio**    | `DW.D.PLZDA.MYQUOTE`: AS-IS con `;£`, TO-BE con `;#` |
| **Commit fix** | `d60ade4` |

**Descrizione**
Due sotto-problemi distinti:

**A) Separatore non rilevato per file specifico dentro ZIP**
`read_csv_from_zip` chiamava `detect_separator(zip_path)` che leggeva sempre
il **primo** CSV nello ZIP per determinare il separatore, usando poi quel valore
per tutti i file successivi. Se i file all'interno dello ZIP (o tra ZIP diversi)
usavano separatori diversi, alcuni venivano letti in modo errato.

**B) Colonna artefatto con nome `#` (o `£`, `|` ecc.)**
Con separatori multi-char tipo `;#`, le righe terminavano con il carattere `#`
isolato. Il vecchio filtro `c != ""` rimuoveva solo le colonne di nome vuoto,
lasciando nel DataFrame una colonna fittizia denominata `#` (o `£`, `|` ecc.)
che generava false differenze strutturali.

**Soluzione A — Detection per-file**
`_read_first_lines()` ora accetta il parametro `zip_entry` per leggere
esattamente il CSV da rilevare. `detect_separator()` espone lo stesso parametro.
`read_csv_from_zip()` lo passa automaticamente:
```python
sep = detect_separator(zip_path, zip_entry=csv_name)
```

**Soluzione B — Detection dinamica del separatore composto**
Aggiunta `_build_sep_candidates()` che analizza il contenuto del file e rileva
automaticamente pattern `;X` (con X qualsiasi carattere speciale), senza
richiedere aggiornamenti alla lista fissa `SEP_CANDIDATES`. Funziona per `;#`,
`;£`, `;|`, `;@` e qualunque variante futura.

**Soluzione C — Filtro colonne artefatto esteso**
Sostituito il filtro `c != ""` con `_is_artifact_col()`:
```python
def _is_artifact_col(col_name):
    stripped = col_name.strip()
    if not stripped:
        return True                   # nome vuoto
    if not re.search(r"\w", stripped):
        return True                   # solo caratteri speciali: #, £, |, ;#, …
    return False
```
Colonne come `#`, `£`, `|`, `;#`, `;£` vengono scartate; colonne come `#ID`
o `_COL` (che contengono almeno un carattere word) vengono mantenute.

---

## Riepilogo issues per categoria

| Categoria                  | Issue          |
|----------------------------|----------------|
| Errore bloccante runtime   | ISSUE-001      |
| Encoding / compatibilità   | ISSUE-002      |
| Risultato errato           | ISSUE-003, ISSUE-004, ISSUE-005, ISSUE-006 |
| Parsing file               | ISSUE-001, ISSUE-006 |

---

*Documento generato il 2026-03-31 — FlowCheck v1.3*
