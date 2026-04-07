"""
Script per scaricare EV/EBITDA da Refinitiv Eikon
per le aziende nel file Excel con ISIN e date IPO.

PREREQUISITI:
  1. Eikon / Workspace aperto e loggato
  2. pip install eikon pandas openpyxl

COME USARE:
  1. Inserisci il tuo App Key sotto (APP_KEY)
  2. Verifica che il path del file Excel sia corretto (INPUT_FILE)
  3. Esegui: python scarica_dati_eikon.py
"""

import eikon as ek
import pandas as pd
from datetime import datetime, timedelta
import time

# ─────────────────────────────────────────────
# CONFIGURAZIONE — modifica questi valori
# ─────────────────────────────────────────────
APP_KEY    = "INSERISCI_QUI_IL_TUO_APP_KEY"   # Eikon: Settings → API → App Key
INPUT_FILE = "claude_lavoro_ipo.xlsx"          # Path del tuo file Excel
OUTPUT_FILE = "risultati_ev_ebitda.xlsx"       # File di output
GIORNI_FINESTRA = 10                           # Giorni post-IPO in cui cercare il dato
# ─────────────────────────────────────────────

ek.set_app_key(APP_KEY)

def isin_to_ric(isin):
    """Converte ISIN in RIC tramite Eikon."""
    try:
        data, _ = ek.get_data(isin, ["TR.RIC"])
        ric = data["RIC"].iloc[0]
        if pd.notna(ric) and ric != "":
            return str(ric).strip()
    except Exception:
        pass
    return None

def get_ev_ebitda(ric, ipo_date, giorni=10):
    """
    Scarica EV/EBITDA cercando il primo valore disponibile
    nella finestra [ipo_date, ipo_date + giorni].
    """
    try:
        start = ipo_date.strftime("%Y-%m-%d")
        end   = (ipo_date + timedelta(days=giorni)).strftime("%Y-%m-%d")
        data, _ = ek.get_data(
            ric,
            ["TR.EVToEBITDA"],
            {"SDate": start, "EDate": end, "Frq": "D"}
        )
        # Prendi il primo valore non nullo
        valori = data["Enterprise Value To EBITDA"].dropna()
        if not valori.empty:
            return round(float(valori.iloc[0]), 4)
    except Exception:
        pass
    return None

# ─────────────────────────────────────────────
# CARICA IL FILE EXCEL
# ─────────────────────────────────────────────
print("Carico il file Excel...")
df = pd.read_excel(INPUT_FILE, sheet_name="Foglio1")
df.columns = ["ISIN", "IPO_DATE"]
df["ISIN"] = df["ISIN"].astype(str).str.strip()
df["IPO_DATE"] = pd.to_datetime(df["IPO_DATE"])
print(f"  → {len(df)} aziende trovate\n")

# ─────────────────────────────────────────────
# STEP 1: CONVERTI ISIN → RIC
# ─────────────────────────────────────────────
print("STEP 1: Conversione ISIN → RIC")
ric_list = []
for i, row in df.iterrows():
    isin = row["ISIN"]
    ric = isin_to_ric(isin)
    ric_list.append(ric)
    stato = ric if ric else "❌ non trovato"
    print(f"  [{i+1}/{len(df)}] {isin} → {stato}")
    time.sleep(0.3)  # evita rate limit

df["RIC"] = ric_list
n_ric_ok  = df["RIC"].notna().sum()
n_ric_no  = df["RIC"].isna().sum()
print(f"\n  ✅ RIC trovati: {n_ric_ok} | ❌ Non trovati: {n_ric_no}\n")

# ─────────────────────────────────────────────
# STEP 2: SCARICA EV/EBITDA
# ─────────────────────────────────────────────
print(f"STEP 2: Download EV/EBITDA (finestra: +{GIORNI_FINESTRA} giorni dall'IPO)")
ev_list = []
for i, row in df.iterrows():
    if pd.isna(row["RIC"]):
        ev_list.append(None)
        print(f"  [{i+1}/{len(df)}] {row['ISIN']} → saltato (nessun RIC)")
        continue

    val = get_ev_ebitda(row["RIC"], row["IPO_DATE"], GIORNI_FINESTRA)
    ev_list.append(val)
    stato = f"{val:.2f}x" if val else "⚠️  NULL"
    print(f"  [{i+1}/{len(df)}] {row['RIC']} ({row['IPO_DATE'].date()}) → {stato}")
    time.sleep(0.3)

df["EV_EBITDA"] = ev_list

# ─────────────────────────────────────────────
# STEP 3: SALVA IL FILE DI OUTPUT
# ─────────────────────────────────────────────
n_ok   = df["EV_EBITDA"].notna().sum()
n_null = df["EV_EBITDA"].isna().sum()

print(f"\nRISULTATI:")
print(f"  ✅ Valori scaricati: {n_ok}/{len(df)}")
print(f"  ⚠️  NULL / non disponibili: {n_null}/{len(df)}")

# Foglio principale
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Dati", index=False)

    # Foglio separato con i NULL per revisione manuale
    df_null = df[df["EV_EBITDA"].isna()].copy()
    if not df_null.empty:
        df_null.to_excel(writer, sheet_name="Da_verificare", index=False)

print(f"\n✅ File salvato: {OUTPUT_FILE}")
if n_null > 0:
    print(f"   → Foglio 'Da_verificare': {n_null} aziende con dato mancante")
    print("   Prova ad aumentare GIORNI_FINESTRA (es. 30) per recuperarne alcune.")
