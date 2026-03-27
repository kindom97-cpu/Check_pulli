#!/usr/bin/env python3
"""
Confronto AS-IS vs TO-BE per file PLZ3A*
=========================================
Cerca due cartelle PLZ3A*:
  - Senza date nel nome  → versione TO-BE
  - Con date nel nome    → versione AS-IS

Join riga-per-riga sulla chiave univoca auto-rilevata, poi confronto
colonna per colonna. Per ogni coppia di file produce un Excel con:

  - Sheet "RIEPILOGO"             : overview tutte le coppie di file
  - Sheet "STRUTTURA_<nome>"      : nomi colonne, tipi, presenza AS-IS/TO-BE
  - Sheet "SINTESI_COLONNE_<nome>": per ogni colonna quante righe differiscono
  - Sheet "DIFF_<nome>"           : vista wide — chiave + colonne AS-IS/TO-BE
                                    affiancate, celle gialle dove i valori divergono
  - Sheet "SOLO_ASIS_<nome>"      : righe presenti solo in AS-IS
  - Sheet "SOLO_TOBE_<nome>"      : righe presenti solo in TO-BE

Uso:
  python3 compare_plz3a.py                          # auto-detect nella CWD
  python3 compare_plz3a.py --base-dir /path/to/dir
  python3 compare_plz3a.py --asis /path/asis --tobe /path/tobe
  python3 compare_plz3a.py --output risultati.xlsx
"""

import argparse
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─── Palette colori ──────────────────────────────────────────────────────────
COLOR_HEADER  = "1F4E79"   # blu scuro   → testo bianco
COLOR_SUBHDR  = "2E75B6"   # blu medio   → testo bianco
COLOR_KEY_HDR = "4472C4"   # blu chiaro  → intestazione colonna chiave
COLOR_DIFF    = "FFE699"   # giallo      → valore discrepante
COLOR_EQUAL   = "FFFFFF"   # bianco      → valore uguale
COLOR_ONLY_AS = "F4CCCC"   # rosso tenue → solo AS-IS
COLOR_ONLY_TO = "D9EAD3"   # verde tenue → solo TO-BE
COLOR_OK      = "E2EFDA"   # verde chiaro → riga riepilogo OK
COLOR_KO      = "FCE4D6"   # arancio      → riga riepilogo con diff
COLOR_WARN    = "FFF2CC"   # giallo chiaro → warning (tipo diverso)


# ─── Helpers stile ───────────────────────────────────────────────────────────

def _fill(c):
    return PatternFill("solid", fgColor=c)

def _font(bold=False, color="000000", size=11):
    return Font(bold=bold, color=color, size=size)

def _border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def _hdr(ws, row, bg, fg="FFFFFF"):
    for cell in ws[row]:
        if cell.value is not None:
            cell.fill      = _fill(bg)
            cell.font      = _font(bold=True, color=fg)
            cell.border    = _border()
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def _row(ws, row, bg=None):
    for cell in ws[row]:
        cell.border    = _border()
        cell.alignment = Alignment(vertical="center")
        if bg:
            cell.fill = _fill(bg)

def _autofit(ws, mn=8, mx=50):
    for col in ws.columns:
        w = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(w + 2, mn), mx)

def _title(ws, text, n_cols):
    ws.append([text])
    if n_cols > 1:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    ws["A1"].fill      = _fill(COLOR_HEADER)
    ws["A1"].font      = _font(bold=True, color="FFFFFF", size=12)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22


# ─── Rilevamento cartelle ─────────────────────────────────────────────────────

TS_FOLDER = re.compile(r"\.\d{14}\.\d{14}$")

def _has_ts(name):
    return bool(TS_FOLDER.search(name))

def find_plz3a_folders(base_dir):
    base = Path(base_dir)
    cands = sorted(p for p in base.iterdir()
                   if p.is_dir() and "PLZ3A" in p.name.upper())
    asis = tobe = None
    for p in cands:
        if _has_ts(p.name):
            asis = p
        else:
            tobe = p
    return asis, tobe


# ─── Matching file ────────────────────────────────────────────────────────────

def _logical(filename):
    """Rimuove i timestamp dal nome file per ottenere il nome logico."""
    stem = Path(filename).stem
    cleaned = re.sub(r"\.\d{14}.*$", "", stem)
    return cleaned.rstrip(".")

def match_files(asis_dir, tobe_dir):
    def _csvs(d):
        if d is None or not d.exists():
            return {}
        return {_logical(f.name): f for f in d.iterdir() if f.suffix.lower() == ".csv"}

    a_map = _csvs(asis_dir)
    t_map = _csvs(tobe_dir)
    keys  = sorted(set(a_map) | set(t_map))
    return [{"logical_name": k,
             "short_name":   k.split(".")[-1] if "." in k else k,
             "asis_path":    a_map.get(k),
             "tobe_path":    t_map.get(k)} for k in keys]


# ─── Lettura CSV ──────────────────────────────────────────────────────────────

def _has_header(filepath):
    with open(filepath, "r", encoding="utf-8", errors="replace") as fh:
        line = fh.readline().rstrip("\r\n")
    fields = [f.strip() for f in line.split(";") if f.strip()]
    if not fields:
        return False
    return sum(1 for f in fields if re.match(r"^[A-Za-z_][A-Za-z0-9_\s]*$", f)) >= max(1, len(fields) * 0.5)

def read_csv(filepath):
    hdr = 0 if _has_header(filepath) else None
    df  = pd.read_csv(filepath, sep=";", header=hdr, dtype=str,
                      encoding="utf-8", encoding_errors="replace",
                      skipinitialspace=True, keep_default_na=False)
    # rimuovi colonne trailing vuote
    df  = df.loc[:, df.apply(lambda c: c.str.strip().ne("").any())]
    df  = df.apply(lambda c: c.str.strip() if c.dtype == object else c)
    df.columns = [str(c) for c in df.columns]
    return df


# ─── Inferenza tipo ───────────────────────────────────────────────────────────

def _infer_type(series):
    s = series.dropna().head(200)
    if s.empty:
        return "VUOTO"
    num = pd.to_numeric(s, errors="coerce").notna().mean()
    if num > 0.9:
        return "DECIMALE" if s.str.contains(r"\.", na=False).any() else "INTERO"
    try:
        pd.to_datetime(s, dayfirst=True, errors="raise")
        return "DATA"
    except Exception:
        pass
    return "TESTO"


# ─── Rilevamento chiave ───────────────────────────────────────────────────────

def detect_key(df_a, df_b):
    """
    Cerca la colonna (o combinazione minima) che sia univoca in entrambi i df.
    Ritorna lista di nomi colonna da usare come chiave.
    """
    common = [c for c in df_a.columns if c in df_b.columns]
    if not common:
        return []

    # 1) singola colonna con tutti valori unici in entrambi
    for c in common:
        if df_a[c].nunique() == len(df_a) and df_b[c].nunique() == len(df_b):
            return [c]

    # 2) combinazione di 2 colonne
    for i, c1 in enumerate(common):
        for c2 in common[i+1:]:
            if (df_a[[c1,c2]].drop_duplicates().shape[0] == len(df_a) and
                df_b[[c1,c2]].drop_duplicates().shape[0] == len(df_b)):
                return [c1, c2]

    # 3) nessuna chiave trovata → confronto posizionale
    return []


# ─── Differenza numerica ──────────────────────────────────────────────────────

def _num_diff(a, b):
    try:
        return float(str(a).replace(",", ".")) - float(str(b).replace(",", "."))
    except (ValueError, TypeError):
        return ""


# ─── Analisi struttura colonne ────────────────────────────────────────────────

def col_structure(df_a, df_b):
    ca, cb = list(df_a.columns), list(df_b.columns)
    all_c  = sorted(set(ca) | set(cb), key=lambda x: (ca.index(x) if x in ca else 9999))
    rows   = []
    for c in all_c:
        in_a = c in ca
        in_b = c in cb
        ta   = _infer_type(df_a[c]) if in_a else "—"
        tb   = _infer_type(df_b[c]) if in_b else "—"
        stat = "OK" if (in_a and in_b) else ("SOLO AS-IS" if in_a else "SOLO TO-BE")
        rows.append({
            "COLONNA":       c,
            "IN AS-IS":      "Sì" if in_a else "No",
            "IN TO-BE":      "Sì" if in_b else "No",
            "POS AS-IS":     ca.index(c) + 1 if in_a else "—",
            "POS TO-BE":     cb.index(c) + 1 if in_b else "—",
            "TIPO AS-IS":    ta,
            "TIPO TO-BE":    tb,
            "TIPO COERENTE": "Sì" if (in_a and in_b and ta == tb) else ("—" if not (in_a and in_b) else "No"),
            "STATUS":        stat,
        })
    return pd.DataFrame(rows)


# ─── Confronto righe (join su chiave) ────────────────────────────────────────

def compare_rows(df_a, df_b, key_cols):
    """
    Ritorna:
      - merged_both   : DataFrame righe comuni (join su key_cols), con suffissi __AS e __TO
      - only_asis     : righe solo in AS-IS
      - only_tobe     : righe solo in TO-BE
    """
    common = [c for c in df_a.columns if c in df_b.columns]

    if key_cols:
        merged = pd.merge(
            df_a[common].reset_index(drop=True),
            df_b[common].reset_index(drop=True),
            on=key_cols, how="outer",
            suffixes=("__AS", "__TO"),
            indicator=True
        )
        only_as   = (merged[merged["_merge"] == "left_only"]
                     .drop(columns=["_merge"])
                     .rename(columns=lambda c: c.replace("__AS","").replace("__TO","")))
        only_to   = (merged[merged["_merge"] == "right_only"]
                     .drop(columns=["_merge"])
                     .rename(columns=lambda c: c.replace("__AS","").replace("__TO","")))
        both      = merged[merged["_merge"] == "both"].drop(columns=["_merge"]).reset_index(drop=True)
    else:
        # posizionale
        n       = min(len(df_a), len(df_b))
        only_as = df_a.iloc[n:].reset_index(drop=True) if len(df_a) > n else pd.DataFrame(columns=df_a.columns)
        only_to = df_b.iloc[n:].reset_index(drop=True) if len(df_b) > n else pd.DataFrame(columns=df_b.columns)
        left    = df_a[common].iloc[:n].reset_index(drop=True).add_suffix("__AS")
        right   = df_b[common].iloc[:n].reset_index(drop=True).add_suffix("__TO")
        both    = pd.concat([left, right], axis=1)

    return both, only_as, only_to


def sintesi_colonne(both, key_cols, value_cols):
    """
    Per ogni colonna di valore: conta righe uguali e diverse.
    """
    rows = []
    for c in value_cols:
        col_as = f"{c}__AS"
        col_to = f"{c}__TO"
        if col_as not in both.columns or col_to not in both.columns:
            continue
        eq    = (both[col_as].fillna("") == both[col_to].fillna("")).sum()
        diff  = len(both) - eq
        rows.append({
            "COLONNA":          c,
            "CHIAVE":           "Sì" if c in key_cols else "No",
            "RIGHE UGUALI":     int(eq),
            "RIGHE DIVERSE":    int(diff),
            "TOT RIGHE":        len(both),
            "% DIVERSE":        f"{diff/len(both)*100:.1f}%" if len(both) else "0%",
            "STATO":            "OK" if diff == 0 else "DIFFERENZE",
        })
    return pd.DataFrame(rows)


# ─── Scrittura Excel ──────────────────────────────────────────────────────────

def build_excel(pairs, output_path):
    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ── RIEPILOGO ────────────────────────────────────────────────────────────
    ws_r = wb.create_sheet("RIEPILOGO")
    n_hdr_cols = 10
    _title(ws_r, "RIEPILOGO CONFRONTO AS-IS vs TO-BE", n_hdr_cols)
    ws_r.append(["FILE LOGICO", "FILE AS-IS", "FILE TO-BE",
                 "RIGHE AS-IS", "RIGHE TO-BE", "DELTA RIGHE",
                 "COLONNE SOLO AS-IS", "COLONNE SOLO TO-BE",
                 "COLONNE CON DIFF", "RIGHE ABBINATE CON DIFF"])
    _hdr(ws_r, 2, COLOR_SUBHDR)

    all_pair_data = []

    for pair in pairs:
        a_path = pair["asis_path"]
        t_path = pair["tobe_path"]

        df_a = read_csv(a_path) if (a_path and a_path.stat().st_size > 0) else pd.DataFrame()
        df_b = read_csv(t_path) if (t_path and t_path.stat().st_size > 0) else pd.DataFrame()

        n_a, n_b = len(df_a), len(df_b)

        # struttura colonne
        strut = col_structure(df_a, df_b) if (not df_a.empty or not df_b.empty) else pd.DataFrame()
        only_a_cols = strut[strut["STATUS"]=="SOLO AS-IS"]["COLONNA"].tolist() if not strut.empty else []
        only_b_cols = strut[strut["STATUS"]=="SOLO TO-BE"]["COLONNA"].tolist() if not strut.empty else []

        # confronto righe
        if not df_a.empty and not df_b.empty:
            key_cols   = detect_key(df_a, df_b)
            common     = [c for c in df_a.columns if c in df_b.columns]
            value_cols = [c for c in common if c not in key_cols]
            both, only_as_rows, only_to_rows = compare_rows(df_a, df_b, key_cols)
            sint       = sintesi_colonne(both, key_cols, value_cols)
            # colonne con almeno una differenza
            cols_w_diff = sint[sint["STATO"]=="DIFFERENZE"]["COLONNA"].tolist() if not sint.empty else []
            # righe abbinate con almeno una differenza
            if value_cols and not both.empty:
                mask = pd.Series([False] * len(both))
                for c in value_cols:
                    if f"{c}__AS" in both.columns and f"{c}__TO" in both.columns:
                        mask = mask | (both[f"{c}__AS"].fillna("") != both[f"{c}__TO"].fillna(""))
                n_diff_rows = int(mask.sum())
            else:
                n_diff_rows = 0
        else:
            key_cols    = []
            value_cols  = []
            both        = pd.DataFrame()
            only_as_rows = df_a.copy()
            only_to_rows = df_b.copy()
            sint         = pd.DataFrame()
            cols_w_diff  = []
            n_diff_rows  = 0

        all_pair_data.append({**pair,
            "df_a": df_a, "df_b": df_b,
            "strut": strut, "sint": sint,
            "key_cols": key_cols, "value_cols": value_cols,
            "both": both,
            "only_as_rows": only_as_rows, "only_to_rows": only_to_rows,
            "only_a_cols": only_a_cols, "only_b_cols": only_b_cols,
            "cols_w_diff": cols_w_diff, "n_diff_rows": n_diff_rows,
        })

        ok = (not only_a_cols and not only_b_cols and n_diff_rows == 0 and n_a == n_b)
        ws_r.append([
            pair["logical_name"],
            a_path.name if a_path else "— MANCANTE —",
            t_path.name if t_path else "— MANCANTE —",
            n_a, n_b, n_a - n_b,
            ", ".join(only_a_cols) or "—",
            ", ".join(only_b_cols) or "—",
            ", ".join(cols_w_diff) or "—",
            n_diff_rows,
        ])
        _row(ws_r, ws_r.max_row, COLOR_OK if ok else COLOR_KO)

    _autofit(ws_r)
    ws_r.freeze_panes = "A3"

    # ── Sheet per ogni coppia ─────────────────────────────────────────────────
    for pd_ in all_pair_data:
        sname      = pd_["short_name"][:19]   # max 31 - len("SINTESI_COL_") = 19
        strut      = pd_["strut"]
        sint       = pd_["sint"]
        both       = pd_["both"]
        only_as    = pd_["only_as_rows"]
        only_to    = pd_["only_to_rows"]
        key_cols   = pd_["key_cols"]
        value_cols = pd_["value_cols"]
        lname      = pd_["logical_name"]

        # ── STRUTTURA COLONNE ────────────────────────────────────────────────
        ws_s = wb.create_sheet(f"STRUTTURA_{sname}")
        _title(ws_s, f"Struttura colonne – {lname}", 9)
        if strut.empty:
            ws_s.append(["(nessun dato)"])
        else:
            ws_s.append(list(strut.columns))
            _hdr(ws_s, 2, COLOR_SUBHDR)
            for _, r in strut.iterrows():
                ws_s.append(list(r))
                rn = ws_s.max_row
                _row(ws_s, rn)
                st = str(r.get("STATUS",""))
                tc = str(r.get("TIPO COERENTE","Sì"))
                if st == "SOLO AS-IS":
                    _row(ws_s, rn, COLOR_ONLY_AS)
                elif st == "SOLO TO-BE":
                    _row(ws_s, rn, COLOR_ONLY_TO)
                elif tc == "No":
                    _row(ws_s, rn, COLOR_WARN)
        _autofit(ws_s)
        ws_s.freeze_panes = "A3"

        # ── SINTESI COLONNE ──────────────────────────────────────────────────
        ws_sc = wb.create_sheet(f"SINTESI_COL_{sname}")
        _title(ws_sc, f"Sintesi differenze per colonna – {lname}", 7)
        if sint.empty:
            ws_sc.append(["(nessun dato comparabile)"])
        else:
            ws_sc.append(list(sint.columns))
            _hdr(ws_sc, 2, COLOR_SUBHDR)
            for _, r in sint.iterrows():
                ws_sc.append(list(r))
                rn = ws_sc.max_row
                stato = str(r.get("STATO",""))
                _row(ws_sc, rn, COLOR_KO if stato == "DIFFERENZE" else COLOR_OK)
        _autofit(ws_sc)
        ws_sc.freeze_panes = "A3"

        # ── DIFF WIDE ────────────────────────────────────────────────────────
        # Intestazioni: KEY | COL1_AS-IS | COL1_TO-BE | DIFF_COL1 | COL2_AS-IS | ...
        ws_d = wb.create_sheet(f"DIFF_{sname}")

        if both.empty or not value_cols:
            _title(ws_d, f"Differenze riga per riga – {lname}", 1)
            ws_d.append(["(nessuna riga comparabile)"])
        else:
            # filtra solo righe con almeno una differenza
            mask = pd.Series([False] * len(both))
            for c in value_cols:
                ca, cb = f"{c}__AS", f"{c}__TO"
                if ca in both.columns and cb in both.columns:
                    mask = mask | (both[ca].fillna("") != both[cb].fillna(""))
            diff_rows_df = both[mask].reset_index(drop=True)

            # costruisci header
            headers = []
            for k in key_cols:
                headers.append(k)
            for c in value_cols:
                headers += [f"{c} [AS-IS]", f"{c} [TO-BE]", f"DIFF {c}"]

            n_cols = len(headers)
            _title(ws_d, f"Differenze riga per riga – {lname} ({len(diff_rows_df)} righe con diff)", n_cols)
            ws_d.append(headers)
            _hdr(ws_d, 2, COLOR_SUBHDR)

            # evidenzia in blu le colonne chiave nell'header
            for idx, k in enumerate(key_cols):
                ws_d.cell(row=2, column=idx+1).fill = _fill(COLOR_KEY_HDR)

            for _, r in diff_rows_df.iterrows():
                row_vals = []
                # chiave
                for k in key_cols:
                    row_vals.append(r.get(k, ""))
                # valori
                for c in value_cols:
                    va = r.get(f"{c}__AS", "")
                    vb = r.get(f"{c}__TO", "")
                    row_vals.append(va)
                    row_vals.append(vb)
                    row_vals.append(_num_diff(va, vb))

                ws_d.append(row_vals)
                rn = ws_d.max_row
                _row(ws_d, rn)

                # evidenzia celle AS-IS / TO-BE dove i valori differiscono
                col_offset = len(key_cols) + 1
                for c in value_cols:
                    va = str(r.get(f"{c}__AS", "")).strip()
                    vb = str(r.get(f"{c}__TO", "")).strip()
                    if va != vb:
                        ws_d.cell(rn, col_offset).fill     = _fill(COLOR_DIFF)
                        ws_d.cell(rn, col_offset + 1).fill = _fill(COLOR_DIFF)
                    col_offset += 3

        _autofit(ws_d)
        ws_d.freeze_panes = f"{get_column_letter(len(key_cols)+1)}3"

        # ── SOLO AS-IS ───────────────────────────────────────────────────────
        ws_a = wb.create_sheet(f"SOLO_ASIS_{sname}")
        nc = max(len(only_as.columns), 1)
        _title(ws_a, f"Righe solo in AS-IS – {lname}", nc)
        if only_as.empty:
            ws_a.append(["(nessuna riga esclusiva)"])
        else:
            ws_a.append(list(only_as.columns))
            _hdr(ws_a, 2, COLOR_SUBHDR)
            for _, r in only_as.iterrows():
                ws_a.append(list(r))
                _row(ws_a, ws_a.max_row, COLOR_ONLY_AS)
        _autofit(ws_a)

        # ── SOLO TO-BE ───────────────────────────────────────────────────────
        ws_t = wb.create_sheet(f"SOLO_TOBE_{sname}")
        nc = max(len(only_to.columns), 1)
        _title(ws_t, f"Righe solo in TO-BE – {lname}", nc)
        if only_to.empty:
            ws_t.append(["(nessuna riga esclusiva)"])
        else:
            ws_t.append(list(only_to.columns))
            _hdr(ws_t, 2, COLOR_SUBHDR)
            for _, r in only_to.iterrows():
                ws_t.append(list(r))
                _row(ws_t, ws_t.max_row, COLOR_ONLY_TO)
        _autofit(ws_t)

    wb.save(output_path)
    print(f"\n✓ Excel salvato in: {output_path}")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Confronto AS-IS vs TO-BE – PLZ3A*")
    parser.add_argument("--base-dir", default=".")
    parser.add_argument("--asis",     default=None)
    parser.add_argument("--tobe",     default=None)
    parser.add_argument("--output",   default=None)
    args = parser.parse_args()

    if args.asis or args.tobe:
        asis_path = Path(args.asis) if args.asis else None
        tobe_path = Path(args.tobe) if args.tobe else None
    else:
        asis_path, tobe_path = find_plz3a_folders(args.base_dir)

    print("=" * 60)
    print("  CONFRONTO AS-IS vs TO-BE – PLZ3A")
    print("=" * 60)
    print(f"  AS-IS : {asis_path or '⚠  NON TROVATA'}")
    print(f"  TO-BE : {tobe_path or '⚠  NON TROVATA'}")

    if asis_path is None and tobe_path is None:
        print("\n✗ Nessuna cartella PLZ3A trovata.")
        sys.exit(1)

    pairs = match_files(asis_path, tobe_path)
    if not pairs:
        print("\n✗ Nessun file CSV trovato.")
        sys.exit(1)

    print(f"\n  File CSV rilevati: {len(pairs)}")
    for p in pairs:
        a = p["asis_path"].name if p["asis_path"] else "MANCANTE"
        t = p["tobe_path"].name if p["tobe_path"] else "MANCANTE"
        print(f"    [{p['short_name']}]  AS-IS: {a}  |  TO-BE: {t}")

    ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = args.output or os.path.join(args.base_dir, f"confronto_PLZ3A_{ts}.xlsx")

    print(f"\n  Elaborazione in corso...")
    build_excel(pairs, out)
    print("  Done.")


if __name__ == "__main__":
    main()
