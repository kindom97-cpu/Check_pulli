#!/usr/bin/env python3
"""
Confronto AS-IS vs TO-BE per file PLZ3A*
=========================================
Cerca due cartelle PLZ3A*:
  - Senza date nel nome  → versione TO-BE
  - Con date nel nome    → versione AS-IS

Per ogni coppia di file corrispondenti produce un Excel con:
  - Sheet "RIEPILOGO"        : overview tutte le coppie di file
  - Sheet "COLONNE_<nome>"   : confronto nomi e tipi colonna
  - Sheet "DIFF_<nome>"      : righe con valori discrepanti (valore per valore)
  - Sheet "SOLO_ASIS_<nome>" : righe presenti solo in AS-IS
  - Sheet "SOLO_TOBE_<nome>" : righe presenti solo in TO-BE

Uso:
  python3 compare_plz3a.py                                 # auto-detect nella CWD
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
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ─── Costanti stile Excel ────────────────────────────────────────────────────
COLOR_HEADER   = "1F4E79"   # blu scuro  → testo bianco
COLOR_SUBHDR   = "2E75B6"   # blu medio  → testo bianco
COLOR_DIFF     = "FFE699"   # giallo     → valore discrepante
COLOR_ONLY_AS  = "F4CCCC"   # rosso tenue → solo AS-IS
COLOR_ONLY_TO  = "D9EAD3"   # verde tenue → solo TO-BE
COLOR_OK       = "E2EFDA"   # verde chiaro
COLOR_KO       = "FCE4D6"   # arancio tenue
COLOR_MISSING  = "D9D9D9"   # grigio     → file mancante


# ─── Utility ─────────────────────────────────────────────────────────────────

def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color="000000", size=11) -> Font:
    return Font(bold=bold, color=color, size=size)

def _border() -> Border:
    side = Side(style="thin")
    return Border(left=side, right=side, top=side, bottom=side)

def _style_header_row(ws, row: int, fill_color: str, font_color="FFFFFF"):
    for cell in ws[row]:
        if cell.value is not None:
            cell.fill    = _fill(fill_color)
            cell.font    = _font(bold=True, color=font_color)
            cell.border  = _border()
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def _style_data_row(ws, row: int, fill_color: str | None = None):
    for cell in ws[row]:
        cell.border = _border()
        cell.alignment = Alignment(vertical="center", wrap_text=False)
        if fill_color:
            cell.fill = _fill(fill_color)

def _autofit(ws, min_w=8, max_w=60):
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 2, min_w), max_w)


# ─── Rilevamento cartelle ─────────────────────────────────────────────────────

TS_PATTERN = re.compile(r"\.\d{14}\.\d{14}$")          # due blocchi di 14 cifre
TS_IN_NAME  = re.compile(r"\.\d{8,14}")                 # almeno 8 cifre consecutive


def _has_timestamps(folder_name: str) -> bool:
    return bool(TS_PATTERN.search(folder_name))


def find_plz3a_folders(base_dir: str) -> tuple[Path | None, Path | None]:
    """Restituisce (asis_path, tobe_path) cercando cartelle PLZ3A* sotto base_dir."""
    base = Path(base_dir)
    candidates = sorted(p for p in base.iterdir()
                        if p.is_dir() and "PLZ3A" in p.name.upper())
    asis, tobe = None, None
    for p in candidates:
        if _has_timestamps(p.name):
            asis = p
        else:
            tobe = p
    return asis, tobe


# ─── Matching file tra le due cartelle ───────────────────────────────────────

def _logical_name(filename: str) -> str:
    """
    Estrae il nome logico del file togliendo i timestamp.
    Esempio:
      DW.D.PLZ3A.ABBINAMENTO_PLZ_DANNI.20260326000000.000L.20260326015133.csv
      → DW.D.PLZ3A.ABBINAMENTO_PLZ_DANNI

    La convenzione di naming è:
      <PREFIX>.<CORE_NAME>.<YYYYMMDDHHMMSS>.<SEQ>.<YYYYMMDDHHMMSS>.csv
    oppure (TO-BE, senza date):
      <PREFIX>.<CORE_NAME>.csv
    """
    stem = Path(filename).stem          # rimuove .csv

    # Pattern: blocco di 14 cifre (timestamp), poi eventuali ".<seq>" e altro timestamp
    # Rimuove tutto a partire dal primo blocco di 14 cifre
    cleaned = re.sub(r"\.\d{14}.*$", "", stem)

    # Se non c'era un blocco di 14 cifre (TO-BE senza date), il stem è già il nome logico
    return cleaned.rstrip(".")


def match_files(asis_dir: Path | None, tobe_dir: Path | None) -> list[dict]:
    """
    Costruisce la lista di coppie di file da confrontare.
    Ritorna lista di dict: { name, asis_path, tobe_path, logical_name }
    """
    def _csv_files(d: Path | None) -> dict[str, Path]:
        if d is None or not d.exists():
            return {}
        return {_logical_name(f.name): f
                for f in d.iterdir()
                if f.suffix.lower() == ".csv"}

    asis_map = _csv_files(asis_dir)
    tobe_map = _csv_files(tobe_dir)

    all_keys = sorted(set(asis_map) | set(tobe_map))
    pairs = []
    for key in all_keys:
        pairs.append({
            "logical_name": key,
            "short_name":   key.split(".")[-1] if "." in key else key,
            "asis_path":    asis_map.get(key),
            "tobe_path":    tobe_map.get(key),
        })
    return pairs


# ─── Lettura CSV ──────────────────────────────────────────────────────────────

def _detect_header(filepath: Path) -> bool:
    """
    Ritorna True se la prima riga sembra un header (contiene almeno un campo
    puramente alfabetico/underscore, non un numero).
    """
    with open(filepath, "r", encoding="utf-8", errors="replace") as fh:
        first_line = fh.readline().rstrip("\r\n")
    fields = [f.strip() for f in first_line.split(";") if f.strip()]
    if not fields:
        return False
    alpha_count = sum(1 for f in fields if re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", f))
    return alpha_count >= max(1, len(fields) * 0.5)


def read_csv(filepath: Path) -> pd.DataFrame:
    """Legge il CSV (sep=;) rilevando automaticamente la presenza di header."""
    has_header = _detect_header(filepath)
    header_row = 0 if has_header else None
    df = pd.read_csv(
        filepath,
        sep=";",
        header=header_row,
        dtype=str,
        encoding="utf-8",
        encoding_errors="replace",
        skipinitialspace=True,
        keep_default_na=False,
    )
    # Rimuovi colonne completamente vuote (trailing semicolon → colonna vuota)
    df = df.loc[:, df.apply(lambda c: c.str.strip().ne("").any())]
    # Strip spazi da tutti i valori
    df = df.apply(lambda c: c.str.strip() if c.dtype == object else c)
    # Nomi colonne come stringa
    df.columns = [str(c) for c in df.columns]
    return df


# ─── Analisi discrepanze ──────────────────────────────────────────────────────

def compare_columns(df_asis: pd.DataFrame, df_tobe: pd.DataFrame) -> pd.DataFrame:
    """
    Confronta nomi e tipo inferred delle colonne.
    """
    def _infer_type(series: pd.Series) -> str:
        sample = series.dropna().head(100)
        if sample.empty:
            return "VUOTO"
        numeric = pd.to_numeric(sample, errors="coerce").notna().mean()
        if numeric > 0.9:
            # distingui intero vs decimale
            has_dot = sample.str.contains(r"\.", na=False).any()
            return "DECIMALE" if has_dot else "INTERO"
        try:
            pd.to_datetime(sample, dayfirst=True, errors="raise")
            return "DATA"
        except Exception:
            pass
        return "TESTO"

    cols_asis = list(df_asis.columns)
    cols_tobe = list(df_tobe.columns)
    all_cols  = sorted(set(cols_asis) | set(cols_tobe),
                       key=lambda c: (cols_asis.index(c) if c in cols_asis else 9999))

    rows = []
    for c in all_cols:
        in_asis  = c in cols_asis
        in_tobe  = c in cols_tobe
        type_as  = _infer_type(df_asis[c]) if in_asis else "—"
        type_to  = _infer_type(df_tobe[c]) if in_tobe else "—"
        pos_as   = cols_asis.index(c) + 1 if in_asis else "—"
        pos_to   = cols_tobe.index(c) + 1 if in_tobe else "—"
        status   = "OK" if in_asis and in_tobe else ("SOLO AS-IS" if in_asis else "SOLO TO-BE")
        tipo_ok  = type_as == type_to if in_asis and in_tobe else False

        rows.append({
            "COLONNA":        c,
            "IN AS-IS":       "Sì" if in_asis else "No",
            "IN TO-BE":       "Sì" if in_tobe else "No",
            "POS AS-IS":      pos_as,
            "POS TO-BE":      pos_to,
            "TIPO AS-IS":     type_as,
            "TIPO TO-BE":     type_to,
            "TIPO COERENTE":  "Sì" if tipo_ok else "No",
            "STATUS":         status,
        })
    return pd.DataFrame(rows)


def compare_rows(df_asis: pd.DataFrame, df_tobe: pd.DataFrame,
                 key_cols: list[str] | None = None) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Confronto riga per riga.
    Ritorna (diff_df, only_asis_df, only_tobe_df).

    Se key_cols è fornito, il join avviene su quelle colonne;
    altrimenti si usa l'indice posizionale.
    """
    common_cols = [c for c in df_asis.columns if c in df_tobe.columns]

    if key_cols and all(k in common_cols for k in key_cols):
        # join su chiave
        merged = pd.merge(
            df_asis[common_cols].reset_index(drop=True).assign(_src="AS-IS"),
            df_tobe[common_cols].reset_index(drop=True).assign(_src="TO-BE"),
            on=key_cols, how="outer", suffixes=("__ASIS", "__TOBE"), indicator=True
        )
        only_asis = merged[merged["_merge"] == "left_only"].drop(columns=["_merge", "_src_x", "_src_y"], errors="ignore")
        only_tobe = merged[merged["_merge"] == "right_only"].drop(columns=["_merge", "_src_x", "_src_y"], errors="ignore")
        both      = merged[merged["_merge"] == "both"].copy()

        value_cols = [c for c in common_cols if c not in key_cols]
        diff_rows  = []
        for _, row in both.iterrows():
            for col in value_cols:
                v_as = row.get(f"{col}__ASIS", "")
                v_to = row.get(f"{col}__TOBE", "")
                if str(v_as).strip() != str(v_to).strip():
                    rec = {k: row[k] for k in key_cols}
                    rec["COLONNA"]        = col
                    rec["VALORE AS-IS"]   = v_as
                    rec["VALORE TO-BE"]   = v_to
                    rec["DIFF_NUMERICA"]  = _numeric_diff(v_as, v_to)
                    diff_rows.append(rec)
        diff_df = pd.DataFrame(diff_rows)

    else:
        # confronto posizionale (stesso numero di righe o tronca al minimo)
        n    = min(len(df_asis), len(df_tobe))
        only_asis = df_asis.iloc[n:].copy() if len(df_asis) > n else pd.DataFrame(columns=df_asis.columns)
        only_tobe = df_tobe.iloc[n:].copy() if len(df_tobe) > n else pd.DataFrame(columns=df_tobe.columns)

        diff_rows = []
        for idx in range(n):
            for col in common_cols:
                v_as = str(df_asis.at[idx, col]).strip()
                v_to = str(df_tobe.at[idx, col]).strip()
                if v_as != v_to:
                    diff_rows.append({
                        "RIGA (0-based)":   idx,
                        "COLONNA":          col,
                        "VALORE AS-IS":     v_as,
                        "VALORE TO-BE":     v_to,
                        "DIFF_NUMERICA":    _numeric_diff(v_as, v_to),
                    })
        diff_df = pd.DataFrame(diff_rows)

    return diff_df, only_asis, only_tobe


def _numeric_diff(a: object, b: object) -> object:
    """Calcola la differenza numerica A - B, oppure '' se non numerici."""
    try:
        fa, fb = float(str(a).replace(",", ".")), float(str(b).replace(",", "."))
        return fa - fb
    except (ValueError, TypeError):
        return ""


# ─── Scrittura Excel ──────────────────────────────────────────────────────────

def _write_df_to_sheet(ws, df: pd.DataFrame, title: str,
                        hdr_color: str = COLOR_SUBHDR,
                        diff_col: str | None = None):
    """Scrive un DataFrame nel foglio ws a partire da riga 1."""
    if df.empty:
        ws.append(["(nessuna riga)"])
        return

    # Header
    ws.append(list(df.columns))
    _style_header_row(ws, ws.max_row, hdr_color)

    # Righe dati
    for _, row_data in df.iterrows():
        ws.append([str(v) if v != "" else "" for v in row_data])
        row_num = ws.max_row
        _style_data_row(ws, row_num)
        # Evidenzia le celle con valore discrepante se diff_col indicato
        if diff_col and diff_col in df.columns:
            pass  # colore già gestito row per row sotto

    _autofit(ws)


def build_excel(pairs: list[dict], output_path: str):
    """Costruisce l'Excel finale."""
    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)   # rimuove il foglio vuoto di default

    # ── Sheet RIEPILOGO ───────────────────────────────────────────────────────
    ws_riepilogo = wb.create_sheet("RIEPILOGO")
    ws_riepilogo.append(["RIEPILOGO CONFRONTO AS-IS vs TO-BE"])
    ws_riepilogo.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
    ws_riepilogo["A1"].fill  = _fill(COLOR_HEADER)
    ws_riepilogo["A1"].font  = _font(bold=True, color="FFFFFF", size=13)
    ws_riepilogo["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws_riepilogo.row_dimensions[1].height = 24

    ws_riepilogo.append([
        "FILE LOGICO", "FILE AS-IS", "FILE TO-BE",
        "RIGHE AS-IS", "RIGHE TO-BE", "DELTA RIGHE",
        "COL SOLO AS-IS", "COL SOLO TO-BE", "RIGHE DISCREPANTI"
    ])
    _style_header_row(ws_riepilogo, 2, COLOR_SUBHDR)

    summary_rows = []

    for pair in pairs:
        lname  = pair["logical_name"]
        sname  = pair["short_name"]
        a_path = pair["asis_path"]
        t_path = pair["tobe_path"]

        # Lettura
        if a_path and a_path.stat().st_size > 0:
            df_asis = read_csv(a_path)
        else:
            df_asis = pd.DataFrame()

        if t_path and t_path.stat().st_size > 0:
            df_tobe = read_csv(t_path)
        else:
            df_tobe = pd.DataFrame()

        pair["df_asis"] = df_asis
        pair["df_tobe"] = df_tobe

        n_as = len(df_asis)
        n_to = len(df_tobe)

        # Analisi colonne
        if not df_asis.empty or not df_tobe.empty:
            col_df = compare_columns(df_asis, df_tobe)
            only_asis_cols = col_df[col_df["STATUS"] == "SOLO AS-IS"]["COLONNA"].tolist()
            only_tobe_cols = col_df[col_df["STATUS"] == "SOLO TO-BE"]["COLONNA"].tolist()
        else:
            col_df         = pd.DataFrame()
            only_asis_cols = []
            only_tobe_cols = []

        # Analisi righe
        if not df_asis.empty and not df_tobe.empty:
            diff_df, only_asis_rows, only_tobe_rows = compare_rows(df_asis, df_tobe)
        else:
            diff_df        = pd.DataFrame()
            only_asis_rows = df_asis.copy()
            only_tobe_rows = df_tobe.copy()

        n_diff_rows = len(diff_df["RIGA (0-based)"].unique()) if "RIGA (0-based)" in diff_df.columns else len(diff_df)

        pair["col_df"]         = col_df
        pair["diff_df"]        = diff_df
        pair["only_asis_rows"] = only_asis_rows
        pair["only_tobe_rows"] = only_tobe_rows
        pair["only_asis_cols"] = only_asis_cols
        pair["only_tobe_cols"] = only_tobe_cols

        # Riga riepilogo
        status_color = COLOR_OK if (not only_asis_cols and not only_tobe_cols and n_diff_rows == 0 and n_as == n_to) else COLOR_KO
        summary_rows.append((
            lname,
            a_path.name if a_path else "— MANCANTE —",
            t_path.name if t_path else "— MANCANTE —",
            n_as, n_to, n_as - n_to,
            ", ".join(only_asis_cols) or "—",
            ", ".join(only_tobe_cols) or "—",
            n_diff_rows,
            status_color,
        ))

    for *row_vals, sc in summary_rows:
        ws_riepilogo.append(row_vals)
        _style_data_row(ws_riepilogo, ws_riepilogo.max_row, sc)

    _autofit(ws_riepilogo)
    ws_riepilogo.freeze_panes = "A3"

    # ── Sheet per ogni coppia ─────────────────────────────────────────────────
    for pair in pairs:
        sname  = pair["short_name"][:18]   # max 18 char per evitare nomi sheet troppo lunghi
        col_df         = pair["col_df"]
        diff_df        = pair["diff_df"]
        only_asis_rows = pair["only_asis_rows"]
        only_tobe_rows = pair["only_tobe_rows"]

        # ── COLONNE ──
        ws_col = wb.create_sheet(f"COLONNE_{sname}")
        ws_col.append([f"Confronto colonne – {pair['logical_name']}"])
        ws_col.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
        ws_col["A1"].fill = _fill(COLOR_HEADER)
        ws_col["A1"].font = _font(bold=True, color="FFFFFF", size=12)
        ws_col.row_dimensions[1].height = 20

        if col_df.empty:
            ws_col.append(["(nessun dato disponibile)"])
        else:
            ws_col.append(list(col_df.columns))
            _style_header_row(ws_col, 2, COLOR_SUBHDR)
            for _, row_data in col_df.iterrows():
                ws_col.append(list(row_data))
                r = ws_col.max_row
                _style_data_row(ws_col, r)
                status_val = str(row_data.get("STATUS", ""))
                tipo_ok    = str(row_data.get("TIPO COERENTE", "Sì"))
                if status_val != "OK":
                    fill_c = COLOR_ONLY_AS if status_val == "SOLO AS-IS" else COLOR_ONLY_TO
                    for cell in ws_col[r]:
                        cell.fill = _fill(fill_c)
                elif tipo_ok == "No":
                    for cell in ws_col[r]:
                        cell.fill = _fill(COLOR_DIFF)

        _autofit(ws_col)
        ws_col.freeze_panes = "A3"

        # ── DIFFERENZE RIGA PER RIGA ──
        ws_diff = wb.create_sheet(f"DIFF_{sname}")
        ws_diff.append([f"Differenze riga per riga – {pair['logical_name']}"])
        ws_diff.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
        ws_diff["A1"].fill = _fill(COLOR_HEADER)
        ws_diff["A1"].font = _font(bold=True, color="FFFFFF", size=12)
        ws_diff.row_dimensions[1].height = 20

        if diff_df.empty:
            ws_diff.append(["(nessuna differenza rilevata)"])
        else:
            ws_diff.append(list(diff_df.columns))
            _style_header_row(ws_diff, 2, COLOR_SUBHDR)
            for _, row_data in diff_df.iterrows():
                ws_diff.append(list(row_data))
                r = ws_diff.max_row
                _style_data_row(ws_diff, r, COLOR_DIFF)

        _autofit(ws_diff)
        ws_diff.freeze_panes = "A3"

        # ── SOLO AS-IS ──
        ws_as = wb.create_sheet(f"SOLO_ASIS_{sname}")
        ws_as.append([f"Righe presenti solo in AS-IS – {pair['logical_name']}"])
        ws_as.merge_cells(start_row=1, start_column=1, end_row=1,
                          end_column=max(len(only_asis_rows.columns), 1))
        ws_as["A1"].fill = _fill(COLOR_HEADER)
        ws_as["A1"].font = _font(bold=True, color="FFFFFF", size=12)
        ws_as.row_dimensions[1].height = 20

        if only_asis_rows.empty:
            ws_as.append(["(nessuna riga esclusiva AS-IS)"])
        else:
            ws_as.append(list(only_asis_rows.columns))
            _style_header_row(ws_as, 2, COLOR_SUBHDR)
            for _, row_data in only_asis_rows.iterrows():
                ws_as.append(list(row_data))
                _style_data_row(ws_as, ws_as.max_row, COLOR_ONLY_AS)

        _autofit(ws_as)

        # ── SOLO TO-BE ──
        ws_to = wb.create_sheet(f"SOLO_TOBE_{sname}")
        ws_to.append([f"Righe presenti solo in TO-BE – {pair['logical_name']}"])
        ws_to.merge_cells(start_row=1, start_column=1, end_row=1,
                          end_column=max(len(only_tobe_rows.columns), 1))
        ws_to["A1"].fill = _fill(COLOR_HEADER)
        ws_to["A1"].font = _font(bold=True, color="FFFFFF", size=12)
        ws_to.row_dimensions[1].height = 20

        if only_tobe_rows.empty:
            ws_to.append(["(nessuna riga esclusiva TO-BE)"])
        else:
            ws_to.append(list(only_tobe_rows.columns))
            _style_header_row(ws_to, 2, COLOR_SUBHDR)
            for _, row_data in only_tobe_rows.iterrows():
                ws_to.append(list(row_data))
                _style_data_row(ws_to, ws_to.max_row, COLOR_ONLY_TO)

        _autofit(ws_to)

    wb.save(output_path)
    print(f"\n✓ Excel salvato in: {output_path}")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Confronto AS-IS vs TO-BE file PLZ3A*"
    )
    parser.add_argument("--base-dir", default=".",
                        help="Directory base dove cercare le cartelle PLZ3A* (default: .)")
    parser.add_argument("--asis",  default=None,
                        help="Path esplicito alla cartella AS-IS (sovrascrive auto-detect)")
    parser.add_argument("--tobe",  default=None,
                        help="Path esplicito alla cartella TO-BE (sovrascrive auto-detect)")
    parser.add_argument("--output", default=None,
                        help="Nome file Excel di output (default: confronto_PLZ3A_YYYYMMDD_HHMMSS.xlsx)")
    args = parser.parse_args()

    # ─ Auto-detect o path espliciti ─
    if args.asis or args.tobe:
        asis_path = Path(args.asis) if args.asis else None
        tobe_path = Path(args.tobe) if args.tobe else None
    else:
        asis_path, tobe_path = find_plz3a_folders(args.base_dir)

    # ─ Stampa cosa è stato trovato ─
    print("=" * 60)
    print("  CONFRONTO AS-IS vs TO-BE – PLZ3A")
    print("=" * 60)
    print(f"  AS-IS : {asis_path or '⚠  NON TROVATA'}")
    print(f"  TO-BE : {tobe_path or '⚠  NON TROVATA'}")

    if asis_path is None and tobe_path is None:
        print("\n✗ Nessuna cartella PLZ3A trovata. Usa --asis / --tobe per specificarle manualmente.")
        sys.exit(1)

    # ─ Match file ─
    pairs = match_files(asis_path, tobe_path)
    if not pairs:
        print("\n✗ Nessun file CSV trovato nelle cartelle.")
        sys.exit(1)

    print(f"\n  File CSV rilevati: {len(pairs)}")
    for p in pairs:
        a = p["asis_path"].name if p["asis_path"] else "MANCANTE"
        t = p["tobe_path"].name if p["tobe_path"] else "MANCANTE"
        print(f"    [{p['short_name']}]")
        print(f"      AS-IS : {a}")
        print(f"      TO-BE : {t}")

    # ─ Output path ─
    if args.output:
        out = args.output
    else:
        ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = os.path.join(args.base_dir, f"confronto_PLZ3A_{ts}.xlsx")

    print(f"\n  Elaborazione in corso...")
    build_excel(pairs, out)
    print("  Done.")


if __name__ == "__main__":
    main()
