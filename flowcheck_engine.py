# flowcheck_engine.py
# Motore generale di confronto CSV / ZIP / Cartella
# Produce un Excel per coppia di file (AS-IS vs TO-BE)

from __future__ import annotations

import io
import os
import re
import traceback
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Separatori candidati (ordine decrescente di specificita')
# ---------------------------------------------------------------------------
SEP_CANDIDATES = [";|", ";£", "£|", "|;", "\t", ";", ",", "|", "£"]

# ---------------------------------------------------------------------------
# Stili Excel
# ---------------------------------------------------------------------------
_CLR = dict(
    header_bg="1F3864", header_fg="FFFFFF",
    ok_bg="C6EFCE",     ok_fg="276221",
    diff_bg="FFEB9C",   diff_fg="9C5700",
    only_a_bg="FFD7D7", only_a_fg="9C0006",
    only_b_bg="D9E1F2", only_b_fg="1F3864",
    err_bg="FFB3B3",    err_fg="7B0000",
    neutral_bg="F2F2F2",
)

def _fill(hex_bg): return PatternFill("solid", fgColor=hex_bg)
def _font(hex_fg, bold=False): return Font(color=hex_fg, bold=bold)
_border = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)


# ---------------------------------------------------------------------------
# Rilevamento separatore
# ---------------------------------------------------------------------------

def _read_first_lines(filepath: Path, n: int = 30,
                      zip_entry: str | None = None) -> list[str]:
    """
    Legge le prime n righe da:
    - file CSV normale
    - CSV specifico dentro uno ZIP (zip_entry = nome file interno)
    - primo CSV trovato in uno ZIP (zip_entry=None)
    """
    try:
        if filepath.suffix.lower() == ".zip":
            with zipfile.ZipFile(filepath) as zf:
                target = zip_entry
                if target is None:
                    csv_names = [x for x in zf.namelist()
                                 if x.lower().endswith(".csv") and not x.startswith("__")]
                    target = csv_names[0] if csv_names else None
                if target:
                    with zf.open(target) as fh:
                        raw = fh.read(16384).decode("utf-8", errors="replace")
                        return raw.splitlines()[:n]
        else:
            with open(filepath, encoding="utf-8", errors="replace") as fh:
                return [fh.readline() for _ in range(n)]
    except Exception:
        pass
    return []


def _build_sep_candidates(lines: list[str]) -> list[str]:
    """
    Costruisce la lista di separatori candidati unendo quelli statici con
    eventuali separatori compositi ;X rilevati dinamicamente nelle righe.
    Questo permette di gestire qualsiasi variante (;#, ;£, ;|, ;@, ...) senza
    dover aggiornare la lista fissa.
    """
    import collections

    cands = list(SEP_CANDIDATES)

    # Cerca il carattere X che segue ';' in modo costante su tutte le righe
    after_semi: collections.Counter = collections.Counter()
    non_empty_lines = [l for l in lines if l.strip()]
    for line in non_empty_lines:
        for i in range(len(line) - 1):
            if line[i] == ";":
                nxt = line[i + 1]
                # ignora spazi, virgolette, newline, alfanumerici, altro ;
                if not nxt.isalnum() and nxt not in (";", " ", "\n", "\r", '"', "'"):
                    after_semi[nxt] += 1

    if after_semi:
        best_char, count = after_semi.most_common(1)[0]
        # accettiamo il candidato se appare in almeno il 50% delle righe non vuote
        threshold = max(1, len(non_empty_lines) * 0.5)
        if count >= threshold:
            compound = f";{best_char}"
            if compound not in cands:
                cands.insert(0, compound)   # massima priorita'

    return cands


def detect_separator(filepath: str | Path,
                     candidates: list[str] | None = None,
                     zip_entry: str | None = None) -> str:
    """
    Rileva il separatore CSV piu' probabile usando un punteggio
    avg_occorrenze / (1 + varianza) su max 30 righe.

    filepath  : percorso al file CSV o ZIP
    candidates: lista separatori da testare; None = auto (statici + dinamici)
    zip_entry : nome del file CSV dentro lo ZIP (None = primo trovato)
    """
    import statistics

    filepath = Path(filepath)
    lines = _read_first_lines(filepath, n=30, zip_entry=zip_entry)
    if not lines:
        return ";"

    cands = list(candidates) if candidates else _build_sep_candidates(lines)

    best_sep, best_score = ";", -1.0
    for sep in cands:
        counts = [line.count(sep) for line in lines if line.strip()]
        if not counts or max(counts) == 0:
            continue
        avg = statistics.mean(counts)
        var = statistics.variance(counts) if len(counts) > 1 else 0
        score = avg / (1.0 + var)
        if score > best_score:
            best_score, best_sep = score, sep
    return best_sep


# ---------------------------------------------------------------------------
# Lettura CSV (con auto-detect del separatore)
# ---------------------------------------------------------------------------

def _clean_str_series(s: pd.Series) -> pd.Series:
    """
    Normalizza una Series di stringhe per la lettura e il confronto:
    - strip leading/trailing whitespace ASCII
    - rimuove spazi non-breaking (U+00A0) da bordi
    - collassa spazi interni multipli in uno singolo
    """
    return (
        s.str.strip()
         .str.replace("\xa0", " ", regex=False)   # non-breaking space -> spazio
         .str.replace(r"[ \t]+", " ", regex=True)  # spazi/tab multipli -> singolo
         .str.strip()
    )


def _clean_df(df: pd.DataFrame) -> pd.DataFrame:
    """Applica _clean_str_series a tutte le colonne stringa del DataFrame."""
    return df.apply(lambda col: _clean_str_series(col) if col.dtype == object else col)


def _dedup_columns(cols: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    result = []
    for c in cols:
        if c in seen:
            seen[c] += 1
            result.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            result.append(c)
    return result


def _has_header(filepath: str | Path, sep: str = ";") -> bool:
    """Euristica: True se la prima riga contiene almeno 1 valore non numerico."""
    try:
        with open(filepath, encoding="utf-8", errors="replace") as fh:
            first = fh.readline()
        parts = first.split(sep)
        for p in parts:
            p = p.strip().strip('"').strip("'")
            if p and not re.match(r"^-?\d+(\.\d+)?$", p):
                return True
    except Exception:
        pass
    return False


def read_csv(filepath: str | Path, sep: str | None = None) -> pd.DataFrame:
    """
    Legge un CSV (anche da dentro un ZIP se filepath e' un Path di ZIP).
    sep=None => auto-detect.
    """
    filepath = Path(filepath)
    if sep is None:
        sep = detect_separator(filepath)

    # engine python obbligatorio per separatori multi-carattere
    engine = "python" if len(sep) > 1 else "c"
    sep_param = re.escape(sep) if engine == "python" else sep

    hdr = 0 if _has_header(filepath, sep) else None

    df = pd.read_csv(
        filepath,
        sep=sep_param,
        header=hdr,
        dtype=str,
        encoding="utf-8",
        encoding_errors="replace",
        skipinitialspace=True,
        keep_default_na=False,
        engine=engine,
    )
    df.columns = _dedup_columns([str(c).strip() for c in df.columns])
    # Rimuovi colonne-artefatto: nome vuoto o solo caratteri speciali (es. '#', '£')
    df = df[[c for c in df.columns if not _is_artifact_col(c)]]
    df = _clean_df(df)
    return df


def _is_artifact_col(col_name: str) -> bool:
    """
    True se il nome colonna e' un artefatto del separatore CSV e va scartato:
    - stringa vuota  (es. trailing ';' -> '')
    - solo caratteri speciali non-word  (es. '#', '£', '|', ';#', ';£')
    I nomi legittimi contengono almeno una lettera, cifra o underscore.
    """
    stripped = col_name.strip()
    if not stripped:
        return True
    # nessun carattere \w (lettera, cifra, _) -> artefatto
    if not re.search(r"\w", stripped):
        return True
    return False


def read_csv_from_zip(zip_path: str | Path, csv_name: str, sep: str | None = None) -> pd.DataFrame:
    """Estrae un CSV da uno ZIP in memoria e lo legge."""
    zip_path = Path(zip_path)
    if sep is None:
        # Rileva il separatore dal CSV specifico, non dal primo dello ZIP
        sep = detect_separator(zip_path, zip_entry=csv_name)
    engine = "python" if len(sep) > 1 else "c"
    sep_param = re.escape(sep) if engine == "python" else sep

    with zipfile.ZipFile(zip_path) as zf:
        with zf.open(csv_name) as fh:
            raw = fh.read().decode("utf-8", errors="replace")

    hdr_flag = _has_header_from_text(raw, sep)
    hdr = 0 if hdr_flag else None
    df = pd.read_csv(
        io.StringIO(raw),
        sep=sep_param,
        header=hdr,
        dtype=str,
        encoding_errors="replace",
        skipinitialspace=True,
        keep_default_na=False,
        engine=engine,
    )
    df.columns = _dedup_columns([str(c).strip() for c in df.columns])
    # Rimuovi colonne-artefatto: nome vuoto o solo caratteri speciali (es. '#', '£')
    df = df[[c for c in df.columns if not _is_artifact_col(c)]]
    df = _clean_df(df)
    return df


def _has_header_from_text(text: str, sep: str) -> bool:
    first = text.split("\n")[0] if text else ""
    for p in first.split(sep):
        p = p.strip().strip('"').strip("'")
        if p and not re.match(r"^-?\d+(\.\d+)?$", p):
            return True
    return False


# ---------------------------------------------------------------------------
# Matching file tra AS-IS e TO-BE
# ---------------------------------------------------------------------------

def _stem_key(name: str) -> str:
    """
    Chiave di matching normalizzata: rimuove timestamp, estensione e
    prefissi tecnici noti (DW, D, M, PLZxx), conservando il nome tabella
    anche quando composto da piu' parti (es. ABK001FW_DANNI, ABBINAMENTO_PLZ_DANNI).

    Esempi:
      DW.D.PLZHA.AVTBCODI.20260326000000.000L.20260326030157.csv  => AVTBCODI
      DW.D.PLZ3A.ABBINAMENTO_PLZ_DANNI.20260326000000.000L.csv    => ABBINAMENTO_PLZ_DANNI
      DW.D.PLZAA.ABK001FW_DANNI.csv                               => ABK001FW_DANNI
      DW.D.PLZBA.STORNI_DANNI.csv                                 => STORNI_DANNI
    """
    name = Path(name).stem.upper()
    # Rimuovi blocchi timestamp (es. 20260326000000) e suffissi tipo 000L
    name = re.sub(r"\b\d{14}\b", "", name)
    name = re.sub(r"\b\d{3}L\b", "", name)
    # Normalizza separatori in underscore, pulisci underscore multipli
    name = re.sub(r"[.\-]+", "_", name)
    name = re.sub(r"_+", "_", name).strip("_")
    # Rimuovi prefissi tecnici noti (DW_D_PLZxx_, DW_M_PLZxx_, D_PLZxx_, ecc.)
    # PLZxx = PLZ seguito da esattamente 2 caratteri alfanumerici (es. 3A, HA, AA ...)
    name = re.sub(r"^(DW_)?[DM]_PLZ\w{2}_", "", name)
    # Rimuovi eventuali residui DW_D_ / DW_M_ / D_ / M_ rimasti
    name = re.sub(r"^(DW_)?[DM]_", "", name)
    name = name.strip("_")
    return name if name else Path(name).stem.upper()


def match_files(
    asis_sources: list[tuple[str, str]],   # (display_name, path_or_zipentry)
    tobe_sources: list[tuple[str, str]],
) -> list[dict]:
    """
    Data due liste di (label, path), restituisce coppie abbinate
    [{'label': str, 'asis': str, 'tobe': str}, ...].
    """
    tobe_by_key: dict[str, tuple[str, str]] = {}
    for label, path in tobe_sources:
        tobe_by_key[_stem_key(label)] = (label, path)

    pairs = []
    for label, path in asis_sources:
        key = _stem_key(label)
        if key in tobe_by_key:
            tb_label, tb_path = tobe_by_key[key]
            pairs.append({"label": key, "asis_label": label, "tobe_label": tb_label,
                          "asis_path": path, "tobe_path": tb_path})
        else:
            pairs.append({"label": key, "asis_label": label, "tobe_label": None,
                          "asis_path": path, "tobe_path": None})
    return pairs


# ---------------------------------------------------------------------------
# Detect join key
# ---------------------------------------------------------------------------

def detect_join_key(df_a: pd.DataFrame, df_b: pd.DataFrame) -> list[str]:
    """
    Cerca colonne comuni che identifichino univocamente le righe in ENTRAMBI
    i DataFrame (combinazione con cardinalita' massima e duplicati minimi).
    Fallback: indice numerico (nessuna chiave).
    """
    common = [c for c in df_a.columns if c in df_b.columns]
    if not common:
        return []

    # Proviamo combinazioni dalla piu' semplice (1 col) alla piu' complessa
    from itertools import combinations

    best_key: list[str] = []
    best_score = -1.0

    for r in range(1, min(5, len(common) + 1)):
        for combo in combinations(common, r):
            combo = list(combo)
            n_unique_a = df_a[combo].drop_duplicates().shape[0]
            n_unique_b = df_b[combo].drop_duplicates().shape[0]
            # punteggio: media delle frazioni di univocita'
            score_a = n_unique_a / max(len(df_a), 1)
            score_b = n_unique_b / max(len(df_b), 1)
            score = (score_a + score_b) / 2
            if score > best_score:
                best_score = score
                best_key = combo
            if best_score >= 0.999:
                break
        if best_score >= 0.999:
            break

    # Se il punteggio e' troppo basso non ha senso usare quella chiave
    if best_score < 0.5:
        return []
    return best_key


# ---------------------------------------------------------------------------
# Confronto DataFrame
# ---------------------------------------------------------------------------

def compare_dataframes(
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    join_key: list[str] | None = None,
    max_diff_rows: int = 10_000,
) -> dict:
    """
    Confronta df_a (AS-IS) con df_b (TO-BE).
    Restituisce un dizionario con:
      - summary: dict con contatori
      - cols_only_a, cols_only_b, cols_common: liste colonne
      - df_only_a, df_only_b: righe uniche
      - df_diff: righe con differenze (max_diff_rows)
      - key_used: lista colonne chiave effettivamente usata
    """
    cols_a = set(df_a.columns)
    cols_b = set(df_b.columns)
    cols_only_a = sorted(cols_a - cols_b)
    cols_only_b = sorted(cols_b - cols_a)
    cols_common = sorted(cols_a & cols_b)

    if join_key is None:
        join_key = detect_join_key(df_a[cols_common] if cols_common else df_a,
                                   df_b[cols_common] if cols_common else df_b)

    # Confronto strutturale
    struct_ok = (cols_only_a == [] and cols_only_b == [])

    if not join_key:
        # Confronto posizionale
        min_rows = min(len(df_a), len(df_b))
        common_df_a = df_a[cols_common].iloc[:min_rows].reset_index(drop=True)
        common_df_b = df_b[cols_common].iloc[:min_rows].reset_index(drop=True)
        # Confronto su valori normalizzati (elimina falsi positivi da spazi)
        norm_a = _clean_df(common_df_a)
        norm_b = _clean_df(common_df_b)
        diff_mask = (norm_a != norm_b)
        diff_rows_idx = diff_mask.any(axis=1)
        # Nel report mostra i valori originali (non normalizzati)
        df_diff_a = common_df_a[diff_rows_idx].head(max_diff_rows)
        df_diff_b = common_df_b[diff_rows_idx].head(max_diff_rows)
        df_only_a = df_a.iloc[min_rows:].head(max_diff_rows) if len(df_a) > min_rows else pd.DataFrame()
        df_only_b = df_b.iloc[min_rows:].head(max_diff_rows) if len(df_b) > min_rows else pd.DataFrame()
        n_diff = int(diff_rows_idx.sum())
        key_used = []
    else:
        # Confronto per chiave — merge sui valori normalizzati della chiave
        df_a_norm = df_a.copy()
        df_b_norm = df_b.copy()
        for k in join_key:
            if k in df_a_norm.columns:
                df_a_norm[k] = _clean_str_series(df_a_norm[k].astype(str))
            if k in df_b_norm.columns:
                df_b_norm[k] = _clean_str_series(df_b_norm[k].astype(str))

        merged = pd.merge(
            df_a_norm[cols_common + [c for c in cols_only_a]],
            df_b_norm[cols_common + [c for c in cols_only_b]],
            on=join_key, how="outer", indicator=True, suffixes=("__A", "__B"),
        )
        df_only_a = merged[merged["_merge"] == "left_only"].drop(columns=["_merge"]).head(max_diff_rows)
        df_only_b = merged[merged["_merge"] == "right_only"].drop(columns=["_merge"]).head(max_diff_rows)
        both = merged[merged["_merge"] == "both"].drop(columns=["_merge"])

        # Confronto colonne non-chiave su valori normalizzati
        compare_cols = [c for c in cols_common if c not in join_key]
        diff_rows_mask = pd.Series(False, index=both.index)
        for c in compare_cols:
            ca, cb = f"{c}__A", f"{c}__B"
            if ca in both.columns and cb in both.columns:
                val_a = _clean_str_series(both[ca].fillna("").astype(str))
                val_b = _clean_str_series(both[cb].fillna("").astype(str))
                diff_rows_mask |= (val_a != val_b)
        df_diff_a = both[diff_rows_mask][join_key + [f"{c}__A" for c in compare_cols if f"{c}__A" in both.columns]].head(max_diff_rows)
        df_diff_b = both[diff_rows_mask][join_key + [f"{c}__B" for c in compare_cols if f"{c}__B" in both.columns]].head(max_diff_rows)
        n_diff = int(diff_rows_mask.sum())
        key_used = join_key

    summary = {
        "rows_a": len(df_a),
        "rows_b": len(df_b),
        "cols_only_a": len(cols_only_a),
        "cols_only_b": len(cols_only_b),
        "cols_common": len(cols_common),
        "rows_only_a": len(df_only_a),
        "rows_only_b": len(df_only_b),
        "rows_diff": n_diff,
        "struct_ok": struct_ok,
    }
    return {
        "summary": summary,
        "cols_only_a": cols_only_a,
        "cols_only_b": cols_only_b,
        "cols_common": cols_common,
        "df_only_a": df_only_a,
        "df_only_b": df_only_b,
        "df_diff_a": df_diff_a,
        "df_diff_b": df_diff_b,
        "key_used": key_used,
    }


# ---------------------------------------------------------------------------
# Scrittura Excel
# ---------------------------------------------------------------------------

def _write_df_to_sheet(ws, df: pd.DataFrame, header_label: str,
                        header_bg: str, header_fg: str,
                        row_bg: str | None = None, row_fg: str | None = None):
    """Scrive un DataFrame in un worksheet openpyxl gia' creato."""
    if df.empty:
        ws.append(["(nessuna riga)"])
        return

    # Intestazione
    for ci, col in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=ci, value=str(col))
        cell.fill = _fill(header_bg)
        cell.font = _font(header_fg, bold=True)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = _border

    # Righe
    for ri, (_, row) in enumerate(df.iterrows(), 2):
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci, value=str(val) if pd.notna(val) else "")
            if row_bg:
                cell.fill = _fill(row_bg)
            if row_fg:
                cell.font = _font(row_fg)
            cell.border = _border

    # Auto-width (max 60)
    for ci in range(1, df.shape[1] + 1):
        max_len = max(
            len(str(ws.cell(row=r, column=ci).value or ""))
            for r in range(1, min(ws.max_row + 1, 200))
        )
        ws.column_dimensions[get_column_letter(ci)].width = min(max_len + 2, 60)


def build_excel_pair(
    label: str,
    df_a: pd.DataFrame,
    df_b: pd.DataFrame,
    asis_label: str,
    tobe_label: str,
    output_path: str | Path,
    join_key: list[str] | None = None,
    max_diff_rows: int = 10_000,
    log_path: str | Path | None = None,
) -> str:
    """
    Confronta df_a e df_b e scrive un Excel con 5 fogli:
    RIEPILOGO, SOLO_AS-IS, SOLO_TO-BE, DIFFERENZE_AS-IS, DIFFERENZE_TO-BE.
    Restituisce il path dell'Excel generato.
    """
    output_path = Path(output_path)
    result = compare_dataframes(df_a, df_b, join_key=join_key, max_diff_rows=max_diff_rows)
    s = result["summary"]

    wb = openpyxl.Workbook()

    # ---- Foglio RIEPILOGO ----
    ws_sum = wb.active
    ws_sum.title = "RIEPILOGO"
    ws_sum.sheet_view.showGridLines = True

    def _hdr(ws, text):
        ws.append([text])
        cell = ws.cell(row=ws.max_row, column=1)
        cell.fill = _fill(_CLR["header_bg"])
        cell.font = _font(_CLR["header_fg"], bold=True)
        cell.alignment = Alignment(horizontal="left")

    def _row(ws, key, val, bg=None, fg=None):
        ws.append([key, str(val)])
        if bg:
            for ci in (1, 2):
                c = ws.cell(row=ws.max_row, column=ci)
                c.fill = _fill(bg)
                if fg: c.font = _font(fg)

    _hdr(ws_sum, "RIEPILOGO CONFRONTO")
    ws_sum.append([])
    _row(ws_sum, "Coppia", label)
    _row(ws_sum, "AS-IS", asis_label)
    _row(ws_sum, "TO-BE", tobe_label)
    _row(ws_sum, "Chiave join", ", ".join(result["key_used"]) if result["key_used"] else "(posizionale)")
    ws_sum.append([])
    _row(ws_sum, "Righe AS-IS", s["rows_a"])
    _row(ws_sum, "Righe TO-BE", s["rows_b"])
    ws_sum.append([])
    struct_bg = _CLR["ok_bg"] if s["struct_ok"] else _CLR["diff_bg"]
    struct_fg = _CLR["ok_fg"] if s["struct_ok"] else _CLR["diff_fg"]
    _row(ws_sum, "Struttura colonne", "OK" if s["struct_ok"] else "DIVERSA", struct_bg, struct_fg)
    _row(ws_sum, "Colonne solo AS-IS", s["cols_only_a"],
         None if s["cols_only_a"] == 0 else _CLR["only_a_bg"],
         None if s["cols_only_a"] == 0 else _CLR["only_a_fg"])
    _row(ws_sum, "Colonne solo TO-BE", s["cols_only_b"],
         None if s["cols_only_b"] == 0 else _CLR["only_b_bg"],
         None if s["cols_only_b"] == 0 else _CLR["only_b_fg"])
    _row(ws_sum, "Colonne comuni", s["cols_common"])
    ws_sum.append([])
    diff_bg = _CLR["ok_bg"] if s["rows_diff"] == 0 else _CLR["diff_bg"]
    diff_fg = _CLR["ok_fg"] if s["rows_diff"] == 0 else _CLR["diff_fg"]
    _row(ws_sum, "Righe con differenze", s["rows_diff"], diff_bg, diff_fg)
    _row(ws_sum, "Righe solo AS-IS", s["rows_only_a"],
         None if s["rows_only_a"] == 0 else _CLR["only_a_bg"],
         None if s["rows_only_a"] == 0 else _CLR["only_a_fg"])
    _row(ws_sum, "Righe solo TO-BE", s["rows_only_b"],
         None if s["rows_only_b"] == 0 else _CLR["only_b_bg"],
         None if s["rows_only_b"] == 0 else _CLR["only_b_fg"])

    if result["cols_only_a"]:
        ws_sum.append([])
        _hdr(ws_sum, "Colonne SOLO AS-IS")
        for c in result["cols_only_a"]:
            ws_sum.append(["", c])
    if result["cols_only_b"]:
        ws_sum.append([])
        _hdr(ws_sum, "Colonne SOLO TO-BE")
        for c in result["cols_only_b"]:
            ws_sum.append(["", c])

    ws_sum.column_dimensions["A"].width = 28
    ws_sum.column_dimensions["B"].width = 60

    # ---- Foglio SOLO_AS-IS ----
    ws_oa = wb.create_sheet("SOLO_AS-IS")
    _write_df_to_sheet(ws_oa, result["df_only_a"], "SOLO_AS-IS",
                       _CLR["header_bg"], _CLR["header_fg"],
                       _CLR["only_a_bg"], _CLR["only_a_fg"])

    # ---- Foglio SOLO_TO-BE ----
    ws_ob = wb.create_sheet("SOLO_TO-BE")
    _write_df_to_sheet(ws_ob, result["df_only_b"], "SOLO_TO-BE",
                       _CLR["header_bg"], _CLR["header_fg"],
                       _CLR["only_b_bg"], _CLR["only_b_fg"])

    # ---- Fogli DIFFERENZE ----
    ws_da = wb.create_sheet("DIFF_AS-IS")
    _write_df_to_sheet(ws_da, result["df_diff_a"], "DIFF_AS-IS",
                       _CLR["header_bg"], _CLR["header_fg"],
                       _CLR["diff_bg"], _CLR["diff_fg"])

    ws_db = wb.create_sheet("DIFF_TO-BE")
    _write_df_to_sheet(ws_db, result["df_diff_b"], "DIFF_TO-BE",
                       _CLR["header_bg"], _CLR["header_fg"],
                       _CLR["diff_bg"], _CLR["diff_fg"])

    wb.save(output_path)
    return str(output_path)


# ---------------------------------------------------------------------------
# Log errori
# ---------------------------------------------------------------------------

def log_error(log_path: str | Path, label: str, exc: Exception) -> None:
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sep = "=" * 70
    msg = (f"\n{sep}\n"
           f"[{ts}] ERRORE su: {label}\n"
           f"Tipo  : {type(exc).__name__}\n"
           f"Motivo: {exc}\n"
           f"--- Traceback ---\n"
           f"{traceback.format_exc()}"
           f"{sep}\n")
    with open(log_path, "a", encoding="utf-8") as fh:
        fh.write(msg)


# ---------------------------------------------------------------------------
# Enumerazione sorgenti (file singolo, ZIP, cartella)
# ---------------------------------------------------------------------------

def enumerate_sources(path: str | Path) -> list[tuple[str, Path, str | None]]:
    """
    Dato un path (file CSV, ZIP, o cartella), restituisce una lista di
    (display_name, zip_or_csv_path, csv_entry_inside_zip_or_None).
    """
    path = Path(path)
    sources = []

    if path.is_dir():
        for f in sorted(path.iterdir()):
            if f.suffix.lower() == ".csv":
                sources.append((f.name, f, None))
            elif f.suffix.lower() == ".zip":
                sources.extend(_enumerate_zip(f))
    elif path.suffix.lower() == ".zip":
        sources.extend(_enumerate_zip(path))
    elif path.suffix.lower() == ".csv":
        sources.append((path.name, path, None))

    return sources


def _enumerate_zip(zip_path: Path) -> list[tuple[str, Path, str]]:
    result = []
    try:
        with zipfile.ZipFile(zip_path) as zf:
            for name in sorted(zf.namelist()):
                if name.lower().endswith(".csv") and not name.startswith("__"):
                    result.append((Path(name).name, zip_path, name))
    except Exception:
        pass
    return result


# ---------------------------------------------------------------------------
# Entry point principale
# ---------------------------------------------------------------------------

def run_comparison(
    asis_path: str | Path,
    tobe_path: str | Path,
    output_dir: str | Path | None = None,
    sep: str | None = None,
    join_key: list[str] | None = None,
    max_diff_rows: int = 10_000,
    progress_cb: Callable[[str], None] | None = None,
) -> list[str]:
    """
    Confronta tutte le coppie di file trovate in asis_path e tobe_path.
    Restituisce la lista degli Excel generati.
    """
    def _log(msg: str):
        if progress_cb:
            progress_cb(msg)
        else:
            print(msg)

    asis_path = Path(asis_path)
    tobe_path = Path(tobe_path)

    if output_dir is None:
        output_dir = asis_path.parent if asis_path.is_file() else asis_path
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = output_dir / f"flowcheck_errori_{ts}.log"

    asis_sources = enumerate_sources(asis_path)
    tobe_sources = enumerate_sources(tobe_path)

    _log(f"AS-IS: {len(asis_sources)} file trovati in {asis_path.name}")
    _log(f"TO-BE: {len(tobe_sources)} file trovati in {tobe_path.name}")

    # Costruisci lista per matching
    asis_list = [(dn, str(zp) if ze is None else f"{zp}::{ze}") for dn, zp, ze in asis_sources]
    tobe_list = [(dn, str(zp) if ze is None else f"{zp}::{ze}") for dn, zp, ze in tobe_sources]

    pairs = match_files(asis_list, tobe_list)
    _log(f"Coppie abbinate: {sum(1 for p in pairs if p['tobe_path'])}/{len(pairs)}")
    _log("")

    generated = []

    for pair in pairs:
        label = pair["label"]
        if pair["tobe_path"] is None:
            _log(f"[SKIP] {label} - nessuna controparte TO-BE trovata")
            continue

        _log(f"  Confronto: {label} ...")

        try:
            # Carica AS-IS
            a_path_str = pair["asis_path"]
            if "::" in a_path_str:
                zp, ze = a_path_str.split("::", 1)
                df_a = read_csv_from_zip(zp, ze, sep=sep)
            else:
                df_a = read_csv(a_path_str, sep=sep)

            # Carica TO-BE
            b_path_str = pair["tobe_path"]
            if "::" in b_path_str:
                zp, ze = b_path_str.split("::", 1)
                df_b = read_csv_from_zip(zp, ze, sep=sep)
            else:
                df_b = read_csv(b_path_str, sep=sep)

            # Genera Excel
            out_name = f"confronto_{label}_{ts}.xlsx"
            out_path = output_dir / out_name

            generated_path = build_excel_pair(
                label=label,
                df_a=df_a,
                df_b=df_b,
                asis_label=pair["asis_label"],
                tobe_label=pair["tobe_label"],
                output_path=out_path,
                join_key=join_key,
                max_diff_rows=max_diff_rows,
                log_path=log_path,
            )
            generated.append(generated_path)
            _log(f"  [OK] Excel salvato: {out_name}")

        except Exception as exc:
            log_error(log_path, label, exc)
            _log(f"  [ERRORE] {label} - vedi log: {log_path.name}")

    _log("")
    _log(f"Completato. {len(generated)}/{len([p for p in pairs if p['tobe_path']])} Excel generati.")
    if log_path.exists():
        _log(f"Log errori: {log_path.name}")
    return generated


# ---------------------------------------------------------------------------
# CLI minimale (test rapido)
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import sys
    if len(sys.argv) < 3:
        print("Uso: python flowcheck_engine.py <AS-IS> <TO-BE> [output_dir]")
        sys.exit(1)
    out = sys.argv[3] if len(sys.argv) > 3 else None
    run_comparison(sys.argv[1], sys.argv[2], output_dir=out)
