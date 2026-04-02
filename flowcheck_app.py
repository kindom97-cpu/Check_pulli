# flowcheck_app.py  —  FlowCheck  |  UI moderna ispirata a VS Code / GitHub Desktop

from __future__ import annotations

import os
import queue
import sys
import time
import threading
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

# ──────────────────────────────────────────────────────────────────────────────
# Design tokens
# ──────────────────────────────────────────────────────────────────────────────
_BG          = "#F5F7FA"   # sfondo generale
_CARD        = "#FFFFFF"   # card
_CARD_HDR    = "#FAFBFC"   # intestazione card
_BORDER      = "#E0E6ED"   # bordo sottile
_SHADOW      = "#D1D9E6"   # simulazione ombra
_PRIMARY     = "#2D7FF9"   # blu principale
_PRIMARY_DK  = "#1C5DC1"   # blu hover / pressed
_PRIMARY_LT  = "#EBF3FF"   # blu pallido (bg secondario)
_SUCCESS     = "#16A34A"
_WARN_BG     = "#FFF8E7"
_ERROR_BG    = "#FFF1F0"
_TEXT        = "#1A202C"   # testo principale
_TEXT_SEC    = "#4A5568"   # testo secondario
_TEXT_MUTED  = "#9CA3AF"   # testo tenue
_WHITE       = "#FFFFFF"

# Tipografia
_FT   = ("Segoe UI", 10)
_FT_S = ("Segoe UI",  9)
_FT_XS= ("Segoe UI",  8)
_FT_B = ("Segoe UI", 10, "bold")
_FT_H1= ("Segoe UI", 18, "bold")
_FT_H2= ("Segoe UI", 11, "bold")
_FT_H3= ("Segoe UI", 10, "bold")
_FT_M = ("Consolas",  9)

SEP_OPTIONS = [
    ("Rilevamento automatico", None),
    (";  punto e virgola",     ";"),
    (",  virgola",             ","),
    ("\\t  tabulazione",       "\t"),
    ("|  pipe",                "|"),
    (";|  composto",           ";|"),
    (";£  composto £",         ";£"),
    ("Personalizzato…",        "__custom__"),
]


# ──────────────────────────────────────────────────────────────────────────────
# Componenti riusabili
# ──────────────────────────────────────────────────────────────────────────────

def _card(parent: tk.Widget, padx: int = 18, pady: int = 14) -> tuple[tk.Frame, tk.Frame]:
    """Card bianca con bordo sottile + simulazione ombra."""
    shadow = tk.Frame(parent, bg=_SHADOW)
    inner  = tk.Frame(shadow, bg=_CARD, padx=padx, pady=pady)
    inner.pack(fill="both", expand=True, padx=(0, 1), pady=(0, 1))
    return shadow, inner


def _card_header(card: tk.Frame, icon: str, title: str, hint: str = "") -> tk.Frame:
    """Barra intestazione di una card con icona + titolo + hint opzionale."""
    hdr = tk.Frame(card, bg=_CARD_HDR, padx=0, pady=8)
    hdr.pack(fill="x", pady=(0, 10))
    tk.Frame(hdr, bg=_BORDER, height=1).pack(fill="x", side="bottom")

    inner = tk.Frame(hdr, bg=_CARD_HDR)
    inner.pack(fill="x", padx=18)
    tk.Label(inner, text=f"{icon}  {title}", font=_FT_H2,
             bg=_CARD_HDR, fg=_TEXT).pack(side="left")
    if hint:
        tk.Label(inner, text=hint, font=_FT_XS,
                 bg=_CARD_HDR, fg=_TEXT_MUTED).pack(side="left", padx=(10, 0))
    return hdr


def _field_label(parent: tk.Frame, text: str):
    tk.Label(parent, text=text, font=_FT_H3,
             bg=_CARD, fg=_TEXT_SEC).pack(anchor="w", pady=(8, 2))


def _hint_label(parent: tk.Frame, text: str):
    tk.Label(parent, text=text, font=_FT_XS,
             bg=_CARD, fg=_TEXT_MUTED, anchor="w").pack(fill="x", pady=(1, 6))


def _divider(parent: tk.Frame):
    tk.Frame(parent, bg=_BORDER, height=1).pack(fill="x", pady=10)


def _icon_button(parent: tk.Frame, icon: str, label: str,
                 command, style: str = "secondary") -> tk.Button:
    """Pulsante con icona + testo. style = 'primary' | 'secondary' | 'ghost'."""
    cfg = {
        "primary":   dict(bg=_PRIMARY,    fg=_WHITE,     ab=_PRIMARY_DK, af=_WHITE),
        "secondary": dict(bg=_WHITE,      fg=_TEXT,      ab="#EEF2F7",   af=_TEXT),
        "ghost":     dict(bg=_BG,         fg=_TEXT_SEC,  ab=_BORDER,     af=_TEXT),
    }[style]
    btn = tk.Button(
        parent, text=f"  {icon}  {label}  ",
        font=_FT_B if style == "primary" else _FT,
        bg=cfg["bg"], fg=cfg["fg"],
        activebackground=cfg["ab"], activeforeground=cfg["af"],
        relief="flat", cursor="hand2", bd=0,
        padx=4, pady=8 if style == "primary" else 6,
        command=command,
    )
    # Hover effect
    btn.bind("<Enter>", lambda _: btn.config(bg=cfg["ab"], fg=cfg["af"]))
    btn.bind("<Leave>", lambda _: btn.config(bg=cfg["bg"], fg=cfg["fg"]))
    return btn


# ──────────────────────────────────────────────────────────────────────────────
# Applicazione
# ──────────────────────────────────────────────────────────────────────────────

class FlowCheckApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("FlowCheck")
        self.geometry("960x820")
        self.minsize(820, 660)
        self.configure(bg=_BG)
        self.resizable(True, True)

        self._q:                queue.Queue[str | None] = queue.Queue()
        self._running:          bool  = False
        self._last_output_dir:  str | None = None
        self._adv_visible:      bool  = False
        self._prog_t_run_start: float | None = None
        self._prog_t_file_start:float | None = None
        self._prog_timer_id:    str | None = None
        self._prog_total:       int   = 1
        self._prog_done:        int   = 0

        self._setup_styles()
        self._build_ui()
        self._poll_queue()

    # ── ttk styles ──────────────────────────────────────────────────────────

    def _setup_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")

        # Entry
        s.configure("FC.TEntry",
                     fieldbackground=_WHITE,
                     bordercolor=_BORDER,
                     lightcolor=_BORDER,
                     darkcolor=_BORDER,
                     relief="solid",
                     padding=(10, 7))
        s.map("FC.TEntry",
              bordercolor=[("focus", _PRIMARY), ("hover", "#A0AEC0")],
              lightcolor =[("focus", _PRIMARY)],
              darkcolor  =[("focus", _PRIMARY)])

        # Combobox
        s.configure("FC.TCombobox",
                     fieldbackground=_WHITE,
                     bordercolor=_BORDER,
                     padding=(8, 6))
        s.map("FC.TCombobox",
              bordercolor=[("focus", _PRIMARY)])

        # Progress bar (in corso)
        s.configure("FC.Horizontal.TProgressbar",
                     troughcolor=_BORDER,
                     background=_PRIMARY,
                     thickness=8,
                     bordercolor=_BORDER,
                     lightcolor=_PRIMARY,
                     darkcolor=_PRIMARY)
        # Progress bar (completata)
        s.configure("FCDone.Horizontal.TProgressbar",
                     troughcolor=_BORDER,
                     background=_SUCCESS,
                     thickness=8,
                     bordercolor=_BORDER,
                     lightcolor=_SUCCESS,
                     darkcolor=_SUCCESS)

    # ── UI principale ───────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Header ──────────────────────────────────────────────────────
        self._build_header()

        # ── Status bar (fondo) ──────────────────────────────────────────
        self._status_var = tk.StringVar(value="Pronto")
        status = tk.Frame(self, bg=_BORDER, height=1)
        status.pack(fill="x", side="bottom")
        self._statusbar = tk.Label(
            self, textvariable=self._status_var,
            font=_FT_XS, bg=_CARD_HDR, fg=_TEXT_MUTED,
            anchor="w", padx=20, pady=5)
        self._statusbar.pack(fill="x", side="bottom")

        # ── Corpo scrollabile ────────────────────────────────────────────
        canvas = tk.Canvas(self, bg=_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        self._body = tk.Frame(canvas, bg=_BG)
        win_id = canvas.create_window((0, 0), window=self._body, anchor="nw")

        def _on_frame_configure(_e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def _on_canvas_configure(e):
            canvas.itemconfig(win_id, width=e.width)

        self._body.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(-1*(e.delta//120), "units"))

        self._build_body(self._body)

    def _build_header(self):
        hdr = tk.Frame(self, bg=_WHITE)
        hdr.pack(fill="x")
        tk.Frame(hdr, bg=_BORDER, height=1).pack(fill="x", side="bottom")

        inner = tk.Frame(hdr, bg=_WHITE, pady=14)
        inner.pack(fill="x", padx=24)

        # Logo / titolo
        logo_frame = tk.Frame(inner, bg=_WHITE)
        logo_frame.pack(side="left")

        tk.Label(logo_frame, text="FlowCheck",
                 font=_FT_H1, bg=_WHITE, fg=_TEXT).pack(anchor="w")
        tk.Label(logo_frame,
                 text="Confronta automaticamente due versioni di file CSV o ZIP",
                 font=_FT_XS, bg=_WHITE, fg=_TEXT_MUTED).pack(anchor="w")

        # Badge versione
        badge = tk.Frame(inner, bg=_PRIMARY_LT, padx=10, pady=4)
        badge.pack(side="right", anchor="n")
        tk.Label(badge, text="v2.0", font=("Segoe UI", 8, "bold"),
                 bg=_PRIMARY_LT, fg=_PRIMARY).pack()

    def _build_body(self, parent: tk.Frame):
        pad = dict(padx=24, pady=8)

        # ── Card 1: File ──────────────────────────────────────────────
        sh, card = _card(parent)
        sh.pack(fill="x", **pad)
        _card_header(card, "📂", "Carica i file")

        _field_label(card, "Versione attuale  (AS-IS)")
        self._asis_var = tk.StringVar()
        self._build_file_row(card, self._asis_var)
        _hint_label(card, "File CSV, archivio ZIP o cartella con i dati di partenza")

        _divider(card)

        _field_label(card, "Nuova versione  (TO-BE)")
        self._tobe_var = tk.StringVar()
        self._build_file_row(card, self._tobe_var)
        _hint_label(card, "File CSV, archivio ZIP o cartella con i nuovi dati da confrontare")

        _divider(card)

        _field_label(card, "Dove salvare i risultati  (facoltativo)")
        out_row = tk.Frame(card, bg=_CARD)
        out_row.pack(fill="x")
        self._out_var = tk.StringVar()
        ttk.Entry(out_row, textvariable=self._out_var,
                  style="FC.TEntry", font=_FT).pack(
            side="left", fill="x", expand=True, padx=(0, 10))
        _icon_button(out_row, "📁", "Sfoglia",
                     lambda: self._browse_dir(self._out_var),
                     "secondary").pack(side="left")
        _hint_label(card,
            "Lascia vuoto per salvare nella stessa cartella dell'AS-IS")

        # ── Toggle avanzate ───────────────────────────────────────────
        tgl_frame = tk.Frame(parent, bg=_BG)
        tgl_frame.pack(fill="x", padx=24, pady=(0, 4))

        self._adv_btn = tk.Button(
            tgl_frame,
            text="▶   Impostazioni avanzate  —  separatore e chiave di collegamento",
            font=("Segoe UI", 9), bg=_BG, fg=_TEXT_SEC,
            relief="flat", cursor="hand2", anchor="w",
            activebackground=_BG, activeforeground=_PRIMARY,
            padx=0, pady=4,
            command=self._toggle_advanced,
        )
        self._adv_btn.pack(side="left")
        self._adv_btn.bind("<Enter>", lambda _: self._adv_btn.config(fg=_PRIMARY))
        self._adv_btn.bind("<Leave>", lambda _: self._adv_btn.config(fg=_TEXT_SEC))

        # ── Card avanzate (nascosta) ───────────────────────────────────
        self._adv_shadow, self._adv_card = _card(parent)
        self._build_advanced(self._adv_card)
        self._adv_row_widget = self._adv_shadow   # per show/hide

        # ── Card azioni ───────────────────────────────────────────────
        act_sh, act_card = _card(parent, padx=18, pady=14)
        act_sh.pack(fill="x", padx=24, pady=(4, 0))
        self._build_actions(act_card)

        # ── Card avanzamento ──────────────────────────────────────────
        prog_sh, prog_card = _card(parent, padx=18, pady=12)
        prog_sh.pack(fill="x", padx=24, pady=(8, 0))
        self._build_progress(prog_card)

        # ── Card log ──────────────────────────────────────────────────
        log_sh, log_card = _card(parent, padx=0, pady=0)
        log_sh.pack(fill="both", expand=True, padx=24, pady=(8, 16))
        self._build_log(log_card)

    # ── Componenti sezione ───────────────────────────────────────────────────

    def _build_file_row(self, parent: tk.Frame, var: tk.StringVar):
        row = tk.Frame(parent, bg=_CARD)
        row.pack(fill="x", pady=(0, 2))
        ttk.Entry(row, textvariable=var, style="FC.TEntry",
                  font=_FT).pack(side="left", fill="x", expand=True, padx=(0, 10))
        _icon_button(row, "📄", "File / ZIP",
                     lambda: self._browse_file(var), "secondary").pack(
            side="left", padx=(0, 6))
        _icon_button(row, "📁", "Cartella",
                     lambda: self._browse_dir(var), "secondary").pack(side="left")

    def _build_advanced(self, parent: tk.Frame):
        _card_header(parent, "⚙", "Impostazioni avanzate",
                     "— modifica solo se il rilevamento automatico non funziona")

        # Separatore
        row1 = tk.Frame(parent, bg=_CARD)
        row1.pack(fill="x", pady=(0, 4))
        tk.Label(row1, text="Separatore campi:", font=_FT_H3,
                 bg=_CARD, fg=_TEXT_SEC, width=20, anchor="w").pack(side="left")

        self._sep_combo_var = tk.StringVar(value=SEP_OPTIONS[0][0])
        sep_cb = ttk.Combobox(row1, textvariable=self._sep_combo_var,
                              values=[o[0] for o in SEP_OPTIONS],
                              state="readonly", width=28,
                              style="FC.TCombobox", font=_FT)
        sep_cb.pack(side="left", padx=(0, 12))
        sep_cb.bind("<<ComboboxSelected>>", self._on_sep_change)

        self._sep_custom_var = tk.StringVar()
        self._sep_custom_entry = ttk.Entry(
            row1, textvariable=self._sep_custom_var,
            width=8, style="FC.TEntry", font=_FT, state="disabled")
        self._sep_custom_entry.pack(side="left")
        tk.Label(row1, text=" ← solo per «Personalizzato»",
                 font=_FT_XS, bg=_CARD, fg=_TEXT_MUTED).pack(side="left")

        _divider(parent)

        # Chiave di collegamento
        row2 = tk.Frame(parent, bg=_CARD)
        row2.pack(fill="x")
        tk.Label(row2, text="Chiave di collegamento:", font=_FT_H3,
                 bg=_CARD, fg=_TEXT_SEC, width=20, anchor="w").pack(side="left")

        self._jk_mode_var = tk.StringVar(value="auto")
        tk.Radiobutton(row2, text="Automatica",
                       variable=self._jk_mode_var, value="auto",
                       bg=_CARD, font=_FT, fg=_TEXT_SEC,
                       selectcolor=_WHITE, activebackground=_CARD,
                       command=self._on_jk_mode_change).pack(side="left")
        tk.Radiobutton(row2, text="Personalizzata:",
                       variable=self._jk_mode_var, value="manual",
                       bg=_CARD, font=_FT, fg=_TEXT_SEC,
                       selectcolor=_WHITE, activebackground=_CARD,
                       command=self._on_jk_mode_change).pack(side="left", padx=(14, 0))

        self._jk_var = tk.StringVar()
        self._jk_entry = ttk.Entry(row2, textvariable=self._jk_var,
                                   width=30, style="FC.TEntry", font=_FT,
                                   state="disabled")
        self._jk_entry.pack(side="left", padx=(6, 0))

        _hint_label(parent,
            "Campo (o campi separati da virgola) usato per abbinare le righe.  "
            "Es.: POLIZZA   oppure   POLIZZA,TIPO_MOV")

    def _build_actions(self, parent: tk.Frame):
        left = tk.Frame(parent, bg=_CARD)
        left.pack(fill="x")

        self._run_btn = _icon_button(
            left, "▶", "Avvia confronto",
            self._start_comparison, "primary")
        self._run_btn.pack(side="left")

        self._open_btn = _icon_button(
            left, "📂", "Apri risultati",
            self._open_output_folder, "secondary")
        self._open_btn.configure(state="disabled")
        self._open_btn.pack(side="left", padx=(10, 0))

        self._clear_btn = _icon_button(
            left, "✕", "Pulisci log",
            self._clear_log, "ghost")
        self._clear_btn.pack(side="right")

    def _build_progress(self, parent: tk.Frame):
        # Riga superiore: titolo + percentuale
        top = tk.Frame(parent, bg=_CARD)
        top.pack(fill="x", pady=(0, 6))

        tk.Label(top, text="Avanzamento", font=_FT_H3,
                 bg=_CARD, fg=_TEXT_SEC).pack(side="left")
        self._prog_pct_label = tk.Label(
            top, text="", font=("Segoe UI", 9, "bold"),
            bg=_CARD, fg=_PRIMARY)
        self._prog_pct_label.pack(side="right")

        # Barra
        self._prog_var = tk.DoubleVar(value=0)
        self._progressbar = ttk.Progressbar(
            parent, variable=self._prog_var,
            maximum=100, mode="determinate",
            style="FC.Horizontal.TProgressbar")
        self._progressbar.pack(fill="x", pady=(0, 6))

        # Riga inferiore: file corrente + timer
        bot = tk.Frame(parent, bg=_CARD)
        bot.pack(fill="x")
        self._prog_file_label = tk.Label(
            bot, text="In attesa…", font=_FT_XS, bg=_CARD, fg=_TEXT_MUTED, anchor="w")
        self._prog_file_label.pack(side="left")
        self._prog_time_label = tk.Label(
            bot, text="", font=_FT_M, bg=_CARD, fg=_TEXT_MUTED, anchor="e")
        self._prog_time_label.pack(side="right")

    def _build_log(self, parent: tk.Frame):
        # Header del log
        log_hdr = tk.Frame(parent, bg="#1B2333", pady=6)
        log_hdr.pack(fill="x")
        tk.Label(log_hdr, text="  📋  Registro elaborazione",
                 font=("Segoe UI", 9, "bold"),
                 bg="#1B2333", fg="#8B949E").pack(side="left")

        self._log = ScrolledText(
            parent, font=_FT_M,
            bg="#161B27", fg="#CDD6F4",
            insertbackground="#CDD6F4",
            selectbackground="#2D4F7C",
            relief="flat", wrap="word",
            state="disabled", height=12,
            padx=14, pady=10,
        )
        self._log.pack(fill="both", expand=True)

        for tag, col in [
            ("ok",   "#A6E3A1"),
            ("err",  "#F38BA8"),
            ("warn", "#F9E2AF"),
            ("info", "#89B4FA"),
            ("dim",  "#6C7086"),
        ]:
            self._log.tag_configure(tag, foreground=col)

    # ── Toggle avanzate ────────────────────────────────────────────────────

    def _toggle_advanced(self):
        self._adv_visible = not self._adv_visible
        if self._adv_visible:
            self._adv_row_widget.pack(fill="x", padx=24, pady=(0, 4))
            self._adv_btn.config(
                text="▼   Impostazioni avanzate  —  separatore e chiave di collegamento")
        else:
            self._adv_row_widget.pack_forget()
            self._adv_btn.config(
                text="▶   Impostazioni avanzate  —  separatore e chiave di collegamento")

    # ── Event handler ─────────────────────────────────────────────────────

    def _browse_file(self, var: tk.StringVar):
        p = filedialog.askopenfilename(
            filetypes=[("CSV / ZIP", "*.csv *.zip"), ("Tutti i file", "*.*")])
        if p:
            var.set(p)

    def _browse_dir(self, var: tk.StringVar):
        p = filedialog.askdirectory()
        if p:
            var.set(p)

    def _on_sep_change(self, _e=None):
        if self._sep_combo_var.get().startswith("Personalizzato"):
            self._sep_custom_entry.configure(state="normal")
            self._sep_custom_entry.focus()
        else:
            self._sep_custom_entry.configure(state="disabled")

    def _on_jk_mode_change(self):
        if self._jk_mode_var.get() == "manual":
            self._jk_entry.configure(state="normal")
            self._jk_entry.focus()
        else:
            self._jk_entry.configure(state="disabled")

    def _get_sep(self) -> str | None:
        sel = self._sep_combo_var.get()
        if sel.startswith("Rilevamento"):
            return None
        if sel.startswith("Personalizzato"):
            return self._sep_custom_var.get() or None
        for label, val in SEP_OPTIONS:
            if label == sel:
                return val
        return None

    def _get_join_key(self) -> list[str] | None:
        if self._jk_mode_var.get() != "manual":
            return None
        cols = [c.strip() for c in self._jk_var.get().split(",") if c.strip()]
        return cols or None

    # ── Log ───────────────────────────────────────────────────────────────

    def _log_write(self, msg: str):
        # ── Messaggi strutturati (non scrivono nel log testuale) ──────────
        if msg.startswith("[FILE_START]"):
            try:
                _, progress, *parts = msg.split()
                done_n, total_n = map(int, progress.split("/"))
                label = " ".join(parts)
                self._prog_total        = total_n
                self._prog_done         = done_n - 1
                self._prog_t_file_start = time.perf_counter()
                self._progressbar.stop()
                self._progressbar.configure(
                    mode="indeterminate",
                    style="FC.Horizontal.TProgressbar")
                self._progressbar.start(10)
                pct_done = ((done_n - 1) / total_n * 100) if total_n else 0
                self._prog_pct_label.configure(
                    text=f"File {done_n} di {total_n}  ({pct_done:.0f}%)",
                    fg=_PRIMARY)
                self._prog_file_label.configure(
                    text=f"▶  {label}")
            except Exception:
                pass
            return

        if msg.startswith("[PROGRESS]"):
            try:
                _, progress = msg.split()
                done, total = map(int, progress.split("/"))
                self._prog_done  = done
                self._prog_total = total
                self._progressbar.stop()
                pct = (done / total * 100) if total else 100
                style = "FCDone.Horizontal.TProgressbar" if done == total \
                        else "FC.Horizontal.TProgressbar"
                self._progressbar.configure(mode="determinate", style=style)
                self._prog_var.set(pct)
                self._prog_pct_label.configure(
                    text=f"File {done} di {total}  ({pct:.0f}%)",
                    fg=_SUCCESS if done == total else _PRIMARY)
                self._prog_file_label.configure(
                    text="" if done == total
                    else f"✔  Completato  —  {total - done} rimasti")
            except Exception:
                pass
            return

        # ── Log testuale ─────────────────────────────────────────────────
        self._log.configure(state="normal")
        ml  = msg.lower()
        tag = ""
        if "[ok]" in ml:                                tag = "ok"
        elif "[errore]" in ml or "traceback" in ml:     tag = "err"
        elif "[attenzione]" in ml or "[skip]" in ml:    tag = "warn"
        elif (msg.startswith("AS-IS") or msg.startswith("TO-BE")
              or msg.startswith("Cop") or msg.startswith("Chi")):
            tag = "info"
        elif msg.startswith("  ") or msg.startswith("---"):
            tag = "dim"
        self._log.insert("end", msg + "\n", tag)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _clear_log(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    def _open_output_folder(self):
        if self._last_output_dir and Path(self._last_output_dir).exists():
            if sys.platform == "win32":
                os.startfile(self._last_output_dir)
            else:
                subprocess.Popen(["xdg-open", self._last_output_dir])

    # ── Timer live ─────────────────────────────────────────────────────────

    @staticmethod
    def _fmt(secs: float) -> str:
        secs = max(0.0, secs)
        if secs < 60:
            return f"{secs:.0f}s"
        m, s = divmod(int(secs), 60)
        return f"{m}m {s:02d}s"

    def _tick_timer(self):
        if not self._running:
            return
        now    = time.perf_counter()
        t_run  = now - self._prog_t_run_start  if self._prog_t_run_start  else 0.0
        t_file = now - self._prog_t_file_start if self._prog_t_file_start else 0.0
        self._prog_time_label.configure(
            text=f"⏱  file {self._fmt(t_file)}  ·  totale {self._fmt(t_run)}")
        self._prog_timer_id = self.after(500, self._tick_timer)

    # ── Stato run ──────────────────────────────────────────────────────────

    def _set_running(self, running: bool):
        self._running = running
        if running:
            self._run_btn.configure(
                state="disabled",
                text="  ⏳  Elaborazione in corso…  ",
                bg="#94B8E8", fg=_WHITE)
            self._run_btn.unbind("<Enter>")
            self._run_btn.unbind("<Leave>")
            self._status_var.set("Elaborazione in corso…")
            # Reset progress
            self._prog_var.set(0)
            self._prog_pct_label.configure(text="", fg=_PRIMARY)
            self._prog_file_label.configure(text="Avvio…")
            self._prog_time_label.configure(text="")
            self._progressbar.configure(
                mode="indeterminate",
                style="FC.Horizontal.TProgressbar")
            self._progressbar.start(10)
            # Timer
            self._prog_t_run_start  = time.perf_counter()
            self._prog_t_file_start = time.perf_counter()
            self._prog_timer_id     = self.after(500, self._tick_timer)
        else:
            if self._prog_timer_id:
                self.after_cancel(self._prog_timer_id)
                self._prog_timer_id = None
            self._progressbar.stop()
            self._progressbar.configure(
                mode="determinate",
                style="FCDone.Horizontal.TProgressbar")
            self._prog_var.set(100)
            self._prog_pct_label.configure(text="Completato ✔", fg=_SUCCESS)
            self._prog_file_label.configure(text="")
            # Ripristina bottone con hover
            self._run_btn.configure(
                state="normal",
                text="  ▶  Avvia confronto  ",
                bg=_PRIMARY, fg=_WHITE)
            self._run_btn.bind("<Enter>",
                lambda _: self._run_btn.config(bg=_PRIMARY_DK))
            self._run_btn.bind("<Leave>",
                lambda _: self._run_btn.config(bg=_PRIMARY))
            self._status_var.set("✔  Completato")
            if self._last_output_dir:
                self._open_btn.configure(state="normal")

    # ── Avvio confronto ────────────────────────────────────────────────────

    def _start_comparison(self):
        asis     = self._asis_var.get().strip()
        tobe     = self._tobe_var.get().strip()
        out      = self._out_var.get().strip() or None
        sep      = self._get_sep()
        join_key = self._get_join_key()

        if not asis:
            messagebox.showwarning(
                "File mancante",
                "Seleziona il file (o cartella) con la versione attuale (AS-IS).")
            return
        if not tobe:
            messagebox.showwarning(
                "File mancante",
                "Seleziona il file (o cartella) con la nuova versione (TO-BE).")
            return

        self._clear_log()
        self._set_running(True)
        self._last_output_dir = out or str(
            Path(asis).parent if Path(asis).is_file() else asis)

        threading.Thread(
            target=self._worker,
            args=(asis, tobe, out, sep, join_key),
            daemon=True,
        ).start()

    def _worker(self, asis: str, tobe: str, out: str | None,
                sep: str | None, join_key: list[str] | None):
        try:
            from flowcheck_engine import run_comparison
            generated = run_comparison(
                asis_path=asis,
                tobe_path=tobe,
                output_dir=out,
                sep=sep,
                join_key=join_key,
                progress_cb=self._q.put,
            )
            if out is None and generated:
                self._last_output_dir = str(Path(generated[0]).parent)
        except Exception as exc:
            import traceback
            self._q.put(f"[ERRORE] {exc}")
            self._q.put(traceback.format_exc())
        finally:
            self._q.put(None)

    # ── Poll coda ──────────────────────────────────────────────────────────

    def _poll_queue(self):
        try:
            while True:
                msg = self._q.get_nowait()
                if msg is None:
                    self._set_running(False)
                else:
                    self._log_write(msg)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)


# ──────────────────────────────────────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────────────────────────────────────

def main():
    app = FlowCheckApp()
    app.mainloop()


if __name__ == "__main__":
    main()
