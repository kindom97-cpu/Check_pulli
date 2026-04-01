# flowcheck_app.py
# Interfaccia grafica tkinter per FlowCheck — design user-friendly

from __future__ import annotations

import os
import queue
import sys
import threading
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

# ---------------------------------------------------------------------------
# Palette colori
# ---------------------------------------------------------------------------
_BG         = "#F0F4F8"   # sfondo generale
_CARD       = "#FFFFFF"   # sfondo card
_HEADER_BG  = "#1E3A5F"   # intestazione
_HEADER_SUB = "#2E5F9A"
_PRIMARY    = "#1A6BBF"   # bottone principale
_PRIMARY_HO = "#155799"   # hover
_SUCCESS    = "#1A7A4A"
_MUTED      = "#718096"   # testo secondario
_BORDER     = "#CBD5E0"   # bordo card
_TEXT       = "#2D3748"
_WHITE      = "#FFFFFF"

_F_TITLE  = ("Segoe UI", 16, "bold")
_F_SUB    = ("Segoe UI",  9)
_F_LABEL  = ("Segoe UI", 10)
_F_BOLD   = ("Segoe UI", 10, "bold")
_F_SECT   = ("Segoe UI", 10, "bold")
_F_HINT   = ("Segoe UI",  8)
_F_MONO   = ("Consolas",  9)
_F_BTN    = ("Segoe UI", 10, "bold")

SEP_OPTIONS = [
    ("Rilevamento automatico",  None),
    (";  (punto e virgola)",    ";"),
    (",  (virgola)",            ","),
    ("\\t  (tabulazione)",      "\t"),
    ("|  (pipe)",               "|"),
    (";|  (composto)",          ";|"),
    (";£  (composto £)",        ";£"),
    ("Personalizzato…",         "__custom__"),
]


# ---------------------------------------------------------------------------
# Utilità: card con bordo arrotondato simulato
# ---------------------------------------------------------------------------

def _card(parent, **kw) -> tk.Frame:
    """Frame bianco con bordo sottile — simula una card."""
    outer = tk.Frame(parent, bg=_BORDER, padx=1, pady=1)
    inner = tk.Frame(outer, bg=_CARD, **kw)
    inner.pack(fill="both", expand=True)
    return outer, inner


def _hint(parent, text: str):
    """Etichetta descrittiva grigia sotto un campo."""
    tk.Label(parent, text=text, font=_F_HINT, bg=_CARD,
             fg=_MUTED, anchor="w").pack(fill="x", padx=2, pady=(0, 6))


def _section_label(parent, text: str):
    tk.Label(parent, text=text, font=_F_SECT, bg=_CARD,
             fg=_TEXT).pack(anchor="w", pady=(10, 2))


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

class FlowCheckApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("FlowCheck")
        self.geometry("920x740")
        self.minsize(800, 620)
        self.configure(bg=_BG)
        self.resizable(True, True)

        self._q: queue.Queue[str | None] = queue.Queue()
        self._running = False
        self._last_output_dir: str | None = None
        self._adv_visible = False   # pannello opzioni avanzate

        self._apply_style()
        self._build_ui()
        self._poll_queue()

    # ------------------------------------------------------------------
    # ttk Style
    # ------------------------------------------------------------------

    def _apply_style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("TEntry",    fieldbackground=_WHITE, bordercolor=_BORDER,
                    relief="flat", padding=5)
        s.configure("TCombobox", fieldbackground=_WHITE, bordercolor=_BORDER,
                    padding=5)
        s.configure("TButton",   background=_BG, relief="flat", padding=(8, 4))
        s.map("TButton",
              background=[("active", "#E2E8F0"), ("disabled", "#EDF2F7")],
              foreground=[("disabled", "#A0AEC0")])

    # ------------------------------------------------------------------
    # Costruzione UI
    # ------------------------------------------------------------------

    def _build_ui(self):
        # ── Header ────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=_HEADER_BG, height=68)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        tk.Label(hdr, text="FlowCheck", font=_F_TITLE,
                 bg=_HEADER_BG, fg=_WHITE).pack(side="left", padx=20, pady=(12, 2),
                                                anchor="sw")
        tk.Label(hdr,
                 text="Confronta automaticamente due versioni di file CSV o ZIP",
                 font=_F_SUB, bg=_HEADER_BG, fg="#A8C4E0").pack(
            side="left", padx=(4, 0), pady=(0, 4), anchor="sw")

        # ── Scroll contenitore principale ─────────────────────────────
        outer = tk.Frame(self, bg=_BG)
        outer.pack(fill="both", expand=True, padx=18, pady=14)
        outer.columnconfigure(0, weight=1)

        row = 0

        # ── Card 1: Scegli i file ─────────────────────────────────────
        f_outer, f_inner = _card(outer, padx=16, pady=6)
        f_outer.grid(row=row, column=0, sticky="ew")
        outer.rowconfigure(row, weight=0)
        row += 1

        tk.Label(f_inner, text="1  Scegli i file", font=("Segoe UI", 11, "bold"),
                 bg=_CARD, fg=_HEADER_BG).pack(anchor="w", pady=(6, 0))

        # AS-IS
        _section_label(f_inner, "Versione attuale  (AS-IS)")
        self._asis_var = tk.StringVar()
        self._build_file_row(f_inner, self._asis_var)
        _hint(f_inner, "File CSV, archivio ZIP o cartella contenente i dati di partenza")

        # Separatore visivo
        tk.Frame(f_inner, bg=_BORDER, height=1).pack(fill="x", pady=4)

        # TO-BE
        _section_label(f_inner, "Nuova versione  (TO-BE)")
        self._tobe_var = tk.StringVar()
        self._build_file_row(f_inner, self._tobe_var)
        _hint(f_inner, "File CSV, archivio ZIP o cartella con i nuovi dati da confrontare")

        # Separatore visivo
        tk.Frame(f_inner, bg=_BORDER, height=1).pack(fill="x", pady=4)

        # Output
        _section_label(f_inner, "Dove salvare i risultati  (facoltativo)")
        out_row = tk.Frame(f_inner, bg=_CARD)
        out_row.pack(fill="x")
        self._out_var = tk.StringVar()
        ttk.Entry(out_row, textvariable=self._out_var,
                  font=_F_LABEL).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(out_row, text="📁  Sfoglia",
                   command=lambda: self._browse_dir(self._out_var)).pack(side="left")
        _hint(f_inner, "Se non indicato, i file vengono salvati nella stessa cartella dell'AS-IS")

        # ── Pannello avanzate (toggle) ─────────────────────────────────
        adv_toggle_row = tk.Frame(outer, bg=_BG)
        adv_toggle_row.grid(row=row, column=0, sticky="ew", pady=(8, 0))
        row += 1

        self._adv_btn = tk.Button(
            adv_toggle_row,
            text="▶  Impostazioni avanzate  (separatore, chiave di collegamento)",
            font=_F_HINT, bg=_BG, fg=_PRIMARY, relief="flat",
            cursor="hand2", activebackground=_BG, activeforeground=_PRIMARY_HO,
            command=self._toggle_advanced,
        )
        self._adv_btn.pack(anchor="w")

        # Card opzioni avanzate (nascosta di default)
        self._adv_outer, self._adv_inner = _card(outer, padx=16, pady=8)
        # NON grid: viene mostrata/nascosta da _toggle_advanced
        self._adv_row = row          # riga da usare se visible
        row += 1

        self._build_advanced(self._adv_inner)

        # ── Bottoni azione ─────────────────────────────────────────────
        btn_card_outer, btn_card = _card(outer, padx=16, pady=10)
        btn_card_outer.grid(row=row, column=0, sticky="ew", pady=(10, 0))
        row += 1

        btn_left = tk.Frame(btn_card, bg=_CARD)
        btn_left.pack(side="left", fill="x", expand=True)

        self._run_btn = tk.Button(
            btn_left,
            text="▶   Avvia confronto",
            font=_F_BTN, bg=_PRIMARY, fg=_WHITE,
            activebackground=_PRIMARY_HO, activeforeground=_WHITE,
            relief="flat", padx=22, pady=8, cursor="hand2",
            command=self._start_comparison,
        )
        self._run_btn.pack(side="left")

        self._open_btn = tk.Button(
            btn_left,
            text="📂  Apri risultati",
            font=_F_LABEL, bg=_BG, fg=_TEXT,
            activebackground="#E2E8F0", relief="flat",
            padx=14, pady=8, cursor="hand2", state="disabled",
            command=self._open_output_folder,
        )
        self._open_btn.pack(side="left", padx=10)

        self._clear_btn = tk.Button(
            btn_left,
            text="✕  Pulisci log",
            font=_F_HINT, bg=_BG, fg=_MUTED,
            activebackground="#E2E8F0", relief="flat",
            padx=10, pady=8, cursor="hand2",
            command=self._clear_log,
        )
        self._clear_btn.pack(side="right")

        # ── Log ────────────────────────────────────────────────────────
        log_outer, log_inner = _card(outer, padx=0, pady=0)
        log_outer.grid(row=row, column=0, sticky="nsew", pady=(10, 0))
        outer.rowconfigure(row, weight=1)
        row += 1

        log_hdr = tk.Frame(log_inner, bg="#161B22", height=28)
        log_hdr.pack(fill="x")
        log_hdr.pack_propagate(False)
        tk.Label(log_hdr, text="  Registro elaborazione",
                 font=_F_HINT, bg="#161B22", fg="#8B949E",
                 anchor="w").pack(fill="both", expand=True, padx=4)

        self._log = ScrolledText(
            log_inner, font=_F_MONO, bg="#0D1117", fg="#C9D1D9",
            insertbackground="white", relief="flat",
            wrap="word", state="disabled", height=10,
        )
        self._log.pack(fill="both", expand=True)

        for tag, color in [("ok",   "#3FB950"), ("err",  "#F85149"),
                           ("warn", "#D29922"), ("info", "#58A6FF"),
                           ("dim",  "#8B949E")]:
            self._log.tag_configure(tag, foreground=color)

        # ── Status bar ─────────────────────────────────────────────────
        self._status_var = tk.StringVar(value="✔  Pronto")
        tk.Label(self, textvariable=self._status_var,
                 font=_F_HINT, bg=_HEADER_BG, fg="#A8C4E0",
                 anchor="w", padx=14).pack(fill="x", side="bottom")

    # ------------------------------------------------------------------
    # File row helper
    # ------------------------------------------------------------------

    def _build_file_row(self, parent: tk.Frame, var: tk.StringVar):
        row = tk.Frame(parent, bg=_CARD)
        row.pack(fill="x", pady=(0, 2))
        ttk.Entry(row, textvariable=var,
                  font=_F_LABEL).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(row, text="📄  File / ZIP",
                   command=lambda: self._browse_file(var)).pack(side="left", padx=(0, 4))
        ttk.Button(row, text="📁  Cartella",
                   command=lambda: self._browse_dir(var)).pack(side="left")

    # ------------------------------------------------------------------
    # Pannello avanzate
    # ------------------------------------------------------------------

    def _build_advanced(self, parent: tk.Frame):
        tk.Label(parent, text="Impostazioni avanzate", font=("Segoe UI", 10, "bold"),
                 bg=_CARD, fg=_HEADER_BG).pack(anchor="w", pady=(4, 8))

        # Separatore
        sep_frame = tk.Frame(parent, bg=_CARD)
        sep_frame.pack(fill="x", pady=(0, 4))

        tk.Label(sep_frame, text="Separatore campi:", font=_F_LABEL,
                 bg=_CARD, fg=_TEXT, width=22, anchor="w").pack(side="left")

        self._sep_combo_var = tk.StringVar(value=SEP_OPTIONS[0][0])
        sep_cb = ttk.Combobox(sep_frame, textvariable=self._sep_combo_var,
                              values=[o[0] for o in SEP_OPTIONS],
                              state="readonly", width=26, font=_F_LABEL)
        sep_cb.pack(side="left", padx=(0, 10))
        sep_cb.bind("<<ComboboxSelected>>", self._on_sep_change)

        self._sep_custom_var = tk.StringVar()
        self._sep_custom_entry = ttk.Entry(sep_frame, textvariable=self._sep_custom_var,
                                           width=8, font=_F_LABEL, state="disabled")
        self._sep_custom_entry.pack(side="left")
        tk.Label(sep_frame, text="← solo per «Personalizzato»",
                 font=_F_HINT, bg=_CARD, fg=_MUTED).pack(side="left", padx=6)

        _hint(parent,
              "Di solito il rilevamento automatico funziona correttamente. "
              "Cambia solo se i dati vengono letti in modo errato.")

        tk.Frame(parent, bg=_BORDER, height=1).pack(fill="x", pady=6)

        # Chiave di collegamento
        jk_label_row = tk.Frame(parent, bg=_CARD)
        jk_label_row.pack(fill="x", pady=(0, 4))
        tk.Label(jk_label_row, text="Chiave di collegamento:", font=_F_LABEL,
                 bg=_CARD, fg=_TEXT, width=22, anchor="w").pack(side="left")

        self._jk_mode_var = tk.StringVar(value="auto")

        tk.Radiobutton(jk_label_row, text="Automatica",
                       variable=self._jk_mode_var, value="auto",
                       bg=_CARD, font=_F_LABEL,
                       command=self._on_jk_mode_change).pack(side="left")
        tk.Radiobutton(jk_label_row, text="Personalizzata:",
                       variable=self._jk_mode_var, value="manual",
                       bg=_CARD, font=_F_LABEL,
                       command=self._on_jk_mode_change).pack(side="left", padx=(14, 0))

        self._jk_var = tk.StringVar()
        self._jk_entry = ttk.Entry(jk_label_row, textvariable=self._jk_var,
                                   width=32, font=_F_LABEL, state="disabled")
        self._jk_entry.pack(side="left", padx=(4, 8))

        _hint(parent,
              "Campo (o campi separati da virgola) usato per abbinare le righe tra i due file.  "
              "Es.: POLIZZA   oppure   POLIZZA,TIPO_MOV")

    # ------------------------------------------------------------------
    # Toggle pannello avanzate
    # ------------------------------------------------------------------

    def _toggle_advanced(self):
        self._adv_visible = not self._adv_visible
        if self._adv_visible:
            self._adv_outer.grid(row=self._adv_row, column=0,
                                 sticky="ew", pady=(4, 0))
            self._adv_btn.configure(
                text="▼  Impostazioni avanzate  (separatore, chiave di collegamento)")
        else:
            self._adv_outer.grid_remove()
            self._adv_btn.configure(
                text="▶  Impostazioni avanzate  (separatore, chiave di collegamento)")

    # ------------------------------------------------------------------
    # Helpers UI
    # ------------------------------------------------------------------

    def _browse_file(self, var: tk.StringVar):
        p = filedialog.askopenfilename(
            filetypes=[("CSV / ZIP", "*.csv *.zip"), ("Tutti i file", "*.*")])
        if p:
            var.set(p)

    def _browse_dir(self, var: tk.StringVar):
        p = filedialog.askdirectory()
        if p:
            var.set(p)

    def _on_sep_change(self, _event=None):
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
        raw = self._jk_var.get().strip()
        cols = [c.strip() for c in raw.split(",") if c.strip()]
        return cols if cols else None

    # ------------------------------------------------------------------
    # Log
    # ------------------------------------------------------------------

    def _log_write(self, msg: str):
        self._log.configure(state="normal")
        ml = msg.lower()
        tag = None
        if "[ok]" in ml:                                   tag = "ok"
        elif "[errore]" in ml or "traceback" in ml:        tag = "err"
        elif "[attenzione]" in ml or "[skip]" in ml:       tag = "warn"
        elif msg.startswith("AS-IS") or msg.startswith("TO-BE") \
             or msg.startswith("Cop") or msg.startswith("Chi"):
            tag = "info"
        elif msg.startswith("  ") or msg.startswith("---"): tag = "dim"
        self._log.insert("end", msg + "\n", tag or "")
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

    # ------------------------------------------------------------------
    # Stato bottone principale
    # ------------------------------------------------------------------

    def _set_running(self, running: bool):
        self._running = running
        if running:
            self._run_btn.configure(
                state="disabled", text="⏳  Elaborazione in corso…",
                bg="#6B9BD2", fg=_WHITE)
            self._status_var.set("⏳  Elaborazione in corso…")
        else:
            self._run_btn.configure(
                state="normal", text="▶   Avvia confronto",
                bg=_PRIMARY, fg=_WHITE)
            self._status_var.set("✔  Completato")
            if self._last_output_dir:
                self._open_btn.configure(state="normal")

    # ------------------------------------------------------------------
    # Avvio confronto
    # ------------------------------------------------------------------

    def _start_comparison(self):
        asis     = self._asis_var.get().strip()
        tobe     = self._tobe_var.get().strip()
        out      = self._out_var.get().strip() or None
        sep      = self._get_sep()
        join_key = self._get_join_key()

        if not asis:
            messagebox.showwarning(
                "File mancante",
                "Seleziona il file (o la cartella) con la versione attuale (AS-IS).")
            return
        if not tobe:
            messagebox.showwarning(
                "File mancante",
                "Seleziona il file (o la cartella) con la nuova versione (TO-BE).")
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
            self._q.put(None)
        except Exception as exc:
            import traceback
            self._q.put(f"[ERRORE] {exc}")
            self._q.put(traceback.format_exc())
            self._q.put(None)

    # ------------------------------------------------------------------
    # Poll coda messaggi (thread-safe)
    # ------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    app = FlowCheckApp()
    app.mainloop()


if __name__ == "__main__":
    main()
