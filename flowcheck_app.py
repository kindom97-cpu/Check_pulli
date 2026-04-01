# flowcheck_app.py
# Interfaccia grafica tkinter per FlowCheck
# Avvia il confronto CSV/ZIP in background e mostra il progresso in tempo reale

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
# Costanti UI
# ---------------------------------------------------------------------------
_APP_TITLE  = "FlowCheck  |  Confronto AS-IS vs TO-BE"
_BG_DARK    = "#1F3864"
_BG_MED     = "#2E4D87"
_BG_LIGHT   = "#F5F7FA"
_ACCENT     = "#2E75B6"
_FG_WHITE   = "#FFFFFF"
_FG_DARK    = "#1A1A2E"
_FG_OK      = "#276221"
_FG_ERR     = "#9C0006"
_FG_WARN    = "#9C5700"
_FONT_TITLE = ("Segoe UI", 15, "bold")
_FONT_LABEL = ("Segoe UI", 10)
_FONT_BOLD  = ("Segoe UI", 10, "bold")
_FONT_MONO  = ("Consolas", 9)
_FONT_SMALL = ("Segoe UI", 8)

SEP_OPTIONS = [
    ("Auto-detect", None),
    (";  (punto e virgola)", ";"),
    (",  (virgola)", ","),
    ("\\t  (tabulazione)", "\t"),
    ("|  (pipe)", "|"),
    (";|  (multi)", ";|"),
    (";£  (multi)", ";£"),
    ("Personalizzato...", "__custom__"),
]


# ---------------------------------------------------------------------------
# Applicazione principale
# ---------------------------------------------------------------------------

class FlowCheckApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(_APP_TITLE)
        self.geometry("900x700")
        self.minsize(780, 580)
        self.configure(bg=_BG_LIGHT)
        self.resizable(True, True)

        self._q: queue.Queue[str | None] = queue.Queue()
        self._running = False
        self._last_output_dir: str | None = None

        self._build_ui()
        self._poll_queue()

    # ------------------------------------------------------------------
    # Costruzione UI
    # ------------------------------------------------------------------

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=_BG_DARK, height=60)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text=_APP_TITLE, font=_FONT_TITLE,
                 bg=_BG_DARK, fg=_FG_WHITE).pack(side="left", padx=20, pady=10)

        # Body
        body = tk.Frame(self, bg=_BG_LIGHT, padx=20, pady=15)
        body.pack(fill="both", expand=True)

        # Grid a 3 colonne: label | entry | bottone(i)
        body.columnconfigure(1, weight=1)

        row = 0

        # -- AS-IS --
        tk.Label(body, text="AS-IS  (CSV / ZIP / Cartella):",
                 font=_FONT_BOLD, bg=_BG_LIGHT, fg=_FG_DARK).grid(
            row=row, column=0, sticky="w", pady=(0, 4))
        row += 1

        self._asis_var = tk.StringVar()
        self._asis_entry = ttk.Entry(body, textvariable=self._asis_var, font=_FONT_LABEL)
        self._asis_entry.grid(row=row, column=0, columnspan=2, sticky="ew", padx=(0, 6))

        btn_frame_a = tk.Frame(body, bg=_BG_LIGHT)
        btn_frame_a.grid(row=row, column=2, sticky="ew")
        ttk.Button(btn_frame_a, text="File/ZIP",
                   command=lambda: self._browse_file(self._asis_var)).pack(side="left", padx=(0, 4))
        ttk.Button(btn_frame_a, text="Cartella",
                   command=lambda: self._browse_dir(self._asis_var)).pack(side="left")
        row += 1

        tk.Frame(body, bg=_BG_LIGHT, height=10).grid(row=row, column=0)
        row += 1

        # -- TO-BE --
        tk.Label(body, text="TO-BE  (CSV / ZIP / Cartella):",
                 font=_FONT_BOLD, bg=_BG_LIGHT, fg=_FG_DARK).grid(
            row=row, column=0, sticky="w", pady=(0, 4))
        row += 1

        self._tobe_var = tk.StringVar()
        self._tobe_entry = ttk.Entry(body, textvariable=self._tobe_var, font=_FONT_LABEL)
        self._tobe_entry.grid(row=row, column=0, columnspan=2, sticky="ew", padx=(0, 6))

        btn_frame_b = tk.Frame(body, bg=_BG_LIGHT)
        btn_frame_b.grid(row=row, column=2, sticky="ew")
        ttk.Button(btn_frame_b, text="File/ZIP",
                   command=lambda: self._browse_file(self._tobe_var)).pack(side="left", padx=(0, 4))
        ttk.Button(btn_frame_b, text="Cartella",
                   command=lambda: self._browse_dir(self._tobe_var)).pack(side="left")
        row += 1

        tk.Frame(body, bg=_BG_LIGHT, height=10).grid(row=row, column=0)
        row += 1

        # -- Output dir --
        tk.Label(body, text="Cartella output:",
                 font=_FONT_BOLD, bg=_BG_LIGHT, fg=_FG_DARK).grid(
            row=row, column=0, sticky="w", pady=(0, 4))
        row += 1

        self._out_var = tk.StringVar()
        ttk.Entry(body, textvariable=self._out_var, font=_FONT_LABEL).grid(
            row=row, column=0, columnspan=2, sticky="ew", padx=(0, 6))
        ttk.Button(body, text="Sfoglia",
                   command=lambda: self._browse_dir(self._out_var)).grid(
            row=row, column=2, sticky="w")
        row += 1

        tk.Frame(body, bg=_BG_LIGHT, height=10).grid(row=row, column=0)
        row += 1

        # -- Separatore --
        sep_row = tk.Frame(body, bg=_BG_LIGHT)
        sep_row.grid(row=row, column=0, columnspan=3, sticky="w")
        tk.Label(sep_row, text="Separatore CSV:", font=_FONT_BOLD,
                 bg=_BG_LIGHT, fg=_FG_DARK).pack(side="left")

        self._sep_combo_var = tk.StringVar(value=SEP_OPTIONS[0][0])
        sep_combo = ttk.Combobox(sep_row, textvariable=self._sep_combo_var,
                                 values=[o[0] for o in SEP_OPTIONS],
                                 state="readonly", width=25, font=_FONT_LABEL)
        sep_combo.pack(side="left", padx=10)
        sep_combo.bind("<<ComboboxSelected>>", self._on_sep_change)

        tk.Label(sep_row, text="Valore:", font=_FONT_LABEL,
                 bg=_BG_LIGHT, fg=_FG_DARK).pack(side="left", padx=(10, 4))
        self._sep_custom_var = tk.StringVar(value="")
        self._sep_custom_entry = ttk.Entry(sep_row, textvariable=self._sep_custom_var,
                                           width=8, font=_FONT_LABEL, state="disabled")
        self._sep_custom_entry.pack(side="left")
        row += 1

        tk.Frame(body, bg=_BG_LIGHT, height=8).grid(row=row, column=0)
        row += 1

        # -- Chiave di join --
        jk_outer = tk.Frame(body, bg=_BG_LIGHT)
        jk_outer.grid(row=row, column=0, columnspan=3, sticky="ew")

        tk.Label(jk_outer, text="Chiave di join:", font=_FONT_BOLD,
                 bg=_BG_LIGHT, fg=_FG_DARK).pack(side="left")

        self._jk_mode_var = tk.StringVar(value="auto")

        tk.Radiobutton(
            jk_outer, text="Auto-detect", variable=self._jk_mode_var,
            value="auto", bg=_BG_LIGHT, font=_FONT_LABEL,
            command=self._on_jk_mode_change,
        ).pack(side="left", padx=(12, 0))

        tk.Radiobutton(
            jk_outer, text="Specifica campi:", variable=self._jk_mode_var,
            value="manual", bg=_BG_LIGHT, font=_FONT_LABEL,
            command=self._on_jk_mode_change,
        ).pack(side="left", padx=(8, 0))

        self._jk_var = tk.StringVar(value="")
        self._jk_entry = ttk.Entry(jk_outer, textvariable=self._jk_var,
                                   width=38, font=_FONT_LABEL, state="disabled")
        self._jk_entry.pack(side="left", padx=(4, 0))

        tk.Label(jk_outer, text="(es. POLIZZA  o  POLIZZA,NUM_CONTR)",
                 font=_FONT_SMALL, bg=_BG_LIGHT, fg="#888888").pack(side="left", padx=(8, 0))

        row += 1

        tk.Frame(body, bg=_BG_LIGHT, height=10).grid(row=row, column=0)
        row += 1

        # -- Bottone Avvia --
        btn_row = tk.Frame(body, bg=_BG_LIGHT)
        btn_row.grid(row=row, column=0, columnspan=3, sticky="ew")

        self._run_btn = tk.Button(
            btn_row, text="  Avvia Confronto  ",
            font=_FONT_BOLD, bg=_ACCENT, fg=_FG_WHITE,
            activebackground=_BG_MED, activeforeground=_FG_WHITE,
            relief="flat", padx=16, pady=6, cursor="hand2",
            command=self._start_comparison,
        )
        self._run_btn.pack(side="left")

        self._open_btn = tk.Button(
            btn_row, text="  Apri cartella output  ",
            font=_FONT_LABEL, bg="#E8ECF3", fg=_FG_DARK,
            activebackground="#D0D8E8", relief="flat", padx=12, pady=6,
            cursor="hand2", state="disabled",
            command=self._open_output_folder,
        )
        self._open_btn.pack(side="left", padx=12)

        self._clear_btn = tk.Button(
            btn_row, text="Pulisci log",
            font=_FONT_SMALL, bg="#E8ECF3", fg=_FG_DARK,
            relief="flat", padx=8, pady=6, cursor="hand2",
            command=self._clear_log,
        )
        self._clear_btn.pack(side="right")
        row += 1

        tk.Frame(body, bg=_BG_LIGHT, height=10).grid(row=row, column=0)
        row += 1

        # -- Log area --
        tk.Label(body, text="Log elaborazione:", font=_FONT_BOLD,
                 bg=_BG_LIGHT, fg=_FG_DARK).grid(
            row=row, column=0, columnspan=3, sticky="w")
        row += 1

        log_frame = tk.Frame(body, bg=_BG_LIGHT)
        log_frame.grid(row=row, column=0, columnspan=3, sticky="nsew")
        body.rowconfigure(row, weight=1)

        self._log = ScrolledText(
            log_frame, font=_FONT_MONO, bg="#0D1117", fg="#C9D1D9",
            insertbackground="white", relief="flat",
            wrap="word", state="disabled",
        )
        self._log.pack(fill="both", expand=True)

        # Tag colori nel log
        self._log.tag_configure("ok",   foreground="#3FB950")
        self._log.tag_configure("err",  foreground="#F85149")
        self._log.tag_configure("warn", foreground="#D29922")
        self._log.tag_configure("info", foreground="#58A6FF")
        self._log.tag_configure("dim",  foreground="#8B949E")
        row += 1

        # -- Status bar --
        self._status_var = tk.StringVar(value="Pronto.")
        status_bar = tk.Label(self, textvariable=self._status_var,
                              font=_FONT_SMALL, bg=_BG_DARK, fg=_FG_WHITE,
                              anchor="w", padx=10)
        status_bar.pack(fill="x", side="bottom")

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
        sel = self._sep_combo_var.get()
        if sel.startswith("Personalizzato"):
            self._sep_custom_entry.configure(state="normal")
        else:
            self._sep_custom_entry.configure(state="disabled")

    def _on_jk_mode_change(self):
        if self._jk_mode_var.get() == "manual":
            self._jk_entry.configure(state="normal")
            self._jk_entry.focus()
        else:
            self._jk_entry.configure(state="disabled")

    def _get_join_key(self) -> list[str] | None:
        """Restituisce None (auto) o la lista di nomi colonna specificati."""
        if self._jk_mode_var.get() != "manual":
            return None
        raw = self._jk_var.get().strip()
        if not raw:
            return None
        cols = [c.strip() for c in raw.split(",") if c.strip()]
        return cols if cols else None

    def _get_sep(self) -> str | None:
        sel = self._sep_combo_var.get()
        if sel.startswith("Auto"):
            return None
        if sel.startswith("Personalizzato"):
            return self._sep_custom_var.get() or None
        for label, val in SEP_OPTIONS:
            if label == sel:
                return val
        return None

    def _log_write(self, msg: str):
        self._log.configure(state="normal")
        tag = None
        ml = msg.lower()
        if "[ok]" in ml:
            tag = "ok"
        elif "[errore]" in ml or "error" in ml:
            tag = "err"
        elif "[skip]" in ml or "[attenzione]" in ml or "warn" in ml:
            tag = "warn"
        elif msg.startswith("AS-IS") or msg.startswith("TO-BE") or msg.startswith("Cop"):
            tag = "info"
        elif msg.startswith("  ") or msg.startswith("---"):
            tag = "dim"
        if tag:
            self._log.insert("end", msg + "\n", tag)
        else:
            self._log.insert("end", msg + "\n")
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

    def _set_running(self, running: bool):
        self._running = running
        state = "disabled" if running else "normal"
        self._run_btn.configure(
            state=state,
            text="  Elaborazione in corso...  " if running else "  Avvia Confronto  ",
            bg="#888" if running else _ACCENT,
        )
        if not running and self._last_output_dir:
            self._open_btn.configure(state="normal")

    # ------------------------------------------------------------------
    # Avvio confronto
    # ------------------------------------------------------------------

    def _start_comparison(self):
        asis = self._asis_var.get().strip()
        tobe = self._tobe_var.get().strip()
        out  = self._out_var.get().strip() or None
        sep  = self._get_sep()

        join_key = self._get_join_key()

        if not asis:
            messagebox.showwarning("Input mancante", "Specifica il percorso AS-IS.")
            return
        if not tobe:
            messagebox.showwarning("Input mancante", "Specifica il percorso TO-BE.")
            return

        self._clear_log()
        self._set_running(True)
        self._status_var.set("Elaborazione in corso...")
        self._last_output_dir = out or str(Path(asis).parent if Path(asis).is_file() else asis)

        t = threading.Thread(
            target=self._worker, args=(asis, tobe, out, sep, join_key), daemon=True)
        t.start()

    def _worker(self, asis: str, tobe: str, out: str | None,
                sep: str | None, join_key: list[str] | None):
        try:
            from flowcheck_engine import run_comparison
            def cb(msg):
                self._q.put(msg)
            generated = run_comparison(
                asis_path=asis,
                tobe_path=tobe,
                output_dir=out,
                sep=sep,
                join_key=join_key,
                progress_cb=cb,
            )
            if out is None:
                # aggiorna la cartella output rilevata
                if generated:
                    self._last_output_dir = str(Path(generated[0]).parent)
            self._q.put(None)  # segnale di fine
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
                    self._status_var.set("Completato.")
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
