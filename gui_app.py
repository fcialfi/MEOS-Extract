"""Tkinter front end for processing MEOS HTML reports.

The window is organised in "cards": input folders (with the number of HTML
reports found in each), output folder, processing options and a run panel
with progress bar. Processing runs in a worker thread so the window stays
responsive; log records and progress updates are passed back to the Tk
thread through a queue and shown in a colour-coded log panel.

Settings (folders, output, options) are remembered between sessions in
``~/.meos_extract_gui.json``.
"""

from pathlib import Path
import json
import logging
import logging.handlers
import os
import queue
import subprocess
import sys
import threading
import time
from tkinter import Tk, filedialog, StringVar, BooleanVar, Text, messagebox, PhotoImage
from tkinter import font as tkfont
from tkinter import ttk

from Extract_all_charts import process_html, generate_combined_polar_plot_artifacts
import pandas as pd


LOG_PATH = Path(__file__).resolve().with_name("gui_app.log")
SETTINGS_PATH = Path.home() / ".meos_extract_gui.json"

handlers = [logging.StreamHandler()]
try:
    handlers.insert(0, logging.FileHandler(LOG_PATH))
except OSError as exc:  # pragma: no cover - log fallback
    print(f"Cannot write log file: {exc}")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=handlers,
    force=True,
)

# --- Palette -------------------------------------------------------------
BG = "#EEF2F7"
CARD = "#FFFFFF"
BORDER = "#D8DEE9"
TEXT = "#1E293B"
MUTED = "#64748B"
ACCENT = "#2563EB"
ACCENT_HOVER = "#1D4ED8"
ACCENT_PRESSED = "#1E40AF"
HEADER = "#0F172A"
HEADER_SUB = "#94A3B8"
SUCCESS = "#16A34A"
DANGER = "#DC2626"
LOG_BG = "#0F172A"
LOG_FG = "#E2E8F0"

PLOT_OPTIONS = [
    ("input_level", "Input Level"),
    ("eb_no", "Eb/N0"),
    ("snr", "Signal/Noise Ratio"),
]


# -------------------- Batch processing (no Tk here) --------------------

def html_reports(folder: Path):
    """HTML reports contained directly in *folder*."""
    try:
        return sorted(Path(folder).glob("*.html"))
    except OSError:
        return []


def run_batch(folders, output_dir: Path, selected_stats, selected_plots,
              make_individual_plots=True, make_combined_plots=False,
              progress=None, stop_event=None):
    """Process every HTML report in *folders* and write the summary files.

    ``progress(done, total, report)`` is called before each report.
    Returns a summary dict (saved files, errors, stopped flag, extra outputs).
    """
    output_dir = Path(output_dir)
    reports = []
    for folder in folders:
        found = html_reports(folder)
        if not found:
            logging.warning("Nessun file HTML in %s", folder)
        reports.extend(found)

    summary = {"saved": [], "errors": [], "stopped": False, "extra": [], "total": len(reports)}
    stats_rows, plot_rows, plot_series_rows = [], [], []

    for i, report in enumerate(reports):
        if stop_event is not None and stop_event.is_set():
            summary["stopped"] = True
            logging.warning("Elaborazione interrotta dall'utente")
            break
        if progress:
            progress(i, len(reports), report)
        try:
            out = process_html(
                report,
                output_dir,
                stats_selectors=selected_stats,
                stats_rows=stats_rows,
                plot_selectors=selected_plots,
                plot_rows=plot_rows,
                plot_series_rows=plot_series_rows,
                generate_individual_plots=make_individual_plots,
            )
            if out in summary["saved"]:
                logging.warning("%s ha la stessa orbita di un report precedente: %s sovrascritto", report.name, out.name)
            logging.info("Salvato: %s", out)
            summary["saved"].append(out)
        except Exception:
            # Keep going: one broken report must not discard the rest
            # of the batch (statistics and plots included).
            logging.exception("Errore durante l'elaborazione di %s", report)
            summary["errors"].append(report)

    if progress:
        progress(len(summary["saved"]) + len(summary["errors"]), len(reports), None)

    if selected_stats and summary["saved"]:
        stats_path = output_dir / "lock_state_stats.xlsx"
        if stats_rows:
            pd.DataFrame(stats_rows, columns=["Orbit Number", "Unlocks"]).to_excel(stats_path, index=False)
        else:
            pd.DataFrame([{"Orbit Number": "N/A", "Unlocks": 0}]).to_excel(stats_path, index=False)
        logging.info("Salvate statistiche: %s", stats_path)
        summary["extra"].append(stats_path)

    if selected_plots and summary["saved"]:
        if make_combined_plots:
            plot_rows.extend(
                generate_combined_polar_plot_artifacts(output_dir, plot_series_rows, selected_plots)
            )
        plot_index = output_dir / "polar_plots_index.xlsx"
        if plot_rows:
            pd.DataFrame(plot_rows).to_excel(plot_index, index=False)
            logging.info("Salvato indice plot polari: %s (%d plot)", plot_index, len(plot_rows))
            summary["extra"].append(plot_index)
        else:
            logging.warning(
                "Plot polari richiesti ma nessuno generato: controllare i warning nel log."
            )

    return summary


# -------------------- GUI helpers --------------------

def open_path(path: Path):
    """Open *path* with the system file manager."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as exc:  # pragma: no cover - platform specific
        logging.error("Impossibile aprire %s: %s", path, exc)


def load_settings():
    try:
        return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_settings(data):
    try:
        SETTINGS_PATH.write_text(json.dumps(data, indent=2), encoding="utf-8")
    except OSError:
        pass


def pick_font_family(root):
    families = set(tkfont.families(root))
    for name in ("Segoe UI", "SF Pro Text", "Helvetica Neue", "Inter", "Ubuntu", "Noto Sans", "DejaVu Sans"):
        if name in families:
            return name
    return tkfont.nametofont("TkDefaultFont").actual("family")


def pick_mono_family(root):
    families = set(tkfont.families(root))
    for name in ("Cascadia Mono", "Consolas", "SF Mono", "Menlo", "Ubuntu Mono", "DejaVu Sans Mono"):
        if name in families:
            return name
    return "TkFixedFont"


def _checkbox_image(root, checked, enabled=True, size=16):
    """Draw a rounded checkbox indicator as a PhotoImage (no external assets)."""
    img = PhotoImage(master=root, width=size, height=size)
    fill = (ACCENT if enabled else "#93C5FD") if checked else ("#FFFFFF" if enabled else "#F1F5F9")
    border = fill if checked else ("#94A3B8" if enabled else "#CBD5E1")
    img.put(fill, to=(1, 1, size - 1, size - 1))
    img.put(border, to=(2, 0, size - 2, 1))
    img.put(border, to=(2, size - 1, size - 2, size))
    img.put(border, to=(0, 2, 1, size - 2))
    img.put(border, to=(size - 1, 2, size, size - 2))
    for x, y in ((1, 1), (size - 2, 1), (1, size - 2), (size - 2, size - 2)):
        img.put(border, to=(x, y, x + 1, y + 1))
    if checked:
        # check mark: short stroke down-right, long stroke up-right (2 px thick)
        pts = [(4, 8), (5, 9), (6, 10), (7, 11)] + [(8, 10), (9, 9), (10, 8), (11, 7), (12, 6), (13, 5)]
        for x, y in pts:
            img.put("#FFFFFF", to=(x - 1, y - 1, x + 1, y + 1))
    return img


def setup_style(root):
    family = pick_font_family(root)
    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
        try:
            tkfont.nametofont(name).configure(family=family, size=10)
        except Exception:
            pass

    style = ttk.Style(root)
    style.theme_use("clam")
    root.configure(bg=BG)

    style.configure(".", background=BG, foreground=TEXT, font=(family, 10))
    style.configure("TFrame", background=BG)
    style.configure("Card.TFrame", background=CARD, relief="solid", borderwidth=1, bordercolor=BORDER)
    style.configure("Inner.TFrame", background=CARD)
    style.configure("Header.TFrame", background=HEADER)

    style.configure("HeaderTitle.TLabel", background=HEADER, foreground="#FFFFFF", font=(family, 18, "bold"))
    style.configure("HeaderSub.TLabel", background=HEADER, foreground=HEADER_SUB, font=(family, 10))
    style.configure("CardTitle.TLabel", background=CARD, foreground=TEXT, font=(family, 11, "bold"))
    style.configure("Card.TLabel", background=CARD, foreground=TEXT)
    style.configure("Muted.TLabel", background=CARD, foreground=MUTED, font=(family, 9))
    style.configure("Section.TLabel", background=CARD, foreground=MUTED, font=(family, 9, "bold"))
    style.configure("Status.TLabel", background=CARD, foreground=MUTED, font=(family, 10))
    style.configure("StatusOk.TLabel", background=CARD, foreground=SUCCESS, font=(family, 10, "bold"))
    style.configure("StatusErr.TLabel", background=CARD, foreground=DANGER, font=(family, 10, "bold"))

    imgs = {
        "off": _checkbox_image(root, False), "on": _checkbox_image(root, True),
        "off_dis": _checkbox_image(root, False, False), "on_dis": _checkbox_image(root, True, False),
    }
    root._checkbox_images = imgs  # keep references alive
    style.element_create(
        "Card.Checkbutton.indicator", "image", imgs["off"],
        ("disabled", "selected", imgs["on_dis"]), ("disabled", imgs["off_dis"]),
        ("selected", imgs["on"]), border=0, sticky="w", width=24,
    )
    style.layout("Card.TCheckbutton", [
        ("Checkbutton.padding", {"sticky": "nswe", "children": [
            ("Card.Checkbutton.indicator", {"side": "left", "sticky": ""}),
            ("Checkbutton.label", {"side": "left", "sticky": "nswe"}),
        ]}),
    ])
    style.configure("Card.TCheckbutton", background=CARD, foreground=TEXT, padding=(0, 2))
    style.map("Card.TCheckbutton", background=[("active", CARD)], foreground=[("disabled", "#A0AEC0")])

    style.configure("TButton", padding=(12, 6), background="#F1F5F9", foreground=TEXT,
                    bordercolor=BORDER, lightcolor="#F1F5F9", darkcolor="#F1F5F9", focuscolor="#F1F5F9")
    style.map("TButton",
              background=[("disabled", "#F8FAFC"), ("pressed", "#E2E8F0"), ("active", "#E2E8F0")],
              foreground=[("disabled", "#A0AEC0")])

    style.configure("Accent.TButton", padding=(18, 10), background=ACCENT, foreground="#FFFFFF",
                    bordercolor=ACCENT, lightcolor=ACCENT, darkcolor=ACCENT, focuscolor=ACCENT,
                    font=(family, 11, "bold"))
    style.map("Accent.TButton",
              background=[("disabled", "#93C5FD"), ("pressed", ACCENT_PRESSED), ("active", ACCENT_HOVER)],
              lightcolor=[("disabled", "#93C5FD"), ("pressed", ACCENT_PRESSED), ("active", ACCENT_HOVER)],
              darkcolor=[("disabled", "#93C5FD"), ("pressed", ACCENT_PRESSED), ("active", ACCENT_HOVER)],
              bordercolor=[("disabled", "#93C5FD")],
              foreground=[("disabled", "#EFF6FF")])

    style.configure("Danger.TButton", padding=(12, 10), background="#FEE2E2", foreground=DANGER,
                    bordercolor="#FECACA", lightcolor="#FEE2E2", darkcolor="#FEE2E2", focuscolor="#FEE2E2")
    style.map("Danger.TButton",
              background=[("disabled", "#F8FAFC"), ("active", "#FECACA")],
              foreground=[("disabled", "#A0AEC0")])

    style.configure("TEntry", fieldbackground="#F8FAFC", bordercolor=BORDER, lightcolor=BORDER, padding=6)

    style.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT,
                    rowheight=28, bordercolor=BORDER, borderwidth=0)
    style.configure("Treeview.Heading", background="#F1F5F9", foreground=MUTED,
                    font=(family, 9, "bold"), relief="flat", padding=(8, 6))
    style.map("Treeview", background=[("selected", "#DBEAFE")], foreground=[("selected", TEXT)])
    style.map("Treeview.Heading", background=[("active", "#E2E8F0")])

    style.configure("Accent.Horizontal.TProgressbar", troughcolor="#E2E8F0", background=ACCENT,
                    bordercolor="#E2E8F0", lightcolor=ACCENT, darkcolor=ACCENT, thickness=10)
    style.configure("Vertical.TScrollbar", background="#E2E8F0", troughcolor=CARD,
                    bordercolor=CARD, arrowcolor=MUTED, lightcolor="#E2E8F0", darkcolor="#E2E8F0")
    return family


def card(parent, title, subtitle=None):
    """Create a white card with a title; returns (outer, body)."""
    outer = ttk.Frame(parent, style="Card.TFrame", padding=14)
    head = ttk.Frame(outer, style="Inner.TFrame")
    head.pack(fill="x")
    ttk.Label(head, text=title, style="CardTitle.TLabel").pack(side="left")
    if subtitle:
        ttk.Label(head, text=subtitle, style="Muted.TLabel").pack(side="left", padx=(8, 0), pady=(2, 0))
    body = ttk.Frame(outer, style="Inner.TFrame")
    body.pack(fill="both", expand=True, pady=(10, 0))
    return outer, body, head


class QueueLogHandler(logging.Handler):
    """Forward log records to the Tk thread through a queue."""

    def __init__(self, q):
        super().__init__()
        self.q = q

    def emit(self, record):
        try:
            self.q.put(("log", record.levelno, self.format(record)))
        except Exception:  # pragma: no cover
            self.handleError(record)


# -------------------- Application --------------------

class App:
    def __init__(self, root):
        self.root = root
        self.queue = queue.Queue()
        self.worker = None
        self.stop_event = threading.Event()
        self.started_at = None
        self.lockable = []

        self.family = setup_style(root)
        self.mono = pick_mono_family(root)

        root.title("MEOS Extract")
        root.geometry("1120x800")
        root.minsize(960, 700)

        settings = load_settings()
        self.output_var = StringVar(value=settings.get("output", ""))
        self.stat_unlock = BooleanVar(value=settings.get("stat_unlock", False))
        self.plot_vars = {key: BooleanVar(value=key in settings.get("plots", [])) for key, _ in PLOT_OPTIONS}
        self.per_file_var = BooleanVar(value=settings.get("per_file", True))
        self.combined_var = BooleanVar(value=settings.get("combined", False))
        self.status_var = StringVar(value="Pronto")
        self.count_var = StringVar()

        self._build()
        for folder in settings.get("folders", []):
            if Path(folder).is_dir():
                self._add_folder(folder, log=False)
        self._update_count()
        self._update_plot_mode_state()
        self.output_entry.xview_moveto(1.0)

        log_handler = QueueLogHandler(self.queue)
        log_handler.setFormatter(logging.Formatter("%(asctime)s  %(levelname)-7s  %(message)s", "%H:%M:%S"))
        logging.getLogger().addHandler(log_handler)

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.bind("<Escape>", lambda e: self.on_close())
        root.bind("<Control-Return>", lambda e: self.start())
        root.after(100, self._poll_queue)
        logging.info("MEOS Extract pronto. Aggiungi le cartelle con i report HTML e premi Avvia.")

    # --- Layout -----------------------------------------------------------
    def _build(self):
        root = self.root

        header = ttk.Frame(root, style="Header.TFrame", padding=(24, 16))
        header.pack(fill="x")
        ttk.Label(header, text="MEOS Extract", style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="Estrazione dei grafici dai report HTML MEOS in Excel, statistiche di lock e plot polari",
            style="HeaderSub.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        content = ttk.Frame(root, padding=16)
        content.pack(fill="both", expand=True)
        content.grid_columnconfigure(0, weight=3, uniform="col")
        content.grid_columnconfigure(1, weight=2, uniform="col")
        content.grid_rowconfigure(0, weight=3, minsize=470)
        content.grid_rowconfigure(1, weight=1, minsize=150)

        # Left column: inputs + output
        left = ttk.Frame(content)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 8))
        left.grid_columnconfigure(0, weight=1)
        left.grid_rowconfigure(0, weight=1)
        self._build_inputs(left).grid(row=0, column=0, sticky="nsew")
        self._build_output(left).grid(row=1, column=0, sticky="ew", pady=(12, 0))

        # Right column: options + run
        right = ttk.Frame(content)
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=(0, 8))
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(0, weight=1)
        self._build_options(right).grid(row=0, column=0, sticky="nsew")
        self._build_run(right).grid(row=1, column=0, sticky="ew", pady=(12, 0))

        # Bottom: log
        self._build_log(content).grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 0))

    def _build_inputs(self, parent):
        outer, body, head = card(parent, "Cartelle di input")
        ttk.Label(head, textvariable=self.count_var, style="Muted.TLabel").pack(side="right")

        tree_frame = ttk.Frame(body, style="Inner.TFrame")
        tree_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_frame, columns=("folder", "reports"), show="headings", selectmode="extended")
        self.tree.heading("folder", text="Cartella", anchor="w")
        self.tree.heading("reports", text="Report HTML", anchor="center")
        self.tree.column("folder", anchor="w", width=380, stretch=True)
        self.tree.column("reports", anchor="center", width=110, stretch=False)
        self.tree.tag_configure("empty", foreground=DANGER)
        self.tree.tag_configure("odd", background="#F8FAFC")
        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tree.bind("<Delete>", lambda e: self.remove_selected())
        self.tree.bind("<BackSpace>", lambda e: self.remove_selected())
        self.tree.bind("<Double-1>", self._open_tree_item)

        btns = ttk.Frame(body, style="Inner.TFrame")
        btns.pack(fill="x", pady=(10, 0))
        for text, cmd in (
            ("+ Aggiungi cartella", self.add_folder),
            ("Aggiungi sottocartelle…", self.add_subfolders),
            ("Rimuovi", self.remove_selected),
            ("Svuota", self.clear_folders),
        ):
            b = ttk.Button(btns, text=text, command=cmd)
            b.pack(side="left", padx=(0, 6))
            self.lockable.append(b)
        return outer

    def _build_output(self, parent):
        outer, body, _ = card(parent, "Cartella di output")
        body.grid_columnconfigure(0, weight=1)
        entry = self.output_entry = ttk.Entry(body, textvariable=self.output_var)
        entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        b_browse = ttk.Button(body, text="Sfoglia…", command=self.select_output)
        b_browse.grid(row=0, column=1, padx=(0, 6))
        b_open = ttk.Button(body, text="Apri", command=self.open_output)
        b_open.grid(row=0, column=2)
        self.lockable.extend([entry, b_browse])
        return outer

    def _build_options(self, parent):
        outer, body, _ = card(parent, "Opzioni")

        ttk.Label(body, text="STATISTICHE", style="Section.TLabel").pack(anchor="w")
        cb = ttk.Checkbutton(body, text="Conta gli unlock del demodulatore (1→0→1)",
                             variable=self.stat_unlock, style="Card.TCheckbutton")
        cb.pack(anchor="w", pady=(4, 0))
        ttk.Label(body, text="→ lock_state_stats.xlsx (Orbit Number, Unlocks)",
                  style="Muted.TLabel").pack(anchor="w", padx=(24, 0))
        self.lockable.append(cb)

        ttk.Separator(body).pack(fill="x", pady=10)
        ttk.Label(body, text="PLOT POLARI / 3D", style="Section.TLabel").pack(anchor="w")
        plots_row = ttk.Frame(body, style="Inner.TFrame")
        plots_row.pack(anchor="w", pady=(4, 0))
        for key, label in PLOT_OPTIONS:
            cb = ttk.Checkbutton(plots_row, text=label, variable=self.plot_vars[key], style="Card.TCheckbutton",
                                 command=self._update_plot_mode_state)
            cb.pack(side="left", padx=(0, 14))
            self.lockable.append(cb)
        ttk.Label(body, text="SNR: i campioni unlocked sono evidenziati in viola",
                  style="Muted.TLabel").pack(anchor="w", pady=(2, 0))

        ttk.Separator(body).pack(fill="x", pady=10)
        ttk.Label(body, text="MODALITÀ PLOT", style="Section.TLabel").pack(anchor="w")
        self.mode_checks = []
        for text, var in (("Un set per ogni report", self.per_file_var),
                          ("Un set combinato per tutti i report", self.combined_var)):
            cb = ttk.Checkbutton(body, text=text, variable=var, style="Card.TCheckbutton")
            cb.pack(anchor="w", pady=(2, 0))
            self.mode_checks.append(cb)
        return outer

    def _build_run(self, parent):
        outer = ttk.Frame(parent, style="Card.TFrame", padding=14)
        row = ttk.Frame(outer, style="Inner.TFrame")
        row.pack(fill="x")
        self.btn_run = ttk.Button(row, text="▶  Avvia elaborazione", style="Accent.TButton", command=self.start)
        self.btn_run.pack(side="left", fill="x", expand=True)
        self.btn_stop = ttk.Button(row, text="■  Stop", style="Danger.TButton", command=self.stop, state="disabled")
        self.btn_stop.pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(outer, style="Accent.Horizontal.TProgressbar", mode="determinate")
        self.progress.pack(fill="x", pady=(14, 6))
        status_row = ttk.Frame(outer, style="Inner.TFrame")
        status_row.pack(fill="x")
        self.status_label = ttk.Label(status_row, textvariable=self.status_var, style="Status.TLabel")
        self.status_label.pack(side="left", fill="x", expand=True)
        self.btn_open_out = ttk.Button(status_row, text="Apri risultati", command=self.open_output, state="disabled")
        self.btn_open_out.pack(side="right")
        return outer

    def _build_log(self, parent):
        outer, body, head = card(parent, "Log")
        ttk.Button(head, text="Pulisci", command=self.clear_log).pack(side="right")
        ttk.Button(head, text="Apri file di log", command=lambda: open_path(LOG_PATH)).pack(side="right", padx=(0, 6))

        self.log = Text(body, height=8, bg=LOG_BG, fg=LOG_FG, insertbackground=LOG_FG, relief="flat",
                        borderwidth=0, padx=10, pady=8, font=(self.mono, 9), wrap="none", state="disabled")
        scroll = ttk.Scrollbar(body, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.log.tag_configure("WARNING", foreground="#FBBF24")
        self.log.tag_configure("ERROR", foreground="#F87171")
        self.log.tag_configure("SUCCESS", foreground="#4ADE80")
        self.log.tag_configure("DEBUG", foreground="#64748B")
        return outer

    # --- Input folders ----------------------------------------------------
    def _folders(self):
        return list(self.tree.get_children())  # item ids are the full paths

    @staticmethod
    def _short_path(path, limit=58):
        return path if len(path) <= limit else "…" + path[-(limit - 1):]

    def _add_folder(self, folder, log=True):
        folder = str(Path(folder))
        if folder in self._folders():
            if log:
                logging.info("Cartella già presente: %s", folder)
            return False
        n = len(html_reports(Path(folder)))
        tags = ["empty"] if n == 0 else []
        if len(self.tree.get_children()) % 2:
            tags.append("odd")
        self.tree.insert("", "end", iid=folder, values=(self._short_path(folder), n if n else "nessuno"), tags=tags)
        if log:
            logging.info("Aggiunta cartella %s (%d report)", folder, n)
        return True

    def _restripe(self):
        for i, item in enumerate(self.tree.get_children()):
            tags = [t for t in self.tree.item(item, "tags") if t != "odd"]
            if i % 2:
                tags.append("odd")
            self.tree.item(item, tags=tags)

    def _update_count(self):
        folders = self._folders()
        n_reports = sum(len(html_reports(Path(f))) for f in folders)
        self.count_var.set(
            f"{len(folders)} cartell{'a' if len(folders) == 1 else 'e'} · {n_reports} report"
        )

    def add_folder(self):
        folder = filedialog.askdirectory(title="Seleziona una cartella con report HTML", mustexist=True)
        if folder:
            self._add_folder(folder)
            self._update_count()

    def add_subfolders(self):
        parent = filedialog.askdirectory(title="Seleziona la cartella padre da scansionare", mustexist=True)
        if not parent:
            return
        parent = Path(parent)
        candidates = [parent] + sorted(p for p in parent.rglob("*") if p.is_dir())
        added = sum(self._add_folder(p, log=False) for p in candidates if html_reports(p))
        logging.info("Scansione di %s: aggiunte %d cartelle con report HTML", parent, added)
        if not added:
            logging.warning("Nessuna nuova cartella con report HTML trovata in %s", parent)
        self._restripe()
        self._update_count()

    def remove_selected(self):
        for item in self.tree.selection():
            self.tree.delete(item)
        self._restripe()
        self._update_count()

    def clear_folders(self):
        self.tree.delete(*self.tree.get_children())
        self._update_count()

    def _open_tree_item(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            open_path(Path(item))

    # --- Output -------------------------------------------------------------
    def select_output(self):
        folder = filedialog.askdirectory(title="Seleziona la cartella di output")
        if folder:
            self.output_var.set(str(Path(folder)))
            self.output_entry.xview_moveto(1.0)

    def open_output(self):
        out = self.output_var.get().strip()
        if out and Path(out).is_dir():
            open_path(Path(out))
        else:
            logging.warning("La cartella di output non esiste ancora")

    def _update_plot_mode_state(self):
        any_plot = any(v.get() for v in self.plot_vars.values())
        for cb in self.mode_checks:
            cb.state(["!disabled"] if any_plot and not self._running() else ["disabled"])

    # --- Run ---------------------------------------------------------------
    def _running(self):
        return self.worker is not None and self.worker.is_alive()

    def _set_running(self, running):
        state = ["disabled"] if running else ["!disabled"]
        for w in self.lockable:
            w.state(state)
        self.btn_run.state(state)
        self.btn_stop.state(["!disabled"] if running else ["disabled"])
        self._update_plot_mode_state()

    def start(self):
        if self._running():
            return
        folders = self._folders()
        output = self.output_var.get().strip()
        selected_stats = ["demodulator_lock_state"] if self.stat_unlock.get() else []
        selected_plots = [k for k, _ in PLOT_OPTIONS if self.plot_vars[k].get()]
        per_file, combined = self.per_file_var.get(), self.combined_var.get()

        if not folders:
            messagebox.showwarning("MEOS Extract", "Aggiungi almeno una cartella con report HTML.")
            return
        if not output:
            messagebox.showwarning("MEOS Extract", "Seleziona la cartella di output.")
            return
        if selected_plots and not (per_file or combined):
            messagebox.showwarning("MEOS Extract", "Seleziona almeno una modalità plot (per report e/o combinato).")
            return

        self._save_settings()
        self.stop_event.clear()
        self.progress.configure(value=0, maximum=1)
        self.btn_open_out.state(["disabled"])
        self.status_label.configure(style="Status.TLabel")
        self.status_var.set("Avvio…")
        self.started_at = time.monotonic()
        self._set_running(True)
        logging.info("Avvio elaborazione: %d cartelle → %s", len(folders), output)

        def progress(done, total, report):
            self.queue.put(("progress", done, total, report))

        def work():
            try:
                summary = run_batch(folders, Path(output), selected_stats, selected_plots,
                                    per_file, combined, progress=progress, stop_event=self.stop_event)
            except Exception as exc:
                logging.exception("Errore inatteso")
                summary = {"saved": [], "errors": [], "stopped": False, "extra": [], "total": 0, "fatal": str(exc)}
            self.queue.put(("done", summary))

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def stop(self):
        if self._running():
            self.stop_event.set()
            self.btn_stop.state(["disabled"])
            self.status_var.set("Interruzione dopo il report corrente…")

    def _finish(self, summary):
        self._set_running(False)
        elapsed = time.monotonic() - (self.started_at or time.monotonic())
        saved, errors = len(set(summary["saved"])), len(summary["errors"])
        if summary.get("fatal"):
            self.status_label.configure(style="StatusErr.TLabel")
            self.status_var.set(f"Errore: {summary['fatal']}")
        elif summary["stopped"]:
            self.status_label.configure(style="StatusErr.TLabel")
            self.status_var.set(f"Interrotto · {saved} Excel creati in {elapsed:.0f}s")
        elif errors:
            self.status_label.configure(style="StatusErr.TLabel")
            self.status_var.set(f"Completato con {errors} errori · {saved} Excel creati in {elapsed:.0f}s")
        else:
            self.status_label.configure(style="StatusOk.TLabel")
            self.status_var.set(f"✓ Completato · {saved} Excel creati in {elapsed:.0f}s")
        self.progress.configure(value=self.progress["maximum"])
        logging.info("Completato: %d Excel creati, %d errori", saved, errors)
        if saved:
            self.btn_open_out.state(["!disabled"])

    # --- Queue / log -------------------------------------------------------
    def _poll_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                kind = msg[0]
                if kind == "log":
                    self._append_log(msg[1], msg[2])
                elif kind == "progress":
                    _, done, total, report = msg
                    self.progress.configure(maximum=max(total, 1), value=done)
                    if report is not None:
                        self.status_var.set(f"Report {done + 1} di {total} · {Path(report).name}")
                elif kind == "done":
                    self._finish(msg[1])
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _append_log(self, levelno, text):
        if levelno >= logging.ERROR:
            tag = "ERROR"
        elif levelno >= logging.WARNING:
            tag = "WARNING"
        elif levelno < logging.INFO:
            tag = "DEBUG"
        elif "Salvat" in text or "Completato" in text:
            tag = "SUCCESS"
        else:
            tag = None
        at_end = self.log.yview()[1] >= 0.999
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n", tag)
        lines = int(self.log.index("end-1c").split(".")[0])
        if lines > 5000:
            self.log.delete("1.0", f"{lines - 5000}.0")
        self.log.configure(state="disabled")
        if at_end:
            self.log.see("end")

    def clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    # --- Settings / close ---------------------------------------------------
    def _save_settings(self):
        save_settings({
            "folders": self._folders(),
            "output": self.output_var.get().strip(),
            "stat_unlock": self.stat_unlock.get(),
            "plots": [k for k, _ in PLOT_OPTIONS if self.plot_vars[k].get()],
            "per_file": self.per_file_var.get(),
            "combined": self.combined_var.get(),
        })

    def on_close(self):
        if self._running():
            if not messagebox.askyesno("MEOS Extract", "Elaborazione in corso. Uscire comunque?"):
                return
            self.stop_event.set()
        self._save_settings()
        self.root.destroy()


def main():
    # License validation removed
    if sys.platform.startswith("win"):
        try:  # crisp text on high-DPI Windows displays
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
