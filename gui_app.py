"""Tkinter-based front end for processing MEOS HTML reports.

The application guides the user through selecting one or more folders
containing ``report.html`` files and an output directory in which to write
the extracted spreadsheets. Two vertically stacked panes divide the main
window: the upper pane hosts all user controls (output entry, list of input
folders and buttons) while the lower pane shows a live log of the operations
performed. Each button is wired to a callback that performs the associated
action and records messages through :mod:`logging`, which are displayed both
on the console and inside the GUI.
"""

from pathlib import Path
from tkinter import Tk, Listbox, filedialog, StringVar, Text, PanedWindow, messagebox
from tkinter import ttk
import logging
import math

from Extract_all_charts import process_html


LOG_PATH = Path(__file__).resolve().with_name("gui_app.log")
LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
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


class TextHandler(logging.Handler):
    """Send logging records to a ``tkinter.Text`` widget."""

    def __init__(self, widget):
        super().__init__()
        self.widget = widget

    def emit(self, record):  # pragma: no cover - UI side effect
        msg = self.format(record)
        self.widget.configure(state="normal")
        self.widget.insert("end", msg + "\n")
        self.widget.configure(state="disabled")
        self.widget.see("end")


def main():
    # License validation removed

    # --- Top-level window setup -----------------------------------------
    root = Tk()  # the root window manages the event loop
    root.title("MEOS Extract GUI")
    root.geometry("600x300")
    # ``minsize`` prevents users from shrinking the window until widgets overlap
    root.minsize(600, 300)
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)

    # PanedWindow splits the interface vertically into controls and log panel
    paned = PanedWindow(root, orient="vertical")
    paned.grid(row=0, column=0, sticky="nsew")

    # Wrap the notebook in a frame so geometry management is explicit
    controls_frame = ttk.Frame(paned)
    controls_frame.grid_rowconfigure(0, weight=1)
    controls_frame.grid_columnconfigure(0, weight=1)
    paned.add(controls_frame, minsize=120)

    # --- Upper pane: user input -----------------------------------------
    notebook = ttk.Notebook(controls_frame)
    notebook.grid(row=0, column=0, sticky="nsew")

    extract_frame = ttk.Frame(notebook)
    notebook.add(extract_frame, text="Estrazione")
    extract_frame.grid_columnconfigure(0, weight=1)
    for r in range(3):
        # a non-zero ``minsize`` keeps rows visible when the window shrinks
        extract_frame.grid_rowconfigure(r, weight=1, minsize=30)
    extract_frame.grid_rowconfigure(3, weight=3)

    style = ttk.Style()
    style.configure("Caption.TLabel", font=("Segoe UI", 10, "bold"))

    # ``StringVar`` keeps the Entry text in sync with ``output_dir``
    output_var = StringVar()
    txt_output = ttk.Entry(
        extract_frame,
        textvariable=output_var,
        state="readonly",
        width=60,
    )
    lbl_output = ttk.Label(extract_frame, text="Output folder:", style="Caption.TLabel")
    lbl_output.grid(row=0, column=0, sticky="w", padx=5, pady=(0, 2))
    txt_output.grid(row=1, column=0, sticky="ew", padx=5, pady=5)

    lbl_input = ttk.Label(extract_frame, text="Input folders:", style="Caption.TLabel")
    lbl_input.grid(row=2, column=0, sticky="w", padx=5, pady=(10, 2))
    listbox = Listbox(extract_frame, width=60, height=8, selectmode="extended")
    listbox.grid(row=3, column=0, sticky="nsew", padx=(5, 0), pady=5)
    list_scroll = ttk.Scrollbar(extract_frame, orient="vertical", command=listbox.yview)
    list_scroll.grid(row=3, column=1, sticky="ns", padx=(0, 5), pady=5)
    listbox.configure(yscrollcommand=list_scroll.set)

    lbl_count = ttk.Label(extract_frame)
    lbl_count.grid(row=4, column=0, sticky="ew", padx=5, pady=5)

    # Button bar uses ``pack`` inside its own frame; mixing layout managers
    # within one container is problematic, but separate frames may use
    # different managers safely.
    btn_frame = ttk.Frame(extract_frame)
    btn_frame.grid(row=5, column=0, sticky="ew", padx=5, pady=5)

    # --- Link budget tab -----------------------------------------------
    budget_frame = ttk.Frame(notebook)
    notebook.add(budget_frame, text="Link budget uplink")
    for r in range(11):
        budget_frame.grid_rowconfigure(r, weight=1, minsize=28)
    budget_frame.grid_columnconfigure(1, weight=1)

    lbl_budget = ttk.Label(budget_frame, text="Calcolo link budget (uplink)", style="Caption.TLabel")
    lbl_budget.grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=(10, 2))

    def add_field(row, text, default):
        label = ttk.Label(budget_frame, text=text)
        label.grid(row=row, column=0, sticky="w", padx=5, pady=2)
        var = StringVar(value=str(default))
        entry = ttk.Entry(budget_frame, textvariable=var)
        entry.grid(row=row, column=1, sticky="ew", padx=5, pady=2)
        return var

    freq_var = add_field(1, "Frequenza (GHz)", 8.0)
    dist_var = add_field(2, "Distanza (km)", 38000)
    tx_power_var = add_field(3, "Potenza TX (dBW)", 20)
    tx_gain_var = add_field(4, "Guadagno antenna TX (dBi)", 35)
    tx_losses_var = add_field(5, "Perdite TX (dB)", 1.0)
    rx_gain_var = add_field(6, "Guadagno antenna RX (dBi)", 45)
    rx_losses_var = add_field(7, "Perdite RX (dB)", 1.5)
    required_var = add_field(8, "Sensibilit\u00e0 richiesta (dBW)", -125)

    result_var = StringVar(
        value=(
            "Inserisci i parametri e premi \"Calcola\" per ottenere: "
            "EIRP, perdita di percorso (FSPL), potenza ricevuta e margine rispetto alla sensibilit\u00e0."
        )
    )
    result_lbl = ttk.Label(
        budget_frame,
        textvariable=result_var,
        wraplength=520,
        justify="left",
    )
    result_lbl.grid(row=9, column=0, columnspan=2, sticky="ew", padx=5, pady=(10, 5))

    def _parse_value(var, name):
        try:
            return float(var.get())
        except ValueError:
            messagebox.showerror("Valore non valido", f"Inserisci un numero per {name}")
            raise

    def calculate_link_budget():
        """Calcola un link budget semplificato per l'uplink."""
        try:
            freq = _parse_value(freq_var, "frequenza")
            distance = _parse_value(dist_var, "distanza")
            tx_power = _parse_value(tx_power_var, "potenza TX")
            tx_gain = _parse_value(tx_gain_var, "guadagno TX")
            tx_losses = _parse_value(tx_losses_var, "perdite TX")
            rx_gain = _parse_value(rx_gain_var, "guadagno RX")
            rx_losses = _parse_value(rx_losses_var, "perdite RX")
            required = _parse_value(required_var, "sensibilit\u00e0")
        except ValueError:
            return

        # Free-space path loss using distance in km and frequency in GHz
        fspl = 92.45 + 20 * math.log10(distance) + 20 * math.log10(freq)
        eirp = tx_power + tx_gain - tx_losses
        received_power = eirp + rx_gain - rx_losses - fspl
        margin = received_power - required

        msg_lines = [
            f"EIRP: {eirp:.2f} dBW",
            f"FSPL: {fspl:.2f} dB",
            f"Potenza ricevuta: {received_power:.2f} dBW",
            f"Margine rispetto alla sensibilit\u00e0: {margin:.2f} dB",
        ]
        result_var.set("\n".join(msg_lines))
        logging.info("Link budget uplink calcolato: %s", "; ".join(msg_lines))

    btn_budget = ttk.Button(budget_frame, text="Calcola", command=calculate_link_budget)
    btn_budget.grid(row=10, column=0, columnspan=2, sticky="e", padx=5, pady=5)

    # --- Lower pane: log output -----------------------------------------
    log_frame = ttk.Frame(paned)
    paned.add(log_frame, minsize=80)
    log_frame.grid_rowconfigure(0, weight=1)
    log_frame.grid_columnconfigure(0, weight=1)
    log_panel = Text(log_frame, state="disabled")
    log_panel.grid(row=0, column=0, sticky="nsew", padx=(5, 0), pady=5)
    log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=log_panel.yview)
    log_scroll.grid(row=0, column=1, sticky="ns", padx=(0, 5), pady=5)
    log_panel.configure(yscrollcommand=log_scroll.set)
    # TextHandler redirects ``logging`` events into the Text widget
    text_handler = TextHandler(log_panel)
    text_handler.setFormatter(logging.getLogger().handlers[0].formatter)
    logging.getLogger().addHandler(text_handler)

    def update_count():
        """Refresh the label showing how many folders are queued."""
        n = listbox.size()
        lbl_count.config(
            text=(
                "Nessuna cartella selezionata"
                if n == 0
                else f"{n} cartella" + ("" if n == 1 else "e") + " in coda"
            ),
        )

    # ``output_dir`` is stored in a dict so inner callbacks can mutate it
    output_dir = {"path": None}

    # --- Callback functions ---------------------------------------------
    def add_folder():
        """Ask the user for a folder and append it to the listbox."""
        folder = filedialog.askdirectory(title="Select folder")
        if folder:
            listbox.insert("end", folder)
            update_count()

    def remove_selected():
        """Remove highlighted folders from the queue."""
        sel = listbox.curselection()
        for idx in reversed(sel):
            listbox.delete(idx)
        update_count()

    def select_output():
        """Prompt for the destination directory and update ``output_var``."""
        folder = filedialog.askdirectory(title="Select output directory")
        if folder:
            output_dir["path"] = Path(folder)
            output_var.set(folder)

    def run():
        """Process each queued folder and log progress to the GUI."""
        if output_dir["path"] is None:
            logging.warning("Output directory not selected")
            messagebox.showwarning(
                "Cartella di output mancante",
                "Seleziona una cartella di destinazione prima di procedere.",
            )
            return

        if listbox.size() == 0:
            logging.warning("No input folders queued")
            messagebox.showwarning(
                "Nessuna cartella da elaborare",
                "Aggiungi almeno una cartella contenente i report HTML.",
            )
            return
        saved = []
        for i in range(listbox.size()):
            folder = Path(listbox.get(i))
            html_files = sorted(folder.glob("*.html"))
            if not html_files:
                logging.warning("No HTML file in %s", folder)
                continue
            if len(html_files) > 1:
                logging.warning(
                    "Multiple HTML files in %s; using %s",
                    folder,
                    html_files[0].name,
                )
            report = html_files[0]
            try:
                out = process_html(report, output_dir["path"])
                logging.info("Saved: %s", out)
                saved.append(out)
            except Exception:
                logging.exception("Error processing %s", folder)
                return
        logging.info("Completed: created %d files", len(saved))
        messagebox.showinfo(
            "Elaborazione completata",
            (
                "Nessun file creato. Controlla i log per eventuali problemi."
                if not saved
                else f"Creati {len(saved)} file nella cartella di output."
            ),
        )

    btn_add = ttk.Button(btn_frame, text="Add folder", command=add_folder)
    btn_remove = ttk.Button(btn_frame, text="Remove selected", command=remove_selected)
    btn_output = ttk.Button(
        btn_frame, text="Output folder destination", command=select_output
    )
    btn_run = ttk.Button(btn_frame, text="Run", command=run)
    btn_exit = ttk.Button(btn_frame, text="Exit", command=root.destroy)

    btn_add.pack(side="left", padx=5, pady=5)
    btn_remove.pack(side="left", padx=5, pady=5)
    btn_output.pack(side="left", padx=5, pady=5)
    btn_run.pack(side="left", padx=5, pady=5)
    btn_exit.pack(side="left", padx=5, pady=5)

    # Escape key uses an event binding to close the window
    root.bind("<Escape>", lambda e: root.destroy())
    update_count()

    root.mainloop()


if __name__ == "__main__":
    main()

