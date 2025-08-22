from pathlib import Path
from tkinter import Tk, Listbox, filedialog, StringVar, Text
from tkinter import ttk
import logging

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
    root = Tk()
    root.title("MEOS Extract GUI")
    root.geometry("600x300")
    root.minsize(600, 300)
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)

    paned = ttk.PanedWindow(root, orient="vertical")
    paned.grid(row=0, column=0, sticky="nsew")

    main_frame = ttk.Frame(paned)
    paned.add(main_frame, weight=3)
    paned.pane(main_frame, minsize=120)
    main_frame.grid_columnconfigure(0, weight=1)
    for r in range(3):
        main_frame.grid_rowconfigure(r, weight=1, minsize=30)
    main_frame.grid_rowconfigure(3, weight=3)

    style = ttk.Style()
    style.configure("Caption.TLabel", font=("Segoe UI", 10, "bold"))

    output_var = StringVar()
    txt_output = ttk.Entry(
        main_frame,
        textvariable=output_var,
        state="readonly",
        width=60,
    )
    lbl_output = ttk.Label(main_frame, text="Output folder:", style="Caption.TLabel")
    lbl_output.grid(row=0, column=0, sticky="w", padx=5, pady=(0, 2))
    txt_output.grid(row=1, column=0, sticky="ew", padx=5, pady=5)

    lbl_input = ttk.Label(main_frame, text="Input folders:", style="Caption.TLabel")
    lbl_input.grid(row=2, column=0, sticky="w", padx=5, pady=(10, 2))
    listbox = Listbox(main_frame, width=60, height=8, selectmode="extended")
    listbox.grid(row=3, column=0, sticky="nsew", padx=(5, 0), pady=5)
    list_scroll = ttk.Scrollbar(main_frame, orient="vertical", command=listbox.yview)
    list_scroll.grid(row=3, column=1, sticky="ns", padx=(0, 5), pady=5)
    listbox.configure(yscrollcommand=list_scroll.set)

    lbl_count = ttk.Label(main_frame)
    lbl_count.grid(row=4, column=0, sticky="ew", padx=5, pady=5)

    btn_frame = ttk.Frame(main_frame)
    btn_frame.grid(row=5, column=0, sticky="ew", padx=5, pady=5)

    log_frame = ttk.Frame(paned)
    paned.add(log_frame, weight=1)
    paned.pane(log_frame, minsize=80)
    log_frame.grid_rowconfigure(0, weight=1)
    log_frame.grid_columnconfigure(0, weight=1)
    log_panel = Text(log_frame, state="disabled")
    log_panel.grid(row=0, column=0, sticky="nsew", padx=(5, 0), pady=5)
    log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=log_panel.yview)
    log_scroll.grid(row=0, column=1, sticky="ns", padx=(0, 5), pady=5)
    log_panel.configure(yscrollcommand=log_scroll.set)
    text_handler = TextHandler(log_panel)
    text_handler.setFormatter(logging.getLogger().handlers[0].formatter)
    logging.getLogger().addHandler(text_handler)

    def update_count():
        n = listbox.size()
        lbl_count.config(
            text=f"{n} folder{'s' if n != 1 else ''} queued"
        )

    output_dir = {"path": None}

    def add_folder():
        folder = filedialog.askdirectory(title="Select folder")
        if folder:
            listbox.insert("end", folder)
            update_count()

    def remove_selected():
        sel = listbox.curselection()
        for idx in reversed(sel):
            listbox.delete(idx)
        update_count()

    def select_output():
        folder = filedialog.askdirectory(title="Select output directory")
        if folder:
            output_dir["path"] = Path(folder)
            output_var.set(folder)

    def run():
        if output_dir["path"] is None:
            logging.warning("Output directory not selected")
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

    root.bind("<Escape>", lambda e: root.destroy())
    update_count()

    root.mainloop()


if __name__ == "__main__":
    main()
