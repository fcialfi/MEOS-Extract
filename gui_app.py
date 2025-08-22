from pathlib import Path
from tkinter import Tk, Listbox, Button, Label, filedialog
from tkinter.scrolledtext import ScrolledText
import logging

from Extract_all_charts import process_html, INPUT_HTML


LOG_PATH = Path(__file__).resolve().with_name("gui_app.log")
LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
handlers = [logging.StreamHandler()]
try:
    handlers.insert(0, logging.FileHandler(LOG_PATH))
except OSError as exc:  # pragma: no cover - log fallback
    print(f"Impossibile scrivere il file di log: {exc}")
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

    listbox = Listbox(root, width=60, height=8, selectmode="extended")
    listbox.pack(fill="x")

    lbl_input = Label(root, anchor="w")
    lbl_input.pack(fill="x")

    lbl_output = Label(root, anchor="w")
    lbl_output.pack(fill="x")

    lbl_count = Label(root)
    lbl_count.pack()

    log_panel = ScrolledText(root, height=10, state="disabled")
    log_panel.pack(fill="both", side="bottom")
    text_handler = TextHandler(log_panel)
    text_handler.setFormatter(logging.getLogger().handlers[0].formatter)
    logging.getLogger().addHandler(text_handler)

    def update_count():
        n = listbox.size()
        lbl_count.config(
            text=f"{n} cartella{'e' if n != 1 else ''} in coda"
        )

    output_dir = {"path": None}

    def on_select(event):
        sel = listbox.curselection()
        if sel:
            lbl_input.config(text=listbox.get(sel[0]))

    def add_folder():
        folder = filedialog.askdirectory(title="Seleziona cartella")
        if folder:
            listbox.insert("end", folder)
            lbl_input.config(text=folder)
            update_count()

    def remove_selected():
        sel = listbox.curselection()
        for idx in reversed(sel):
            listbox.delete(idx)
        update_count()

    def select_output():
        folder = filedialog.askdirectory(title="Seleziona directory di output")
        if folder:
            output_dir["path"] = Path(folder)
            lbl_output.config(text=folder)

    def run():
        if output_dir["path"] is None:
            logging.warning("Directory di output non selezionata")
            return
        saved = []
        for i in range(listbox.size()):
            folder = Path(listbox.get(i))
            report = folder / INPUT_HTML.name
            try:
                out = process_html(report, output_dir["path"])
                logging.info("Salvato: %s", out)
                saved.append(out)
            except Exception:
                logging.exception("Errore elaborando %s", folder)
                return
        logging.info("Completato: creati %d file", len(saved))

    listbox.bind("<<ListboxSelect>>", on_select)

    btn_add = Button(root, text="Aggiungi cartella", command=add_folder)
    btn_remove = Button(root, text="Rimuovi selezionata", command=remove_selected)
    btn_output = Button(root, text="Seleziona output", command=select_output)
    btn_run = Button(root, text="Run", command=run)
    btn_exit = Button(root, text="Exit", command=root.destroy)

    btn_add.pack(side="left")
    btn_remove.pack(side="left")
    btn_output.pack(side="left")
    btn_run.pack(side="left")
    btn_exit.pack(side="left")

    root.bind("<Escape>", lambda e: root.destroy())
    update_count()

    root.mainloop()


if __name__ == "__main__":
    main()
