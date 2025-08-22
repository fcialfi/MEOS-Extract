from pathlib import Path
from tkinter import Tk, Listbox, Button, Label, filedialog, Frame
from tkinter.scrolledtext import ScrolledText
import logging

from Extract_all_charts import process_html, INPUT_HTML


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

    main_frame = Frame(root)
    main_frame.pack(fill="both", expand=True)

    listbox = Listbox(main_frame, width=60, height=8, selectmode="extended")
    listbox.pack(fill="x")

    lbl_input = Label(main_frame, anchor="w")
    lbl_input.pack(fill="x")

    lbl_output = Label(main_frame, anchor="w")
    lbl_output.pack(fill="x")

    lbl_count = Label(main_frame)
    lbl_count.pack(fill="x")

    bottom_frame = Frame(root)
    bottom_frame.pack(fill="x", side="bottom")

    log_panel = ScrolledText(bottom_frame, height=10, state="disabled")
    log_panel.pack(fill="both", expand=True, side="bottom")
    text_handler = TextHandler(log_panel)
    text_handler.setFormatter(logging.getLogger().handlers[0].formatter)
    logging.getLogger().addHandler(text_handler)

    btn_frame = Frame(bottom_frame)
    btn_frame.pack(fill="x")

    def update_count():
        n = listbox.size()
        lbl_count.config(
            text=f"{n} folder{'s' if n != 1 else ''} queued"
        )

    output_dir = {"path": None}

    def on_select(event):
        sel = listbox.curselection()
        if sel:
            lbl_input.config(text=listbox.get(sel[0]))

    def add_folder():
        folder = filedialog.askdirectory(title="Select folder")
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
        folder = filedialog.askdirectory(title="Select output directory")
        if folder:
            output_dir["path"] = Path(folder)
            lbl_output.config(text=folder)

    def run():
        if output_dir["path"] is None:
            logging.warning("Output directory not selected")
            return
        saved = []
        for i in range(listbox.size()):
            folder = Path(listbox.get(i))
            report = folder / INPUT_HTML.name
            try:
                out = process_html(report, output_dir["path"])
                logging.info("Saved: %s", out)
                saved.append(out)
            except Exception:
                logging.exception("Error processing %s", folder)
                return
        logging.info("Completed: created %d files", len(saved))

    listbox.bind("<<ListboxSelect>>", on_select)

    btn_add = Button(btn_frame, text="Add folder", command=add_folder)
    btn_remove = Button(btn_frame, text="Remove selected", command=remove_selected)
    btn_output = Button(btn_frame, text="Output folder destination", command=select_output)
    btn_run = Button(btn_frame, text="Run", command=run)
    btn_exit = Button(btn_frame, text="Exit", command=root.destroy)

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
