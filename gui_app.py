from pathlib import Path
from tkinter import Tk, Listbox, Button, messagebox, filedialog

from Extract_all_charts import process_html, INPUT_HTML


def main():
    root = Tk()
    root.title("MEOS Extract GUI")

    listbox = Listbox(root, width=60)
    listbox.pack(fill="both", expand=True)

    output_dir = {"path": None}

    def add_folder():
        folder = filedialog.askdirectory(title="Seleziona cartella")
        if folder:
            listbox.insert("end", folder)

    def remove_selected():
        sel = listbox.curselection()
        for idx in reversed(sel):
            listbox.delete(idx)

    def select_output():
        folder = filedialog.askdirectory(title="Seleziona directory di output")
        if folder:
            output_dir["path"] = Path(folder)

    def run():
        if output_dir["path"] is None:
            messagebox.showwarning("Output mancante", "Seleziona una directory di output")
            return
        saved = []
        for i in range(listbox.size()):
            folder = Path(listbox.get(i))
            report = folder / INPUT_HTML.name
            try:
                out = process_html(report, output_dir["path"])
                print(f"Salvato: {out}")
                saved.append(out)
            except Exception as exc:
                messagebox.showerror("Errore", f"{folder}: {exc}")
                return
        messagebox.showinfo("Completato", f"Creati {len(saved)} file")

    btn_add = Button(root, text="Aggiungi cartella", command=add_folder)
    btn_remove = Button(root, text="Rimuovi selezionata", command=remove_selected)
    btn_output = Button(root, text="Seleziona output", command=select_output)
    btn_run = Button(root, text="Run", command=run)

    btn_add.pack(side="left")
    btn_remove.pack(side="left")
    btn_output.pack(side="left")
    btn_run.pack(side="left")

    root.mainloop()


if __name__ == "__main__":
    main()
