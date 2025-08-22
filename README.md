# MEOS-Extract

Script per estrarre i grafici dal report HTML generato da MEOS.

## Dipendenze

- Python 3
- beautifulsoup4
- pandas
- numpy
- openpyxl
- tkinter (se non incluso, installare ad es. `sudo apt install python3-tk`)

Installazione rapida:
```bash
pip install beautifulsoup4 pandas numpy openpyxl
```

## Interfaccia grafica

Un'interfaccia Tkinter è disponibile per elaborare più cartelle.

1. Avviare con:
   ```bash
   python gui_app.py
   ```
2. Utilizzare il pulsante **Aggiungi cartella** per selezionare le cartelle contenenti `report.html`.
3. Selezionare la cartella di destinazione con **Seleziona output**.
4. Premere **Run** per generare gli Excel; ogni file salvato verrà segnalato.
5. La GUI richiede `tkinter`. Se non è già presente, installarlo come indicato nella sezione *Dipendenze*.

I log dell'applicazione sono salvati nel file `gui_app.log` nella stessa directory dello script. Se il file non è scrivibile, i messaggi vengono mostrati solo in console.

