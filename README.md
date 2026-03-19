# MEOS-Extract

Script per estrarre i grafici dal report HTML generato da MEOS.

Per una spiegazione dettagliata del codice, consultare [CODE_MANUAL.md](CODE_MANUAL.md).

## Dipendenze

- Python 3
- beautifulsoup4
- pandas
- numpy
- openpyxl
- tkinter (se non incluso, installare ad es. `sudo apt install python3-tk`)
- matplotlib (opzionale, per PNG polari/3D)
- plotly (opzionale, per plot HTML interattivi con hover)

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
4. (Opzionale) Abilitare le opzioni nel riquadro **Statistics**:
   - `demodulator_lock_state`: conta gli eventi di unlock interni al pass (pattern stabile `1→0→1`).
   - `Polar plot Input Level`, `Polar plot Eb/No`, `Polar plot SNR`: genera grafici polari a colori rispetto ad azimuth/elevation; nel plot SNR gli eventuali campioni con `demodulator_lock_state = 0` vengono evidenziati in viola.
5. Scegliere la modalità dei plot:
   - `one set per file`: genera i plot separati per ciascun report elaborato.
   - `one combined set for all files`: genera anche un solo plot per parametro selezionato, aggregando i dati di tutti i report elaborati.
   - Le due opzioni sono selezionabili insieme.
6. Premere **Run** per generare gli Excel; ogni file salvato verrà segnalato.
7. Se è abilitata la statistica lock, viene creato `lock_state_stats.xlsx` (solo colonne `Orbit Number` e `Unlocks`).
8. Se sono abilitati i plot, vengono creati PNG polari **e 3D sferici (cupola del cielo con antenna al centro)** e, se `plotly` è disponibile, anche versioni HTML interattive con hover che mostrano orbita, azimuth, elevation e valore della metrica. Viene inoltre salvato un indice `polar_plots_index.xlsx`.
9. La GUI richiede `tkinter`. Se non è già presente, installarlo come indicato nella sezione *Dipendenze*.

I log dell'applicazione sono salvati nel file `gui_app.log` nella stessa directory dello script. Se il file non è scrivibile, i messaggi vengono mostrati solo in console.

## File di licenza

Il programma richiede un file `license.key` nella stessa cartella di `Extract_all_charts.py` o `gui_app.py`.
Nel repository è presente un file di esempio che viene incluso anche nel bundle PyInstaller,
ma **deve essere sostituito con una chiave valida prima dell'esecuzione**.
Il file deve contenere la stringa esadecimale generata con la stessa chiave segreta usata dall'applicazione:

```python
import hmac, hashlib
secret = b"demo-secret"  # sostituire con il proprio segreto
print(hmac.new(secret, b"MEOS-Extract", hashlib.sha256).hexdigest())
```

Salvare l'output in `license.key`. Se il file manca o la chiave non è valida
l'esecuzione termina mostrando "Invalid or missing license".

**Nota:** al momento il controllo della licenza è disattivato tramite la variabile
d'ambiente `MEOS_SKIP_LICENSE` ed è previsto che venga riabilitato in seguito.

## Licenza

Questo progetto è distribuito con licenza [MIT](LICENSE).
