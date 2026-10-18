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

Avvio:
```bash
python gui_app.py
```

1. **Cartelle di input**: *+ Aggiungi cartella* aggiunge una cartella con i report `*.html`;
   *Aggiungi sottocartelle…* scansiona una cartella padre e aggiunge tutte le sottocartelle che contengono report.
   Per ogni cartella viene mostrato il numero di report trovati (in rosso se nessuno).
   *Rimuovi* (o tasto Canc) e *Svuota* gestiscono la lista; doppio clic apre la cartella.
2. **Cartella di output**: *Sfoglia…* per sceglierla, *Apri* per aprirla nel file manager.
3. **Opzioni** (facoltative):
   - *Conta gli unlock del demodulatore*: conta gli eventi di unlock interni al pass (pattern stabile `1→0→1`) e crea `lock_state_stats.xlsx` (colonne `Orbit Number` e `Unlocks`).
   - *Plot polari / 3D* per Input Level, Eb/N0 e Signal/Noise Ratio: grafici polari e 3D sferici (cupola del cielo con antenna al centro) colorati rispetto ad azimuth/elevation; nel plot SNR i campioni con `demodulator_lock_state = 0` sono evidenziati in viola. Se `plotly` è installato vengono create anche versioni HTML interattive con hover (orbita, azimuth, elevation, valore). Viene salvato un indice `polar_plots_index.xlsx`.
   - *Modalità plot*: un set per ogni report e/o un set combinato per tutti i report.
   - *Analisi multi-orbita → Dopo l'elaborazione*: a fine elaborazione crea `meos_processed.xlsx` nella cartella di output usando **tutti** gli Excel MEOS presenti, anche quelli di esecuzioni precedenti (vedi *Analisi multi-orbita* sotto).
   - *Analizza Excel già estratti…*: esegue subito l'analisi multi-orbita su una cartella qualsiasi di Excel già estratti (vecchie campagne, file ricevuti da altri), senza elaborare report HTML; `meos_processed.xlsx` viene salvato in quella cartella.
4. **Avvia elaborazione** (o Ctrl+Invio): l'elaborazione gira in background, con barra di avanzamento e stato; *Stop* interrompe dopo il report corrente. Un report con errori viene segnalato e saltato senza fermare gli altri. A fine elaborazione *Apri risultati* apre la cartella di output.

Il log è mostrato a colori nella finestra (avvisi in giallo, errori in rosso) ed è salvato anche in `gui_app.log` nella cartella dello script.
Cartelle e opzioni vengono ricordate tra un avvio e l'altro (`~/.meos_extract_gui.json`).
La GUI richiede `tkinter` (vedi *Dipendenze*).

Note sui dati estratti:
- L'asse tempo è ricavato dalle etichette orarie del grafico (HH:MM), quindi `t_sec_rel` è in secondi da *Start time* ed è coerente tra tutte le sezioni.
- I grafici di stato (es. FEP Lock State: Idle/Search/Check/Lock/...) hanno come valore l'indice del livello dal basso (0, 1, 2, ...) e una colonna aggiuntiva `<sezione>_state` con il nome dello stato; per `demodulator_lock_state` resta 0 = unlocked, 1 = locked.
- Per i grafici con più serie viene estratta una sola curva (quella con più punti nell'area del grafico).


## Analisi multi-orbita

`report_analysis.py` (prima nel repository separato *Report_Analysis*, di cui conserva la cronologia) confronta l'SNR di passaggi diversi in funzione della posizione dell'antenna, partendo dagli Excel prodotti da MEOS-Extract.

Per ogni orbita:
1. prende le serie SNR, azimuth ed elevation (formato vecchio con fogli `6_1_azimuth`/`6_2_elevation` o formato attuale con il foglio antenna);
2. abbina ogni campione SNR all'azimuth/elevation più vicino nel tempo (entro 2 s);
3. taglia la coda dopo la fine del passaggio (antenna che torna in parcheggio: tutto ciò che segue l'ultimo campione ≤ 5°);
4. divide il passaggio in salita (prima del picco di elevazione) e discesa.

Le orbite che ripercorrono la stessa traccia a terra (ciclo di **9 giorni**, stessa posizione nel ciclo entro ±30 min) formano un gruppo. Per ogni gruppo l'output contiene:
- foglio **Group N**: tempo, azimuth, elevation, SNR ed elevazione con segno di ogni orbita, più la tabella *SNR Correlation* con l'SNR delle orbite nei punti azimuth/elevation comuni (e relativo grafico);
- foglio **Group N Chart**: grafici SNR vs elevazione per la salita e per la discesa, una serie per orbita.

Se l'output esiste già, le sue orbite vengono mantenute e unite a quelle nuove (un'orbita presente in entrambi viene aggiornata). Gli Excel che non sono export MEOS (`lock_state_stats.xlsx`, `polar_plots_index.xlsx`, ...) vengono saltati.

Da riga di comando:
```bash
python report_analysis.py <cartella_excel> -o meos_processed.xlsx
```

| Opzione | Default | Significato |
|---|---|---|
| `-o`, `--output` | `meos_processed.xlsx` | file Excel di output |
| `--repeat-days` | `9` | ciclo di ripetizione della traccia a terra (giorni) |
| `--group-tolerance-min` | `30` | scarto massimo (minuti) nel ciclo per stare nello stesso gruppo |
| `--azimuth-tolerance` | `0.1` | passo (gradi) dei bin di azimuth della tabella di correlazione |
| `--elevation-tolerance` | `0.1` | passo (gradi) dei bin di elevation della tabella di correlazione |
| `--replace` | – | ignora l'output esistente e lo rigenera solo dagli Excel della cartella |

`AWS-PFM_orbit_2648.xlsx` è un Excel di esempio nel formato vecchio.

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
