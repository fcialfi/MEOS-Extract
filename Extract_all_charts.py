# extract_all_charts.py
# ------------------------------------------
# Estrae i grafici dal report HTML (sezioni: Input Level, SNR, Eb/N0,
# Carrier Offset, Phase Loop Error, Reed-Solomon Frames, Frame Error Rate,
# Demodulator/FEP Lock State), mappa X a Start→Stop (secondi),
# ricostruisce Y dai tick numerici, rimuove segmenti spurii (legend/diagonale)
# e salva un unico Excel con un foglio per grafico.
#
# Nome file: <prefix>_orbit_<num>.xls/.xlsx (prefix e orbit number letti dall'HTML)
#   - Se manca il prefix: "orbit_<num>..."
#   - Se manca anche il numero: "orbit..."
#
# Uso standalone:
#   - Metti "report.html" nella stessa cartella
#   - Esegui:  python extract_all_charts.py
#
# Dipendenze:
#   pip install beautifulsoup4 pandas numpy openpyxl
#   (opzionale per .xls: pip install xlwt)
# ------------------------------------------

from pathlib import Path
import argparse
import re
from datetime import datetime, timezone, timedelta
import logging

import numpy as np
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl.chart import ScatterChart, Reference, Series


DEFAULT_HTML = Path("report.html")  # used if directory lacks .html

logging.basicConfig(level=logging.INFO)


# -------------------- Utilities --------------------

def parse_iso_utc(s: str):
    """Parse 'YYYY-MM-DD HH:MM:SSZ' → datetime tz-aware (UTC)."""
    if not s:
        return None
    s = s.strip()
    m = re.match(r"(\d{4}-\d{2}-\d{2})\s+(\d{2}):(\d{2}):(\d{2})Z", s)
    if not m:
        return None
    y, mo, d = map(int, m.group(1).split("-"))
    hh, mm, ss = int(m.group(2)), int(m.group(3)), int(m.group(4))
    return datetime(y, mo, d, hh, mm, ss, tzinfo=timezone.utc)


def parse_transforms(transform: str):
    """Accorpa scale()/translate() in (sx, sy, tx, ty)."""
    sx, sy, tx, ty = 1.0, 1.0, 0.0, 0.0
    if not transform:
        return sx, sy, tx, ty
    for m in re.finditer(r"(translate|scale)\(\s*([^)]+)\)", transform):
        kind = m.group(1)
        args = [float(v) for v in re.split(r"[, \t]+", m.group(2).strip()) if v]
        if kind == "scale":
            if len(args) == 1:
                sx *= args[0]; sy *= args[0]
            else:
                sx *= args[0]; sy *= args[1]
        else:  # translate
            if len(args) == 1:
                tx += args[0]
            else:
                tx += args[0]; ty += args[1]
    return sx, sy, tx, ty


def cumulative_transform(tag):
    """Accumula scale/translate lungo la gerarchia SVG del nodo."""
    Sx, Sy, Tx, Ty = 1.0, 1.0, 0.0, 0.0
    chain = []
    cur = tag
    while cur is not None and getattr(cur, "name", None) is not None:
        tr = cur.get("transform")
        if tr:
            chain.append(tr)
        cur = cur.parent
    for tr in reversed(chain):
        sx, sy, tx, ty = parse_transforms(tr)
        Tx = sx * Tx + tx
        Ty = sy * Ty + ty
        Sx *= sx
        Sy *= sy
    return Sx, Sy, Tx, Ty


def apply_tr(x, y, Sx, Sy, Tx, Ty):
    return Sx * x + Tx, Sy * y + Ty


def parse_path_subpaths(d_attr: str):
    """
    Parsea un 'path' SVG supportando M/m L/l H/h V/v e lo spezza in subpath.
    Un nuovo subpath inizia a ogni 'M'/'m'. Ritorna lista di liste di (x,y).
    """
    tokens = re.findall(r"([MmLlHhVv])|([-+]?\d*\.?\d+(?:e[-+]?\d+)?)", d_attr or "")
    flat = []
    for a, b in tokens:
        if a:
            flat.append(a)
        else:
            flat.append(float(b))

    subpaths = []
    x = y = 0.0
    cmd = None
    i = 0
    current = []

    def flush():
        nonlocal current
        if len(current) >= 2:
            subpaths.append(current)
        current = []

    while i < len(flat):
        t = flat[i]
        if isinstance(t, str):
            cmd = t
            if cmd in ("M", "m"):
                flush()
            i += 1
            continue

        if cmd == "M":
            if i + 1 < len(flat) and not isinstance(flat[i + 1], str):
                x, y = flat[i], flat[i + 1]
                current = [(x, y)]
                i += 2
            else:
                i += 1
        elif cmd == "m":
            if i + 1 < len(flat) and not isinstance(flat[i + 1], str):
                x += flat[i]; y += flat[i + 1]
                current = [(x, y)]
                i += 2
            else:
                i += 1
        elif cmd == "L":
            if i + 1 < len(flat) and not isinstance(flat[i + 1], str):
                x, y = flat[i], flat[i + 1]
                current.append((x, y))
                i += 2
            else:
                i += 1
        elif cmd == "l":
            if i + 1 < len(flat) and not isinstance(flat[i + 1], str):
                x += flat[i]; y += flat[i + 1]
                current.append((x, y))
                i += 2
            else:
                i += 1
        elif cmd == "H":
            x = flat[i]; current.append((x, y)); i += 1
        elif cmd == "h":
            x += flat[i]; current.append((x, y)); i += 1
        elif cmd == "V":
            y = flat[i]; current.append((x, y)); i += 1
        elif cmd == "v":
            y += flat[i]; current.append((x, y)); i += 1
        else:
            i += 1

    flush()
    return subpaths


def svg_axes_from_ticks(svg):
    """
    Estrae tick (label testo) e posizioni pixel assolute.
    Definisce il riquadro assi:
      - X: min/max posizione pixel dei tick orari (HH:MM o HH:MM:SS)
      - Y: min/max posizione pixel dei tick numerici
    """
    rows = []
    for t in svg.find_all("text"):
        content = (t.get_text() or "").strip()
        if not content:
            continue
        is_num = re.fullmatch(r"[-+]?\d+(\.\d+)?", content) is not None or content.lstrip("+-").isdigit()
        is_time = (
            re.fullmatch(r"\d{2}:\d{2}", content) is not None or
            re.fullmatch(r"\d{2}:\d{2}:\d{2}", content) is not None
        )
        if not (is_num or is_time):
            continue
        Sx, Sy, Tx, Ty = cumulative_transform(t)
        x_px, y_px = apply_tr(0.0, 0.0, Sx, Sy, Tx, Ty)
        rows.append({"text": content, "x_px": x_px, "y_px": y_px, "kind": "num" if is_num else "time"})

    ticks = pd.DataFrame(rows)

    if ticks.empty:
        return ticks, (None, None, None, None)

    x_tick_px_min = ticks.loc[ticks["kind"] == "time", "x_px"].min()
    x_tick_px_max = ticks.loc[ticks["kind"] == "time", "x_px"].max()
    y_tick_px_min = ticks.loc[ticks["kind"] == "num", "y_px"].min()
    y_tick_px_max = ticks.loc[ticks["kind"] == "num", "y_px"].max()

    return ticks, (x_tick_px_min, x_tick_px_max, y_tick_px_min, y_tick_px_max)


def extract_curve_for_header_id(soup: BeautifulSoup, hdr_id: str):
    """
    Per una sezione (h2/h3) identificata da id (es. '_input_level'), raccoglie gli SVG
    sottostanti fino al prossimo h2/h3, e sceglie il sottopercorso dati migliore
    (massimo numero di punti dentro il riquadro assi).
    """
    hdr = soup.find(id=hdr_id)
    if not hdr:
        return pd.DataFrame(), pd.DataFrame()

    svgs = []
    cur = hdr
    while True:
        cur = cur.find_next_sibling()
        if cur is None or cur.name in ("h2", "h3"):
            break
        svgs.extend(cur.find_all("svg"))
    if not svgs:
        return pd.DataFrame(), pd.DataFrame()

    # Scegli lo svg con più gruppi "gnuplot_plot_*"
    def count_groups(s):
        return len([g for g in s.find_all("g") if (g.get("id") or "").startswith("gnuplot_plot_")])

    best_svg = max(svgs, key=count_groups)
    ticks, axes = svg_axes_from_ticks(best_svg)
    x_min_tick, x_max_tick, y_min_tick, y_max_tick = axes

    best_pts = []
    best_score = -1

    groups = [g for g in best_svg.find_all("g") if (g.get("id") or "").startswith("gnuplot_plot_")]
    for g in groups:
        # PATH: split in subpath e valuta punti dentro assi
        for p in g.find_all("path"):
            d = p.get("d")
            if not d:
                continue
            Sx, Sy, Tx, Ty = cumulative_transform(p)
            for sp in parse_path_subpaths(d):
                pts = [apply_tr(x, y, Sx, Sy, Tx, Ty) for x, y in sp]
                if None in (x_min_tick, x_max_tick, y_min_tick, y_max_tick):
                    score = len(pts)
                else:
                    m = 2.0
                    score = sum(
                        (x_min_tick - m <= x <= x_max_tick + m) and
                        (y_min_tick - m <= y <= y_max_tick + m)
                        for x, y in pts
                    )
                if score > best_score:
                    best_pts = pts
                    best_score = score

        # POLYLINE: fallback
        for pl in g.find_all("polyline"):
            raw = (pl.get("points") or "").strip()
            if not raw:
                continue
            raw = re.sub(r"\s+", " ", raw)
            pairs = re.findall(
                r"([-+]?\d*\.?\d+(?:e[-+]?\d+)?)\s*,\s*([-+]?\d*\.?\d+(?:e[-+]?\d+)?)",
                raw
            )
            pts_local = [(float(x), float(y)) for x, y in pairs]
            Sx, Sy, Tx, Ty = cumulative_transform(pl)
            pts = [apply_tr(x, y, Sx, Sy, Tx, Ty) for x, y in pts_local]

            if None in (x_min_tick, x_max_tick, y_min_tick, y_max_tick):
                score = len(pts)
            else:
                m = 2.0
                score = sum(
                    (x_min_tick - m <= x <= x_max_tick + m) and
                    (y_min_tick - m <= y <= y_max_tick + m)
                    for x, y in pts
                )
            if score > best_score:
                best_pts = pts
                best_score = score

    curve_px = pd.DataFrame(best_pts, columns=["x_px", "y_px"])
    return curve_px, ticks


def map_x_to_time(df: pd.DataFrame, start_dt: datetime, stop_dt: datetime):
    """Mappa x_px in tempo assoluto usando Start/Stop (in secondi)."""
    if df.empty or not start_dt or not stop_dt:
        df["t_sec_rel"] = np.nan
        df["time_HH:MM:SS"] = None
        df["time_iso_utc"] = None
        return df

    dur = (stop_dt - start_dt).total_seconds()
    x_min = df["x_px"].min()
    x_max = df["x_px"].max()
    if x_max == x_min:
        df["t_sec_rel"] = 0.0
        df["time_HH:MM:SS"] = start_dt.strftime("%H:%M:%S")
        df["time_iso_utc"] = start_dt.strftime("%Y-%m-%d %H:%M:%S")
        return df

    df["t_sec_rel"] = (df["x_px"] - x_min) / (x_max - x_min) * dur
    df["time_iso_utc"] = df["t_sec_rel"].apply(
        lambda s: (start_dt + timedelta(seconds=float(s))).strftime("%Y-%m-%d %H:%M:%S")
    )
    df["time_HH:MM:SS"] = df["t_sec_rel"].apply(
        lambda s: (start_dt + timedelta(seconds=float(s))).strftime("%H:%M:%S")
    )
    return df


def map_y_from_ticks(df: pd.DataFrame, ticks: pd.DataFrame, colname: str):
    """Fit lineare: y_px → valore asse (da tick numerici)."""
    if df.empty or ticks is None or ticks.empty:
        df[colname] = np.nan
        return df
    y_ticks = ticks[ticks["kind"] == "num"].copy()
    if y_ticks.empty:
        df[colname] = np.nan
        return df
    y_ticks["value"] = y_ticks["text"].astype(float)
    Y = np.vstack([y_ticks["y_px"].values, np.ones(len(y_ticks))]).T
    a, b = np.linalg.lstsq(Y, y_ticks["value"].values, rcond=None)[0]
    df[colname] = a * df["y_px"] + b
    return df


def safe_sheet_name(title: str):
    """Excel: max 31 char, niente []:*?/\\ ."""
    forbidden = set('[]:*?/\\')
    t = ''.join('_' if ch in forbidden else ch for ch in title)
    t = t.strip()
    return t[:31] if len(t) > 31 else t


def derive_orbit_filename(soup: BeautifulSoup):
    """
    Ricava (prefix, orbit_no) dall'HTML, senza fallback fissi.
    - orbit_no: 'orbit 1234' / 'Orbit: 1234'
    - prefix: prova in tabella "Session" (Spacecraft/Satellite/Mission/Name/Platform/...),
              quindi <title> e header h1/h2/h3, infine pattern tipo 'AWS-PFM' nel testo.
      Se non si trova nulla, prefix=None.
    """
    txt = soup.get_text(" ", strip=True)

    # Orbit number
    m = re.search(r"\borbit\s*[:#]?\s*(\d{1,7})\b", txt, flags=re.I)
    orbit_no = m.group(1) if m else None

    candidates = []

    # 1) Dalla tabella Session
    h2 = soup.find(id="_session")
    if h2:
        tbl = h2.find_next("table")
        if tbl:
            for row in tbl.find_all("tr"):
                cells = row.find_all(["td", "th"])
                if len(cells) < 2:
                    continue
                key = cells[0].get_text(" ", strip=True).lower()
                val = cells[1].get_text(" ", strip=True)
                if not val:
                    continue
                if any(k in key for k in (
                    "spacecraft", "satellite", "mission", "name", "platform", "asset", "receiver", "modem"
                )):
                    candidates.append(val)

    # 2) <title> + headers
    if soup.title and soup.title.get_text(strip=True):
        candidates.append(soup.title.get_text(" ", strip=True))
    for hdr in soup.find_all(["h1", "h2", "h3"]):
        t = hdr.get_text(" ", strip=True)
        if t:
            candidates.append(t)

    # 3) pattern tipo AWS-PFM da testo generale
    m2 = re.search(r"\b([A-Z]{2,}(?:-[A-Z0-9]{2,})+)\b", txt)
    if m2:
        candidates.append(m2.group(1))

    prefix = None
    for raw in candidates:
        cleaned = re.sub(r"[^A-Za-z0-9_-]+", " ", raw).strip()
        m = re.search(r"\b([A-Za-z0-9]+(?:[-_][A-Za-z0-9]+)+)\b", cleaned)
        if not m:
            m = re.search(r"\b([A-Z0-9]{3,})\b", cleaned)
        if m:
            prefix = m.group(1)
            break

    return prefix, orbit_no


# -------------------- Main --------------------

def process_html(html_path: Path, output_dir: Path) -> Path:
    """Elabora un report HTML e salva i grafici in un file Excel.

    Parameters
    ----------
    html_path : Path
        Percorso del file HTML del report.
    output_dir : Path
        Directory in cui salvare l'Excel risultante.

    Returns
    -------
    Path
        Percorso del file Excel creato.
    """
    html = Path(html_path)
    if not html.exists():
        alt = Path(__file__).resolve().parent / html.name
        if alt.exists():
            html = alt
        else:
            raise FileNotFoundError(
                f"File HTML non trovato: {html_path} (cwd) o {alt} (script dir)"
            )

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    with html.open("r", encoding="utf-8") as f:
        soup = BeautifulSoup(f, "html.parser")

    # Tempi di sessione
    start_dt = stop_dt = rep_dt = None
    h2 = soup.find(id="_session")
    if h2:
        tbl = h2.find_next("table")
        if tbl:
            for row in tbl.find_all("tr"):
                cells = row.find_all(["td", "th"])
                if len(cells) < 2:
                    continue
                label = cells[0].get_text(" ", strip=True).lower()
                value = cells[1].get_text(" ", strip=True)
                if "start time" in label:
                    start_dt = parse_iso_utc(value)
                elif "stop time" in label:
                    stop_dt = parse_iso_utc(value)
                elif "report creation time" in label:
                    rep_dt = parse_iso_utc(value)

    # Sezioni target (id degli header h2/h3 del report)
    targets = [
        ("_input_level", "Input Level", "input_level"),
        ("_signalnoise_ratio", "Signal/Noise Ratio", "snr_db"),
        ("_ebn0", "Eb/N0", "ebn0_db"),
        ("_carrier_offset", "Carrier Offset", "carrier_offset"),
        ("_phase_loop_error", "Phase Loop Error", "phase_loop_error"),
        ("_reed_solomon_frames", "Reed-Solomon Frames", "rs_frames"),
        ("_frame_error_rate", "Frame Error Rate", "fer"),
        ("_demodulator_lock_state", "Demodulator Lock State", "lock_state"),
        ("_fep_lock_state", "FEP Lock State", "fep_lock_state"),
    ]

    # Nome file in base a (prefix, orbit_no) trovati nell'HTML
    prefix, orbit_no = derive_orbit_filename(soup)
    base = (
        f"{prefix}_orbit_{orbit_no}" if prefix and orbit_no else
        f"{prefix}_orbit" if prefix and not orbit_no else
        f"orbit_{orbit_no}" if orbit_no else
        "orbit"
    )

    # writer: prova .xls se xlwt presente; altrimenti .xlsx
    try:
        import xlwt  # noqa: F401
        out_path = output_dir / (base + ".xls")
        writer = pd.ExcelWriter(out_path, engine="xlwt")
    except Exception:
        out_path = output_dir / (base + ".xlsx")
        writer = pd.ExcelWriter(out_path)

    with writer as wr:
        # Meta
        pd.DataFrame([
            {
                "start_time_utc": start_dt.strftime("%Y-%m-%d %H:%M:%S") if start_dt else None,
                "stop_time_utc": stop_dt.strftime("%Y-%m-%d %H:%M:%S") if stop_dt else None,
                "report_time_utc": rep_dt.strftime("%Y-%m-%d %H:%M:%S") if rep_dt else None,
            }
        ]).to_excel(wr, sheet_name="__meta__", index=False)

        # Per ogni sezione, estrai e salva in un foglio
        for hdr_id, title, ycol in targets:
            df, ticks = extract_curve_for_header_id(soup, hdr_id)
            df = map_x_to_time(df, start_dt, stop_dt)
            df = map_y_from_ticks(df, ticks, colname=ycol)

            cols = [
                "x_px",
                "y_px",
                "t_sec_rel",
                "time_HH:MM:SS",
                "time_iso_utc",
                ycol,
            ]
            df = df.reindex(cols, axis=1)
            if not df.empty:
                for col in ("time_HH:MM:SS", "time_iso_utc"):
                    if col in df.columns:
                        df[col] = pd.to_datetime(df[col])

            sheet = safe_sheet_name(title)
            if df.empty:
                pd.DataFrame([{"note": "nessun dato estratto"}]).to_excel(
                    wr, sheet_name=sheet, index=False
                )
            else:
                df.to_excel(wr, sheet_name=sheet, index=False)

            # Charts rely on openpyxl and are skipped when generating .xls (xlwt)
            if wr.engine == "openpyxl" and not df.empty:
                wb = wr.book
                ws_data = wr.sheets[sheet]
                csheet = safe_sheet_name(title + " chart")
                # create chart sheet after data sheet (tick sheet will be added later)
                ws_chart = wb.create_sheet(
                    csheet, index=wb.sheetnames.index(ws_data.title) + 1
                )
                wr.sheets[csheet] = ws_chart

                chart = ScatterChart()
                chart.scatterStyle = "lineMarker"   # draw a single line through points
                chart.varyColors = False            # ensure a monochromatic trace
                chart.title = title
                chart.x_axis.title = "Time"
                chart.y_axis.title = ycol
                chart.x_axis.number_format = "hh:mm:ss"
                x_ref = Reference(ws_data, min_col=5, min_row=2, max_row=len(df) + 1)
                y_ref = Reference(ws_data, min_col=6, min_row=2, max_row=len(df) + 1)
                series = Series(values=y_ref, xvalues=x_ref, title=ycol)
                series.smooth = True
                chart.series = []
                chart.series.append(series)
                ws_chart.add_chart(chart, "A1")

            # Ticks in foglio dedicato
            tname = safe_sheet_name(title + " ticks")
            (ticks if not ticks.empty else pd.DataFrame([{"note": "no ticks"}])).to_excel(
                wr, sheet_name=tname, index=False
            )

    return out_path


def main_cli():
    parser = argparse.ArgumentParser(
        description="Estrae i grafici da un report HTML e li salva in un Excel unico."
    )
    parser.add_argument(
        "path",
        nargs="?",
        default=Path("."),
        type=Path,
        help="File HTML o directory contenente l'HTML (default: cartella corrente)",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default=Path("."),
        type=Path,
        help="Directory in cui salvare l'Excel (default: cartella corrente)",
    )
    args = parser.parse_args()

    html_path = args.path
    if html_path.is_dir():
        html_files = sorted(html_path.glob("*.html"))
        if html_files:
            html_path = html_files[0]
        else:
            html_path = html_path / DEFAULT_HTML

    out_path = process_html(html_path, args.output_dir)
    logging.info("Salvato: %s", out_path)


if __name__ == "__main__":
    main_cli()
