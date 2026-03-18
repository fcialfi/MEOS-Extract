# extract_all_charts.py
# ------------------------------------------
# Estrae i grafici dal report HTML (sezioni: Input Level, SNR, Eb/N0,
# Carrier Offset, Phase Loop Error, Reed-Solomon Frames, Frame Error Rate,
# Demodulator/FEP Lock State), mappa X a Start→Stop (secondi),
# ricostruisce Y dai tick numerici, rimuove segmenti spurii (legend/diagonale)
# e salva un unico Excel con un foglio per grafico.
#
# Nome file: <prefix>_orbit_<num>.xlsx (prefix e orbit number letti dall'HTML)
#   - Se manca il prefix: "orbit_<num>..."
#   - Se manca anche il numero: "orbit..."
#
# Uso standalone:
#   - Metti "report.html" nella stessa cartella
#   - Esegui:  python extract_all_charts.py
#
# Dipendenze:
#   pip install beautifulsoup4 pandas numpy openpyxl
# ------------------------------------------

from pathlib import Path
import argparse
import re
from datetime import datetime, timezone, timedelta
import logging
import base64
import urllib.request
import sys

import numpy as np
import pandas as pd
from bs4 import BeautifulSoup, FeatureNotFound


DEFAULT_HTML = Path("report.html")  # used if directory lacks .html

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


# -------------------- Utilities --------------------

def parse_iso_utc(s: str):
    """Return a timezone aware :class:`datetime` from a ``YYYY-MM-DD HH:MM:SSZ`` string.

    Parameters
    ----------
    s : str
        Timestamp in the strict ISO format used in the HTML reports.

    Returns
    -------
    datetime | None
        ``datetime`` object in UTC or ``None`` if the input is missing or
        does not match the expected pattern.

    Notes
    -----
    The function validates the string with a regular expression before
    constructing the ``datetime`` object. Only the ``Z`` (Zulu/UTC) timezone
    designator is accepted.
    """
    if not s:
        return None  # Empty field – nothing to parse
    s = s.strip()  # Remove leading/trailing whitespace

    # Match date and time components with a regular expression. ``m`` is a
    # ``re.Match`` object if the pattern is found; otherwise ``None``.
    m = re.match(r"(\d{4}-\d{2}-\d{2})\s+(\d{2}):(\d{2}):(\d{2})Z", s)
    if not m:
        return None  # The string is not in the expected ISO format

    # Split the first capture group (YYYY-MM-DD) into integers using ``map``
    # and unpack the remaining groups for hours, minutes and seconds.
    y, mo, d = map(int, m.group(1).split("-"))
    hh, mm, ss = int(m.group(2)), int(m.group(3)), int(m.group(4))

    # Create and return an aware ``datetime`` with the UTC timezone.
    return datetime(y, mo, d, hh, mm, ss, tzinfo=timezone.utc)


def parse_transforms(transform: str):
    """Collapse an SVG ``transform`` chain into scale and translation factors.

    Parameters
    ----------
    transform : str
        Value of the ``transform`` attribute, e.g. ``"scale(2) translate(3,4)"``.

    Returns
    -------
    tuple[float, float, float, float]
        Aggregate scale ``(sx, sy)`` and translation ``(tx, ty)`` values.

    Explanation
    -----------
    SVG allows multiple ``scale`` and ``translate`` operations to be combined
    in a single string.  This function walks through each transformation in
    order and accumulates the resulting scale and translation.  Only these two
    operations are handled because the charts produced by the MEOS report use
    a simple transformation chain.
    """
    sx, sy, tx, ty = 1.0, 1.0, 0.0, 0.0  # Start with the identity transform
    if not transform:
        return sx, sy, tx, ty  # Nothing to parse → return defaults

    # ``re.finditer`` yields each ``scale`` or ``translate`` call. Captured
    # arguments are later split into a Python list of numbers.
    for m in re.finditer(r"(translate|scale)\(\s*([^)]+)\)", transform):
        kind = m.group(1)  # Either 'scale' or 'translate'
        # ``re.split`` handles comma- or space-separated numbers. ``list``
        # comprehension converts each token to ``float``.
        args = [float(v) for v in re.split(r"[, \t]+", m.group(2).strip()) if v]
        if kind == "scale":
            # Apply scaling. If only one value is supplied, it scales both axes.
            if len(args) == 1:
                sx *= args[0]; sy *= args[0]
            else:
                sx *= args[0]; sy *= args[1]
        else:  # ``translate`` case
            # Translation accepts one or two numbers (x and optionally y).
            if len(args) == 1:
                tx += args[0]
            else:
                tx += args[0]; ty += args[1]
    return sx, sy, tx, ty


def cumulative_transform(tag):
    """Compute the combined transform of an SVG element and its ancestors.

    Parameters
    ----------
    tag : bs4.element.Tag
        Current SVG node whose effective transform is required.

    Returns
    -------
    tuple[float, float, float, float]
        Overall scale ``(Sx, Sy)`` and translation ``(Tx, Ty)`` factors.

    Details
    -------
    SVG applies parent transforms to child elements.  The function therefore
    walks up the DOM tree collecting ``transform`` attributes, then applies
    them from outermost to innermost.  A simple loop is used instead of
    recursion to avoid building deep call stacks.
    """
    # Accumulators start as an identity transform: scale 1 and translation 0.
    Sx, Sy, Tx, Ty = 1.0, 1.0, 0.0, 0.0
    chain = []  # Will store transform strings encountered on the path to root
    cur = tag

    # Traverse ancestors and collect their ``transform`` attributes. The
    # ``getattr`` guard ensures compatibility with objects that may not be
    # BeautifulSoup ``Tag`` instances.
    while cur is not None and getattr(cur, "name", None) is not None:
        tr = cur.get("transform")
        if tr:
            chain.append(tr)
        cur = cur.parent  # Move one level up the tree

    # Apply transforms from root to current element. ``reversed`` produces a
    # generator that yields items in reverse order without creating a copy.
    for tr in reversed(chain):
        sx, sy, tx, ty = parse_transforms(tr)
        # First transform existing translation, then accumulate new offsets.
        Tx = sx * Tx + tx
        Ty = sy * Ty + ty
        # Update the global scale factors.
        Sx *= sx
        Sy *= sy
    return Sx, Sy, Tx, Ty


def apply_tr(x, y, Sx, Sy, Tx, Ty):
    """Apply scale and translation factors to a point.

    Parameters
    ----------
    x, y : float
        Coordinates in the local SVG space.
    Sx, Sy, Tx, Ty : float
        Scale and translation returned by :func:`cumulative_transform`.

    Returns
    -------
    tuple[float, float]
        The transformed point in absolute pixel coordinates.
    """
    # Simple arithmetic uses Python's float type. Returning a tuple allows
    # callers to unpack the two coordinates at once.
    return Sx * x + Tx, Sy * y + Ty


def parse_path_subpaths(d_attr: str):
    """Split an SVG ``path`` definition into absolute subpaths.

    Parameters
    ----------
    d_attr : str
        The content of the ``d`` attribute from a ``<path>`` element.

    Returns
    -------
    list[list[tuple[float, float]]]
        Each inner list contains ``(x, y)`` pairs describing one continuous
        polyline. A new subpath starts whenever an ``M`` or ``m`` command is
        encountered.

    Description
    -----------
    Only the commands ``M/m`` (move), ``L/l`` (line), ``H/h`` (horizontal) and
    ``V/v`` (vertical) are supported because the gnuplot output embedded in the
    report uses these primitives exclusively.  Numbers are converted to
    ``float`` for easier numeric processing.
    """
    # Extract drawing commands and numbers. The regular expression yields a
    # list of tuples where only one item of each tuple is non-empty.
    tokens = re.findall(r"([MmLlHhVv])|([-+]?\d*\.?\d+(?:e[-+]?\d+)?)", d_attr or "")

    flat = []
    for a, b in tokens:
        if a:
            flat.append(a)  # SVG command as a string
        else:
            flat.append(float(b))  # Coordinate or length as Python float

    # ``subpaths`` collects individual polylines; ``current`` tracks the one
    # currently being built. ``x`` and ``y`` hold the cursor position.
    subpaths = []
    x = y = 0.0
    cmd = None
    i = 0
    current = []

    def flush():
        """Append ``current`` polyline to ``subpaths`` if it has at least two points."""
        nonlocal current
        if len(current) >= 2:
            subpaths.append(current)
        current = []  # Reset for the next subpath

    # Iterate through the token list using an index ``i`` so we can look ahead
    # when commands require multiple numeric parameters.
    while i < len(flat):
        t = flat[i]
        if isinstance(t, str):
            # ``t`` is a command letter. Store it and, for ``M/m``, close the
            # previous subpath.
            cmd = t
            if cmd in ("M", "m"):
                flush()
            i += 1
            continue

        # At this point ``t`` is a numeric value associated with the last
        # command seen.
        if cmd == "M":  # Absolute move-to command
            if i + 1 < len(flat) and not isinstance(flat[i + 1], str):
                x, y = flat[i], flat[i + 1]  # Set new absolute position
                current = [(x, y)]  # Start a new subpath list
                i += 2
            else:
                i += 1  # Malformed data; skip value
        elif cmd == "m":  # Relative move-to command
            if i + 1 < len(flat) and not isinstance(flat[i + 1], str):
                x += flat[i]; y += flat[i + 1]  # Update position relative to current point
                current = [(x, y)]
                i += 2
            else:
                i += 1
        elif cmd == "L":  # Absolute line-to command
            if i + 1 < len(flat) and not isinstance(flat[i + 1], str):
                x, y = flat[i], flat[i + 1]
                current.append((x, y))  # Append point to current polyline
                i += 2
            else:
                i += 1
        elif cmd == "l":  # Relative line-to command
            if i + 1 < len(flat) and not isinstance(flat[i + 1], str):
                x += flat[i]; y += flat[i + 1]
                current.append((x, y))
                i += 2
            else:
                i += 1
        elif cmd == "H":  # Absolute horizontal line
            x = flat[i]; current.append((x, y)); i += 1
        elif cmd == "h":  # Relative horizontal line
            x += flat[i]; current.append((x, y)); i += 1
        elif cmd == "V":  # Absolute vertical line
            y = flat[i]; current.append((x, y)); i += 1
        elif cmd == "v":  # Relative vertical line
            y += flat[i]; current.append((x, y)); i += 1
        else:
            i += 1  # Unsupported command; skip token

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
        is_num = re.fullmatch(r"[-+]?\d+(?:[.,]\d+)?(?:\s*(dB|°|deg))?", content) is not None
        is_time = (
            re.fullmatch(r"\d{2}:\d{2}", content) is not None or
            re.fullmatch(r"\d{2}:\d{2}:\d{2}", content) is not None
        )
        is_lock_state = re.fullmatch(
            r"(?i)(lock(?:ed)?|unlock(?:ed)?|no\s*lock|out\s*of\s*lock|loss\s*of\s*lock)",
            content,
        ) is not None
        if not (is_num or is_time or is_lock_state):
            continue
        Sx, Sy, Tx, Ty = cumulative_transform(t)
        x_px, y_px = apply_tr(0.0, 0.0, Sx, Sy, Tx, Ty)
        if is_num:
            kind = "num"
        elif is_time:
            kind = "time"
        else:
            kind = "state"
        rows.append({"text": content, "x_px": x_px, "y_px": y_px, "kind": kind})

    ticks = pd.DataFrame(rows)

    if ticks.empty:
        return ticks, (None, None, None, None)

    x_ticks = ticks.loc[ticks["kind"] == "time", "x_px"]
    y_ticks = ticks.loc[ticks["kind"] == "num", "y_px"]

    x_tick_px_min = x_ticks.min() if not x_ticks.empty else None
    x_tick_px_max = x_ticks.max() if not x_ticks.empty else None
    y_tick_px_min = y_ticks.min() if not y_ticks.empty else None
    y_tick_px_max = y_ticks.max() if not y_ticks.empty else None

    return ticks, (x_tick_px_min, x_tick_px_max, y_tick_px_min, y_tick_px_max)


def extract_curve_for_header(hdr):
    """
    Per una sezione (h2/h3) già individuata, raccoglie gli SVG sottostanti fino
    al prossimo h2/h3 e sceglie il sottopercorso dati migliore (massimo numero
    di punti dentro il riquadro assi).
    """
    if hdr is None:
        return pd.DataFrame(), pd.DataFrame()

    svgs = []
    for el in hdr.next_elements:
        name = getattr(el, "name", None)
        if name in ("h2", "h3"):
            break
        if name == "svg":
            svgs.append(el)
        elif name == "object" and el.get("type") == "image/svg+xml":
            data = el.get("data", "")
            m = re.match(r"^data:image/svg\+xml(;charset=[^;]+)?;base64,(.*)$", data, re.I)
            try:
                if m:  # Base64 inline data
                    svg_bytes = base64.b64decode(m.group(2))
                elif data.startswith(("http://", "https://")):  # Remote file
                    with urllib.request.urlopen(data) as resp:
                        svg_bytes = resp.read()
                elif data:  # Local file path
                    with open(data, "rb") as f:
                        svg_bytes = f.read()
                else:
                    continue
                try:
                    svg_soup = BeautifulSoup(svg_bytes, "xml")
                except FeatureNotFound:
                    logger.warning(
                        "lxml parser not found; falling back to html.parser. Install lxml for full XML support."
                    )
                    try:
                        svg_soup = BeautifulSoup(svg_bytes, "html.parser")
                    except FeatureNotFound:
                        import xml.etree.ElementTree as ET

                        svg_soup = BeautifulSoup(
                            ET.tostring(ET.fromstring(svg_bytes)), "html.parser"
                        )
                if svg_soup.svg:
                    svgs.append(svg_soup.svg)
            except Exception as exc:
                logger.warning("Failed to load SVG from %s: %s", data, exc)
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

    groups = [
        g
        for g in best_svg.find_all("g")
        if (g.get("id") or "").startswith("gnuplot_plot_") and (g.find("path") or g.find("polyline"))
    ]
    if not groups:
        groups = [g for g in best_svg.find_all("g") if g.find("path") or g.find("polyline")]
    def has_missing_axes(values):
        return any(v is None or (isinstance(v, float) and np.isnan(v)) for v in values)

    for g in groups:
        # PATH: split in subpath e valuta punti dentro assi
        for p in g.find_all("path"):
            d = p.get("d")
            if not d:
                continue
            Sx, Sy, Tx, Ty = cumulative_transform(p)
            for sp in parse_path_subpaths(d):
                pts = [apply_tr(x, y, Sx, Sy, Tx, Ty) for x, y in sp]
                if has_missing_axes((x_min_tick, x_max_tick, y_min_tick, y_max_tick)):
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

            if has_missing_axes((x_min_tick, x_max_tick, y_min_tick, y_max_tick)):
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


def extract_curves_for_header(hdr):
    """Extract multiple candidate curves for a header (multi-series friendly)."""
    if hdr is None:
        return []

    svgs = []
    for el in hdr.next_elements:
        name = getattr(el, "name", None)
        if name in ("h2", "h3"):
            break
        if name == "svg":
            svgs.append(el)
        elif name == "object" and el.get("type") == "image/svg+xml":
            data = el.get("data", "")
            m = re.match(r"^data:image/svg\+xml(;charset=[^;]+)?;base64,(.*)$", data, re.I)
            try:
                if m:
                    svg_bytes = base64.b64decode(m.group(2))
                elif data.startswith(("http://", "https://")):
                    with urllib.request.urlopen(data) as resp:
                        svg_bytes = resp.read()
                elif data:
                    with open(data, "rb") as f:
                        svg_bytes = f.read()
                else:
                    continue
                try:
                    svg_soup = BeautifulSoup(svg_bytes, "xml")
                except FeatureNotFound:
                    svg_soup = BeautifulSoup(svg_bytes, "html.parser")
                if svg_soup.svg:
                    svgs.append(svg_soup.svg)
            except Exception:
                continue

    if not svgs:
        return []

    candidates = []

    def has_missing_axes(values):
        return any(v is None or (isinstance(v, float) and np.isnan(v)) for v in values)

    for sidx, svg in enumerate(svgs, start=1):
        ticks, axes = svg_axes_from_ticks(svg)
        x_min_tick, x_max_tick, y_min_tick, y_max_tick = axes

        def score_pts(pts):
            if has_missing_axes((x_min_tick, x_max_tick, y_min_tick, y_max_tick)):
                return len(pts)
            m = 2.0
            return sum(
                (x_min_tick - m <= x <= x_max_tick + m) and
                (y_min_tick - m <= y <= y_max_tick + m)
                for x, y in pts
            )

        groups = [
            g
            for g in svg.find_all("g")
            if (g.get("id") or "").startswith("gnuplot_plot_") and (g.find("path") or g.find("polyline"))
        ]
        if not groups:
            groups = [g for g in svg.find_all("g") if g.find("path") or g.find("polyline")]

        for idx, g in enumerate(groups, start=1):
            title_tag = g.find("title")
            base_title = title_tag.get_text(" ", strip=True) if title_tag else f"svg{sidx}_series_{idx}"

            for p in g.find_all("path"):
                d = p.get("d")
                if not d:
                    continue
                Sx, Sy, Tx, Ty = cumulative_transform(p)
                for sp_i, sp in enumerate(parse_path_subpaths(d), start=1):
                    pts = [apply_tr(x, y, Sx, Sy, Tx, Ty) for x, y in sp]
                    if len(pts) < 3:
                        continue
                    candidates.append((score_pts(pts), f"{base_title}_p{sp_i}", pd.DataFrame(pts, columns=["x_px", "y_px"]), ticks))

            for pl_i, pl in enumerate(g.find_all("polyline"), start=1):
                raw = (pl.get("points") or "").strip()
                if not raw:
                    continue
                raw = re.sub(r"\s+", " ", raw)
                pairs = re.findall(
                    r"([-+]?\d*\.?\d+(?:e[-+]?\d+)?)\s*,\s*([-+]?\d*\.?\d+(?:e[-+]?\d+)?)",
                    raw,
                )
                pts_local = [(float(x), float(y)) for x, y in pairs]
                if len(pts_local) < 3:
                    continue
                Sx, Sy, Tx, Ty = cumulative_transform(pl)
                pts = [apply_tr(x, y, Sx, Sy, Tx, Ty) for x, y in pts_local]
                candidates.append((score_pts(pts), f"{base_title}_pl{pl_i}", pd.DataFrame(pts, columns=["x_px", "y_px"]), ticks))

    if not candidates:
        return []

    candidates.sort(key=lambda x: x[0], reverse=True)
    picked = []
    seen = set()
    for score, title, df, ticks in candidates:
        xs = df["x_px"].to_numpy()
        ys = df["y_px"].to_numpy()
        key = (round(float(xs.min()), 1), round(float(xs.max()), 1), round(float(ys.min()), 1), round(float(ys.max()), 1), len(df))
        if key in seen:
            continue
        seen.add(key)
        picked.append((title, df, ticks))
        if len(picked) >= 10:
            break

    return picked


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
    if not y_ticks.empty:
        y_ticks["value"] = y_ticks["text"].apply(
            lambda s: float(re.sub(r"[^0-9+\-.,]", "", s).replace(",", "."))
        )
        Y = np.vstack([y_ticks["y_px"].values, np.ones(len(y_ticks))]).T
        a, b = np.linalg.lstsq(Y, y_ticks["value"].values, rcond=None)[0]
        df[colname] = a * df["y_px"] + b
        return df

    state_ticks = ticks[ticks["kind"] == "state"].copy()
    if not state_ticks.empty:
        state_ticks["value"] = state_ticks["text"].apply(
            lambda s: 0.0 if re.search(r"(?i)unlock|no\s*lock|out\s*of\s*lock|loss", s) else 1.0
        )
        if len(state_ticks) >= 2:
            Y = np.vstack([state_ticks["y_px"].values, np.ones(len(state_ticks))]).T
            a, b = np.linalg.lstsq(Y, state_ticks["value"].values, rcond=None)[0]
            df[colname] = np.round(a * df["y_px"] + b).clip(0, 1)
        else:
            df[colname] = state_ticks["value"].iloc[0]
        return df

    if "lock" in colname.lower() and not df.empty:
        threshold = float(df["y_px"].median())
        df[colname] = (df["y_px"] <= threshold).astype(int)
        return df

    df[colname] = np.nan
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



def count_unlock_events(values):
    """Count unlock events as stable 1→0→1 patterns."""
    stable = []
    for raw in values:
        if pd.isna(raw):
            continue
        v = int(round(float(raw)))
        if not stable or stable[-1] != v:
            stable.append(v)
    if len(stable) < 3:
        return 0
    return sum(
        1
        for i in range(1, len(stable) - 1)
        if stable[i - 1] == 1 and stable[i] == 0 and stable[i + 1] == 1
    )


def summarize_selected_stats(orbit_no: str, section_frames: dict, selectors):
    """Create summary rows for selected statistics."""
    rows = []
    wanted = {s.lower() for s in (selectors or [])}
    if "demodulator_lock_state" in wanted:
        unlocks_total = 0
        for ycol, df in section_frames.items():
            if "demodulator_lock_state" in ycol.lower() and ycol in df:
                unlocks_total += count_unlock_events(df[ycol])
        rows.append({"Orbit Number": orbit_no or "N/A", "Unlocks": unlocks_total})
    return rows


def _normalized_label(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (s or "").lower())


def _find_section_by_predicate(section_frames: dict, predicate):
    for ycol, df in section_frames.items():
        if predicate(ycol):
            return ycol, df
    return None, None


def _infer_az_el_from_antenna(section_frames: dict):
    """Fallback: infer az/el from multiple antenna-like series when labels are generic."""
    antenna = []
    for ycol, df in section_frames.items():
        n = _normalized_label(ycol)
        if "antenna" not in n:
            continue
        if ycol not in df or df.empty:
            continue
        vals = pd.to_numeric(df[ycol], errors="coerce").dropna()
        if vals.empty:
            continue
        antenna.append((ycol, df, float(vals.median()), float(vals.max()), float(vals.min())))

    if len(antenna) < 2:
        return None, None, None, None

    # prefer a high-span/high-magnitude curve as azimuth and lower-magnitude as elevation
    az_item = max(antenna, key=lambda t: (t[3] - t[4], t[3]))
    el_candidates = [a for a in antenna if a[0] != az_item[0]]
    el_item = min(el_candidates, key=lambda t: t[3])
    return az_item[0], az_item[1], el_item[0], el_item[1]


def _is_azimuth_label(label: str) -> bool:
    n = _normalized_label(label)
    return "azimuth" in n or n.startswith("az") or "antennaaz" in n


def _is_elevation_label(label: str) -> bool:
    n = _normalized_label(label)
    return "elevation" in n or n.startswith("el") or "antennael" in n


def _is_input_level_label(label: str) -> bool:
    n = _normalized_label(label)
    return "inputlevel" in n or "iflevel" in n or ("input" in n and "level" in n)


def _is_ebno_label(label: str) -> bool:
    n = _normalized_label(label)
    return "ebn0" in n or "ebno" in n or "esn0" in n or "esno" in n or ("eb" in n and ("n0" in n or "no" in n))


def _is_snr_label(label: str) -> bool:
    n = _normalized_label(label)
    return "snr" in n or "signaltonoiseratio" in n or "signalnoiseratio" in n or "cn0" in n or "cno" in n


def _find_metric_section(section_frames: dict, selector: str):
    metric_map = {
        "input_level": _is_input_level_label,
        "eb_no": _is_ebno_label,
        "snr": _is_snr_label,
    }
    matcher = metric_map.get(selector)
    if matcher is None:
        return None, None

    col, df = _find_section_by_predicate(section_frames, matcher)
    if df is not None and col is not None:
        return col, df

    token_map = {
        "input_level": ["input", "level", "iflevel"],
        "eb_no": ["eb", "n0", "no", "esn0", "esno"],
        "snr": ["snr", "cn0", "cno", "signal", "noise", "ratio"],
    }
    tokens = token_map.get(selector, [])
    scored = []
    for ycol, cdf in section_frames.items():
        n = _normalized_label(ycol)
        if "antenna" in n or "azimuth" in n or "elevation" in n or "lock" in n:
            continue
        score = sum(1 for t in tokens if t in n)
        if score <= 0 or ycol not in cdf:
            continue
        vals = pd.to_numeric(cdf[ycol], errors="coerce").dropna()
        if vals.empty:
            continue
        scored.append((score, len(vals), ycol, cdf))

    if not scored:
        return None, None
    scored.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return scored[0][2], scored[0][3]


def _align_metric_with_az_el(metric: pd.DataFrame, az: pd.DataFrame, el: pd.DataFrame):
    """Align metric samples with azimuth/elevation on common time interval."""
    # common overlap window
    t0 = max(float(metric["t_sec_rel"].min()), float(az["t_sec_rel"].min()), float(el["t_sec_rel"].min()))
    t1 = min(float(metric["t_sec_rel"].max()), float(az["t_sec_rel"].max()), float(el["t_sec_rel"].max()))
    if t1 <= t0:
        return pd.DataFrame()

    metric = metric[(metric["t_sec_rel"] >= t0) & (metric["t_sec_rel"] <= t1)].copy()
    az = az[(az["t_sec_rel"] >= t0) & (az["t_sec_rel"] <= t1)].copy()
    el = el[(el["t_sec_rel"] >= t0) & (el["t_sec_rel"] <= t1)].copy()
    if metric.empty or az.empty or el.empty:
        return pd.DataFrame()

    # Deduplicate and sort by time.
    metric = metric.sort_values("t_sec_rel").groupby("t_sec_rel", as_index=False)["metric"].median()
    az = az.sort_values("t_sec_rel").drop_duplicates("t_sec_rel")
    el = el.sort_values("t_sec_rel").drop_duplicates("t_sec_rel")
    if len(metric) < 2 or len(az) < 2 or len(el) < 2:
        return pd.DataFrame()

    # Build a regular time base to avoid clustered/jagged point clouds.
    n_samples = int(np.clip(max(len(metric), 200), 120, 800))
    t = np.linspace(t0, t1, n_samples)

    metric_t = metric["t_sec_rel"].to_numpy(dtype=float)
    metric_v = metric["metric"].to_numpy(dtype=float)
    metric_interp = np.interp(t, metric_t, metric_v)

    az_t = az["t_sec_rel"].to_numpy(dtype=float)
    el_t = el["t_sec_rel"].to_numpy(dtype=float)

    # interpolate azimuth with angular continuity: unwrap in radians, then wrap back.
    az_rad = np.deg2rad(az["azimuth"].to_numpy(dtype=float))
    az_unwrapped = np.unwrap(az_rad)
    az_interp_unwrapped = np.interp(t, az_t, az_unwrapped)
    az_interp = np.rad2deg(az_interp_unwrapped) % 360.0

    # linear interpolation for elevation.
    el_interp = np.interp(t, el_t, el["elevation"].to_numpy(dtype=float))

    aligned = pd.DataFrame({
        "t_sec_rel": t,
        "metric": metric_interp,
        "azimuth": az_interp,
        "elevation": el_interp,
    })
    aligned = aligned.dropna(subset=["metric", "azimuth", "elevation"])
    return aligned


def _spherical_to_cartesian(az_deg, el_deg):
    """Convert azimuth/elevation angles to unit-sphere Cartesian coordinates."""
    az = np.deg2rad(np.asarray(az_deg, dtype=float))
    el = np.deg2rad(np.asarray(el_deg, dtype=float))
    r_xy = np.cos(el)
    x = r_xy * np.cos(az)
    y = r_xy * np.sin(az)
    z = np.sin(el)
    return x, y, z


def _build_base_track(az: pd.DataFrame, el: pd.DataFrame, n_points: int = 500):
    """Build a smooth antenna base track from azimuth/elevation time series."""
    t0 = max(float(az["t_sec_rel"].min()), float(el["t_sec_rel"].min()))
    t1 = min(float(az["t_sec_rel"].max()), float(el["t_sec_rel"].max()))
    if t1 <= t0:
        return np.array([]), np.array([])

    az_s = az.sort_values("t_sec_rel").drop_duplicates("t_sec_rel")
    el_s = el.sort_values("t_sec_rel").drop_duplicates("t_sec_rel")
    if len(az_s) < 2 or len(el_s) < 2:
        return np.array([]), np.array([])

    t = np.linspace(t0, t1, n_points)
    az_t = az_s["t_sec_rel"].to_numpy(dtype=float)
    el_t = el_s["t_sec_rel"].to_numpy(dtype=float)

    az_unwrapped = np.unwrap(np.deg2rad(az_s["azimuth"].to_numpy(dtype=float)))
    az_interp = np.rad2deg(np.interp(t, az_t, az_unwrapped)) % 360.0
    el_interp = np.interp(t, el_t, el_s["elevation"].to_numpy(dtype=float))
    el_interp = np.clip(el_interp, 0.0, 90.0)
    return az_interp, el_interp


def generate_polar_plot_artifacts(out_path: Path, section_frames: dict, selectors):
    """Generate polar color plots (metric over azimuth/elevation)."""
    wanted = {s.lower() for s in (selectors or [])}
    if not wanted:
        return []

    az_col, az_df = _find_section_by_predicate(section_frames, _is_azimuth_label)
    el_col, el_df = _find_section_by_predicate(section_frames, _is_elevation_label)
    if az_df is None or el_df is None:
        az_col, az_df, el_col, el_df = _infer_az_el_from_antenna(section_frames)
    if az_df is None or el_df is None:
        logger.warning(
            "Polar plots skipped: azimuth/elevation charts not found. Available sections: %s",
            ", ".join(section_frames.keys()),
        )
        return []

    try:
        import matplotlib.pyplot as plt
    except ImportError:
        logger.warning("matplotlib not available: skipping polar plot generation")
        return []

    artifacts = []

    az = az_df[["t_sec_rel", az_col]].copy()
    el = el_df[["t_sec_rel", el_col]].copy()
    az["t_sec_rel"] = pd.to_numeric(az["t_sec_rel"], errors="coerce")
    el["t_sec_rel"] = pd.to_numeric(el["t_sec_rel"], errors="coerce")
    az[az_col] = pd.to_numeric(az[az_col], errors="coerce")
    el[el_col] = pd.to_numeric(el[el_col], errors="coerce")
    az = az.dropna().sort_values("t_sec_rel")
    el = el.dropna().sort_values("t_sec_rel")
    if az.empty or el.empty:
        logger.warning("Polar plots skipped: azimuth/elevation numeric samples are empty")
        return []

    for selector in ("input_level", "eb_no", "snr"):
        if selector not in wanted:
            continue
        metric_col, metric_df = _find_metric_section(section_frames, selector)
        if metric_df is None or metric_col not in metric_df:
            continue

        metric = metric_df[["t_sec_rel", metric_col]].copy()
        metric["t_sec_rel"] = pd.to_numeric(metric["t_sec_rel"], errors="coerce")
        metric[metric_col] = pd.to_numeric(metric[metric_col], errors="coerce")
        metric = metric.dropna().sort_values("t_sec_rel")
        if metric.empty:
            continue
        metric = metric.rename(columns={metric_col: "metric"})

        az_tmp = az[["t_sec_rel", az_col]].rename(columns={az_col: "azimuth"})
        el_tmp = el[["t_sec_rel", el_col]].rename(columns={el_col: "elevation"})
        aligned = _align_metric_with_az_el(metric, az_tmp, el_tmp)
        if aligned.empty:
            logger.warning(
                "Polar plot '%s' skipped: no overlapping/aligned time samples with azimuth/elevation",
                metric_col,
            )
            continue

        az_vals = np.mod(aligned["azimuth"].to_numpy(dtype=float), 360.0)
        el_vals = aligned["elevation"].to_numpy(dtype=float)
        metric_vals = aligned["metric"].to_numpy(dtype=float)
        valid = np.isfinite(az_vals) & np.isfinite(el_vals) & np.isfinite(metric_vals)
        valid &= (el_vals >= 0.0) & (el_vals <= 90.0)
        az_vals, el_vals, metric_vals = az_vals[valid], el_vals[valid], metric_vals[valid]
        if len(az_vals) < 3:
            logger.warning("Polar/3D plot '%s' skipped: insufficient valid angle samples", metric_col)
            continue

        # 2D polar styled like sky-map reference view.
        theta = np.deg2rad(az_vals)
        radius_norm = 1.0 - (el_vals / 90.0)
        track_az, track_el = _build_base_track(az_tmp, el_tmp)
        track_theta = np.deg2rad(track_az) if len(track_az) else np.array([])
        track_radius = 1.0 - (track_el / 90.0) if len(track_el) else np.array([])

        fig_p, ax_p = plt.subplots(subplot_kw={"projection": "polar"}, figsize=(8, 6))
        if len(track_theta):
            ax_p.plot(track_theta, track_radius, color="black", linewidth=1.6, alpha=0.95, zorder=1)
        sc_p = ax_p.scatter(theta, radius_norm, c=metric_vals, cmap="jet", s=34, zorder=3, edgecolors="none")
        ax_p.scatter(theta[:1], radius_norm[:1], c="#00AA88", marker="^", s=72, zorder=4)
        ax_p.scatter(theta[-1:], radius_norm[-1:], c="#66CCFF", marker="D", s=70, zorder=4)
        ax_p.set_theta_zero_location("N")
        ax_p.set_theta_direction(-1)
        ax_p.set_ylim(0.0, 1.02)
        ax_p.set_rticks([0.0, 0.5, 1.0])
        ax_p.set_yticklabels(["0", "0.5", "1"])
        ax_p.set_rlabel_position(18)
        ax_p.grid(alpha=0.35)
        ax_p.set_title(f"{metric_col} on antenna track")
        cbar_p = fig_p.colorbar(sc_p, ax=ax_p, pad=0.10)
        cbar_p.set_label("SNR (dB)" if "snr" in metric_col.lower() or "noise" in metric_col.lower() else metric_col)
        polar_name = f"{out_path.stem}_{metric_col}_polar.png"
        polar_path = out_path.with_name(polar_name)
        fig_p.savefig(polar_path, dpi=150, bbox_inches="tight")
        plt.close(fig_p)
        artifacts.append({"plot": metric_col, "kind": "polar", "path": str(polar_path)})

        # 3D spherical sky-view: antenna at center + hemisphere surface.
        x, y, z = _spherical_to_cartesian(az_vals, el_vals)
        tx, ty, tz = _spherical_to_cartesian(track_az, track_el) if len(track_az) else (np.array([]), np.array([]), np.array([]))

        fig3d = plt.figure(figsize=(9, 7))
        ax3d = fig3d.add_subplot(111, projection="3d")

        # hemisphere wireframe (z >= 0): antenna sky dome
        az_grid = np.linspace(0, 2 * np.pi, 72)
        el_grid = np.linspace(0, np.pi / 2, 28)
        AZ, EL = np.meshgrid(az_grid, el_grid)
        Xs = np.cos(EL) * np.cos(AZ)
        Ys = np.cos(EL) * np.sin(AZ)
        Zs = np.sin(EL)
        ax3d.plot_wireframe(Xs, Ys, Zs, rstride=3, cstride=6, color="lightgray", linewidth=0.5, alpha=0.45)

        # reference horizon circle
        hz = np.linspace(0, 2 * np.pi, 240)
        ax3d.plot(np.cos(hz), np.sin(hz), np.zeros_like(hz), color="gray", linewidth=1.0, alpha=0.8)

        # track and colored metric samples on sphere
        if len(tx):
            ax3d.plot(tx, ty, tz, color="black", alpha=0.75, linewidth=1.4)
        sc3d = ax3d.scatter(x, y, z, c=metric_vals, cmap="turbo", s=24, depthshade=False)

        # antenna center marker
        ax3d.scatter([0.0], [0.0], [0.0], c="black", s=42)
        ax3d.text(0.02, 0.02, 0.02, "Antenna", fontsize=8)

        ax3d.scatter([x[0]], [y[0]], [z[0]], c="white", edgecolors="black", s=60)
        ax3d.scatter([x[-1]], [y[-1]], [z[-1]], c="black", s=50)
        ax3d.set_title(f"{metric_col} 3D spherical sky-view")
        ax3d.set_xlabel("X")
        ax3d.set_ylabel("Y")
        ax3d.set_zlabel("Z")
        lim = 1.05
        ax3d.set_xlim(-lim, lim)
        ax3d.set_ylim(-lim, lim)
        ax3d.set_zlim(0.0, lim)
        ax3d.view_init(elev=24, azim=48)
        cbar3d = fig3d.colorbar(sc3d, ax=ax3d, pad=0.08)
        cbar3d.set_label(metric_col)

        png3d_name = f"{out_path.stem}_{metric_col}_3d.png"
        png3d_path = out_path.with_name(png3d_name)
        fig3d.savefig(png3d_path, dpi=150, bbox_inches="tight")
        plt.close(fig3d)
        artifacts.append({"plot": metric_col, "kind": "3d", "path": str(png3d_path)})

    if wanted and not artifacts:
        logger.warning(
            "Polar plots requested but no matching metric sections found. Requested=%s; available=%s",
            sorted(wanted),
            ", ".join(section_frames.keys()),
        )

    return artifacts


def _numeric_tick_groups(ticks: pd.DataFrame):
    """Split numeric ticks into one or two X-side groups (for dual Y axes)."""
    if ticks is None or ticks.empty:
        return []
    num = ticks[ticks.get("kind") == "num"].copy()
    if num.empty:
        return []
    num["val"] = num["text"].apply(lambda t: float(re.sub(r"[^0-9+\-.,]", "", str(t)).replace(",", ".")))
    num = num.dropna(subset=["x_px", "y_px", "val"])
    if len(num) < 2:
        return [num]

    xm = float(num["x_px"].median())
    left = num[num["x_px"] <= xm]
    right = num[num["x_px"] > xm]
    groups = []
    if len(left) >= 2:
        groups.append(left)
    if len(right) >= 2:
        groups.append(right)
    if not groups:
        groups.append(num)
    return groups


def _sanitize_curve_timebase(raw_df: pd.DataFrame):
    """Ensure curve is single-valued over x (time axis) by collapsing duplicate x samples."""
    if raw_df is None or raw_df.empty:
        return pd.DataFrame(columns=["x_px", "y_px"])
    df = raw_df[["x_px", "y_px"]].copy()
    df["x_px"] = pd.to_numeric(df["x_px"], errors="coerce")
    df["y_px"] = pd.to_numeric(df["y_px"], errors="coerce")
    df = df.dropna(subset=["x_px", "y_px"])
    if df.empty:
        return df

    # Collapse near-duplicate x coordinates that create vertical jumps.
    df["x_key"] = df["x_px"].round(2)
    df = (
        df.groupby("x_key", as_index=False)
        .agg({"x_px": "median", "y_px": "median"})
        .sort_values("x_px")
        .reset_index(drop=True)
    )
    return df[["x_px", "y_px"]]


def _map_curve_with_ticks_group(raw_df: pd.DataFrame, tick_group: pd.DataFrame, out_col: str):
    """Map y_px to values using a specific numeric tick group."""
    if raw_df.empty or tick_group is None or tick_group.empty:
        tmp = raw_df.copy()
        tmp[out_col] = np.nan
        return tmp
    Y = np.vstack([tick_group["y_px"].values, np.ones(len(tick_group))]).T
    a, b = np.linalg.lstsq(Y, tick_group["val"].values, rcond=None)[0]
    tmp = raw_df.copy()
    tmp[out_col] = a * tmp["y_px"] + b
    return tmp


def _regularize_antenna_series(df: pd.DataFrame, az_col: str, el_col: str):
    """Regularize antenna az/el time series to reduce extraction jitter."""
    out = df.copy()
    out = out.sort_values("t_sec_rel")

    # Collapse duplicated/near-duplicated timestamps by median value.
    out["t_key"] = out["t_sec_rel"].round(3)
    agg = out.groupby("t_key", as_index=False).agg({
        "t_sec_rel": "median",
        "time_HH:MM:SS": "first",
        "time_iso_utc": "first",
        "x_px_az": "median",
        "y_px_az": "median",
        az_col: "median",
        "x_px_el": "median",
        "y_px_el": "median",
        el_col: "median",
    })

    az = pd.to_numeric(agg[az_col], errors="coerce").to_numpy(dtype=float)
    el = pd.to_numeric(agg[el_col], errors="coerce").to_numpy(dtype=float)
    if len(agg) < 8:
        return agg.drop(columns=["t_key"], errors="ignore")

    # Smooth elevation with piecewise monotonic profile (up then down).
    peak = int(np.nanargmax(el)) if np.isfinite(el).any() else len(el) // 2
    el_up = np.maximum.accumulate(el[: peak + 1])
    el_dn = np.minimum.accumulate(el[peak:])
    el_s = np.concatenate([el_up, el_dn[1:]]) if len(el_dn) > 1 else el_up
    el_s = np.clip(el_s, 0.0, 90.0)

    # Smooth azimuth as mostly monotonic in unwrapped space.
    az_u = np.rad2deg(np.unwrap(np.deg2rad(az)))
    direction = np.nanmedian(np.diff(az_u))
    if np.isnan(direction):
        direction = 0.0
    if direction >= 0:
        az_s = np.maximum.accumulate(az_u)
    else:
        az_s = np.minimum.accumulate(az_u)
    az_s = np.mod(az_s, 360.0)

    agg[az_col] = az_s
    agg[el_col] = el_s
    return agg.drop(columns=["t_key"], errors="ignore")


def build_antenna_combined_df(ycol, curves, start_dt, stop_dt):
    """Build a single antenna dataframe with azimuth and elevation columns."""
    mapped = []
    for idx, (_series_name, raw_df, ticks) in enumerate(curves, start=1):
        base_curve = _sanitize_curve_timebase(raw_df.copy())
        base = map_x_to_time(base_curve, start_dt, stop_dt)
        if base.empty:
            continue

        candidates = []
        # candidate 1: all numeric ticks together
        all_num = ticks[ticks.get("kind") == "num"].copy() if ticks is not None and not ticks.empty else pd.DataFrame()
        if not all_num.empty:
            all_num["val"] = all_num["text"].apply(lambda t: float(re.sub(r"[^0-9+\-.,]", "", str(t)).replace(",", ".")))
            all_num = all_num.dropna(subset=["x_px", "y_px", "val"])
            if len(all_num) >= 2:
                dfa = _map_curve_with_ticks_group(base, all_num, f"value_{idx}_all")
                candidates.append((dfa, f"value_{idx}_all"))

        # candidate 2..n: per-side numeric tick groups (dual-axis charts)
        for g_i, grp in enumerate(_numeric_tick_groups(ticks), start=1):
            dfg = _map_curve_with_ticks_group(base, grp, f"value_{idx}_g{g_i}")
            candidates.append((dfg, f"value_{idx}_g{g_i}"))

        for cand_df, cand_col in candidates:
            vals = pd.to_numeric(cand_df[cand_col], errors="coerce")
            if vals.dropna().empty:
                continue
            mapped.append((idx, cand_df, vals, cand_col))

    if len(mapped) < 2:
        return None

    # choose azimuth/elevation with value-domain aware scoring.
    scored = []
    for idx, df, vals, val_col in mapped:
        tmp = df[["t_sec_rel", val_col]].copy()
        tmp = tmp.dropna().sort_values("t_sec_rel")
        if len(tmp) < 8:
            continue
        v = pd.to_numeric(tmp[val_col], errors="coerce").dropna().to_numpy(dtype=float)
        if len(v) < 8:
            continue
        vmin, vmax = float(np.min(v)), float(np.max(v))
        vrng = float(vmax - vmin)

        d1 = np.diff(v)
        nz = d1[np.abs(d1) > 1e-9]
        if len(nz) == 0:
            continue
        frac_pos = float(np.mean(nz > 0))
        frac_neg = float(np.mean(nz < 0))
        mono_score = max(frac_pos, frac_neg)

        # bell-shape score: signs should go + ... - with at most one transition.
        signs = np.sign(nz)
        trans = np.sum(signs[1:] * signs[:-1] < 0)
        bell_score = 1.0 / (1.0 + trans)

        frac_az = float(np.mean((v >= -5) & (v <= 365)))
        frac_el = float(np.mean((v >= -2) & (v <= 92)))

        az_score = 2.0 * frac_az + 1.8 * mono_score + (1.2 if vmax > 120 else 0.0) + 0.001 * vrng
        el_score = 2.0 * frac_el + 1.7 * bell_score + (1.0 if 5 <= vmax <= 95 else -0.8) - 0.001 * max(0.0, vrng - 90.0)

        scored.append((idx, df, val_col, az_score, el_score, vmin, vmax, vrng, mono_score, bell_score))

    if len(scored) < 2:
        return None

    az_item = max(scored, key=lambda t: t[3])
    el_candidates = [t for t in scored if t[0] != az_item[0]]
    if not el_candidates:
        # fallback to different mapping candidate from same raw curve
        el_candidates = [t for t in scored if t[2] != az_item[2]]
        if not el_candidates:
            return None
    el_item = max(el_candidates, key=lambda t: t[4])

    az_idx, az_df, az_col = az_item[0], az_item[1], az_item[2]
    el_idx, el_df, el_col = el_item[0], el_item[1], el_item[2]

    az = az_df[["t_sec_rel", "time_HH:MM:SS", "time_iso_utc", "x_px", "y_px", az_col]].copy()
    az = az.rename(columns={"x_px": "x_px_az", "y_px": "y_px_az", az_col: f"{ycol}_azimuth"})
    el = el_df[["t_sec_rel", "x_px", "y_px", el_col]].copy()
    el = el.rename(columns={"x_px": "x_px_el", "y_px": "y_px_el", el_col: f"{ycol}_elevation"})

    az["t_sec_rel"] = pd.to_numeric(az["t_sec_rel"], errors="coerce")
    el["t_sec_rel"] = pd.to_numeric(el["t_sec_rel"], errors="coerce")
    az = az.dropna(subset=["t_sec_rel"]).sort_values("t_sec_rel")
    el = el.dropna(subset=["t_sec_rel"]).sort_values("t_sec_rel")
    if az.empty or el.empty:
        return None

    merged = pd.merge_asof(az, el, on="t_sec_rel", direction="nearest")
    # sanitize physical ranges for pointing angles
    merged[f"{ycol}_azimuth"] = pd.to_numeric(merged[f"{ycol}_azimuth"], errors="coerce") % 360.0
    merged[f"{ycol}_elevation"] = pd.to_numeric(merged[f"{ycol}_elevation"], errors="coerce").clip(0.0, 90.0)
    merged = merged.dropna(subset=[f"{ycol}_azimuth", f"{ycol}_elevation"])
    merged = _regularize_antenna_series(merged, f"{ycol}_azimuth", f"{ycol}_elevation")
    return merged


# -------------------- Main --------------------

def process_html(
    html_path: Path,
    output_dir: Path,
    stats_selectors=None,
    stats_rows=None,
    plot_selectors=None,
    plot_rows=None,
) -> Path:
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

    # Sezioni target: cerca dinamicamente tutti gli header h2/h3 e verifica
    # se contengono grafici (svg/object) prima del prossimo header.
    targets = []
    for hdr in soup.find_all(["h2", "h3"]):
        title = hdr.get_text(" ", strip=True)
        if not title:
            continue
        cur = hdr
        has_chart = False
        while True:
            cur = cur.find_next_sibling()
            if cur is None or cur.name in ("h2", "h3"):
                break
            if cur.find("svg") or cur.find("object", type="image/svg+xml"):
                has_chart = True
                break
        if not has_chart:
            continue
        key = re.sub(r"\W+", "_", title.lower()).strip("_") or "section"
        EXCLUDE = {"session", "activities", "channel"}
        if key in EXCLUDE or any(key.endswith(f"_{e}") for e in EXCLUDE):
            continue
        targets.append((hdr, key))

    # Nome file in base a (prefix, orbit_no) trovati nell'HTML
    prefix, orbit_no = derive_orbit_filename(soup)
    base = (
        f"{prefix}_orbit_{orbit_no}" if prefix and orbit_no else
        f"{prefix}_orbit" if prefix and not orbit_no else
        f"orbit_{orbit_no}" if orbit_no else
        "orbit"
    )

    # writer: salva sempre in .xlsx
    out_path = output_dir / (base + ".xlsx")
    section_frames = {}
    with pd.ExcelWriter(out_path, engine="openpyxl") as wr:
        # Meta
        pd.DataFrame([
            {
                "start_time_utc": start_dt.strftime("%Y-%m-%d %H:%M:%S") if start_dt else None,
                "stop_time_utc": stop_dt.strftime("%Y-%m-%d %H:%M:%S") if stop_dt else None,
                "report_time_utc": rep_dt.strftime("%Y-%m-%d %H:%M:%S") if rep_dt else None,
            }
        ]).to_excel(wr, sheet_name="__meta__", index=False)

        # Per ogni sezione, estrai e salva in un foglio
        def write_section(sheet_key, df, ticks):
            df = map_x_to_time(df, start_dt, stop_dt)
            df = map_y_from_ticks(df, ticks, colname=sheet_key)
            section_frames[sheet_key] = df.copy()

            cols = ["x_px", "y_px", "t_sec_rel", "time_HH:MM:SS", "time_iso_utc", sheet_key]
            df = df.reindex(cols, axis=1)
            sheet = safe_sheet_name(sheet_key)
            if df.empty:
                pd.DataFrame([{"note": "nessun dato estratto"}]).to_excel(wr, sheet_name=sheet, index=False)
            else:
                df.to_excel(wr, sheet_name=sheet, index=False)

            tname = safe_sheet_name(sheet_key + "_ticks")
            (ticks if not ticks.empty else pd.DataFrame([{"note": "no ticks"}])).to_excel(
                wr, sheet_name=tname, index=False
            )

        for hdr, ycol in targets:
            is_antenna = any(k in ycol.lower() for k in ("antenna", "azimuth", "elevation"))
            if is_antenna:
                multi = extract_curves_for_header(hdr)
                combined = build_antenna_combined_df(ycol, multi, start_dt, stop_dt) if multi else None
                if combined is not None:
                    section_frames[f"{ycol}_azimuth"] = combined[["t_sec_rel", f"{ycol}_azimuth"]].copy()
                    section_frames[f"{ycol}_elevation"] = combined[["t_sec_rel", f"{ycol}_elevation"]].copy()
                    sheet = safe_sheet_name(ycol)
                    combined.to_excel(wr, sheet_name=sheet, index=False)
                    # Save ticks of first detected curve as reference for this combined sheet
                    tname = safe_sheet_name(ycol + "_ticks")
                    ref_ticks = multi[0][2] if multi else pd.DataFrame([{"note": "no ticks"}])
                    (ref_ticks if not ref_ticks.empty else pd.DataFrame([{"note": "no ticks"}])).to_excel(
                        wr, sheet_name=tname, index=False
                    )
                    continue

            df, ticks = extract_curve_for_header(hdr)
            write_section(ycol, df, ticks)
    wb = wr.book
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    if stats_rows is not None:
        stats_rows.extend(summarize_selected_stats(orbit_no, section_frames, stats_selectors))

    if plot_rows is not None:
        plot_rows.extend(generate_polar_plot_artifacts(out_path, section_frames, plot_selectors))

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
