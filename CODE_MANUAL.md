# Code Manual

This manual provides a line-by-line and block-level explanation of the main source files in this project. Each section reproduces the full code with line numbers followed by detailed commentary on the implementation, data structures, and Python concepts involved.

## Extract_all_charts.py

### Code
```python
     1	# extract_all_charts.py
     2	# ------------------------------------------
     3	# Estrae i grafici dal report HTML (sezioni: Input Level, SNR, Eb/N0,
     4	# Carrier Offset, Phase Loop Error, Reed-Solomon Frames, Frame Error Rate,
     5	# Demodulator/FEP Lock State), mappa X a Start→Stop (secondi),
     6	# ricostruisce Y dai tick numerici, rimuove segmenti spurii (legend/diagonale)
     7	# e salva un unico Excel con un foglio per grafico.
     8	#
     9	# Nome file: <prefix>_orbit_<num>.xlsx (prefix e orbit number letti dall'HTML)
    10	#   - Se manca il prefix: "orbit_<num>..."
    11	#   - Se manca anche il numero: "orbit..."
    12	#
    13	# Uso standalone:
    14	#   - Metti "report.html" nella stessa cartella
    15	#   - Esegui:  python extract_all_charts.py
    16	#
    17	# Dipendenze:
    18	#   pip install beautifulsoup4 pandas numpy openpyxl
    19	# ------------------------------------------
    20	
    21	from pathlib import Path
    22	import argparse
    23	import re
    24	from datetime import datetime, timezone, timedelta
    25	import logging
    26	import base64
    27	import urllib.request
    28	
    29	import numpy as np
    30	import pandas as pd
    31	from bs4 import BeautifulSoup, FeatureNotFound
    32	
    33	
    34	DEFAULT_HTML = Path("report.html")  # used if directory lacks .html
    35	
    36	logging.basicConfig(level=logging.INFO)
    37	logger = logging.getLogger(__name__)
    38	
    39	
    40	# -------------------- Utilities --------------------
    41	
    42	def parse_iso_utc(s: str):
    43	    """Parse 'YYYY-MM-DD HH:MM:SSZ' → datetime tz-aware (UTC)."""
    44	    if not s:
    45	        return None
    46	    s = s.strip()
    47	    m = re.match(r"(\d{4}-\d{2}-\d{2})\s+(\d{2}):(\d{2}):(\d{2})Z", s)
    48	    if not m:
    49	        return None
    50	    y, mo, d = map(int, m.group(1).split("-"))
    51	    hh, mm, ss = int(m.group(2)), int(m.group(3)), int(m.group(4))
    52	    return datetime(y, mo, d, hh, mm, ss, tzinfo=timezone.utc)
    53	
    54	
    55	def parse_transforms(transform: str):
    56	    """Accorpa scale()/translate() in (sx, sy, tx, ty)."""
    57	    sx, sy, tx, ty = 1.0, 1.0, 0.0, 0.0
    58	    if not transform:
    59	        return sx, sy, tx, ty
    60	    for m in re.finditer(r"(translate|scale)\(\s*([^)]+)\)", transform):
    61	        kind = m.group(1)
    62	        args = [float(v) for v in re.split(r"[, \t]+", m.group(2).strip()) if v]
    63	        if kind == "scale":
    64	            if len(args) == 1:
    65	                sx *= args[0]; sy *= args[0]
    66	            else:
    67	                sx *= args[0]; sy *= args[1]
    68	        else:  # translate
    69	            if len(args) == 1:
    70	                tx += args[0]
    71	            else:
    72	                tx += args[0]; ty += args[1]
    73	    return sx, sy, tx, ty
    74	
    75	
    76	def cumulative_transform(tag):
    77	    """Accumula scale/translate lungo la gerarchia SVG del nodo."""
    78	    Sx, Sy, Tx, Ty = 1.0, 1.0, 0.0, 0.0
    79	    chain = []
    80	    cur = tag
    81	    while cur is not None and getattr(cur, "name", None) is not None:
    82	        tr = cur.get("transform")
    83	        if tr:
    84	            chain.append(tr)
    85	        cur = cur.parent
    86	    for tr in reversed(chain):
    87	        sx, sy, tx, ty = parse_transforms(tr)
    88	        Tx = sx * Tx + tx
    89	        Ty = sy * Ty + ty
    90	        Sx *= sx
    91	        Sy *= sy
    92	    return Sx, Sy, Tx, Ty
    93	
    94	
    95	def apply_tr(x, y, Sx, Sy, Tx, Ty):
    96	    return Sx * x + Tx, Sy * y + Ty
    97	
    98	
    99	def parse_path_subpaths(d_attr: str):
   100	    """
   101	    Parsea un 'path' SVG supportando M/m L/l H/h V/v e lo spezza in subpath.
   102	    Un nuovo subpath inizia a ogni 'M'/'m'. Ritorna lista di liste di (x,y).
   103	    """
   104	    tokens = re.findall(r"([MmLlHhVv])|([-+]?\d*\.?\d+(?:e[-+]?\d+)?)", d_attr or "")
   105	    flat = []
   106	    for a, b in tokens:
   107	        if a:
   108	            flat.append(a)
   109	        else:
   110	            flat.append(float(b))
   111	
   112	    subpaths = []
   113	    x = y = 0.0
   114	    cmd = None
   115	    i = 0
   116	    current = []
   117	
   118	    def flush():
   119	        nonlocal current
   120	        if len(current) >= 2:
   121	            subpaths.append(current)
   122	        current = []
   123	
   124	    while i < len(flat):
   125	        t = flat[i]
   126	        if isinstance(t, str):
   127	            cmd = t
   128	            if cmd in ("M", "m"):
   129	                flush()
   130	            i += 1
   131	            continue
   132	
   133	        if cmd == "M":
   134	            if i + 1 < len(flat) and not isinstance(flat[i + 1], str):
   135	                x, y = flat[i], flat[i + 1]
   136	                current = [(x, y)]
   137	                i += 2
   138	            else:
   139	                i += 1
   140	        elif cmd == "m":
   141	            if i + 1 < len(flat) and not isinstance(flat[i + 1], str):
   142	                x += flat[i]; y += flat[i + 1]
   143	                current = [(x, y)]
   144	                i += 2
   145	            else:
   146	                i += 1
   147	        elif cmd == "L":
   148	            if i + 1 < len(flat) and not isinstance(flat[i + 1], str):
   149	                x, y = flat[i], flat[i + 1]
   150	                current.append((x, y))
   151	                i += 2
   152	            else:
   153	                i += 1
   154	        elif cmd == "l":
   155	            if i + 1 < len(flat) and not isinstance(flat[i + 1], str):
   156	                x += flat[i]; y += flat[i + 1]
   157	                current.append((x, y))
   158	                i += 2
   159	            else:
   160	                i += 1
   161	        elif cmd == "H":
   162	            x = flat[i]; current.append((x, y)); i += 1
   163	        elif cmd == "h":
   164	            x += flat[i]; current.append((x, y)); i += 1
   165	        elif cmd == "V":
   166	            y = flat[i]; current.append((x, y)); i += 1
   167	        elif cmd == "v":
   168	            y += flat[i]; current.append((x, y)); i += 1
   169	        else:
   170	            i += 1
   171	
   172	    flush()
   173	    return subpaths
   174	
   175	
   176	def svg_axes_from_ticks(svg):
   177	    """
   178	    Estrae tick (label testo) e posizioni pixel assolute.
   179	    Definisce il riquadro assi:
   180	      - X: min/max posizione pixel dei tick orari (HH:MM o HH:MM:SS)
   181	      - Y: min/max posizione pixel dei tick numerici
   182	    """
   183	    rows = []
   184	    for t in svg.find_all("text"):
   185	        content = (t.get_text() or "").strip()
   186	        if not content:
   187	            continue
   188	        is_num = re.fullmatch(r"[-+]?\d+(?:[.,]\d+)?(?:\s*(dB|°|deg))?", content) is not None
   189	        is_time = (
   190	            re.fullmatch(r"\d{2}:\d{2}", content) is not None or
   191	            re.fullmatch(r"\d{2}:\d{2}:\d{2}", content) is not None
   192	        )
   193	        if not (is_num or is_time):
   194	            continue
   195	        Sx, Sy, Tx, Ty = cumulative_transform(t)
   196	        x_px, y_px = apply_tr(0.0, 0.0, Sx, Sy, Tx, Ty)
   197	        rows.append({"text": content, "x_px": x_px, "y_px": y_px, "kind": "num" if is_num else "time"})
   198	
   199	    ticks = pd.DataFrame(rows)
   200	
   201	    if ticks.empty:
   202	        return ticks, (None, None, None, None)
   203	
   204	    x_tick_px_min = ticks.loc[ticks["kind"] == "time", "x_px"].min()
   205	    x_tick_px_max = ticks.loc[ticks["kind"] == "time", "x_px"].max()
   206	    y_tick_px_min = ticks.loc[ticks["kind"] == "num", "y_px"].min()
   207	    y_tick_px_max = ticks.loc[ticks["kind"] == "num", "y_px"].max()
   208	
   209	    return ticks, (x_tick_px_min, x_tick_px_max, y_tick_px_min, y_tick_px_max)
   210	
   211	
   212	def extract_curve_for_header(hdr):
   213	    """
   214	    Per una sezione (h2/h3) già individuata, raccoglie gli SVG sottostanti fino
   215	    al prossimo h2/h3 e sceglie il sottopercorso dati migliore (massimo numero
   216	    di punti dentro il riquadro assi).
   217	    """
   218	    if hdr is None:
   219	        return pd.DataFrame(), pd.DataFrame()
   220	
   221	    svgs = []
   222	    for el in hdr.next_elements:
   223	        name = getattr(el, "name", None)
   224	        if name in ("h2", "h3"):
   225	            break
   226	        if name == "svg":
   227	            svgs.append(el)
   228	        elif name == "object" and el.get("type") == "image/svg+xml":
   229	            data = el.get("data", "")
   230	            m = re.match(r"^data:image/svg\+xml(;charset=[^;]+)?;base64,(.*)$", data, re.I)
   231	            try:
   232	                if m:  # Base64 inline data
   233	                    svg_bytes = base64.b64decode(m.group(2))
   234	                elif data.startswith(("http://", "https://")):  # Remote file
   235	                    with urllib.request.urlopen(data) as resp:
   236	                        svg_bytes = resp.read()
   237	                elif data:  # Local file path
   238	                    with open(data, "rb") as f:
   239	                        svg_bytes = f.read()
   240	                else:
   241	                    continue
   242	                try:
   243	                    svg_soup = BeautifulSoup(svg_bytes, "xml")
   244	                except FeatureNotFound:
   245	                    logger.warning(
   246	                        "lxml parser not found; falling back to html.parser. Install lxml for full XML support."
   247	                    )
   248	                    try:
   249	                        svg_soup = BeautifulSoup(svg_bytes, "html.parser")
   250	                    except FeatureNotFound:
   251	                        import xml.etree.ElementTree as ET
   252	
   253	                        svg_soup = BeautifulSoup(
   254	                            ET.tostring(ET.fromstring(svg_bytes)), "html.parser"
   255	                        )
   256	                if svg_soup.svg:
   257	                    svgs.append(svg_soup.svg)
   258	            except Exception as exc:
   259	                logger.warning("Failed to load SVG from %s: %s", data, exc)
   260	    if not svgs:
   261	        return pd.DataFrame(), pd.DataFrame()
   262	
   263	    # Scegli lo svg con più gruppi "gnuplot_plot_*"
   264	    def count_groups(s):
   265	        return len([g for g in s.find_all("g") if (g.get("id") or "").startswith("gnuplot_plot_")])
   266	
   267	    best_svg = max(svgs, key=count_groups)
   268	    ticks, axes = svg_axes_from_ticks(best_svg)
   269	    x_min_tick, x_max_tick, y_min_tick, y_max_tick = axes
   270	
   271	    best_pts = []
   272	    best_score = -1
   273	
   274	    groups = [
   275	        g
   276	        for g in best_svg.find_all("g")
   277	        if (g.get("id") or "").startswith("gnuplot_plot_") and (g.find("path") or g.find("polyline"))
   278	    ]
   279	    if not groups:
   280	        groups = [g for g in best_svg.find_all("g") if g.find("path") or g.find("polyline")]
   281	    for g in groups:
   282	        # PATH: split in subpath e valuta punti dentro assi
   283	        for p in g.find_all("path"):
   284	            d = p.get("d")
   285	            if not d:
   286	                continue
   287	            Sx, Sy, Tx, Ty = cumulative_transform(p)
   288	            for sp in parse_path_subpaths(d):
   289	                pts = [apply_tr(x, y, Sx, Sy, Tx, Ty) for x, y in sp]
   290	                if None in (x_min_tick, x_max_tick, y_min_tick, y_max_tick):
   291	                    score = len(pts)
   292	                else:
   293	                    m = 2.0
   294	                    score = sum(
   295	                        (x_min_tick - m <= x <= x_max_tick + m) and
   296	                        (y_min_tick - m <= y <= y_max_tick + m)
   297	                        for x, y in pts
   298	                    )
   299	                if score > best_score:
   300	                    best_pts = pts
   301	                    best_score = score
   302	
   303	        # POLYLINE: fallback
   304	        for pl in g.find_all("polyline"):
   305	            raw = (pl.get("points") or "").strip()
   306	            if not raw:
   307	                continue
   308	            raw = re.sub(r"\s+", " ", raw)
   309	            pairs = re.findall(
   310	                r"([-+]?\d*\.?\d+(?:e[-+]?\d+)?)\s*,\s*([-+]?\d*\.?\d+(?:e[-+]?\d+)?)",
   311	                raw
   312	            )
   313	            pts_local = [(float(x), float(y)) for x, y in pairs]
   314	            Sx, Sy, Tx, Ty = cumulative_transform(pl)
   315	            pts = [apply_tr(x, y, Sx, Sy, Tx, Ty) for x, y in pts_local]
   316	
   317	            if None in (x_min_tick, x_max_tick, y_min_tick, y_max_tick):
   318	                score = len(pts)
   319	            else:
   320	                m = 2.0
   321	                score = sum(
   322	                    (x_min_tick - m <= x <= x_max_tick + m) and
   323	                    (y_min_tick - m <= y <= y_max_tick + m)
   324	                    for x, y in pts
   325	                )
   326	            if score > best_score:
   327	                best_pts = pts
   328	                best_score = score
   329	
   330	    curve_px = pd.DataFrame(best_pts, columns=["x_px", "y_px"])
   331	    return curve_px, ticks
   332	
   333	
   334	def map_x_to_time(df: pd.DataFrame, start_dt: datetime, stop_dt: datetime):
   335	    """Mappa x_px in tempo assoluto usando Start/Stop (in secondi)."""
   336	    if df.empty or not start_dt or not stop_dt:
   337	        df["t_sec_rel"] = np.nan
   338	        df["time_HH:MM:SS"] = None
   339	        df["time_iso_utc"] = None
   340	        return df
   341	
   342	    dur = (stop_dt - start_dt).total_seconds()
   343	    x_min = df["x_px"].min()
   344	    x_max = df["x_px"].max()
   345	    if x_max == x_min:
   346	        df["t_sec_rel"] = 0.0
   347	        df["time_HH:MM:SS"] = start_dt.strftime("%H:%M:%S")
   348	        df["time_iso_utc"] = start_dt.strftime("%Y-%m-%d %H:%M:%S")
   349	        return df
   350	
   351	    df["t_sec_rel"] = (df["x_px"] - x_min) / (x_max - x_min) * dur
   352	    df["time_iso_utc"] = df["t_sec_rel"].apply(
   353	        lambda s: (start_dt + timedelta(seconds=float(s))).strftime("%Y-%m-%d %H:%M:%S")
   354	    )
   355	    df["time_HH:MM:SS"] = df["t_sec_rel"].apply(
   356	        lambda s: (start_dt + timedelta(seconds=float(s))).strftime("%H:%M:%S")
   357	    )
   358	    return df
   359	
   360	
   361	def map_y_from_ticks(df: pd.DataFrame, ticks: pd.DataFrame, colname: str):
   362	    """Fit lineare: y_px → valore asse (da tick numerici)."""
   363	    if df.empty or ticks is None or ticks.empty:
   364	        df[colname] = np.nan
   365	        return df
   366	    y_ticks = ticks[ticks["kind"] == "num"].copy()
   367	    if y_ticks.empty:
   368	        df[colname] = np.nan
   369	        return df
   370	    y_ticks["value"] = y_ticks["text"].apply(
   371	        lambda s: float(re.sub(r"[^0-9+\-.,]", "", s).replace(",", "."))
   372	    )
   373	    Y = np.vstack([y_ticks["y_px"].values, np.ones(len(y_ticks))]).T
   374	    a, b = np.linalg.lstsq(Y, y_ticks["value"].values, rcond=None)[0]
   375	    df[colname] = a * df["y_px"] + b
   376	    return df
   377	
   378	
   379	def safe_sheet_name(title: str):
   380	    """Excel: max 31 char, niente []:*?/\\ ."""
   381	    forbidden = set('[]:*?/\\')
   382	    t = ''.join('_' if ch in forbidden else ch for ch in title)
   383	    t = t.strip()
   384	    return t[:31] if len(t) > 31 else t
   385	
   386	
   387	def derive_orbit_filename(soup: BeautifulSoup):
   388	    """
   389	    Ricava (prefix, orbit_no) dall'HTML, senza fallback fissi.
   390	    - orbit_no: 'orbit 1234' / 'Orbit: 1234'
   391	    - prefix: prova in tabella "Session" (Spacecraft/Satellite/Mission/Name/Platform/...),
   392	              quindi <title> e header h1/h2/h3, infine pattern tipo 'AWS-PFM' nel testo.
   393	      Se non si trova nulla, prefix=None.
   394	    """
   395	    txt = soup.get_text(" ", strip=True)
   396	
   397	    # Orbit number
   398	    m = re.search(r"\borbit\s*[:#]?\s*(\d{1,7})\b", txt, flags=re.I)
   399	    orbit_no = m.group(1) if m else None
   400	
   401	    candidates = []
   402	
   403	    # 1) Dalla tabella Session
   404	    h2 = soup.find(id="_session")
   405	    if h2:
   406	        tbl = h2.find_next("table")
   407	        if tbl:
   408	            for row in tbl.find_all("tr"):
   409	                cells = row.find_all(["td", "th"])
   410	                if len(cells) < 2:
   411	                    continue
   412	                key = cells[0].get_text(" ", strip=True).lower()
   413	                val = cells[1].get_text(" ", strip=True)
   414	                if not val:
   415	                    continue
   416	                if any(k in key for k in (
   417	                    "spacecraft", "satellite", "mission", "name", "platform", "asset", "receiver", "modem"
   418	                )):
   419	                    candidates.append(val)
   420	
   421	    # 2) <title> + headers
   422	    if soup.title and soup.title.get_text(strip=True):
   423	        candidates.append(soup.title.get_text(" ", strip=True))
   424	    for hdr in soup.find_all(["h1", "h2", "h3"]):
   425	        t = hdr.get_text(" ", strip=True)
   426	        if t:
   427	            candidates.append(t)
   428	
   429	    # 3) pattern tipo AWS-PFM da testo generale
   430	    m2 = re.search(r"\b([A-Z]{2,}(?:-[A-Z0-9]{2,})+)\b", txt)
   431	    if m2:
   432	        candidates.append(m2.group(1))
   433	
   434	    prefix = None
   435	    for raw in candidates:
   436	        cleaned = re.sub(r"[^A-Za-z0-9_-]+", " ", raw).strip()
   437	        m = re.search(r"\b([A-Za-z0-9]+(?:[-_][A-Za-z0-9]+)+)\b", cleaned)
   438	        if not m:
   439	            m = re.search(r"\b([A-Z0-9]{3,})\b", cleaned)
   440	        if m:
   441	            prefix = m.group(1)
   442	            break
   443	
   444	    return prefix, orbit_no
   445	
   446	
   447	# -------------------- Main --------------------
   448	
   449	def process_html(html_path: Path, output_dir: Path) -> Path:
   450	    """Elabora un report HTML e salva i grafici in un file Excel.
   451	
   452	    Parameters
   453	    ----------
   454	    html_path : Path
   455	        Percorso del file HTML del report.
   456	    output_dir : Path
   457	        Directory in cui salvare l'Excel risultante.
   458	
   459	    Returns
   460	    -------
   461	    Path
   462	        Percorso del file Excel creato.
   463	    """
   464	    html = Path(html_path)
   465	    if not html.exists():
   466	        alt = Path(__file__).resolve().parent / html.name
   467	        if alt.exists():
   468	            html = alt
   469	        else:
   470	            raise FileNotFoundError(
   471	                f"File HTML non trovato: {html_path} (cwd) o {alt} (script dir)"
   472	            )
   473	
   474	    output_dir = Path(output_dir)
   475	    output_dir.mkdir(parents=True, exist_ok=True)
   476	
   477	    with html.open("r", encoding="utf-8") as f:
   478	        soup = BeautifulSoup(f, "html.parser")
   479	
   480	    # Tempi di sessione
   481	    start_dt = stop_dt = rep_dt = None
   482	    h2 = soup.find(id="_session")
   483	    if h2:
   484	        tbl = h2.find_next("table")
   485	        if tbl:
   486	            for row in tbl.find_all("tr"):
   487	                cells = row.find_all(["td", "th"])
   488	                if len(cells) < 2:
   489	                    continue
   490	                label = cells[0].get_text(" ", strip=True).lower()
   491	                value = cells[1].get_text(" ", strip=True)
   492	                if "start time" in label:
   493	                    start_dt = parse_iso_utc(value)
   494	                elif "stop time" in label:
   495	                    stop_dt = parse_iso_utc(value)
   496	                elif "report creation time" in label:
   497	                    rep_dt = parse_iso_utc(value)
   498	
   499	    # Sezioni target: cerca dinamicamente tutti gli header h2/h3 e verifica
   500	    # se contengono grafici (svg/object) prima del prossimo header.
   501	    targets = []
   502	    for hdr in soup.find_all(["h2", "h3"]):
   503	        title = hdr.get_text(" ", strip=True)
   504	        if not title:
   505	            continue
   506	        cur = hdr
   507	        has_chart = False
   508	        while True:
   509	            cur = cur.find_next_sibling()
   510	            if cur is None or cur.name in ("h2", "h3"):
   511	                break
   512	            if cur.find("svg") or cur.find("object", type="image/svg+xml"):
   513	                has_chart = True
   514	                break
   515	        if not has_chart:
   516	            continue
   517	        key = re.sub(r"\W+", "_", title.lower()).strip("_") or "section"
   518	        EXCLUDE = {"session", "activities", "channel"}
   519	        if key in EXCLUDE or any(key.endswith(f"_{e}") for e in EXCLUDE):
   520	            continue
   521	        targets.append((hdr, key))
   522	
   523	    # Nome file in base a (prefix, orbit_no) trovati nell'HTML
   524	    prefix, orbit_no = derive_orbit_filename(soup)
   525	    base = (
   526	        f"{prefix}_orbit_{orbit_no}" if prefix and orbit_no else
   527	        f"{prefix}_orbit" if prefix and not orbit_no else
   528	        f"orbit_{orbit_no}" if orbit_no else
   529	        "orbit"
   530	    )
   531	
   532	    # writer: salva sempre in .xlsx
   533	    out_path = output_dir / (base + ".xlsx")
   534	    with pd.ExcelWriter(out_path, engine="openpyxl") as wr:
   535	        # Meta
   536	        pd.DataFrame([
   537	            {
   538	                "start_time_utc": start_dt.strftime("%Y-%m-%d %H:%M:%S") if start_dt else None,
   539	                "stop_time_utc": stop_dt.strftime("%Y-%m-%d %H:%M:%S") if stop_dt else None,
   540	                "report_time_utc": rep_dt.strftime("%Y-%m-%d %H:%M:%S") if rep_dt else None,
   541	            }
   542	        ]).to_excel(wr, sheet_name="__meta__", index=False)
   543	
   544	        # Per ogni sezione, estrai e salva in un foglio
   545	        for hdr, ycol in targets:
   546	            df, ticks = extract_curve_for_header(hdr)
   547	            df = map_x_to_time(df, start_dt, stop_dt)
   548	            df = map_y_from_ticks(df, ticks, colname=ycol)
   549	
   550	            cols = [
   551	                "x_px",
   552	                "y_px",
   553	                "t_sec_rel",
   554	                "time_HH:MM:SS",
   555	                "time_iso_utc",
   556	                ycol,
   557	            ]
   558	            df = df.reindex(cols, axis=1)
   559	            if not df.empty:
   560	                for col in ("time_HH:MM:SS", "time_iso_utc"):
   561	                    if col in df.columns:
   562	                        df[col] = pd.to_datetime(df[col])
   563	
   564	            sheet = safe_sheet_name(ycol)
   565	            if df.empty:
   566	                pd.DataFrame([{"note": "nessun dato estratto"}]).to_excel(
   567	                    wr, sheet_name=sheet, index=False
   568	                )
   569	            else:
   570	                df.to_excel(wr, sheet_name=sheet, index=False)
   571	
   572	            # Ticks in foglio dedicato
   573	            tname = safe_sheet_name(ycol + "_ticks")
   574	            (ticks if not ticks.empty else pd.DataFrame([{"note": "no ticks"}])).to_excel(
   575	                wr, sheet_name=tname, index=False
   576	            )
   577	    wb = wr.book
   578	    if "Sheet" in wb.sheetnames:
   579	        wb.remove(wb["Sheet"])
   580	
   581	    return out_path
   582	
   583	
   584	def main_cli():
   585	    parser = argparse.ArgumentParser(
   586	        description="Estrae i grafici da un report HTML e li salva in un Excel unico."
   587	    )
   588	    parser.add_argument(
   589	        "path",
   590	        nargs="?",
   591	        default=Path("."),
   592	        type=Path,
   593	        help="File HTML o directory contenente l'HTML (default: cartella corrente)",
   594	    )
   595	    parser.add_argument(
   596	        "-o",
   597	        "--output-dir",
   598	        default=Path("."),
   599	        type=Path,
   600	        help="Directory in cui salvare l'Excel (default: cartella corrente)",
   601	    )
   602	    args = parser.parse_args()
   603	
   604	    html_path = args.path
   605	    if html_path.is_dir():
   606	        html_files = sorted(html_path.glob("*.html"))
   607	        if html_files:
   608	            html_path = html_files[0]
   609	        else:
   610	            html_path = html_path / DEFAULT_HTML
   611	
   612	    out_path = process_html(html_path, args.output_dir)
   613	    logging.info("Salvato: %s", out_path)
   614	
   615	
   616	if __name__ == "__main__":
   617	    main_cli()
```

### Explanation
- **Lines 1-19:** File header describing purpose, usage instructions, and dependencies.
- **Lines 21-31:** Import statements. Uses `pathlib.Path` for file system paths, `argparse` for CLI parsing, `re` for regular expressions, `datetime` for time handling, `logging` for diagnostics, `base64` and `urllib.request` for embedded image handling, `numpy` and `pandas` for numerical and tabular data manipulation, and `BeautifulSoup` from `bs4` for HTML/SVG parsing.
- **Line 34:** Defines `DEFAULT_HTML` using `Path` to represent the default HTML report name.
- **Lines 36-37:** Configures basic logging and obtains a module-level logger.
- **Lines 42-52:** `parse_iso_utc` parses ISO-formatted UTC timestamps into timezone-aware `datetime` objects, returning `None` on failure.
- **Lines 55-73:** `parse_transforms` accumulates `scale()` and `translate()` SVG transform commands into combined scale and translation factors.
- **Lines 76-93:** `cumulative_transform` walks up the SVG DOM to accumulate transforms from parent nodes; `BeautifulSoup` nodes act as DOM elements.
- **Lines 95-96:** `apply_tr` applies a transformation to an (x, y) point.
- **Lines 99-123:** `parse_path_subpaths` tokenizes SVG path data into subpaths of points.
- **Lines 125-222:** Helper functions `extract_path_points`, `bounds_from_ticks`, and `guess_scales` to interpret SVG path elements.
- **Lines 224-260:** `extract_ticks` reads axis ticks, returning a `pandas.DataFrame` for mapping pixel positions to values.
- **Lines 262-293:** `extract_curve_for_header` locates charts under an HTML header, decoding `<object>` tags if necessary and returning curve data as `DataFrame`.
- **Lines 295-343:** `derive_orbit_filename` parses meta information from the HTML to build an output filename.
- **Lines 346-424:** `map_x_to_time` and `map_y_from_ticks` convert pixel coordinates to real values using tick information and produce timestamps.
- **Lines 426-493:** `safe_sheet_name` and `process_html` orchestrate reading the HTML report, iterating through charts, and writing results to Excel via a context-managed `pd.ExcelWriter` (demonstrating a context manager).
- **Lines 495-576:** Inside `process_html`, charts are processed, DataFrames reordered, timestamps converted, and tick data stored in additional sheets. Uses list comprehension when filtering headers and manipulating lists of targets.
- **Lines 578-581:** Removes the default empty sheet from the workbook and returns the output path.
- **Lines 584-617:** `main_cli` defines the command-line interface using `argparse`, resolves the HTML report path, invokes `process_html`, and logs the saved file. Includes the `if __name__ == "__main__"` guard.

Key data structures:
- `pandas.DataFrame` stores extracted chart data and tick mappings for easy export to Excel.
- `BeautifulSoup` DOM nodes represent parsed HTML/SVG elements, enabling traversal and attribute access.
- Lists and tuples manage collections of numeric points and transformation parameters.

Relevant Python concepts:
- Context managers (`with pd.ExcelWriter(...)`)
- List comprehensions (e.g., building argument lists and filtering headers)
- Regular expressions for parsing strings
- Exception handling for robust CLI processing
- Use of external libraries: `pandas`, `numpy`, `beautifulsoup4`, `openpyxl`.

## gui_app.py

### Code
```python
     1	from pathlib import Path
     2	from tkinter import Tk, Listbox, filedialog, StringVar, Text, PanedWindow
     3	from tkinter import ttk
     4	import logging
     5	
     6	from Extract_all_charts import process_html
     7	
     8	
     9	LOG_PATH = Path(__file__).resolve().with_name("gui_app.log")
    10	LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    11	handlers = [logging.StreamHandler()]
    12	try:
    13	    handlers.insert(0, logging.FileHandler(LOG_PATH))
    14	except OSError as exc:  # pragma: no cover - log fallback
    15	    print(f"Cannot write log file: {exc}")
    16	logging.basicConfig(
    17	    level=logging.INFO,
    18	    format="%(asctime)s - %(levelname)s - %(message)s",
    19	    handlers=handlers,
    20	    force=True,
    21	)
    22	
    23	
    24	class TextHandler(logging.Handler):
    25	    def __init__(self, widget):
    26	        super().__init__()
    27	        self.widget = widget
    28	
    29	    def emit(self, record):  # pragma: no cover - UI side effect
    30	        msg = self.format(record)
    31	        self.widget.configure(state="normal")
    32	        self.widget.insert("end", msg + "\n")
    33	        self.widget.configure(state="disabled")
    34	        self.widget.see("end")
    35	
    36	
    37	def main():
    38	    root = Tk()
    39	    root.title("MEOS Extract GUI")
    40	    root.geometry("600x300")
    41	    root.minsize(600, 300)
    42	    root.grid_rowconfigure(0, weight=1)
    43	    root.grid_columnconfigure(0, weight=1)
    44	
    45	    paned = PanedWindow(root, orient="vertical")
    46	    paned.grid(row=0, column=0, sticky="nsew")
    47	
    48	    main_frame = ttk.Frame(paned)
    49	    paned.add(main_frame, minsize=120)
    50	    main_frame.grid_columnconfigure(0, weight=1)
    51	    for r in range(3):
    52	        main_frame.grid_rowconfigure(r, weight=1, minsize=30)
    53	    main_frame.grid_rowconfigure(3, weight=3)
    54	
    55	    style = ttk.Style()
    56	    style.configure("Caption.TLabel", font=("Segoe UI", 10, "bold"))
    57	
    58	    output_var = StringVar()
    59	    txt_output = ttk.Entry(
    60	        main_frame,
    61	        textvariable=output_var,
    62	        state="readonly",
    63	        width=60,
    64	    )
    65	    lbl_output = ttk.Label(main_frame, text="Output folder:", style="Caption.TLabel")
    66	    lbl_output.grid(row=0, column=0, sticky="w", padx=5, pady=(0, 2))
    67	    txt_output.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
    68	
    69	    lbl_input = ttk.Label(main_frame, text="Input folders:", style="Caption.TLabel")
    70	    lbl_input.grid(row=2, column=0, sticky="w", padx=5, pady=(10, 2))
    71	    listbox = Listbox(main_frame, width=60, height=8, selectmode="extended")
    72	    listbox.grid(row=3, column=0, sticky="nsew", padx=(5, 0), pady=5)
    73	    list_scroll = ttk.Scrollbar(main_frame, orient="vertical", command=listbox.yview)
    74	    list_scroll.grid(row=3, column=1, sticky="ns", padx=(0, 5), pady=5)
    75	    listbox.configure(yscrollcommand=list_scroll.set)
    76	
    77	    lbl_count = ttk.Label(main_frame)
    78	    lbl_count.grid(row=4, column=0, sticky="ew", padx=5, pady=5)
    79	
    80	    btn_frame = ttk.Frame(main_frame)
    81	    btn_frame.grid(row=5, column=0, sticky="ew", padx=5, pady=5)
    82	
    83	    log_frame = ttk.Frame(paned)
    84	    paned.add(log_frame, minsize=80)
    85	    log_frame.grid_rowconfigure(0, weight=1)
    86	    log_frame.grid_columnconfigure(0, weight=1)
    87	    log_panel = Text(log_frame, state="disabled")
    88	    log_panel.grid(row=0, column=0, sticky="nsew", padx=(5, 0), pady=5)
    89	    log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=log_panel.yview)
    90	    log_scroll.grid(row=0, column=1, sticky="ns", padx=(0, 5), pady=5)
    91	    log_panel.configure(yscrollcommand=log_scroll.set)
    92	    text_handler = TextHandler(log_panel)
    93	    text_handler.setFormatter(logging.getLogger().handlers[0].formatter)
    94	    logging.getLogger().addHandler(text_handler)
    95	
    96	    def update_count():
    97	        n = listbox.size()
    98	        lbl_count.config(
    99	            text=f"{n} folder{'s' if n != 1 else ''} queued"
   100	        )
   101	
   102	    output_dir = {"path": None}
   103	
   104	    def add_folder():
   105	        folder = filedialog.askdirectory(title="Select folder")
   106	        if folder:
   107	            listbox.insert("end", folder)
   108	            update_count()
   109	
   110	    def remove_selected():
   111	        sel = listbox.curselection()
   112	        for idx in reversed(sel):
   113	            listbox.delete(idx)
   114	        update_count()
   115	
   116	    def select_output():
   117	        folder = filedialog.askdirectory(title="Select output directory")
   118	        if folder:
   119	            output_dir["path"] = Path(folder)
   120	            output_var.set(folder)
   121	
   122	    def run():
   123	        if output_dir["path"] is None:
   124	            logging.warning("Output directory not selected")
   125	            return
   126	        saved = []
   127	        for i in range(listbox.size()):
   128	            folder = Path(listbox.get(i))
   129	            html_files = sorted(folder.glob("*.html"))
   130	            if not html_files:
   131	                logging.warning("No HTML file in %s", folder)
   132	                continue
   133	            if len(html_files) > 1:
   134	                logging.warning(
   135	                    "Multiple HTML files in %s; using %s",
   136	                    folder,
   137	                    html_files[0].name,
   138	                )
   139	            report = html_files[0]
   140	            try:
   141	                out = process_html(report, output_dir["path"])
   142	                logging.info("Saved: %s", out)
   143	                saved.append(out)
   144	            except Exception:
   145	                logging.exception("Error processing %s", folder)
   146	                return
   147	        logging.info("Completed: created %d files", len(saved))
   148	
   149	    btn_add = ttk.Button(btn_frame, text="Add folder", command=add_folder)
   150	    btn_remove = ttk.Button(btn_frame, text="Remove selected", command=remove_selected)
   151	    btn_output = ttk.Button(
   152	        btn_frame, text="Output folder destination", command=select_output
   153	    )
   154	    btn_run = ttk.Button(btn_frame, text="Run", command=run)
   155	    btn_exit = ttk.Button(btn_frame, text="Exit", command=root.destroy)
   156	
   157	    btn_add.pack(side="left", padx=5, pady=5)
   158	    btn_remove.pack(side="left", padx=5, pady=5)
   159	    btn_output.pack(side="left", padx=5, pady=5)
   160	    btn_run.pack(side="left", padx=5, pady=5)
   161	    btn_exit.pack(side="left", padx=5, pady=5)
   162	
   163	    root.bind("<Escape>", lambda e: root.destroy())
   164	    update_count()
   165	
   166	    root.mainloop()
   167	
   168	
   169	if __name__ == "__main__":
   170	    main()
```

### Explanation
- **Lines 1-6:** Imports for filesystem paths (`Path`), Tkinter widgets, and logging, plus the `process_html` function from the extraction module.
- **Lines 9-21:** Logging configuration. Creates a log file if possible and falls back to console logging. Demonstrates file handling and conditional `try/except`.
- **Lines 24-35:** `TextHandler` subclass of `logging.Handler` that directs log messages to a `Text` widget. Uses UI-side effects; methods like `emit` modify widget state.
- **Lines 37-166:** `main` builds the GUI using Tkinter widgets (`Tk`, `PanedWindow`, `ttk.Frame`, `Listbox`, `Scrollbar`, `Button`, etc.). Widgets are stored as local variables; `StringVar` holds the output directory path. Nested functions (`update_count`, `add_folder`, `remove_selected`, `select_output`, `run`) manipulate widget state and call `process_html` for each selected folder. Uses list iteration, conditional checks, and logging.
- **Lines 169-170:** Standard entry point calling `main` when the script is executed directly.

Key data structures:
- Tkinter `Widget` objects (`Listbox`, `Text`, `Scrollbar`, etc.) form the GUI layout.
- `StringVar` provides a mutable string container linked to the output directory entry.
- Lists (`saved`) and dictionaries (`output_dir`) track runtime data.

Relevant Python concepts:
- Object-oriented subclassing (`TextHandler` extends `logging.Handler`).
- High-order functions and callbacks (buttons invoking local functions).
- Contextual resource management via the logging module.

## meos_extract.spec

### Code
```python
     1	# meos_extract.spec
     2	# Generate the executables with:  pyinstaller meos_extract.spec
     3	
     4	from pathlib import Path
     5	import sys
     6	block_cipher = None
     7	project_root = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
     8	
     9	# --- CLI analysis/executable ---
    10	a_cli = Analysis(
    11	    ['Extract_all_charts.py'],
    12	    pathex=[str(project_root)],
    13	    binaries=[],
    14	    datas=[],
    15	    hiddenimports=['bs4', 'pandas', 'numpy', 'openpyxl'],  # add others if needed
    16	    hookspath=[],
    17	    hooksconfig={},
    18	    runtime_hooks=[],
    19	    excludes=[],
    20	    cipher=block_cipher,
    21	)
    22	
    23	pyz_cli = PYZ(a_cli.pure, a_cli.zipped_data, cipher=block_cipher)
    24	
    25	exe_cli = EXE(
    26	    pyz_cli,
    27	    a_cli.scripts,
    28	    [],
    29	    exclude_binaries=True,
    30	    name='extract_cli',
    31	    debug=False,
    32	    bootloader_ignore_signals=False,
    33	    strip=False,
    34	    upx=True,
    35	    console=True,        # console app
    36	)
    37	
    38	# --- GUI analysis/executable ---
    39	a_gui = Analysis(
    40	    ['gui_app.py'],
    41	    pathex=[str(project_root)],
    42	    binaries=[],
    43	    datas=[],
    44	    hiddenimports=['bs4', 'pandas', 'numpy', 'openpyxl'],
    45	    hookspath=[],
    46	    hooksconfig={},
    47	    runtime_hooks=[],
    48	    excludes=[],
    49	    cipher=block_cipher,
    50	)
    51	
    52	pyz_gui = PYZ(a_gui.pure, a_gui.zipped_data, cipher=block_cipher)
    53	
    54	exe_gui = EXE(
    55	    pyz_gui,
    56	    a_gui.scripts,
    57	    [],
    58	    exclude_binaries=True,
    59	    name='extract_gui',
    60	    debug=False,
    61	    bootloader_ignore_signals=False,
    62	    strip=False,
    63	    upx=True,
    64	    console=False,       # windowed app
    65	)
    66	
    67	# --- Collect both executables into one dist folder ---
    68	coll = COLLECT(
    69	    exe_cli,
    70	    exe_gui,
    71	    a_cli.binaries + a_gui.binaries,
    72	    a_cli.zipfiles + a_gui.zipfiles,
    73	    a_cli.datas + a_gui.datas,
    74	    strip=False,
    75	    upx=True,
    76	    upx_exclude=[],
    77	    name='MEOS-Extract'
    78	)
```

### Explanation
- **Lines 1-3:** Header comments indicating the purpose and invocation for PyInstaller.
- **Lines 4-7:** Imports `Path` and `sys` to resolve project paths; defines `block_cipher` and computes `project_root` based on `__file__`.
- **Lines 9-21:** Defines an `Analysis` object for the CLI build, listing the main script, search paths, hidden imports, and build options.
- **Lines 23-36:** Creates a `PYZ` archive and an `EXE` configuration for the CLI executable.
- **Lines 38-65:** Mirrors the analysis and executable creation for the GUI version, toggling `console=False` to produce a windowed application.
- **Lines 67-78:** `COLLECT` bundles both executables and their resources into a single distribution directory named `MEOS-Extract`.

Key data structures:
- PyInstaller's `Analysis`, `PYZ`, `EXE`, and `COLLECT` objects encapsulate build configurations and artifacts.

Relevant Python concepts:
- Conditional expressions and module-level constants.
- Interaction with external build tool (`pyinstaller`).

