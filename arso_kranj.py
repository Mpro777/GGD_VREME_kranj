#!/usr/bin/env python3
"""Varčen zajem 48-urnih meritev ARSO za postajo Kranj.

Brez zunanjih Python knjižnic. Ob vsakem zagonu:
- prenese ARSO HTML,
- prebere največjo tabelo,
- shrani samo nove vrstice v SQLite,
- izdela/posodobi Excel datoteko ARSO_Kranj.xlsx,
- se zapre.
"""
from __future__ import annotations

import datetime as dt
import html
import os
import re
import sqlite3
import sys
import urllib.request
import zipfile
import unicodedata
from html.parser import HTMLParser
from pathlib import Path
from xml.sax.saxutils import escape

URL = "https://meteo.arso.gov.si/uploads/probase/www/observ/surface/text/sl/observationAms_KRANJ_history.html"
BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "arso_kranj.sqlite"
XLSX_DIR = BASE_DIR / "tabela_vremena"
XLSX_DIR.mkdir(parents=True, exist_ok=True)
XLSX_PATH = XLSX_DIR / "ARSO_Kranj.xlsx"
LOG_PATH = BASE_DIR / "arso_kranj.log"
TIMEOUT_SECONDS = 30

DATE_RE = re.compile(r"(?:\d{1,2}\.\d{1,2}\.\d{4}|\d{4}-\d{2}-\d{2}).*?\d{1,2}:\d{2}")
SPACE_RE = re.compile(r"\s+")


class TableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.tables: list[list[list[str]]] = []
        self._table_depth = 0
        self._current_table: list[list[str]] | None = None
        self._current_row: list[str] | None = None
        self._in_cell = False
        self._cell_parts: list[str] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        tag = tag.lower()
        if tag == "table":
            self._table_depth += 1
            if self._table_depth == 1:
                self._current_table = []
        elif self._table_depth == 1 and tag == "tr":
            self._current_row = []
        elif self._table_depth == 1 and tag in ("td", "th"):
            self._in_cell = True
            self._cell_parts = []
        elif self._in_cell and tag in ("br", "p", "div"):
            self._cell_parts.append(" ")

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if self._table_depth == 1 and tag in ("td", "th") and self._in_cell:
            text = SPACE_RE.sub(" ", "".join(self._cell_parts)).strip()
            if self._current_row is not None:
                self._current_row.append(text)
            self._in_cell = False
            self._cell_parts = []
        elif self._table_depth == 1 and tag == "tr":
            if self._current_table is not None and self._current_row and any(self._current_row):
                self._current_table.append(self._current_row)
            self._current_row = None
        elif tag == "table":
            if self._table_depth == 1 and self._current_table:
                self.tables.append(self._current_table)
                self._current_table = None
            self._table_depth = max(0, self._table_depth - 1)

    def handle_data(self, data: str) -> None:
        if self._in_cell:
            self._cell_parts.append(data)


def log(message: str) -> None:
    timestamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {message}"
    print(line)
    with LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def download_page() -> str:
    request = urllib.request.Request(
        URL,
        headers={"User-Agent": "Mozilla/5.0 ARSO-Kranj-archive/1.0"},
    )
    with urllib.request.urlopen(request, timeout=TIMEOUT_SECONDS) as response:
        raw = response.read()
        charset = response.headers.get_content_charset() or "utf-8"
    try:
        return raw.decode(charset)
    except UnicodeDecodeError:
        return raw.decode("utf-8", errors="replace")


def clean_cell(value: str) -> str:
    return SPACE_RE.sub(" ", html.unescape(value)).strip()


def select_weather_table(page: str) -> list[list[str]]:
    parser = TableParser()
    parser.feed(page)
    if not parser.tables:
        raise RuntimeError("Na strani ni bilo mogoče najti nobene tabele.")
    # Vremenska tabela ima daleč največ vrstic.
    return max(parser.tables, key=lambda t: (len(t), sum(len(r) for r in t)))


def normalize_table(table: list[list[str]]) -> tuple[list[str], list[list[str]]]:
    rows = [[clean_cell(c) for c in row] for row in table if any(clean_cell(c) for c in row)]
    data_start = next((i for i, row in enumerate(rows) if row and DATE_RE.search(row[0])), None)
    if data_start is None:
        raise RuntimeError("Ni bilo mogoče prepoznati vrstic z datumom in uro meritve.")

    data_rows = rows[data_start:]
    col_count = max(len(r) for r in data_rows)
    data_rows = [r + [""] * (col_count - len(r)) for r in data_rows]

    # Poskus sestave naslovov iz vrstic pred podatki.
    header_candidates = rows[:data_start]
    headers: list[str] = []
    if header_candidates:
        candidate = max(header_candidates, key=len)
        headers = candidate + [""] * (col_count - len(candidate))
        headers = headers[:col_count]

    fallback = [
        "Datum in čas", "Velja za", "UTC čas", "Postaja", "Nadmorska višina [m]",
        "Širina", "Dolžina", "Ocenjena oblačnost / pojavi", "Temperatura [°C]",
        "Vlažnost [%]", "Hitrost vetra [km/h]", "Smer vetra", "Smer vetra [°]",
        "Sunki vetra [km/h]", "Zračni tlak [hPa]", "Tlak nereduciran [hPa]",
        "Padavine [mm]", "Vsota padavin [mm]", "Sončno obsevanje [W/m²]",
        "Difuzno sončno obsevanje [W/m²]", "Višina snežne odeje [cm]", "Temperatura vode [°C]"
    ]
    out_headers: list[str] = []
    used: dict[str, int] = {}
    for i in range(col_count):
        name = headers[i] if i < len(headers) and headers[i] else (fallback[i] if i < len(fallback) else f"Podatek {i+1}")
        name = clean_cell(name)
        used[name] = used.get(name, 0) + 1
        if used[name] > 1:
            name = f"{name} ({used[name]})"
        out_headers.append(name)

    return out_headers, data_rows


def row_key(row: list[str]) -> str:
    # Celotna vrstica je ključ; tako se izognemo podvajanju tudi ob spremembi strukture vira.
    return "\x1f".join(row)


def init_db(conn: sqlite3.Connection) -> None:
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS measurements (
            row_key TEXT PRIMARY KEY,
            observed_at TEXT,
            row_data TEXT NOT NULL,
            inserted_at TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS metadata (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        )
    """)


def parse_observed_at(text: str) -> str:
    patterns = [
        r"(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{2})",
        r"(\d{4})-(\d{2})-(\d{2})\s+(\d{1,2}):(\d{2})",
    ]
    m = re.search(patterns[0], text)
    if m:
        d, mo, y, h, mi = map(int, m.groups())
        return f"{y:04d}-{mo:02d}-{d:02d} {h:02d}:{mi:02d}:00"
    m = re.search(patterns[1], text)
    if m:
        y, mo, d, h, mi = map(int, m.groups())
        return f"{y:04d}-{mo:02d}-{d:02d} {h:02d}:{mi:02d}:00"
    return text


def save_rows(headers: list[str], rows: list[list[str]]) -> int:
    now = dt.datetime.now().isoformat(timespec="seconds")
    inserted = 0
    import json
    with sqlite3.connect(DB_PATH) as conn:
        init_db(conn)
        for row in rows:
            key = row_key(row)
            cur = conn.execute(
                "INSERT OR IGNORE INTO measurements(row_key, observed_at, row_data, inserted_at) VALUES (?, ?, ?, ?)",
                (key, parse_observed_at(row[0]), json.dumps(row, ensure_ascii=False), now),
            )
            inserted += cur.rowcount
        conn.execute("INSERT OR REPLACE INTO metadata(key, value) VALUES ('headers', ?)", (json.dumps(headers, ensure_ascii=False),))
        conn.execute("INSERT OR REPLACE INTO metadata(key, value) VALUES ('source_url', ?)", (URL,))
        conn.commit()
    return inserted


def column_letter(n: int) -> str:
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def numeric_value(value: str):
    v = value.strip().replace("−", "-")
    if not v or DATE_RE.search(v):
        return None
    # Decimalna vejica ali pika; brez tisočic.
    if re.fullmatch(r"[-+]?\d+(?:[.,]\d+)?", v):
        try:
            return float(v.replace(",", "."))
        except ValueError:
            return None
    return None


def normalize_header_name(value: str) -> str:
    """Poenoti naslov stolpca za zanesljivo iskanje ne glede na šumnike in enote."""
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = value.lower().replace("/", " ")
    return SPACE_RE.sub(" ", value).strip()


def find_column(headers: list[str], required: tuple[str, ...], excluded: tuple[str, ...] = ()) -> int | None:
    normalized = [normalize_header_name(h) for h in headers]
    for i, header in enumerate(normalized):
        if all(word in header for word in required) and not any(word in header for word in excluded):
            return i
    return None


def safe_value(row: list[str], index: int | None, fallback_index: int | None = None) -> str:
    if index is not None and index < len(row):
        return row[index]
    if fallback_index is not None and fallback_index < len(row):
        return row[fallback_index]
    return ""


def split_date_time(value: str) -> tuple[str, str]:
    """Vrne datum DD.MM.LLLL in uro HH:MM iz ARSO zapisa."""
    patterns = [
        r"(\d{1,2})\.(\d{1,2})\.(\d{4}).*?(\d{1,2}):(\d{2})",
        r"(\d{4})-(\d{2})-(\d{2}).*?(\d{1,2}):(\d{2})",
    ]
    m = re.search(patterns[0], value)
    if m:
        d, mo, y, h, mi = map(int, m.groups())
        return f"{d:02d}.{mo:02d}.{y:04d}", f"{h:02d}:{mi:02d}"
    m = re.search(patterns[1], value)
    if m:
        y, mo, d, h, mi = map(int, m.groups())
        return f"{d:02d}.{mo:02d}.{y:04d}", f"{h:02d}:{mi:02d}"
    return value, ""


def prepare_excel_data(headers: list[str], rows: list[list[str]]) -> tuple[list[str], list[list[str]]]:
    """Pripravi zahtevane stolpce iz ARSO tabele za postajo Kranj.

    ARSO v HTML za isti podatek pogosto vsebuje dve zaporedni celici
    (vidna vrednost in pomožna/dostopnostna vrednost). Pred temperaturo sta
    tudi prazni celici za oblačnost/pojave. Zato uporabljamo preverjeno
    razporeditev glede na celico z imenom postaje, ne glede na sestavljeno glavo.
    """
    export_headers = [
        "Datum", "Ura", "Kraj", "Temperatura [°C]", "Vlažnost [%]",
        "Padavine [mm]", "Vsota padavin [mm]"
    ]
    export_rows: list[list[str]] = []

    def first_nonempty(*values: str) -> str:
        for value in values:
            if value is not None and str(value).strip() != "":
                return str(value).strip()
        return ""

    def at(row: list[str], idx: int) -> str:
        return row[idx].strip() if 0 <= idx < len(row) else ""

    for row in rows:
        date_value, time_value = split_date_time(row[0] if row else "")
        station_idx = next(
            (i for i, value in enumerate(row[:12]) if value.strip().casefold() == "kranj"),
            None,
        )

        if station_idx is None:
            continue

        place = "Kranj"

        # Dejanska ARSO struktura vsebuje podvojene vrednosti. Na podlagi
        # prikazane tabele sta temperatura in drugi podatki v drugem členu para.
        temperature = first_nonempty(at(row, station_idx + 6), at(row, station_idx + 7))
        humidity = first_nonempty(at(row, station_idx + 8), at(row, station_idx + 9))
        # Padavine in vsota padavin sta pozneje v vrstici, prav tako v parih.
        precipitation = first_nonempty(at(row, station_idx + 22), at(row, station_idx + 23))
        precipitation_sum = first_nonempty(at(row, station_idx + 24), at(row, station_idx + 25))

        export_rows.append([
            date_value,
            time_value,
            place,
            temperature,
            humidity,
            precipitation,
            precipitation_sum,
        ])

    return export_headers, export_rows

def write_xlsx(headers: list[str], rows: list[list[str]]) -> None:
    """Ustvari preprost, veljaven XLSX z uporabo samo standardne knjižnice."""
    max_cols = max([len(headers)] + [len(r) for r in rows])
    headers = headers + [f"Podatek {i+1}" for i in range(len(headers), max_cols)]
    rows = [r + [""] * (max_cols - len(r)) for r in rows]

    def inline_cell(ref: str, value: str, style: int = 0) -> str:
        return f'<c r="{ref}" t="inlineStr" s="{style}"><is><t xml:space="preserve">{escape(value)}</t></is></c>'

    sheet_rows = []
    header_cells = "".join(inline_cell(f"{column_letter(i+1)}1", h, 1) for i, h in enumerate(headers))
    sheet_rows.append(f'<row r="1" ht="24" customHeight="1">{header_cells}</row>')

    for r_idx, row in enumerate(rows, start=2):
        cells = []
        for c_idx, value in enumerate(row, start=1):
            ref = f"{column_letter(c_idx)}{r_idx}"
            num = numeric_value(value)
            if num is not None:
                cells.append(f'<c r="{ref}" s="2"><v>{num}</v></c>')
            else:
                cells.append(inline_cell(ref, value, 0))
        sheet_rows.append(f'<row r="{r_idx}">{"".join(cells)}</row>')

    last_col = column_letter(max_cols)
    last_row = len(rows) + 1
    widths = []
    for i, h in enumerate(headers, 1):
        width = 22 if i == 1 else min(28, max(11, len(h) + 2))
        widths.append(f'<col min="{i}" max="{i}" width="{width}" customWidth="1"/>')

    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>{"".join(widths)}</cols>
  <sheetData>{"".join(sheet_rows)}</sheetData>
  <autoFilter ref="A1:{last_col}{last_row}"/>
</worksheet>'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font><sz val="10"/><name val="Calibri"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>
  <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF2F6B3B"/><bgColor indexed="64"/></patternFill></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"><alignment horizontal="right"/></xf></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    files = {
        "[Content_Types].xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>''',
        "_rels/.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>''',
        "xl/workbook.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="ARSO Kranj" sheetId="1" r:id="rId1"/></sheets></workbook>''',
        "xl/_rels/workbook.xml.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>''',
        "xl/worksheets/sheet1.xml": sheet_xml,
        "xl/styles.xml": styles_xml,
        "docProps/app.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>ARSO Kranj zajem</Application></Properties>''',
        "docProps/core.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>ARSO Kranj vremenske meritve</dc:title><dc:creator>ARSO Kranj program</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">{dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "")}Z</dcterms:created></cp:coreProperties>''',
    }
    temp = XLSX_PATH.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(temp, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for name, content in files.items():
            zf.writestr(name, content.encode("utf-8"))
    os.replace(temp, XLSX_PATH)


def export_all() -> int:
    import json
    with sqlite3.connect(DB_PATH) as conn:
        init_db(conn)
        meta = conn.execute("SELECT value FROM metadata WHERE key='headers'").fetchone()
        headers = json.loads(meta[0]) if meta else ["Datum in čas"]
        db_rows = conn.execute("SELECT row_data FROM measurements ORDER BY observed_at ASC, inserted_at ASC").fetchall()
    rows = [json.loads(r[0]) for r in db_rows]
    excel_headers, excel_rows = prepare_excel_data(headers, rows)
    write_xlsx(excel_headers, excel_rows)
    return len(rows)


def main() -> int:
    try:
        page = download_page()
        table = select_weather_table(page)
        headers, rows = normalize_table(table)
        inserted = save_rows(headers, rows)
        total = export_all()
        log(f"Uspešno: prebranih {len(rows)} vrstic, novih {inserted}, skupaj v arhivu {total}. Excel: {XLSX_PATH.name}")
        return 0
    except Exception as exc:
        log(f"NAPAKA: {type(exc).__name__}: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
