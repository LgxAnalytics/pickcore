#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
PickCore — Picking List Converter  (FULL / automated)
=====================================================
Pelny pipeline: BC zapisuje PDF -> watcher lapie -> arkusz HTML (Twoj layout)
-> Edge headless render do PDF (ulotny) -> druk na wybrana drukarke -> kasacja PDF
-> kopia HTML do archiwum. Plus baza klientow (kod -> pelna nazwa + reminder).

ZALEZNOSCI (build na maszynie z Pythonem 3.12):
    pip install pdfplumber pywin32 pyinstaller
Edge headless = wbudowany w Windows 11, zero instalacji.

BUILD .EXE (v1.0 - ONEDIR, nie onefile):
    py -3.12 -m PyInstaller --onedir --windowed --name PickCore ^
       --collect-all pdfplumber --collect-all pdfminer ^
       pick_converter_full.py

    Deploy: skopiuj CALY folder dist\PickCore (exe + _internal) w docelowe miejsce,
    skrot na pulpicie do PickCore.exe. Archiwa i Logi powstaja obok exe jak dotad.
    DLACZEGO onedir: --onefile rozpakowywal ~100 MB do Temp\_MEIxxxxx przy kazdym
    starcie; AV trzymal uchwyty przy zamykaniu -> bootloader nie mogl skasowac
    katalogu -> dialog "Failed to remove temporary directory". Onedir nie tworzy
    _MEI wcale (start szybszy, zero dialogu). Po przenosinach przelacz ponownie
    "Start with Windows" w Settings (wpis w rejestrze wskazuje na stary exe).
"""
import os, sys, re, html, json, time, shutil, tempfile, threading, subprocess, webbrowser, glob
from pathlib import Path
from datetime import datetime
from collections import defaultdict
import getpass

import pdfplumber

try:
    import win32print                      # v3.0 fix: print_zpl_raw used win32print without importing it (NameError)
except Exception:
    win32print = None                      # environments without pywin32 (dev/test) - RAW printing unavailable

def resource_path(rel):
    """Path to a bundled resource, working inside a packed .exe through PyInstaller _MEIPASS."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)

def app_base_dir():
    """Directory holding the .exe (or the script) - archives and logs are created here."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def ensure_app_folders():
    """Creates the archive and log folders next to the .exe. Returns (pick_dir, pa_dir, log_dir)."""
    base = app_base_dir()
    pick_dir = os.path.join(base, "Archive_Picks")
    pa_dir   = os.path.join(base, "Archive_PutAways")
    log_dir  = os.path.join(base, "Logs")
    for d in (pick_dir, pa_dir, log_dir):
        try: os.makedirs(d, exist_ok=True)
        except Exception: pass
    return pick_dir, pa_dir, log_dir

# A silent disk write failure once cost a whole customer file: "except: pass"
# hid a NameError and nobody knew something had failed to load. The catch is that
# write_log_file and log_event ARE the logging mechanism - they cannot report their own
# errors through logln without looping. So they are collected here and the GUI pulls
# them into the Logs view in drain_io_fails().
IO_FAILS = []          # [(where, reason)] - queue to surface in Logs

def note_io_fail(where, err):
    """Records failed I/O without using logln. Safe to call from worker threads."""
    try:
        if len(IO_FAILS) < 200:            # upper bound: if the GUI stops draining, this must not grow without limit
            IO_FAILS.append((where, str(err)[:120]))
    except Exception:
        pass

def write_log_file(line):
    """Appends an entry to the persistent log file (next to the .exe)."""
    try:
        _, _, log_dir = ensure_app_folders()
        with open(os.path.join(log_dir, "pickcore_log.txt"), "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {line}\n")
    except Exception as e:
        note_io_fail("write_log_file", e)

def local_ip():
    """This station's IP on the warehouse network (no ipconfig): the UDP socket sends nothing,
       but the OS still picks the outbound interface - the same one the scanner will arrive on."""
    import socket as _s
    try:
        k = _s.socket(_s.AF_INET, _s.SOCK_DGRAM); k.settimeout(0.3)
        k.connect(("10.255.255.255", 1)); ip = k.getsockname()[0]; k.close()
        if ip and not ip.startswith("127."): return ip
    except Exception: pass
    try:
        return _s.gethostbyname(_s.gethostname())
    except Exception:
        return "127.0.0.1"

def events_path():
    _, _, log_dir = ensure_app_folders()
    return os.path.join(log_dir, "pickcore_events.jsonl")

def log_event(etype, **data):
    """Structured event for metrics and audit (JSONL, one row = one event).
       To zrodlo dashboardu 'ile zlapano bledow / uchroniono przed strata'.
       Every event carries the operator (Windows login) and the app version - 'who' and 'with what' for free."""
    try:
        rec = {"ts": datetime.now().isoformat(timespec="seconds"), "type": etype,
               "op": get_picker(), "app": APP_VERSION}
        rec.update(data)
        if etype == "PICK_TELEMETRY":
            rec.pop("op", None)   # PRIVACY (GDPR): telemetry measures the WAREHOUSE, not people - operator omitted
        with open(events_path(), "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except Exception as e:
        # This feeds the Impact dashboard. A silent failure means metrics simply vanish,
        # and nobody notices, because the chart still draws, just flat.
        note_io_fail(f"log_event({etype})", e)

def load_events():
    """Loads every event from the JSONL file."""
    out = []
    try:
        p = events_path()
        if not os.path.exists(p): return out
        with open(p, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line: continue
                try: out.append(json.loads(line))
                except Exception: pass
    except Exception: pass
    return out

# events counted as "error caught / loss prevented"
ERROR_EVENTS = ("WRONG_ITEM", "OVERPICK", "BULK_BLOCKED")

APP_VERSION = "1.0"
APP_TITLE = f"PickCore {APP_VERSION} — Cockpit"
HOME = Path.home()
CONFIG_PATH    = HOME / ".pickcore_converter.json"
CUSTOMERS_PATH = HOME / ".pickcore_customers.json"
SERIAL_SKUS_PATH = HOME / ".pickcore_serial_skus.json"
BULK_PROFILE_PATH = HOME / ".pickcore_bulk_profile.json"
KNOWN_SKUS_PATH  = HOME / ".pickcore_known_skus.json"
SKU_DESC_PATH    = HOME / ".pickcore_sku_desc.json"   # SKU -> description, learned from picks and put-aways

# ==================================================================== parser v2
COLS = {"qty":(0,90),"item":(90,240),"desc":(240,457),"bin":(457,535),
        "cust":(535,607),"order":(607,682),"due":(682,9999)}
FURNITURE = re.compile(r"(PICKING LIST|Warehouse Activity|Location Code|^No\.$|^Page$|"
                       r"OPERATOR|USER|^\d{1,2}:\d{2}|January|February|March|April|May|June|July|"
                       r"August|September|October|November|December)", re.I)

def _to_qty(s):
    """Ilosc z komorki QTY, tolerancyjnie na separatory tysiecy ('1,000' / '1.000' -> 1000).
       Returns an int, or None when the value is not a number."""
    s = (s or "").strip()
    if not re.fullmatch(r"[\d.,\s]+", s): return None
    digits = re.sub(r"[^\d]", "", s)
    return int(digits) if digits else None

def _col(x0):
    for n,(lo,hi) in COLS.items():
        if lo <= x0 < hi: return n
    return None

def _cluster(words, tol=5):
    words = sorted(words, key=lambda w:(w["top"], w["x0"]))
    out, cur, top = [], [], None
    for w in words:
        if top is None or abs(w["top"]-top) <= tol:
            cur.append(w); top = w["top"] if top is None else top
        else:
            out.append(cur); cur=[w]; top=w["top"]
    if cur: out.append(cur)
    return out

def parse_pick(path):
    rows, warnings, header_no = [], [], ""
    started = False
    with pdfplumber.open(path) as pdf:
        for pno, page in enumerate(pdf.pages, 1):
            words = page.extract_words(); full = page.extract_text() or ""
            if not header_no:
                m = re.search(r"No\.:?\s*(PI\d+)", full)
                if m: header_no = m.group(1)
            for line in _cluster(words):
                cells = defaultdict(list)
                raw = " ".join(w["text"] for w in sorted(line, key=lambda x:x["x0"]))
                for w in sorted(line, key=lambda x:x["x0"]):
                    c = _col(w["x0"])
                    if c: cells[c].append(w["text"])
                t = {k:" ".join(v).strip() for k,v in cells.items()}
                if t.get("qty") == "QTY": started=True; continue
                if not started: continue
                qv = _to_qty(t.get("qty",""))
                if qv is not None:
                    rows.append({"qty":qv,"item":t.get("item",""),"desc":t.get("desc",""),
                                 "bin":t.get("bin",""),"cust":t.get("cust",""),
                                 "order":t.get("order",""),"due":t.get("due",""),"page":pno})
                else:
                    if rows and t.get("desc") and set(cells.keys())<= {"desc"} and not FURNITURE.search(raw):
                        rows[-1]["desc"] += " " + t["desc"]
    for i,it in enumerate(rows,1):
        if not it["item"]: warnings.append(f"Line {i}: missing SKU")
        if not it["bin"]:  warnings.append(f"Line {i}: missing location ({it['item']})")
        if it["qty"] <= 0: warnings.append(f"Line {i}: QTY<=0 ({it['item']})")
    if not rows: warnings.append("No lines detected - may not be a Warehouse Pick PDF")
    return header_no, rows, warnings

# --- PUT-AWAY: column layout differs from picking ---
PA_COLS = {"item":(0,150),"desc":(150,370),"qty":(370,410),"bin":(410,475),
           "source":(475,545),"due":(545,599),"action":(599,710),"notes":(710,9999)}

def detect_type(full_text):
    if re.search(r"Type:\s*Put-away", full_text, re.I): return "putaway"
    if re.search(r"Type:\s*Pick", full_text, re.I):     return "pick"
    return "unknown"

def parse_putaway(path):
    rows, warnings, header_no, source = [], [], "", ""
    started = False
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            words = page.extract_words(); full = page.extract_text() or ""
            if not header_no:
                m = re.search(r"No\.:?\s*(PA\d+)", full)
                if m: header_no = m.group(1)
            for line in _cluster(words):
                cells = defaultdict(list); raw = " ".join(w["text"] for w in sorted(line,key=lambda x:x["x0"]))
                for w in sorted(line, key=lambda x:x["x0"]):
                    c = _col2(w["x0"], PA_COLS)
                    if c: cells[c].append(w["text"])
                t = {k:" ".join(v).strip() for k,v in cells.items()}
                if t.get("item")=="Item" or t.get("qty")=="QTY": started=True; continue
                if not started: continue
                qv = _to_qty(t.get("qty",""))
                if qv is not None:
                    rows.append({"item":t.get("item",""),"desc":t.get("desc",""),"qty":qv,
                                 "bin":t.get("bin",""),"source":t.get("source",""),"due":t.get("due","")})
                    if not source: source = t.get("source","")
                else:
                    if rows and t.get("desc") and set(cells.keys())<= {"desc"} and not FURNITURE.search(raw):
                        rows[-1]["desc"] += " " + t["desc"]
    for i,it in enumerate(rows,1):
        if not it["item"]: warnings.append(f"Line {i}: missing SKU")
        if it["qty"]<=0:   warnings.append(f"Line {i}: QTY<=0 ({it['item']})")
    if not rows: warnings.append("No lines detected - may not be a Put-away PDF")
    return header_no, source, rows, warnings

def _col2(x0, cols):
    for n,(lo,hi) in cols.items():
        if lo <= x0 < hi: return n
    return None

def parse_bc_lines_xlsx(path):
    """Zrodlo PRAWDY: eksport Lines z karty picka BC (Lines -> Open in Excel).
       Picking-list PDF agreguje linie per (item,bin) PONAD orderami i gubi przypisania AS
       (a real case: 2+6 units collapsed to "8" under a single assembly). XLSX carries the full structure.
       Only Action Type=Take is taken (picked from a bin); Place is a put-down (DISPATCH/ASSEMBLY)."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    it = ws.iter_rows(values_only=True)
    hdr = [str(h or "").strip() for h in next(it)]
    def col(name):
        for i, h in enumerate(hdr):
            if h.lower() == name.lower(): return i
        return None
    ix = {k: col(k) for k in ("Action Type","Item No.","Description","Bin Code",
                              "Quantity","Due Date","Source Document","Source No.","Destination No.")}
    if ix["Item No."] is None or ix["Source No."] is None or ix["Action Type"] is None:
        return "unknown", "", os.path.basename(path), [], ["not a BC Lines export"]
    rows, warns = [], []
    for r in it:
        if not r: continue
        g = lambda k: (r[ix[k]] if ix[k] is not None and ix[k] < len(r) else None)
        if str(g("Action Type") or "").strip().lower() != "take": continue
        item = str(g("Item No.") or "").strip()
        if not item: continue
        try: qty = int(float(g("Quantity") or 0))
        except Exception: qty = 0
        if qty <= 0: continue
        due = g("Due Date")
        due_s = due.strftime("%m/%d/%y") if hasattr(due, "strftime") else str(due or "").strip()
        rows.append({
            "qty": qty, "item": item,
            "desc": str(g("Description") or "").replace("\n", " ").strip(),
            "bin": str(g("Bin Code") or "").strip().upper(),
            "cust": str(g("Destination No.") or "").strip(),
            "order": str(g("Source No.") or "").strip(),
            "due": due_s, "page": 1,
            "src": str(g("Source Document") or "").strip(),
        })
    if not rows: return "unknown", "", os.path.basename(path), [], ["no Take lines"]
    hdr_no = primary_order(rows, "")          # document title = primary SO
    return "pick", hdr_no, os.path.basename(path), rows, warns

def parse_document(path):
    if str(path).lower().endswith((".xlsx", ".xlsm")):
        return parse_bc_lines_xlsx(path)
    """Dispatcher: wykrywa typ i routuje do wlasciwego parsera.
       Zwraca (doc_type, header_no, source/'', rows, warnings).
       The type is searched on ALL pages - the header is not always on page 1."""
    dt = "unknown"
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            d = detect_type(page.extract_text() or "")
            if d != "unknown":
                dt = d; break
    if dt == "putaway":
        hdr, src, rows, warns = parse_putaway(path)
        return "putaway", hdr, src, rows, warns
    elif dt == "pick":
        hdr, rows, warns = parse_pick(path)
        return "pick", hdr, "", rows, warns
    return "unknown", "", "", [], ["Unknown document type (neither Pick nor Put-away)"]

def natural_bin_key(b):
    special = "-" not in b
    return (1 if special else 0, [int(p) if p.isdigit() else p for p in re.split(r"(\d+)", b)])

def get_picker():
    try: raw = getpass.getuser()
    except Exception: raw = os.environ.get("USERNAME") or "UNKNOWN"
    n = raw.strip().replace(" ", ".").upper()
    return n.split("\\")[-1] if "\\" in n else (n or "UNKNOWN")

def _uniform(rows, key):
    vals = {r[key] for r in rows if r.get(key)}
    return "-" if not vals else (next(iter(vals)) if len(vals)==1 else "MULTIPLE")

def order_list(rows):
    """Distinct job numbers, in document order."""
    seen, out = set(), []
    for r in rows:
        o = r.get("order","")
        if o and o not in seen: seen.add(o); out.append(o)
    return out

def primary_order(rows, header_no):
    """Primary number (header plus file name): first sales order, else first job, else shipment.
       The sales order is preferred, because that is the key used to look the job up in the ERP."""
    orders = order_list(rows)
    so = [o for o in orders if o.upper().startswith("SO")]
    if so: return so[0]
    if orders: return orders[0]
    return header_no or "PICK"

def bin_sort_key(b):
    """Natural location order: 01-D11-A1 sorts before 01-D2-A1 lexically, but a walk route reverses it.
       Split into segments and compare numbers as numbers, the way they read on the printout."""
    out = []
    for seg in re.split(r"[-_/ ]+", (b or "").upper()):
        m = re.match(r"^([A-Z]*)(\d*)([A-Z]*)$", seg)
        if m:
            out.append((m.group(1), int(m.group(2)) if m.group(2) else -1, m.group(3)))
        else:
            out.append((seg, -1, ""))
    return out

def _norm_sku(s):
    """Canonical SKU normalisation: upper case, no -ASM suffix, no region markers
       ('MDR11SDGANQ1AN EU' na liscie = skan 'MDR11SDGANQ1AN'), bez spacji.
       Marker = krotkie (<=3) czysto-literowe tokeny po pierwszym (EU, UK, NA...)."""
    s = (s or "").strip().upper()
    if s.endswith("-ASM"): s = s[:-4]
    if " " in s:
        parts = s.split()
        if all(len(p) <= 3 and p.isalpha() for p in parts[1:]):
            s = parts[0]                                   # 'X EU' -> 'X'
        else:
            s = "".join(parts)                             # other spaces: join
    return s

def sku_match(code, sku):
    """Whether a SCANNED code matches a SKU on the picking list, tolerating barcode variants:
       - dodatkowy prefiks (np. skan 'PPMNN4491E' = lista 'PMNN4491E')
       - sufiks 'B' = bulk (np. skan 'PMNN4491EB' = lista 'PMNN4491E')
       - oba naraz; plus marker regionu na LISCIE ('MDR11SDGANQ1AN EU' = skan bazy).
       BEZPIECZNIK: tolerancje TYLKO dla SKU zawierajacych litery (warianty wystepuja na urzadzeniach OEM).
       Czysto-cyfrowe SKU (katalog dostawcy: 900000123, 900000456) = wylacznie DOKLADNE dopasowanie,
       because the numeric space is dense and a fuzzy match could accept a DIFFERENT real item."""
    c, s = _norm_sku(code), _norm_sku(sku)
    if not c or not s: return False
    if c == s: return True                                          # exact (any format)
    if not any(ch.isalpha() for ch in s): return False              # numeric SKU: no tolerance
    if c.endswith("B") and c[:-1] == s: return True                # suffix B (bulk)
    if len(c) == len(s) + 1 and c[1:] == s: return True            # extra leading letter
    if c.endswith("B") and len(c) == len(s) + 2 and c[1:-1] == s: return True   # prefix + B
    return False

QTY_SPEED_MAX_GAP_MS = 80     # scanner: inter-char usually 1-50ms (configurable); a human sustaining <80ms is not realistic
def qty_speed_ok(times_ms):
    """Scanner-versus-human heuristic on HARDWARE event timestamps (event.time, ms).
       Odporna na lag petli Tk (mierzy czas nacisniecia, nie obslugi). Zwraca (ok, n, max_ms, med_ms).
       Guard na wrap 32-bit licznika systemowego."""
    n = len(times_ms)
    if n < 2: return False, n, 0, 0
    gaps = []
    for a, b in zip(times_ms, times_ms[1:]):
        d = b - a
        if d < 0: d += 4294967296          # uint32 wrap (~49.7 days of uptime)
        gaps.append(d)
    gaps_s = sorted(gaps)
    med = gaps_s[len(gaps_s)//2]
    mx = gaps_s[-1]
    return mx <= QTY_SPEED_MAX_GAP_MS, n, int(mx), int(med)

def parse_scanned_qty(code):
    """Quantity from a scanner QTY code. Formats: 'Q'+number (Q30), 'qty'+number (qty20) or a bare number (30).
       UWAGA anti-cheat: gdy dopuszczamy samą liczbę, format przestaje chronic - caly ciezar bierze
       speed lock (a scanner fires instantly, a hand-typed '30' leaves gaps above 50ms)."""
    c = (code or "").strip()
    m = re.match(r"^Q(?:TY)?\s*[:#x*\-]?\s*(\d+)$", c, re.I)
    if m: return int(m.group(1))
    if c.isdigit() and len(c) <= 4: return int(c)   # bare number (max 4 digits - box qty; longer means EAN, not qty)
    return None

# Example prefix families. Every catalogue has its own; swap these for yours.
RADIO_PREFIXES = ("UNA", "UNB", "UNC", "TRX", "TRM", "HRT", "HRP")
def is_radio(sku):
    """A main unit is recognised by a known model prefix (a serialised item).
       The list covers the common families; anything else can be flagged manually at the station."""
    return (sku or "").strip().upper().startswith(RADIO_PREFIXES)

# OEM accessory catalogue number families (chargers/batteries/cables/audio/antennas...)
# - finite and distinctive; device SERIALS NEVER start this way (they start with digits).
ACCESSORY_PREFIXES = ("ACB","ACC","ACD","ACK","ACL","AUD","ANT",
                      "BAT","CBL","CHG","MNT","PSU","RMT")
def looks_like_sku(code):
    """Structural recognition of an OEM part BY PREFIX (device or accessory).
       Closes the registry cold-start gap: a wrong item from a catalogue family is an error
       even BEFORE the registry has seen it (for example a near-miss catalogue number one digit apart)."""
    c = _norm_sku(code)
    return c.startswith(RADIO_PREFIXES) or c.startswith(ACCESSORY_PREFIXES)

SERIAL_MIN, SERIAL_MAX = 6, 18
def valid_serial(code, pick_skus=None):
    """Whether a code looks like a valid serial number rather than a SKU, a qty or noise.
       Returns (True,'') or (False, reason). Catches a malformed code scanned in place of a serial.
       pick_skus = the SKUs on the pick, used to detect a SKU scanned in place of a serial."""
    c = (code or "").strip(); pick_skus = pick_skus or []
    if len(c) < SERIAL_MIN: return False, f"too short ({len(c)} chars)"
    if len(c) > SERIAL_MAX: return False, f"too long ({len(c)} chars)"
    if not re.fullmatch(r"[A-Za-z0-9\-]+", c): return False, "invalid characters"
    if not any(ch.isalpha() for ch in c): return False, "no letters (looks like a number/qty)"
    if any(sku_match(c, s) for s in pick_skus): return False, "looks like a SKU, not a serial"
    if looks_like_sku(c): return False, "looks like a SKU, not a serial"   # OEM part after the prefix
    if parse_scanned_qty(c) is not None: return False, "looks like a QTY code"
    return True, ""

def bc_serial_line(serial):
    """An Item Tracking Lines row for a SINGLE unit:
       Serial No <TAB> Availability,Serial=Yes <TAB> Lot No(empty) <TAB> Availability,Lot=Yes <TAB> Qty(Base)=1 <TAB> Qty to Handle=1.
       Format confirmed against a real ERP export."""
    return f"{serial}\tYes\t\tYes\t1\t1"

def bc_serial_export(serials, pack=1):
    """Blok do wklejenia w BC Item Tracking Lines.
       Jednostka = OPAKOWANIE: jeden wiersz = jedno pudelko z 'pack' radiami.
       pack=1: serial<TAB>Yes<TAB><TAB>Yes<TAB>1<TAB>1 (potwierdzone).
       pack=2/4/6: 'pack' seriali w JEDNEJ komorce oddzielonych SPACJA (nie osobne kolumny) + kolumny formatu.
       for a dual pack: 'SN1 SN2<TAB>Yes<TAB><TAB>Yes<TAB>1<TAB>1', the next box on the row below."""
    pack = max(1, pack)
    rows = []
    for i in range(0, len(serials), pack):
        group = serials[i:i+pack]
        rows.append(" ".join(group) + "\tYes\t\tYes\t1\t1")   # SPACE between serials in a row
    return "\n".join(rows)

# ==================================================================== customers
def load_customers():
    if CUSTOMERS_PATH.exists():
        try: return json.loads(CUSTOMERS_PATH.read_text(encoding="utf-8"))
        except Exception: return {}
    return {}

def save_customers(d):
    CUSTOMERS_PATH.write_text(json.dumps(d, indent=2, ensure_ascii=False), encoding="utf-8")

def load_serial_skus():
    """SKUs manually flagged as serialised, normalised and stored as a set."""
    if SERIAL_SKUS_PATH.exists():
        try: return set(json.loads(SERIAL_SKUS_PATH.read_text(encoding="utf-8")))
        except Exception: return set()
    return set()
def save_serial_skus(s):
    SERIAL_SKUS_PATH.write_text(json.dumps(sorted(s), ensure_ascii=False), encoding="utf-8")

# ---------- KNOWN SKU REGISTRY (learns passively from every pick/put-away) ----------
def load_known_skus():
    """Catalogue of SKUs gathered from every processed document. The basis of triage:
       a code NOT on the pick but IN the registry is a genuine wrong item; outside it, noise."""
    if KNOWN_SKUS_PATH.exists():
        try: return set(json.loads(KNOWN_SKUS_PATH.read_text(encoding="utf-8")))
        except Exception: return set()
    return set()
SKU_DESC = {}

def load_sku_desc():
    """Description registry: {SKU: description}, fed by every processed document."""
    global SKU_DESC
    if SKU_DESC: return SKU_DESC
    try:
        if SKU_DESC_PATH.exists():
            SKU_DESC = json.loads(SKU_DESC_PATH.read_text(encoding="utf-8"))
    except Exception:
        SKU_DESC = {}
    return SKU_DESC

def sku_desc(sku):
    """Item description when known. An empty string rather than an error: it is a convenience, not a requirement."""
    return (load_sku_desc().get(_norm_sku(sku or "")) or "").strip()

def learn_sku_desc(rows):
    """Zapamietuje opisy z dokumentu. Nowszy dokument aktualizuje starszy opis."""
    try:
        reg = dict(load_sku_desc()); changed = False
        for r in rows:
            sk = _norm_sku(r.get("item","") or "")
            d = (r.get("desc") or "").strip()
            if sk and d and reg.get(sk) != d:
                reg[sk] = d[:60]; changed = True
        if changed:
            SKU_DESC_PATH.write_text(json.dumps(reg, ensure_ascii=False, indent=0), encoding="utf-8")
            globals()["SKU_DESC"] = reg
    except Exception: pass

def import_item_descriptions(path):
    """Item catalogue import (columns: No.;Description). Fills the description registry IMMEDIATELY,
       instead of waiting for the SKU to appear in a pick. Separator auto-detected (; , TAB |)."""
    import csv as _csv
    try:
        if not path or not os.path.exists(path):
            return 0, "file not found"
        raw = open(path, "r", encoding="utf-8-sig", errors="ignore").read()
        lines = raw.splitlines()
        if not lines: return 0, "file is empty"
        head = lines[0]
        delim = max([";", ",", "\t", "|"], key=lambda d: head.count(d))
        if head.count(delim) == 0: delim = ";"
        rd = list(_csv.DictReader(lines, delimiter=delim))
        if not rd: return 0, "no rows"
        cols = [c for c in (rd[0].keys() or []) if c]
        def pick(*frags):
            for c in cols:
                lc = c.lower().replace(" ", "").replace("_", "").replace(".", "")
                if any(f in lc for f in frags): return c
            return None
        c_no  = pick("no", "item", "nr", "sku")
        c_dsc = pick("description", "opis", "name")
        if not (c_no and c_dsc):
            return 0, f"No./Description columns not found (seen: {', '.join(cols[:6])})"
        reg = dict(load_sku_desc()); n = 0
        for r in rd:
            sk = _norm_sku((r.get(c_no) or "").strip())
            d  = (r.get(c_dsc) or "").strip()
            if sk and d and reg.get(sk) != d:
                reg[sk] = d[:60]; n += 1
        if n:
            SKU_DESC_PATH.write_text(json.dumps(reg, ensure_ascii=False, indent=0), encoding="utf-8")
            globals()["SKU_DESC"] = reg
        return n, f"{len(rd)} rows read, {n} descriptions added/updated, {len(reg)} in registry"
    except Exception as e:
        return 0, f"read error: {str(e)[:80]}"

def learn_skus(rows):
    """Adds SKUs from a processed document to the registry, only when they are new."""
    learn_sku_desc(rows)
    try:
        known = load_known_skus()
        fresh = {_norm_sku(r.get("item","")) for r in rows if r.get("item")}
        fresh.discard("")
        if fresh - known:
            KNOWN_SKUS_PATH.write_text(json.dumps(sorted(known | fresh), ensure_ascii=False), encoding="utf-8")
    except Exception: pass

# ---------- BULK PROFILE: registry of "which items move in bulk and in what size" ----------
# Deterministic and auditable (like known-SKU) - learns from picks AND serial sessions.
def load_bulk_profile():
    if BULK_PROFILE_PATH.exists():
        try: return json.loads(BULK_PROFILE_PATH.read_text(encoding="utf-8"))
        except Exception: return {}
    return {}
def learn_bulk(sku, qty):
    """Remembers the bulk qty for a SKU, keeping the last 20 observations."""
    try:
        if qty < 2: return
        p = load_bulk_profile(); k = _norm_sku(sku)
        p.setdefault(k, []).append(int(qty)); p[k] = p[k][-20:]
        BULK_PROFILE_PATH.write_text(json.dumps(p, ensure_ascii=False), encoding="utf-8")
    except Exception: pass
def bulk_suggest(sku):
    """(sugeruj?, dominanta_qty) - sugestia od >=2 obserwacji; dominanta = najczestsze qty."""
    p = load_bulk_profile().get(_norm_sku(sku), [])
    if len(p) < 2: return False, 0
    from collections import Counter as _C
    return True, _C(p).most_common(1)[0][0]

def classify_scan(code, pick_skus=None, known=None):
    """TRIAGE of a scan that missed - separates a REAL error from noise:
       'qty'     - kod ilosci (Q30) poza bulkiem            -> szum
       'sku'     - SKU z rejestru znanych, nie na tym picku -> PRAWDZIWY ZLY ITEM
       'numeric' - same cyfry NIEZNANE rejestrowi (EAN itp.)-> szum
       'serial'  - wyglada na numer seryjny                 -> szum (za wczesnie zeskanowany)
       'unknown' - nierozpoznany                            -> nie liczony jako strata
       KOLEJNOSC MA ZNACZENIE: rejestr sprawdzany PRZED testem cyfrowym, bo katalog zawiera
       purely numeric supplier SKUs count too - those are real wrong items, not EAN codes."""
    c = (code or "").strip()
    if not c: return "unknown"
    known = load_known_skus() if known is None else known
    if any(sku_match(c, s) for s in known): return "sku"     # REGISTRY STRICTLY FIRST:
    # (parse_scanned_qty now catches bare digits, so numeric supplier SKUs must be errors, not qty noise)
    if parse_scanned_qty(c) is not None: return "qty"
    if looks_like_sku(c): return "sku"     # STRUCTURE: an OEM part outside the registry is a REAL error too
    if c.isdigit(): return "numeric"
    okv, _ = valid_serial(c, pick_skus or [])
    if okv: return "serial"
    return "unknown"

# ============================ WAREHOUSE ANALYTICS (slotting intelligence) ============================
def bin_zone(b):
    """Strefa z formatu binu: 'A-12-3'->'A', '05-01-02'->'05', 'CROSSDOCKING'->'CROSSDOCKING'."""
    b = (b or "").strip().upper()
    if not b: return "?"
    return b.split("-")[0] if "-" in b else b

def export_star_schema(out_dir):
    """Eksport modelu gwiazdy (star schema) z telemetrii do CSV - gotowe pod Power BI / DuckDB / SQL.
       Fakt: fact_pick_lines (ziarno = jedna linia picka). Wymiary: dim_item, dim_zone, dim_date, dim_pick.
       To jest 'jezyk rozmowy kwalifikacyjnej' + warstwa ladowania do BI jednym ruchem.
       Returns (paths[], fact_count), or ([], 0) when there is no data."""
    import csv
    from datetime import datetime as _dt
    tele = [e for e in load_events() if e.get("type") == "PICK_TELEMETRY" and e.get("lines")]
    if not tele: return [], 0
    facts = []       # grain: one pick line
    dim_pick = {}    # pick_id -> header attributes
    for pid, ev in enumerate(tele, 1):
        ts = ev.get("ts", ""); day = ts[:10]
        pick_id = f"P{pid:05d}"
        dim_pick[pick_id] = {
            "pick_id": pick_id, "doc": ev.get("doc",""), "order": ev.get("pick",""),
            "ts": ts, "date": day, "n_lines": ev.get("n_lines", len(ev.get("lines",[]))),
            "units": ev.get("units", 0), "cycle_s": ev.get("cycle_s") or "",
            "validation_s": ev.get("validation_s") or "", "queued": int(bool(ev.get("queued"))),
        }
        for li, ln in enumerate(ev.get("lines", []), 1):
            sku = _norm_sku(ln.get("sku","")); binv = ln.get("bin",""); zone = bin_zone(binv)
            facts.append({
                "fact_id": f"{pick_id}L{li:02d}", "pick_id": pick_id, "date": day,
                "sku": sku, "zone": zone, "bin": binv,
                "qty": int(ln.get("qty", 1)), "seq": li, "t_offset_s": ln.get("t", 0),
            })
    # item dimension: velocity aggregates (the ABC foundation)
    it = {}
    for f in facts:
        d = it.setdefault(f["sku"], {"sku": f["sku"], "picks": 0, "units": 0, "zones": {}})
        d["picks"] += 1; d["units"] += f["qty"]; d["zones"][f["zone"]] = d["zones"].get(f["zone"],0)+1
    ranked = sorted(it.values(), key=lambda d: -d["units"])
    cum = 0; tot = sum(d["units"] for d in ranked) or 1
    for d in ranked:                                  # ABC classification by cumulative share (Pareto)
        cum += d["units"]; share = cum / tot
        d["abc"] = "A" if share <= 0.8 else ("B" if share <= 0.95 else "C")
        d["home_zone"] = max(d["zones"], key=d["zones"].get) if d["zones"] else "?"
    dim_item = [{"sku": d["sku"], "home_zone": d["home_zone"], "picks": d["picks"],
                 "units": d["units"], "abc_class": d["abc"]} for d in ranked]
    # zone dimension
    zt = {}
    for f in facts:
        z = zt.setdefault(f["zone"], {"zone": f["zone"], "lines": 0, "units": 0})
        z["lines"] += 1; z["units"] += f["qty"]
    dim_zone = sorted(zt.values(), key=lambda z: -z["lines"])
    # date dimension
    dts = {}
    for f in facts:
        try: dd = _dt.strptime(f["date"], "%Y-%m-%d")
        except Exception: continue
        dts[f["date"]] = {"date": f["date"], "year": dd.year, "month": dd.month,
                          "day": dd.day, "weekday": dd.strftime("%A"), "week": int(dd.strftime("%V"))}
    dim_date = sorted(dts.values(), key=lambda d: d["date"])

    os.makedirs(out_dir, exist_ok=True)
    def _w(name, rows, cols):
        p = os.path.join(out_dir, name)
        with open(p, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=cols); w.writeheader()
            for r in rows: w.writerow(r)
        return p
    paths = [
        _w("fact_pick_lines.csv", facts, ["fact_id","pick_id","date","sku","zone","bin","qty","seq","t_offset_s"]),
        _w("dim_item.csv", dim_item, ["sku","home_zone","picks","units","abc_class"]),
        _w("dim_zone.csv", dim_zone, ["zone","lines","units"]),
        _w("dim_date.csv", dim_date, ["date","year","month","day","weekday","week"]),
        _w("dim_pick.csv", list(dim_pick.values()),
           ["pick_id","doc","order","ts","date","n_lines","units","cycle_s","validation_s","queued"]),
    ]
    # README with the model plus ready SQL (evidence of a designed warehouse schema)
    readme = (
        "PICKCORE — STAR SCHEMA EXPORT (warehouse slotting analytics)\n"
        "="*60 + "\n\n"
        "GRAIN: fact_pick_lines = one picked line on one pick.\n\n"
        "FACT\n  fact_pick_lines (fact_id, pick_id FK, date FK, sku FK, zone FK, bin, qty, seq, t_offset_s)\n\n"
        "DIMENSIONS\n"
        "  dim_item (sku PK, home_zone, picks, units, abc_class)   -- ABC via cumulative Pareto (A<=80%, B<=95%)\n"
        "  dim_zone (zone PK, lines, units)                        -- heat / capacity view\n"
        "  dim_date (date PK, year, month, day, weekday, week)     -- time intelligence\n"
        "  dim_pick (pick_id PK, doc, order, ts, date, n_lines, units, cycle_s, validation_s, queued)\n\n"
        "PRIVACY: no operator field anywhere — process data only (GDPR/AVG safe).\n\n"
        "LOAD (DuckDB example):\n"
        "  CREATE TABLE fact AS SELECT * FROM read_csv_auto('fact_pick_lines.csv');\n"
        "  CREATE TABLE dim_item AS SELECT * FROM read_csv_auto('dim_item.csv');\n\n"
        "SAMPLE QUERY — ABC-A units sitting in slow zones (slotting candidates):\n"
        "  SELECT i.sku, i.units, i.home_zone\n"
        "  FROM dim_item i\n"
        "  WHERE i.abc_class='A'\n"
        "  ORDER BY i.units DESC;\n\n"
        "POWER BI: load all 5 CSVs, connect fact_pick_lines[sku|zone|date|pick_id]\n"
        "to the matching dim PKs (star), build DAX measures on fact.\n"
    )
    rp = os.path.join(out_dir, "README_schema.txt")
    Path(rp).write_text(readme, encoding="utf-8"); paths.append(rp)
    return paths, len(facts)

# ---------------- PICKMAP HEAT (integration with pickmap_template.html) ----------------
# rack-view template embedded in the exe as zlib+base64; an external file takes priority
PICKMAP_TPL_B64 = (
"eNrVPe1y28iR//0Us/DeFRkTNAF+maSkhBIJaHPr2BU7SWV1qjNIDkXEFMAAoGSF0XPd/3uy6+7Bx0yTopS9TVXOLkloTHdP"
"T09Pd88MMDj5bvLh4vOfP07FKrtdn706wT9iHUQ3p5aMLLwhgwX8uZVZIOarIElldmpts6X9zipuR8GtPLXuQnm/iZPMEvM4"
"ymQEaPfhIludLuRdOJc2AQ0RRmEWBms7nQdreeogkyzM1vJsHCbpQyo24fyruA02IpO3m3WQyZO3qvzVyXe2/UqI8Q+///Tn"
"T+LjDxf/Id6PP4rP0/cffxx/ngpRu5GRTIBkIWYPogD+CzkCw+bmof4K6H8A4W6gIIwjMf74g1jGiVB1f/gkaqmUYhZnWXwr"
"TtJ5Em6ys/oQyYT4CHzeA59VeLNaw09WQ85pHYroQpyKK8tt2VOnbXuu1RDNZvNaAPeddnco2lTwSCyP/LsNEmCZruMsFfdh"
"thLZSoomaSfNHtZSBNFCBOKv2YOYBYsbaYoogwyaXFsEWdCINxkIiZcgoSlLX8kyEomMFjJJgWHXTjO5EcjhORGpC8UyAV2t"
"47nS6DzeRiBxDStt3gbfxDzYpCQ7YddHzzFdSui+MCOV2uswwgrkX7cymj+IjUzKmoz2ztcySGp1nU8ib+M7CS1ar5VNvaU2"
"wZ9UruWchCUdG4xuZFabxwtZ1xll2yRSbZh8eC+A+hasm+wm0Bu+MPtgAQKYIgFysFCM5O1MLhbQ0k+fP4AdpxFoaRVnorYE"
"kxdVL9f3unVisn2eZYC9mhsEcVvE8y22oBksFtM7uPgxhC6HwVKz8qEyVCoCE5bi9AxNpP5st4UJKBv0gBYr5mtgBNSghCwI"
"12h3oKgGWmsj179cPOKwAgGSh5woSJIQmKCl2oB+av30k/37jx/tHz9ZTfFRJvYymEtxGyZJnITRzRA4oY1a6g6Yc5ZsJcoQ"
"CMIMozRcSPHj+M8f/vAZTDIR21SSrrI4Xs+CBByanH+dxd/qTeDzCcjn0NfBGj3YMrwhe0NOQ2HF0frBErW1vJNrYZ+hWcX3"
"6GiCh7TeQCkWcpOtPkBzEqgUUKEEEVXTcFQoiwFzDGGAEKOc8m9xJKF7dfZ4ayQ2gC0zcR8kchWj7LM4BqewXMfIKRXtVkPM"
"ZHCbEuBAK2wb/CRZzhlwHiZxnO2o62x7djN87XW9ztQbAbQJIrmGG/QPb4TR1+Fr13W77hjB2y300PB1b9pv9XujnAUOx+Hr"
"yWDSu7gY5bA7fD32xueDaYGENgSMz6ft8XSUwznlxeDd2Dkvb1KV3YuO2zoviFMUypl0vYmDaEvQG0jRmr4bd70CBxR7Pxz0"
"Nt8QA9Vrr4btlgJxFA63oX0bR3G6ga5rWBcBeJ5FGIj3cM9qXMRRGq+DtPFeRuu4USIid3TKv9qBPdhp+De0sFmcgFu04Q4W"
"zeLFww58xk0YDVujWTD/epOQgHfgfFDB9dE8XsdJDkPzKne3hJg4dDqbb2+dZqcrINpAfLO3YcP6JG9iKf7wg9X4fQyRJ26M"
"E4iPjTSIUlBHEi6Lqq9oYPwtjm9PrdS63uWa6LvQ9FIRLlRxiGBdEThuT6dov0MKIMFIL5PdJk5D9GjDNINOehhl8QZa+zdo"
"zkJ+G3b3202mpDv2UmsYRofO5psAjYcLodDRFOqjDTgf1LADHSccEEGjX4QpRP6H4XItv43wl32fgFPCX6MbuCAaF5sdQByO"
"7BBUmQ7n4Mlkgm1vzhIIjjtUObiM23D9kEtKnn1E96GD5dDpAg8YYUBnoxGgQM1WT95WXMRM8bmXGPGHvVZLKwOaaKd3OY2a"
"+kjZiL2Wy2z4br8KxQJ9ULoz2oqNe3ewWUwNFQNwcpttdpU9bORpJr9l1ztlaW1Qkd7oQt/Q9WIAVVA2NnRp3LwyO+5gj7n1"
"Ud6tCQymbYqMdFN47U3g//neANAknW3BHiJNvmoQVMZuCOpgH+9X+1J5n5duNN8mKdzZxGFpPLqwwxXkEMkuF+FI29bBTK53"
"mmVB+0YHLGOvu7uHu/uAXGsJGe1i32DQqRy0mEoY97AwupkG2yymauarcFNWEkaoS3sG0fhrbjLUspUaDXRt9k/7/9Q/oGxw"
"OjA7oPYMbRQ8lzKhGrvKu5GUlAzv9t0R3C2NVW91GYjqFQuIN3nvDl0mJRRpiKBpvapEbiCjAjslhkFi32D7QfG1TnchbxoZ"
"+AcY7gkmii0BUjdeTztTZ9LDawFjvE4u9zYIo13pCnHoF5q/gSyCutdVOvhm5+p337VKpQxbouw3zFN2T/nmIx3Bx7VjeIRS"
"tg4Ox3JM4qiALOTe/jY0BLAxhhwwUe6rZwFoFyrn3j3v6yJytPLeLlmLlVsFYM3A0SscctMVYTOKM3nIV+vDpNndq7BJqduO"
"D5XRM5wKefbDD/V7E7uXj7Kyy1ujf9QRl/1VdMxwFcJcIHoqWCu7luvC+PPRddQ+jobzUoBb0JKy1DwpgoxDVbe+W5uWccBn"
"/WULScfywc5XEIrbFdNOR2stRRLoarEX7Q70juFtxt7A65Z8ntWB+yIlqMQgeDBbmbvK3KdQyqUw8XqHKENn9DP1UqZqiYRp"
"Y3gnq35QUdY9oheKf4dipBF6uDM/ainxNiNLjpdLmJqR+y6bOoScO7PBl64XptG1KpQ82uoyOVNv2q8wljB3Te27MA1na7nL"
"63vadSPN0TCxFx9IK4dcSclKZdUwk70egk+pleCp5VjX9eEwWILadqXaii4LsiypFbj1quuCGUgOJkp5Nqh1pLRSehByKAN9"
"jJOKiwhcJdKgg/1w3DHC8SEd5HkQ6QGng/1pr2qvCo8wHVoFC3AqMKGWGM/wf2VaprZpyrYLovCW1kXApIN0JZymmwoJXl+0"
"EfU3X+XDMgluJU5moXzX+reG02r927HKtJB6cPGs+zR55R5Iuvpj6f+adzE4Yt3iwN4m027VHlw52T1pEFj6/84g9ObmM/6q"
"/+FflfUc0c5rbPpTOem7p+dmUKwCzy+ZTB2KP52e1vj/W3gjwaGDfm78yvvG7R0MXu4/N3iV0qPq1TzlADvdwjut1v6suJvP"
"iudxEnEnrVi8IIxS8rOMYxwQT+W0mGRiiiVoblpYc4svEjyfgBnDoklBRFVNa7W7Y/nZb24lrhhtEoiBauSqxRE+a4/iSCpn"
"RGtDulKWy+UIFDr7GmY28VFzETtYoJEM5bdgno2eKlA885w+kQF6aVzLHAY4Is3JzevBYPBIi1cnb/OVv5O3+a4NSiW0VaDb"
"YkdHJrg+eLII78QcHHB6atGKhnV2MjtTGywnb2dn1SbLCS51nBVbMuKu1eyI//lvWqZMhdPCa5d+t1t221Gg27NbfbvVA7GQ"
"+OQt1MZrJXVaZ9TcE1rKEOHi1FqGIIugRQ0LVzUsAfXO5Speg+SnlgfF5co7LjPfNEW1rSHSDTgYWtoFTsE6lUUFalJPNcyy"
"yKZazpAZNJaKDuPhSr51hovwAi+PI6NbtM4u1Xr7C/Bp68I6u8A/DJMG65mmlnyhO1dMsXhtnYn3VCDAJCElOnmrCPeqQxuw"
"AQEY4I7aqfXpFheYoXb7aRmJKIxKmh9hlCLJG53kYN+q4FDoXtkA/i7KMcKAwSnrAMvK1D5AfuMYFW2QlKS0W1KjTTBxilsM"
"9ZewgKyl5FDsSLyEDpxoSTdP4jS1gzBdS3BPtjN4CQMcwSWHKBYJOA21HbBHTR1QxdmSqtK4Guk4ml+d4GKCMhMlIsJ4Xzk9"
"Itrg+C72OSTuNsXgee5w4S0EDzqkQU/tScXE9kWyBY6RgNiLhrIeiXixEEWmkqpdFnABAn0T8I+XkOPNV4oDjMs7GWnYuPES"
"bwhWFE0xJtVNRJjCjehmLW28v6C9zWI7iPbdaJsDa4jWDyLIcNtFOH3Cwx0Y4Tq2OwB+Qu+SYguwEoE6Se1YStp4SuJ7Ib/N"
"5SajfoAafGxjQ9yvQmgIyIWhPIy28Tal2kCaRC5jUN5XKTep8B2o9lMGRUFCsgzF2Ml3aNCaG2Li2BO3ITzH9tq43eM7tt9u"
"iEvHvmw3xe+x96nOYa4CtWNEW51oFA0BNuJR+xvqjvitrgZolEOSOX3cxLoA15fSxiao0TLMyxKr4E4iu83qIcVlttKHNk/e"
"bioD+ci3nWrKz7dbb9tObiS0owsJfMWr0HK+SQVmk93HeYSYPdD+AoXivBHj3PrqMGYRqdzMyrfGRDCDzisLndxYUiDE1p5z"
"80LjKIyLenBMxooh2c5iG/+qffRzIgTyKEtIidgUXMsqGwBDrSnOt+uv+CAEbiFDb29x7xSSkk6+KZeCa5ebwiRUn0Xb2xl0"
"W6hslvYDK63UHGiLskuwvyirj3Lu4xZYw7g1AP7UsJZLFbapgtxAUrVhnaK4Y6dljx03R+7b560BjZWHfCvwwp7CyK2h8u9x"
"+xNIIEADNlTQ6jdAtkTKqkgxag3q0ORWuxx11BMXDbjXVdeTpvhRyet0iGTs9PR60fbo/jTvOdQB9XBDzLcJJu9gs8rNQ+Ag"
"KwlgEm+a3qd4maHdiRUUK0uT+zvEJ5ipnRX7xCdvCWyKC+gZ5ETg/tMaV/g8Rj3Hps3YcrM9FXGyT5k/RLGjTUW1h/to0BtP"
"S+SPQ9T2HmDAvPPgIwxoArT1joal6jY34MuW7W9+AyW6tjAqxVbb3JQzplfhdVOh5iwK16F8HZkSDNDDhOhciprRX+kP1SiK"
"A4/WlPWk8TbBHfPeu45IsxidD+ggbYie0xbbKAQ9vBKlCuD+APx4mVjS8y7KJE7eFoHrVfEkzitwxmlWbOjjwyz4/JE1hAsc"
"93Bh5U8RTbdJvJHivPnHptUQFuUuWPyn0qcVzxlhMVhYCsIgAiS2eIfCCDkcvMkCCCKUjxpheZXrWo9QRpqE+1c7K6TyiQ2u"
"CKnQlykChIKC/wQhjHwI5KilyEWM/M9tqzXrk1Mj1Aa6d9VoM3YSfo6tAmcZL/KQSdGS6pQ3+CQICXs1aEAogR9wR04XfvrX"
"DXHlwh0X7rhwxwXv4Q6u4balfDSRWWPg5CAutcOlKw+u2nTll1eXdIXU5QMblJ038qcrsBsnyNLBevbrfQTdFhqd2ujEj6h0"
"qqu0wGU6nRZaokjClbofOSiUTvb09g6aD7HLgSDvdOCnR3qDOy7cceGO24Ofd7+43gx1HDew6fMGNj1mYE/owhOFZv8lDElT"
"iPecfXgvsA/v59jHVOT6/VezEe8ZG/GetxHvZ9iI/y9rI/5zNuK/wEb8n2Mj3r+qjfjP2Ih/zEbyqQwzEVHDjR9qd11XjTav"
"yikOTn+KwnISVNyoJkHl/Od5I4OfwcHYclynTqlJp9SkuvotXR2Pab+lmOaQkesx7CdM5HVlO6ayx0fN7yecFkFin6sDJ7Ll"
"lErXs5qKavkWqHYDM5xwvl0HCWae1Xx/T4FgeGB3YHZkm09o6ahFtVtjGmb2DOZTemPbrQLCRx+L2G+1HUqhXqiEfKaqUp7x"
"3lAsbsyquVzLxSxtXz9U2FGTujf5jCefyBTmiXZ5qqZ10GxzNveU5g6qrFMaVueFysPRc1x5F7nyGhjsy0v/n6ZS8mNgwI5r"
"+DttBlroHldqXG3W2VC3WqKYoe7pTnnBg6pTnuyivJoyP4c05oO4qBwYWEPhPr5I1WQJ/7Cuf3kFK4t0OrbzMnNNzen33qR7"
"X8sqxjxtoBfl1fRlpspix88Z4TymHFZWGVsOjO+23RocD8EbnadaDTKV3mrnayLaQsye+iB+QPiA6DF40RjfN8q2ehfEGuDf"
"Z2zzfC9VedoqDeWev9wUz3+W+1TrRKanRFTAxIWwJ5T3vHe8OKI590UaAyN5QdDRxvLkWQUetc7zF1pnV1ffU0tuxr3JQeN7"
"0citDLBMR7DFV+gz1WS7e/0iTf6iEegXUCrFH1xyGbwg/sBI5vGn1X86/jw5qH9e9Ok/EXyuH0f5gpZ6Y4cWtDAnhEE2dpCo"
"IRTcUbBTwD1W/s4sp5+83G3Zk9bA9oryAnYZ3DZhn8GXOuw4VX057HGY4fuuCV+6mnwIG/htxq9tyosww/dZuc/KL3V9IGzg"
"d1l7uqz+Lqu/y+rvsvq6qj4dNtrbVe0t4T6rvw8/ZvnENcs9BvsMvmT8DP26rP8AnnDYNWGPw4yfz+h9vX2uY+rbZf3ttln9"
"bVY/sweX2YPL7AFgn+FfcpjRm/J0zfEDsCkPsweX2QPAvmvS+6z8ktGb9bP+dll/I8zwfQab7e2z+rg9DFh7B6w/Bkwe5k9c"
"5j8A9hn+JcPX5Znq/iuHJ6zc0/RJcNss9xm9z+h9hn/J8Lk8lwb+gMln6odg14Q9Vu4xfj4r9xm9UX/hzzXY47DLYEbvs3Kj"
"/cyfI6z7H4Jds9zjcNuEfcbPZ+Vm+1xWv2vqF2CzPpfV57L6XFafy+prG/aO8ITDronvMdisr83022b1m/6GYKO8w9rfYfJ0"
"mDwdJk+HydNh9XdYf3eYPF1Wf5fV12X21mX677L6u0wfXSZPl8nTY/X3WP/3mDw9Jk+PydNj8vSYPD0mT4/po8f6p8/k6zP5"
"+kyePuufPpOvz+TpM3n6TD+m/566zB9QjDdhj8OM3mflPi9vmzCv39APyyemLJ+YsnwCYY+Ve7yc8fcZvs/wfYZ/yfAvGb4p"
"P/M/rmuOP4A9Vm7Ky/yRy/yRy/yR22b1tZm+mP9h+c+U5T9Tlv9MWf5DMMPn8lwyfFNfzD+5zB8B7LFyn5Wb9TF/g/kVg01+"
"zD+4zB+4vT18sz1s/Lps/LJ8a8ryrSnLp6buO0bP8hWX5Rsuyzdclk+wfGvK8q0py6cI1vh7LH9CeMJh14Q9Vu7xcsbfZ/g+"
"w+fyXDJ8fb7jsXzKY/mUx/Ipj+VTXpE/afx8hm/KM2DymPmnx/Itj+VbCBvysPzLY/mXx/IvhH2G77PyS1ZuyuMweUz/6rH8"
"zGP5mMfyMY/Nxz2Wn3ksH/NYPkawlo8j7LFyUx8uk8dl9bt79Zv6MPMzj83/PZaPeSwf81i+5bF8y2P5lsfyLYQ9hu+5Jr7H"
"+PuM/pLRX/Jyg77L5Osy/XeZfF3Gv8v49xh+j7Wnx/TXY/L3WP/0jfzZY/mQx9YrPJYfeSwfQthn5T6j19cTPJYfeSw/8lh+"
"5LH8iGCG77Fyj5e3Tdhn+D7DN/RV5EduBU8YbNbvsPrZ+GX5jcfyG89l45XlKx7LVzyWrxDM8A39u2x8snzGY/mMx/IZj+Uz"
"HstfPJa/eCx/8Vi+4rF8BGGPw20TvjT032X9wcYXy0c8tv7juV0mHxtvLD9B2LSPHmtPn+mTja8iX3Er2GwfywdYvuKxfAVh"
"Uz4zPvssv/BZfuEX+YVbwR4r1+XzWT7hF/mERu8z/EuGr9u7z9Zr/BaXf8DkNddffBbvEZ4w2OOwa8Kcn8/KL1m5ng/5jukf"
"fBbffRa/Ef6tAbtMfpfRu0Z/+yxe+yxe+2w9xWfx2WfxmWBGb/SH02bytZl8bVZfh+F3GH6HtafD5O/s8fN5eduELw39d1h7"
"ukyeLpOny+rrMfwew+8x+XusP8z1DJ/FZ5/FZ4QvWfklLzf49Zm99Zl8fVZfn9kbG18OG1+OmW/7LB77LB77LB77LP76LP76"
"LN76LN76bL3CZ+sVPluv8Nl6hc/irc/irc/WC3wWb322XuCzeOsX8dWtYM7flJeNHxZffRYvfRYffZeNHxYvfTZ/91m8RNhn"
"5T6j9xk+r/+S4ZvtY+OLxWOEfVZu8u8xfDa+XDZ+XDZ+XDZeXDZeWDz23T6rj40Xtl7gs3jss3jss3iMsM/KfV7eNmFen9ke"
"Nj5dLf612X5zAbusvM3KO3p5m9G3Fb0Ot3V8zb8XsMtgo74Oq6/L6LuMvsvq6zL6HqPvMfoeo+8p+lKePqPvM/o+o++z+geM"
"fsDoBya9vp9PsGvSO66pb4fp12HtcXqGPOctk995y+R3zvr/nPX3edHfjgZz/I4Odxh9h9F3GH3e/+0C7jJ5u4ze7P9z1v/n"
"+vMTBewyuM3wDfo+q7/P6h+w8sF+edk+h8ZTOT4LuM3gjg5r86kCNvA7jF+H4Xf28Q3+XVbeZeU9xr/H+PcYfY/R9xl9n9H3"
"GX2f0Q8q/1vALivX6WH8XDgMdjV6B/dYzPKpa8I+K/eNcpfxdxV/Hdb7G+Apw58y/GnbhH2Gz+v3NXwcr7o8BLsM1uTB8Wni"
"t/fLOzrcYfgdxr/D6DuKvizvGv1PsMvgNoMN+h6rv8fq76n6ddiQv8/o+4y+b/QPwa5J77Ny36AfMP4Dxn/A5Buw+gaGPRDM"
"8H2Gz+tX9vA4epU/YbYJFuJURHia86cM3x+tRfUm3PyUBUlWw4a1rHrxOBq+yflhCfi1ZQNfYG6I9d26IdI6kL8S4sv3u9qy"
"mT/6J/7+d7F7rF8BxjVeq4JH+/vdskmP8D1+jwcc1oBNHe8CHtxJH7+AaMttpM7hVqePe8FcYo3qMbnxel3Ho1eUSOoeivTd"
"d8v81da6+O70tMLGs2HwwcFU3uBR7Mtm8UxfE1/iTbHpaTNdh1BJnQ4MCpeilnPaFVT45wB+M8G3kFOkrC5HdF6wko9exD8V"
"V9fImJgs42QazFe1GqgvRM2JHdUI14jc3GzTVS3aQitH2o1ms5kiYxCw5K2eQvwcb6hVCtoXbFTip3KO7ykUp3zPExlkcqpO"
"La9Z+avkVl1JOm/S2RS/w1f5T9Wrq9aI7uPr1anM6L1gqjlcaELRyzd4ly7ICCoJ1KvvWGo8DcnRgvX6XOlN6yx8J9toDZ2A"
"r7U8kYstNL0WNMQVno5wnZulEIF4U/BsLsN1JpPaDBX/Hcp4pewzv2yG0Xy9Xci0NqvvnW1eVvGtIWbUdd+Ac43acjW7Fr/+"
"tYjqDdGin0KNYRTJ5PLz+x9B1i/6MSjlWZLW2cnKPcOBQc+6Pp68BfCLeHPgMLMvxrEh9HQrEeLVY34IyEso6Z11JC3UspbR"
"TbZ6VPb6P/8tYCiifh+Vmo2Ter5ovY1HUx6xKUBX9oR4pkHhHassovOJ6PDLz/kb3xfxensbQQdTW77g+Y4gFA0IHIbUf/AL"
"BjqOFfFrYeG5XpYYCks7YNKqN/8Sh1HNElb9UQmO7+fXlPRX5L4iauI1HlJSjqi6OtwpH2fB7CVtFIhoNhL4WyO6jYcVXajD"
"v6AA7isKanywwfeuLvAExBrg5sw0OWcoGx0oX5wWhw6jaj24qT1Gcq1Ea1j5wWoWOLfi0BKpXFTBSdk+Oup///diUNyxsWCc"
"Uwcy3UEzqjqwAnWADVRyp+wYFRDF5RtmVnVG156wd7poOVohYOlH9yuEm1Z5xJvCiy5lIou3sYoxSR1c4JFeKR4AmgMuTZyc"
"KjK4fvNmr6EY9ABTxT4MQ7My7I0YarrOZVQHIIGYWLOl4aXrJp7RhKrJcUZ4r3CpoKu8KrpdqJFuGDy4NYEbSg2M/U840KcX"
"rIaoqZijuQhAJ6tF3GYW39ysJYaDtWV8lKO0fzzjLMjmK+Jfi+S9uNimWXyr4L1PRRi+aKe++zDErz4MsVn06Yfhm5qmBLjx"
"97+jDy1OXxoaAqKdBGGU5iI+PtY1MR/r+qmWD4aRpdVZ9IVt7Rki0ORIj/mx/0cGljrxzlL1Hxmueaal2y9zXZAEieIQPnJh"
"cKPQ/ob1tkmpzp0AEsylCtexJ/OmXnzFACOSXoKoVKi+ZoLlo1ePWgoGImfBTQMCBzSnCONPu0PAhbFs+nogHRX8F+h5eIJX"
"0xO6YKNzv5FZzvr84YdFjY6wInnxdBE9tFrWiCWFkPod5aReuqg36bg0uWCduMRO1E9AqVOVuu6eyk7rpEHF5nI6/kymht8A"
"wmM025MemM5rbzLpD8Z01R/3Oud4NXV7HcfDK5iJTVsX1vVI4/JfP/zuPxQX9WUMxKuuOmN3jG8QWflRrdrVtZ5R44E0P9Kh"
"YbVtKhea4mfxt2Pq0k4bK5Pk75AFJk900cQTFsFGgBHvmbz39cR4A/MLyowhsyVqPHYphFhTv26mMcw9MI9TSVZw1boWtpjB"
"nzoF/9rVDHO8uCFW4XWV5n156iw59cmcU0s7hvH7HeoUAsSjeWhidb88VA3mJjGNuVUIIw6uh/jrjbBsC36vwjLvIr3w1udC"
"0UFA6OqqJE2pIM9PLN1o8oOG8L0f9EAyG+ZR6Kzqn79uZfLwiTxknNS+XFUfx/meziV6tK6/0Fdk+FepdJ9EZ5RCNeMkAWcZ"
"pvS3+HrVr9UJSKTxOdZ9NW84oO2h+DD7C1Rb9pjCz1MgicMvxVf2gW9rL5W5Kp0+JVwkQBly8xAOhNkqTKtvLo20XOW7CEws"
"r+HNm0MZTaQFCwiDKiJhXhLpIUadCFj5+mLaR4zrJEoMKel9kES1L3mHDCEDzTEeqUdqoCZIvQWdHowfGMKzrUBfyhQeG+rj"
"KdUXtwR+/Qq79TFvM7Uz/0jV6EDHME0jk2IIzBvijmwf+0W8getiblO7KovuxFk+Fcm1R4zzbL9eDEutYvwy16koP9IFY/t9"
"kK3wGqegilrVX9VRN0SfbcGbYqC6w0JkAvmVA9aEAyuvGUazk/fWMK8ABgFDaKgSGq+12h3eqou30EfAkq5/JThJ3RSF3BPE"
"SkhRoPdq/5iJ3v3TDLTIEk5zXdVKDdw1UGFVQrNnynSIqFas5k2VXwOeuf/iOOTe8mIMJQaKOS7uzJQaqchDY3PBI4NNXN2h"
"fgo05fyxELxy2ZYE3DVg1RuVBSVXDt26ruvjjoek0T9xNOZDbWd+78zwpRDBa1Z1qnujPPC8oZ1j3qhO+YYMoljSifTMWu87"
"deBk7okalLg2LGIEf1/QpxhCeUda5YRqAZJn0uzGXME5W03FamWp0If6Ft1OmXMVKei9U1xuKjzYxMDLXZqBhWfmqsFQBJhy"
"2tQE53Vbq8PU4g+QPSUXQbEs9dyAKoeT0f9W2f9lZw8ho1Z0RZ6JzSxGHGhunsTr9Q9RFv8xlPe13UyugrsQwr2V3sZxtsKE"
"HD9lM7TUwdwWfo6SviBQ3Cg0eahXVUfixBcmwIChvnLwJzzCerQ/hAvsV9U3BEWUd8jj6NU96DG+b1YpQH4FSdyTuVl5FnH9"
"2Xmf+RFCkOIo09y+X8q0spTnGKvDjV/OOR+3z7Glg5uPcAUtm3zJap9kmbO7C9ZbebTuJ+v9Kh+gNyOqWVZrvbIJBZRTWlMy"
"r7opEQwZPE05y6tW677Pz2QOtHsVRDf4Iruaq5Qr+T99+PD+E80mUvBAt/CzxnnCk1UUhz0/32OIWXOOaqs8bvqF3GxiV05h"
"6J4xc9HnLXiyeTUTDKGojEiuFopaDaWDJn297sOyNiu9J/KvQ3q+UHHYLAB+RHcVXlPCXkxgR3jIen4Y5clbFILOYKZv5/4v"
"8Tk7nA==")
def _embedded_template():
    import zlib as _z, base64 as _b
    return _z.decompress(_b.b64decode(PICKMAP_TPL_B64)).decode("utf-8")

PICKMAP_LOC = re.compile(r"^\d{2}-[A-Z]\d{2}-[A-Z]\d$")   # ZZ-RPP-LS grammar (matches the template data-loc)

def pickmap_heat_data():
    """Aggregates telemetry into {exact_location: VISIT count} - exactly what
       czego oczekuje PickMap.heatmap() (czestotliwosc WIZYT).
       Codes outside the grammar (CROSSDOCKING/WARRANTY/...) become unmapped and are listed in the banner."""
    from collections import Counter
    heat, unmapped = Counter(), Counter()
    for ev in load_telemetry():
        seen = set()                                   # VISITS: one pick times one location counts as one, regardless of scan mode and qty
        for ln in ev.get("lines") or []:
            b = (ln.get("bin") or "").strip().upper()
            if b: seen.add(b)
        for b in seen:
            if PICKMAP_LOC.match(b): heat[b] += 1
            else: unmapped[b] += 1
    return heat, unmapped

# ================= KIT LABELS (natywny port Label Selector HG v11.1) =================
# Recovered from v11: hardcoded dimensions/DPI, DESC_MAP, ZPL generator with a truncation flag,
# copies validation, audit CSV, RAW printing. Fixes: (1) kit source is order-level grouping
# rather than a brittle adjacency parser, (2) logging into Logs\, (3) printer taken from cfg.
LABEL_W_MM, LABEL_H_MM, LABEL_DPI = 40, 30, 203
_DPMM = 8                                   # 203dpi is about 8 dots/mm (GK420d)
LABEL_PW, LABEL_LL = LABEL_W_MM*_DPMM, LABEL_H_MM*_DPMM      # 320 x 240
FONT_HEADER, FONT_ITEM = 26, 20
HEADER_Y, ITEM_ROW_H, LABEL_BOTTOM_PAD = 56, 26, 8

DESC_MAP = {"CHARGER":("CHGR","CHARGER","DESKTOP"),
            "BATTERY":("BATT","BATTERY","LI-ION","IMPRES"),
            "ANTENNA":("ANT","ANTENNA","WHIP","STUBBY"),
            "BELT CLIP":("CLIP","BELT"),
            "DUST COVER":("COVER","DUST"),
            "MICROPHONE":("MIC","MICROPHONE","REMOTE"),
            "POWER SUPPLY":("POWER","SUPPLY","ADAPTER")}
def standardize_description(desc):
    """Canonical component name for the label, from a keyword map."""
    u = (desc or "").upper()
    for label, kws in DESC_MAP.items():
        if any(k in u for k in kws): return label
    t = (desc or "").strip()
    return (t[:15] + "…") if len(t) > 16 else t

def extract_label_kits(rows):
    """Label kits built from ALREADY parsed pick lines, grouped by order like the sheet.
       -> [{'kit_id': radio_sku, 'order': AS..., 'items': [{'sku','desc'},...]}]"""
    from collections import OrderedDict as _OD
    g=_OD()
    for it in rows: g.setdefault((it.get("order") or "").strip(), []).append(it)
    kits=[]
    for o, its in g.items():
        rads=[r for r in its if is_radio(r.get("item",""))]
        if not (o.upper().startswith("AS") and rads and len(its)>1): continue
        main=rads[0]
        kits.append({"kit_id": _norm_sku(main["item"]), "order": o,
                     "items": [{"sku": _norm_sku(r["item"]),
                                "desc": standardize_description(r.get("desc",""))}
                               for r in its if r is not main]})
    return kits

def generate_kit_zpl(kit_id, items, copies):
    """(zpl, truncated, n_fit) - header, a line, then 'SKU  - DESC' rows; ^PQ=copies."""
    y = HEADER_Y
    body, fit = [], 0
    for it in items:
        if y + ITEM_ROW_H > LABEL_LL - LABEL_BOTTOM_PAD: break
        body.append(f"^FO20,{y}^A0N,{FONT_ITEM},{FONT_ITEM}^FD{it['sku']}^FS")
        body.append(f"^FO135,{y}^A0N,{FONT_ITEM},{FONT_ITEM}^FD- {it['desc']}^FS")
        y += ITEM_ROW_H; fit += 1
    zpl = ("^XA" f"^PW{LABEL_PW}" f"^LL{LABEL_LL}" "^CI28"
           f"^FO20,20^A0N,{FONT_HEADER},{FONT_HEADER}^FD{kit_id}^FS"
           f"^FO20,48^GB{LABEL_PW-40},2,2^FS"
           + "".join(body) + f"^PQ{int(copies)}^XZ")
    return zpl, fit < len(items), fit

def validate_copies(value):
    try:
        c = int(str(value).strip())
        if c < 1: raise ValueError
        return c
    except (ValueError, TypeError):
        return None

def print_zpl_raw(zpl, printer):
    """Druk RAW ZPL do kolejki Zebry (port _print_zpl z v11)."""
    h = win32print.OpenPrinter(printer)
    try:
        win32print.StartDocPrinter(h, 1, ("PickCore Kit Label", None, "RAW"))
        win32print.StartPagePrinter(h)
        win32print.WritePrinter(h, zpl.encode("utf-8"))
        win32print.EndPagePrinter(h); win32print.EndDocPrinter(h)
    finally:
        win32print.ClosePrinter(h)

def append_print_log(kit_id, copies, printer, status):
    """Print audit trail written to Logs/print_log.csv (an earlier version wrote to CWD)."""
    try:
        import csv as _csv
        _,_,ld = ensure_app_folders()
        p = os.path.join(ld, "print_log.csv"); new = not os.path.exists(p)
        with open(p, "a", newline="", encoding="utf-8-sig") as f:
            w=_csv.writer(f)
            if new: w.writerow(["Timestamp","Kit ID","Copies","Printer","Status"])
            w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), kit_id, copies, printer, status])
    except Exception: pass

# ---------------- WINDOWS AUTOSTART (HKCU\\...\\Run - no admin rights needed) ----------------
_RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
_RUN_VAL = "PickCore"
def _autostart_target():
    """Launch path: the frozen exe, or the script when running from source."""
    if getattr(sys, "frozen", False): return sys.executable
    return os.path.abspath(sys.argv[0])
def autostart_get():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _RUN_KEY, 0, winreg.KEY_READ) as k:
            winreg.QueryValueEx(k, _RUN_VAL)
        return True
    except Exception:
        return False
def autostart_set(enable):
    """Enables or disables start with Windows. Returns (ok, msg)."""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
            if enable:
                winreg.SetValueEx(k, _RUN_VAL, 0, winreg.REG_SZ, f'"{_autostart_target()}"')
            else:
                try: winreg.DeleteValue(k, _RUN_VAL)
                except FileNotFoundError: pass
        return True, ""
    except Exception as e:
        return False, str(e)

# ========== SHIPMENT LABELS (port Label Selector HG v10.2 -> widok cockpit, v3.0) ==========
SHP_DESC_MAP = {
    "CHARGER":      ["CHGR", "CHARGER", "SUC", "UNIT", "DESKTOP"],
    "BATTERY":      ["BATT", "BATTERY", "LI-ION", "IMPRES"],
    "ANTENNA":      ["ANT", "ANTENNA", "WHIP", "STUBBY"],
    "BELT CLIP":    ["CLIP", "BELT"],
    "DUST COVER":   ["COVER", "DUST"],
    "MICROPHONE":   ["MIC", "MICROPHONE", "REMOTE"],
    "POWER SUPPLY": ["PSU", "POWER", "SUPPLY", "ADAPTER"],
}

def shp_standardize(raw_desc):
    d = (raw_desc or "").upper()
    for std, keys in SHP_DESC_MAP.items():
        if any(k in d for k in keys): return std
    return (raw_desc or "")[:15].strip() + ".."

def parse_shipment_pdf(path):
    """Shipment PDF -> {kit_id: [{sku,desc},...]} (main units plus *-ASM lines).
       Model prefixes come from RADIO_PREFIXES, so a different catalogue needs no code change.
       Guards against empty pages, keeps no globals, deduplicates per kit."""
    with pdfplumber.open(path) as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)
    kits, cur_id, cur_items, seen = {}, None, [], set()
    for line in text.splitlines():
        line = line.strip()
        _hit = next((p for p in RADIO_PREFIXES if p in line), None)
        if _hit:
            if cur_id and cur_items: kits[cur_id] = cur_items
            m = re.search(rf"{_hit}[A-Z0-9]+", line)
            c = re.search(r"([A-Z0-9]{9})$", line)
            model = m.group() if m else "UNIT"
            cur_id = f"{model}_{c.group(1)}" if c else model
            cur_items, seen = [], set()
            continue
        am = re.search(r"([A-Z0-9\-]+-ASM)\s+(.+)", line)
        if am and cur_id:
            sku = am.group(1).replace("-ASM", "")
            if sku in seen: continue
            desc = shp_standardize(re.sub(r"\s+\d{2}-[A-Z0-9]+.*$", "", am.group(2)).strip())
            cur_items.append({"sku": sku, "desc": desc}); seen.add(sku)
    if cur_id and cur_items: kits[cur_id] = cur_items
    return kits

def generate_shipment_zpl(kit_id, items, w_mm, h_mm, dpi, copies):
    """ZPL etykiety zawartosci kitu, konfigurowalne W/H/DPI (dziedzictwo v10.2).
       Zwraca (zpl, shown, total). Przy obcieciu drukuje marker '+N more'
       instead of the silent break the original had."""
    mult = 8 if int(dpi) == 203 else 12
    pw, ll = int(w_mm) * mult, int(h_mm) * mult
    z = ["^XA", f"^PW{pw}", f"^LL{ll}", "^CI28",
         f"^FO20,20^A0N,22,22^FD{kit_id}^FS",
         f"^FO20,48^GB{pw-40},2,2^FS"]
    y, shown, total = 65, 0, len(items)
    for it in items:
        if y > (ll - 30): break
        z.append(f"^FO20,{y}^A0N,20,20^FD{it['sku']}^FS")
        z.append(f"^FO135,{y}^A0N,18,18^FD- {it['desc']}^FS")
        y += 26; shown += 1
    if shown < total:
        if y > (ll - 20) and shown:       # no room for the marker - sacrifice the last line
            z = z[:-2]; shown -= 1; y -= 26
        z.append(f"^FO20,{y}^A0N,16,16^FD+{total-shown} more^FS")
    z.append(f"^PQ{int(copies)}"); z.append("^XZ")
    return "\n".join(z), shown, total


# ================= FORWARDER (shipment -> payload for the forwarder portal) =================
# Extracted from a standalone prototype once the address parser stabilised.
# The parser identifies address lines BY CHARACTER - the "Delivery Address" block has no fixed structure.
PARCEL_PRESETS = [
    ("XS  20x15x10", 20, 15, 10),
    ("S   30x20x15", 30, 20, 15),
    ("M   40x30x20", 40, 30, 20),
    ("L   60x40x40", 60, 40, 40),
    ("XL  80x60x50", 80, 60, 50),
    ("Pallet 120x80x100", 120, 80, 100),
]

# Postcode patterns - how many TOKENS of the line belong to the code and how many to the city.
# Without this "3542 AW Utrecht" yielded code "3542" and city "AW Utrecht" -> error_postal_code_format.
POSTCODE_RULES = [
    ("Netherlands",    r"^\d{4}\s*[A-Za-z]{2}$",                    2),
    ("United Kingdom", r"^[A-Za-z]{1,2}\d[A-Za-z\d]?\s*\d[A-Za-z]{2}$", 2),
    ("Sweden",         r"^\d{3}\s*\d{2}$",                          2),
    ("Poland",         r"^\d{2}-\d{3}$",                            1),
    ("Portugal",       r"^\d{4}-\d{3}$",                            1),
]

def split_postcode(line, country=""):
    """Returns (postcode, city). Country pattern first, then a general heuristic."""
    toks = (line or "").split()
    if not toks:
        return "", ""
    for cname, rx, n in POSTCODE_RULES:
        if country and country.lower() != cname.lower():
            continue
        for take in (2, 1):
            cand = " ".join(toks[:take])
            if re.match(rx, cand):
                return cand, " ".join(toks[take:]).strip()
    # general case: "1234 AB City" - digits plus a short letter block is still a postcode
    if len(toks) >= 2 and re.match(r"^\d{3,5}$", toks[0]) and re.match(r"^[A-Za-z]{2}$", toks[1]):
        return f"{toks[0]} {toks[1]}", " ".join(toks[2:]).strip()
    return toks[0], " ".join(toks[1:]).strip()

# Country from the document -> name in the forwarder list (extend as needed)
COUNTRY_FIX = {"UK": "United Kingdom", "GB": "United Kingdom", "NL": "Netherlands",
               "DE": "Germany", "SE": "Sweden", "FR": "France", "BE": "Belgium"}


# ---------------------------------------------------------------- parser
COUNTRIES = {
    "netherlands": "Netherlands", "nederland": "Netherlands", "holland": "Netherlands",
    "sweden": "Sweden", "sverige": "Sweden", "germany": "Germany", "deutschland": "Germany",
    "united kingdom": "United Kingdom", "uk": "United Kingdom", "england": "United Kingdom",
    "belgium": "Belgium", "belgie": "Belgium", "france": "France", "poland": "Poland",
    "denmark": "Denmark", "norway": "Norway", "finland": "Finland", "ireland": "Ireland",
    "portugal": "Portugal", "italy": "Italy", "spain": "Spain", "austria": "Austria", "switzerland": "Switzerland",
}

# Registration and tax numbers can sit inside the address block - they are NOT address lines
REG_RX = re.compile(r"^(kvk|btw|vat|coc|reg|tax|ust|org)[\s.:\-]?\w*\d"
                    r"|^[A-Za-z]{2}\d{8,12}[A-Za-z]?\d{0,2}$"     # VAT with a country prefix: PL1234567890, NL123456789B01
                    r"|^kvk\d+$", re.I)

# Postcode patterns used to DETECT the line (country independent)
PC_DETECT = [
    (r"^\d{4}\s*[A-Za-z]{2}\b", "Netherlands"),
    (r"^[A-Za-z]{1,2}\d[A-Za-z\d]?\s+\d[A-Za-z]{2}\b", "United Kingdom"),
    (r"^\d{3}\s\d{2}\b", "Sweden"),
    (r"^\d{5}\b", ""),          # DE / SE without a space / FR - the country is never guessed
    (r"^\d{2}-\d{3}\b", "Poland"),
    (r"^\d{4}-\d{3}\b", "Portugal"),
]


def _looks_like_person(line):
    """Contact person versus district or extra address line. Person names in these documents
       maja 2-3 czlony pisane wielka litera ("Edwin Venema", "Daniel Kantor Kanon"),
       while districts are single words. Position relative to the street is not enough on its own."""
    # Name particles such as "de", "van", "von", "da" are lowercase - still a person
    PARTICLES = {"de", "van", "der", "den", "von", "la", "le", "du", "di", "da", "dos", "el", "ter"}
    t = (line or "").split()
    if not (2 <= len(t) <= 4):
        return False
    if not any(w[:1].isupper() for w in t):
        return False
    for w in t:
        if not (w[:1].isupper() or w.lower() in PARTICLES):
            return False
        if not re.fullmatch(r"[^\W\d_]+[-'\u2019]?[^\W\d_]*\.?", w, re.UNICODE):
            return False
    return True

def classify_address(lines):
    """The 'Delivery Address' block has NO fixed structure. Lines are classified by character,
       a pozostale ustawiamy WZGLEDEM ULICY - to jedyna stabilna os w tych dokumentach:
       whatever sits BEFORE the street is the contact person, whatever sits AFTER it is a district or extra line."""
    ln = [l.strip() for l in lines if l and l.strip()]
    out = {"company": "", "contact": "", "address1": "", "address2": "", "address3": "",
           "postcode": "", "city": "", "country": ""}
    if not ln:
        return out

    # 1) country on its own line
    for i, l in enumerate(ln):
        if l.lower().strip(",. ") in COUNTRIES:
            out["country"] = COUNTRIES[l.lower().strip(",. ")]
            ln.pop(i); break

    # 2) line carrying the postcode
    for i, l in enumerate(ln):
        hit = False
        for rx, guess in PC_DETECT:
            if re.match(rx, l):
                out["postcode"], out["city"] = split_postcode(l, out["country"] or guess)
                if not out["country"] and guess:
                    out["country"] = guess       # country inferred from the code format when the document omits it
                ln.pop(i); hit = True; break
        if hit:
            break

    # 3) registration/tax numbers: KVK/VAT/BTW and lines made up ONLY of digits
    #    (Portuguese NIF, Dutch KVK) - not part of the address
    ln = [l for l in ln
          if not REG_RX.match(l.replace(" ", ""))
          and not re.fullmatch(r"[\d\s.\-]{6,}", l)]

    if not ln:
        return out
    out["company"] = ln.pop(0)

    # 4) street: the last remaining line containing a digit (house number)
    street_idx = None
    for i, l in enumerate(ln):
        if re.search(r"\d", l):
            street_idx = i
    if street_idx is not None:
        out["address1"] = ln.pop(street_idx)
        before, after = ln[:street_idx], ln[street_idx:]
    else:
        before, after = ln, []

    # 5) before the street = contact person, after it = district or extra address lines
    rest = list(before) + list(after)
    person = next((l for l in rest if _looks_like_person(l)), "")
    if person:
        out["contact"] = person
        rest.remove(person)
    extra = [l for l in rest if l]
    if extra:
        out["address2"] = extra[0]
        if len(extra) > 1:
            out["address3"] = extra[1]
    return out
def _value_below(words, page_width, label_tokens, ymax=26):
    """The value sitting below a label. Takes the WHOLE ROW, not the first word:
       "Your Order No" potrafi miec wartosc wieloczlonowa (np. ORDER 260806/VOORRAAD).
       The column is determined from the label position, so text from the neighbouring one is not pulled in."""
    mid = page_width / 2.0
    seq = list(label_tokens)
    for i, w in enumerate(words):
        if w["text"] != seq[0]:
            continue
        chunk = words[i:i + len(seq)]
        if [c["text"] for c in chunk] != seq:
            continue
        lx = chunk[0]["x0"]
        ly = max(c["bottom"] for c in chunk)
        right = lx > mid
        cand = [v for v in words
                if ly - 2 < v["top"] < ly + ymax and ((v["x0"] > mid) if right else (v["x0"] < mid))]
        if not cand:
            continue
        row_top = min(v["top"] for v in cand)
        row = [v for v in cand if abs(v["top"] - row_top) < 4]
        row.sort(key=lambda v: v["x0"])
        return " ".join(v["text"] for v in row).strip()
    return ""


def parse_shipment(path):
    """Sales shipment PDF into a data dictionary. The address is read POSITIONALLY (left column),
       because in linear text it interleaves with the reference column on the right."""
    if pdfplumber is None:
        raise RuntimeError("pdfplumber is missing. Install it with: py -m pip install pdfplumber")
    with pdfplumber.open(path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words()
        text = page.extract_text() or ""

    def first_top(token):
        cand = [w["top"] for w in words if w["text"] == token]
        return min(cand) if cand else None

    out = {"file": os.path.basename(path)}

    # --- references: taken POSITIONALLY, not from linear text ---
    # In linear text the right column (Shipment No, Sales Order No.) interleaves
    # with the address block, so newline regexes picked up values from the neighbouring column.
    # Take the word sitting DIRECTLY BELOW the label, within the same X band.
    def value_below(*label_tokens, ymax=26):
        return _value_below(words, page.width, list(label_tokens), ymax)

    out["shipment_no"] = value_below("Shipment", "No")
    out["sales_order"] = value_below("Sales", "Order", "No.")
    out["your_order"]  = value_below("Your", "Order", "No")
    out["account"]     = value_below("Customer", "Account", "No.")
    out["date"]        = value_below("Date")
    # Sender phone from the header - the portal rejects the "00000000" filler as an invalid number,
    # and the sender number is valid and sensible: on a problem the courier calls us.
    mt = re.search(r"T:\s*(\+\d[\d()\s\-]{6,}\d)", text)
    if mt:
        _ph = re.sub(r"\(\s*0\s*\)", "", mt.group(1))          # "(0)" is domestic notation, redundant in the international format
        out["sender_phone"] = re.sub(r"\s{2,}", " ", _ph.replace("(", "").replace(")", "")).strip()
    else:
        out["sender_phone"] = ""

    m = re.search(r"Dispatch date:\s*([0-9]{1,2}-[A-Za-z]{3}-[0-9]{4})", text)
    out["dispatch"] = m.group(1) if m else ""

    # --- address block: left column between "Delivery Address" and "Customer Account No." ---
    top = first_top("Delivery")
    bot = first_top("Customer")
    addr = []
    if top is not None and bot is not None:
        left = [w for w in words if w["x0"] < 300 and top + 5 < w["top"] < bot - 2]
        rows = {}
        for w in sorted(left, key=lambda w: (round(w["top"]), w["x0"])):
            rows.setdefault(round(w["top"] / 3), []).append(w["text"])
        addr = [" ".join(v).strip() for _, v in sorted(rows.items()) if " ".join(v).strip()]

    out["address_raw"] = addr
    out.update(classify_address(addr))

    # --- line items ---
    items = []
    started = False
    for ln in text.splitlines():
        if ln.strip().startswith("Quantity Item"):
            started = True
            continue
        if not started or ln.startswith("Dispatch date"):
            continue
        m = re.match(r"^\s*(\d+)?\s*([A-Z0-9][A-Z0-9\-\.]{3,})\s+(.*?)\s+(\d+)\s*$", ln)
        if m:
            qty = int(m.group(1)) if m.group(1) else 1
            items.append({"qty": qty, "sku": m.group(2), "desc": m.group(3).strip()})
    # freight charge lines are not goods
    out["items"] = [i for i in items if not re.match(r"^DEL[\s\-]", i["sku"])]
    out["all_items"] = items
    return out


FIELD_MAX = {"company_name": 50, "contact_name": 50, "address1": 50, "address2": 50,
             "address3": 50, "city": 50, "description": 120}


def _cap(value, key):
    """Portal ma twarde limity dlugosci (Company name: 50 znakow) i odrzuca dluzsze wpisy."""
    v = (value or "").strip()
    n = FIELD_MAX.get(key)
    return v[:n].strip() if n and len(v) > n else v


def build_payload(data, ui):
    """Payload for the bookmarklet. Keys map to the forwarder form fields."""
    # The parcel description is a fixed constant rather than a per-item value:
    # forwarder portals validate it against a commercial description list.
    desc = ui.get("description") or "Two way radios"
    # Reference mapping: the internal order number becomes the sender reference,
    # the customer order number becomes the delivery reference.
    ref = ui.get("reference") or data.get("sales_order", "")
    return {
        "your_reference": ref,
        "delivery_reference": ui.get("delivery_reference") or data.get("your_order", ""),
        "country": data.get("country", ""),
        "company_name": _cap(data.get("company", "") or "Unknown", "company_name"),
        "contact_name": _cap(data.get("contact", "") or data.get("company", "") or "Unknown", "contact_name"),
        "address1": _cap(data.get("address1", ""), "address1"),
        "address2": _cap(data.get("address2", ""), "address2"),
        "address3": _cap(data.get("address3", ""), "address3"),
        "postcode": data.get("postcode", ""),
        "city": _cap(data.get("city", ""), "city"),
        # The portal VALIDATES the phone format - "00000000" is rejected. When the consignee has no number,
        # the sender number from the document header is sent: it is valid and the courier has someone to call.
        "telephone": ((ui.get("telephone") or "").strip()
                      or (data.get("sender_phone") or "").strip()
                      or "+31 475252041"),
        "email": ui.get("email", ""),
        "parcels": int(ui.get("parcels") or 1),
        "description": _cap(desc, "description"),
        "weight_kg": ui.get("weight") or "0",
        "length_cm": ui.get("length") or "0",
        "width_cm": ui.get("width") or "0",
        "height_cm": ui.get("height") or "0",
        "value": ui.get("value") or "0",
        "packing_type": "Box",          # always parcels through the forwarder
        "_source": data.get("file", ""),
    }




# ---------- CUSTOMER FILE (enriches shipment data) ----------
CUSTOMER_DB = {"by_no": {}, "by_name": {}, "n": 0, "file": ""}
_CUST_SUFFIX = re.compile(
    r"\b(b\.?v\.?|n\.?v\.?|gmbh|ltd|limited|sarl|sas|s\.?a\.?|srl|s\.?r\.?l\.?|"
    r"sp\.? ?z ?o\.?o\.?|spa|ab|as|a/s|aps|oy|kft|d\.?o\.?o\.?|vof|v\.?o\.?f\.?|"
    r"lda|slu|sl|kg|co|inc|ug|mbh|holding|group)\b\.?", re.I)

def _cust_key(name):
    """Name key tolerant of spelling variants: 'Example B.V.' = 'Example BV' = 'EXAMPLE'."""
    n = (name or "").lower()
    n = n.split(" t/a ")[0].split("(")[0]
    n = _CUST_SUFFIX.sub(" ", n)
    n = re.sub(r"[^a-z0-9]+", "", n)
    return n

def load_customer_db(path):
    """Customer file export (No.;Name;Address;City;Contact;Phone No.;Email).
       Used ONLY to fill gaps - the address on the document always wins."""
    import csv as _csv
    if not path or not os.path.exists(path):
        return 0, "file not found"
    try:
        raw = open(path, "r", encoding="utf-8-sig", errors="ignore").read()
        lines = raw.splitlines()
        if not lines: return 0, "file is empty"
        delim = max([";", ",", "\t", "|"], key=lambda d: lines[0].count(d))
        rd = list(_csv.DictReader(lines, delimiter=delim))
        if not rd: return 0, "no rows"
        cols = [c for c in (rd[0].keys() or []) if c]
        def pick(*frags):
            for c in cols:
                lc = c.lower().replace(" ", "").replace(".", "").replace("_", "")
                if any(f in lc for f in frags): return c
            return None
        c_no, c_nm = pick("no", "customerno"), pick("name")
        c_ad, c_ci = pick("address", "street"), pick("city", "town")
        c_ct, c_ph = pick("contact"), pick("phone", "tel")
        c_em = pick("email", "mail")
        if not (c_no and c_nm): return 0, "No./Name columns not found"
        by_no, by_name = {}, {}
        for r in rd:
            no = (r.get(c_no) or "").strip().upper()
            nm = (r.get(c_nm) or "").strip()
            if not no: continue
            rec = {"no": no, "name": nm,
                   "address": (r.get(c_ad) or "").strip() if c_ad else "",
                   "city": (r.get(c_ci) or "").strip() if c_ci else "",
                   "contact": (r.get(c_ct) or "").strip() if c_ct else "",
                   "phone": (r.get(c_ph) or "").strip() if c_ph else "",
                   "email": (r.get(c_em) or "").strip() if c_em else ""}
            # NOTE: "un known" is a DELIBERATE filler - the target portal requires a contact person,
            # so it is written on purpose when the contact is unknown. Only the spelling is normalised.
            # The "-" and empty variants are ignored, since they add nothing.
            for k in ("contact", "phone", "email"):
                _v = rec[k].strip()
                if _v in ("-", "--", "n/a", "N/A", "brak"):
                    rec[k] = ""
            if rec["contact"].lower().replace(" ", "") in ("unknown", "unkown"):
                rec["contact"] = "Unknown"
            by_no[no] = rec
            key = _cust_key(nm)
            if key and key not in by_name:      # first occurrence wins
                by_name[key] = rec
        CUSTOMER_DB.update(by_no=by_no, by_name=by_name, n=len(by_no), file=path)
        return len(by_no), f"{len(by_no)} customers loaded"
    except Exception as e:
        return 0, f"read error: {str(e)[:80]}"

def match_customer(account="", name=""):
    """The account number on the document is the UNAMBIGUOUS key; the name is a fallback only."""
    if account:
        rec = CUSTOMER_DB["by_no"].get(account.strip().upper())
        if rec: return rec, "account"
    key = _cust_key(name)
    if key:
        rec = CUSTOMER_DB["by_name"].get(key)
        if rec: return rec, "name"
        for k, r in CUSTOMER_DB["by_name"].items():   # the document is sometimes truncated
            if key and (k.startswith(key) or key.startswith(k)) and min(len(k), len(key)) >= 6:
                return r, "name~"
    return None, ""

def enrich_from_customers(data):
    """Fills GAPS from the customer file. The document address always takes precedence -
       dropshipments and customers with several delivery addresses are the norm, not the exception."""
    rec, how = match_customer(data.get("account", ""), data.get("company", ""))
    out = dict(data)
    out["match"] = {"found": bool(rec), "how": how, "no": (rec or {}).get("no", ""),
                    "name": (rec or {}).get("name", ""), "address_differs": False}
    if not rec:
        return out
    for src, dst in (("phone", "telephone"), ("email", "email"), ("contact", "contact")):
        if not (out.get(dst) or "").strip() and rec.get(src):
            out[dst] = rec[src]
    doc_addr = re.sub(r"[^a-z0-9]", "", (data.get("address1", "") + data.get("city", "")).lower())
    db_addr = re.sub(r"[^a-z0-9]", "", (rec.get("address", "") + rec.get("city", "")).lower())
    if doc_addr and db_addr and doc_addr != db_addr:
        out["match"]["address_differs"] = True     # dropshipment or a different delivery address
    return out


# ---------- BROWSER BOOKMARK INSTALLATION ----------
# The bookmarklet is percent-encoded and contains NO double quotes, so it can be placed
# in an href="..." attribute without further escaping. Apostrophes are harmless there.
FORWARDER_BOOKMARKLET = r"""javascript%3A%28async%20function%28%29%20%7B%20%20%20let%20raw%3B%20%20%20try%20%7B%20%20%20%20%20raw%20%3D%20await%20navigator.clipboard.readText%28%29%3B%20%20%20%7D%20catch%20%28e%29%20%7B%20%20%20%20%20alert%28%22Cannot%20read%20the%20clipboard.%5CnClick%20once%20on%20the%20page%20background%20and%20try%20again.%22%29%3B%20%20%20%20%20return%3B%20%20%20%7D%20%20%20let%20d%3B%20%20%20try%20%7B%20%20%20%20%20d%20%3D%20JSON.parse%28raw%29%3B%20%20%20%7D%20catch%20%28e%29%20%7B%20%20%20%20%20alert%28%20%20%20%20%20%20%20%22The%20clipboard%20holds%20no%20shipment%20payload%20%28JSON%29.%5CnCopy%20the%20payload%20from%20the%20Forwarder%20view%20first.%22%20%20%20%20%20%20%20%29%3B%20%20%20%20%20return%3B%20%20%20%7D%20%20%20const%20MAP_MAIN%20%3D%20%5B%20%20%20%20%20%5B%22your_reference%22%2C%20%5B%22your%20reference%22%5D%5D%2C%20%20%20%20%20%5B%22delivery_reference%22%2C%20%5B%22delivery%20reference%22%5D%5D%2C%20%20%20%20%20%5B%22company_name%22%2C%20%5B%22company%20name%22%5D%5D%2C%20%20%20%20%20%5B%22contact_name%22%2C%20%5B%22contact%20name%22%5D%5D%2C%20%20%20%20%20%5B%22address1%22%2C%20%5B%22address%20line%201%22%2C%20%22street%20name%22%5D%5D%2C%20%20%20%20%20%5B%22address2%22%2C%20%5B%22address%20line%202%22%5D%5D%2C%20%20%20%20%20%5B%22address3%22%2C%20%5B%22address%20line%203%22%5D%5D%2C%20%20%20%20%20%5B%22postcode%22%2C%20%5B%22postal%20code%22%2C%20%22postcode%22%2C%20%229999%20aa%22%5D%5D%2C%20%20%20%20%20%5B%22city%22%2C%20%5B%22town%22%2C%20%22city%22%5D%5D%2C%20%20%20%20%20%5B%22telephone%22%2C%20%5B%22telephone%22%2C%20%22phone%22%5D%5D%2C%20%20%20%20%20%5B%22email%22%2C%20%5B%22email%22%5D%5D%20%20%20%5D%3B%20%20%20const%20MAP_PARCEL%20%3D%20%5B%20%20%20%20%20%5B%22description%22%2C%20%5B%22parcel%20description%22%5D%5D%2C%20%20%20%20%20%5B%22weight_kg%22%2C%20%5B%22parcel%20weight%22%5D%5D%2C%20%20%20%20%20%5B%22length_cm%22%2C%20%5B%22parcel%20length%22%5D%5D%2C%20%20%20%20%20%5B%22width_cm%22%2C%20%5B%22parcel%20width%22%5D%5D%2C%20%20%20%20%20%5B%22height_cm%22%2C%20%5B%22parcel%20height%22%5D%5D%2C%20%20%20%20%20%5B%22value%22%2C%20%5B%22value%22%5D%5D%20%20%20%5D%3B%20%20%20const%20norm%20%3D%20%28s%29%20%3D%3E%20%28s%20%7C%7C%20%22%22%29.toLowerCase%28%29.replace%28%2F%5Cs%2B%2Fg%2C%20%22%20%22%29.trim%28%29%3B%20%20%20%20function%20scan%28%29%20%7B%20%20%20%20%20return%20%5B...document.querySelectorAll%28%22input%2C%20select%2C%20textarea%22%29%5D.filter%28%28e%29%20%3D%3E%20e.type%20%21%3D%3D%20%20%20%20%20%20%20%22hidden%22%20%26%26%20%21e.disabled%20%26%26%20e.offsetParent%20%21%3D%3D%20null%29.map%28%28e%29%20%3D%3E%20%7B%20%20%20%20%20%20%20let%20lab%20%3D%20%22%22%3B%20%20%20%20%20%20%20if%20%28e.labels%20%26%26%20e.labels%5B0%5D%29%20lab%20%3D%20e.labels%5B0%5D.innerText%3B%20%20%20%20%20%20%20if%20%28%21lab%29%20%7B%20%20%20%20%20%20%20%20%20const%20cell%20%3D%20e.closest%28%22td%2C%20div%2C%20tr%22%29%3B%20%20%20%20%20%20%20%20%20if%20%28cell%29%20%7B%20%20%20%20%20%20%20%20%20%20%20const%20prev%20%3D%20cell.previousElementSibling%3B%20%20%20%20%20%20%20%20%20%20%20if%20%28prev%29%20lab%20%3D%20prev.innerText%3B%20%20%20%20%20%20%20%20%20%7D%20%20%20%20%20%20%20%7D%20%20%20%20%20%20%20return%20%7B%20%20%20%20%20%20%20%20%20el%3A%20e%2C%20%20%20%20%20%20%20%20%20txt%3A%20norm%28%5Blab%2C%20e.name%2C%20e.id%2C%20e.placeholder%5D.join%28%22%20%22%29%29%20%20%20%20%20%20%20%7D%3B%20%20%20%20%20%7D%29%3B%20%20%20%7D%20%20%20const%20findIn%20%3D%20%28list%2C%20frags%29%20%3D%3E%20%7B%20%20%20%20%20for%20%28const%20f%20of%20frags%29%20%7B%20%20%20%20%20%20%20const%20hit%20%3D%20list.find%28%28x%29%20%3D%3E%20x.txt.includes%28f%29%20%26%26%20%21x.el.dataset.fwFilled%29%3B%20%20%20%20%20%20%20if%20%28hit%29%20return%20hit%3B%20%20%20%20%20%7D%20%20%20%20%20return%20null%3B%20%20%20%7D%3B%20%20%20const%20sleep%20%3D%20%28ms%29%20%3D%3E%20new%20Promise%28%28r%29%20%3D%3E%20setTimeout%28r%2C%20ms%29%29%3B%20%20%20const%20setValue%20%3D%20%28el%2C%20val%29%20%3D%3E%20%7B%20%20%20%20%20const%20v%20%3D%20String%28val%20%3D%3D%20null%20%3F%20%22%22%20%3A%20val%29%3B%20%20%20%20%20if%20%28el.tagName%20%3D%3D%3D%20%22SELECT%22%29%20%7B%20%20%20%20%20%20%20if%20%28%21el.options%20%7C%7C%20%21el.options.length%29%20return%20false%3B%20%20%20%20%20%20%20const%20opt%20%3D%20%5B...el.options%5D.find%28%28o%29%20%3D%3E%20norm%28o.text%29%20%3D%3D%3D%20norm%28v%29%20%7C%7C%20norm%28o.value%29%20%3D%3D%3D%20%20%20%20%20%20%20%20%20norm%28v%29%29%20%7C%7C%20%5B...el.options%5D.find%28%28o%29%20%3D%3E%20norm%28o.text%29.startsWith%28norm%28v%29%29%29%3B%20%20%20%20%20%20%20if%20%28%21opt%29%20return%20false%3B%20%20%20%20%20%20%20el.value%20%3D%20opt.value%3B%20%20%20%20%20%7D%20else%20%7B%20%20%20%20%20%20%20el.focus%28%29%3B%20%20%20%20%20%20%20el.value%20%3D%20v%3B%20%20%20%20%20%7D%20%20%20%20%20el.dispatchEvent%28new%20Event%28%22input%22%2C%20%7B%20%20%20%20%20%20%20bubbles%3A%20true%20%20%20%20%20%7D%29%29%3B%20%20%20%20%20el.dispatchEvent%28new%20Event%28%22change%22%2C%20%7B%20%20%20%20%20%20%20bubbles%3A%20true%20%20%20%20%20%7D%29%29%3B%20%20%20%20%20el.dispatchEvent%28new%20Event%28%22blur%22%2C%20%7B%20%20%20%20%20%20%20bubbles%3A%20true%20%20%20%20%20%7D%29%29%3B%20%20%20%20%20el.style.outline%20%3D%20%222px%20solid%20%2327c66d%22%3B%20%20%20%20%20el.dataset.fwFilled%20%3D%20%221%22%3B%20%20%20%20%20return%20true%3B%20%20%20%7D%3B%20%20%20%20function%20tickCheckbox%28labelFrag%29%20%7B%20%20%20%20%20const%20box%20%3D%20%5B...document.querySelectorAll%28%27input%5Btype%3D%22checkbox%22%5D%27%29%5D.find%28%28e%29%20%3D%3E%20%7B%20%20%20%20%20%20%20let%20lab%20%3D%20e.labels%20%26%26%20e.labels%5B0%5D%20%3F%20e.labels%5B0%5D.innerText%20%3A%20%22%22%3B%20%20%20%20%20%20%20if%20%28%21lab%29%20%7B%20%20%20%20%20%20%20%20%20const%20par%20%3D%20e.closest%28%22td%2C%20div%2C%20label%2C%20tr%22%29%3B%20%20%20%20%20%20%20%20%20lab%20%3D%20par%20%3F%20par.innerText%20%3A%20%22%22%3B%20%20%20%20%20%20%20%7D%20%20%20%20%20%20%20return%20norm%28%5Blab%2C%20e.name%2C%20e.id%5D.join%28%22%20%22%29%29.includes%28labelFrag%29%3B%20%20%20%20%20%7D%29%3B%20%20%20%20%20if%20%28%21box%29%20return%20false%3B%20%20%20%20%20if%20%28%21box.checked%29%20box.click%28%29%3B%20%20%20%20%20box.style.outline%20%3D%20%222px%20solid%20%2327c66d%22%3B%20%20%20%20%20return%20box.checked%3B%20%20%20%7D%20%20%20%20function%20setCountry%28want%29%20%7B%20%20%20%20%20const%20w%20%3D%20norm%28want%29%3B%20%20%20%20%20if%20%28%21w%29%20return%20false%3B%20%20%20%20%20const%20sels%20%3D%20%5B...document.querySelectorAll%28%22select%22%29%5D%3B%20%20%20%20%20for%20%28const%20sel%20of%20sels%29%20%7B%20%20%20%20%20%20%20if%20%28%21sel.options%20%7C%7C%20%21sel.options.length%29%20continue%3B%20%20%20%20%20%20%20const%20opt%20%3D%20%5B...sel.options%5D.find%28%28o%29%20%3D%3E%20norm%28o.text%29%20%3D%3D%3D%20w%20%7C%7C%20norm%28o.value%29%20%3D%3D%3D%20w%29%3B%20%20%20%20%20%20%20if%20%28%21opt%29%20continue%3B%20%20%20%20%20%20%20if%20%28norm%28sel.value%29%20%3D%3D%3D%20norm%28opt.value%29%29%20%7B%20%20%20%20%20%20%20%20%20sel.style.outline%20%3D%20%222px%20solid%20%2327c66d%22%3B%20%20%20%20%20%20%20%20%20return%20true%3B%20%20%20%20%20%20%20%7D%20%20%20%20%20%20%20sel.value%20%3D%20opt.value%3B%20%20%20%20%20%20%20sel.dispatchEvent%28new%20Event%28%22input%22%2C%20%7B%20%20%20%20%20%20%20%20%20bubbles%3A%20true%20%20%20%20%20%20%20%7D%29%29%3B%20%20%20%20%20%20%20sel.dispatchEvent%28new%20Event%28%22change%22%2C%20%7B%20%20%20%20%20%20%20%20%20bubbles%3A%20true%20%20%20%20%20%20%20%7D%29%29%3B%20%20%20%20%20%20%20try%20%7B%20%20%20%20%20%20%20%20%20if%20%28window.jQuery%29%20window.jQuery%28sel%29.trigger%28%22change%22%29%3B%20%20%20%20%20%20%20%7D%20catch%20%28e%29%20%7B%7D%20%20%20%20%20%20%20sel.style.outline%20%3D%20%222px%20solid%20%2327c66d%22%3B%20%20%20%20%20%20%20return%20true%3B%20%20%20%20%20%7D%20%20%20%20%20return%20false%3B%20%20%20%7D%20%20%20%20function%20setPackaging%28want%29%20%7B%20%20%20%20%20const%20w%20%3D%20norm%28want%29%3B%20%20%20%20%20const%20isPack%20%3D%20%28t%29%20%3D%3E%20t.includes%28%22packaging%20type%22%29%20%7C%7C%20t.includes%28%22packing%20type%22%29%3B%20%20%20%20%20const%20sels%20%3D%20%5B...document.querySelectorAll%28%22select%22%29%5D.filter%28%28e%29%20%3D%3E%20isPack%28norm%28%5B%28e%20%20%20%20%20%20%20%20%20.labels%20%26%26%20e.labels%5B0%5D%20%3F%20e.labels%5B0%5D.innerText%20%3A%20%22%22%29%2C%20e.name%2C%20e.id%5D.join%28%22%20%22%29%29%29%20%7C%7C%20%20%20%20%20%20%20isPack%28norm%28%28e.closest%28%22td%2C%20div%2C%20tr%22%29%20%7C%7C%20%7B%7D%29.innerText%20%7C%7C%20%22%22%29%29%29%3B%20%20%20%20%20for%20%28const%20sel%20of%20sels%29%20%7B%20%20%20%20%20%20%20if%20%28%21sel.options%20%7C%7C%20%21sel.options.length%29%20continue%3B%20%20%20%20%20%20%20const%20opt%20%3D%20%5B...sel.options%5D.find%28%28o%29%20%3D%3E%20norm%28o.text%29%20%3D%3D%3D%20w%20%7C%7C%20norm%28o.value%29%20%3D%3D%3D%20w%29%3B%20%20%20%20%20%20%20if%20%28opt%29%20%7B%20%20%20%20%20%20%20%20%20sel.value%20%3D%20opt.value%3B%20%20%20%20%20%20%20%20%20sel.dispatchEvent%28new%20Event%28%22input%22%2C%20%7B%20%20%20%20%20%20%20%20%20%20%20bubbles%3A%20true%20%20%20%20%20%20%20%20%20%7D%29%29%3B%20%20%20%20%20%20%20%20%20sel.dispatchEvent%28new%20Event%28%22change%22%2C%20%7B%20%20%20%20%20%20%20%20%20%20%20bubbles%3A%20true%20%20%20%20%20%20%20%20%20%7D%29%29%3B%20%20%20%20%20%20%20%20%20try%20%7B%20%20%20%20%20%20%20%20%20%20%20if%20%28window.jQuery%29%20window.jQuery%28sel%29.trigger%28%22change%22%29%3B%20%20%20%20%20%20%20%20%20%7D%20catch%20%28e%29%20%7B%7D%20%20%20%20%20%20%20%20%20sel.style.outline%20%3D%20%222px%20solid%20%2327c66d%22%3B%20%20%20%20%20%20%20%20%20return%20true%3B%20%20%20%20%20%20%20%7D%20%20%20%20%20%7D%20%20%20%20%20const%20trigger%20%3D%20%5B...document.querySelectorAll%28%22button%2C%20a%2C%20div%2C%20span%22%29%5D.find%28%28e%29%20%3D%3E%20norm%28e%20%20%20%20%20%20%20%20%20.textContent%29%20%3D%3D%3D%20%22select%20packaging%20type%22%20%7C%7C%20norm%28e.textContent%29%20%3D%3D%3D%20%20%20%20%20%20%20%22select%20packing%20type%22%29%3B%20%20%20%20%20if%20%28trigger%29%20%7B%20%20%20%20%20%20%20trigger.click%28%29%3B%20%20%20%20%20%20%20const%20item%20%3D%20%5B...document.querySelectorAll%28%22li%2C%20option%2C%20div%2C%20span%2C%20a%22%29%5D.find%28%28e%29%20%3D%3E%20norm%28e%20%20%20%20%20%20%20%20%20.textContent%29%20%3D%3D%3D%20w%20%26%26%20e.offsetParent%20%21%3D%3D%20null%29%3B%20%20%20%20%20%20%20if%20%28item%29%20%7B%20%20%20%20%20%20%20%20%20item.click%28%29%3B%20%20%20%20%20%20%20%20%20return%20true%3B%20%20%20%20%20%20%20%7D%20%20%20%20%20%7D%20%20%20%20%20return%20false%3B%20%20%20%7D%20%20%20const%20filled%20%3D%20%5B%5D%2C%20%20%20%20%20missed%20%3D%20%5B%5D%3B%20%20%20const%20applied%20%3D%20%5B%5D%3B%20%20%20%20function%20fillGroup%28map%2C%20list%29%20%7B%20%20%20%20%20for%20%28const%20%5Bkey%2C%20frags%5D%20of%20map%29%20%7B%20%20%20%20%20%20%20const%20val%20%3D%20d%5Bkey%5D%3B%20%20%20%20%20%20%20if%20%28val%20%3D%3D%3D%20undefined%20%7C%7C%20val%20%3D%3D%3D%20%22%22%20%7C%7C%20val%20%3D%3D%3D%20null%29%20continue%3B%20%20%20%20%20%20%20const%20f%20%3D%20findIn%28list%2C%20frags%29%3B%20%20%20%20%20%20%20if%20%28f%20%26%26%20setValue%28f.el%2C%20val%29%29%20%7B%20%20%20%20%20%20%20%20%20filled.push%28key%29%3B%20%20%20%20%20%20%20%20%20applied.push%28%7B%20%20%20%20%20%20%20%20%20%20%20el%3A%20f.el%2C%20%20%20%20%20%20%20%20%20%20%20key%3A%20key%2C%20%20%20%20%20%20%20%20%20%20%20val%3A%20String%28val%29%20%20%20%20%20%20%20%20%20%7D%29%3B%20%20%20%20%20%20%20%7D%20else%20%7B%20%20%20%20%20%20%20%20%20missed.push%28key%29%3B%20%20%20%20%20%20%20%7D%20%20%20%20%20%7D%20%20%20%7D%20%20%20if%20%28d.country%29%20%7B%20%20%20%20%20if%20%28setCountry%28d.country%29%29%20%7B%20%20%20%20%20%20%20filled.push%28%22country%22%29%3B%20%20%20%20%20%20%20await%20sleep%28800%29%3B%20%20%20%20%20%7D%20else%20missed.push%28%22country%22%29%3B%20%20%20%7D%20%20%20fillGroup%28MAP_MAIN%2C%20scan%28%29%29%3B%20%20%20if%20%28d.parcels%20%21%3D%3D%20undefined%20%26%26%20d.parcels%20%21%3D%3D%20%22%22%20%26%26%20String%28d.parcels%29%20%21%3D%3D%20%221%22%29%20%7B%20%20%20%20%20const%20pf%20%3D%20findIn%28scan%28%29%2C%20%5B%22number%20of%20parcels%22%5D%29%3B%20%20%20%20%20if%20%28pf%29%20%7B%20%20%20%20%20%20%20setValue%28pf.el%2C%20d.parcels%29%3B%20%20%20%20%20%20%20filled.push%28%22parcels%22%29%3B%20%20%20%20%20%20%20await%20sleep%28700%29%3B%20%20%20%20%20%7D%20else%20%7B%20%20%20%20%20%20%20missed.push%28%22parcels%22%29%3B%20%20%20%20%20%7D%20%20%20%7D%20%20%20fillGroup%28MAP_PARCEL%2C%20scan%28%29%29%3B%20%20%20if%20%28tickCheckbox%28%22ready%20now%22%29%29%20filled.push%28%22ready_now%22%29%3B%20%20%20else%20missed.push%28%22ready_now%22%29%3B%20%20%20if%20%28setPackaging%28d.packing_type%20%7C%7C%20%22Box%22%29%29%20filled.push%28%22packing_type%22%29%3B%20%20%20else%20missed.push%28%22packing_type%22%29%3B%20%20%20await%20sleep%28600%29%3B%20%20%20let%20repaired%20%3D%200%3B%20%20%20for%20%28const%20a%20of%20applied%29%20%7B%20%20%20%20%20try%20%7B%20%20%20%20%20%20%20if%20%28a.el.isConnected%20%26%26%20String%28a.el.value%20%7C%7C%20%22%22%29%20%21%3D%3D%20a.val%20%26%26%20a.el.tagName%20%21%3D%3D%20%22SELECT%22%29%20%7B%20%20%20%20%20%20%20%20%20a.el.dataset.fwFilled%20%3D%20%22%22%3B%20%20%20%20%20%20%20%20%20if%20%28setValue%28a.el%2C%20a.val%29%29%20repaired%2B%2B%3B%20%20%20%20%20%20%20%7D%20%20%20%20%20%7D%20catch%20%28e%29%20%7B%7D%20%20%20%7D%20%20%20const%20box%20%3D%20document.createElement%28%22div%22%29%3B%20%20%20box.style.cssText%20%3D%20%20%20%20%20%22position%3Afixed%3Bright%3A16px%3Bbottom%3A16px%3Bz-index%3A999999%3Bbackground%3A%23161922%3Bcolor%3A%23e9edf2%3B%22%20%2B%20%20%20%20%20%22font%3A12px%2F1.5%20Segoe%20UI%2CArial%3Bpadding%3A12px%2014px%3Bborder%3A1px%20solid%20%232a2f3a%3Bborder-radius%3A8px%3B%22%20%2B%20%20%20%20%20%22max-width%3A340px%3Bbox-shadow%3A0%206px%2024px%20rgba%280%2C0%2C0%2C.4%29%22%3B%20%20%20box.innerHTML%20%3D%20%27%3Cb%20style%3D%22color%3A%234d8dff%22%3EFill%20shipment%3C%2Fb%3E%3Cbr%3E%27%20%2B%20%20%20%20%20%27%3Cspan%20style%3D%22color%3A%2327c66d%22%3Efilled%3A%20%27%20%2B%20filled.length%20%2B%20%22%3C%2Fspan%3E%22%20%2B%20%28missed.length%20%3F%20%20%20%20%20%20%20%27%3Cbr%3E%3Cspan%20style%3D%22color%3A%23f5b342%22%3Enot%20found%3A%20%27%20%2B%20missed.join%28%22%2C%20%22%29%20%2B%20%22%3C%2Fspan%3E%22%20%3A%20%20%20%20%20%20%20%27%3Cbr%3E%3Cspan%20style%3D%22color%3A%238a93a0%22%3Eall%20mapped%20fields%20filled%3C%2Fspan%3E%27%29%20%2B%20%28repaired%20%3F%20%20%20%20%20%20%20%27%3Cbr%3E%3Cspan%20style%3D%22color%3A%238a93a0%22%3Ere-applied%20after%20form%20refresh%3A%20%27%20%2B%20repaired%20%2B%20%22%3C%2Fspan%3E%22%20%3A%20%20%20%20%20%20%20%22%22%29%20%2B%20%27%3Cbr%3E%3Cspan%20style%3D%22color%3A%235a6675%22%3ECheck%20the%20form%20and%20submit%20manually.%3C%2Fspan%3E%27%3B%20%20%20document.body.appendChild%28box%29%3B%20%20%20setTimeout%28%28%29%20%3D%3E%20box.remove%28%29%2C%209000%29%3B%20%7D%29%28%29%3B"""

def find_chrome():
    """Path to chrome.exe: the App Paths registry key first, then the usual install locations."""
    try:
        import winreg
        for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            try:
                with winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe") as k:
                    path = winreg.QueryValue(k, None)
                    if path and os.path.exists(path):
                        return path
            except Exception:
                continue
    except Exception:
        pass
    for env in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
        base = os.environ.get(env) or ""
        cand = os.path.join(base, "Google", "Chrome", "Application", "chrome.exe")
        if base and os.path.exists(cand):
            return cand
    return ""

def build_setup_page(has_chrome=True):
    """Bookmark installation page. Dragging the link onto the bookmarks bar is the only
       bezpieczny sposob: NIE ruszamy pliku profilu Chrome (suma kontrolna, synchronizacja,
       multiple profiles, a browser that must be closed) and it needs no admin rights."""
    chrome_note = ("" if has_chrome else
        '<div class="warn"><b>Chrome not detected on this station.</b><br>'
        'The console and this bookmark work best in Chrome. Install it, then reopen this page.'
        '<br><a class="btn" href="https://www.google.com/chrome/" target="_blank">Download Chrome</a>'
        '<a class="btn ghost" href="/setup">I have installed it &mdash; check again</a></div>')
    return """<!doctype html><html><head><meta charset="utf-8">
<title>PickCore &middot; Forwarder bookmark setup</title>
<style>
body{background:#1a1714;color:#f2ece1;font:15px/1.6 "Bahnschrift",Arial;margin:0;padding:34px}
.wrap{max-width:760px;margin:0 auto}
h1{color:#2fb5a8;font-size:22px;margin:0 0 4px}
.sub{color:#a99a84;font-size:13px;margin-bottom:26px}
.card{background:#241f1a;border:1px solid #45392c;border-radius:10px;padding:20px 22px;margin-bottom:16px}
.drag{display:inline-block;background:#28a745;color:#fff;text-decoration:none;font-weight:700;
      padding:14px 26px;border-radius:8px;font-size:16px;cursor:grab}
.step{display:flex;gap:12px;margin:10px 0}
.num{flex:0 0 26px;height:26px;border-radius:50%;background:#2fb5a8;color:#fff;text-align:center;
     font-weight:700;font-size:13px;line-height:26px}
code{background:#0b0d12;padding:2px 6px;border-radius:4px;font-family:Consolas,monospace;font-size:13px}
.muted{color:#a99a84;font-size:13px}
.warn{background:#2a1f12;border:1px solid #f5b342;border-radius:10px;padding:16px 18px;margin-bottom:16px}
.btn{display:inline-block;background:#2fb5a8;color:#fff;text-decoration:none;padding:8px 16px;
     border-radius:6px;margin:10px 8px 0 0;font-size:14px}
.btn.ghost{background:#383026}
button{background:#383026;color:#f2ece1;border:0;border-radius:6px;padding:9px 16px;font-size:14px;cursor:pointer}
</style></head><body><div class="wrap">
<h1>Forwarder bookmark setup</h1>
<div class="sub">One-time setup, about 30 seconds. Nothing is installed and no browser settings are changed.</div>
""" + chrome_note + """
<div class="card">
  <div class="step"><div class="num">1</div><div>Press <code>Ctrl+Shift+B</code> to show the bookmarks bar (skip if it is already visible).</div></div>
  <div class="step"><div class="num">2</div><div><b>Drag the green button below onto the bookmarks bar.</b></div></div>
  <div class="step"><div class="num">3</div><div>Done. On the Forwarder shipment form, copy the payload from PickCore and click the bookmark.</div></div>
  <p style="margin:22px 0 6px"><a class="drag" href="%%BOOKMARK%%">&#128230; Fill shipment</a></p>
  <div class="muted">Clicking this link here does nothing useful &mdash; it only works on the Forwarder form.</div>
</div>
<div class="card">
  <div class="muted">If dragging is not possible (some managed browsers block it), copy the address and
  create the bookmark manually: <b>Ctrl+Shift+O</b> &rarr; right-click the bookmarks bar &rarr; Add new bookmark.</div>
  <p><button onclick="copyIt()">Copy bookmark address</button> <span id="msg" class="muted"></span></p>
  <script>
  function copyIt(){
    var a=document.querySelector('.drag').getAttribute('href');
    navigator.clipboard.writeText(a).then(function(){document.getElementById('msg').textContent='copied';},
      function(){document.getElementById('msg').textContent='copy failed - drag the button instead';});
  }
  </script>
</div>
<div class="card"><div class="muted">
  <b>What the bookmark does:</b> reads the shipment payload from your clipboard and fills the Forwarder form
  in your own logged-in tab. It sends nothing over the network, stores no passwords, and never presses Submit.
</div></div>
</div></body></html>""".replace("%%BOOKMARK%%", FORWARDER_BOOKMARKLET)

# ================= /FORWARDER =================

# ========== LOCATION INDEX (where this item already sits) ==========
# Source #1 (works immediately, no credentials): own PA_LINE / RELOC_LINE telemetry.
# Source #2 (optional, when access is granted): the BC adapter - see bc_bin_contents().
# Contract: {sku: [{"bin":..., "n":hit_count, "last":ISO, "src":"local|bc"}, ...]} sorted descending.

LOC_INDEX = {"map": {}, "built": "", "n": 0}

# Sound signatures - shared by the desktop (winsound) and the handheld console (Web Audio).
# Retro-arcade styled - ORIGINAL sequences, not transcriptions of anyone's music.
# (No copyrighted audio anywhere in the project.)
SOUND_KINDS = ("ok", "err", "warn", "line", "done")
SOUND_EXT_PC  = (".wav",)                      # winsound plays WAV ONLY
SOUND_EXT_WEB = (".mp3", ".ogg", ".wav", ".m4a")   # the handheld browser decodes more formats

def sound_dir(cfg):
    d = (cfg.get("sound_dir") or "").strip()
    return d if d and os.path.isdir(d) else os.path.join(app_base_dir(), "sounds")

def find_sound(cfg, kind, exts):
    """Sound file for an event when the operator supplied one. The name is the event kind
       (ok / err / warn / line / done). With no file, a generated tone is played instead."""
    base = sound_dir(cfg)
    for ext in exts:
        f = os.path.join(base, kind + ext)
        if os.path.exists(f):
            return f
    return ""

BEEP_SEQS = {
    # "coin": short blip plus a held higher tone - the most frequent sound of a shift, keep it light
    "ok":   [(988, 55), (1319, 150)],
    # error: a falling second, clearly downward - audible over noise, impossible to mistake for ok
    "err":  [(392, 120), (294, 240)],
    # warning: two mid blips
    "warn": [(660, 80), (660, 80)],
    # line complete: a rising triad
    "line": [(784, 70), (988, 70), (1319, 160)],
    # whole pick done: an arpeggio upward with an accent at the end
    "done": [(523,90),(659,90),(784,90),(1047,110),(988,90),(1047,90),(1319,340)],
}
snd_state = {"id": 0, "kind": ""}

def build_loc_index(max_events=60000):
    """Builds a SKU to bin index from event history. Reads the file backwards, since recent events matter more."""
    idx = {}
    try:
        pth = events_path()
        if not os.path.exists(pth): 
            LOC_INDEX.update(map={}, built=datetime.now().isoformat(timespec="seconds"), n=0); return LOC_INDEX
        with open(pth, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()[-max_events:]
        for ln in lines:
            try: e = json.loads(ln)
            except Exception: continue
            t = e.get("type")
            if t == "PA_LINE":
                sku, b = _norm_sku(e.get("sku") or ""), (e.get("bin") or "").strip().upper()
            elif t == "RELOC_LINE":
                sku, b = _norm_sku(e.get("sku") or ""), (e.get("to") or "").strip().upper()
            else:
                continue
            if not sku or not b: continue
            slot = idx.setdefault(sku, {})
            rec = slot.setdefault(b, {"bin": b, "n": 0, "last": "", "src": "local"})
            rec["n"] += 1
            ts = e.get("ts") or ""
            if ts > rec["last"]: rec["last"] = ts
        # RELOC_LINE marks a move out - the source bin loses weight
        for ln in lines:
            try: e = json.loads(ln)
            except Exception: continue
            if e.get("type") != "RELOC_LINE": continue
            sku, frm = _norm_sku(e.get("sku") or ""), (e.get("frm") or "").strip().upper()
            if sku in idx and frm in idx[sku]:
                idx[sku][frm]["n"] = max(0, idx[sku][frm]["n"] - 2)
        out = {}
        for sku, bins in idx.items():
            ranked = sorted([b for b in bins.values() if b["n"] > 0],
                            key=lambda r: (r["last"], r["n"]), reverse=True)
            if ranked: out[sku] = ranked[:3]
        LOC_INDEX.update(map=out, built=datetime.now().isoformat(timespec="seconds"), n=len(out))
    except Exception as e:
        # Without this index Inbound shows NO suggestions at all and the operator sees
        # an empty column instead of a message. This exact symptom was reported from the floor.
        note_io_fail("build_loc_index", e)
    return LOC_INDEX

def loc_suggest(sku):
    """Returns up to three candidate bins for a SKU, or an empty list."""
    return LOC_INDEX["map"].get(_norm_sku(sku or ""), [])

# ---------- BUSINESS CENTRAL ADAPTER (Bin Contents over OData V4 + OAuth2) ----------
# RULES (non-negotiable):
#  1. Query ONLY the SKUs on the current document, never the whole stock. No background polling.
#  2. The secret NEVER lives in the config or the source: environment variable PICKCORE_BC_SECRET
#     or Windows Credential Manager (keyring). The config holds only tenant/client_id/URL.
#  3. Every failure degrades silently to the local index - operations must not stall on an API.
#  4. Short timeouts: the operator is waiting for a list, not for the network.
BC_TOKEN = {"tok": "", "exp": 0.0}
BC_CACHE = {}          # sku -> (ts, [records])
BC_STATE = {"last": "", "err": "", "n": 0}
bc_state = BC_STATE          # alias used by the GUI (the "refreshing" flag)
BC_TTL = 900           # 15 min - bin contents do not change by the second

def _bc_secret():
    """The secret comes from external sources only. No secret means the adapter stays off."""
    v = os.environ.get("PICKCORE_BC_SECRET")
    if v: return v.strip()
    try:
        import keyring
        return (keyring.get_password("PickCore", "bc_client_secret") or "").strip()
    except Exception:
        return ""

def _bc_token(cfg):
    """OAuth2 client credentials (Entra ID). Token cache'owany do wygasniecia."""
    import urllib.request, urllib.parse, time as _t
    if BC_TOKEN["tok"] and BC_TOKEN["exp"] > _t.time() + 60:
        return BC_TOKEN["tok"]
    tenant, cid, sec = cfg.get("bc_tenant",""), cfg.get("bc_client_id",""), _bc_secret()
    if not (tenant and cid and sec): raise RuntimeError("BC credentials not configured")
    url = cfg.get("bc_token_url") or f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    data = urllib.parse.urlencode({
        "grant_type": "client_credentials", "client_id": cid, "client_secret": sec,
        "scope": cfg.get("bc_scope") or "https://api.businesscentral.dynamics.com/.default",
    }).encode()
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=8) as r:
        j = json.loads(r.read().decode("utf-8"))
    BC_TOKEN["tok"] = j.get("access_token", "")
    BC_TOKEN["exp"] = _t.time() + float(j.get("expires_in", 3600))
    if not BC_TOKEN["tok"]: raise RuntimeError("no access_token in response")
    return BC_TOKEN["tok"]

def bc_bin_contents(skus, cfg, force=False):
    """For a list of SKUs returns [{"sku","bin","qty"}] from ERP Bin Contents.
       Requires page 7379 'Bin Contents' published as a web service (its name lives in cfg['bc_ws'])."""
    import urllib.request, urllib.parse, time as _t
    out, need = [], []
    now = _t.time()
    for sk in {(_norm_sku(x) or "") for x in skus if x}:
        c = BC_CACHE.get(sk)
        if c and not force and (now - c[0]) < BC_TTL: out.extend(c[1])
        else: need.append(sk)
    if not need or not cfg.get("bc_enabled"): 
        return out
    base = (cfg.get("bc_base_url") or "").rstrip("/")
    comp = cfg.get("bc_company") or ""
    ws   = cfg.get("bc_ws") or "BinContents"
    if not (base and comp): return out
    try:
        tok = _bc_token(cfg)
        for i in range(0, len(need), 15):                     # batches of 15 SKUs
            chunk = need[i:i+15]
            flt = " or ".join(f"Item_No eq '{sk}'" for sk in chunk)
            url = (f"{base}/ODataV4/Company('{urllib.parse.quote(comp)}')/{ws}"
                   f"?$filter={urllib.parse.quote(flt)}&$select=Item_No,Bin_Code,Quantity&$top=200")
            req = urllib.request.Request(url, headers={"Authorization": f"Bearer {tok}",
                                                       "Accept": "application/json"})
            with urllib.request.urlopen(req, timeout=10) as r:
                j = json.loads(r.read().decode("utf-8"))
            got = {}
            for row in j.get("value", []):
                sk = _norm_sku(row.get("Item_No") or "")
                b  = (row.get("Bin_Code") or "").strip().upper()
                if not (sk and b): continue
                try: q = float(row.get("Quantity") or 0)
                except Exception: q = 0
                if q <= 0: continue                            # an empty bin is not a suggestion
                rec = {"sku": sk, "bin": b, "qty": q}
                got.setdefault(sk, []).append(rec); out.append(rec)
            for sk in chunk:
                BC_CACHE[sk] = (now, got.get(sk, []))           # empty results are cached too
        BC_STATE.update(last=datetime.now().strftime("%H:%M:%S"), err="", n=len(out))
    except Exception as e:
        BC_STATE.update(err=str(e)[:120])                       # degrade to the local index
    return out

def load_bin_export(path):
    """Third location source: a CSV/TSV exported from Power Query (Excel authenticates
       JAKO UZYTKOWNIK - omija ograniczenia service principala).
       Kolumny wykrywane elastycznie: cokolwiek zawierajace item / bin / qty|quantity|base.
       Returns (records, info)."""
    import csv as _csv
    rows, info = [], ""
    try:
        if not path or not os.path.exists(path):
            return [], "file not found"
        raw = open(path, "r", encoding="utf-8-sig", errors="ignore").read()
        if not raw.strip():
            return [], "file is empty"
        # Separator detected from the first line: Excel under EU locales writes SEMICOLONS.
        head = raw.splitlines()[0] if raw.splitlines() else ""
        delim = max([";", ",", "\t", "|"], key=lambda d: head.count(d))
        if head.count(delim) == 0: delim = ","
        rd = list(_csv.DictReader(raw.splitlines(), delimiter=delim))
        if not rd: return [], "no rows"
        cols = [c for c in (rd[0].keys() or []) if c]
        def pick(*frags):
            for c in cols:
                lc = c.lower().replace(" ", "").replace("_", "")
                if any(f in lc for f in frags): return c
            return None
        c_item = pick("itemno", "item", "nr")
        # Header aliases cover English and Polish exports, so a file saved in either
        # locale is read without a manual column mapping.
        c_bin  = pick("bincode", "bin", "lokacja")
        c_qty  = pick("qtybase", "quantitybase", "base", "quantity", "qty", "ilosc")
        c_desc = pick("description", "opis", "itemdescription")
        if not (c_item and c_bin):
            return [], f"item/bin columns not found (seen: {', '.join(cols[:8])})"
        for r in rd:
            sku = _norm_sku((r.get(c_item) or "").strip())
            b = (r.get(c_bin) or "").strip().upper()
            if not (sku and b): continue
            q = 0.0
            if c_qty:
                try: q = float(str(r.get(c_qty) or 0).replace(",", ".").replace(" ", "") or 0)
                except Exception: q = 0.0
                if q <= 0: continue          # an empty bin is not a suggestion
            rows.append({"sku": sku, "bin": b, "qty": q})
            if c_desc:
                d = (r.get(c_desc) or "").strip()
                if d: learn_sku_desc([{"item": sku, "desc": d}])
        ts = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")
        info = f"{len(rows)} rows, file from {ts}"
    except Exception as e:
        return [], f"read error: {str(e)[:80]}"
    return rows, info

def _agg_by_sku(rows):
    """The ERP returns separate rows per lot and unit of measure, so they are summed by (sku, bin),
       so one bin cannot take two of the three suggestion slots."""
    acc = {}
    for r in rows:
        key = (r["sku"], r["bin"])
        acc[key] = acc.get(key, 0.0) + float(r.get("qty") or 0)
    by = {}
    for (sk, b), q in acc.items():
        by.setdefault(sk, []).append({"bin": b, "n": 999, "last": "9999", "src": "bc", "qty": q})
    return by

def merge_export_into_index(path):
    """Merges the export into LOC_INDEX with priority (src='bc'), exactly like the API path."""
    rows, info = load_bin_export(path)
    if not rows: return 0, info
    by = _agg_by_sku(rows)
    for sk, recs in by.items():
        recs.sort(key=lambda x: -x["qty"])
        local = [c for c in LOC_INDEX["map"].get(sk, []) if c.get("src") != "bc"]
        seen = {c["bin"] for c in recs}
        LOC_INDEX["map"][sk] = (recs + [c for c in local if c["bin"] not in seen])[:3]
    return len(by), info

def merge_bc_into_index(skus, cfg):
    """Scala Bin Contents z BC do LOC_INDEX. Rekordy z BC maja pierwszenstwo (src='bc')."""
    rows = bc_bin_contents(skus, cfg)
    if not rows: return 0
    by = _agg_by_sku(rows)
    for sk, bc_recs in by.items():
        bc_recs.sort(key=lambda x: -x["qty"])
        local = [c for c in LOC_INDEX["map"].get(sk, []) if c.get("src") != "bc"]
        seen = {c["bin"] for c in bc_recs}
        LOC_INDEX["map"][sk] = (bc_recs + [c for c in local if c["bin"] not in seen])[:3]
    return len(by)

# ================= INBOUND (PUT-AWAY SESSION) - helpers =================
SPECIAL_BINS = {"CROSSDOCKING","WARRANTY","ASSEMBLY","DISPATCH"}
def format_bin_input(raw):
    """Maska 'sztywnych myslnikow': '30a03a2' -> '30-A03-A2'. Zwraca (sformatowany, ok).
       Already formatted codes and special bins are accepted as well."""
    s = (raw or "").strip().upper()
    if not s: return "", False
    if s in SPECIAL_BINS: return s, True
    if PICKMAP_LOC.match(s): return s, True
    t = re.sub(r"[^A-Z0-9]", "", s)
    if len(t) == 7:
        f = f"{t[0:2]}-{t[2:5]}-{t[5:7]}"
        if PICKMAP_LOC.match(f): return f, True
    return s, False

LAST_PUTAWAY = {"hdr": "", "rows": []}      # last parsed put-away (prefills the session)
_TEST_HOOKS = {}                             # handles for the smoke-test harness (no runtime impact)

def export_inbound_day_html(entries, out_path):
    """One file per day for downstream tools: Item / Description / Qty / PO (PO at the end of the line,
       zgodnie z prosba: 'note a PO number at the end with each line ... 1 file is fine')."""
    rows_html = "".join(
        f'<tr><td style="border:1px solid #cfc3b0;padding:6px 10px;font-family:Consolas,monospace">{html.escape(e.get("sku",""))}</td>'
        f'<td style="border:1px solid #cfc3b0;padding:6px 10px">{html.escape(e.get("desc",""))}</td>'
        f'<td style="border:1px solid #cfc3b0;padding:6px 10px;text-align:center;font-weight:700">{e.get("qty") or ""}</td>'
        f'<td style="border:1px solid #cfc3b0;padding:6px 10px;font-family:Consolas,monospace;color:#1c6f68;font-weight:700">{html.escape(e.get("po",""))}</td></tr>'
        for e in entries)
    units=sum(int(e.get("qty") or 0) for e in entries); pos=len({e.get("po") for e in entries})
    TH='style="background:#1c6f68;color:#fff;border:1px solid #1c6f68;padding:7px 10px;text-align:left"'
    doc_html=f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>Incoming put-aways</title></head>
<body style="font-family:Arial,sans-serif;background:#f4f6fa;margin:0;padding:22px">
<div style="max-width:820px;margin:auto;background:#fff;border:1px solid #d5dbe6;border-radius:8px;padding:20px 24px">
<div style="font-size:17px;font-weight:700;color:#1c6f68">INCOMING PUT-AWAYS — {datetime.now():%Y-%m-%d}</div>
<div style="font-size:11px;color:#7a6c58;margin:2px 0 12px">{pos} PO · {len(entries)} lines · {units} units · Exported: {datetime.now():%Y-%m-%d %H:%M:%S} · PickCore {APP_VERSION}</div>
<table id="tbl" style="border-collapse:collapse;width:100%">
<tr><th {TH}>Item No.</th><th {TH}>Description</th><th {TH.replace("left","center")}>Qty</th><th {TH}>PO</th></tr>
{rows_html}
<tr><td colspan="2" style="border:1px solid #cfc3b0;padding:6px 10px;text-align:right;font-weight:700">Total units</td>
<td style="border:1px solid #cfc3b0;padding:6px 10px;text-align:center;font-weight:700">{units}</td>
<td style="border:1px solid #cfc3b0"></td></tr></table>
<div style="margin-top:12px">
<button onclick="var r=document.createRange();r.selectNode(document.getElementById('tbl'));
var s=window.getSelection();s.removeAllRanges();s.addRange(r);document.execCommand('copy');s.removeAllRanges();
this.textContent='✓ Copied — paste into Teams';"
style="background:#0E7490;color:#fff;border:0;border-radius:6px;padding:8px 14px;font-weight:700;cursor:pointer">📋 Copy table</button>
<span style="font-size:11px;color:#a99a84;margin-left:10px">→ Teams · Incoming put-aways</span>
</div></div></body></html>"""
    Path(out_path).write_text(doc_html, encoding="utf-8")
    return out_path

def find_pickmap_template():
    """Looks for the rack-view template: exact name first, then a typo-tolerant glob (*pick*template*.html)
       obok exe / folder wyzej / Logi. Gdy brak - materializuje WBUDOWANY szablon do Logi
       (the rack view is self-contained like the 3D view; an external file still takes priority)."""
    import glob as _g
    base = app_base_dir()
    dirs = (base, os.path.dirname(base), os.path.join(base, "Logs"))
    for d in dirs:
        p = os.path.join(d, "pickmap_template.html")
        if os.path.exists(p): return p
    for d in dirs:
        for p in sorted(_g.glob(os.path.join(d, "*pick*template*.html"))):
            if os.path.isfile(p): return p
    try:
        _,_,log_dir = ensure_app_folders()
        p = os.path.join(log_dir, "pickmap_template.html")
        Path(p).write_text(_embedded_template(), encoding="utf-8")
        return p
    except Exception:
        return None

def build_pickmap_heat_html(template_path, heat, unmapped, out_path):
    """Injects live data into the template: PickMap.heatmap(data) once it has loaded,
       plus a banner with the date and pick count and a list of off-map locations. Returns out_path."""
    txt = Path(template_path).read_text(encoding="utf-8")
    payload = json.dumps(dict(heat), ensure_ascii=False)
    n_locs = len(heat); n_lines = sum(heat.values())
    un_rows = "".join(f"<div>{html.escape(k)} · {v}</div>"
                      for k, v in sorted(unmapped.items(), key=lambda kv: -kv[1])[:8])
    banner_un = ("" if not unmapped else
        f"<div style='position:fixed;right:14px;bottom:14px;z-index:9;background:#FFF7E6;"
        f"border:1px solid #C98A1B;border-radius:8px;padding:8px 12px;font:12px/1.5 Consolas,monospace;"
        f"color:#5C420B'><b>Off map ({sum(unmapped.values())} lines):</b>{un_rows}</div>")
    inject = (
        "\n<script>\n"
        f"PickMap.heatmap({payload});\n"
        "var _b=document.querySelector('.brand span');"
        f"if(_b) _b.textContent=' HEAT · {n_lines} pick visits · {n_locs} locations · "
        f"{datetime.now():%Y-%m-%d} · PickCore {APP_VERSION}';\n"
        "var _bd=document.getElementById('btn-demo'); if(_bd) _bd.style.display='none';\n"
        "var _bh=document.getElementById('btn-heat'); if(_bh) _bh.textContent='Reload heat',"
        f"_bh.onclick=function(){{PickMap.heatmap({payload});}};\n"
        "</script>\n" + banner_un + "\n")
    txt = txt.replace("</body>", inject + "</body>")
    Path(out_path).write_text(txt, encoding="utf-8")
    return out_path

# ---------------- PICKMAP ISO (faked 3D projection - a view for decision makers) ----------------
def pickmap_iso_data():
    """Agregacja telemetrii per (grupa_stref, alejka, strona, bay) pod widok izometryczny.
       Two zone codes can share one physical position (floor and upper levels) and are summed as one."""
    from collections import Counter
    heat, un = Counter(), Counter()
    for ev in load_telemetry():
        seen = set()                                   # VISITS: locations deduplicated within a single pick
        for ln in ev.get("lines") or []:
            b = (ln.get("bin") or "").strip().upper()
            if b: seen.add(b)
        for b in seen:
            if not PICKMAP_LOC.match(b): un[b] += 1; continue
            zone, aisle, bay = b[:2], b[3], int(b[4:6])
            zg = "30" if zone in ("30", "31") else zone
            side = "odd" if bay % 2 else "even"
            heat[(zg, aisle, side, bay)] += 1
    return heat, un

# Synthetic demo layout. Deliberately generic: one entry per rack block, the
# faces it covers, its bay numbers and its position on the isometric grid.
# Replace this list with your own layout - nothing else in the renderer depends
# on these values, and any location code that finds no block is reported as
# unmapped rather than silently dropped.
_LA, _LB = list(range(2, 12)), list(range(14, 24))
ISO_BLOCKS = [
    # shelf hall (zone 01): rows anchored at the D-odd wall, pairs split by a cross-aisle
    {"label":"D odd",          "faces":[("01","D","odd")],                   "bays":list(range(3,24,2)), "gx":0.0,  "gy":0.0},
    {"label":"D even \u00b7 C odd", "faces":[("01","D","even"),("01","C","odd")], "bays":_LA,            "gx":0.0,  "gy":2.2},
    {"label":"D even \u00b7 C odd", "faces":[("01","D","even"),("01","C","odd")], "bays":_LB,            "gx":13.0, "gy":2.2},
    {"label":"C even \u00b7 B odd", "faces":[("01","C","even"),("01","B","odd")], "bays":_LA,            "gx":0.0,  "gy":4.4},
    {"label":"C even \u00b7 B odd", "faces":[("01","C","even"),("01","B","odd")], "bays":_LB,            "gx":13.0, "gy":4.4},
    {"label":"B even \u00b7 A odd", "faces":[("01","B","even"),("01","A","odd")], "bays":_LA,            "gx":0.0,  "gy":6.6},
    {"label":"B even \u00b7 A odd", "faces":[("01","B","even"),("01","A","odd")], "bays":_LB,            "gx":13.0, "gy":6.6},
    {"label":"Zone 03 \u00b7 A", "faces":[("03","A","even")],                  "bays":[2,4,6,8,10],        "gx":0.0,  "gy":9.4, "vert":True},
    # pallet hall (zone 02): staggered blocks, entry on the lower edge
    {"label":"B odd 07-09",  "faces":[("02","B","odd")],  "bays":[7,9],     "gx":24.0, "gy":0.0, "tall":True},
    {"label":"B odd 05",     "faces":[("02","B","odd")],  "bays":[5],       "gx":28.5, "gy":0.6, "tall":True},
    {"label":"B odd 03",     "faces":[("02","B","odd")],  "bays":[3],       "gx":31.5, "gy":1.4, "tall":True},
    {"label":"B even 02-06", "faces":[("02","B","even")], "bays":[2,4,6],   "gx":25.0, "gy":3.4, "tall":True},
    {"label":"A odd 03-09",  "faces":[("02","A","odd")],  "bays":[3,5,7,9], "gx":26.5, "gy":5.0, "tall":True},
    {"label":"A even 14-16", "faces":[("02","A","even")], "bays":[14,16],   "gx":22.5, "gy":7.6, "tall":True},
    {"label":"A even 10-12", "faces":[("02","A","even")], "bays":[10,12],   "gx":26.0, "gy":7.6, "tall":True},
    {"label":"A even 02-06", "faces":[("02","A","even")], "bays":[2,4,6],   "gx":29.5, "gy":7.6, "tall":True},
]
_ISO_HEAT = ["#3a3f4a", "#FDD79A", "#F7A64B", "#E2641F", "#B02E0C"]   # 0 = empty plus four levels (template palette)

def build_pickmap_iso_html(heat, unmapped, out_path):
    """Samowystarczalny HTML/SVG: izometryczny rzut obu hal. Jeden szescian = jeden bay;
       kolor gory = czestotliwosc wizyt (obie sciany regalu zsumowane), badge = liczba,
       tooltip = rozbicie per sciana. Malarz: rysowanie tyl->przod po (x+y)."""
    U = 27.5
    def iso(x, y, z=0.0):
        return (560 + (x - y) * U * 0.92, 118 + (x + y) * U * 0.50 - z * U * 0.95)
    mx = max(heat.values()) if heat else 1
    def bucket(v):
        if v <= 0: return 0
        if mx <= 1: return 4
        return 1 + min(3, round((v - 1) / (mx - 1) * 3))
    def pg(pts, fill, op="1"):
        s = " ".join(f"{a:.1f},{b:.1f}" for a, b in pts)
        return f'<polygon points="{s}" fill="{fill}" fill-opacity="{op}" stroke="#1a1714" stroke-width="1"/>'
    cells, labels = [], []
    for blk in ISO_BLOCKS:
        h = 1.7 if blk.get("tall") else 1.15
        for i, b in enumerate(blk["bays"]):
            x = blk["gx"] + (0 if blk.get("vert") else i)
            y = blk["gy"] + (i if blk.get("vert") else 0)
            v, det = 0, []
            for (z, a, s) in blk["faces"]:
                if (b % 2 == 1) == (s == "odd"):
                    n = heat.get((z, a, s, b), 0)
                    if n: det.append(f"{z}-{a}{b:02d} {s}: {n}")
                    v += n
            cells.append((x + y, x, y, h, b, v, blk["label"], " · ".join(det) or "no picks"))
        if blk.get("vert"):
            px, py = iso(blk["gx"] + 0.5, blk["gy"] + len(blk["bays"]) + 0.7)
        else:
            px, py = iso(blk["gx"] + len(blk["bays"]) / 2, blk["gy"] + 1.65)
        labels.append((px, py, blk["label"]))
    cells.sort(key=lambda c: c[0])
    P = []
    for _, x, y, h, b, v, lab, det in cells:
        _, B0, C0, D0 = iso(x, y), iso(x+1, y), iso(x+1, y+1), iso(x, y+1)   # A0 is not needed by the painter algorithm
        A1, B1, C1, D1 = iso(x, y, h), iso(x+1, y, h), iso(x+1, y+1, h), iso(x, y+1, h)
        col = _ISO_HEAT[bucket(v)]
        P.append(f'<g><title>{lab} · bay {b:02d}\n{det}</title>'
                 + pg([A1, B1, C1, D1], col)
                 + pg([D0, C0, C1, D1], col if v else "#12151c", "0.55")
                 + pg([B0, C0, C1, B1], col if v else "#191d26", "0.75"))
        tx, ty = iso(x + .5, y + .5, h)
        if v > 0:
            ink = "#5C420B" if bucket(v) < 3 else "#FFF7E6"
            P.append(f'<text x="{tx:.0f}" y="{ty+4:.0f}" text-anchor="middle" font-size="11" font-weight="700" fill="{ink}">{v}</text>')
        else:
            P.append(f'<text x="{tx:.0f}" y="{ty+3:.0f}" text-anchor="middle" font-size="8" fill="#7a6c58">{b:02d}</text>')
        P.append("</g>")
    lab_svg = "".join(f'<text x="{x:.0f}" y="{y:.0f}" text-anchor="middle" font-size="11" fill="#a99a84">{t}</text>'
                      for x, y, t in labels)
    gx_, gy_ = iso(11.7, 8.2); ex, ey = iso(33.6, 9.9)
    extra = (f'<text x="{gx_:.0f}" y="{gy_:.0f}" text-anchor="middle" font-size="10" fill="#7a6c58">cross-aisle</text>'
             f'<text x="{ex:.0f}" y="{ey:.0f}" text-anchor="middle" font-size="12" font-weight="700" fill="#46d17f">▲ ENTRY</text>')
    n_lines, n_locs = sum(heat.values()), len(heat)
    leg = "".join(f'<span style="display:inline-flex;align-items:center;gap:5px;margin-right:12px">'
                  f'<span style="width:13px;height:13px;background:{c};border:1px solid #45392c;display:inline-block"></span>{t}</span>'
                  for c, t in zip(_ISO_HEAT, ["0", "low", "mid", "high", "max"]))
    un_html = "" if not unmapped else (
        '<div style="position:fixed;right:16px;bottom:16px;background:#241f1a;border:1px solid #C98A1B;'
        'border-radius:8px;padding:10px 14px;font:12px Consolas,monospace;color:#f5b342">'
        f'<b>Off map ({sum(unmapped.values())} lines)</b><br>'
        + "<br>".join(f"{html.escape(k)} · {v}" for k, v in sorted(unmapped.items(), key=lambda kv: -kv[1])[:8]) + "</div>")
    doc = f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>PickCore · Warehouse Heat 3D</title>
<style>body{{margin:0;background:#1a1714;color:#f2ece1;font:14px 'Bahnschrift',sans-serif}}
header{{padding:13px 22px;border-bottom:1px solid #45392c;display:flex;gap:16px;align-items:baseline;flex-wrap:wrap}}
h1{{margin:0;font:700 17px Consolas,monospace;color:#2fb5a8}} .sub{{color:#a99a84;font-size:12px}}
g:hover polygon{{filter:brightness(1.28)}} svg{{display:block;margin:0 auto}}</style></head><body>
<header><h1>PICKCORE · WAREHOUSE HEAT — 3D HALL</h1>
<span class="sub">{n_lines} pick visits · {n_locs} bays · {datetime.now():%Y-%m-%d} · PickCore {APP_VERSION} · frequency: 1 pick at a location = 1 (qty ignored)</span>
<span style="margin-left:auto">{leg}</span></header>
<svg viewBox="0 0 1460 780" width="100%">{"".join(P)}{lab_svg}{extra}</svg>{un_html}</body></html>"""
    Path(out_path).write_text(doc, encoding="utf-8")
    return out_path

def load_telemetry():
    """PICK_TELEMETRY events from the event file (process data only, no operator identity)."""
    return [e for e in load_events() if e.get("type") == "PICK_TELEMETRY" and e.get("lines")]

def analyze_warehouse(tele):
    """Analiza slottingowa z telemetrii pickow. METODOLOGIA:
       - cykl (wydruk -> pierwszy skan) = realny czas zbierania; skany przy stacji NIE mierza chodzenia
       - koszt strefy = MEDIANA cykli pickow 1-LINIOWYCH per strefa (naturalna kalibracja:
         staly narzut wziecia kartki skraca sie przy POROWNYWANIU stref)
       - picki z kolejki (queued) wykluczone z czasow - stacja byla zajeta, cykl zawyzony
       Returns: items (ABC), zones (heat), zone_cost, movers, pairs, thru, summary."""
    from statistics import median
    from itertools import combinations
    items = {}; zones = {}; pairs = {}
    zone_singles = {}; multi_spl = []; validations = []
    total_units = 0; days = set()
    for ev in tele:
        lines = ev.get("lines") or []
        days.add((ev.get("ts","") or "")[:10])
        if ev.get("validation_s"): validations.append(ev["validation_s"])
        cyc = ev.get("cycle_s"); queued = ev.get("queued")
        skus_on_pick = set()
        for ln in lines:
            sku = _norm_sku(ln.get("sku","")); z = bin_zone(ln.get("bin",""))
            qty = int(ln.get("qty", 1))
            if not sku: continue
            rec = items.setdefault(sku, {"picks":0, "units":0, "zones":{}})
            rec["picks"] += 1; rec["units"] += qty
            rec["zones"][z] = rec["zones"].get(z, 0) + 1
            zones[z] = zones.get(z, 0) + 1
            total_units += qty
            skus_on_pick.add(sku)
        for a, b in combinations(sorted(skus_on_pick), 2):
            pairs[(a,b)] = pairs.get((a,b), 0) + 1
        if cyc and not queued and cyc > 0:
            if ev.get("n_lines") == 1 and lines:
                zone_singles.setdefault(bin_zone(lines[0].get("bin","")), []).append(cyc)
            elif ev.get("n_lines", 0) > 1:
                multi_spl.append(cyc / ev["n_lines"])
    items_rank = sorted(
        [(s, d["picks"], d["units"], max(d["zones"], key=d["zones"].get)) for s, d in items.items()],
        key=lambda x: (-x[1], -x[2]))
    zones_rank = sorted(zones.items(), key=lambda kv: -kv[1])
    zone_cost = {z: (round(median(v),1), len(v)) for z, v in zone_singles.items() if len(v) >= 3}
    all_singles = [c for v in zone_singles.values() for c in v]
    med1 = round(median(all_singles),1) if all_singles else 0.0
    movers = []
    hot = items_rank[:max(5, len(items_rank)//5)]        # top ~20% (ABC class A)
    for sku, picks, units, z in hot:
        zc = zone_cost.get(z)
        if zc and med1 and zc[0] > med1 * 1.3:           # zone more than 30% slower than the single-line median
            movers.append((sku, picks, z, zc[0]))
    pairs_rank = sorted([p for p in pairs.items() if p[1] >= 2], key=lambda kv: -kv[1])
    summary = {"picks": len(tele), "days": len(days), "units": total_units,
               "med_cycle_1line": med1, "n_singles": len(all_singles),
               "s_per_line": round(median(multi_spl),1) if multi_spl else 0.0,
               "med_validation": round(median(validations),1) if validations else 0.0}
    return {"items": items_rank, "zones": zones_rank, "zone_cost": zone_cost,
            "movers": movers, "pairs": pairs_rank, "summary": summary}

def parse_customer_file(path):
    """Loads customers from Excel (.xlsx/.xls) or CSV. Returns {CODE: name}.
       Column A is the code, column B the name. The header row is skipped."""
    out = {}
    # Same idea: accept both English and Polish header spellings.
    HEADERS = ("no.","no","code","kod","number","nr","customer","klient")
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if not row or len(row) < 2: continue
            code = ("" if row[0] is None else str(row[0])).strip()
            name = ("" if row[1] is None else str(row[1])).strip()
            if not code: continue
            if i == 0 and code.lower() in HEADERS: continue  # header
            out[code.upper()] = name
    else:  # CSV (separator auto-detected, comma or semicolon)
        import csv as _csv
        with open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as f:
            sample = f.read(4096); f.seek(0)
            try:
                dialect = _csv.Sniffer().sniff(sample, delimiters=";,\t")
            except Exception:
                dialect = _csv.excel
                dialect.delimiter = ";" if sample.count(";") >= sample.count(",") else ","
            for i, row in enumerate(_csv.reader(f, dialect)):
                if not row or len(row) < 2: continue
                code = row[0].strip(); name = row[1].strip()
                if not code: continue
                if i == 0 and code.lower() in HEADERS: continue
                out[code.upper()] = name
    return out

# ==================================================================== sheet HTML
SHEET_CSS = """
@page{size:A4;margin:0}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Inter','Bahnschrift',sans-serif;color:#000;background:#5c5f66;padding:24px 12px}
.bar{max-width:210mm;margin:0 auto 16px;display:flex;justify-content:flex-end;gap:12px}
.bar .hint{margin-right:auto;color:#f2ece1;font-size:13px}
.pbtn{font-weight:700;font-size:14px;border:none;cursor:pointer;background:#111;color:#fff;padding:11px 20px;border-radius:8px}
.sheet{width:210mm;min-height:297mm;margin:0 auto;background:#fff;padding:8mm 9mm}
.head{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:3px solid #000;padding-bottom:5px}
.title{font-size:10px;font-weight:800;letter-spacing:.2em;color:#000}
.pino{font-family:'Courier New',monospace;font-size:27px;font-weight:800;margin-top:1px;letter-spacing:.02em}
.docno{font-family:'Courier New',monospace;font-size:11px;font-weight:700;color:#555;margin-top:2px}
.hr{text-align:right;font-size:10px;color:#000}
.hr .b{font-family:'Courier New',monospace;font-size:12px;font-weight:700;color:#000}
.cust-banner{margin:7px 0 0;padding:7px 11px;background:#fff;border:2px solid #000;font-size:13px}
.cust-banner b{font-family:'Courier New',monospace}
.reminder{margin:6px 0 0;border:2.5px solid #000;font-size:13px;font-weight:700;color:#000}
.reminder .rt{font-family:'Courier New',monospace;font-size:10px;letter-spacing:.14em;color:#fff;background:#000;display:block;padding:3px 11px}
.reminder .rb{display:block;padding:7px 11px}
.meta{display:flex;border:1.5px solid #000;border-top:none;margin:7px 0}
.meta div{flex:1;padding:4px 9px;border-right:1px solid #999}
.meta div:last-child{border-right:none}
.meta .k{font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:#000;font-weight:700}
.meta .v{font-family:'Courier New',monospace;font-size:12px;font-weight:700;margin-top:1px}
.zone{display:flex;align-items:center;gap:9px;margin:6px 0 3px;break-after:avoid}
.zone .z{font-family:'Courier New',monospace;font-size:10px;font-weight:800;background:#000;color:#fff;padding:2px 8px}
.zone .zl{font-size:9px;color:#444;letter-spacing:.08em;text-transform:uppercase;font-weight:600}
.zone .ln{flex:1;height:2px;background:#000}
.row{display:grid;grid-template-columns:10mm 7mm 1fr 20mm;align-items:stretch;border:1px solid #000;border-top:none;min-height:9.2mm;break-inside:avoid}
.zone + .row{border-top:1px solid #000}
.c{padding:3px 8px;display:flex;flex-direction:column;justify-content:center}
.chk{border-right:1px solid #000;align-items:center}
.bx{width:6mm;height:6mm;border:2px solid #000;border-radius:1px}
.sq{border-right:1px solid #000;align-items:center;justify-content:center;font-family:'Courier New',monospace;font-size:15px;font-weight:800;color:#000}
.mid{border-right:1px solid #000;gap:1px}
.br{display:flex;align-items:baseline;gap:10px;flex-wrap:wrap}
.sku{font-family:'Courier New',monospace;font-size:18px;font-weight:800;letter-spacing:.01em}
.bin{font-family:'Courier New',monospace;font-size:14px;font-weight:700;color:#000;background:#fff;border:1.5px solid #000;padding:0 7px}
.bin.special{color:#fff;background:#000;border:1.5px solid #000}
.flag{font-family:'Courier New',monospace;font-size:9px;font-weight:800;padding:1px 6px;color:#fff;background:#000}
.desc{font-size:9.5px;color:#333;line-height:1.2}
.qc{align-items:center;justify-content:center;text-align:center}
.ql{font-size:7px;letter-spacing:.16em;color:#000;font-weight:700}
.qty{font-family:'Courier New',monospace;font-size:22px;font-weight:800}
.row.service{background:#e8e8e8}
.build{border:1.5px solid #000;border-left:5px solid #000;margin:9px 0;overflow:hidden;break-inside:auto}
.build-hd{background:#dcdcdc;padding:6px 11px;display:flex;align-items:center;gap:10px;break-after:avoid;border-bottom:1px solid #000}
.build-hd.plain{background:#f2ece1}
.build-hd .tag{font-family:'Courier New',monospace;font-size:9px;font-weight:800;letter-spacing:.08em;color:#fff;background:#000;padding:2px 7px}
.build-hd.plain .tag{color:#000;background:#fff;border:1.5px solid #000}
.build-hd .ord{font-family:'Courier New',monospace;font-size:13px;font-weight:800;color:#000}
.build-hd .radio{font-family:'Courier New',monospace;font-size:12px;font-weight:700;color:#000;margin-left:auto}
.build-body .row{border-left:none;border-right:none}
.build-body .row:first-child{border-top:none}
.build-body .row:last-child{border-bottom:none}
.row.main{background:#f2ece1}
.row.part .sq{color:#888}
.foot{margin-top:8px;border-top:3px solid #000;padding-top:7px;display:flex;justify-content:space-between;align-items:flex-end;break-inside:avoid}
.tot{display:flex;gap:20px}
.tot .k{font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:#000;font-weight:700}
.tot .v{font-family:'Courier New',monospace;font-size:18px;font-weight:800}
.sg{display:flex;gap:18px;text-align:center}
.sg .l{width:40mm;border-bottom:1.5px solid #000;height:24px}
.sg .s{font-size:8px;letter-spacing:.1em;text-transform:uppercase;color:#000;margin-top:3px;font-weight:600}
@media print{body{background:#fff;padding:0}.bar{display:none}.sheet{margin:0;width:auto;min-height:auto}}
"""

def _is_asm(it):     return it["item"].upper().endswith("-ASM")
def _is_service(it): return (it["bin"].upper()=="WARRANTY") or ("SERVICE" in it["item"].upper())

def _row_html(i, it, role=""):
    loc=it["bin"] or "—"; special="-" not in loc
    binclass="bin special" if special else "bin"
    flag='<span class="flag">NON-PICK · confirm</span>' if _is_service(it) else ''
    cls="row"+(" service" if _is_service(it) else "")+((" "+role) if role else "")
    return (f'<div class="{cls}"><div class="c chk"><div class="bx"></div></div>'
            f'<div class="c sq">{i}</div>'
            f'<div class="c mid"><div class="br"><span class="sku">{html.escape(it["item"] or "—")}</span>'
            f'<span class="{binclass}">{html.escape(loc)}</span>{flag}</div>'
            f'<div class="desc">{html.escape(it["desc"] or "")}</div></div>'
            f'<div class="c qc"><span class="ql">PICK</span><span class="qty">{it["qty"]}</span></div></div>')

def build_sheet_html(header_no, rows, customers=None):
    customers = customers or {}
    # --- Kits grouped by the ORDER COLUMN (a hard key from the ERP), not by adjacency ---
    # The ERP does not guarantee line order: parts can appear BEFORE the main unit, and an item
    # without the -ASM suffix used to break the stream (missing frames, main unit landing in LOOSE, malformed kits).
    # Kit = an order group containing a main unit; main = that line; the rest of the group are parts.
    from collections import OrderedDict as _OD
    _by_order = _OD()
    for it in rows:
        _by_order.setdefault((it.get("order") or "").strip(), []).append(it)
    kits, standalone = [], []
    for _ord, _its in _by_order.items():
        _radios = [r for r in _its if is_radio(r.get("item",""))]
        _is_asm_ord = _ord.upper().startswith("AS") or any(
            str(r.get("src","")).lower().startswith("assembly") for r in _its)
        if _is_asm_ord and _radios and len(_its) > 1:
            _main = _radios[0]
            kits.append({'main': _main, 'order': _ord,
                         'parts': [r for r in _its if r is not _main]})
        else:
            standalone.extend(_its)

    body, seq = [], 0
    # 1) assembly kits (document order - main unit before its parts)
    for g in kits:
        radio = g['main']; order = g.get('order') or radio.get("order","")
        rlbl = f'{html.escape(radio["item"])} ×{radio["qty"]}'
        body.append(f'<div class="build assembly"><div class="build-hd"><span class="tag">🔧 ASSEMBLY KIT</span>'
                    f'<span class="ord">{html.escape(order)}</span><span class="radio">{rlbl}</span></div><div class="build-body">')
        seq += 1; body.append(_row_html(seq, radio, "main"))
        for it in sorted(g['parts'], key=lambda r: natural_bin_key(r["bin"])):
            seq += 1; body.append(_row_html(seq, it, "part"))
        body.append('</div></div>')
    # 2) standalone lines (zone walk, sorted by location)
    if standalone:
        if kits:
            body.append('<div class="zone"><span class="z">LOOSE</span>'
                        '<span class="zl">standalone items — walk by location</span><span class="ln"></span></div>')
        last = None
        for it in sorted(standalone, key=lambda r: natural_bin_key(r["bin"])):
            loc = it["bin"] or "—"; special = "-" not in loc
            zone = loc.split("-")[0] if not special else loc
            if zone != last and not kits:
                zl = "non-shelf location" if special else f"aisle {zone} — walk in order"
                body.append(f'<div class="zone"><span class="z">{html.escape(zone)}</span>'
                            f'<span class="zl">{html.escape(zl)}</span><span class="ln"></span></div>'); last = zone
            seq += 1; body.append(_row_html(seq, it))
    rows_html = "".join(body)

    code = _uniform(rows,"cust")
    cinfo = customers.get(code, {})
    cname = cinfo.get("name","")
    reminder = cinfo.get("reminder","")
    cust_display = f'{html.escape(code)}' + (f' — {html.escape(cname)}' if cname else '')
    cust_banner = (f'<div class="cust-banner">Customer: <b>{html.escape(code)}</b>'
                   + (f' — <b>{html.escape(cname)}</b>' if cname else '')
                   + (' (name not in database)' if not cname and code not in ("-","MULTIPLE") else '')
                   + '</div>') if code not in ("-",) else ''
    reminder_html = (f'<div class="reminder"><span class="rt">⚠ CUSTOMER REMINDER</span>'
                     f'<span class="rb">{html.escape(reminder)}</span></div>' if reminder else '')

    orders = order_list(rows); due=html.escape(_uniform(rows,"due")); picker=html.escape(get_picker())
    # The sales order is ALWAYS the primary number, the shipment number a small addition; extra orders are listed after it
    prim = primary_order(rows, header_no)
    others = [o for o in orders if o != prim]
    main_no = html.escape(prim)
    doc_bits = []
    if header_no: doc_bits.append(f"doc {html.escape(header_no)}")
    if others: doc_bits.append("also: " + ", ".join(html.escape(o) for o in others))
    doc_no = "  ·  ".join(doc_bits)
    order_meta = html.escape(prim) + (f" (+{len(others)})" if others else "")
    zones=len({(r["bin"].split("-")[0] if "-" in r["bin"] else r["bin"]) for r in rows})
    units=sum(r["qty"] for r in rows); printed=datetime.now().strftime("%d/%m/%y, %H:%M")

    return f"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<title>Pick Sheet — {main_no}</title><style>{SHEET_CSS}</style></head><body>
<div class="bar"><span class="hint">Print preview — use Print / Save as PDF (this bar won't appear on paper).</span>
<button class="pbtn" onclick="window.print()">🖨 Print / Save PDF</button></div>
<div class="sheet">
<div class="head"><div><div class="title">WAREHOUSE PICK SHEET</div><div class="pino">{main_no}</div>
<div class="docno">{doc_no}</div></div>
<div class="hr"><div>Location <span class="b">MAIN</span></div><div style="margin-top:3px">Printed <span class="b">{printed}</span></div></div></div>
{cust_banner}{reminder_html}
<div class="meta"><div><div class="k">Customer</div><div class="v">{cust_display}</div></div>
<div><div class="k">Order</div><div class="v">{order_meta}</div></div>
<div><div class="k">Due date</div><div class="v">{due}</div></div>
<div><div class="k">Picker</div><div class="v">{picker}</div></div></div>
{rows_html}
<div class="foot"><div class="tot"><div><div class="k">Lines</div><div class="v">{len(rows)}</div></div>
<div><div class="k">Total units</div><div class="v">{units}</div></div>
<div><div class="k">Zones</div><div class="v">{zones}</div></div></div>
<div class="sg"><div><div class="l"></div><div class="s">Picked by / time</div></div>
<div><div class="l"></div><div class="s">Checked by / time</div></div></div></div>
</div></body></html>"""

# ==================================================================== put-away sheet
PA_CSS = """
@page{size:A4;margin:0}*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Inter','Bahnschrift',sans-serif;color:#000;background:#5c5f66;padding:24px 12px}
.bar{max-width:210mm;margin:0 auto 16px;display:flex;justify-content:flex-end;gap:12px}.bar .hint{margin-right:auto;color:#f2ece1;font-size:13px}
.pbtn{font-weight:700;font-size:14px;border:none;cursor:pointer;background:#111;color:#fff;padding:11px 20px;border-radius:8px}
.sheet{width:210mm;min-height:297mm;margin:0 auto;background:#fff;padding:8mm 9mm}
.head{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:3px double #000;padding-bottom:5px}
.title{font-size:10px;font-weight:800;letter-spacing:.2em;color:#000}
.pino{font-family:'Courier New',monospace;font-size:27px;font-weight:800;margin-top:1px;letter-spacing:.02em}
.docno{font-family:'Courier New',monospace;font-size:11px;font-weight:700;color:#555;margin-top:2px}
.hr{text-align:right;font-size:10px;color:#000}.hr .b{font-family:'Courier New',monospace;font-size:12px;font-weight:700;color:#000}
.meta{display:flex;border:1.5px solid #000;border-top:none;margin:7px 0}
.meta div{flex:1;padding:4px 9px;border-right:1px solid #999}.meta div:last-child{border-right:none}
.meta .k{font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:#000;font-weight:700}
.meta .v{font-family:'Courier New',monospace;font-size:12px;font-weight:700;margin-top:1px}
.build{border:1.5px solid #000;border-left:5px solid #000;margin:9px 0;overflow:hidden;break-inside:auto}
.build-hd{background:#dcdcdc;padding:6px 11px;display:flex;align-items:center;gap:10px;break-after:avoid;border-bottom:1px solid #000}
.build-hd.plain{background:#f2ece1}
.build-hd .tag{font-family:'Courier New',monospace;font-size:9px;font-weight:800;letter-spacing:.08em;color:#fff;background:#000;padding:2px 7px}
.build-hd.plain .tag{color:#000;background:#fff;border:1.5px solid #000}
.build-hd .ord{font-family:'Courier New',monospace;font-size:13px;font-weight:800;color:#000}
.build-hd .radio{font-family:'Courier New',monospace;font-size:12px;font-weight:700;color:#000;margin-left:auto}
.row{display:grid;grid-template-columns:10mm 7mm 1fr 56mm 18mm;align-items:stretch;border:1px solid #000;border-top:none;min-height:11mm;break-inside:avoid}
.build-body .row:first-child{border-top:none}.build-body .row:last-child{border-bottom:none}
.row.main{background:#f2ece1}.row.part .sq{color:#888}
.c{padding:4px 8px;display:flex;flex-direction:column;justify-content:center}
.chk{border-right:1px solid #000;align-items:center}.bx{width:6mm;height:6mm;border:2px solid #000;border-radius:1px}
.sq{border-right:1px solid #000;align-items:center;justify-content:center;font-family:'Courier New',monospace;font-size:15px;font-weight:800;color:#000}
.mid{border-right:1px solid #000;gap:1px}.br{display:flex;align-items:baseline;gap:10px}
.sku{font-family:'Courier New',monospace;font-size:18px;font-weight:800}
.desc{font-size:9.5px;color:#333;line-height:1.2}
.loc{border-right:1px solid #000;align-items:center;gap:3px}
.dash{font:700 5mm Consolas;margin:0 1px;align-self:center}
.ll{font-size:7px;letter-spacing:.16em;color:#000;font-weight:800}
.boxes{display:flex;gap:3px}.bxc{display:inline-flex;align-items:center;justify-content:center;width:7mm;height:9mm;border:2px solid #000;border-radius:1px;background:#fff;font:700 4.6mm Consolas,monospace;line-height:1}
.qc{align-items:center;justify-content:center;text-align:center}.ql{font-size:7px;letter-spacing:.16em;color:#000;font-weight:700}
.qty{font-family:'Courier New',monospace;font-size:20px;font-weight:800}
.foot{margin-top:8px;border-top:3px double #000;padding-top:7px;display:flex;justify-content:space-between;align-items:flex-end;break-inside:avoid}
.tot{display:flex;gap:20px}.tot .k{font-size:8px;letter-spacing:.12em;text-transform:uppercase;color:#000;font-weight:700}
.tot .v{font-family:'Courier New',monospace;font-size:18px;font-weight:800}
.sg{display:flex;gap:18px;text-align:center}.sg .l{width:40mm;border-bottom:1.5px solid #000;height:24px}
.sg .s{font-size:8px;letter-spacing:.1em;text-transform:uppercase;color:#000;margin-top:3px;font-weight:600}
@media print{body{background:#fff;padding:0}.bar{display:none}.sheet{margin:0;width:auto;min-height:auto}}
"""

def _pa_boxes(bin_code=""):
    """Bin template 2-3-2 ([_][_]-[_][_][_]-[_][_]). Empty means it is filled in by hand on the printout;
       with bin_code the boxes are PRE-FILLED, giving a digital twin of the paper sheet."""
    ch=re.sub(r"[^A-Z0-9]","",(bin_code or "").upper())[:7].ljust(7)
    def grp(a,b): return ''.join(f'<span class="bxc">{c.strip()}</span>' for c in ch[a:b])
    boxes = ('<div class="boxes">'+grp(0,2)+'<span class="dash">-</span>'
             +grp(2,5)+'<span class="dash">-</span>'+grp(5,7)+'</div>')
    if (bin_code or "").strip():
        # Read-only field: selecting with the MOUSE plus Ctrl+C yields PLAIN TEXT.
        # Selecting a plain <div> copies HTML structure and the ERP then sees "9 rows".
        boxes += (f'<input class="bintxt" readonly value="{html.escape(bin_code.strip().upper())}" '
                  f'onclick="this.select()" title="Zaznacz i skopiuj (Ctrl+C)">')
    return boxes

def _pa_row(i, it, role="", bin_txt=""):
    cls="row"+((" "+role) if role else "")
    return (f'<div class="{cls}"><div class="c chk"><div class="bx"></div></div>'
            f'<div class="c sq">{i}</div>'
            f'<div class="c mid"><div class="br"><span class="sku">{html.escape(it["item"])}</span></div>'
            f'<div class="desc">{html.escape(it["desc"])}</div></div>'
            f'<div class="c loc"><span class="ll">PUT TO</span>{_pa_boxes(bin_txt)}</div>'
            + f'<div class="c qc"><span class="ql">QTY</span><span class="qty">{it["qty"]}</span></div></div>')

def build_putaway_html(header_no, source, rows, filled=False):
    keys, groups = [], {}
    for r in rows:
        k=r.get("source") or "—"
        if k not in groups: groups[k]=[]; keys.append(k)
        groups[k].append(r)
    body, seq = [], 0
    for k in keys:
        grp=groups[k]; parts=[x for x in grp if x["item"].upper().endswith("-ASM")]; mains=[x for x in grp if not x["item"].upper().endswith("-ASM")]
        if parts and mains:
            radio=mains[0]
            body.append(f'<div class="build"><div class="build-hd"><span class="tag">🔧 ASSEMBLY KIT</span>'
                        f'<span class="ord">{html.escape(k)}</span><span class="radio">{html.escape(radio["item"])} ×{radio["qty"]}</span></div><div class="build-body">')
            for it in mains: seq+=1; body.append(_pa_row(seq,it,"main", it.get("bin","") if filled else ""))
            for it in parts: seq+=1; body.append(_pa_row(seq,it,"part", it.get("bin","") if filled else ""))
            body.append('</div></div>')
        else:
            body.append(f'<div class="build"><div class="build-hd plain"><span class="tag">SOURCE PO</span>'
                        f'<span class="ord">{html.escape(k)}</span></div><div class="build-body">')
            for it in grp: seq+=1; body.append(_pa_row(seq,it,"", it.get("bin","") if filled else ""))
            body.append('</div></div>')
    units=sum(r["qty"] for r in rows); printed=datetime.now().strftime("%d/%m/%y, %H:%M"); operator=get_picker()
    # Location copied as PLAIN TEXT - a mouse selection copies HTML structure
    # and the ERP sees "9 rows" instead of one code. writeText() puts the bare string on the clipboard.
    copy_js = ("""<style>
.bintxt{display:block;margin-top:4px;width:120px;border:1px solid #c7cfd8;border-radius:4px;
  background:#fbfcfe;color:#14477d;font:700 13px Consolas,monospace;letter-spacing:.5px;
  padding:3px 6px;text-align:center;cursor:text}
.bintxt:focus{outline:2px solid #2fb5a8;background:#fff}
@media print{.bintxt{display:none}}
</style>
""" if filled else "")
    main_no = html.escape(source or "PUT-AWAY"); doc_no = f"doc {html.escape(header_no)}" if header_no else ""
    return f"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<title>Put-away Sheet — {main_no}</title><style>{PA_CSS}</style></head><body>
<div class="bar"><span class="hint">Put-away sheet — write the location you place each item into the boxes.</span>
<button class="pbtn" onclick="window.print()">🖨 Print / Save PDF</button></div>
<div class="sheet">
<div class="head"><div><div class="title">WAREHOUSE PUT-AWAY SHEET</div><div class="pino">{main_no}</div>
<div class="docno">{doc_no}</div></div>
<div class="hr"><div>Location <span class="b">MAIN</span></div><div style="margin-top:3px">Printed <span class="b">{printed}</span></div></div></div>
<div class="meta"><div><div class="k">Source PO</div><div class="v">{html.escape(source or "-")}</div></div>
<div><div class="k">Lines</div><div class="v">{len(rows)}</div></div>
<div><div class="k">Total units</div><div class="v">{units}</div></div>
<div><div class="k">Operator</div><div class="v">{html.escape(operator)}</div></div></div>
{''.join(body)}
<div class="foot"><div class="tot"><div><div class="k">Lines</div><div class="v">{len(rows)}</div></div>
<div><div class="k">Units</div><div class="v">{units}</div></div></div>
<div class="sg"><div><div class="l"></div><div class="s">Put away by / time</div></div>
<div><div class="l"></div><div class="s">Checked by / time</div></div></div></div>
</div>{copy_js}</body></html>"""

# ==================================================================== Edge render + print
def find_edge():
    for p in [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
              r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"]:
        if os.path.exists(p): return p
    return None

def html_to_pdf(html_path, pdf_path):
    """Renders HTML to PDF through headless Edge (1:1 layout, no admin rights).
       IMPORTANT: Edge exits BEFORE the write completes, so we wait for the file to appear and settle."""
    edge = find_edge()
    if not edge: raise RuntimeError("Microsoft Edge not found - cannot render PDF")
    edge_profile = os.path.join(tempfile.gettempdir(), "pickcore_edge_profile")
    uri = Path(html_path).as_uri()
    # delete the old file so it cannot be confused with the previous render
    try:
        if os.path.exists(pdf_path): os.remove(pdf_path)
    except Exception: pass

    def _wait_pdf(timeout=20):
        """Wait until the PDF exists AND its size stops growing, meaning the write has finished."""
        t0 = time.time(); last = -1
        while time.time() - t0 < timeout:
            if os.path.exists(pdf_path):
                sz = os.path.getsize(pdf_path)
                if sz > 0 and sz == last:   # a stable size means the write has finished
                    return True
                last = sz
            time.sleep(0.5)
        return os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0

    attempts = [
        [edge, "--headless", "--disable-gpu", "--no-pdf-header-footer",
         f"--user-data-dir={edge_profile}", f"--print-to-pdf={pdf_path}", uri],
        [edge, "--headless", "--disable-gpu", "--no-pdf-header-footer",
         f"--print-to-pdf={pdf_path}", uri],
    ]
    for cmd in attempts:
        try:
            subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            continue
        if _wait_pdf(20):
            return pdf_path
    raise RuntimeError("Edge did not produce a PDF")

LAST_PRINT = {"html": None, "hdr": ""}   # last sheet available for reprint (HTML path in the archive)

def print_pdf(pdf_path, printer_name=None):
    """Prints a PDF with no helper files.
       1) 'printto' - drukuje na DOKLADNIE wskazana drukarke (Adobe bierze ja z argumentu,
          NIE z wlasnej pamieci - rozwiazuje 'przyklejona' Zebre). Dziala gdy handler PDF go wspiera.
       2) fallback when printto is unsupported (a different handler): temporary default plus the 'print' verb."""
    if printer_name:
        try:
            import win32api
            r = win32api.ShellExecute(0, "printto", pdf_path, f'"{printer_name}"', ".", 0)
            if isinstance(r, int) and r > 32:    # >32 means success; 32 or less (31 NOASSOC) means no verb
                time.sleep(4)
                return True
        except Exception:
            pass   # fall through to the fallback
    # --- fallback: temporary default printer plus 'print', the most widely supported verb ---
    prev = None
    if printer_name:
        try:
            import win32print
            prev = win32print.GetDefaultPrinter()
            if prev and prev != printer_name:
                win32print.SetDefaultPrinter(printer_name)
            else:
                prev = None
        except Exception:
            prev = None
    try:
        try:
            import win32api
            win32api.ShellExecute(0, "print", pdf_path, None, ".", 0)
        except Exception:
            os.startfile(pdf_path, "print")
        time.sleep(5)
        return True
    except Exception as e:
        print("Print error:", e); return False
    finally:
        if prev:
            try:
                import win32print
                win32print.SetDefaultPrinter(prev)
            except Exception: pass

def list_printers():
    try:
        import win32print
        return [p[2] for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    except Exception:
        return []

# ==================================================================== config
DEFAULT_CFG = {"watch_dir":"","archive_dir_pick":"","archive_dir_pa":"","printer":"","auto_print":False,
               "watcher_on":False,"name_filter":"Warehouse Pick*PI*","pa_filter":"PO*","euro_per_error":150,
               "analytics_owners":[],
               }
def _ver_tuple(v):
    """'1.10' > '1.9' - compared numerically, not lexically."""
    out = []
    for part in str(v or "").strip().split("."):
        d = "".join(ch for ch in part if ch.isdigit())
        out.append(int(d) if d else 0)
    return tuple(out + [0] * (4 - len(out)))[:4]

def check_update(update_dir):
    """Reads the manifest from the network share. Returns (is_newer, version, info)."""
    try:
        d = (update_dir or "").strip()
        if not d: return False, "", "no path set"
        mf = os.path.join(d, "VERSION.txt")
        if not os.path.exists(mf): return False, "", "VERSION.txt not found in the update folder"
        raw = open(mf, "r", encoding="utf-8-sig", errors="ignore").read().strip().splitlines()
        ver = (raw[0] if raw else "").strip()
        note = " ".join(l.strip() for l in raw[1:4]) if len(raw) > 1 else ""
        if not ver: return False, "", "VERSION.txt is empty"
        newer = _ver_tuple(ver) > _ver_tuple(APP_VERSION)
        return newer, ver, (note or ("new version " + ver if newer else "up to date"))
    except Exception as e:
        return False, "", f"read error: {str(e)[:70]}"

def profile_path():
    """Deployment profile sitting NEXT TO the exe - the settings template for a new station."""
    return os.path.join(app_base_dir(), "pickcore_profile.json")

# Station-specific keys: NOT carried between machines.
STATION_KEYS = {"scanner_allow", "printer", "zebra_printer", "analytics_owners"}

def load_cfg():
    # First start on a new machine: with no config but a profile next to the exe, seed from it.
    try:
        if not CONFIG_PATH.exists() and os.path.exists(profile_path()):
            prof = json.loads(open(profile_path(), "r", encoding="utf-8").read())
            seed = {**DEFAULT_CFG, **{k: v for k, v in prof.items()
                                      if k not in STATION_KEYS and not k.startswith("_")}}
            CONFIG_PATH.write_text(json.dumps(seed, indent=1, ensure_ascii=False), encoding="utf-8")
    except Exception as e:
        # A failure here means a new station starts with an empty config despite the profile next to the exe,
        # so "install and it works" silently stops being true.
        note_io_fail("load_cfg(seed z profilu)", e)
    return _load_cfg_raw()

def _load_cfg_raw():
    if CONFIG_PATH.exists():
        try: cfg = {**DEFAULT_CFG, **json.loads(CONFIG_PATH.read_text(encoding="utf-8"))}
        except Exception: cfg = dict(DEFAULT_CFG)
    else:
        cfg = dict(DEFAULT_CFG)
    # folders next to the .exe are the default when unset, which keeps "install and it works" true
    pick_dir, pa_dir, _ = ensure_app_folders()
    if not cfg.get("archive_dir_pick"): cfg["archive_dir_pick"] = pick_dir
    if not cfg.get("archive_dir_pa"):   cfg["archive_dir_pa"]   = pa_dir
    # Same rule for the location file. ONE station publishes bin_contents.csv,
    # the others only read it. Earlier versions required the path to be typed in
    # manually in Settings, which is why some stations ended up with NO bin
    # suggestions at all. Not a code defect, just an option nobody filled in.
    # Search order:
    #   1. the copy next to the exe - excluded from installer overwrites, so it survives updates,
    #   2. the update folder, where the build script places it.
    # A manually entered path ALWAYS wins and is never overwritten.
    # The result is NOT written back to the config: moving the source is then detected
    # on the next start, and startup avoids a pointless disk write.
    if not (cfg.get("bin_export_path") or "").strip():
        cands = [os.path.join(app_base_dir(), "bin_contents.csv")]
        upd = (cfg.get("update_dir") or "").strip()
        if upd: cands.append(os.path.join(upd, "bin_contents.csv"))
        for cand in cands:
            # exists() on a dead UNC share can hang for seconds, so the local copy
            # is checked FIRST and usually ends the loop immediately.
            try: found = os.path.exists(cand)
            except OSError: found = False
            if found:
                cfg["bin_export_path"] = cand
                cfg["_bin_export_auto"] = True     # a trace for Logs, not a setting
                break
    return cfg
def save_cfg(c): CONFIG_PATH.write_text(json.dumps(c, indent=2, ensure_ascii=False), encoding="utf-8")

def wait_stable(path, timeout=15):
    last=-1; t0=time.time()
    while time.time()-t0 < timeout:
        try: sz=os.path.getsize(path)
        except OSError: return False
        if sz==last and sz>0: return True
        last=sz; time.sleep(0.5)
    return os.path.exists(path)

def pick_base_name(header_no, rows):
    """Archive file name is the primary sales order, since that is the key used to find the job in the ERP."""
    return primary_order(rows, header_no)

def unique_pick_path(folder, base, ext=".html"):
    """Pick: base -> base_DATE -> base_DATE_n  (kolizja: data, potem licznik)."""
    folder = Path(folder); folder.mkdir(parents=True, exist_ok=True)
    p = folder / f"{base}{ext}"
    if not p.exists(): return p
    date = datetime.now().strftime("%Y-%m-%d")
    p = folder / f"{base}_{date}{ext}"
    if not p.exists(): return p
    n = 2
    while True:
        p = folder / f"{base}_{date}_{n}{ext}"
        if not p.exists(): return p
        n += 1

def unique_pa_path(folder, base, ext=".html"):
    """Put-away: base -> base (1) -> base (2)  (kolizja: licznik w nawiasie)."""
    folder = Path(folder); folder.mkdir(parents=True, exist_ok=True)
    p = folder / f"{base}{ext}"
    if not p.exists(): return p
    n = 1
    while True:
        p = folder / f"{base} ({n}){ext}"
        if not p.exists(): return p
        n += 1

# ==================================================================== PIPELINE
def process_pdf(pdf_path, cfg, customers, logfn=print, on_pick=None, on_putaway=None):
    """Pelny przeplyw. Wykrywa Pick vs Put-away i routuje do wlasciwego arkusza.
       on_pick(pick_data) - optional callback handing the items to the validation station."""
    name = os.path.basename(pdf_path)
    dtype, hdr, source, rows, warns = parse_document(pdf_path)
    if dtype == "unknown" or not rows:
        logfn(f"✕ {name}: not a Pick/Put-away PDF (ignored)"); return False, "ignored"
    learn_skus(rows)   # known SKU registry - the foundation of scan triage, learning from EVERY document
    if dtype == "putaway":
        LAST_PUTAWAY["hdr"] = hdr; LAST_PUTAWAY["rows"] = list(rows)
        if on_putaway:
            try: on_putaway(hdr, list(rows))
            except Exception as e:
                # This is the entry point to the pa_pending FIFO queue. A silent failure loses
                # the put-away list entirely, which has happened once before.
                logfn(f"   ! put-away queue hook failed: {e}")
                note_io_fail("process_pdf(on_putaway)", e)
        sheet_html = build_putaway_html(hdr, source, rows)
        kind = "PUT-AWAY"
        folder = cfg.get("archive_dir_pa") or tempfile.gettempdir()
        base = source or hdr or "PUTAWAY"
        html_out = unique_pa_path(folder, base)
        title = source or hdr; sub = f"doc {hdr}"
    else:
        sheet_html = build_sheet_html(hdr, rows, customers)
        kind = "PICK"
        folder = cfg.get("archive_dir_pick") or tempfile.gettempdir()
        base = pick_base_name(hdr, rows)
        html_out = unique_pick_path(folder, base)
        code = _uniform(rows, "cust"); cname = customers.get(code, {}).get("name", "")
        cust = code + (f" — {cname}" if cname else "")
        title = primary_order(rows, hdr); sub = f"doc {hdr}" + (f"  ·  {cust}" if code not in ("-",) else "")
    html_out.write_text(sheet_html, encoding="utf-8")
    msg = (f"[{kind}] {hdr}: {len(rows)} lines -> archived {html_out.name}"
           + (f"  [{len(warns)} warnings]" if warns else ""))
    logfn("✓ " + msg); write_log_file(msg)
    # hand the items to the validation station (scanning)
    # only PICKS reach the validation station - a put-away is a different document
    if on_pick and kind == "PICK":
        try:
            remembered = load_serial_skus()    # SKUs manually marked as serialised
            items = [{"sku": r["item"], "need": int(r["qty"]), "bin": r.get("bin",""),
                      "scanned": 0,
                      "serial": is_radio(r["item"]) or _norm_sku(r["item"]) in remembered}
                     for r in rows if r.get("item")]
            if items:                          # guard: a pick with no lines must never enter the station (zombie document)
                on_pick({"type": kind, "hdr": hdr, "title": title, "sub": sub, "items": items,
                         "printed_ts": time.time()})   # cycle anchor: print to first scan is the real picking time
            else:
                logfn(f"   ! {hdr}: no scannable items — sheet archived, station skipped")
        except Exception as e:
            logfn(f"   ! could not load into station: {e}")
    LAST_PRINT["html"] = str(html_out); LAST_PRINT["hdr"] = hdr   # used by the 'Print again' button
    if cfg.get("auto_print"):
        try:
            tmp_pdf = Path(tempfile.gettempdir()) / f"{hdr}_print_{int(time.time())}.pdf"
            html_to_pdf(str(html_out), str(tmp_pdf))
            ok = print_pdf(str(tmp_pdf), cfg.get("printer") or None)
            time.sleep(3)
            try: tmp_pdf.unlink()
            except Exception: pass
            logfn(f"   🖨 printed to {cfg.get('printer') or 'default'}" if ok else "   ! print failed")
            write_log_file(f"    printed={ok} printer={cfg.get('printer') or 'default'}")
        except Exception as e:
            logfn(f"   ! auto-print failed ({e}) — opening in browser for manual print")
            write_log_file(f"    auto-print FAILED: {e} (opened in browser)")
            try: webbrowser.open(html_out.as_uri())
            except Exception: pass
    else:
        webbrowser.open(html_out.as_uri())
    return True, hdr

# ==================================================================== GUI
def _sweep_stale_mei():
    """Usuwa osierocone katalogi _MEI* poprzednich instancji (taskkill /F, race z AV).
       Wlasny sys._MEIPASS pomijany. Ciche - brak uprawnien = trudno."""
    try:
        import tempfile, glob
        cur = os.path.normcase(getattr(sys, "_MEIPASS", "") or "")
        for d in glob.glob(os.path.join(tempfile.gettempdir(), "_MEI*")):
            if cur and os.path.normcase(d) == cur: continue
            shutil.rmtree(d, ignore_errors=True)
    except Exception:
        pass

def run_gui():
    _sweep_stale_mei()
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog

    cfg = load_cfg(); customers = load_customers()
    seen = set(); scan_state = {"last": 0.0}

    # AppUserModelID BEFORE the window is created - otherwise Windows groups it under python.exe
    # and shows the default taskbar icon instead of ours.
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("PickCore.PickCore.PickConverter.1")
    except Exception:
        pass

    root = tk.Tk(); root.title(APP_TITLE); root.geometry("1080x800"); root.configure(bg="#1a1714")
    root.minsize(1040, 740)  # the rail plus the bottom bar must stay visible at the default window size
    try:
        ico = resource_path("pickcore.ico")
        root.iconbitmap(default=ico)          # default= sets the icon for the window AND the taskbar
        root.wm_iconbitmap(ico)
    except Exception:
        pass

    # ---- a single colour system (design tokens) ----
    UI = {"bg":"#1a1714", "panel":"#241f1a", "panel2":"#2e2720", "border":"#45392c",
          "text":"#f2ece1", "muted":"#a99a84", "faint":"#7a6c58",
          "accent":"#2fb5a8", "ok":"#46d17f", "warn":"#f5b342", "err":"#ff5a5f", "serial":"#cf7fbf",
          "ok_bg":"#13301f", "warn_bg":"#332813", "err_bg":"#3a1416", "todo_bg":"#2a241e"}

    style = ttk.Style(); style.theme_use("clam")
    # Treeview - taller rows, clean typography, dark background
    style.configure("Pick.Treeview", background=UI["panel"], fieldbackground=UI["panel"],
                    foreground=UI["text"], rowheight=34, borderwidth=0, font=("Bahnschrift",11))
    style.configure("Pick.Treeview.Heading", background=UI["panel2"], foreground=UI["muted"],
                    relief="flat", font=("Bahnschrift",9,"bold"), padding=(6,6))
    style.map("Pick.Treeview.Heading", background=[("active", UI["panel2"])])
    style.map("Pick.Treeview", background=[("selected", "#243044")], foreground=[("selected", UI["text"])])
    # Progressbar - thicker, in the accent colour
    style.configure("Pick.Horizontal.TProgressbar", troughcolor=UI["panel2"], background=UI["ok"],
                    borderwidth=0, thickness=18)
    # ================== COCKPIT SHELL: rail + header + status bar ==================
    class SideNav(tk.Frame):
        """Cockpit shell v1.1: zwijany rail ikonowy + pasek kontekstu nad trescia.
           Rail startuje waski (same ikony, etykieta w dymku), przypinany klikiem
           w logo. API zgodne z uzywanym podzbiorem ttk.Notebook (add/insert/select/
           tab/alert + wirtualne <<NotebookTabChanged>>), wiec zawartosc zakladek
           requires no changes at all."""
        GROUPS = ("OUTBOUND", "INBOUND", "OPERATIONS", "LABELS", "INSIGHTS", "SYSTEM")
        W_MIN, W_MAX, ROW_H = 60, 210, 38
        def __init__(self, master):
            super().__init__(master, bg=UI["bg"])
            self.open = False
            self._tip = None
            self.rail = tk.Frame(self, bg=UI["panel"], width=self.W_MIN)
            self.rail.pack(side="left", fill="y"); self.rail.pack_propagate(False)
            tk.Frame(self, bg=UI["border"], width=1).pack(side="left", fill="y")
            right = tk.Frame(self, bg=UI["bg"]); right.pack(side="left", fill="both", expand=True)
            # Context bar: with the rail collapsed this is the only place showing
            # the full name of the active view. Without it the icons are a guessing game.
            self.ctx = tk.Frame(right, bg=UI["bg"], height=36)
            self.ctx.pack(side="top", fill="x"); self.ctx.pack_propagate(False)
            self.ctx_ico = tk.Label(self.ctx, text="", fg=UI["accent"], bg=UI["bg"],
                                    font=("Bahnschrift", 13))
            self.ctx_ico.pack(side="left", padx=(16, 7))
            self.ctx_lbl = tk.Label(self.ctx, text="", fg=UI["text"], bg=UI["bg"],
                                    font=("Bahnschrift", 12, "bold"))
            self.ctx_lbl.pack(side="left")
            self.ctx_key = tk.Label(self.ctx, text="", fg=UI["faint"], bg=UI["bg"],
                                    font=("Cascadia Mono", 8, "bold"))
            self.ctx_key.pack(side="left", padx=9)
            tk.Frame(right, bg=UI["border"], height=1).pack(side="top", fill="x")
            self.body = tk.Frame(right, bg=UI["bg"]); self.body.pack(side="top", fill="both", expand=True)
            self._items = []; self._cur = None
        # ---- item registration (unchanged API) ----
        def _nav_register(self, pos, kind, frame, text, group, fkey, icon, cmd=None):
            it = {"kind": kind, "frame": frame, "text": text.strip(), "group": group,
                  "fkey": fkey, "icon": icon, "cmd": cmd}
            self._items.insert(len(self._items) if pos is None else pos, it)
            if frame is not None:
                frame.place(in_=self.body, x=0, y=0, relwidth=1.0, relheight=1.0)
            self._rebuild()
            if kind == "view" and self._cur is None: self.select(frame)
        def add(self, child, text="", group="SYSTEM", fkey=None, icon=""):
            self._nav_register(None, "view", child, text, group, fkey, icon)
        def insert(self, pos, child, text="", group="OPERATIONS", fkey=None, icon=""):
            self._nav_register(pos, "view", child, text, group, fkey, icon)
        def add_action(self, text, command, group="OPERATIONS", fkey=None, icon=""):
            self._nav_register(None, "action", None, text, group, fkey, icon, cmd=command)

        # ---- collapsing ----
        def toggle(self):
            self.open = not self.open
            self.rail.config(width=(self.W_MAX if self.open else self.W_MIN))
            self._hide_tip(); self._rebuild()

        def _show_tip(self, it):
            """Tooltip with the item name while the rail is collapsed. An undecorated Toplevel, because
               a label inside the rail would stretch it to the width of the text."""
            if self.open or not it.get("row"): return
            self._hide_tip()
            try:
                x = self.rail.winfo_rootx() + self.W_MIN + 6
                y = it["row"].winfo_rooty() + 6
                t = tk.Toplevel(self); t.wm_overrideredirect(True); t.wm_geometry(f"+{x}+{y}")
                tk.Label(t, text=f'  {it["text"]}  {it["fkey"] or ""}  ', bg=UI["panel2"],
                         fg=UI["text"], font=("Bahnschrift", 9, "bold"),
                         bd=1, relief="solid", pady=3).pack()
                self._tip = t
            except Exception:
                self._tip = None

        def _hide_tip(self):
            if self._tip is not None:
                try: self._tip.destroy()
                except Exception: pass
                self._tip = None
        # ---- rail rendering ----
        def _rebuild(self):
            for w in self.rail.winfo_children(): w.destroy()
            brand = tk.Frame(self.rail, bg=UI["panel"], height=46, cursor="hand2")
            brand.pack(fill="x"); brand.pack_propagate(False)
            mark = tk.Label(brand, text="\u25E7", fg=UI["accent"], bg=UI["panel"],
                            font=("Bahnschrift", 15, "bold"))
            mark.pack(side="left", padx=(20, 0) if not self.open else (18, 8))
            wid = [brand, mark]
            if self.open:
                nm = tk.Label(brand, text="PickCore", fg=UI["text"], bg=UI["panel"],
                              font=("Bahnschrift", 12, "bold"))
                nm.pack(side="left"); wid.append(nm)
            for w in wid:
                w.bind("<Button-1>", lambda e: self.toggle())
            tk.Frame(self.rail, bg=UI["border"], height=1).pack(fill="x")
            for g in self.GROUPS:
                grp = [it for it in self._items if it["group"] == g]
                if not grp: continue
                if self.open:
                    tk.Label(self.rail, text=g, fg=UI["faint"], bg=UI["panel"],
                             font=("Bahnschrift", 7, "bold"), anchor="w").pack(fill="x", padx=14, pady=(11, 2))
                else:
                    tk.Frame(self.rail, bg=UI["panel"], height=7).pack(fill="x")
                    tk.Frame(self.rail, bg=UI["border"], height=1).pack(fill="x", padx=16)
                    tk.Frame(self.rail, bg=UI["panel"], height=5).pack(fill="x")
                for it in grp:
                    self._row(it)
            self._restyle()

        def _row(self, it):
            row = tk.Frame(self.rail, bg=UI["panel"], cursor="hand2", height=self.ROW_H)
            row.pack(fill="x", pady=1); row.pack_propagate(False)
            bar = tk.Frame(row, bg=UI["panel"], width=3); bar.pack(side="left", fill="y")
            ico = tk.Label(row, text=it["icon"] or "\u25AB", fg=UI["muted"], bg=UI["panel"],
                           font=("Bahnschrift", 13), width=2)
            ico.pack(side="left", padx=(11, 0))
            lbl = tk.Label(row, text=it["text"], fg=UI["muted"], bg=UI["panel"],
                           font=("Bahnschrift", 10, "bold"), anchor="w")
            key = tk.Label(row, text=it["fkey"] or "", fg=UI["faint"], bg=UI["panel"],
                           font=("Cascadia Mono", 8), anchor="e")
            dot = tk.Label(row, text="", fg="#ff5c5c", bg=UI["panel"], font=("Bahnschrift", 11, "bold"))
            if self.open:
                lbl.pack(side="left", fill="x", expand=True, padx=(8, 2))
                key.pack(side="right", padx=(0, 10))
                dot.pack(side="right", padx=(0, 2))
            else:
                dot.place(relx=0.80, rely=0.18)
            it.update(row=row, bar=bar, ico=ico, lbl=lbl, key=key, dot=dot)
            if it.get("alert"): dot.config(text="\u25CF")
            for w in (row, bar, ico, lbl, key):
                w.bind("<Button-1>", lambda e, it=it: self._activate(it))
                w.bind("<Enter>", lambda e, it=it: (self._hover(it, True), self._show_tip(it)))
                w.bind("<Leave>", lambda e, it=it: (self._hover(it, False), self._hide_tip()))
        # ---- visual state ----
        def _is_on(self, it): return it["kind"] == "view" and it["frame"] is self._cur
        def _hover(self, it, on):
            if self._is_on(it) or "row" not in it: return
            bg = UI["panel2"] if on else UI["panel"]
            for k in ("row", "lbl", "key", "bar", "ico", "dot"):
                try: it[k].config(bg=bg)
                except Exception: pass
        def _restyle(self):
            for it in self._items:
                if "row" not in it: continue
                on = self._is_on(it)
                bg = UI["panel2"] if on else UI["panel"]
                fg = UI["accent"] if on else UI["muted"]
                for k in ("row", "lbl", "key", "bar", "ico", "dot"):
                    try: it[k].config(bg=bg)
                    except Exception: pass
                it["bar"].config(bg=(UI["accent"] if on else bg))
                it["lbl"].config(fg=fg); it["ico"].config(fg=fg)
                it["key"].config(fg=(UI["accent"] if on else UI["faint"]))
                if on:
                    self.ctx_ico.config(text=it["icon"] or "")
                    self.ctx_lbl.config(text=it["text"])
                    self.ctx_key.config(text=(it["fkey"] or ""))
        def alert(self, text, on=True):
            """A red dot on a nav item: something arrived in a view you are not currently in."""
            for it in self._items:
                if it["text"] == text:
                    it["alert"] = bool(on)
                    if "dot" in it:
                        try: it["dot"].config(text=("\u25CF" if on else ""))
                        except Exception: pass
                    break
        def _activate(self, it):
            if it["kind"] == "action":
                if it["cmd"]: it["cmd"]()
                return
            self.select(it["frame"])
        def select(self, child=None):
            if child is None:
                return str(self._cur) if self._cur is not None else ""
            if isinstance(child, str):
                child = next((it["frame"] for it in self._items
                              if it["kind"] == "view" and str(it["frame"]) == child), None)
                if child is None: return
            self._cur = child; child.lift()
            for it in self._items:
                if it["kind"] == "view" and it["frame"] is child and it.get("alert"):
                    self.alert(it["text"], False)
            self._restyle()
            self.event_generate("<<NotebookTabChanged>>")
        def tab(self, tab_id, option=None):
            tid = tab_id if isinstance(tab_id, str) else str(tab_id)
            for it in self._items:
                if it["kind"] == "view" and str(it["frame"]) == tid:
                    return it["text"] if option == "text" else dict(text=it["text"])
            return "" if option == "text" else {}
        def bind_fkeys(self, top):
            for it in self._items:
                if it.get("fkey"):
                    top.bind(f'<{it["fkey"]}>', lambda e, it=it: self._activate(it))
            top.bind("<Control-b>", lambda e: self.toggle())

    # --- HEADER (global): brand | current pick | watcher | printer | operator ---
    hdr = tk.Frame(root, bg=UI["panel"]); hdr.pack(side="top", fill="x")
    tk.Frame(root, bg=UI["border"], height=1).pack(side="top", fill="x")
    tk.Label(hdr, text="PickCore", fg=UI["accent"], bg=UI["panel"],
             font=("Bahnschrift",13,"bold")).pack(side="left", padx=(14,3), pady=10)
    tk.Label(hdr, text=APP_VERSION, fg=UI["faint"], bg=UI["panel"],
             font=("Cascadia Mono",8,"bold")).pack(side="left", padx=(0,8), pady=(16,6))
    hdr_op = tk.Label(hdr, text=f"⛭ {get_picker()}", fg=UI["muted"], bg=UI["panel"], font=("Cascadia Mono",9,"bold"))
    hdr_op.pack(side="right", padx=(6,14))
    hdr_prn = tk.Label(hdr, text="", fg=UI["muted"], bg=UI["panel"], font=("Cascadia Mono",9))
    hdr_prn.pack(side="right", padx=6)
    hdr_led = tk.Label(hdr, text="● watcher", fg=UI["faint"], bg=UI["panel"], font=("Cascadia Mono",9,"bold"))
    hdr_led.pack(side="right", padx=6)
    def _hdr_prn_refresh():
        hdr_prn.config(text=f'🖨 {cfg.get("printer") or "(default)"}')
    _hdr_prn_refresh()
    hdr_mid = tk.Frame(hdr, bg=UI["panel"]); hdr_mid.pack(side="left", fill="both", expand=True, padx=(14,8))

    # --- STATUS BAR (global): last event plus shortcuts ---
    sbar = tk.Frame(root, bg=UI["panel"]); sbar.pack(side="bottom", fill="x")
    tk.Frame(root, bg=UI["border"], height=1).pack(side="bottom", fill="x")
    status_lbl = tk.Label(sbar, text="Ready.", fg=UI["muted"], bg=UI["panel"],
                          font=("Cascadia Mono",9), anchor="w")
    status_lbl.pack(side="left", fill="x", expand=True, padx=12, pady=4)
    tk.Label(sbar, text="F1-F9 — switch views", fg=UI["faint"], bg=UI["panel"],
             font=("Bahnschrift",8)).pack(side="right", padx=12)

    live_ring = []; cloud_state = {"dirty": True}   # feeds the synced-folder live board (hooked into logln)
    web_alerts = {"pick": 0, "inbound": 0}          # notification counters for the handheld console
    hdr_scn = tk.Label(hdr, text="", fg=UI["faint"], bg=UI["panel"], font=("Cascadia Mono",9,"bold"))
    upd_state = {"ver": "", "info": ""}
    hdr_upd = tk.Label(hdr, text="", fg=UI["warn"], bg=UI["panel"], font=("Bahnschrift",9,"bold"), cursor="hand2")
    nb = SideNav(root); nb.pack(fill="both", expand=True)

    # ---------- TAB: PICK STATION ----------
    import queue as _queue
    station = {"cur": None, "queue": []}
    bulk = {"on": False, "idx": None, "step": None}
    serial_mode = {"on": False, "radios": [], "active": None, "seen": set(), "history": []}
    new_picks = _queue.Queue()
    log_q = _queue.Queue()          # logs from worker threads reach the main thread through a queue (Tk is not thread safe)
    flagged = {"wrong": set(), "noise": set(), "over": set()}   # dedupe zdarzen per pick: 3x ten sam zly kod = 1 zdarzenie
    pick_t = {"t0": None, "scans": []}   # telemetria picka: (t_offset, sku, bin, qty) - dane do slottingu, BEZ operatora
    session = {"on": False, "items": [], "active": None, "seen": set(), "await": None, "suggest": 0}
    pa = {"on": False, "po": "", "doc": "", "lines": [], "active": None, "await": None}
    def _pa_path():
        _,_,ld = ensure_app_folders(); return os.path.join(ld, "pa_session_active.json")
    def pa_save():
        try:
            Path(_pa_path()).write_text(json.dumps({k: pa[k] for k in ("po","doc","lines","await","active")},
                ensure_ascii=False), encoding="utf-8")
        except Exception as e:
            # A silent failure means the session comes back EMPTY after a restart while the operator believes it was saved.
            note_io_fail("pa_save", e)
    def pa_clear_file(archive=False):
        try:
            p=_pa_path()
            if os.path.exists(p):
                if archive and pa["lines"]:
                    shutil.copy2(p, p.replace("active", datetime.now().strftime("%Y%m%d_%H%M%S")))
                os.remove(p)
        except Exception: pass
    def _sess_path():
        _,_,ld = ensure_app_folders(); return os.path.join(ld, "serial_session_active.json")
    def sess_save():
        try:
            Path(_sess_path()).write_text(json.dumps({"items": session["items"], "await": session["await"],
                "active": session["active"]}, ensure_ascii=False), encoding="utf-8")
        except Exception as e:
            note_io_fail("sess_save", e)
    def sess_clear_file(archive=False):
        try:
            p = _sess_path()
            if os.path.exists(p):
                if archive and session["items"]:
                    shutil.copy2(p, p.replace("active", datetime.now().strftime("%Y%m%d_%H%M%S")))
                os.remove(p)
        except Exception: pass

    def _beep(kind):
        """Distinguishable sound signatures (winsound.Beep on its own thread, so the GUI never freezes).
           ok=krotki tik | err=niski szorstki | warn=podwojny sredni | line=wznoszacy (item 10/10)
           | done=fanfara C-E-G-C (caly pick 74/74).
           v1.0: ta sama sygnatura gra na TC22 - konsola dostaje (id, kind) w snapshocie
           i odtwarza SEKWENCJE Z BEEP_SEQS (jedno zrodlo prawdy, zero duplikacji nut w JS)."""
        seqs = BEEP_SEQS
        snd_state["id"] += 1; snd_state["kind"] = kind if kind in seqs else "ok"
        # An operator-supplied file takes precedence over the generated tone.
        # winsound handles WAV only - MP3 works exclusively on the handheld console.
        _wav = find_sound(cfg, snd_state["kind"], SOUND_EXT_PC)
        if _wav:
            try:
                import winsound as _ws
                threading.Thread(target=lambda: _ws.PlaySound(_wav, _ws.SND_FILENAME | _ws.SND_ASYNC),
                                 daemon=True).start()
                return
            except Exception:
                pass
        def play():
            try:
                import winsound
                for f, d in seqs.get(kind, seqs["ok"]):
                    winsound.Beep(int(f), int(d))
            except Exception: pass
        threading.Thread(target=play, daemon=True).start()

    t1 = tk.Frame(nb, bg=UI["bg"]); nb.add(t1, text="Pick Station", group="OUTBOUND", fkey="F1", icon="🎯")

    # --- HEADER PICKA: przeniesiony do globalnego cockpit-headera (hdr_mid) - v3.0 ---
    head = tk.Frame(hdr_mid, bg=UI["panel"]); head.pack(fill="both", expand=True, pady=4)
    head_l = tk.Frame(head, bg=UI["panel"]); head_l.pack(side="left", fill="x", expand=True, anchor="w")
    st_title = tk.Label(head_l, text="No active pick", fg=UI["accent"], bg=UI["panel"],
                        font=("Bahnschrift",16,"bold"), anchor="w"); st_title.pack(anchor="w")
    st_sub = tk.Label(head_l, text="Load a pick PDF (watcher or manual) to begin scanning",
                     fg=UI["muted"], bg=UI["panel"], font=("Bahnschrift",8), anchor="w"); st_sub.pack(anchor="w")
    st_badge = tk.Label(head, text="IDLE", fg=UI["faint"], bg=UI["panel2"], font=("Bahnschrift",10,"bold"),
                       padx=12, pady=4); st_badge.pack(side="right", anchor="e", padx=(0,10))

    # --- PROGRESS: gruby pasek + duze X / Y + procent ---
    prog_fr = tk.Frame(t1, bg=UI["bg"]); prog_fr.pack(fill="x", padx=16, pady=(10,2))
    prog = ttk.Progressbar(prog_fr, mode="determinate", maximum=100, style="Pick.Horizontal.TProgressbar")
    prog.pack(side="left", fill="x", expand=True)
    st_prog = tk.Label(prog_fr, text="0 / 0", fg=UI["text"], bg=UI["bg"], font=("Bahnschrift",15,"bold"), width=11)
    st_prog.pack(side="left", padx=(12,0))
    st_queue = tk.Label(t1, text="", fg=UI["warn"], bg=UI["bg"], font=("Bahnschrift",9,"bold"), anchor="w")
    st_queue.pack(fill="x", padx=18)

    # --- LISTA POZYCJI ---
    tree_wrap = tk.Frame(t1, bg=UI["border"]); tree_wrap.pack(fill="both", expand=True, padx=16, pady=8)
    itree = ttk.Treeview(tree_wrap, columns=("st","sku","bin","cnt"), show="headings", height=8, style="Pick.Treeview")
    _ysb = ttk.Scrollbar(tree_wrap, orient="vertical", command=itree.yview)
    itree.configure(yscrollcommand=_ysb.set); _ysb.pack(side="right", fill="y")
    for c,(h_,w_,anch) in {"st":("",46,"center"),"sku":("ITEM",300,"w"),"bin":("LOCATION",130,"center"),"cnt":("DONE / NEED",130,"center")}.items():
        itree.heading(c,text=h_); itree.column(c,width=w_,anchor=anch)
    itree.tag_configure("done", background=UI["ok_bg"], foreground=UI["ok"])
    itree.tag_configure("part", background=UI["warn_bg"], foreground=UI["warn"])
    itree.tag_configure("todo", background=UI["todo_bg"], foreground=UI["text"])
    itree.pack(fill="both", expand=True, padx=1, pady=1)

    # --- PASEK SKANU (wyrozniony panel) ---
    scan_fr = tk.Frame(t1, bg=UI["panel"], highlightbackground=UI["border"], highlightthickness=1)
    scan_fr.pack(fill="x", padx=16, pady=(4,4))
    tk.Label(scan_fr, text="▶ SCAN", fg=UI["ok"], bg=UI["panel"], font=("Bahnschrift",12,"bold")).pack(side="left", padx=(10,4), pady=8)
    scan_var = tk.StringVar()
    scan_entry = tk.Entry(scan_fr, textvariable=scan_var, font=("Cascadia Mono",17,"bold"), bg="#120f0d",
                          fg=UI["text"], insertbackground=UI["ok"], relief="flat")
    scan_entry.pack(side="left", fill="x", expand=True, padx=8, ipady=6)
    # pomiar predkosci: SPRZETOWY timestamp zdarzenia (event.time, ms) - odporny na lag petli Tk
    # (previously: handler execution time, which inflated whenever the main thread was busy and caused false scan rejects)
    scan_keys = {"times": []}
    _dbg = {"n": 0, "mx": 0, "md": 0}
    _CTRL = ("Return","Tab","Shift_L","Shift_R","Control_L","Control_R","Alt_L","Alt_R","Caps_Lock",
             "Left","Right","Up","Down","BackSpace","Delete","Home","End")
    def _on_keypress(event):
        if event.keysym in _CTRL: return
        scan_keys["times"].append(int(getattr(event, "time", 0) or 0))
    scan_entry.bind("<KeyPress>", _on_keypress, add="+")
    def _qty_was_scanned():
        ok, n, mx, md = qty_speed_ok(scan_keys["times"])
        _dbg["n"], _dbg["mx"], _dbg["md"] = n, mx, md
        return ok
    btn_bulk = tk.Button(scan_fr, text="BULK", bg=UI["serial"], fg="white", relief="flat", font=("Bahnschrift",10,"bold"), width=7, cursor="hand2")
    btn_bulk.pack(side="left", padx=3, pady=6)
    btn_manual = tk.Button(scan_fr, text="Manual", bg="#5a4a38", fg="white", relief="flat", font=("Bahnschrift",10), width=8, cursor="hand2")
    btn_manual.pack(side="left", padx=3, pady=6)
    btn_serial = tk.Button(scan_fr, text="🔖 Serial", bg=UI["accent"], fg="white", relief="flat", font=("Bahnschrift",10), width=9, cursor="hand2")
    btn_serial.pack(side="left", padx=(3,10), pady=6)
    # serial mode buttons, hidden until that mode is entered
    btn_back = tk.Button(scan_fr, text="← Pick", bg=UI["panel2"], fg=UI["text"], relief="flat", font=("Bahnschrift",10,"bold"), width=7, cursor="hand2")
    btn_copy = tk.Button(scan_fr, text="📋 Copy serials", bg=UI["ok"], fg="white", relief="flat", font=("Bahnschrift",10,"bold"), width=14, cursor="hand2")
    btn_undo = tk.Button(scan_fr, text="↶ Undo", bg=UI["serial"], fg="white", relief="flat", font=("Bahnschrift",10,"bold"), width=8, cursor="hand2")
    btn_finish = tk.Button(scan_fr, text="✓ Finish pick", bg="#FF7B00", fg="white", relief="flat", font=("Bahnschrift",10,"bold"), width=13, cursor="hand2")
    # button into serial capture, shown during validation when the pick has serialised items
    btn_to_serials = tk.Button(scan_fr, text="🔢 Serials →", bg=UI["serial"], fg="white", relief="flat", font=("Bahnschrift",10,"bold"), width=11, cursor="hand2")

    # --- BANER FEEDBACKU: duzy, zmienia tlo wg stanu (sygnal z 2m) ---
    fb = tk.Label(t1, text="Ready to scan.", fg=UI["muted"], bg=UI["panel"], font=("Bahnschrift",16,"bold"),
                  anchor="w", padx=14, pady=10)
    fb.pack(fill="x", padx=16, pady=(4,2))
    bulk_hint = tk.Label(t1, text="", fg=UI["serial"], bg=UI["bg"], font=("Cascadia Mono",10), anchor="w")
    bulk_hint.pack(fill="x", padx=18)
    # pack mode selector for the ACTIVE unit, visible only in serial mode
    pack_fr = tk.Frame(t1, bg=UI["bg"])
    tk.Label(pack_fr, text="Pack mode:", fg=UI["muted"], bg=UI["bg"], font=("Bahnschrift",9,"bold")).pack(side="left", padx=(0,8))
    PACK_MODES = {"Single":1, "Dual":2, "Quad":4, "6-pack":6}
    pack_btns = {}
    def _highlight_pack(active):
        for size,b in pack_btns.items():
            on = (size == active)
            b.config(bg=(UI["accent"] if on else UI["panel2"]), fg=("white" if on else UI["muted"]))
    def _set_pack(size):
        if session["on"]:
            if session["active"] is None: return
            session["items"][session["active"]]["pack"] = size
            _highlight_pack(size); sess_save(); refresh_station(); scan_entry.focus_set(); return
        if serial_mode["active"] is None: return
        serial_mode["radios"][serial_mode["active"]]["pack"] = size
        _highlight_pack(size); refresh_station(); scan_entry.focus_set()
    for _name,_size in PACK_MODES.items():
        _b = tk.Button(pack_fr, text=_name, relief="flat", font=("Bahnschrift",9,"bold"),
                       bg=UI["panel2"], fg=UI["muted"], width=7, cursor="hand2",
                       command=lambda s=_size: _set_pack(s))
        _b.pack(side="left", padx=2); pack_btns[_size] = _b
    _btm = tk.Frame(t1, bg=UI["bg"]); _btm.pack(pady=(2,8))
    tk.Button(_btm, text="📄  Load PDF manually", command=lambda: manual_load(), bg=UI["panel2"], fg=UI["muted"],
              relief="flat", font=("Bahnschrift",9), cursor="hand2").pack(side="left", padx=4)
    def print_again():
        if not LAST_PRINT["html"] or not os.path.exists(LAST_PRINT["html"]):
            feedback("Nothing to reprint yet — process a pick first.", "warn"); return
        feedback(f"🖨 Reprinting {LAST_PRINT['hdr']}…", "info")
        def job():
            try:
                tmp = Path(tempfile.gettempdir()) / f"reprint_{int(time.time())}.pdf"
                html_to_pdf(LAST_PRINT["html"], str(tmp))
                ok = print_pdf(str(tmp), cfg.get("printer") or None)
                time.sleep(3)
                try: tmp.unlink()
                except Exception: pass
                log_q.put(f"🖨 Reprint {LAST_PRINT['hdr']}: {'ok' if ok else 'FAILED'}")
            except Exception as e:
                log_q.put(f"! Reprint failed: {e}")
        threading.Thread(target=job, daemon=True).start()
    # Picking is scanner-driven now, so reprinting the A4 sheet is an exception, not routine.
    # The button shrinks to icon size: still within reach, but no longer competing
    # for attention with the actions that actually get used. The function itself is unchanged.
    _btn_reprint = tk.Button(_btm, text="🖨", command=print_again, bg=UI["panel2"], fg=UI["faint"],
                             relief="flat", font=("Bahnschrift",9), cursor="hand2", width=3)
    _btn_reprint.pack(side="left", padx=4)
    # An unlabelled icon does not read as clickable, so it is highlighted on hover.
    # Tk has no native tooltips and a Toplevel tooltip is another window object to maintain,
    # not worth a dedicated button. If an operator asks what it is,
    # so a shared hint mechanism for the whole UI comes only once it earns its keep.
    _btn_reprint.bind("<Enter>", lambda e: _btn_reprint.config(fg=UI["accent"]))
    _btn_reprint.bind("<Leave>", lambda e: _btn_reprint.config(fg=UI["faint"]))
    def remove_last_pick():
        """Removes the MOST RECENTLY added pick: the tail of the queue first, losing no work,
           and when the queue is empty it cancels the current one, with a confirmation and no VERIFIED metric."""
        if station["queue"]:
            pd = station["queue"].pop()
            log_event("PICK_CANCELLED", pick=pd.get("title",""), doc=pd.get("hdr",""), scope="queued")
            logln(f"🗑 Removed from queue: {pd.get('title','?')} (doc {pd.get('hdr','')})")
            feedback(f"Removed queued pick {pd.get('title','?')}.", "warn"); _beep("warn")
            refresh_station(); scan_entry.focus_set(); return
        cur = station["cur"]
        if not cur:
            feedback("Nothing to remove.", "warn"); scan_entry.focus_set(); return
        title, hdr = cur["title"], cur["hdr"]
        done = sum(i["scanned"] for i in cur["items"]); need = sum(i["need"] for i in cur["items"])
        nser = sum(len(r["serials"]) for r in serial_mode["radios"]) if serial_mode["radios"] else 0
        extra = (f"\n\nProgress will be lost: {done}/{need} scanned." if done else "")
        if nser: extra += f"\n{nser} collected serial(s) will be discarded."
        if not messagebox.askyesno("Remove pick",
                f"Cancel current pick {title} (doc {hdr})?{extra}\n\nThis will NOT count as verified."):
            scan_entry.focus_set(); return
        log_event("PICK_CANCELLED", pick=title, doc=hdr, scope="current", scanned=done, need=need, serials=nser)
        write_log_file(f"    CANCELLED {title} (doc {hdr}) {done}/{need}")
        logln(f"🗑 Cancelled current pick {title}")
        station["cur"] = None
        serial_mode["on"]=False; serial_mode["radios"]=[]; serial_mode["active"]=None
        serial_mode["seen"]=set(); serial_mode["history"]=[]
        _show_validation_buttons(); bulk_reset()
        pick_t["t0"]=None; pick_t["scans"]=[]; pick_t["sent"]=False
        feedback(f"Pick {title} cancelled.", "warn"); _beep("warn")
        refresh_station()
        if station["queue"]: root.after(400, load_next)
        scan_entry.focus_set()
    tk.Button(_btm, text="🗑  Remove last pick", command=remove_last_pick, bg=UI["panel2"], fg=UI["muted"],
              relief="flat", font=("Bahnschrift",9), cursor="hand2").pack(side="left", padx=4)
    btn_session = tk.Button(_btm, text="⚡  Serial session", command=lambda: toggle_session(), bg=UI["panel2"], fg=UI["muted"],
              relief="flat", font=("Bahnschrift",9), cursor="hand2"); btn_session.pack(side="left", padx=4)

    def _set_badge(text, kind):
        cmap = {"idle":(UI["faint"],UI["panel2"]), "scan":(UI["accent"],"#16243a"),
                "serial":(UI["serial"],"#241a3a"), "ok":(UI["ok"],UI["ok_bg"])}
        fg,bg = cmap.get(kind, (UI["faint"],UI["panel2"]))
        st_badge.config(text=text, fg=fg, bg=bg)

    def feedback(text, kind="info"):
        col = {"ok":UI["ok"],"err":UI["err"],"warn":UI["warn"],"info":UI["muted"],"done":UI["ok"]}.get(kind, UI["muted"])
        bg  = {"ok":UI["ok_bg"],"err":UI["err_bg"],"warn":UI["warn_bg"],"done":UI["ok_bg"]}.get(kind, UI["panel"])
        fg  = UI["text"] if kind in ("ok","err","warn","done") else col
        fb.config(text=text, fg=fg, bg=bg)
        if kind == "err":
            scan_entry.config(bg=UI["err_bg"])
            t1.after(350, lambda: scan_entry.config(bg="#120f0d"))

    def refresh_station():
        if session["on"]:
            refresh_session(); return
        itree.delete(*itree.get_children())
        # --- TRYB ZBIERANIA SERIALI (po weryfikacji picka z radiami) ---
        if serial_mode["on"] and station["cur"]:
            cur = station["cur"]; radios = serial_mode["radios"]
            _set_badge("● SERIALS", "serial")
            st_title.config(text=f"{cur['title']}")
            st_sub.config(text="Serial collection · click a radio to switch · Copy to clipboard · Finish when done")
            done = sum(min(len(r["serials"]), r["qty"]*r.get("pack",1)) for r in radios)
            need = sum(r["qty"]*r.get("pack",1) for r in radios)
            prog.config(value=(done/need*100 if need else 0)); st_prog.config(text=f"{done} / {need}")
            st_queue.config(text=(f"Queue: {len(station['queue'])} pick(s) waiting" if station["queue"] else ""))
            for idx, r in enumerate(radios):
                pk = r.get("pack", 1)
                target = r["qty"] * pk                 # ile seriali razem (pudelka x pack)
                full = len(r["serials"]) >= target
                tag = "done" if full else ("part" if r["serials"] else "todo")
                mark = "✓" if full else ("▶" if idx == serial_mode["active"] else "")
                pid = f"r{idx}"
                loc = f"{r['qty']} box" if pk <= 1 else f"{r['qty']} box · {pk}x"
                itree.insert("","end",iid=pid, open=(idx==serial_mode["active"]),
                    values=(mark, r["sku"], loc, f'{len(r["serials"])}/{target}'), tags=(tag,))
                # serials: for packs larger than one show the box number, otherwise the line index
                for j,s in enumerate(r["serials"]):
                    label = f"#{j+1}" if pk <= 1 else f"box {j//pk + 1}"
                    itree.insert(pid,"end",iid=f"{pid}_s{j}", values=("", s, "", label), tags=("todo",))
            if serial_mode["active"] is not None:
                try: itree.selection_set(f"r{serial_mode['active']}")
                except Exception: pass
            return
        cur = station["cur"]
        if not cur:
            _set_badge("IDLE", "idle")
            st_title.config(text="No active pick"); st_sub.config(text="Load a pick PDF (watcher or manual) to begin scanning")
            prog.config(value=0); st_prog.config(text="0 / 0")
            st_queue.config(text=(f"Queue: {len(station['queue'])} pick(s) waiting" if station["queue"] else ""))
            if btn_to_serials.winfo_ismapped(): btn_to_serials.pack_forget()
            return
        _set_badge("● SCANNING", "scan")
        st_title.config(text=f"{cur['title']}")
        st_sub.config(text=f"{cur['type']} · {cur['sub']}")
        done_u = sum(min(i["scanned"],i["need"]) for i in cur["items"])
        need_u = sum(i["need"] for i in cur["items"])
        prog.config(value=(done_u/need_u*100 if need_u else 0))
        st_prog.config(text=f"{done_u} / {need_u}")
        st_queue.config(text=(f"Queue: {len(station['queue'])} pick(s) waiting" if station["queue"] else ""))
        for idx,i in enumerate(cur["items"]):
            full = i["scanned"] >= i["need"]
            tag = "done" if full else ("part" if i["scanned"]>0 else "todo")
            mark = "✓" if full else ("▶" if i["scanned"]>0 else "")
            sku_disp = ("🔖 " + i["sku"]) if i.get("serial") else i["sku"]
            itree.insert("","end",iid=str(idx),
                values=(mark, sku_disp, i["bin"] or "—", f'{i["scanned"]}/{i["need"]}'), tags=(tag,))
        # the "Serials →" button appears when the pick holds serialised items and serial mode is off
        has_ser = any(i.get("serial") for i in cur["items"])
        if has_ser and not btn_to_serials.winfo_ismapped():
            btn_to_serials.pack(side="left", padx=(10,3), pady=6)
        elif not has_ser and btn_to_serials.winfo_ismapped():
            btn_to_serials.pack_forget()

    def bulk_reset():
        bulk["on"]=False; bulk["idx"]=None; bulk["step"]=None; bulk["retry"]=None
        bulk_hint.config(text=""); btn_bulk.config(text="BULK", bg=UI["serial"])

    def load_next():
        if session["on"]: return
        if station["cur"] is None and station["queue"]:
            # a new pick resets serial mode: the previous data is no longer valid
            serial_mode["on"]=False; serial_mode["radios"]=[]; serial_mode["active"]=None
            serial_mode["seen"]=set(); serial_mode["history"]=[]
            flagged["wrong"]=set(); flagged["noise"]=set(); flagged["over"]=set()
            pick_t["t0"] = None; pick_t["scans"] = []; pick_t["sent"] = False
            _show_validation_buttons()
            station["cur"] = station["queue"].pop(0); bulk_reset(); refresh_station()
            feedback(f"Loaded {station['cur']['title']} — scan items to verify.", "info")
        else:
            refresh_station()

    def _pick_alert():
        """Red dot plus counter for the console when a pick arrives while the operator is elsewhere."""
        try:
            if nb.select() != str(t1): nb.alert("Pick Station", True)
            web_alerts["pick"] = web_alerts.get("pick", 0) + 1
        except Exception: pass
    def on_pick_arrived(pd):
        pd["queued"] = not (station["cur"] is None and not station["queue"])  # czekal na stacje = cykl zawyzony
        station["queue"].append(pd)
        _pick_alert()
        if station["cur"] is None: load_next()
        else: refresh_station()

    def drain_picks():
        try:
            while True: on_pick_arrived(new_picks.get_nowait())
        except _queue.Empty: pass
        root.after(300, drain_picks)

    def _find_line(code):
        cur = station["cur"]; c = _norm_sku(code)
        # pass 1: dokladne dopasowanie (najbezpieczniejsze)
        for idx,i in enumerate(cur["items"]):
            if _norm_sku(i["sku"]) == c and i["scanned"] < i["need"]: return idx
        # pass 2: tolerancyjne (prefiks/sufiks B z kodu kreskowego)
        for idx,i in enumerate(cur["items"]):
            if sku_match(code, i["sku"]) and i["scanned"] < i["need"]: return idx
        return None
    def _sku_exists(code):
        return any(sku_match(code, i["sku"]) for i in station["cur"]["items"])

    def start_bulk():
        cur = station["cur"]
        if bulk["on"]:                       # toggle: ponowne klikniecie wylacza bulk
            bulk_reset(); feedback("Bulk mode off.", "info"); scan_entry.focus_set(); return
        if not cur: feedback("No active pick.", "warn"); return
        sel = itree.selection()
        if not sel: feedback("Select the line first, then press BULK.", "warn"); return
        try: idx = int(sel[0])
        except (ValueError, IndexError): return
        line = cur["items"][idx]
        if line["scanned"] >= line["need"]: feedback("That line is already complete.", "warn"); return
        bulk["on"]=True; bulk["idx"]=idx; bulk["step"]="sku"
        btn_bulk.config(text="BULK ●", bg="#cf7fbf")
        bulk_hint.config(text=f"BULK on {line['sku']}   →   step 1/2: scan the BOX SKU  (press BULK again to exit)")
        feedback(f"BULK mode: scan box SKU for {line['sku']}", "info"); scan_entry.focus_set()

    def handle_bulk(code, was_scanned):
        line = station["cur"]["items"][bulk["idx"]]
        if bulk["step"] == "sku":
            if sku_match(code, line["sku"]):
                bulk["step"]="qty"
                bulk_hint.config(text=f"BULK on {line['sku']}   →   step 2/2: scan the QTY barcode (e.g. Q30)")
                feedback("Box SKU OK. Now scan the QTY barcode.", "ok"); _beep("ok")
            else:
                feedback(f"Box SKU mismatch — expected {line['sku']}, got {code}", "err"); _beep("err")
        elif bulk["step"] == "qty":
            qty = parse_scanned_qty(code)
            # LOCK 1 (format): a quantity code must be a number (Q30 / qty20 / 30)
            if qty is None:
                feedback(f"'{code}' is not a QTY code. Scan the box qty label (e.g. Q30 or 30).", "err"); _beep("err"); return
            # RELIEF VALVE: the same code rescanned within 8s of a speed reject counts as label confirmation
            retry = bulk.get("retry")
            double_ok = bool(retry and retry["code"] == code and (time.time() - retry["ts"]) <= 8.0)
            # LOCK 2 (predkosc): sprzetowe timestampy zdarzen; odrzut loguje realne ms do kalibracji
            if not was_scanned and not double_ok:
                bulk["retry"] = {"code": code, "ts": time.time()}
                logln(f"⚠ QTY speed-reject: code={code} n={_dbg['n']} max={_dbg['mx']}ms med={_dbg['md']}ms (limit {QTY_SPEED_MAX_GAP_MS}ms)")
                log_event("QTY_SPEED_REJECT", code=code, n=_dbg["n"], max_ms=_dbg["mx"], med_ms=_dbg["md"])
                feedback("⚠ Quantity must be SCANNED — scan the SAME label again to confirm.", "err"); _beep("err"); return
            if double_ok and not was_scanned:
                logln(f"✓ QTY accepted via double-scan confirm: {code}")
                log_event("QTY_DOUBLE_CONFIRM", code=code)
            bulk["retry"] = None
            remaining = line["need"] - line["scanned"]
            if qty > remaining:
                feedback(f"BLOCKED: box has {qty}, only {remaining} needed. Bulk off — scan remainder as singles.", "err"); _beep("err")
                log_event("BULK_BLOCKED", sku=line["sku"], scanned_qty=qty, remaining=remaining,
                          pick=station["cur"]["title"], doc=station["cur"]["hdr"])
                bulk_reset(); refresh_station(); return
            ok = messagebox.askokcancel("Verify box quantity",
                f"VERIFY PHYSICALLY:\n\nThe box must contain exactly the label quantity:  {qty}\n\n"
                f"Count or weigh if unsure.\n\nPress OK only if the box truly holds {qty}.")
            if ok:
                line["scanned"] += qty
                learn_bulk(line["sku"], qty)          # profil bulkow uczy sie z produkcji
                _all = all(i["scanned"] >= i["need"] for i in station["cur"]["items"])
                if _all:      pass
                elif line["scanned"] >= line["need"]: _beep("line")
                else:         _beep("ok")
                _now = time.time()
                if pick_t["t0"] is None: pick_t["t0"] = _now
                pick_t["scans"].append((round(_now - pick_t["t0"], 1), line["sku"], line.get("bin",""), qty))
                refresh_station(); check_complete()
                if line["scanned"] >= line["need"]:
                    bulk_reset()                          # linia kompletna -> wyjdz z bulka
                    feedback(f"BULK +{qty} — {line['sku']} complete ({line['scanned']}/{line['need']}).", "ok")
                else:
                    bulk["step"] = "sku"                  # ZOSTAN w bulku na kolejne pudelko
                    bulk_hint.config(text=f"BULK on {line['sku']} — scan NEXT box SKU  ({line['scanned']}/{line['need']}, press BULK to exit)")
                    feedback(f"BULK +{qty}   ·   {line['sku']}   ({line['scanned']}/{line['need']}) — scan next box", "ok")
            else:
                bulk["step"] = "sku"                       # anulowano -> wroc do skanu pudelka
                bulk_hint.config(text=f"BULK on {line['sku']} — cancelled, scan box SKU again")

    pa_q = _queue.Queue()
    # ---------- TAB: INBOUND (editable grid with a BIN field per item) ----------
    t_in = tk.Frame(nb, bg=UI["bg"])   # wpiecie na pozycje #2 w boocie
    pa_hdr = tk.Label(t_in, text="\U0001F4E6 INBOUND", fg=UI["accent"], bg=UI["bg"], font=("Bahnschrift",16,"bold"))
    pa_hdr.pack(anchor="w", padx=16, pady=(12,0))
    pa_sub = tk.Label(t_in, text="", fg=UI["muted"], bg=UI["bg"], font=("Cascadia Mono",9)); pa_sub.pack(anchor="w", padx=16)
    _top = tk.Frame(t_in, bg=UI["bg"]); _top.pack(fill="x", padx=16, pady=(6,2))
    tk.Label(_top, text="PO:", fg=UI["muted"], bg=UI["bg"], font=("Bahnschrift",10,"bold")).pack(side="left")
    pa_po_var = tk.StringVar()
    pa_po_e = tk.Entry(_top, textvariable=pa_po_var, width=18, bg=UI["panel2"], fg=UI["text"],
                       insertbackground=UI["text"], relief="flat", font=("Cascadia Mono",12))
    pa_po_e.pack(side="left", padx=8, ipady=3)
    def _commit_po(_=None):
        pa["po"]=pa_po_var.get().strip().upper()
        pa_save(); _pa_counts()
        if pa["po"]:
            log_event("PA_SESSION_OPEN", po=pa["po"], doc=pa["doc"], preloaded=len(pa["lines"]))
            _focus_first_empty()
        return "break"
    pa_po_e.bind("<Return>", _commit_po); pa_po_e.bind("<FocusOut>", lambda e: (_commit_po(), None) and None)
    # scrollable row grid
    _wrap = tk.Frame(t_in, bg=UI["panel"]); _wrap.pack(fill="both", expand=True, padx=16, pady=8)
    _cv = tk.Canvas(_wrap, bg=UI["panel"], highlightthickness=0)
    _sb = ttk.Scrollbar(_wrap, orient="vertical", command=_cv.yview)
    _in = tk.Frame(_cv, bg=UI["panel"])
    _cv_win = _cv.create_window((0,0), window=_in, anchor="nw")
    def _pa_sync_scroll(_=None):
        """Scroll-lock fix: the scrollregion is never smaller than the viewport, so the view
           cannot drift above the content, and when everything fits it stays pinned to the top."""
        try:
            ch, cw = _cv.winfo_height(), _cv.winfo_width()
            iw, ih = _in.winfo_reqwidth(), _in.winfo_reqheight()
            _cv.itemconfigure(_cv_win, width=max(cw, iw))
            _cv.configure(scrollregion=(0, 0, max(cw, iw), max(ih, ch)))
            if ih <= ch: _cv.yview_moveto(0.0)
        except Exception: pass
    _in.bind("<Configure>", _pa_sync_scroll)
    _cv.bind("<Configure>", _pa_sync_scroll)
    _cv.configure(yscrollcommand=_sb.set)
    _cv.pack(side="left", fill="both", expand=True); _sb.pack(side="right", fill="y")
    def _pa_wheel(e):
        if pa["on"] and _in.winfo_reqheight() > _cv.winfo_height():
            _cv.yview_scroll(int(-e.delta/120), "units")
    _cv.bind_all("<MouseWheel>", _pa_wheel)
    pa_rows = []            # [{'e':Entry,'var':StringVar,'idx':int,'lab':Label}]
    pa_pending = []         # v1.0 FIX: kolejka FIFO put-awayow czekajacych (wczesniej: dane przepadaly!)
    bc_q = _queue.Queue()   # wyniki zapytan do BC (watek roboczy -> glowny watek)
    pa_status = tk.Label(t_in, text="Process a put-away to load items, or \u2795 add lines manually.",
                         bg="#241f1a", fg=UI["muted"], font=("Bahnschrift",11,"bold"), anchor="w", padx=12, pady=8)
    pa_status.pack(fill="x", padx=16, pady=(0,4))
    def pa_feedback(msg, kind="info"):
        pa_status.config(text=msg, fg={"ok":UI["ok"],"err":"#ff6b6b","warn":"#f5b342",
                                       "done":UI["ok"],"info":UI["muted"]}.get(kind, UI["muted"]))
    def _pa_counts():
        done=sum(1 for l in pa["lines"] if l.get("bin"))
        pa_hdr.config(text="\U0001F4E6 INBOUND \u00B7 PO " + (pa["po"] or "\u2014")
                      + (("  \u00B7  " + pa["doc"]) if pa["doc"] else ""))
        try: _dn=len(day_load()["entries"])
        except Exception: _dn=0
        pa_sub.config(text=f"{done}/{len(pa['lines'])} bins set \u00B7 day log: {_dn} lines \u00B7 ENTER \u2192 next \u00B7 Ctrl+Z undo")
    def _row_paint(i):
        l=pa["lines"][i]; r=pa_rows[i]
        col = UI["ok"] if l.get("bin") else UI["text"]
        r["lab"].config(fg=col); r["e"].config(fg=col)
    def _focus_bin(i):
        if 0 <= i < len(pa_rows):
            e=pa_rows[i]["e"]; e.focus_set(); e.icursor("end"); e.select_range(0,"end")
            _cv.yview_moveto(max(0, (i-3)/max(1,len(pa_rows))))
    def _focus_first_empty():
        for i,l in enumerate(pa["lines"]):
            if not l.get("bin"): _focus_bin(i); return
        if pa_rows: _focus_bin(len(pa_rows)-1)
    def _copy_if_filled(i):
        l=pa["lines"][i]
        if l.get("bin"):
            try: root.clipboard_clear(); root.clipboard_append(l["bin"]); root.update()
            except Exception: pass
            pa_feedback(f"\U0001F4CB {l['bin']} \u2192 clipboard (line {i+1}) \u2014 paste into BC.","done")
    def _commit_bin(i, jump=True):
        var=pa_rows[i]["var"]; raw=var.get().strip()
        l=pa["lines"][i]
        if not raw:
            if l.get("bin"): l["bin"]=""; pa_save(); _row_paint(i); _pa_counts()
            if jump: _focus_bin(i+1) if i+1<len(pa_rows) else None
            return "break"
        fmt, okb = format_bin_input(raw)
        if not okb:
            pa_feedback(f"'{raw}' \u2192 invalid bin (ZZ-RPP-LS, e.g. 30-A03-A2).","err"); _beep("err")
            pa_rows[i]["e"].config(fg="#ff6b6b"); return "break"
        var.set(fmt); l["bin"]=fmt
        try: root.clipboard_clear(); root.clipboard_append(fmt); root.update()
        except Exception: pass
        log_event("PA_LINE", po=pa["po"], sku=l["sku"], qty=l.get("qty"), bin=fmt)
        _beep("ok"); pa_save(); _row_paint(i); _pa_counts()
        pa_feedback(f"\U0001F4CB {fmt} copied \u2014 paste into BC.","done")
        if jump:
            nxt=[j for j in range(i+1,len(pa["lines"])) if not pa["lines"][j].get("bin")] \
                or [j for j in range(0,len(pa["lines"])) if not pa["lines"][j].get("bin")]
            if nxt: _focus_bin(nxt[0])
            else:
                pa_feedback("\u2713 All bins set \u2014 paste them into BC (click a field = copy), then \u2705 Confirm.","ok")
        return "break"
    def _pa_set_bin(i, binv, src):
        """The only path for setting a bin from the handheld: it records the previous value for undo."""
        if not (0 <= i < len(pa["lines"])): return False
        prev = pa["lines"][i].get("bin", "")
        pa["lines"][i]["bin"] = binv
        pa_last.update(id=pa_last["id"]+1, i=i, sku=pa["lines"][i].get("sku",""),
                       bin=binv, prev=prev, src=src)
        pa_save(); rebuild_rows()
        if binv:
            log_event("PA_LINE", po=pa.get("po",""), sku=pa["lines"][i].get("sku",""),
                      qty=pa["lines"][i].get("qty",0), bin=binv, src=src)
        return True
    def _maybe_refresh_export():
        """Refreshes the export ONLY when it is older than the threshold, instead of hitting the ERP every few minutes.
           Skrypt idzie w watku roboczym; sugestie z biezacego pliku dzialaja od razu,
           fresher data appears once the script finishes, through rebuild_rows on the queue."""
        try:
            if not cfg.get("bin_export_auto"): return
            path = (cfg.get("bin_export_path") or "").strip()
            ps1 = os.path.join(app_base_dir(), "Refresh-BinExport.ps1")
            if not (path and os.path.exists(ps1)): return
            if bc_state.get("refreshing"): return
            age_min = int(cfg.get("bin_export_max_age") or 60)
            if os.path.exists(path):
                age = (time.time() - os.path.getmtime(path)) / 60.0
                if age < age_min: return
            else:
                age = -1
            bc_state["refreshing"] = True
            logln(f"\u21BB Bin export older than {age_min} min ({age:.0f} min) \u2014 refreshing in background\u2026")
            def run():
                try:
                    import subprocess
                    r = subprocess.run(
                        ["powershell", "-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass", "-File", ps1],
                        capture_output=True, text=True, timeout=300)
                    bc_q.put(("export", "ok" if r.returncode == 0 else f"kod {r.returncode}"))
                except Exception as e:
                    bc_q.put(("export", f"error: {str(e)[:90]}"))
                finally:
                    bc_state["refreshing"] = False
            threading.Thread(target=run, daemon=True).start()
        except Exception:
            bc_state["refreshing"] = False
    def _bc_enrich(skus):
        """Worker thread: queries the ERP for the document SKUs; the result returns to the main thread on a queue."""
        try:
            n = merge_bc_into_index(skus, cfg)
            bc_q.put(("ok", n))
        except Exception as e:
            bc_q.put(("err", str(e)[:120]))
    def drain_bc():
        try:
            while True:
                kind, val = bc_q.get_nowait()
                if kind == "ok" and val:
                    rebuild_rows(); cloud_state["dirty"] = True
                    logln(f"\u2601 BC bin contents: {val} SKU enriched")
                elif kind == "export":
                    if val == "ok":
                        n_ex, info_ex = merge_export_into_index((cfg.get("bin_export_path") or "").strip())
                        rebuild_rows(); cloud_state["dirty"] = True
                        logln(f"\u21BB Bin export refreshed: {n_ex} SKU ({info_ex})")
                    else:
                        logln(f"\u26A0 Export refresh failed ({val}) \u2014 using previous file")
                elif kind == "test":
                    logln(f"\u2601 BC test: {val}")
                    messagebox.showinfo("Business Central", val)
                elif kind == "err":
                    logln(f"\u26A0 BC lookup failed ({val}) \u2014 using local history")
        except _queue.Empty:
            pass
        root.after(500, drain_bc)
    def _set_bin_from_suggestion(i, binv):
        """Clicking a suggested bin writes it into the field and commits, the same path as typing it."""
        try:
            for r in pa_rows:
                if r["idx"] == i:
                    r["var"].set(binv); _commit_bin(i, False)
                    logln(f"\U0001F4CD Suggested bin applied: {pa['lines'][i]['sku']} \u2192 {binv}")
                    log_event("LOC_SUGGEST_USED", sku=pa["lines"][i].get("sku",""), bin=binv)
                    break
        except Exception: pass
    def rebuild_rows(focus_first=False):
        for w in _in.winfo_children(): w.destroy()
        pa_rows.clear()
        hdrf=tk.Frame(_in, bg=UI["panel"]); hdrf.pack(fill="x")
        for txt,wd in (("#",4),("ITEM",26),("QTY",6),("BIN LOCATION",16),("SUGGESTED (click)",22)):
            tk.Label(hdrf, text=txt, width=wd, anchor="w", bg=UI["panel"], fg=UI["muted"],
                     font=("Cascadia Mono",9,"bold")).pack(side="left", padx=(8,2))
        for i,l in enumerate(pa["lines"]):
            rf=tk.Frame(_in, bg=UI["panel"]); rf.pack(fill="x", pady=1)
            tk.Label(rf, text=f"{i+1}", width=4, anchor="w", bg=UI["panel"], fg=UI["muted"],
                     font=("Cascadia Mono",10)).pack(side="left", padx=(8,2))
            lab=tk.Label(rf, text=l["sku"], width=26, anchor="w", bg=UI["panel"],
                         font=("Cascadia Mono",11,"bold")); lab.pack(side="left", padx=2)
            tk.Label(rf, text=str(l.get("qty") or "?"), width=6, anchor="w", bg=UI["panel"], fg=UI["muted"],
                     font=("Cascadia Mono",11)).pack(side="left", padx=2)
            var=tk.StringVar(value=l.get("bin",""))
            e=tk.Entry(rf, textvariable=var, width=16, bg=UI["panel2"], insertbackground=UI["text"],
                       relief="flat", font=("Cascadia Mono",12))
            e.pack(side="left", padx=2, ipady=2)
            e.bind("<Return>",   lambda ev, i=i: _commit_bin(i, True))
            e.bind("<Tab>",      lambda ev, i=i: _commit_bin(i, True))
            e.bind("<Down>",     lambda ev, i=i: _commit_bin(i, True))
            e.bind("<Up>",       lambda ev, i=i: (_commit_bin(i, False), _focus_bin(i-1), "break")[-1])
            e.bind("<FocusOut>", lambda ev, i=i: _commit_bin(i, False))
            e.bind("<FocusIn>",  lambda ev, i=i: _copy_if_filled(i))
            e.bind("<Control-z>",lambda ev: (undo_serial(), "break")[-1])
            sug = loc_suggest(l["sku"])
            if sug:
                sf = tk.Frame(rf, bg=UI["panel"]); sf.pack(side="left", padx=(6,2))
                for c in sug[:2]:
                    if c.get("src") == "bc":
                        q = c.get("qty") or 0
                        _tip_txt = f"{int(q) if float(q).is_integer() else q} in stock"
                    else:
                        _tip_txt = f'{c.get("n",0)}x, last {(c.get("last") or "")[:10]}'
                    txt = f'\u21b3 {c["bin"]}' + (" \u2601" if c.get("src") == "bc" else "")
                    b = tk.Label(sf, text=txt, bg=UI["panel"], fg=UI["ok"], cursor="hand2",
                                 font=("Cascadia Mono",10,"bold"))
                    b.pack(side="left", padx=3)
                    b.bind("<Button-1>", lambda ev, i=i, v=c["bin"]: (_set_bin_from_suggestion(i, v), "break")[-1])
                    tk.Label(sf, text=_tip_txt, bg=UI["panel"], fg=UI["faint"],
                             font=("Bahnschrift",7)).pack(side="left", padx=(0,4))
            else:
                tk.Label(rf, text="\u2014 new item", bg=UI["panel"], fg=UI["faint"],
                         font=("Bahnschrift",8)).pack(side="left", padx=(8,2))
            pa_rows.append({"e":e,"var":var,"idx":i,"lab":lab})
            _row_paint(i)
        _pa_counts()
        if focus_first: _focus_first_empty()
    # dokladanie pozycji spoza dokumentu
    _add = tk.Frame(t_in, bg=UI["bg"]); _add.pack(fill="x", padx=16, pady=(0,2))
    tk.Label(_add, text="\u2795 Add:", fg=UI["muted"], bg=UI["bg"], font=("Bahnschrift",9,"bold")).pack(side="left")
    _a_sku=tk.Entry(_add, width=18, bg=UI["panel2"], fg=UI["text"], insertbackground=UI["text"],
                    relief="flat", font=("Cascadia Mono",11)); _a_sku.pack(side="left", padx=4, ipady=2)
    _a_qty=tk.Entry(_add, width=6, bg=UI["panel2"], fg=UI["text"], insertbackground=UI["text"],
                    relief="flat", font=("Cascadia Mono",11)); _a_qty.pack(side="left", padx=4, ipady=2)
    def add_line(_=None):
        skv=_norm_sku(_a_sku.get())
        if not skv: pa_feedback("Add: type item SKU first.","warn"); return "break"
        try: q=int(_a_qty.get() or 0)
        except Exception: q=0
        pa["lines"].append({"sku":skv,"desc":"","qty":q,"bin":""})
        _a_sku.delete(0,"end"); _a_qty.delete(0,"end")
        pa_save(); rebuild_rows(); _focus_bin(len(pa["lines"])-1); return "break"
    tk.Button(_add, text="Add", command=add_line, bg=UI["panel2"], fg=UI["muted"], relief="flat",
              font=("Bahnschrift",9)).pack(side="left", padx=4)
    _a_qty.bind("<Return>", add_line)
    def _pa_load(hdr, rows, announce=True):
        # --- LOCATIONS: local history, then the ERP export, then optionally the API; all before render ---
        try:
            build_loc_index()
            _maybe_refresh_export()
            _ex = (cfg.get("bin_export_path") or "").strip()
            if _ex:
                n_ex, info_ex = merge_export_into_index(_ex)
                # "auto" = the path found by _load_cfg_raw itself, not typed in by hand.
                # On failure the path is logged too: "file not found" without naming WHICH
                # file once cost half an hour of guessing on a live station.
                _src = " auto" if cfg.get("_bin_export_auto") else ""
                logln(f"\U0001F4C4 Bin export{_src}: {n_ex} SKU ({info_ex})" if n_ex
                      else f"\u26A0 Bin export{_src}: {info_ex} \u2014 {_ex}")
            if cfg.get("bc_enabled"):
                threading.Thread(target=_bc_enrich,
                                 args=([_norm_sku(r.get("item","")) for r in rows],), daemon=True).start()
        except Exception as e:
            logln(f"\u26A0 Location index: {str(e)[:90]}")
        pa["doc"]=hdr; pa["active"]=None; pa["await"]=None
        _srcs=[str(r.get("source") or "").strip() for r in rows if r.get("source")]
        if _srcs and not pa["po"]:
            from collections import Counter as _C
            pa["po"]=_C(_srcs).most_common(1)[0][0].upper()
        pa["lines"]=[{"sku":_norm_sku(r.get("item","")), "desc":r.get("desc",""),
                      "qty":int(r.get("qty") or 0), "bin":""} for r in rows]
        pa_save(); pa_po_var.set(pa["po"]); rebuild_rows()
        if announce:
            pa_feedback(f"\U0001F4E6 {hdr} loaded ({len(rows)} items) \u2014 type PO, then fill BIN fields.","info")
        (pa_po_e if not pa["po"] else pa_rows[0]["e"] if pa_rows else pa_po_e).focus_set()
    def _day_path():
        _,_,ld=ensure_app_folders(); return os.path.join(ld,"inbound_day.json")
    def day_load():
        try:
            d=json.loads(Path(_day_path()).read_text(encoding="utf-8"))
            today=datetime.now().strftime("%Y-%m-%d")
            if d.get("date")!=today:
                if d.get("entries"):
                    shutil.copy2(_day_path(), _day_path().replace(".json", f"_{d.get('date','old')}.json"))
                d={"date":today,"entries":[]}
                Path(_day_path()).write_text(json.dumps(d,ensure_ascii=False),encoding="utf-8")
            return d
        except Exception:
            return {"date":datetime.now().strftime("%Y-%m-%d"),"entries":[]}
    def day_append(entries):
        d=day_load(); d["entries"].extend(entries)
        try: Path(_day_path()).write_text(json.dumps(d,ensure_ascii=False),encoding="utf-8")
        except Exception: pass
        return len(d["entries"])
    def _pa_archive_dir():
        p=cfg.get("archive_dir_pa") or os.path.join(app_base_dir(),"Archive_PutAways")
        os.makedirs(p, exist_ok=True); return p
    _confirming={"on":False}
    def confirm_putaway(auto=False):
        if _confirming["on"] or not pa["lines"]: return
        _commit_po()
        if not pa["po"]:
            pa_feedback("Set the PO number before confirming.","warn"); _beep("warn"); pa_po_e.focus_set(); return
        missing=[l for l in pa["lines"] if not l.get("bin")]
        if missing and not messagebox.askyesno("Confirm put-away",
                f"{len(missing)} line(s) without a bin. Confirm anyway?"):
            return
        _confirming["on"]=True
        try:
            po=pa["po"]
            rows=[{"item":l["sku"],"desc":l.get("desc",""),"qty":l.get("qty") or 0,
                   "bin":l.get("bin",""),"source":po,"due":""} for l in pa["lines"]]
            arch=_pa_archive_dir(); outp=os.path.join(arch, f"{po}.html")
            if os.path.exists(outp):
                try: shutil.move(outp, os.path.join(arch, f"{po}_old_{datetime.now():%H%M%S}.html"))
                except Exception as e:
                    # Nieudane odlozenie starej wersji znaczy, ze za chwile ja NADPISZEMY.
                    logln(f"⚠ Archive: nie moge odlozyc poprzedniej wersji {po}: {str(e)[:70]}")
            Path(outp).write_text(build_putaway_html(pa["doc"], po, rows, filled=True), encoding="utf-8")
            n=day_append([{"sku":l["sku"],"desc":l.get("desc",""),"qty":l.get("qty") or 0,"po":po}
                          for l in pa["lines"]])
            log_event("PA_CONFIRMED", po=po, doc=pa["doc"], lines=len(pa["lines"]),
                      missing=len(missing), file=f"{po}.html", auto=auto)
            logln(f"\u2705 Put-away confirmed: {po} ({len(pa['lines'])} lines) \u2192 {po}.html · day log {n} lines")
            pa_clear_file(archive=True)
            pa["po"]=""; pa["doc"]=""; pa["lines"]=[]; pa["active"]=None; pa["await"]=None
            pa_po_var.set(""); pa_save(); rebuild_rows()
            pa_feedback(f"\u2705 {po} archived (boxes filled) \u2014 load next put-away.","done"); _beep("done")
            try:
                pa_done.update(doc=pa.get("doc",""), po=po,
                               lines=len(rows), missing=sum(1 for r in rows if not r.get("bin")),
                               ts=datetime.now().strftime("%H:%M:%S"), file=os.path.basename(outp))
            except Exception as e:
                # pa_done zasila "LAST CONFIRMED" na skanerze - ochrone przed
                # an accidental Confirm. A silent failure leaves STALE data there.
                logln(f"⚠ LAST CONFIRMED nie zaktualizowane: {str(e)[:70]}")
            if pa_pending:
                root.after(600, lambda: load_next_pa(auto=True))   # nastepna lista wskakuje sama po potwierdzeniu
        except Exception as e:
            messagebox.showerror("Confirm failed", str(e)); _beep("err")
        finally:
            _confirming["on"]=False
    def load_last_pa():
        if not LAST_PUTAWAY["rows"]:
            pa_feedback("No put-away processed yet.","warn"); _beep("warn"); return
        if pa["lines"] and not messagebox.askyesno("Inbound",
                f"Replace current session ({len(pa['lines'])} lines) with {LAST_PUTAWAY['hdr']}?"):
            return
        pa["po"]=""; _pa_load(LAST_PUTAWAY["hdr"], LAST_PUTAWAY["rows"]); _beep("ok")
    def _pa_pending_sync():
        n = len(pa_pending)
        if n:
            btn_pa_next.config(text=f"\u23ED  Next put-away ({n})")
            if not btn_pa_next.winfo_ismapped():
                btn_pa_next.pack(side="left", padx=5, after=btn_pa_last)
        else:
            btn_pa_next.pack_forget()
    def load_next_pa(auto=False):
        """Laduje nastepny put-away z kolejki pa_pending (FIFO)."""
        if not pa_pending:
            pa_feedback("No pending put-aways.","warn"); _beep("warn"); return
        if pa["lines"] and not auto and not messagebox.askyesno("Inbound",
                f"Replace current session ({len(pa['lines'])} lines) with next pending put-away?"):
            return
        hdr, rows = pa_pending.pop(0)
        pa["po"]=""; _pa_load(hdr, rows)
        if pa_pending:
            pa_feedback(f"\U0001F4E6 {hdr} loaded ({len(rows)} items) \u2014 {len(pa_pending)} more pending (\u23ED Next).","info")
        _pa_pending_sync()
    def pa_load_pdf():
        """Manual put-away PDF import (watcher off, a file from email, or a retro entry).
           Routed through pa_q, the same path the watcher uses (FIFO, auto-load), WITHOUT printing a sheet."""
        f = filedialog.askopenfilename(title="Select put-away PDF", filetypes=[("PDF","*.pdf")])
        if not f: return
        try:
            dtype, hdr, source, rows, warns = parse_document(f)
        except Exception as e:
            messagebox.showerror("PDF error", str(e)); _beep("err"); return
        if dtype != "putaway" or not rows:
            pa_feedback(f"\u2715 {os.path.basename(f)}: not a put-away PDF.", "warn"); _beep("warn"); return
        learn_skus(rows)
        LAST_PUTAWAY["hdr"] = hdr; LAST_PUTAWAY["rows"] = list(rows)
        for w in warns[:3]: logln(f"\u26a0 {os.path.basename(f)}: {w}")
        pa_q.put((hdr, list(rows)))
        pa_feedback(f"\U0001F4C2 {hdr or os.path.basename(f)} parsed ({len(rows)} lines) \u2014 queued.", "info")
    def export_inbound():
        d=day_load()
        if not d["entries"]:
            pa_feedback("Day log is empty \u2014 confirm at least one put-away first.","warn"); _beep("warn"); return
        if pa["lines"]:
            pa_feedback("\u26A0 Current put-away NOT confirmed \u2014 it is not in this export.","warn")
        _,_,ld=ensure_app_folders()
        exd=os.path.join(ld,"Inbound_Exports"); os.makedirs(exd, exist_ok=True)
        fn=f"Inbound_{datetime.now():%Y-%m-%d_%H%M%S}.html"
        out=os.path.join(exd,fn)
        try:
            export_inbound_day_html(d["entries"], out)
            tsv="Item\tDescription\tQty\tPO\n"+"\n".join(
                f'{e["sku"]}\t{e.get("desc","")}\t{e.get("qty") or ""}\t{e.get("po","")}' for e in d["entries"])
            try: root.clipboard_clear(); root.clipboard_append(tsv); root.update()
            except Exception: pass
            log_event("PA_EXPORTED", n=len(d["entries"]), pos=len({e.get("po") for e in d["entries"]}), file=fn)
            logln(f"\U0001F4E4 Day export: {fn} ({len(d['entries'])} lines)")
            webbrowser.open(Path(out).as_uri())
            pa_feedback("\U0001F4E4 Day file opened \u2014 \U0001F4CB Copy table \u2192 Teams.","done"); _beep("line")
            if messagebox.askyesno("Day log", "Reset the day log now? (Export file keeps everything.)"):
                try:
                    shutil.copy2(_day_path(), _day_path().replace(".json", f"_exported_{datetime.now():%H%M%S}.json"))
                    Path(_day_path()).write_text(json.dumps({"date":d["date"],"entries":[]},ensure_ascii=False),encoding="utf-8")
                except Exception: pass
        except Exception as e:
            messagebox.showerror("Export failed", str(e)); _beep("err")
    def clear_inbound():
        if pa["lines"] and not messagebox.askyesno("Clear session",
                f"Clear inbound session ({len(pa['lines'])} lines)?\nArchived JSON copy stays in Logi."):
            return
        log_event("PA_SESSION_CLOSED", po=pa["po"], n=len(pa["lines"]))
        pa_clear_file(archive=True)
        pa["po"]=""; pa["doc"]=""; pa["lines"]=[]; pa["active"]=None; pa["await"]=None
        pa_po_var.set(""); pa_save(); rebuild_rows()
        _pa_pending_sync()
        if pa_pending:
            pa_feedback(f"\U0001F9F9 Session cleared \u2014 {len(pa_pending)} put-away(s) pending (\u23ED Next).","info")
        else:
            pa_feedback("\U0001F9F9 Session cleared \u2014 process a put-away or add lines.","info")
    _pb = tk.Frame(t_in, bg=UI["bg"]); _pb.pack(pady=(2,10))
    tk.Button(_pb, text="\U0001F4C2  Load PDF", command=pa_load_pdf, bg=UI["panel2"], fg=UI["muted"],
              relief="flat", font=("Bahnschrift",9)).pack(side="left", padx=5)
    btn_pa_last = tk.Button(_pb, text="\U0001F4E5  Load last put-away", command=load_last_pa, bg=UI["panel2"], fg=UI["muted"],
              relief="flat", font=("Bahnschrift",9))
    btn_pa_last.pack(side="left", padx=5)
    btn_pa_next = tk.Button(_pb, text="\u23ED  Next put-away", command=load_next_pa, bg="#B45309", fg="white",
              relief="flat", font=("Bahnschrift",9,"bold"))   # pakowany dynamicznie gdy kolejka niepusta
    tk.Button(_pb, text="\u2705  Confirm put-away", command=lambda: confirm_putaway(False), bg="#1E7D46", fg="white",
              relief="flat", font=("Bahnschrift",9,"bold")).pack(side="left", padx=5)
    tk.Button(_pb, text="\U0001F4E4  Export day (Teams)", command=export_inbound, bg="#0E7490", fg="white",
              relief="flat", font=("Bahnschrift",9,"bold")).pack(side="left", padx=5)
    tk.Button(_pb, text="\U0001F9F9  Clear session", command=clear_inbound, bg=UI["panel2"], fg=UI["muted"],
              relief="flat", font=("Bahnschrift",9)).pack(side="left", padx=5)
    def _on_tab_changed(_=None):
        pa["on"] = (nb.select() == str(t_in))
        if pa["on"]:
            pa_po_var.set(pa["po"]); rebuild_rows(); _focus_first_empty()
    nb.bind("<<NotebookTabChanged>>", _on_tab_changed, add="+")
    def drain_pa():
        try:
            while True:
                hdr, rows = pa_q.get_nowait()
                pa_pending.append((hdr, rows))          # v1.0 FIX: NIC nie przepada - wszystko do kolejki FIFO
                try:
                    if nb.select() != str(t_in): nb.alert("Put-away", True)
                    web_alerts["inbound"] = web_alerts.get("inbound", 0) + 1
                except Exception: pass
                if not pa["lines"]:
                    load_next_pa(auto=True)             # pusta sesja -> pierwszy z kolejki wskakuje od razu
                    nb.select(t_in)
                else:
                    _pa_pending_sync()
                    pa_feedback(f"\U0001F195 {hdr} queued \u2014 {len(pa_pending)} pending. Finish current, then \u23ED Next.","warn")
                    logln(f"\U0001F195 Put-away queued: {hdr} ({len(rows)} lines) \u2014 {len(pa_pending)} pending")
                _beep("line")
        except _queue.Empty:
            pass
        root.after(400, drain_pa)
    def toggle_session():
        if session["on"]:
            inc = [it for it in session["items"] if len(it["serials"]) < it["qty"]*it.get("pack",1)]
            if inc and not messagebox.askyesno("End session",
                    f"{len(inc)} item(s) incomplete — end anyway?\n(Serials already copied per completed item.)"):
                scan_entry.focus_set(); return
            tot = sum(len(it["serials"]) for it in session["items"])
            log_event("SESSION_CLOSED", items=len(session["items"]), serials=tot)
            logln(f"⚡ Session closed: {len(session['items'])} item(s), {tot} serial(s)")
            sess_clear_file(archive=True)
            session["on"]=False; session["items"]=[]; session["active"]=None
            session["seen"]=set(); session["await"]=None; session["suggest"]=0
            btn_session.config(text="⚡  Serial session", bg=UI["panel2"], fg=UI["muted"])
            pack_fr.pack_forget()
            feedback("Serial session ended.", "info"); refresh_station(); scan_entry.focus_set(); return
        if station["cur"]:
            feedback("Finish or remove the current pick first.", "warn"); _beep("warn"); return
        session["on"]=True; session["await"]="sku"
        log_event("SESSION_OPEN")
        btn_session.config(text="✕  End session", bg="#B02E0C", fg="white")
        pack_fr.pack(fill="x", padx=16, pady=(0,2)); _highlight_pack(1)
        feedback("⚡ SERIAL SESSION — scan an item SKU to begin.", "info")
        refresh_station(); scan_entry.focus_set()

    def handle_session(code):
        aw = session["await"]
        if aw == "qty":
            if not code and session["suggest"]:
                q = session["suggest"]                       # Enter = akceptacja podpowiedzi profilu
            else:
                q = parse_scanned_qty(code)
                if q is None and code.isdigit(): 
                    try: q = int(code)
                    except Exception: q = None
            if not q or q <= 0:
                feedback(f"'{code}': enter/scan a quantity (e.g. 30 or Q30).", "warn"); _beep("warn"); return
            it = session["items"][session["active"]]
            it["qty"] = q; learn_bulk(it["sku"], q)
            session["await"]="serials"; session["suggest"]=0
            feedback(f"{it['sku']} ×{q} — scan serial(s); set PACK below for dual/quad/6.", "ok"); _beep("ok")
            sess_save(); refresh_station(); return
        if aw == "serials":
            it = session["items"][session["active"]]
            tgt = it["qty"] * it.get("pack", 1)
            if len(it["serials"]) < tgt:
                okv, why = valid_serial(code, [it["sku"]])
                if not okv:
                    if looks_like_sku(code) or classify_scan(code, [it["sku"]]) == "sku":
                        feedback(f"{it['sku']}: {len(it['serials'])}/{tgt} — finish serials before next SKU.", "warn"); _beep("warn"); return
                    feedback(f"REJECTED: {code} — {why}", "err"); _beep("err")
                    log_event("SERIAL_REJECTED", code=code, reason=why, pick="SESSION"); return
                if code in session["seen"]:
                    feedback(f"DUPLICATE serial {code}.", "err"); _beep("err")
                    log_event("DUP_SERIAL", code=code, pick="SESSION"); return
                it["serials"].append(code); session["seen"].add(code); _beep("ok")
                feedback(f"✓ {code}   ({len(it['serials'])}/{tgt})", "ok")
                if len(it["serials"]) >= tgt:
                    content = bc_serial_export(it["serials"], it.get("pack", 1))
                    try: root.clipboard_clear(); root.clipboard_append(content); root.update()
                    except Exception: pass
                    _beep("line")
                    log_event("SESSION_ITEM", sku=it["sku"], qty=it["qty"])
                    write_log_file(f"    SESSION {it['sku']} x{it['qty']} -> clipboard")
                    pk = it.get("pack", 1); rows = (len(it["serials"]) + pk - 1)//pk
                    pinfo = "" if pk <= 1 else f" ({pk}x → {rows} box row(s))"
                    feedback(f"📋 {it['sku']} — {len(it['serials'])} serial(s) copied{pinfo}. Paste to BC, then scan NEXT SKU.", "done")
                    session["await"]="sku"
                sess_save(); refresh_station(); return
            session["await"]="sku"   # komplet - przelot do nowego SKU
        # aw == "sku": start nowego itemu
        sku = _norm_sku(code)
        if not sku or parse_scanned_qty(code) is not None:
            feedback("Scan an item SKU first.", "warn"); _beep("warn"); return
        _okv,_ = valid_serial(code, [])
        if _okv and not looks_like_sku(code):
            feedback(f"{code}: looks like a SERIAL — scan an item SKU to start the next item.", "warn"); _beep("warn"); return
        session["items"].append({"sku": sku, "qty": 0, "pack": 1, "serials": []})
        session["active"] = len(session["items"]) - 1
        sug, q = bulk_suggest(sku)
        session["suggest"] = q if sug else 0
        session["await"] = "qty"
        if sug:
            feedback(f"{sku}: usually bulk ×{q} — press ENTER to accept, or scan/type qty.", "info")
        else:
            feedback(f"{sku}: scan/type quantity (e.g. 30).", "info")
        _beep("ok"); sess_save(); refresh_station()

    def refresh_session():
        st_title.config(text="⚡ SERIAL SESSION")
        hints = {"sku": "scan item SKU", "qty": "scan/type quantity (ENTER = accept suggestion)",
                 "serials": "scan serials"}
        st_sub.config(text=f"Free serial capture · {hints.get(session['await'],'')} · Ctrl+Z undo · ✕ ends")
        done = sum(min(len(it["serials"]), it["qty"]*it.get("pack",1)) for it in session["items"])
        need = sum(it["qty"]*it.get("pack",1) for it in session["items"]) or 1
        prog.config(value=done/need*100); st_prog.config(text=f"{done} / {need}")
        st_queue.config(text="")
        itree.delete(*itree.get_children())
        for idx, it in enumerate(session["items"]):
            pk = it.get("pack", 1); tgt = it["qty"] * pk
            full = tgt and len(it["serials"]) >= tgt
            tag = "done" if full else ("part" if it["serials"] else "todo")
            mark = "✓" if full else ("▶" if idx == session["active"] else "")
            pid = f"s{idx}"
            itree.insert("", "end", iid=pid, open=(idx == session["active"]),
                         values=(mark, it["sku"], (f'{it["qty"]} box · {pk}x' if pk>1 else "SESSION"),
                                 f'{len(it["serials"])}/{tgt or "?"}'), tags=(tag,))
            for j, s in enumerate(it["serials"]):
                lbl = f"#{j+1}" if pk <= 1 else f"box {j//pk + 1}"
                itree.insert(pid, "end", iid=f"{pid}_{j}", values=("", s, "", lbl), tags=("todo",))
        _sync_pack_combo()

    def on_scan(_=None):
        code = scan_var.get().strip()
        was_scanned = _qty_was_scanned()          # snapshot predkosci PRZED resetem
        scan_keys["times"] = []                    # reset na kolejny input
        scan_var.set(""); scan_entry.focus_set()
        if session["on"]:
            handle_session(code); return
        if not code: return
        if serial_mode["on"]: handle_serial(code); return     # tryb zbierania seriali
        if not station["cur"]: feedback("No active pick to scan against.", "warn"); _beep("err"); return
        if bulk["on"]: handle_bulk(code, was_scanned); return
        cur = station["cur"]
        # --- pick ALREADY verified (suspended serials view): a scan here is NOT a picker error ---
        if all(i["scanned"] >= i["need"] for i in cur["items"]):
            if serial_mode["radios"]:
                okv, _w = valid_serial(code, [i["sku"] for i in cur["items"]])
                if okv:                                   # serial? auto-wznow zbieranie
                    enter_serial_mode()
                    if len(serial_mode["radios"]) == 1:   # 1 model = jednoznaczne przypisanie
                        handle_serial(code)
                    else:                                 # kilka modeli = najpierw wybierz radio
                        feedback("Serial mode resumed — click the radio this serial belongs to, then re-scan.", "warn"); _beep("warn")
                    return
                feedback("Pick verified — press 'Serials →' to resume serial scanning.", "warn"); _beep("warn")
            else:
                feedback("Pick verified.", "info")
            return
        idx = _find_line(code)
        if idx is None:
            # ===== TRIAGE of a missed scan: real error versus noise =====
            if _sku_exists(code):
                # an item on the pick with its line already full is an overpick attempt (deduplicated per pick per code)
                feedback(f"{code}: LINE ALREADY COMPLETE", "warn"); _beep("err")
                key = _norm_sku(code)
                if key not in flagged["over"]:
                    flagged["over"].add(key)
                    log_event("OVERPICK", code=code, pick=cur["title"], doc=cur["hdr"])
                return
            kind = classify_scan(code, [i["sku"] for i in cur["items"]])
            key = code.strip().upper()
            if kind == "sku":
                # a SKU from the catalogue registry that is not on this pick is a REAL wrong item, counted as a loss
                feedback(f"{code}: NOT ON THIS PICK", "err"); _beep("err")
                if key not in flagged["wrong"]:
                    flagged["wrong"].add(key)
                    log_event("WRONG_ITEM", code=code, pick=cur["title"], doc=cur["hdr"])
            else:
                # noise: qty outside bulk, an EAN, a serial scanned too early, or an unknown code - NOT counted as a loss
                msg = {"qty":     "Q-code scanned — quantity codes only work in BULK mode.",
                       "numeric": f"{code}: numeric code not matched — not on this pick / not in catalog yet.",
                       "serial":  f"{code}: looks like a SERIAL — serials are scanned after verification.",
                       "unknown": f"{code}: not recognized — check the barcode."}[kind]
                feedback(msg, "warn"); _beep("warn")
                if key not in flagged["noise"]:
                    flagged["noise"].add(key)
                    log_event("NOISE_SCAN", code=code, reason=kind, pick=cur["title"], doc=cur["hdr"])
            return
        line = station["cur"]["items"][idx]; line["scanned"] += 1
        _all = all(i["scanned"] >= i["need"] for i in cur["items"])
        if _all:      pass                                   # caly pick gotowy -> fanfare w check_complete (bez dublu)
        elif line["scanned"] == line["need"]: _beep("line")  # ta linia wlasnie skompletowana (10/10)
        else:         _beep("ok")                             # zwykly poprawny skan
        feedback(f"{line['sku']}   +1   ({line['scanned']}/{line['need']})", "ok")
        _now = time.time()
        if pick_t["t0"] is None: pick_t["t0"] = _now
        pick_t["scans"].append((round(_now - pick_t["t0"], 1), line["sku"], line.get("bin",""), 1))
        try: itree.see(str(idx)); itree.selection_set(str(idx))
        except Exception: pass
        refresh_station(); check_complete()

    def manual_sku():
        from tkinter import simpledialog
        code = simpledialog.askstring("Manual SKU",
            "Type the SKU (counts as 1 unit — quantity can never be typed):", parent=root)
        if code: scan_var.set(code.strip()); on_scan()

    def toggle_serial():
        """Manually flags or unflags an item as serialised (combined with auto-detection).
           When enabled it offers to REMEMBER the SKU for future picks."""
        if serial_mode["on"]: return
        cur = station["cur"]
        if not cur: feedback("No active pick.", "warn"); return
        sel = itree.selection()
        if not sel: feedback("Select a line, then press 🔖 Serial to mark it.", "warn"); return
        try: idx = int(sel[0])
        except (ValueError, IndexError): return
        line = cur["items"][idx]
        line["serial"] = not line.get("serial", False)
        key = _norm_sku(line["sku"])
        if line["serial"]:
            feedback(f"{line['sku']}: serial tracking ON 🔖", "info")
            if messagebox.askyesno("Remember this item?",
                f"Remember {line['sku']} as a serial item for ALL future picks?\n\n"
                "Yes = it will auto-flag every time it appears."):
                s = load_serial_skus(); s.add(key); save_serial_skus(s)
                logln(f"🔖 Remembered {line['sku']} as serial item")
        else:
            feedback(f"{line['sku']}: serial tracking OFF", "info")
            remembered = load_serial_skus()
            if key in remembered and messagebox.askyesno("Forget this item?",
                f"{line['sku']} is remembered as a serial item.\n\nForget it (stop auto-flagging)?"):
                remembered.discard(key); save_serial_skus(remembered)
                logln(f"   Forgot {line['sku']} as serial item")
        refresh_station()
        try: itree.selection_set(str(idx))
        except Exception: pass

    def emergency_override(_=None):
        if serial_mode["on"]: return              # override dziala tylko w walidacji
        cur = station["cur"]
        if not cur: return
        sel = itree.selection()
        if not sel: feedback("Select the line to override first.", "warn"); return
        try: idx = int(sel[0])
        except (ValueError, IndexError): return
        line = cur["items"][idx]
        if not messagebox.askyesno("Emergency override",
            f"Force-complete this line?\n\n{line['sku']}   ({line['scanned']}/{line['need']})\n\n"
            "Use ONLY if the system is wrong. This action is logged."): return
        line["scanned"] = line["need"]
        write_log_file(f"    OVERRIDE: {line['sku']} forced complete on {cur['title']} (doc {cur['hdr']})")
        log_event("OVERRIDE", sku=line["sku"], pick=cur["title"], doc=cur["hdr"])
        logln(f"⚠ OVERRIDE: {line['sku']} forced complete on {cur['title']}")
        feedback(f"Override logged: {line['sku']} marked complete.", "warn")
        refresh_station(); check_complete()

    def check_complete():
        cur = station["cur"]
        if not cur: return
        if all(i["scanned"] >= i["need"] for i in cur["items"]):
            units = sum(i["need"] for i in cur["items"])
            write_log_file(f"    VERIFIED {cur['title']} (doc {cur['hdr']})")
            log_event("PICK_VERIFIED", pick=cur["title"], doc=cur["hdr"], units=units, lines=len(cur["items"]))
            # --- SLOTTING TELEMETRY (once per pick, no operator identity) ---
            if pick_t["scans"] and not pick_t.get("sent"):
                pick_t["sent"] = True
                printed = cur.get("printed_ts")
                cycle = round(pick_t["t0"] - printed, 1) if (printed and pick_t["t0"]) else None
                log_event("PICK_TELEMETRY", pick=cur["title"], doc=cur["hdr"],
                          n_lines=len(cur["items"]), units=units,
                          cycle_s=cycle,                                   # wydruk -> pierwszy skan = zbieranie
                          validation_s=round(time.time() - pick_t["t0"], 1),
                          queued=bool(cur.get("queued")),
                          lines=[{"t": t, "sku": s, "bin": b, "qty": q} for t, s, b, q in pick_t["scans"]])
            logln(f"✓ VERIFIED {cur['title']} (doc {cur['hdr']})")
            models = {_norm_sku(i["sku"]) for i in cur["items"] if i.get("serial")}
            if models:
                _beep("line")     # weryfikacja OK, ale pick jeszcze trwa (seriale) - krotki wznoszacy
                feedback(f"✓ PICK VERIFIED — now collect serials for {len(models)} item type(s).", "done")
                root.after(900, enter_serial_mode)
            else:
                _beep("done")     # pick faktycznie ZAKONCZONY - fanfara
                feedback(f"✓ PICK VERIFIED — {cur['title']}   (no serial items — closing)", "done")
                station["cur"] = None; root.after(1600, load_next)

    # ====================== TRYB ZBIERANIA SERIALI ======================
    def _show_validation_buttons():
        btn_back.pack_forget(); btn_copy.pack_forget(); btn_undo.pack_forget()
        btn_finish.pack_forget(); pack_fr.pack_forget()
        btn_bulk.pack(side="left", padx=3, pady=6); btn_manual.pack(side="left", padx=3, pady=6)
        btn_serial.pack(side="left", padx=(3,10), pady=6)

    def enter_serial_mode():
        cur = station["cur"]
        if not cur: return
        if not serial_mode["radios"]:                 # buduj TYLKO przy pierwszym wejsciu
            merged = {}
            for i in cur["items"]:
                if not i.get("serial"): continue
                k = _norm_sku(i["sku"])
                if k in merged: merged[k]["qty"] += i["need"]
                else: merged[k] = {"sku": i["sku"], "qty": i["need"], "serials": [], "pack": 1}
            serial_mode["radios"] = list(merged.values())
            serial_mode["active"] = 0 if serial_mode["radios"] else None
            serial_mode["seen"] = set(); serial_mode["history"] = []
        if not serial_mode["radios"]:
            feedback("No serial items in this pick.", "warn"); return
        serial_mode["on"] = True
        bulk_reset()
        btn_bulk.pack_forget(); btn_manual.pack_forget(); btn_serial.pack_forget()
        if btn_to_serials.winfo_ismapped(): btn_to_serials.pack_forget()
        btn_back.pack(side="left", padx=3, pady=6); btn_copy.pack(side="left", padx=3, pady=6)
        btn_undo.pack(side="left", padx=3, pady=6); btn_finish.pack(side="left", padx=3, pady=6)
        pack_fr.pack(fill="x", padx=16, pady=(0,2)); _sync_pack_combo()
        refresh_station()
        if serial_mode["active"] is not None:
            r0 = serial_mode["radios"][serial_mode["active"]]
            feedback(f"SERIAL COLLECTION — scan serials for {r0['sku']} ({len(r0['serials'])}/{r0['qty']}). Click a radio to switch.", "info")
        scan_entry.focus_set()

    def suspend_serial_mode():
        """Returns to the pick view WITHOUT losing captured serials; resume through 'Serials →'."""
        serial_mode["on"] = False
        _show_validation_buttons()
        feedback("Back to pick view. Press 'Serials →' to resume serial scanning.", "info")
        refresh_station(); scan_entry.focus_set()

    def exit_serial_mode():
        """Full exit plus a wipe of serial data, on closing a pick or starting a new one."""
        serial_mode["on"]=False; serial_mode["radios"]=[]; serial_mode["active"]=None
        serial_mode["seen"]=set(); serial_mode["history"]=[]
        _show_validation_buttons()

    def _sync_pack_combo():
        """Podswietl przycisk paka aktywnego radia."""
        if session["on"]:
            if session["active"] is not None:
                _highlight_pack(session["items"][session["active"]].get("pack", 1))
            return
        if serial_mode["active"] is None: return
        pk = serial_mode["radios"][serial_mode["active"]].get("pack", 1)
        _highlight_pack(pk)

    def copy_radio(idx, auto=False):
        r = serial_mode["radios"][idx]
        if not r["serials"]:
            feedback(f"No serials scanned for {r['sku']} yet.", "warn"); return
        pk = r.get("pack", 1)
        content = bc_serial_export(r["serials"], pk)
        try:
            root.clipboard_clear(); root.clipboard_append(content); root.update()
        except Exception as e:
            feedback(f"Clipboard error: {e}", "err"); _beep("err"); return
        _beep("line" if auto else "ok")
        how = "auto-copied" if auto else "copied"
        n = len(r["serials"]); rows = (n + pk - 1)//pk
        pinfo = "" if pk <= 1 else f" ({pk}x pack → {rows} box row(s), {pk}/row)"
        feedback(f"📋 {r['sku']} — {n} serial(s) {how}{pinfo}. Paste into BC (Ctrl+V).", "done")
        write_log_file(f"    SERIALS {r['sku']} x{len(r['serials'])} pack={pk} -> clipboard (pick {station['cur']['title']})")
        logln(f"📋 {r['sku']}: {len(r['serials'])} serials (pack {pk}x) → clipboard")

    def copy_active_radio():
        if serial_mode["active"] is None: feedback("Select a radio first.", "warn"); return
        copy_radio(serial_mode["active"])

    def handle_serial(code):
        radios = serial_mode["radios"]
        if serial_mode["active"] is None or not radios:
            feedback("No radio selected — click a radio row first.", "warn"); _beep("err"); return
        parts = [c.strip() for c in re.split(r"[\n\r,;|]", code) if c.strip()]  # wzorzec z BC Scanner
        if not parts: return
        pick_skus = [i["sku"] for i in station["cur"]["items"]]
        r = radios[serial_mode["active"]]
        for serial in parts:
            okv, why = valid_serial(serial, pick_skus)          # walidacja formatu seriala
            if not okv:
                feedback(f"REJECTED '{serial}': {why}", "err"); _beep("err")
                log_event("SERIAL_REJECTED", code=serial, reason=why, pick=station["cur"]["title"]); continue
            if serial in serial_mode["seen"]:
                feedback(f"DUPLICATE serial: {serial} (already scanned in this pick)", "err"); _beep("err")
                log_event("DUP_SERIAL", serial=serial, pick=station["cur"]["title"]); continue
            target = r["qty"] * r.get("pack",1)     # pudelka x pack = ile seriali razem
            if len(r["serials"]) >= target:
                feedback(f"{r['sku']} already has all {target} serials. Select another radio.", "warn"); _beep("err"); continue
            r["serials"].append(serial); serial_mode["seen"].add(serial)
            serial_mode["history"].append((serial_mode["active"], serial))   # do cofania
            _beep("ok")
            feedback(f"✓ {serial}   ({len(r['serials'])}/{target})   {r['sku']}", "ok")
        refresh_station()
        if len(r["serials"]) >= r["qty"] * r.get("pack",1):   # radio kompletne -> auto-kopiuj do schowka
            copy_radio(serial_mode["active"], auto=True)

    def undo_serial(_=None):
        if pa["on"]:
            for idx in range(len(pa["lines"])-1,-1,-1):
                l=pa["lines"][idx]
                if l.get("bin"):
                    l["bin"]=""
                    pa_feedback(f"\u21B6 Bin cleared for {l['sku']} \u2014 enter again.","warn"); _beep("ok")
                    pa_save(); rebuild_rows(); _focus_bin(idx); return
            pa_feedback("Nothing to undo.","warn"); return
        if session["on"]:
            for it in reversed(session["items"]):
                if it["serials"]:
                    s = it["serials"].pop(); session["seen"].discard(s)
                    session["await"] = "serials"; session["active"] = session["items"].index(it)
                    _beep("ok"); feedback(f"↶ Undone: {s} removed from {it['sku']} ({len(it['serials'])}/{it['qty']})  ⚠ re-copy before pasting!", "warn")
                    sess_save(); refresh_station(); scan_entry.focus_set(); return
            feedback("Nothing to undo.", "warn"); return

        """Undoes the MOST RECENT serial scan, from any unit."""
        if not serial_mode["on"]: return
        hist = serial_mode["history"]
        if not hist: feedback("Nothing to undo.", "warn"); return
        idx, serial = hist.pop()
        r = serial_mode["radios"][idx]
        tgt = r["qty"] * r.get("pack", 1)
        was_complete = len(r["serials"]) >= tgt          # schowek mogl dostac auto-copy - bedzie nieaktualny
        if serial in r["serials"]: r["serials"].remove(serial)
        serial_mode["seen"].discard(serial)
        serial_mode["active"] = idx              # wroc na radio, z ktorego cofamy
        _sync_pack_combo(); _beep("ok")
        stale = "  ⚠ CLIPBOARD OUTDATED — press Copy again before pasting!" if was_complete else ""
        feedback(f"↶ Undone: {serial} removed from {r['sku']}  ({len(r['serials'])}/{tgt}){stale}", "warn")
        refresh_station(); scan_entry.focus_set()

    def on_tree_select(_=None):
        if not serial_mode["on"]:
            scan_entry.focus_set()               # klik w liste NIE moze ukrasc fokusu skanera (zgubiony skan!)
            return
        sel = itree.selection()
        if not sel: return
        m = re.match(r"r(\d+)", sel[0])
        if not m: return
        idx = int(m.group(1))
        if idx != serial_mode["active"]:
            serial_mode["active"] = idx
            r = serial_mode["radios"][idx]
            _sync_pack_combo()
            tgt = r["qty"] * r.get("pack", 1)
            feedback(f"Selected {r['sku']} — scan its serials ({len(r['serials'])}/{tgt}).", "info")
            refresh_station()
        scan_entry.focus_set()

    def finish_pick():
        cur = station["cur"]
        if not cur: return
        radios = serial_mode["radios"]
        incomplete = [r for r in radios if len(r["serials"]) < r["qty"]*r.get("pack",1)]
        if incomplete:
            names = "\n".join(f"  · {r['sku']}  ({len(r['serials'])}/{r['qty']*r.get('pack',1)})" for r in incomplete)
            if not messagebox.askyesno("Finish pick?",
                f"Some radios are missing serials:\n\n{names}\n\nFinish anyway?"): return
        write_log_file(f"    PICK CLOSED {cur['title']} (doc {cur['hdr']}) — " +
                       ", ".join(f"{r['sku']}x{len(r['serials'])}" for r in radios))
        log_event("PICK_CLOSED", pick=cur["title"], doc=cur["hdr"],
                  serials=sum(len(r["serials"]) for r in radios), radios=len(radios))
        logln(f"✓ PICK CLOSED {cur['title']}")
        exit_serial_mode(); station["cur"] = None
        _beep("done")     # PICK ZAMKNIETY (po serialach) - fanfara nalezy sie tutaj
        feedback("Pick closed. Ready for next.", "info"); refresh_station()
        root.after(300, load_next)

    scan_entry.bind("<Return>", on_scan)
    btn_bulk.config(command=start_bulk); btn_manual.config(command=manual_sku); btn_serial.config(command=toggle_serial)
    btn_copy.config(command=copy_active_radio); btn_finish.config(command=finish_pick); btn_undo.config(command=undo_serial)
    btn_back.config(command=suspend_serial_mode); btn_to_serials.config(command=enter_serial_mode)
    itree.bind("<<TreeviewSelect>>", on_tree_select)
    root.bind("<Control-Shift-O>", emergency_override); root.bind("<Control-Shift-o>", emergency_override)
    root.bind("<Control-z>", undo_serial); root.bind("<Control-Z>", undo_serial)

    def manual_load():
        f=filedialog.askopenfilename(title="Select pick / put-away PDF", filetypes=[("Pick documents","*.pdf *.xlsx"),("PDF","*.pdf"),("BC Lines (Excel)","*.xlsx")])
        if f: threading.Thread(target=lambda: process_pdf(
            f, cfg, customers, log_q.put, on_pick=lambda pd: new_picks.put(pd), on_putaway=lambda h,r: pa_q.put((h,r))), daemon=True).start()

    # ---------- TAB: LOGS ----------
    tlog = tk.Frame(nb, bg="#1a1714"); nb.add(tlog, text="Logs", group="SYSTEM", fkey="F8", icon="🧾")
    tk.Label(tlog, text="Reader log — picks & put-aways processed by the converter", fg="#f2ece1",
             bg="#1a1714", font=("Bahnschrift",11,"bold")).pack(pady=(12,2))
    wstat = tk.Label(tlog, text="Watcher: OFF", fg="#a99a84", bg="#1a1714", font=("Cascadia Mono",10)); wstat.pack(pady=(0,6))
    log = tk.Text(tlog, height=18, bg="#2a241e", fg="#f2ece1", relief="flat", font=("Cascadia Mono",9), wrap="word")
    log.pack(fill="both", expand=True, padx=16, pady=10); log.insert("end","Ready.\n"); log.config(state="disabled")
    def logln(s):
        log.config(state="normal"); log.insert("end", s+"\n"); log.see("end"); log.config(state="disabled")
        try: status_lbl.config(text=s.replace("\n"," ")[:160])
        except Exception: pass
        try:
            live_ring.append((datetime.now().strftime("%H:%M:%S"), s.replace("\n"," ")[:120]))
            del live_ring[:-12]; cloud_state["dirty"] = True
        except Exception: pass

    def drain_io_fails():
        """Wyciaga do Logs awarie I/O zebrane przez note_io_fail().
           write_log_file i log_event nie moga wolac logln (byloby zapetlenie),
           so they accumulate at module level and reach the operator from there."""
        try:
            while IO_FAILS:
                where, err = IO_FAILS.pop(0)
                logln(f"⚠ I/O {where}: {err}")
        except Exception:
            pass
        root.after(5000, drain_io_fails)
    root.after(5000, drain_io_fails)

    # ---------- TAB: IMPACT DASHBOARD ----------
    tstat = tk.Frame(nb, bg="#1a1714"); nb.add(tstat, text="Impact", group="INSIGHTS", fkey="F6", icon="📊")
    tk.Label(tstat, text="PickCore — Impact Dashboard", fg="#2fb5a8", bg="#1a1714",
             font=("Bahnschrift",16,"bold")).pack(pady=(12,1))
    tk.Label(tstat, text="What the system caught and prevented — proof of process control",
             fg="#a99a84", bg="#1a1714", font=("Bahnschrift",9)).pack()

    cards_fr = tk.Frame(tstat, bg="#1a1714"); cards_fr.pack(fill="x", padx=16, pady=12)
    card_vals = {}
    def _make_card(parent, key, label, color):
        c = tk.Frame(parent, bg="#2a241e", highlightbackground="#45392c", highlightthickness=1)
        c.pack(side="left", expand=True, fill="both", padx=5)
        v = tk.Label(c, text="0", fg=color, bg="#2a241e", font=("Bahnschrift",26,"bold")); v.pack(pady=(12,0))
        tk.Label(c, text=label, fg="#a99a84", bg="#2a241e", font=("Bahnschrift",9), wraplength=130).pack(pady=(0,12))
        card_vals[key] = v
    _make_card(cards_fr, "verified", "Picks verified", "#46d17f")
    _make_card(cards_fr, "errors", "Errors caught", "#ff5a5f")
    _make_card(cards_fr, "loss", "Loss prevented (est.)", "#f5d24a")
    _make_card(cards_fr, "overrides", "Overrides (audit)", "#cf7fbf")
    tk.Label(tstat, text="INBOUND", fg="#7a6c58", bg="#1a1714",
             font=("Bahnschrift",8,"bold"), anchor="w").pack(fill="x", padx=21, pady=(2,0))
    in_cards_fr = tk.Frame(tstat, bg="#1a1714"); in_cards_fr.pack(fill="x", padx=16, pady=(2,12))
    _make_card(in_cards_fr, "pa_conf",    "Put-aways confirmed", "#46d17f")
    _make_card(in_cards_fr, "pa_lines",   "Lines binned", "#2fb5a8")
    _make_card(in_cards_fr, "pa_units",   "Units put away", "#f5d24a")
    _make_card(in_cards_fr, "pa_exports", "Day exports to office", "#cf7fbf")

    tk.Label(tstat, text="Errors caught — last 14 days", fg="#f2ece1", bg="#1a1714",
             font=("Bahnschrift",10,"bold")).pack(anchor="w", padx=20, pady=(4,0))
    chart = tk.Canvas(tstat, height=150, bg="#2a241e", highlightthickness=0)
    chart.pack(fill="x", padx=16, pady=(4,8))

    brk = tk.Label(tstat, text="", fg="#cfc3b0", bg="#1a1714", font=("Cascadia Mono",10), justify="left", anchor="w")
    brk.pack(fill="x", padx=20, pady=(0,4))

    ctl = tk.Frame(tstat, bg="#1a1714"); ctl.pack(fill="x", padx=16, pady=(2,10))
    tk.Label(ctl, text="€ per prevented error:", fg="#a99a84", bg="#1a1714", font=("Bahnschrift",9)).pack(side="left")
    euro_var = tk.StringVar(value=str(cfg.get("euro_per_error",150)))
    euro_entry = tk.Entry(ctl, textvariable=euro_var, width=7, bg="#0e1014", fg="#f2ece1", relief="flat", justify="center")
    euro_entry.pack(side="left", padx=6, ipady=2)

    def _draw_chart(per_day):
        chart.delete("all")
        w = chart.winfo_width() or 700; h = 150
        days = list(per_day.items())
        if not days: return
        mx = max((v for _,v in days), default=0) or 1
        n = len(days); pad = 24; bw = (w - 2*pad) / n
        for i,(d,v) in enumerate(days):
            x0 = pad + i*bw + 3; x1 = pad + (i+1)*bw - 3
            bh = (h - 40) * (v / mx)
            y1 = h - 22; y0 = y1 - bh
            col = "#ff5a5f" if v>0 else "#45392c"
            chart.create_rectangle(x0, y0, x1, y1, fill=col, outline="")
            if v>0: chart.create_text((x0+x1)/2, y0-7, text=str(v), fill="#f2ece1", font=("Bahnschrift",8,"bold"))
            chart.create_text((x0+x1)/2, h-10, text=d[5:], fill="#7a6c58", font=("Bahnschrift",7))

    def refresh_stats():
        from collections import Counter
        from datetime import date, timedelta
        evs = load_events()
        by = Counter(e.get("type") for e in evs)
        errors = sum(by[t] for t in ERROR_EVENTS)
        verified = by.get("PICK_VERIFIED",0)
        overrides = by.get("OVERRIDE",0)
        try: epe = float(euro_var.get())
        except ValueError: epe = 150.0
        if cfg.get("euro_per_error") != epe:          # cfg zapis TYLKO gdy wartosc faktycznie zmieniona
            cfg["euro_per_error"] = epe; save_cfg(cfg)
        card_vals["verified"].config(text=str(verified))
        card_vals["errors"].config(text=str(errors))
        card_vals["loss"].config(text=f"€{int(errors*epe):,}".replace(",", " "))
        card_vals["overrides"].config(text=str(overrides))
        pa_units = sum(int(e.get("qty") or 0) for e in evs if e.get("type") == "PA_LINE")
        card_vals["pa_conf"].config(text=str(by.get("PA_CONFIRMED",0)))
        card_vals["pa_lines"].config(text=str(by.get("PA_LINE",0)))
        card_vals["pa_units"].config(text=str(pa_units))
        card_vals["pa_exports"].config(text=str(by.get("PA_EXPORTED",0)))
        today = date.today()
        per_day = {(today - timedelta(days=i)).isoformat(): 0 for i in range(13,-1,-1)}
        for e in evs:
            if e.get("type") in ERROR_EVENTS:
                d = (e.get("ts","") or "")[:10]
                if d in per_day: per_day[d]+=1
        stats_cache["per_day"] = per_day              # cache dla przerysowan przy resize
        _draw_chart(per_day)
        noise = by.get("NOISE_SCAN",0)
        lines = [
            f"  Wrong item scanned (caught)   {by.get('WRONG_ITEM',0):>5}",
            f"  Over-pick blocked             {by.get('OVERPICK',0):>5}",
            f"  Bulk quantity blocked         {by.get('BULK_BLOCKED',0):>5}",
            "  ─────────────────────────────────",
            f"  Noise scans filtered out      {noise:>5}   (qty/EAN/serial — NOT counted as loss)",
            f"  Picks verified                {verified:>5}",
            f"  Picks closed (with serials)   {by.get('PICK_CLOSED',0):>5}",
            f"  Serials rejected (bad format) {by.get('SERIAL_REJECTED',0):>5}",
            f"  Duplicate serials caught      {by.get('DUP_SERIAL',0):>5}",
            f"  Emergency overrides           {overrides:>5}",
            "  \u2500\u2500 INBOUND \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500",
            f"  Put-away sessions opened      {by.get('PA_SESSION_OPEN',0):>5}",
            f"  Put-aways confirmed           {by.get('PA_CONFIRMED',0):>5}",
            f"  Lines binned / units          {by.get('PA_LINE',0):>5} / {pa_units}",
            f"  Day exports to office         {by.get('PA_EXPORTED',0):>5}",
        ]
        brk.config(text="\n".join(lines))

    btn_fr = tk.Frame(tstat, bg="#1a1714"); btn_fr.pack(pady=(0,10))
    tk.Button(btn_fr, text="🔄 Refresh", command=refresh_stats, bg="#2fb5a8", fg="white",
              relief="flat", font=("Bahnschrift",10,"bold")).pack(side="left", padx=4)
    def open_events():
        try: os.startfile(events_path())
        except Exception:
            try: webbrowser.open(Path(events_path()).as_uri())
            except Exception: pass
    tk.Button(btn_fr, text="📂 Open data file", command=open_events, bg="#5a4a38", fg="#cfc3b0",
              relief="flat", font=("Bahnschrift",9)).pack(side="left", padx=4)

    def clean_history():
        """Retro-klasyfikacja: stare WRONG_ITEM bedace szumem (qty/EAN/serial-podobne) -> NOISE_SCAN.
           Conservative by design: codes in the known-SKU registry REMAIN errors. A .bak backup is written first."""
        evs = load_events()
        if not evs: messagebox.showinfo("Clean history","No events to clean."); return
        known = load_known_skus()
        recls = 0
        for e in evs:
            if e.get("type") != "WRONG_ITEM": continue
            c = (e.get("code") or "").strip()
            if not c: continue
            if any(sku_match(c, s) for s in known): continue      # prawdziwy SKU katalogowy - zostaje bledem
            k = classify_scan(c, [], known)
            if k in ("qty","numeric","serial","unknown"):
                e["type"]="NOISE_SCAN"; e["reason"]=k; e["reclassified"]=True; recls += 1
        if not recls:
            messagebox.showinfo("Clean history","Nothing to reclassify — history is clean."); return
        if not messagebox.askyesno("Clean history",
            f"Reclassify {recls} noisy WRONG_ITEM event(s) as NOISE_SCAN?\n\n"
            "Codes matching the known-SKU registry stay as real errors.\n"
            "A backup (.bak) will be created."): return
        p = events_path()
        try:
            shutil.copy2(p, p + ".bak")
            with open(p, "w", encoding="utf-8") as f:
                for e in evs: f.write(json.dumps(e, ensure_ascii=False) + "\n")
            log_event("HISTORY_CLEANED", reclassified=recls)
            refresh_stats()
            messagebox.showinfo("Clean history", f"Done — {recls} event(s) reclassified as noise.\nBackup: {p}.bak")
        except Exception as ex:
            messagebox.showerror("Clean history", f"Failed: {ex}")
    tk.Button(btn_fr, text="🧹 Clean noise from history", command=clean_history, bg="#f5b342", fg="#2a241e",
              relief="flat", font=("Bahnschrift",9,"bold")).pack(side="left", padx=4)

    stats_cache = {"per_day": {}}
    _resize_job = {"id": None}
    def _on_chart_resize(_e=None):
        """A resize only redraws from cache, debounced, with ZERO disk I/O."""
        if _resize_job["id"]:
            try: chart.after_cancel(_resize_job["id"])
            except Exception: pass
        _resize_job["id"] = chart.after(150, lambda: _draw_chart(stats_cache["per_day"]))
    euro_entry.bind("<Return>", lambda e: refresh_stats())
    chart.bind("<Configure>", _on_chart_resize)
    def _on_tab_changed(_e=None):
        try:
            if nb.tab(nb.select(), "text").strip() == "Impact": refresh_stats()
        except Exception: pass
    nb.bind("<<NotebookTabChanged>>", _on_tab_changed, add="+")   # v3.0 fix: bez add="+" kasowal handler Inbound

    # ---------- TAB: WAREHOUSE (slotting analytics) ----------
    _owners = [o.strip().upper() for o in (cfg.get("analytics_owners") or [])]
    # An empty owner list means everyone: analytics is hidden only when the list
    # actually names someone.
    if not _owners or get_picker() in _owners:
        tware = tk.Frame(nb, bg=UI["bg"]); nb.add(tware, text="Warehouse", group="INSIGHTS", fkey="F7", icon="🏭")
        tk.Label(tware, text="Warehouse Intelligence — slotting analytics", fg=UI["accent"], bg=UI["bg"],
                 font=("Bahnschrift",15,"bold")).pack(pady=(12,1))
        tk.Label(tware, text="Process data only (no operators) · zone costs from 1-line picks · compare zones RELATIVELY",
                 fg=UI["muted"], bg=UI["bg"], font=("Bahnschrift",9)).pack()
        wtxt = tk.Text(tware, bg=UI["panel"], fg=UI["text"], relief="flat", font=("Cascadia Mono",9), wrap="none")
        wtxt.pack(fill="both", expand=True, padx=16, pady=8)
        def _wh_report_text(a):
            s = a["summary"]; L = []
            L.append(f"DATA: {s['picks']} picks · {s['days']} day(s) · {s['units']} units · "
                     f"median 1-line cycle {s['med_cycle_1line']}s (n={s['n_singles']}) · "
                     f"multi-line {s['s_per_line']}s/line · validation {s['med_validation']}s")
            L.append("")
            L.append("TOP ITEMS (ABC — picks · units · home zone)")
            tot_p = sum(p for _,p,_,_ in a["items"]) or 1
            for sku,p,u,z in a["items"][:15]:
                L.append(f"  {sku:<20} {p:>4} picks  {u:>5} u  [{z}]  {100*p/tot_p:>4.1f}%")
            L.append("")
            L.append("ZONE HEAT (scan lines per zone)")
            mx = a["zones"][0][1] if a["zones"] else 1
            for z,n in a["zones"][:12]:
                L.append(f"  {z:<14} {n:>5}  {'█'*max(1,int(24*n/mx))}")
            L.append("")
            L.append("ZONE ACCESS COST (median cycle of 1-line picks — fixed overhead cancels between zones)")
            if a["zone_cost"]:
                for z,(m,n) in sorted(a["zone_cost"].items(), key=lambda kv:-kv[1][0]):
                    L.append(f"  {z:<14} {m:>7.1f}s   (n={n})")
            else:
                L.append("  — need ≥3 single-line picks per zone (collecting…)")
            L.append("")
            L.append("SLOTTING CANDIDATES (hot item in slow zone → consider moving closer)")
            if a["movers"]:
                for sku,p,z,c in a["movers"][:10]:
                    L.append(f"  {sku:<20} {p:>4} picks   zone {z:<8} {c:>6.1f}s")
            else:
                L.append("  — none yet (or not enough 1-line calibration data)")
            L.append("")
            L.append("TOP CO-PICKED PAIRS (slot near each other)")
            for (x,y),n in a["pairs"][:10]:
                L.append(f"  {x:<20} + {y:<20} ×{n}")
            return "\n".join(L)
        def refresh_warehouse():
            tele = load_telemetry()
            wtxt.config(state="normal"); wtxt.delete("1.0","end")
            if not tele:
                wtxt.insert("end","No telemetry yet — data is collected automatically with every verified pick.\n")
            else:
                wtxt.insert("end", _wh_report_text(analyze_warehouse(tele)))
            wtxt.config(state="disabled")
        def export_warehouse():
            tele = load_telemetry()
            if not tele: messagebox.showinfo("Warehouse","No telemetry yet."); return
            a = analyze_warehouse(tele)
            body = html.escape(_wh_report_text(a))
            doc = (f"<!DOCTYPE html><html><head><meta charset='utf-8'><title>Warehouse Report</title>"
                   f"<style>body{{background:#1a1714;color:#f2ece1;font:13px 'Bahnschrift'}}"
                   f"pre{{background:#241f1a;padding:16px;border-radius:8px;font:12px Consolas}}"
                   f"h1{{color:#2fb5a8;font-size:20px}}</style></head><body>"
                   f"<h1>PickCore {APP_VERSION} — Warehouse Intelligence</h1>"
                   f"<p>{datetime.now():%Y-%m-%d %H:%M} · process data only (no operators)</p>"
                   f"<pre>{body}</pre></body></html>")
            _,_,log_dir = ensure_app_folders()
            p = os.path.join(log_dir, f"Warehouse_Report_{datetime.now():%Y-%m-%d}.html")
            Path(p).write_text(doc, encoding="utf-8")
            try: webbrowser.open(Path(p).as_uri())
            except Exception: pass
            feedback_w = f"Report saved: {p}"
            messagebox.showinfo("Warehouse", feedback_w)
        wb = tk.Frame(tware, bg=UI["bg"]); wb.pack(pady=(0,10))
        tk.Button(wb, text="🔄 Analyze", command=refresh_warehouse, bg=UI["accent"], fg="white",
                  relief="flat", font=("Bahnschrift",10,"bold")).pack(side="left", padx=4)
        tk.Button(wb, text="📤 Export HTML report", command=export_warehouse, bg=UI["ok"], fg="white",
                  relief="flat", font=("Bahnschrift",10,"bold")).pack(side="left", padx=4)
        def export_star():
            _,_,log_dir = ensure_app_folders()
            out = os.path.join(log_dir, f"star_schema_{datetime.now():%Y-%m-%d}")
            paths, nfacts = export_star_schema(out)
            if not nfacts:
                messagebox.showinfo("Star schema","No telemetry yet — nothing to export."); return
            try: os.startfile(out)
            except Exception:
                try: webbrowser.open(Path(out).as_uri())
                except Exception: pass
            messagebox.showinfo("Star schema export",
                f"Exported {nfacts} fact rows + 4 dimensions + README.\n\nFolder:\n{out}\n\n"
                "Load the 5 CSVs into Power BI (star) or DuckDB. See README_schema.txt.")
        tk.Button(wb, text="🗄 Export star schema (BI/SQL)", command=export_star, bg=UI["serial"], fg="white",
                  relief="flat", font=("Bahnschrift",10,"bold")).pack(side="left", padx=4)
        def export_heatmap():
            tpl = find_pickmap_template()
            if not tpl:
                messagebox.showinfo("Pick heat map",
                    "pickmap_template.html not found.\n\nPlace your generated template next to "
                    "PickConverter.exe (C:\\converter\\dist\\pickmap_template.html) and try again."); return
            heat, unmapped = pickmap_heat_data()
            if not heat and not unmapped:
                messagebox.showinfo("Pick heat map","No telemetry yet — verify a few picks first."); return
            _,_,log_dir = ensure_app_folders()
            out = os.path.join(log_dir, f"PickMap_Heat_{datetime.now():%Y-%m-%d}.html")
            try:
                build_pickmap_heat_html(tpl, heat, unmapped, out)
                webbrowser.open(Path(out).as_uri())
                logln(f"🗺 Heat map: {sum(heat.values())} visits → {len(heat)} locations ({os.path.basename(out)})")
            except Exception as e:
                messagebox.showerror("Pick heat map", f"Failed: {e}")
        tk.Button(wb, text="🗺 Heat — rack view", command=export_heatmap, bg="#E2641F", fg="white",
                  relief="flat", font=("Bahnschrift",10,"bold")).pack(side="left", padx=4)
        def export_iso():
            heat, unmapped = pickmap_iso_data()
            if not heat and not unmapped:
                messagebox.showinfo("Heat 3D","No telemetry yet — verify a few picks first."); return
            _,_,log_dir = ensure_app_folders()
            out = os.path.join(log_dir, f"PickMap_ISO_{datetime.now():%Y-%m-%d}.html")
            try:
                build_pickmap_iso_html(heat, unmapped, out)
                webbrowser.open(Path(out).as_uri())
                logln(f"🧊 Heat 3D: {sum(heat.values())} visits → {len(heat)} bays ({os.path.basename(out)})")
            except Exception as e:
                messagebox.showerror("Heat 3D", f"Failed: {e}")
        tk.Button(wb, text="🧊 Heat — 3D hall", command=export_iso, bg="#7c4ddf", fg="white",
                  relief="flat", font=("Bahnschrift",10,"bold")).pack(side="left", padx=4)

    # ---------- TAB: LABELS (integracja LabelSelector - launcher + autostart) ----------
    tlab = tk.Frame(nb, bg=UI["bg"]); nb.add(tlab, text="Station kits", group="LABELS", fkey="F4", icon="🏷")
    tk.Label(tlab, text="Label Generator — Zebra GK420d", fg=UI["accent"], bg=UI["bg"],
             font=("Bahnschrift",15,"bold")).pack(pady=(14,2))
    tk.Label(tlab, text=f"KIT LABELS (native)   ·   {LABEL_W_MM}×{LABEL_H_MM} mm  |  {LABEL_DPI} DPI  |  Zebra RAW",
             fg=UI["muted"], bg=UI["bg"], font=("Cascadia Mono",9,"bold")).pack()
    _kl = tk.Frame(tlab, bg=UI["bg"]); _kl.pack(pady=6)
    tk.Label(_kl, text="Kit:", fg=UI["muted"], bg=UI["bg"]).grid(row=0,column=0,sticky="e",padx=4)
    kit_combo = ttk.Combobox(_kl, state="readonly", width=34); kit_combo.grid(row=0,column=1,padx=4)
    tk.Label(_kl, text="Copies:", fg=UI["muted"], bg=UI["bg"]).grid(row=0,column=2,sticky="e",padx=4)
    e_copies = tk.Entry(_kl, width=5, justify="center"); e_copies.insert(0,"1"); e_copies.grid(row=0,column=3,padx=4)
    tk.Label(_kl, text="Printer:", fg=UI["muted"], bg=UI["bg"]).grid(row=1,column=0,sticky="e",padx=4,pady=4)
    prn_combo = ttk.Combobox(_kl, state="readonly", width=34,
                             values=(list_printers() or ["[no printers found]"]))
    prn_combo.grid(row=1,column=1,padx=4,pady=4)
    _prns = list_printers()
    _pref = cfg.get("zebra_printer") or next((p for p in _prns if "ZEBRA" in p.upper() or "ZDESIGNER" in p.upper()), "")
    if _pref in _prns: prn_combo.set(_pref)
    elif _prns: prn_combo.current(0)
    lbl_warn = tk.Label(tlab, text="", fg=UI["warn"], bg=UI["bg"], font=("Bahnschrift",9,"bold")); lbl_warn.pack()
    _kits_cache = {"kits": []}
    def _labels_refresh_kits():
        rows=[]; 
        if station["cur"]: rows += [dict(sku=i["sku"], item=i["sku"], desc=i.get("desc",""), order=i.get("order","")) for i in station["cur"]["items"]]
        for pd_ in station["queue"]:
            rows += [dict(item=i["sku"], desc=i.get("desc",""), order=i.get("order","")) for i in pd_["items"]]
        ks = extract_label_kits(rows)
        _kits_cache["kits"]=ks
        kit_combo["values"]=[f'{k["kit_id"]}  ({k["order"]}, {len(k["items"])} items)' for k in ks] or ["[no assembly kits on station]"]
        kit_combo.current(0)
        lbl_warn.config(text="")
    def _labels_print():
        ks=_kits_cache["kits"]
        if not ks: feedback("No kits on station — load a pick with AS orders.", "warn"); _beep("warn"); return
        k=ks[max(0, kit_combo.current())]
        c=validate_copies(e_copies.get())
        if c is None:
            messagebox.showerror("Copies","'Copies' must be a whole number of 1 or more."); return
        zpl, trunc, fit = generate_kit_zpl(k["kit_id"], k["items"], c)
        if trunc and not messagebox.askyesno("Truncated assemblies",
                f"Only {fit}/{len(k['items'])} assemblies fit on the label.\nPrint anyway?"):
            return
        prn = prn_combo.get()
        if not prn or prn.startswith("["):
            f=filedialog.asksaveasfilename(defaultextension=".txt", initialfile=f"{k['kit_id']}.txt",
                                           filetypes=[("Text files","*.txt")])
            if f: Path(f).write_text(zpl, encoding="utf-8"); append_print_log(k["kit_id"], c, "FILE", "SAVED")
            return
        try:
            print_zpl_raw(zpl, prn)
            cfg["zebra_printer"]=prn; save_cfg(cfg)
            append_print_log(k["kit_id"], c, prn, "OK")
            logln(f"🏷 Kit label: {k['kit_id']} ×{c} → {prn}"); _beep("ok")
            feedback(f"🏷 Printed {k['kit_id']} ×{c} ({fit} items).", "done")
        except Exception as e:
            append_print_log(k["kit_id"], c, prn, f"ERROR: {e}")
            messagebox.showerror("Print error", str(e)); _beep("err")
    _kb = tk.Frame(tlab, bg=UI["bg"]); _kb.pack(pady=4)
    tk.Button(_kb, text="🔄  Refresh kits", command=_labels_refresh_kits, bg=UI["panel2"], fg=UI["muted"],
              relief="flat", font=("Bahnschrift",9)).pack(side="left", padx=4)
    tk.Button(_kb, text="🚀  PRINT KIT LABEL", command=_labels_print, bg="#28A745", fg="white",
              relief="flat", font=("Bahnschrift",10,"bold"), padx=12, pady=4).pack(side="left", padx=4)
    tk.Frame(tlab, bg=UI["panel2"], height=1).pack(fill="x", padx=30, pady=10)
    tk.Label(tlab, text="Kit labels are generated natively \u2014 no external tool required.",
             fg=UI["muted"], bg=UI["bg"], font=("Bahnschrift",9,"italic")).pack(pady=(4,0))

    # ---------- VIEW: SHIPMENT PDF (etykiety zawartosci kitow z PDF wysylkowego) ----------
    tshp = tk.Frame(nb, bg=UI["bg"]); nb.add(tshp, text="Shipment labels", group="LABELS", fkey="F5", icon="📄")
    tk.Label(tshp, text="Shipment kit labels — parse a shipment PDF", fg=UI["accent"], bg=UI["bg"],
             font=("Bahnschrift",15,"bold")).pack(pady=(14,2))
    tk.Label(tshp, text="main units + *-ASM lines  ·  configurable size / DPI  ·  Zebra RAW",
             fg=UI["muted"], bg=UI["bg"], font=("Cascadia Mono",9)).pack()
    shp = {"kits": {}, "file": ""}
    _sf = tk.Frame(tshp, bg=UI["bg"]); _sf.pack(pady=8)
    shp_file_lbl = tk.Label(_sf, text="No PDF loaded", fg=UI["faint"], bg=UI["bg"], font=("Cascadia Mono",9))
    def shp_load():
        f = filedialog.askopenfilename(title="Select shipment PDF", filetypes=[("PDF","*.pdf")])
        if not f: return
        try:
            kits = parse_shipment_pdf(f)
        except Exception as e:
            messagebox.showerror("PDF error", str(e)); _beep("err"); return
        shp["kits"] = kits; shp["file"] = os.path.basename(f)
        shp_kit["values"] = list(kits.keys()) or ["[no kits found in PDF]"]
        shp_kit.current(0)
        shp_file_lbl.config(text=f'{shp["file"]} — {len(kits)} kit(s)', fg=(UI["ok"] if kits else UI["warn"]))
        logln(f"📄 Shipment PDF: {shp['file']} → {len(kits)} kit(s)")
        shp_preview()
    tk.Button(_sf, text="📂  Load shipment PDF", command=shp_load, bg=UI["accent"], fg="white",
              relief="flat", font=("Bahnschrift",10,"bold"), padx=12, pady=5).pack(side="left", padx=6)
    shp_file_lbl.pack(side="left", padx=8)
    _sp = tk.Frame(tshp, bg=UI["bg"]); _sp.pack(pady=4)
    tk.Label(_sp, text="Kit:", fg=UI["muted"], bg=UI["bg"]).grid(row=0, column=0, sticky="e", padx=4)
    shp_kit = ttk.Combobox(_sp, state="readonly", width=40, font=("Cascadia Mono",10))
    shp_kit.grid(row=0, column=1, columnspan=7, padx=4, sticky="w")
    def _shp_field(col, label, default):
        tk.Label(_sp, text=label, fg=UI["muted"], bg=UI["bg"]).grid(row=1, column=col, sticky="e", padx=(10,2), pady=8)
        e = tk.Entry(_sp, width=5, justify="center"); e.insert(0, default)
        e.grid(row=1, column=col+1, pady=8, sticky="w"); return e
    shp_w = _shp_field(0, "W(mm):", "40"); shp_h = _shp_field(2, "H(mm):", "30"); shp_c = _shp_field(4, "Copies:", "1")
    tk.Label(_sp, text="DPI:", fg=UI["muted"], bg=UI["bg"]).grid(row=1, column=6, sticky="e", padx=(10,2))
    shp_dpi = ttk.Combobox(_sp, values=["203","300"], width=5, state="readonly")
    shp_dpi.current(0); shp_dpi.grid(row=1, column=7, sticky="w")
    shp_txt = tk.Text(tshp, height=11, bg=UI["panel"], fg=UI["warn"], relief="flat",
                      font=("Cascadia Mono",9), insertbackground=UI["text"])
    shp_txt.pack(fill="both", expand=True, padx=20, pady=6)
    shp_warn = tk.Label(tshp, text="", fg=UI["warn"], bg=UI["bg"], font=("Bahnschrift",9,"bold")); shp_warn.pack()
    def _shp_params():
        try:
            w, h, c = int(shp_w.get()), int(shp_h.get()), int(shp_c.get())
            if w < 10 or h < 10 or not (1 <= c <= 999): raise ValueError
            return w, h, c, int(shp_dpi.get())
        except Exception:
            return None
    def shp_preview(*_):
        kid = shp_kit.get()
        shp_txt.delete("1.0", "end"); shp_warn.config(text="", fg=UI["warn"])
        if kid not in shp["kits"]: return
        pr = _shp_params()
        if pr is None:
            shp_warn.config(text="W/H ≥ 10 mm, Copies 1-999 — whole numbers only."); return
        w, h, c, dpi = pr
        zpl, shown, total = generate_shipment_zpl(kid, shp["kits"][kid], w, h, dpi, c)
        shp_txt.insert("end", zpl)
        if shown < total:
            shp_warn.config(text=f"⚠ Label fits {shown}/{total} assemblies — '+{total-shown} more' marker on label.")
    shp_kit.bind("<<ComboboxSelected>>", shp_preview)
    shp_dpi.bind("<<ComboboxSelected>>", shp_preview)
    for _e in (shp_w, shp_h, shp_c): _e.bind("<KeyRelease>", shp_preview)
    _sb2 = tk.Frame(tshp, bg=UI["bg"]); _sb2.pack(pady=(2,12))
    tk.Label(_sb2, text="Printer:", fg=UI["muted"], bg=UI["bg"]).pack(side="left", padx=4)
    shp_prn = ttk.Combobox(_sb2, state="readonly", width=34, values=(list_printers() or ["[no printers found]"]))
    _sprns = list_printers()
    _spref = cfg.get("zebra_printer") or next((p for p in _sprns if "ZEBRA" in p.upper() or "ZDESIGNER" in p.upper()), "")
    if _spref in _sprns: shp_prn.set(_spref)
    elif _sprns: shp_prn.current(0)
    shp_prn.pack(side="left", padx=4)
    def shp_out(save_only=False):
        kid = shp_kit.get()
        if kid not in shp["kits"]:
            shp_warn.config(text="Load a shipment PDF and choose a kit first.", fg=UI["warn"]); _beep("warn"); return
        pr = _shp_params()
        if pr is None:
            shp_warn.config(text="W/H ≥ 10 mm, Copies 1-999 — whole numbers only.", fg=UI["warn"]); _beep("warn"); return
        w, h, c, dpi = pr
        zpl, shown, total = generate_shipment_zpl(kid, shp["kits"][kid], w, h, dpi, c)
        if shown < total and not messagebox.askyesno("Truncated assemblies",
                f"The label fits {shown}/{total} assemblies (+ marker).\nContinue?"):
            return
        prn = shp_prn.get()
        if save_only or not prn or prn.startswith("["):
            f = filedialog.asksaveasfilename(defaultextension=".txt", initialfile=f"{kid}.txt",
                                             filetypes=[("Text files","*.txt")])
            if f:
                Path(f).write_text(zpl, encoding="utf-8")
                append_print_log(kid, c, "FILE", "SAVED(SHIPMENT)")
                logln(f"📄 Shipment label saved: {kid} ×{c}")
            return
        try:
            print_zpl_raw(zpl, prn)
            cfg["zebra_printer"] = prn; save_cfg(cfg)
            append_print_log(kid, c, prn, "OK(SHIPMENT)")
            logln(f"📄 Shipment label: {kid} ×{c} → {prn}"); _beep("ok")
            shp_warn.config(text=f"Printed {kid} ×{c} ({shown} items).", fg=UI["ok"])
        except Exception as e:
            append_print_log(kid, c, prn, f"ERROR: {e}")
            messagebox.showerror("Print error", str(e)); _beep("err")
    tk.Button(_sb2, text="🚀  PRINT LABEL", command=shp_out, bg="#28A745", fg="white",
              relief="flat", font=("Bahnschrift",10,"bold"), padx=14, pady=5).pack(side="left", padx=6)
    tk.Button(_sb2, text="💾  Save .txt", command=lambda: shp_out(True), bg=UI["panel2"], fg=UI["muted"],
              relief="flat", font=("Bahnschrift",9)).pack(side="left", padx=4)


    # ---------- VIEW: RELOCATIONS (bin-to-bin, feeds the ERP reclassification journal) ----------
    trel = tk.Frame(nb, bg=UI["bg"]); nb.add(trel, text="Relocations", group="OPERATIONS", icon="\U0001F500")
    rel = {"lines": [], "pend": {}}
    REL_PATH = os.path.join(ensure_app_folders()[2], "reloc_session.json")
    rel_hdr = tk.Label(trel, text="\U0001F500 RELOCATIONS", fg=UI["accent"], bg=UI["bg"], font=("Bahnschrift",16,"bold"))
    rel_hdr.pack(anchor="w", padx=16, pady=(12,0))
    rel_sub = tk.Label(trel, text="", fg=UI["muted"], bg=UI["bg"], font=("Cascadia Mono",9)); rel_sub.pack(anchor="w", padx=16)
    _rtop = tk.Frame(trel, bg=UI["bg"]); _rtop.pack(fill="x", padx=16, pady=(6,2))
    tk.Label(_rtop, text="Scan:", fg=UI["muted"], bg=UI["bg"], font=("Bahnschrift",10,"bold")).pack(side="left")
    rel_scan_var = tk.StringVar()
    rel_scan = tk.Entry(_rtop, textvariable=rel_scan_var, width=24, bg="#120f0d", fg=UI["ok"],
                        insertbackground=UI["ok"], relief="flat", font=("Cascadia Mono",13,"bold"))
    rel_scan.pack(side="left", padx=8, ipady=3)
    rel_pend_lbl = tk.Label(_rtop, text="", fg=UI["warn"], bg=UI["bg"], font=("Cascadia Mono",10,"bold"))
    rel_pend_lbl.pack(side="left", padx=10)
    _rwrap = tk.Frame(trel, bg=UI["panel"]); _rwrap.pack(fill="both", expand=True, padx=16, pady=8)
    _rcv = tk.Canvas(_rwrap, bg=UI["panel"], highlightthickness=0)
    _rsb = ttk.Scrollbar(_rwrap, orient="vertical", command=_rcv.yview)
    _rin = tk.Frame(_rcv, bg=UI["panel"])
    _rcv_win = _rcv.create_window((0,0), window=_rin, anchor="nw")
    def _rel_sync_scroll(_=None):
        try:
            ch, cw = _rcv.winfo_height(), _rcv.winfo_width()
            iw, ih = _rin.winfo_reqwidth(), _rin.winfo_reqheight()
            _rcv.itemconfigure(_rcv_win, width=max(cw, iw))
            _rcv.configure(scrollregion=(0, 0, max(cw, iw), max(ih, ch)))
            if ih <= ch: _rcv.yview_moveto(0.0)
        except Exception: pass
    _rin.bind("<Configure>", _rel_sync_scroll); _rcv.bind("<Configure>", _rel_sync_scroll)
    _rcv.configure(yscrollcommand=_rsb.set)
    _rcv.pack(side="left", fill="both", expand=True); _rsb.pack(side="right", fill="y")
    rel_status = tk.Label(trel, text="Scan ITEM \u2192 FROM bin \u2192 TO bin (or type + ENTER). Edit qty in the grid.",
                          bg="#241f1a", fg=UI["muted"], font=("Bahnschrift",11,"bold"), anchor="w", padx=12, pady=8)
    rel_status.pack(fill="x", padx=16, pady=(0,4))
    def rel_feedback(msg, kind="info"):
        rel_status.config(text=msg, fg={"ok":UI["ok"],"err":"#ff6b6b","warn":"#f5b342","info":UI["muted"]}.get(kind, UI["muted"]))
    def _rel_counts():
        rel_hdr.config(text=f"\U0001F500 RELOCATIONS \u00B7 {len(rel['lines'])} moves")
        rel_sub.config(text="click a cell = copy for BC \u00B7 journal columns: Item \u00B7 Bin \u00B7 New Bin \u00B7 Qty")
    def rel_save():
        try: Path(REL_PATH).write_text(json.dumps({"lines": rel["lines"]}, indent=1), encoding="utf-8")
        except Exception: pass
    def _rel_copy(val, what):
        """FIX: update_idletasks() does NOT process clipboard events, so Windows saw an empty or stale buffer.
           A full update() is required. The value always goes out as plain text, trimmed."""
        try:
            val = str(val).strip()
            root.clipboard_clear(); root.clipboard_append(val); root.update()
            rel_feedback(f"\U0001F4CB {what}: {val} \u2014 paste into BC.", "ok"); _beep("ok")
        except Exception as e:
            rel_feedback(f"Clipboard error: {e}", "err")
    def rel_rebuild():
        for w in _rin.winfo_children(): w.destroy()
        hdrr = tk.Frame(_rin, bg=UI["panel"]); hdrr.pack(fill="x", pady=(6,2))
        for txt, wd in (("#",3),("ITEM",15),("DESCRIPTION",22),("LOC",8),("NEW LOC",8),("FROM",12),("TO",12),("QTY",5)):
            tk.Label(hdrr, text=txt, fg=UI["faint"], bg=UI["panel"], font=("Cascadia Mono",9,"bold"),
                     width=wd, anchor="w").pack(side="left", padx=6)
        for i, l in enumerate(rel["lines"]):
            row = tk.Frame(_rin, bg=UI["panel"]); row.pack(fill="x", pady=1)
            tk.Label(row, text=str(i+1), fg=UI["faint"], bg=UI["panel"], font=("Cascadia Mono",10),
                     width=4, anchor="w").pack(side="left", padx=6)
            sk = tk.Label(row, text=l["sku"], fg=UI["text"], bg=UI["panel"], font=("Cascadia Mono",11,"bold"),
                          width=16, anchor="w", cursor="hand2")
            sk.pack(side="left", padx=6)
            sk.bind("<Button-1>", lambda e, v=l["sku"]: _rel_copy(v, "Item No."))
            dl = tk.Label(row, text=(l.get("desc") or sku_desc(l["sku"]) or "\u2014")[:22], fg=UI["muted"],
                          bg=UI["panel"], font=("Bahnschrift",9), width=22, anchor="w", cursor="hand2")
            dl.pack(side="left", padx=6)
            dl.bind("<Button-1>", lambda e, l=l: _rel_copy(l.get("desc") or sku_desc(l["sku"]), "Description"))
            for key, name, w in (("loc","Location Code",8), ("nloc","New Location Code",8)):
                lv = tk.StringVar(value=l.get(key) or (cfg.get("default_location") or "MAIN"))
                le = tk.Entry(row, textvariable=lv, width=w, bg=UI["panel2"], fg=UI["text"],
                              insertbackground=UI["text"], relief="flat", font=("Cascadia Mono",10))
                le.pack(side="left", padx=6)
                def _lcommit(_e=None, l=l, key=key, lv=lv, name=name, copy=False):
                    l[key] = (lv.get() or "").strip().upper() or "MAIN"; lv.set(l[key]); rel_save()
                    if copy: _rel_copy(l[key], name)
                    return "break"
                le.bind("<Return>", lambda e, f=_lcommit: f(copy=True))
                le.bind("<FocusOut>", lambda e, f=_lcommit: f(copy=False))
            qv = tk.StringVar(value=str(l.get("qty",1)))
            qe = tk.Entry(row, textvariable=qv, width=5, justify="center", bg=UI["panel2"], fg=UI["text"],
                          insertbackground=UI["text"], relief="flat", font=("Cascadia Mono",11))
            qe.pack(side="left", padx=6)
            def _qcommit(_e=None, i=i, qv=qv, copy=True):
                """BUGFIX: the clipboard is written ONLY on an explicit ENTER.
                   Wczesniej wisialo tez na <FocusOut>, wiec klik w lokacje wygladal tak:
                   klik kopiuje bin -> pole ilosci traci fokus -> handler NADPISUJE schowek qty.
                   The effect was that the quantity pasted every time and the location code never did."""
                try:
                    q = int(qv.get())
                    if q < 1: raise ValueError
                    rel["lines"][i]["qty"] = q; rel_save()
                    if copy: _rel_copy(str(q), "Quantity")
                except Exception:
                    qv.set(str(rel["lines"][i].get("qty",1))); rel_feedback("Qty must be a number >= 1.","warn")
                return "break"
            qe.bind("<Return>", _qcommit)
            qe.bind("<FocusOut>", lambda e, f=_qcommit: f(copy=False))   # zapis bez dotykania schowka
            for key, name in (("frm","Bin Code"),("to","New Bin Code")):
                lb = tk.Label(row, text=l[key], fg=(UI["ok"] if key=="to" else UI["text"]), bg=UI["panel"],
                              font=("Cascadia Mono",11), width=13, anchor="w", cursor="hand2")
                lb.pack(side="left", padx=6)
                lb.bind("<Button-1>", lambda e, v=l[key], n=name: _rel_copy(v, n))
        _rel_counts(); _rel_sync_scroll()
    def _rel_pend_paint():
        p = rel["pend"]
        rel_pend_lbl.config(text=(f"{p.get('sku','?')} \u00B7 FROM {p.get('frm','\u2014')} \u00B7 TO \u2026" if p else ""))
    def rel_feed(code):
        code = (code or "").strip().upper()
        if not code: return
        # Location codes: each has its own barcode and sets the Location Code,
        # not the bin. The list is configurable so extra codes need no code change.
        locs = [x.strip().upper() for x in
                (cfg.get("location_codes") or "MAIN,DISPATCHED,QUARANTINE,TRANSIT").split(",") if x.strip()]
        if code in locs:
            if rel["lines"] and not rel["pend"]:
                rel["lines"][-1]["loc"] = code; rel["lines"][-1]["nloc"] = code
                rel_save(); rel_rebuild()
                rel_feedback(f"\U0001F4CD Location {code} \u2014 applied to last move.", "ok")
            else:
                rel["pend"]["loc"] = code; rel["pend"]["nloc"] = code
                rel_feedback(f"\U0001F4CD Location {code} \u2014 will be used for this move.", "info")
            _beep("ok"); _rel_pend_paint(); return
        b, okb = format_bin_input(code)
        p = rel["pend"]
        if okb:
            if not p.get("sku"):
                rel_feedback(f"\u2715 {b}: scan the ITEM first.", "warn"); _beep("warn"); return
            if not p.get("frm"):
                p["frm"] = b; rel_feedback(f"FROM {b} \u2014 now scan the destination bin.", "info"); _beep("line")
            else:
                rel["lines"].append({"sku": p["sku"], "frm": p["frm"], "to": b, "qty": 1,
                                     "desc": sku_desc(p["sku"]),
                                     "loc": p.get("loc") or (cfg.get("default_location") or "MAIN"),
                                     "nloc": p.get("nloc") or p.get("loc") or (cfg.get("default_location") or "MAIN")})
                log_event("RELOC_LINE", sku=p["sku"], frm=p["frm"], to=b, qty=1)
                done = rel["lines"][-1]; rel["pend"] = {}
                rel_save(); rel_rebuild()
                rel_feedback(f"\u2705 {done['sku']}: {done['frm']} \u2192 {done['to']} (qty 1 \u2014 edit in the grid).", "ok")
                _beep("ok")
        else:
            if p.get("sku") and not p.get("to"):
                logln(f"\u26A0 Reloc: {p['sku']} abandoned (incomplete) \u2014 new item {code}")
            rel["pend"] = {"sku": code}
            rel_feedback(f"ITEM {code} \u2014 scan the source bin (FROM).", "info"); _beep("line")
        _rel_pend_paint()
    def _rel_scan_enter(_=None):
        c = rel_scan_var.get().strip(); rel_scan_var.set("")
        if c: rel_feed(c)
        rel_scan.focus_set(); return "break"
    rel_scan.bind("<Return>", _rel_scan_enter)
    _rbot = tk.Frame(trel, bg=UI["bg"]); _rbot.pack(pady=(2,10))
    def _rel_tsv(l):
        """A row to paste into the ERP journal. THE ORDER IS CONFIGURABLE, because the ERP allows
           personalizowac uklad kolumn i rozni sie miedzy uzytkownikami/stacjami.
           Klucze: sku, desc, loc, nloc, bin, nbin, qty."""
        vals = {"sku": l.get("sku",""),
                "desc": l.get("desc","") or sku_desc(l.get("sku","")),
                "loc": l.get("loc","MAIN"),
                "nloc": l.get("nloc", l.get("loc","MAIN")),
                "bin": l.get("frm",""),
                "nbin": l.get("to",""),
                "qty": str(l.get("qty",1))}
        order = [k.strip().lower() for k in
                 (cfg.get("bc_paste_order") or "sku,desc,loc,nloc,bin,nbin,qty").split(",") if k.strip()]
        return "\t".join(vals.get(k, "") for k in order)
    def rel_copy_tsv():
        if not rel["lines"]:
            rel_feedback("Nothing to copy.", "warn"); _beep("warn"); return
        tsv = "\n".join(_rel_tsv(l) for l in rel["lines"])
        _rel_copy(tsv, f"TSV \u00B7 {len(rel['lines'])} lines (full BC journal columns)")
        log_event("RELOC_EXPORTED", n=len(rel["lines"]))
    def rel_remove_last():
        if not rel["lines"]: return
        l = rel["lines"].pop(); rel_save(); rel_rebuild()
        rel_feedback(f"\u21A9 Removed: {l['sku']} {l['frm']} \u2192 {l['to']}.", "info")
    def rel_clear():
        if rel["lines"] and not messagebox.askyesno("Relocations", f"Clear {len(rel['lines'])} lines?"): return
        rel["lines"] = []; rel["pend"] = {}
        rel_save(); rel_rebuild(); _rel_pend_paint()
        log_event("RELOC_CLEARED"); rel_feedback("Session cleared.", "info")
    tk.Button(_rbot, text="\U0001F4CB  Copy TSV for records", command=rel_copy_tsv, bg="#2fb5a8", fg="white",
              relief="flat", font=("Bahnschrift",9,"bold")).pack(side="left", padx=5)
    tk.Button(_rbot, text="\u21A9  Remove last", command=rel_remove_last, bg=UI["panel2"], fg=UI["muted"],
              relief="flat", font=("Bahnschrift",9)).pack(side="left", padx=5)
    tk.Button(_rbot, text="\U0001F9F9  Clear", command=rel_clear, bg=UI["panel2"], fg=UI["muted"],
              relief="flat", font=("Bahnschrift",9)).pack(side="left", padx=5)
    def _rel_on_view(_=None):
        if nb.select() == str(trel): rel_scan.focus_set()
    nb.bind("<<NotebookTabChanged>>", _rel_on_view, add="+")
    try:
        if os.path.exists(REL_PATH):
            rel["lines"] = list(json.loads(Path(REL_PATH).read_text(encoding="utf-8")).get("lines", []))
            if rel["lines"]: logln(f"\U0001F500 Relocations: restored {len(rel['lines'])} lines from the session")
    except Exception: pass
    rel_rebuild()


    # ---------- VIEW: FORWARDER (shipment PDF -> payload for the forwarder portal) ----------
    tsb = tk.Frame(nb, bg=UI["bg"]); nb.add(tsb, text="Forwarder", group="OUTBOUND", fkey="F11", icon="\U0001F69A")
    sb = {"data": None}
    tk.Label(tsb, text="\U0001F69A  Forwarder shipment", fg=UI["accent"], bg=UI["bg"],
             font=("Bahnschrift",15,"bold")).pack(anchor="w", padx=16, pady=(12,0))
    tk.Label(tsb, text="Sales Shipment PDF \u2192 clipboard \u2192 'Fill shipment' bookmark in the portal",
             fg=UI["muted"], bg=UI["bg"], font=("Cascadia Mono",9)).pack(anchor="w", padx=16)

    _sbtop = tk.Frame(tsb, bg=UI["bg"]); _sbtop.pack(fill="x", padx=16, pady=(8,4))
    sb_file = tk.Label(_sbtop, text="No shipment loaded", fg=UI["faint"], bg=UI["bg"], font=("Cascadia Mono",9))
    sb_match = tk.Label(_sbtop, text="", fg=UI["faint"], bg=UI["bg"], font=("Cascadia Mono",9))
    _sbmain = tk.Frame(tsb, bg=UI["bg"]); _sbmain.pack(fill="both", expand=True, padx=16, pady=4)
    _sbL = tk.Frame(_sbmain, bg=UI["panel"]); _sbL.pack(side="left", fill="both", expand=True, padx=(0,6))
    _sbR = tk.Frame(_sbmain, bg=UI["panel"]); _sbR.pack(side="left", fill="both", expand=True, padx=(6,0))

    sbv = {}
    def _sbrow(parent, key, label, w=30):
        f = tk.Frame(parent, bg=UI["panel"]); f.pack(fill="x", padx=10, pady=2)
        tk.Label(f, text=label, bg=UI["panel"], fg=UI["muted"], font=("Bahnschrift",9),
                 width=16, anchor="w").pack(side="left")
        v = tk.StringVar()
        tk.Entry(f, textvariable=v, width=w, bg=UI["panel2"], fg=UI["text"], relief="flat",
                 insertbackground=UI["text"], font=("Cascadia Mono",10)).pack(side="left", fill="x", expand=True)
        sbv[key] = v
        return v

    tk.Label(_sbL, text="DELIVERY ADDRESS", bg=UI["panel"], fg=UI["faint"],
             font=("Bahnschrift",8,"bold")).pack(anchor="w", padx=10, pady=(8,2))
    for k, lab in (("company","Company name"), ("contact","Contact name"), ("address1","Address line 1"),
                   ("postcode","Postal code"), ("city","Town / City"), ("country","Country"),
                   ("telephone","Telephone"), ("email","Email")):
        _sbrow(_sbL, k, lab)
    tk.Label(_sbL, text="REFERENCES", bg=UI["panel"], fg=UI["faint"],
             font=("Bahnschrift",8,"bold")).pack(anchor="w", padx=10, pady=(10,2))
    _sbrow(_sbL, "reference", "Your ref (SO no.)")
    _sbrow(_sbL, "delivery_reference", "Delivery ref")

    tk.Label(_sbR, text="PARCEL", bg=UI["panel"], fg=UI["faint"],
             font=("Bahnschrift",8,"bold")).pack(anchor="w", padx=10, pady=(8,2))
    sb_preset = tk.StringVar(value=PARCEL_PRESETS[2][0])
    sb_dims = {"l": tk.StringVar(value="40"), "w": tk.StringVar(value="30"), "h": tk.StringVar(value="20")}
    def _sb_apply_preset():
        for nm, L, W, H in PARCEL_PRESETS:
            if nm == sb_preset.get():
                sb_dims["l"].set(str(L)); sb_dims["w"].set(str(W)); sb_dims["h"].set(str(H)); break
    for nm, *_ in PARCEL_PRESETS:
        tk.Radiobutton(_sbR, text=nm, variable=sb_preset, value=nm, command=_sb_apply_preset,
                       bg=UI["panel"], fg=UI["text"], selectcolor=UI["panel2"], activebackground=UI["panel"],
                       activeforeground=UI["accent"], font=("Cascadia Mono",10), anchor="w").pack(fill="x", padx=16)
    _sbd = tk.Frame(_sbR, bg=UI["panel"]); _sbd.pack(fill="x", padx=10, pady=(6,2))
    for k, lab in (("l","L"), ("w","W"), ("h","H")):
        tk.Label(_sbd, text=lab, bg=UI["panel"], fg=UI["muted"], font=("Bahnschrift",9)).pack(side="left", padx=(0,2))
        tk.Entry(_sbd, textvariable=sb_dims[k], width=5, justify="center", bg=UI["panel2"], fg=UI["text"],
                 relief="flat", insertbackground=UI["text"], font=("Cascadia Mono",10)).pack(side="left", padx=(0,8))
    tk.Label(_sbd, text="cm", bg=UI["panel"], fg=UI["faint"], font=("Bahnschrift",8)).pack(side="left")
    sb_weight = tk.StringVar(); sb_parcels = tk.StringVar(value="1"); sb_value = tk.StringVar(value="0")
    for var, lab, col in ((sb_weight,"Weight (kg)",UI["ok"]), (sb_parcels,"Parcels",UI["text"]),
                          (sb_value,"Value",UI["text"])):
        f = tk.Frame(_sbR, bg=UI["panel"]); f.pack(fill="x", padx=10, pady=2)
        tk.Label(f, text=lab, bg=UI["panel"], fg=UI["muted"], font=("Bahnschrift",9),
                 width=12, anchor="w").pack(side="left")
        tk.Entry(f, textvariable=var, width=8, justify="center", bg=UI["panel2"], fg=col, relief="flat",
                 insertbackground=UI["text"], font=("Cascadia Mono",11,"bold")).pack(side="left")
    tk.Label(_sbR, text="Payload preview", bg=UI["panel"], fg=UI["faint"],
             font=("Bahnschrift",8,"bold")).pack(anchor="w", padx=10, pady=(10,2))
    sb_prev = tk.Text(_sbR, height=10, bg=UI["bg"], fg=UI["warn"], relief="flat",
                      font=("Cascadia Mono",8), insertbackground=UI["text"])
    sb_prev.pack(fill="both", expand=True, padx=10, pady=(0,10))

    def _sb_collect():
        d = dict(sb["data"] or {})
        for k in ("company","contact","address1","postcode","city","country"):
            d[k] = sbv[k].get().strip()
        return build_payload(d, {
            "reference": sbv["reference"].get().strip(),
            "delivery_reference": sbv["delivery_reference"].get().strip(),
            "telephone": sbv["telephone"].get().strip(), "email": sbv["email"].get().strip(),
            "weight": sb_weight.get().strip(), "parcels": sb_parcels.get().strip() or "1",
            "value": sb_value.get().strip() or "0",
            "length": sb_dims["l"].get().strip(), "width": sb_dims["w"].get().strip(),
            "height": sb_dims["h"].get().strip()})

    def _sb_preview(*_):
        if not sb["data"]: return
        sb_prev.delete("1.0","end"); sb_prev.insert("end", json.dumps(_sb_collect(), indent=1, ensure_ascii=False))

    def _sb_load(path=None):
        # W Downloads wydruki z BC (Sales - Shipment - ...pdf) leza WYMIESZANE z
        # etykietami kuriera (label_*.pdf) pobieranymi po utworzeniu przesylki.
        # Wybranie etykiety dawalo pusty formularz i mylacy komunikat o kartotece.
        # The dialog therefore opens in Downloads with the newest export ALREADY SELECTED.
        # Filtering uses initialfile rather than a filetypes pattern, because the native
        # Windows dialog does not handle patterns beyond the extension reliably.
        _kw = {}
        _dl = os.path.join(os.path.expanduser("~"), "Downloads")
        if os.path.isdir(_dl):
            _kw["initialdir"] = _dl
            try:
                _kand = sorted((os.path.getmtime(os.path.join(_dl, x)), x)
                               for x in os.listdir(_dl)
                               if x.lower().endswith(".pdf") and "shipment" in x.lower())
                if _kand:
                    _kw["initialfile"] = _kand[-1][1]
            except OSError as e:
                logln(f"\u26A0 Forwarder: nie moge przejrzec {_dl}: {str(e)[:60]}")
        f = path or filedialog.askopenfilename(title="Sales Shipment PDF",
                                               filetypes=[("PDF","*.pdf")], **_kw)
        if not f: return
        try:
            d = parse_shipment(f)
        except Exception as e:
            messagebox.showerror("Shipment", str(e)); return
        d = enrich_from_customers(d)     # telefon / e-mail / kontakt z kartoteki, adres z DOKUMENTU
        # A document the parser does not recognise used to yield an EMPTY form and red
        # "customer not found in database" - komunikat o ZLYM problemie. Operator
        # the operator then hunts for a fault in the customer file, when the real cause is
        # a courier waybill (Shipper's / Consignee's Name columns) was loaded instead of
        # a sales shipment (Delivery Address block). The message says so explicitly.
        if not (str(d.get("company") or "").strip() or str(d.get("shipment_no") or "").strip()):
            _nazwa = os.path.basename(f)
            sb_match.config(text="\u26A0 to nie wyglada na Sales Shipment z BC", fg=UI["warn"])
            sb_file.config(text=_nazwa, fg=UI["warn"])
            logln(f"\u26A0 Forwarder: {_nazwa} \u2014 no 'Delivery Address' block, nothing for the parser to read")
            messagebox.showwarning("Forwarder",
                f"Could not read the delivery address from:\n{_nazwa}\n\n"
                "The parser looks for a 'Delivery Address' block and did not find one.\n"
                "Previously loaded data has been kept - nothing was overwritten.\n\n"
                "To see the document structure without sending the file anywhere:\n"
                "  py -3.14 C:\\converter\\tests\\sbdiag.py \"<this file>\"")
            return
        sb["data"] = d
        for k in ("company","contact","address1","postcode","city","country"):
            sbv[k].set(d.get(k,""))
        for k in ("telephone","email"):
            if d.get(k): sbv[k].set(d[k])
        _m = d.get("match") or {}
        if _m.get("found"):
            _txt = f'\u2713 {_m["no"]} \u00b7 {_m["name"][:28]}' + (" \u00b7 different delivery address" if _m.get("address_differs") else "")
            sb_match.config(text=_txt, fg=(UI["warn"] if _m.get("address_differs") else UI["ok"]))
            logln(f'\U0001F465 Customer matched: {_m["no"]} ({_m["how"]})'
                  + (" \u2014 delivery address differs from the card (dropshipment?)" if _m.get("address_differs") else ""))
        else:
            sb_match.config(text="\u2717 customer not found in database", fg=UI["faint"])
        sbv["reference"].set(d.get("sales_order",""))
        sbv["delivery_reference"].set(d.get("your_order",""))
        sb_file.config(text=f'{d.get("shipment_no","")} \u00b7 {d.get("company","")}', fg=UI["ok"])
        logln(f'\U0001F69A Forwarder: {d.get("shipment_no","")} \u2192 {d.get("company","")}, {d.get("country","")}')
        _sb_preview()

    def _sb_copy():
        if not sb["data"]:
            messagebox.showwarning("Forwarder", "Load a Sales Shipment PDF first."); return
        pl = _sb_collect()
        if not pl.get("weight_kg") or pl["weight_kg"] == "0":
            if not messagebox.askyesno("Forwarder", "Weight is empty. Copy anyway?"): return
        try:
            root.clipboard_clear(); root.clipboard_append(json.dumps(pl, ensure_ascii=False)); root.update()
            log_event("FORWARDER_PAYLOAD", shipment=sb["data"].get("shipment_no",""),
                      country=pl.get("country",""), parcels=pl.get("parcels",1))
            logln("\U0001F4CB Forwarder payload copied \u2014 open the portal and click 'Fill shipment'")
            _beep("ok")
        except Exception as e:
            messagebox.showerror("Forwarder", str(e))

    tk.Button(_sbtop, text="\U0001F4C2  Load shipment PDF", command=lambda: _sb_load(), bg=UI["accent"],
              fg="white", relief="flat", font=("Bahnschrift",10,"bold"), padx=12, pady=5).pack(side="left")
    sb_file.pack(side="left", padx=10)
    sb_match.pack(side="left", padx=6)
    def _sb_setup():
        """Strona instalacyjna zakladki. Serwujemy ja z wlasnego serwera - dzieki temu
           przeciagniecie linku na pasek dziala tak samo na kazdej stacji, bez dotykania
           the Chrome profile (bookmark file checksum, sync, multiple profiles)."""
        url = f"http://127.0.0.1:{int(cfg.get('http_port') or 8080)}/setup"
        chrome = find_chrome()
        if not cfg.get("http_serve"):
            messagebox.showwarning("Bookmark setup",
                "Enable 'Serve live view over LAN' in Settings -> Integrations first,\n"
                "then restart PickCore. The setup page is served by that same server.")
            return
        try:
            if chrome:
                subprocess.Popen([chrome, url])
                logln("\U0001F4CE Bookmark setup opened in Chrome")
            else:
                webbrowser.open(url)
                logln("\u26A0 Chrome not found \u2014 setup page opened in the default browser")
        except Exception as e:
            messagebox.showerror("Bookmark setup", f"{e}\n\nOpen manually: {url}")
    tk.Button(_sbtop, text="\U0001F4CE  Set up bookmark", command=_sb_setup, bg="#5a4a38", fg="#cfc3b0",
              relief="flat", font=("Bahnschrift",9), padx=10, pady=5).pack(side="right", padx=8)
    tk.Button(_sbtop, text="\U0001F4CB  Copy payload", command=_sb_copy, bg="#28a745", fg="white",
              relief="flat", font=("Bahnschrift",10,"bold"), padx=12, pady=5).pack(side="right")
    for _v in list(sbv.values()) + [sb_weight, sb_parcels, sb_value] + list(sb_dims.values()):
        _v.trace_add("write", _sb_preview)

    # ---------- TAB: CUSTOMERS ----------
    t2 = tk.Frame(nb, bg="#1a1714"); nb.add(t2, text="Customers", group="SYSTEM", icon="👥")
    tk.Label(t2, text="Customer database  (code → full name + reminder)", fg="#f2ece1", bg="#1a1714",
             font=("Bahnschrift",12,"bold")).pack(pady=(14,8))
    ctree = ttk.Treeview(t2, columns=("code","name","reminder"), show="headings", height=10)
    for c,(h_,w_) in {"code":("Code",90),"name":("Full name",240),"reminder":("Reminder",330)}.items():
        ctree.heading(c,text=h_); ctree.column(c,width=w_,anchor="w")
    ctree.pack(fill="both", expand=True, padx=14, pady=6)
    def refresh_customers():
        ctree.delete(*ctree.get_children())
        for code,info in sorted(customers.items()):
            ctree.insert("","end",iid=code,values=(code,info.get("name",""),info.get("reminder","")))
    form = tk.Frame(t2, bg="#1a1714"); form.pack(fill="x", padx=14, pady=6)
    tk.Label(form,text="Code",fg="#a99a84",bg="#1a1714").grid(row=0,column=0,sticky="w")
    tk.Label(form,text="Full name",fg="#a99a84",bg="#1a1714").grid(row=0,column=1,sticky="w")
    tk.Label(form,text="Reminder (optional)",fg="#a99a84",bg="#1a1714").grid(row=0,column=2,sticky="w")
    e_code=tk.Entry(form,width=12); e_code.grid(row=1,column=0,padx=2)
    e_name=tk.Entry(form,width=30); e_name.grid(row=1,column=1,padx=2)
    e_rem=tk.Entry(form,width=42);  e_rem.grid(row=1,column=2,padx=2)
    def add_customer():
        code=e_code.get().strip().upper()
        if not code: return
        customers[code]={"name":e_name.get().strip(),"reminder":e_rem.get().strip()}
        save_customers(customers); refresh_customers()
        e_code.delete(0,"end"); e_name.delete(0,"end"); e_rem.delete(0,"end")
    def del_customer():
        sel=ctree.selection()
        if sel and sel[0] in customers:
            del customers[sel[0]]; save_customers(customers); refresh_customers()
    def load_to_form(_=None):
        sel=ctree.selection()
        if sel and sel[0] in customers:
            info=customers[sel[0]]
            e_code.delete(0,"end"); e_code.insert(0,sel[0])
            e_name.delete(0,"end"); e_name.insert(0,info.get("name",""))
            e_rem.delete(0,"end");  e_rem.insert(0,info.get("reminder",""))
    ctree.bind("<<TreeviewSelect>>", load_to_form)
    def import_customers():
        path = filedialog.askopenfilename(title="Import customers (Excel or CSV)",
                filetypes=[("Excel / CSV","*.xlsx *.xls *.xlsm *.csv"),
                           ("Excel","*.xlsx *.xls *.xlsm"),("CSV","*.csv"),("All files","*.*")])
        if not path: return
        try:
            imported = parse_customer_file(path)
        except ImportError:
            messagebox.showerror("Import failed",
                "Reading Excel needs the 'openpyxl' module which isn't bundled.\n"
                "Save the file as CSV and import that instead."); return
        except Exception as e:
            messagebox.showerror("Import failed", f"Could not read the file:\n{e}"); return
        if not imported:
            messagebox.showwarning("Import","No customers found (expected code in column A, name in column B)."); return
        added = updated = 0
        for code, name in imported.items():
            if code in customers:
                customers[code]["name"] = name          # zachowaj istniejacy reminder
                customers[code].setdefault("reminder","")
                updated += 1
            else:
                customers[code] = {"name": name, "reminder": ""}
                added += 1
        save_customers(customers); refresh_customers()
        messagebox.showinfo("Import complete",
            f"Imported {len(imported)} customers.\n"
            f"New: {added}    Updated: {updated}\n\nExisting reminders were preserved.")

    btns=tk.Frame(t2,bg="#1a1714"); btns.pack(pady=6)
    tk.Button(btns,text="📥 Import Excel / CSV",command=import_customers,bg="#2fb5a8",fg="white",relief="flat",
              font=("Bahnschrift",10,"bold")).pack(side="left",padx=4)
    tk.Button(btns,text="Add / Update",command=add_customer,bg="#27c66d",fg="white",relief="flat").pack(side="left",padx=4)
    tk.Button(btns,text="Delete selected",command=del_customer,bg="#ef4655",fg="white",relief="flat").pack(side="left",padx=4)

    # ---------- TAB: SETTINGS ----------
    t3 = tk.Frame(nb, bg="#1a1714"); nb.add(t3, text="Settings", group="SYSTEM", fkey="F9", icon="⚙")
    def fld(parent,label,key,row):
        tk.Label(parent,text=label,fg="#a99a84",bg="#1a1714",font=("Bahnschrift",10)).grid(row=row,column=0,sticky="w",pady=6)
        var=tk.StringVar(value=cfg.get(key,""))
        ent=tk.Entry(parent,textvariable=var,width=46); ent.grid(row=row,column=1,padx=6)
        def browse():
            d=filedialog.askdirectory()
            if d: var.set(d)
        tk.Button(parent,text="Browse",command=browse,bg="#5a4a38",fg="white",relief="flat").grid(row=row,column=2)
        return var
    # Settings grew with every feature and stopped fitting an unmaximised window.
    # Canvas plus scrollbar plus mouse wheel: ALL settings reachable at any window size.
    _s_wrap = tk.Frame(t3, bg="#1a1714"); _s_wrap.pack(fill="both", expand=True)
    _s_cv   = tk.Canvas(_s_wrap, bg="#1a1714", highlightthickness=0)
    _s_sb   = ttk.Scrollbar(_s_wrap, orient="vertical", command=_s_cv.yview)
    _s_cv.configure(yscrollcommand=_s_sb.set)
    _s_cv.pack(side="left", fill="both", expand=True); _s_sb.pack(side="right", fill="y")
    _s_host = tk.Frame(_s_cv, bg="#1a1714")
    _s_win  = _s_cv.create_window((0,0), window=_s_host, anchor="nw")
    def _s_sync(_=None):
        try:
            _s_cv.configure(scrollregion=_s_cv.bbox("all"))
            _s_cv.itemconfigure(_s_win, width=max(_s_cv.winfo_width(), _s_host.winfo_reqwidth()))
        except Exception: pass
    _s_host.bind("<Configure>", _s_sync); _s_cv.bind("<Configure>", _s_sync)
    def _s_wheel(e):
        if nb.select() == str(t3) and _s_host.winfo_reqheight() > _s_cv.winfo_height():
            _s_cv.yview_scroll(int(-e.delta/120), "units")
    _s_cv.bind_all("<MouseWheel>", _s_wheel, add="+")

    # --- pasek zakladek wewnetrznych ---------------------------------------------------
    # All controls sit on ONE grid with unique row numbers, so pages are built by
    # implemented with grid_remove()/grid() over row ranges. grid_remove preserves the
    # hiding rows rather than rebuilding: grid_remove() keeps the cell config, so returning is lossless.
    SET_PAGES = [
        ("Station",     "\U0001F5A5", list(range(0, 8))),
        ("Integrations","\U0001F50C", list(range(8, 22))),
        ("Data",        "\U0001F5C4", [22, 26, 27, 28, 30]),
        ("Deployment",  "\U0001F4E6", [23, 24, 25, 29]),
    ]
    _tabbar = tk.Frame(_s_host, bg="#1a1714")
    _tabbar.pack(fill="x", padx=18, pady=(14, 0))
    _tab_btns = {}
    _set_page = {"cur": "Station"}

    sf=tk.Frame(_s_host,bg="#1a1714"); sf.pack(padx=18,pady=(10,14),anchor="w")
    sf.grid_columnconfigure(0, minsize=210, pad=6)   # rowna kolumna etykiet
    sf.grid_columnconfigure(1, pad=6)

    _set_map = {"rows": None}

    # --- dimming options that resolve themselves ---------------------------------------
    # Paczka instalacyjna dostarcza dzis Items.csv, customers.csv, sounds\ i bin_contents.csv,
    # and archive folders are created next to the exe. Choosing a path for something that
    # arrives with the update anyway is only an opportunity for error. The row disappears
    # when the setting points exactly at what the package delivers.
    # Empty values are DELIBERATELY left visible: empty means auto-detection FAILED,
    # which is exactly when the field is needed to fix things.
    # watch_dir and update_dir are never hidden: they define the station and are not delivered.
    SETTINGS_AUTO_ROWS = {
        1:  ("archive_dir_pick", "Archive_Picks"),
        2:  ("archive_dir_pa",   "Archive_PutAways"),
        22: ("bin_export_path",  "bin_contents.csv"),
        27: ("items_csv",        "Items.csv"),
        29: ("sound_dir",        "sounds"),
        30: ("customers_csv",    "customers.csv"),
    }
    _set_adv = {"on": bool(cfg.get("settings_show_all"))}

    def _auto_rows():
        """Rows dimmed because the setting matches what the package already provides."""
        base = app_base_dir()
        hide = set()
        for _row, (_key, _name) in SETTINGS_AUTO_ROWS.items():
            # The WHOLE body sits in a try: this runs on EVERY switch of the Settings
            # page, so an exception here would take down the entire page render.
            try:
                val = (cfg.get(_key) or "").strip()
                if not val:
                    continue
                if os.path.normcase(os.path.abspath(val)) == os.path.normcase(os.path.join(base, _name)):
                    hide.add(_row)
            except Exception as e:
                logln(f"⚠ Settings: cannot compare path {_key}: {str(e)[:60]}")
        # The location file usually lives in the update folder rather than next to the exe, hence a separate trace.
        if cfg.get("_bin_export_auto"):
            hide.add(22)
        return hide

    def _show_page(name):
        # NOTE: grid_remove() drops the widget from grid_slaves(), so a hidden page
        # could never be found again. The widget-to-row map is built ONCE, before anything disappears.
        if _set_map["rows"] is None:
            snap = []
            for w in sf.grid_slaves():
                try:
                    snap.append((w, int(w.grid_info().get("row", -1))))
                except Exception:
                    pass
            _set_map["rows"] = snap
        rows = next((r for n, _i, r in SET_PAGES if n == name), [])
        hidden = set() if _set_adv["on"] else _auto_rows()
        for w, r in _set_map["rows"]:
            try:
                if r in rows and r not in hidden:
                    w.grid()
                else:
                    w.grid_remove()
            except Exception:
                pass
        _set_page["cur"] = name
        for n, b in _tab_btns.items():
            on = (n == name)
            b.config(bg=("#2e2720" if on else "#1a1714"),
                     fg=("#2fb5a8" if on else "#7a6c58"),
                     font=("Bahnschrift", 10, "bold" if on else "normal"))
        try:
            _s_cv.yview_moveto(0.0); _s_sync()
        except Exception:
            pass

    for _n, _ico, _rows in SET_PAGES:
        _b = tk.Button(_tabbar, text=f"{_ico}  {_n}", relief="flat", bd=0, cursor="hand2",
                       bg="#1a1714", fg="#7a6c58", font=("Bahnschrift", 10),
                       padx=16, pady=8, command=lambda n=_n: _show_page(n))
        _b.pack(side="left", padx=(0, 2))
        _tab_btns[_n] = _b
    # An escape hatch for dimmed rows. Options are hidden, never removed:
    # when auto-detection fails in a way nobody predicted, there has to be a way
    # to fix it without dictating JSON edits over the phone.
    v_adv = tk.BooleanVar(value=_set_adv["on"])
    def _toggle_adv():
        _set_adv["on"] = bool(v_adv.get())
        cfg["settings_show_all"] = _set_adv["on"]
        save_cfg(cfg)
        _show_page(_set_page["cur"])
    tk.Checkbutton(_tabbar, text="Show all settings", variable=v_adv, command=_toggle_adv,
                   fg="#7a6c58", bg="#1a1714", selectcolor="#2a241e",
                   activebackground="#1a1714", activeforeground="#a99a84",
                   font=("Bahnschrift", 8)).pack(side="right", padx=(0, 8))
    tk.Frame(_s_host, bg="#2e2720", height=1).pack(fill="x", padx=18)
    v_watch=fld(sf,"Watch folder (BC saves PDF here)","watch_dir",0)
    v_arch_pick=fld(sf,"Pick archive folder","archive_dir_pick",1)
    v_arch_pa  =fld(sf,"Put-away archive folder","archive_dir_pa",2)
    tk.Label(sf,text="Pick filter",fg="#a99a84",bg="#1a1714",font=("Bahnschrift",10)).grid(row=3,column=0,sticky="w",pady=6)
    v_filter=tk.StringVar(value=cfg.get("name_filter")); tk.Entry(sf,textvariable=v_filter,width=46).grid(row=3,column=1,padx=6)
    tk.Label(sf,text="Put-away filter",fg="#a99a84",bg="#1a1714",font=("Bahnschrift",10)).grid(row=4,column=0,sticky="w",pady=6)
    v_pafilter=tk.StringVar(value=cfg.get("pa_filter")); tk.Entry(sf,textvariable=v_pafilter,width=46).grid(row=4,column=1,padx=6)
    tk.Label(sf,text="Printer",fg="#a99a84",bg="#1a1714",font=("Bahnschrift",10)).grid(row=5,column=0,sticky="w",pady=6)
    printers=list_printers() or ["(default)"]
    v_printer=tk.StringVar(value=cfg.get("printer") or printers[0])
    ttk.Combobox(sf,textvariable=v_printer,values=printers,width=43,state="readonly").grid(row=5,column=1,padx=6)
    v_boot=tk.BooleanVar(value=autostart_get())
    def _toggle_boot():
        ok,msg = autostart_set(v_boot.get())
        if not ok:
            messagebox.showerror("Autostart", f"Could not change the autostart entry:\n{msg}")
        v_boot.set(autostart_get())          # stan zawsze z rejestru (prawda systemowa)
        if ok:
            logln(f"🚀 Start with Windows: {'ON' if v_boot.get() else 'OFF'} ({_autostart_target()})")
    tk.Checkbutton(sf,text="🚀 Start PickCore with Windows (current user)",variable=v_boot,command=_toggle_boot,
                   fg="#f2ece1",bg="#1a1714",activebackground="#1a1714",activeforeground="#f2ece1",
                   selectcolor="#1c1f27",font=("Bahnschrift",10)).grid(row=6,column=0,columnspan=2,sticky="w",pady=(2,0))
    tk.Label(sf,text="The HKCU\\Run entry points at the current exe - re-toggle after moving the file.",
             fg="#7a6c58",bg="#1a1714",font=("Bahnschrift",8)).grid(row=7,column=0,columnspan=3,sticky="w")
    v_auto=tk.BooleanVar(value=cfg.get("auto_print"))
    tk.Checkbutton(sf,text="Auto-print after conversion",variable=v_auto,fg="#f2ece1",bg="#1a1714",
                   selectcolor="#2a241e",activebackground="#1a1714").grid(row=6,column=1,sticky="w",pady=6)
    v_won=tk.BooleanVar(value=cfg.get("watcher_on"))
    tk.Checkbutton(sf,text="Watcher ON (auto-detect new PDFs)",variable=v_won,fg="#f2ece1",bg="#1a1714",
                   selectcolor="#2a241e",activebackground="#1a1714").grid(row=7,column=1,sticky="w",pady=6)
    # --- DEMO INTEGRATIONS: TC22 (DataWedge IP) + synced folder live board ---
    tk.Label(sf,text="DEMO INTEGRATIONS",fg="#7a6c58",bg="#1a1714",font=("Bahnschrift",8,"bold")).grid(row=8,column=1,sticky="w",pady=(12,0))
    _ipf=tk.Frame(sf,bg="#1a1714"); _ipf.grid(row=8,column=2,sticky="w",padx=(14,0),pady=(12,0))
    _ip_lbl=tk.Label(_ipf,text="",fg=UI["ok"],bg="#1a1714",font=("Cascadia Mono",10,"bold"))
    _ip_lbl.pack(side="left")
    def _ip_refresh():
        ip=local_ip()
        _ip_lbl.config(text=f"THIS PC: {ip}")
        return ip
    def _ip_copy():
        ip=_ip_refresh()
        try:
            root.clipboard_clear()
            root.clipboard_append(f'http://{ip}:{cfg.get("http_port") or 8080}'); root.update()
            logln(f"\U0001F4CB Console URL copied: http://{ip}:{cfg.get('http_port') or 8080}")
        except Exception: pass
    tk.Button(_ipf,text="\U0001F4CB URL",command=_ip_copy,bg="#5a4a38",fg="#cfc3b0",relief="flat",
              font=("Bahnschrift",8)).pack(side="left",padx=6)
    tk.Button(_ipf,text="\u21BB",command=_ip_refresh,bg="#5a4a38",fg="#cfc3b0",relief="flat",
              font=("Bahnschrift",8)).pack(side="left")
    _ip_refresh()
    v_scn=tk.BooleanVar(value=bool(cfg.get("scanner_listen")))
    tk.Checkbutton(sf,text="TC22 scanner listener (TCP \u2014 DataWedge IP output)",variable=v_scn,fg="#f2ece1",bg="#1a1714",
                   selectcolor="#2a241e",activebackground="#1a1714").grid(row=9,column=1,sticky="w",pady=2)
    _scnfr=tk.Frame(sf,bg="#1a1714"); _scnfr.grid(row=10,column=1,sticky="w")
    tk.Label(_scnfr,text="Port:",fg="#a99a84",bg="#1a1714").pack(side="left")
    v_scnp=tk.StringVar(value=str(cfg.get("scanner_port") or 5577))
    tk.Entry(_scnfr,textvariable=v_scnp,width=6,justify="center").pack(side="left",padx=(4,10))
    tk.Label(_scnfr,text="Allow IP (empty = any):",fg="#a99a84",bg="#1a1714").pack(side="left")
    v_scna=tk.StringVar(value=cfg.get("scanner_allow") or "")
    tk.Entry(_scnfr,textvariable=v_scna,width=14).pack(side="left",padx=4)
    tk.Label(_scnfr,text="(restart applies)",fg="#7a6c58",bg="#1a1714",font=("Bahnschrift",8)).pack(side="left",padx=6)
    tk.Label(sf,text="BUSINESS CENTRAL (bin contents lookup)",fg="#7a6c58",bg="#1a1714",
             font=("Bahnschrift",8,"bold")).grid(row=15,column=1,sticky="w",pady=(12,0))
    v_bce=tk.BooleanVar(value=bool(cfg.get("bc_enabled")))
    tk.Checkbutton(sf,text="Query BC for existing bin locations (on put-away load)",variable=v_bce,
                   fg="#f2ece1",bg="#1a1714",selectcolor="#2a241e",activebackground="#1a1714").grid(row=16,column=1,sticky="w",pady=2)
    _bcf=tk.Frame(sf,bg="#1a1714"); _bcf.grid(row=17,column=1,sticky="w")
    def _bcfield(lbl, key, w=26, default=""):
        tk.Label(_bcf,text=lbl,fg="#a99a84",bg="#1a1714",font=("Bahnschrift",8)).pack(side="left",padx=(0,2))
        v=tk.StringVar(value=cfg.get(key) or default)
        tk.Entry(_bcf,textvariable=v,width=w).pack(side="left",padx=(0,8))
        return v
    v_bcurl=_bcfield("Base URL:","bc_base_url",30)
    v_bccmp=_bcfield("Company:","bc_company",14)
    _bcf2=tk.Frame(sf,bg="#1a1714"); _bcf2.grid(row=18,column=1,sticky="w",pady=(4,0))
    tk.Label(_bcf2,text="Tenant:",fg="#a99a84",bg="#1a1714",font=("Bahnschrift",8)).pack(side="left",padx=(0,2))
    v_bctn=tk.StringVar(value=cfg.get("bc_tenant") or ""); tk.Entry(_bcf2,textvariable=v_bctn,width=26).pack(side="left",padx=(0,8))
    tk.Label(_bcf2,text="Client ID:",fg="#a99a84",bg="#1a1714",font=("Bahnschrift",8)).pack(side="left",padx=(0,2))
    v_bccid=tk.StringVar(value=cfg.get("bc_client_id") or ""); tk.Entry(_bcf2,textvariable=v_bccid,width=26).pack(side="left")
    _exf=tk.Frame(sf,bg="#1a1714"); _exf.grid(row=22,column=1,sticky="w",pady=(6,0))
    tk.Label(_exf,text="Bin export file (Excel/Power Query \u2192 CSV):",fg="#a99a84",bg="#1a1714",
             font=("Bahnschrift",8)).pack(side="left",padx=(0,4))
    v_bex=tk.StringVar(value=cfg.get("bin_export_path") or "")
    tk.Entry(_exf,textvariable=v_bex,width=34).pack(side="left")
    def _pick_export():
        f=filedialog.askopenfilename(title="Bin contents export",filetypes=[("CSV/TSV","*.csv;*.txt;*.tsv"),("All","*.*")])
        if f: v_bex.set(f)
    tk.Button(_exf,text="\u2026",command=_pick_export,bg="#5a4a38",fg="#cfc3b0",relief="flat").pack(side="left",padx=4)
    def _test_export():
        n,info = merge_export_into_index(v_bex.get().strip())
        messagebox.showinfo("Bin export", f"{n} SKU loaded\n{info}" if n else f"Nothing loaded\n{info}")
        logln(f"\U0001F4C4 Bin export test: {n} SKU ({info})")
    tk.Button(_exf,text="\u21BB Load now",command=_test_export,bg="#5a4a38",fg="#cfc3b0",relief="flat",
              font=("Bahnschrift",8)).pack(side="left",padx=4)
    _exf2=tk.Frame(sf,bg="#1a1714"); _exf2.grid(row=22,column=2,sticky="w",padx=(14,0),pady=(6,0))
    _pof=tk.Frame(sf,bg="#1a1714"); _pof.grid(row=28,column=1,sticky="w",pady=(8,0))
    tk.Label(_pof,text="BC paste column order:",fg="#a99a84",bg="#1a1714",font=("Bahnschrift",8)).pack(side="left",padx=(0,3))
    v_ord=tk.StringVar(value=cfg.get("bc_paste_order") or "sku,desc,loc,nloc,bin,nbin,qty")
    tk.Entry(_pof,textvariable=v_ord,width=42).pack(side="left")
    tk.Label(_pof,text="keys: sku desc loc nloc bin nbin qty \u2014 match YOUR Item Reclassification Journal layout",
             fg="#7a6c58",bg="#1a1714",font=("Bahnschrift",8)).pack(side="left",padx=6)
    _itf=tk.Frame(sf,bg="#1a1714"); _itf.grid(row=27,column=1,sticky="w",pady=(8,0))
    tk.Label(_itf,text="Item catalogue (No.;Description CSV):",fg="#a99a84",bg="#1a1714",
             font=("Bahnschrift",8)).pack(side="left",padx=(0,4))
    v_items=tk.StringVar(value=cfg.get("items_csv") or "")
    tk.Entry(_itf,textvariable=v_items,width=30).pack(side="left")
    def _pick_items():
        f=filedialog.askopenfilename(title="Item catalogue export",
                                     filetypes=[("CSV/TSV","*.csv;*.txt;*.tsv"),("All","*.*")])
        if f: v_items.set(f); _import_items()
    def _import_items():
        n, info = import_item_descriptions(v_items.get().strip())
        cfg["items_csv"] = v_items.get().strip(); save_cfg(cfg)
        logln(f"\U0001F4D6 Item catalogue: {info}")
        messagebox.showinfo("Item catalogue", info)
        try: rel_rebuild()
        except Exception: pass
    tk.Button(_itf,text="\u2026",command=_pick_items,bg="#5a4a38",fg="#cfc3b0",relief="flat").pack(side="left",padx=4)
    tk.Button(_itf,text="\u2193 Import",command=_import_items,bg="#5a4a38",fg="#cfc3b0",relief="flat",
              font=("Bahnschrift",8)).pack(side="left",padx=4)
    tk.Label(_itf,text="(fills item descriptions at once - no waiting for picks)",fg="#7a6c58",bg="#1a1714",
             font=("Bahnschrift",8)).pack(side="left",padx=6)
    _cuf=tk.Frame(sf,bg="#1a1714"); _cuf.grid(row=30,column=1,sticky="w",pady=(8,0))
    tk.Label(_cuf,text="Customer database (No.;Name;...;Phone;Email):",fg="#a99a84",bg="#1a1714",
             font=("Bahnschrift",8)).pack(side="left",padx=(0,4))
    v_cust=tk.StringVar(value=cfg.get("customers_csv") or "")
    tk.Entry(_cuf,textvariable=v_cust,width=28).pack(side="left")
    def _pick_cust():
        f=filedialog.askopenfilename(title="Customer export (CSV)",
                                     filetypes=[("CSV","*.csv;*.txt"),("All","*.*")])
        if f: v_cust.set(f); _load_cust()
    def _load_cust():
        n,info=load_customer_db(v_cust.get().strip())
        cfg["customers_csv"]=v_cust.get().strip(); save_cfg(cfg)
        logln(f"\U0001F465 Customer database: {info}")
        messagebox.showinfo("Customer database", info)
    tk.Button(_cuf,text="\u2026",command=_pick_cust,bg="#5a4a38",fg="#cfc3b0",relief="flat").pack(side="left",padx=4)
    tk.Button(_cuf,text="\u2193 Load",command=_load_cust,bg="#5a4a38",fg="#cfc3b0",relief="flat",
              font=("Bahnschrift",8)).pack(side="left",padx=4)
    tk.Label(_cuf,text="(fills phone / e-mail / contact; the delivery address always comes from the document)",
             fg="#7a6c58",bg="#1a1714",font=("Bahnschrift",8)).pack(side="left",padx=6)
    _sndf=tk.Frame(sf,bg="#1a1714"); _sndf.grid(row=29,column=1,sticky="w",pady=(8,0))
    tk.Label(_sndf,text="Sound folder (ok/err/warn/line/done):",fg="#a99a84",bg="#1a1714",
             font=("Bahnschrift",8)).pack(side="left",padx=(0,4))
    v_snd=tk.StringVar(value=cfg.get("sound_dir") or "")
    tk.Entry(_sndf,textvariable=v_snd,width=28).pack(side="left")
    def _pick_snd():
        d=filedialog.askdirectory(title="Folder with sound files")
        if d: v_snd.set(d); _snd_scan()
    def _snd_scan():
        probe=dict(cfg); probe["sound_dir"]=v_snd.get().strip()
        pc=[k for k in SOUND_KINDS if find_sound(probe,k,SOUND_EXT_PC)]
        web=[k for k in SOUND_KINDS if find_sound(probe,k,SOUND_EXT_WEB)]
        msg=(f"Folder: {sound_dir(probe)}\n\n"
             f"PC (WAV only): {', '.join(pc) or 'none - generated tones will be used'}\n"
             f"Scanner (mp3/ogg/wav): {', '.join(web) or 'none - generated tones will be used'}\n\n"
             "Name the files after the event: ok, err, warn, line, done.")
        messagebox.showinfo("Sounds", msg); logln(f"\U0001F50A Sounds \u00b7 PC: {len(pc)} \u00b7 scanner: {len(web)}")
    tk.Button(_sndf,text="\u2026",command=_pick_snd,bg="#5a4a38",fg="#cfc3b0",relief="flat").pack(side="left",padx=4)
    tk.Button(_sndf,text="\U0001F50A Check",command=_snd_scan,bg="#5a4a38",fg="#cfc3b0",relief="flat",
              font=("Bahnschrift",8)).pack(side="left",padx=4)
    tk.Label(_sndf,text="(WAV on PC \u00b7 mp3/ogg/wav on the scanner \u00b7 missing file = generated tone)",
             fg="#7a6c58",bg="#1a1714",font=("Bahnschrift",8)).pack(side="left",padx=6)
    _locf=tk.Frame(sf,bg="#1a1714"); _locf.grid(row=26,column=1,sticky="w",pady=(8,0))
    tk.Label(_locf,text="Default location:",fg="#a99a84",bg="#1a1714",font=("Bahnschrift",8)).pack(side="left",padx=(0,3))
    v_dloc=tk.StringVar(value=cfg.get("default_location") or "MAIN")
    tk.Entry(_locf,textvariable=v_dloc,width=10).pack(side="left",padx=(0,10))
    tk.Label(_locf,text="Scannable location codes (comma separated):",fg="#a99a84",bg="#1a1714",
             font=("Bahnschrift",8)).pack(side="left",padx=(0,3))
    v_locs=tk.StringVar(value=cfg.get("location_codes") or "MAIN,DISPATCHED,QUARANTINE,TRANSIT")
    tk.Entry(_locf,textvariable=v_locs,width=34).pack(side="left")
    v_bea=tk.BooleanVar(value=bool(cfg.get("bin_export_auto")))
    tk.Checkbutton(_exf2,text="auto-refresh when older than",variable=v_bea,fg="#f2ece1",bg="#1a1714",
                   selectcolor="#2a241e",activebackground="#1a1714",font=("Bahnschrift",8)).pack(side="left")
    v_bem=tk.StringVar(value=str(cfg.get("bin_export_max_age") or 60))
    tk.Entry(_exf2,textvariable=v_bem,width=4,justify="center").pack(side="left",padx=3)
    tk.Label(_exf2,text="min (on put-away load)",fg="#7a6c58",bg="#1a1714",
             font=("Bahnschrift",8)).pack(side="left")
    _depf=tk.Frame(sf,bg="#1a1714"); _depf.grid(row=23,column=1,sticky="w",pady=(10,0))
    tk.Label(_depf,text="DEPLOYMENT",fg="#7a6c58",bg="#1a1714",font=("Bahnschrift",8,"bold")).pack(side="left",padx=(0,8))
    def _export_profile():
        """Writes the settings template next to the exe. A new station only needs the folder copied."""
        try:
            save_settings()                      # najpierw utrwal to, co na ekranie
            prof = {k: v for k, v in cfg.items() if k not in STATION_KEYS}
            prof["_exported"] = datetime.now().isoformat(timespec="seconds")
            prof["_note"] = ("Template for a new station. Station-specific keys (printers, allow-IP, owners) "
                             "are intentionally omitted - set them locally. No secrets are stored here.")
            Path(profile_path()).write_text(json.dumps(prof, indent=1, ensure_ascii=False), encoding="utf-8")
            logln(f"\U0001F4E6 Deployment profile saved: {profile_path()}")
            messagebox.showinfo("Deployment",
                f"Template saved:\n{profile_path()}\n\n"
                "New station: copy the whole PickCore folder.\n"
                "On first start the settings load automatically;\n"
                "station IP and printers are detected locally.")
        except Exception as e:
            messagebox.showerror("Deployment", str(e))
    tk.Button(_depf,text="\U0001F4E6  Export station profile",command=_export_profile,bg="#5a4a38",
              fg="#cfc3b0",relief="flat",font=("Bahnschrift",9)).pack(side="left")
    def _install_refresh_task():
        """Registers a Windows scheduled task at user level, with NO admin rights.
           UWAGA: tylko na stacji, ktora ma skoroszyt i poswiadczenia BC. Pozostale stanowiska
           read the finished file from a network share, so credentials are never duplicated."""
        ps1 = os.path.join(app_base_dir(), "Refresh-BinExport.ps1")
        if not os.path.exists(ps1):
            messagebox.showwarning("Auto-refresh",
                f"Script not found:\n{ps1}\n\nCopy Refresh-BinExport.ps1 next to the exe."); return
        hh = simpledialog.askstring("Auto-refresh", "Daily refresh time (HH:MM):",
                                    initialvalue=cfg.get("refresh_task_time") or "06:30", parent=root)
        if not hh: return
        hh = hh.strip()
        if not re.match(r"^\d{1,2}:\d{2}$", hh):
            messagebox.showerror("Auto-refresh", "Time format: HH:MM"); return
        cmd = ('schtasks /Create /F /TN "PickCore BinExport" /SC DAILY /ST ' + hh +
               ' /TR "powershell -WindowStyle Hidden -ExecutionPolicy Bypass -File \\"' + ps1 + '\\""')
        try:
            import subprocess
            r = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=25)
            out = (r.stdout or "") + (r.stderr or "")
            if r.returncode == 0:
                cfg["refresh_task_time"] = hh; save_cfg(cfg)
                logln(f"\u23F0 Auto-refresh task installed \u00B7 daily {hh}")
                messagebox.showinfo("Auto-refresh",
                    f"Task created: daily at {hh}.\n\n"
                    "Run it now:\n  schtasks /Run /TN \"PickCore BinExport\"\n\n"
                    "Runs while you are logged in - BC credentials live in your profile.\n"
                    "Do NOT install this on the other stations: they read the ready file from the share.")
            else:
                messagebox.showerror("Auto-refresh", f"schtasks returned an error:\n{out[:400]}")
        except Exception as e:
            messagebox.showerror("Auto-refresh", str(e))
    tk.Button(_depf,text="\u23F0  Install auto-refresh",command=_install_refresh_task,bg="#5a4a38",
              fg="#cfc3b0",relief="flat",font=("Bahnschrift",9)).pack(side="left",padx=6)
    _updf=tk.Frame(sf,bg="#1a1714"); _updf.grid(row=24,column=1,sticky="w",pady=(6,0))
    tk.Label(_updf,text="Update folder (share):",fg="#a99a84",bg="#1a1714",font=("Bahnschrift",8)).pack(side="left",padx=(0,4))
    v_upd=tk.StringVar(value=cfg.get("update_dir") or "")
    tk.Entry(_updf,textvariable=v_upd,width=32).pack(side="left")
    def _pick_upd():
        d=filedialog.askdirectory(title="Folder with the new version (network share)")
        if d: v_upd.set(d)
    tk.Button(_updf,text="\u2026",command=_pick_upd,bg="#5a4a38",fg="#cfc3b0",relief="flat").pack(side="left",padx=4)
    def _check_now():
        cfg["update_dir"]=v_upd.get().strip(); save_cfg(cfg); _update_check(announce=True)
    tk.Button(_updf,text="\u21BB Check",command=_check_now,bg="#5a4a38",fg="#cfc3b0",relief="flat",
              font=("Bahnschrift",8)).pack(side="left",padx=4)

    def _install_update():
        """IN-PLACE replacement (same path keeps the firewall rule), after the app has closed.
           The script waits for the process to exit, backs up, copies the new version and restarts."""
        src = (v_upd.get() or "").strip()
        newer, ver, info = check_update(src)
        if not newer:
            messagebox.showinfo("Update", f"No newer version available.\nYou have {APP_VERSION}. {info}"); return
        # A synced folder or share replicates files in arbitrary order, so VERSION.txt can
        # arrive before the exe. Without this check an update would start on an incomplete package.
        try:
            exes = [f for f in os.listdir(src) if f.lower().endswith(".exe")]
            internal = os.path.join(src, "_internal")
            probs = []
            if not exes: probs.append("missing .exe")
            if not os.path.isdir(internal): probs.append("missing _internal folder")
            else:
                if len(os.listdir(internal)) < 5: probs.append("_internal wyglada na niekompletny")
            if exes and os.path.getsize(os.path.join(src, exes[0])) < 500_000:
                probs.append("exe suspiciously small (cloud placeholder?)")
            # The installer runs robocopy /MIR on the exe folder, so a file MISSING from the
            # package is DELETED on the station, not skipped. Hence the check
            # obecnosc I rozmiar: synced folder potrafi dostarczyc zaslepke 0-bajtowa.
            # These files are DELIBERATELY not added to /XF: they should travel with the update,
            # and the real protection belongs here, before copying starts.
            for _f, _min in (("Items.csv", 100_000), ("customers.csv", 1_000)):
                _p = os.path.join(src, _f)
                if not os.path.exists(_p):
                    probs.append(f"missing {_f}")
                elif os.path.getsize(_p) < _min:
                    probs.append(f"{_f} suspiciously small (cloud placeholder?)")
            if probs:
                messagebox.showwarning("Update",
                    "The update package looks incomplete:\n  \u2022 " + "\n  \u2022 ".join(probs) +
                    "\n\nIf the source is a synced folder, wait for sync to finish\n"
                    "or set the folder to 'Always keep on this device'.")
                logln(f"\u26A0 Update {ver} halted: {', '.join(probs)}")
                return
        except Exception as e:
            messagebox.showerror("Update", f"Cannot read the update folder:\n{e}"); return
        if pa["lines"] or station.get("cur"):
            if not messagebox.askyesno("Update",
                "A session is in progress (pick or put-away).\n"
                "The update will close the application.\n\nContinue anyway?"): return
        # The /XD below is NOT cosmetic. The installer runs robocopy /MIR,
        # so EVERY folder present in the package overwrites its counterpart, and a folder
        # absent from the package is DELETED. Logs and archives are OPERATOR data,
        # created locally on every station:
        #   - Logs\pickcore_events.jsonl feeds the Impact dashboard,
        #   - the put-away archive holds signed receipt sheets.
        # If anyone built a package after running the app from dist (a single smoke
        # test is enough), those folders would ship inside it and wipe the archives
        # on every station. *.log is not enough, because events live in .jsonl.
        dst = app_base_dir()
        bak = os.path.join(os.path.dirname(dst.rstrip("\\/")), f"PickCore_backup_{APP_VERSION}")
        exe = sys.executable if getattr(sys, "frozen", False) else ""
        ps = f'''$ErrorActionPreference="SilentlyContinue"
$pid_ = {os.getpid()}
Write-Host ""
Write-Host "  PickCore - aktualizacja {APP_VERSION} -> {ver}" -ForegroundColor Cyan
Write-Host "  ---------------------------------------------"
Write-Host "  [1/4] Waiting for the application to close (PID $pid_)..."
for ($i=0; $i -lt 60; $i++) {{ if (-not (Get-Process -Id $pid_ -EA SilentlyContinue)) {{ break }}; Start-Sleep -Milliseconds 500 }}
Start-Sleep -Seconds 1
Write-Host "  [2/4] Kopia zapasowa..."
if (Test-Path "{bak}") {{ Remove-Item "{bak}" -Recurse -Force }}
Copy-Item "{dst}" "{bak}" -Recurse -Force
Write-Host "  [3/4] Kopiowanie nowej wersji..."
robocopy "{src}" "{dst}" /MIR /XF config.json pickcore_profile.json bin_contents.csv *.log *.jsonl /XD Logi Archive_Picks Archive_PutAways /NFL /NDL /NJH /NJS
if ($LASTEXITCODE -ge 8) {{
  Write-Host "COPY FAILED - restoring the backup"
  robocopy "{bak}" "{dst}" /MIR /NFL /NDL /NJH /NJS
}}
Write-Host "  [4/4] Uruchamiam PickCore..." -ForegroundColor Green
Start-Process "{exe}"
Start-Sleep -Seconds 2
'''
        try:
            up = os.path.join(os.environ.get("TEMP") or dst, "pickcore_update.ps1")
            Path(up).write_text(ps, encoding="utf-8")
            import subprocess
            subprocess.Popen(["powershell", "-WindowStyle", "Normal", "-ExecutionPolicy", "Bypass", "-File", up])
            log_event("APP_UPDATE", frm=APP_VERSION, to=ver)
            logln(f"\u2191 Update {APP_VERSION} \u2192 {ver}: closing application\u2026")
            # Progress window: the operator sees something happening instead of a vanishing app.
            win = tk.Toplevel(root); win.title("PickCore - update")
            win.configure(bg=UI["panel"]); win.attributes("-topmost", True)
            win.geometry("430x170"); win.resizable(False, False)
            win.protocol("WM_DELETE_WINDOW", lambda: None)
            tk.Label(win, text=f"Updating {APP_VERSION} \u2192 {ver}", bg=UI["panel"], fg=UI["accent"],
                     font=("Bahnschrift",13,"bold")).pack(pady=(16,4))
            stp = tk.Label(win, text="Closing application\u2026", bg=UI["panel"], fg=UI["text"],
                           font=("Bahnschrift",10)); stp.pack()
            pb = ttk.Progressbar(win, mode="indeterminate", length=360); pb.pack(pady=12); pb.start(12)
            tk.Label(win, text="The application will restart automatically.\nDo not close the PowerShell window.",
                     bg=UI["panel"], fg=UI["muted"], font=("Bahnschrift",8)).pack()
            steps = ["Closing application\u2026", "Backing up current version\u2026",
                     "Copying new files\u2026", "Restarting\u2026"]
            def _tick(k=0):
                if k < len(steps):
                    stp.config(text=steps[k]); win.after(1200, lambda: _tick(k+1))
            _tick()
            root.after(4200, root.destroy)
        except Exception as e:
            messagebox.showerror("Update", str(e))
    tk.Button(_updf,text="\u2191  Install update",command=_install_update,bg="#B45309",fg="white",
              relief="flat",font=("Bahnschrift",9,"bold")).pack(side="left",padx=8)
    tk.Label(sf,text="Update: in-place replacement + backup; local config and station profile are preserved",
             fg="#7a6c58",bg="#1a1714",font=("Bahnschrift",8)).grid(row=25,column=1,sticky="w")
    tk.Label(_depf,text="(writes pickcore_profile.json next to the exe - template for further stations)",
             fg="#7a6c58",bg="#1a1714",font=("Bahnschrift",8)).pack(side="left",padx=8)
    _bcf3=tk.Frame(sf,bg="#1a1714"); _bcf3.grid(row=19,column=1,sticky="w",pady=(4,0))
    tk.Label(_bcf3,text="Web service:",fg="#a99a84",bg="#1a1714",font=("Bahnschrift",8)).pack(side="left",padx=(0,2))
    v_bcws=tk.StringVar(value=cfg.get("bc_ws") or "Bin_Contents_List")
    tk.Entry(_bcf3,textvariable=v_bcws,width=22).pack(side="left",padx=(0,10))
    def bc_test():
        """Diagnostics: token, one row, field names. No guessing what the ERP actually exposes."""
        probe = dict(cfg)
        probe.update(bc_enabled=True, bc_base_url=v_bcurl.get().strip().rstrip("/"),
                     bc_company=v_bccmp.get().strip(), bc_tenant=v_bctn.get().strip(),
                     bc_client_id=v_bccid.get().strip(), bc_ws=v_bcws.get().strip())
        def run():
            import urllib.request, urllib.parse
            tok = ""
            try:
                if not _bc_secret():
                    bc_q.put(("err","no secret: set PICKCORE_BC_SECRET or use Credential Manager")); return
                tok = _bc_token(probe)
                url = (f'{probe["bc_base_url"]}/ODataV4/Company(\'{urllib.parse.quote(probe["bc_company"])}\')'
                       f'/{probe["bc_ws"]}?$top=1')
                req = urllib.request.Request(url, headers={"Authorization": f"Bearer {tok}",
                                                           "Accept": "application/json"})
                with urllib.request.urlopen(req, timeout=12) as r:
                    j = json.loads(r.read().decode("utf-8"))
                vals = j.get("value") or []
                if not vals:
                    bc_q.put(("test","connection OK, but the service returned 0 rows")); return
                keys = sorted(vals[0].keys())
                need = [k for k in ("Item_No","Bin_Code","Quantity") if k not in keys]
                msg = f"connection OK. Fields: {', '.join(keys[:12])}"
                if need: msg += f"  |  MISSING expected: {', '.join(need)} - mapping required"
                bc_q.put(("test", msg))
            except Exception as e:
                msg = str(e)[:140]
                if "404" in msg:
                    # 404 means authentication succeeded but the path is wrong. Query the service catalogue and show the real names.
                    try:
                        import urllib.request as _u
                        root_url = f'{probe["bc_base_url"]}/ODataV4/'
                        rq = _u.Request(root_url, headers={"Authorization": f"Bearer {tok}",
                                                           "Accept": "application/json"})
                        with _u.urlopen(rq, timeout=12) as r:
                            doc = json.loads(r.read().decode("utf-8"))
                        names = [v.get("name") or v.get("url") for v in doc.get("value", [])]
                        hit = [n for n in names if n and ("bin" in n.lower() or "content" in n.lower())]
                        if hit:
                            bc_q.put(("test", "404: wrong service name. Published matches: " + ", ".join(hit[:8])))
                        elif names:
                            bc_q.put(("test", f"404: no service with 'bin' in the name. Published ({len(names)}): "
                                              + ", ".join([n for n in names if n][:12])))
                        else:
                            bc_q.put(("err", "404 and the service catalog is empty - check the company name and whether page "
                                             "7379 Bin Contents is published as a web service"))
                    except Exception as e2:
                        bc_q.put(("err", f"404 on the path; service catalog unavailable too ({str(e2)[:70]}). "
                                         f"Check Base URL (without /ODataV4/...) and company name without %20"))
                else:
                    bc_q.put(("err", f"connection test: {msg}"))
        threading.Thread(target=run, daemon=True).start()
        logln("\u2601 BC: testing connection\u2026")
    tk.Button(_bcf3,text="\U0001F50C  Test connection",command=bc_test,bg="#5a4a38",fg="#cfc3b0",
              relief="flat",font=("Bahnschrift",9)).pack(side="left")
    tk.Label(sf,text="Base URL = .../v2.0/<tenant>/<environment>  (no /ODataV4/...) \u00B7 Company = plain name, no URL encoding",
             fg="#7a6c58",bg="#1a1714",font=("Bahnschrift",8)).grid(row=20,column=1,sticky="w")
    tk.Label(sf,text="Secret is NEVER stored here - set env var PICKCORE_BC_SECRET or Windows Credential Manager",
             fg="#7a6c58",bg="#1a1714",font=("Bahnschrift",8)).grid(row=21,column=1,sticky="w")
    v_cld=tk.BooleanVar(value=bool(cfg.get("cloud_publish")))
    tk.Checkbutton(sf,text="Publish the live board to the synced folder (pickcore_live.html)",variable=v_cld,fg="#f2ece1",bg="#1a1714",
                   selectcolor="#2a241e",activebackground="#1a1714").grid(row=11,column=1,sticky="w",pady=2)
    v_http=tk.BooleanVar(value=bool(cfg.get("http_serve")))
    _htfr=tk.Frame(sf,bg="#1a1714")
    tk.Checkbutton(sf,text="Serve live view over LAN (TC22 preview in Chrome)",variable=v_http,fg="#f2ece1",bg="#1a1714",
                   selectcolor="#2a241e",activebackground="#1a1714").grid(row=13,column=1,sticky="w",pady=2)
    _htfr.grid(row=14,column=1,sticky="w")
    tk.Label(_htfr,text="HTTP port:",fg="#a99a84",bg="#1a1714").pack(side="left")
    v_httpp=tk.StringVar(value=str(cfg.get("http_port") or 8080))
    tk.Entry(_htfr,textvariable=v_httpp,width=6,justify="center").pack(side="left",padx=(4,10))
    tk.Label(_htfr,text="(restart applies; open http://IP_PC:port on TC22)",fg="#7a6c58",bg="#1a1714",
             font=("Bahnschrift",8)).pack(side="left")
    _cldfr=tk.Frame(sf,bg="#1a1714"); _cldfr.grid(row=12,column=1,sticky="w")
    v_cldd=tk.StringVar(value=cfg.get("cloud_dir") or "")
    tk.Entry(_cldfr,textvariable=v_cldd,width=44).pack(side="left")
    def _pick_cloud_dir():
        d=filedialog.askdirectory(title="Choose the synced folder")
        if d: v_cldd.set(d)
    tk.Button(_cldfr,text="\u2026",command=_pick_cloud_dir,bg="#5a4a38",fg="#cfc3b0",relief="flat").pack(side="left",padx=4)

    # --- WATCHER: folder scan every 2s, tolerant of temp-plus-rename writes ---
    def _prime_seen(d):
        """Marks existing files as seen, so only NEW ones are processed after startup."""
        try:
            for fn in os.listdir(d):
                if fn.lower().endswith(".pdf"): seen.add(os.path.join(d, fn))
        except Exception: pass

    def start_watcher():
        d = cfg.get("watch_dir")
        if cfg.get("watcher_on") and d and os.path.isdir(d):
            _prime_seen(d)
            wstat.config(text=f"Watcher: ON  (scanning {d})", fg="#27c66d")
            hdr_led.config(fg=UI["ok"])
        elif cfg.get("watcher_on") and d:
            wstat.config(text=f"Watcher: ON but folder not found: {d}", fg="#ef4655")
            hdr_led.config(fg=UI["err"])
        else:
            wstat.config(text="Watcher: OFF", fg="#a99a84")
            hdr_led.config(fg=UI["faint"])

    def save_settings():
        cfg.update(watch_dir=v_watch.get().strip(),
                   archive_dir_pick=v_arch_pick.get().strip(), archive_dir_pa=v_arch_pa.get().strip(),
                   name_filter=v_filter.get().strip(), pa_filter=v_pafilter.get().strip(),
                   printer=("" if v_printer.get()=="(default)" else v_printer.get()),
                   auto_print=v_auto.get(), watcher_on=v_won.get(),
                   scanner_listen=v_scn.get(),
                   scanner_port=(int(v_scnp.get()) if (v_scnp.get() or "").strip().isdigit() else 5577),
                   scanner_allow=v_scna.get().strip(),
                   cloud_publish=v_cld.get(), cloud_dir=v_cldd.get().strip(),
                   bc_enabled=v_bce.get(), bc_base_url=v_bcurl.get().strip(),
                   bc_company=v_bccmp.get().strip(), bc_tenant=v_bctn.get().strip(),
                   bc_client_id=v_bccid.get().strip(), bc_ws=v_bcws.get().strip(),
                   update_dir=v_upd.get().strip(),
                   items_csv=v_items.get().strip(), sound_dir=v_snd.get().strip(),
                   customers_csv=v_cust.get().strip(),
                   bc_paste_order=(v_ord.get().strip().lower() or "sku,desc,loc,nloc,bin,nbin,qty"),
                   default_location=(v_dloc.get().strip().upper() or "MAIN"),
                   location_codes=v_locs.get().strip().upper(),
                   bin_export_path=v_bex.get().strip(), bin_export_auto=v_bea.get(),
                   bin_export_max_age=(int(v_bem.get()) if (v_bem.get() or "").strip().isdigit() else 60),
                   http_serve=v_http.get(),
                   http_port=(int(v_httpp.get()) if (v_httpp.get() or "").strip().isdigit() else 8080))
        save_cfg(cfg); start_watcher(); _hdr_prn_refresh()
        messagebox.showinfo("Settings","Saved. Watcher "+("ON" if cfg["watcher_on"] else "OFF"))
    root.after(600, lambda: _show_page("Station"))   # domyslna zakladka po zbudowaniu kontrolek
    tk.Button(t3,text="💾  Save settings",command=save_settings,bg="#2fb5a8",fg="white",
              relief="flat",font=("Bahnschrift",11,"bold"),padx=20,pady=8).pack(pady=10)

    def scan_folder():
        d = cfg.get("watch_dir")
        if not (cfg.get("watcher_on") and d and os.path.isdir(d)): return
        try:
            files = os.listdir(d)
        except Exception:
            return
        for fn in files:
            if not fn.lower().endswith((".pdf", ".xlsx")): continue
            full = os.path.join(d, fn)
            if full in seen: continue
            seen.add(full)
            # process_pdf detects the type from the CONTENT; the file name is irrelevant
            def work(path=full):
                if not wait_stable(path): return
                process_pdf(path, cfg, customers, log_q.put, on_pick=lambda pd: new_picks.put(pd), on_putaway=lambda h,r: pa_q.put((h,r)))
            threading.Thread(target=work, daemon=True).start()

    def drain_logs():
        """Logs entries from worker threads on the MAIN thread (Tkinter is not thread safe)."""
        try:
            while True: logln(log_q.get_nowait())
        except _queue.Empty: pass
        root.after(200, drain_logs)

    def poll():
        now = time.time()
        if now - scan_state["last"] >= 2.0:
            scan_state["last"] = now
            try: scan_folder()
            except Exception: pass
        root.after(500, poll)

    refresh_customers(); refresh_station(); start_watcher()
    nb.insert(1, t_in, text="Put-away", group="INBOUND", fkey="F2", icon="📥")
    def _nav_serial():
        """Serial jako pozycja pierwszej klasy w sidebarze: przelacza na stacje i odpala
           the existing mode (serials within a pick, or a free session) with no widget re-parenting."""
        nb.select(t1)
        if serial_mode["on"] or session["on"]:
            scan_entry.focus_set(); return
        if station["cur"]: toggle_serial()
        else: toggle_session()
    nb.add_action("Serial session", _nav_serial, group="OUTBOUND", fkey="F3", icon="⚡")
    nb.bind_fkeys(root)
    nb.select(t1)
    root.after(500, poll); root.after(300, drain_picks); root.after(200, drain_logs); root.after(350, drain_pa)

    # ---------- TC22 SCANNER LISTENER (DataWedge IP Output -> TCP) \u2014 tryb demo ----------
    scan_net_q = _queue.Queue()
    scanner_state = {"last": 0.0, "peer": "", "on": False, "err": ""}
    def _scanner_srv(port, allow):
        import socket
        try:
            srv = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            srv.bind(("0.0.0.0", port)); srv.listen(8); scanner_state["on"] = True
        except Exception as e:
            scanner_state["err"] = f"port {port}: {e}"; return
        def _serve(conn, addr):
            """One connection, one thread. Without it a SECOND scanner would wait for the first to disconnect
               (DataWedge holds the TCP session open), which is the normal case with several stations."""
            try:
                scanner_state["peer"] = addr[0]
                scanner_state["peers"] = sorted(set(scanner_state.get("peers", []) + [addr[0]]))
                with conn:
                    conn.settimeout(600); buf = b""
                    while True:
                        chunk = conn.recv(1024)
                        if not chunk: break
                        buf += chunk
                        while b"\n" in buf:
                            raw, buf = buf.split(b"\n", 1)
                            code = raw.decode("utf-8", "ignore").strip()
                            if code:
                                scan_net_q.put(code); scanner_state["last"] = time.time()
                    tail = buf.decode("utf-8", "ignore").strip()   # skan bez terminatora (DataWedge bez Send ENTER)
                    if tail:
                        scan_net_q.put(tail); scanner_state["last"] = time.time()
            except Exception:
                pass
        while True:
            try:
                conn, addr = srv.accept()
                if allow and addr[0] not in [a.strip() for a in allow.split(",") if a.strip()]:
                    conn.close(); continue
                threading.Thread(target=_serve, args=(conn, addr), daemon=True).start()
            except Exception:
                time.sleep(0.5)
    def drain_scanner():
        try:
            while True:
                code = scan_net_q.get_nowait()
                logln(f"\U0001F4E1 TC22 scan: {code}")
                log_event("SCAN_REMOTE", code=code, src=scanner_state.get("peer",""))
                # TTL raised to 30 min: with the handheld screen off the browser stops polling,
                # so the heartbeat goes quiet and a short TTL used to divert scans back to the desktop station.
                fresh = (time.time() - web_focus["ts"]) < 1800
                hv = web_focus["view"] if fresh else ""
                b, okb = format_bin_input(code)
                # Trwajaca sekwencja relokacji ma PIERWSZENSTWO nad wygaslym fokusem:
                # with an item or a source bin open, the next scan belongs to that sequence.
                if rel.get("pend", {}).get("sku"):
                    rel_feed(code); render_web_state(); continue
                if hv == "inbound" and pa["lines"] and okb:
                    # operator stoi w Inbound na TC22 i zeskanowal LOKACJE -> wpisz w wybrana linie
                    i = web_focus["i"]
                    if not (0 <= i < len(pa["lines"])) or pa["lines"][i].get("bin"):
                        i = next((k for k, l in enumerate(pa["lines"]) if not l.get("bin")), -1)
                    if i >= 0:
                        _pa_set_bin(i, b, "tc22")
                        logln(f"\U0001F4E1 TC22 bin: line {i+1} ({pa['lines'][i].get('sku','')}) \u2192 {b}")
                        nxt = next((k for k, l in enumerate(pa["lines"]) if not l.get("bin")), -1)
                        web_focus["i"] = nxt; web_focus["ts"] = time.time()
                        _beep("line" if nxt < 0 else "ok")
                    else:
                        _beep("warn")
                elif hv == "reloc" or nb.select() == str(trel):
                    rel_feed(code)                     # Relocations: z konsoli albo z widoku na PC
                else:
                    scan_var.set(code); on_scan()
                render_web_state()                     # natychmiastowe odswiezenie konsoli na TC22
        except _queue.Empty:
            pass
        if cfg.get("scanner_listen"):
            fresh = (time.time() - scanner_state["last"]) < 12
            hdr_scn.config(text="\U0001F4E1 TC22",
                           fg=(UI["ok"] if fresh else (UI["err"] if scanner_state["err"] else UI["faint"])))
        root.after(250, drain_scanner)
    if cfg.get("scanner_listen"):
        hdr_scn.pack(side="right", padx=6)
        threading.Thread(target=_scanner_srv, args=(int(cfg.get("scanner_port") or 5577),
                         (cfg.get("scanner_allow") or "").strip()), daemon=True).start()
    root.after(400, drain_scanner)

    # ---------- CLOUD LIVE BOARD (synced folder; station state only, no sensitive data) ----------
    def _live_html(refresh=5):
        rows = "".join(f"<tr><td class='t'>{html.escape(t)}</td><td>{html.escape(m)}</td></tr>"
                       for t, m in reversed(live_ring))
        done = sum(1 for l in pa["lines"] if l.get("bin"))
        hero = html.escape(live_ring[-1][1]) if live_ring else "Waiting for first event\u2026"
        return ("<!doctype html><html><head><meta charset='utf-8'>"
            "<meta name='viewport' content='width=device-width, initial-scale=1'>"
            f"<meta http-equiv='refresh' content='{refresh}'>"
            "<title>PickCore live</title><style>body{background:#1a1714;color:#f2ece1;font-family:Bahnschrift,Arial;margin:24px}"
            "h1{color:#2fb5a8;font-size:20px;margin:0}.sub{color:#a99a84;font-size:12px;margin-bottom:14px}"
            ".card{background:#241f1a;border:1px solid #45392c;border-radius:8px;padding:12px 16px;margin:8px 0}"
            "table{width:100%;border-collapse:collapse;font-family:Consolas,monospace;font-size:12px}"
            "td{padding:3px 8px;border-bottom:1px solid #2e2720}.t{color:#7a6c58;width:70px}</style></head><body>"
            "<h1>PickCore \u2014 live station board</h1>"
            f"<div class='sub'>Operator {html.escape(get_picker())} \u00B7 PickCore {APP_VERSION} \u00B7 "
            f"updated {datetime.now().strftime('%H:%M:%S')} \u00B7 auto-refresh {refresh} s</div>"
            f"<div class='card' style='font-size:17px;font-weight:600'>{hero}</div>"
            f"<div class='card'><b>Pick station:</b> {html.escape(st_title.cget('text'))} \u2014 {html.escape(st_badge.cget('text'))}</div>"
            f"<div class='card'><b>Inbound:</b> PO {html.escape(pa.get('po') or '\u2014')} \u00B7 {done}/{len(pa['lines'])} bins set "
            f"\u00B7 {len(pa_pending)} pending</div>"
            f"<div class='card'><b>Last events</b><table>{rows}</table></div></body></html>")
    def live_cache_tick():
        """Rendered on the main Tk thread, where reading widgets is safe; consumers read the finished string."""
        try:
            if cloud_state.get("dirty"):
                cloud_state["html"] = _live_html(refresh=2)
                cloud_state["dirty"] = False
            render_web_state()
        except Exception:
            pass
        root.after(1000, live_cache_tick)
    root.after(900, live_cache_tick)
    def cloud_tick():
        try:
            d = (cfg.get("cloud_dir") or "").strip()
            h = cloud_state.get("html") or ""
            if cfg.get("cloud_publish") and d and os.path.isdir(d) and h and h != cloud_state.get("written"):
                tmp = os.path.join(d, "._pickcore_live.tmp"); out = os.path.join(d, "pickcore_live.html")
                with open(tmp, "w", encoding="utf-8") as f: f.write(h)
                os.replace(tmp, out); cloud_state["written"] = h
        except Exception:
            pass
        root.after(4000, cloud_tick)
    root.after(2500, cloud_tick)

    # ================= HANDHELD WEB CONSOLE (LAN only) =================
    # Kontrakt (przenosny 1:1 do portu portfolio):
    #   - the HTTP thread NEVER touches Tk widgets. It reads only the JSON snapshot (web_state)
    #     rendered on the main thread, and pushes commands onto web_cmd_q.
    #   - glowny watek: render_web_state() + drain_web() -> pelna serializacja, zero wyscigow.
    #   - Powierzchnia API: GET / (konsola), GET /api/state, POST /api/cmd.
    http_state = {"on": False, "err": ""}
    web_state = {"json": "{}"}
    web_focus = {"view": "", "i": -1, "ts": 0.0}   # gdzie stoi operator na TC22 (routing skanow)
    pa_last = {"id": 0, "i": -1, "sku": "", "bin": "", "prev": "", "src": ""}   # ostatni wpis binu -> undo
    pa_done = {"doc": "", "po": "", "lines": 0, "missing": 0, "ts": "", "file": ""}  # ostatni POTWIERDZONY
    web_cmd_q = _queue.Queue()

    def render_web_state():
        """Station state snapshot for the handheld console. Main Tk thread only."""
        try:
            lines = [{"i": i, "sku": l.get("sku",""), "desc": (l.get("desc") or "")[:34],
                      "qty": l.get("qty",""), "bin": l.get("bin",""),
                      "sug": [{"b": c["bin"],
                               "q": (int(c["qty"]) if c.get("src") == "bc" and c.get("qty") is not None
                                     and float(c["qty"]).is_integer() else c.get("qty")),
                               "s": c.get("src","local")}
                              for c in loc_suggest(l.get("sku",""))[:2]]}
                     for i, l in enumerate(pa["lines"])]
            rel_rows = [{"i": rel["lines"].index(l), "sku": l["sku"], "frm": l["frm"], "to": l["to"],
                         "qty": l["qty"], "loc": l.get("loc","MAIN")}
                        for l in rel["lines"][-8:]]
            cur = station.get("cur")
            plines, pdone, pneed = [], 0, 0
            if cur and not session["on"]:
                # same order as the printout: ascending by location, not the order found in the PDF
                _order = sorted(range(len(cur["items"])),
                                key=lambda k: bin_sort_key(cur["items"][k].get("bin","")))
                for idx in _order:
                    it = cur["items"][idx]
                    sc, nd = int(it.get("scanned", 0)), int(it.get("need", 0))
                    pdone += min(sc, nd); pneed += nd
                    plines.append({"i": idx, "sku": it.get("sku",""), "bin": it.get("bin","") or "\u2014",
                                   "sc": sc, "nd": nd, "ser": bool(it.get("serial")),
                                   "st": ("done" if sc >= nd else ("part" if sc > 0 else "todo"))})
            snap = {
                "ts": datetime.now().strftime("%H:%M:%S"),
                "op": get_picker(),
                "ver": APP_VERSION,
                "pick": {"title": st_title.cget("text"), "sub": st_sub.cget("text"),
                         "badge": st_badge.cget("text"), "fb": fb.cget("text"),
                         "fbk": ("ok" if fb.cget("bg") == UI["ok_bg"] else
                                 "err" if fb.cget("bg") == UI["err_bg"] else
                                 "warn" if fb.cget("bg") == UI["warn_bg"] else "info"),
                         "lines": plines, "done": pdone, "need": pneed,
                         "queue": len(station.get("queue") or []),
                         "serial": bool(serial_mode["on"]), "bulk": bool(bulk["on"])},
                "alerts": dict(web_alerts),
                "focus": {"view": web_focus["view"],
                          "i": (web_focus["i"] if (time.time() - web_focus["ts"]) < 1800 else -1)},
                "last": {"id": pa_last["id"], "i": pa_last["i"], "sku": pa_last["sku"],
                         "bin": pa_last["bin"], "prev": pa_last["prev"], "src": pa_last["src"]},
                "done": dict(pa_done),
                "inb": {"po": pa.get("po") or "", "doc": pa.get("doc") or "",
                        "lines": lines, "pending": len(pa_pending),
                        "done": sum(1 for l in pa["lines"] if l.get("bin"))},
                "rel": {"rows": rel_rows, "pend": rel.get("pend", {}), "total": len(rel["lines"])},
                "log": [{"t": t, "m": m} for t, m in list(reversed(live_ring))[:8]],
                "ser": ({"on": True,
                         "radios": [{"i": i, "sku": r.get("sku",""), "bin": r.get("bin",""),
                                     "qty": int(r.get("qty",0)), "pack": int(r.get("pack",1)),
                                     "n": len(r.get("serials", [])),
                                     "need": int(r.get("qty",0)) * int(r.get("pack",1)),
                                     "act": (i == serial_mode.get("active"))}
                                    for i, r in enumerate(serial_mode.get("radios") or [])],
                         "packs": dict(PACK_MODES)}
                        if serial_mode.get("on") else {"on": False, "radios": [], "packs": dict(PACK_MODES)}),
                "scn": {"on": bool(scanner_state.get("on")),
                        "err": (scanner_state.get("err") or "")[:60],
                        "peers": scanner_state.get("peers") or [],
                        "ago": (int(time.time() - scanner_state["last"]) if scanner_state.get("last") else -1),
                        "port": int(cfg.get("scanner_port") or 5577),
                        "ip": local_ip()},
                "snd": {"id": snd_state["id"], "k": snd_state["kind"],
                        "files": [k for k in SOUND_KINDS if find_sound(cfg, k, SOUND_EXT_WEB)]},
                "tones": BEEP_SEQS,
            }
            web_state["json"] = json.dumps(snap, ensure_ascii=False)
        except Exception as e:
            # The MOST DANGEROUS silent exception in this file: the handheld console freezes
            # na ostatnim dobrym snapshocie, HTTP odpowiada 200, testy API przechodza,
            # and the operator at the rack sees no updates. This class of failure cost a week.
            web_state["err"] = str(e)[:120]
            logln(f"⚠ Console snapshot FAILED: {str(e)[:90]}")

    def drain_web():
        """Executes commands from the handheld console on the main thread."""
        try:
            while True:
                c = web_cmd_q.get_nowait()
                op = c.get("op")
                if op == "bin":
                    i, v = int(c.get("i", -1)), (c.get("v") or "").strip()
                    if 0 <= i < len(pa["lines"]):
                        b, okb = format_bin_input(v)
                        if okb:
                            _pa_set_bin(i, b, "handheld")
                            logln(f"\U0001F4F1 Handheld: line {i+1} \u2192 {b}")
                        else:
                            logln(f"\u2715 Handheld: invalid bin '{v}' (line {i+1})")
                            _beep("err")
                elif op == "bin_clear":
                    i = int(c.get("i", -1))
                    if 0 <= i < len(pa["lines"]):
                        old = pa["lines"][i].get("bin","")
                        _pa_set_bin(i, "", "clear")
                        web_focus.update(view="inbound", i=i, ts=time.time())
                        logln(f"\u21A9 Handheld: line {i+1} cleared (was {old or '\u2014'}) \u2014 rescan")
                        _beep("warn")
                elif op == "bin_undo":
                    i = pa_last.get("i", -1)
                    if 0 <= i < len(pa["lines"]) and pa_last.get("id"):
                        was = pa_last.get("bin",""); prev = pa_last.get("prev","")
                        _pa_set_bin(i, prev, "undo")
                        web_focus.update(view="inbound", i=i, ts=time.time())
                        logln(f"\u21A9 Undo: line {i+1} {was or '\u2014'} \u2192 {prev or '(empty)'}")
                        _beep("warn")
                elif op == "confirm":
                    if pa["lines"]: confirm_putaway(False)
                elif op == "next_pa":
                    if pa_pending: load_next_pa(auto=True)
                elif op == "scan":
                    code = (c.get("v") or "").strip()
                    if code: rel_feed(code)
                elif op == "view":
                    web_focus["view"] = (c.get("v") or "").strip()
                    web_focus["ts"] = time.time()
                    if c.get("i") is not None:
                        try: web_focus["i"] = int(c.get("i"))
                        except (TypeError, ValueError):
                            logln(f"⚠ Console: bad line index in view command: {c.get('i')!r}")
                elif op == "target":
                    try:
                        web_focus["i"] = int(c.get("i")); web_focus["view"] = "inbound"
                        web_focus["ts"] = time.time()
                    except (TypeError, ValueError):
                        logln(f"⚠ Console: bad line index in target command: {c.get('i')!r}")
                elif op == "rel_qty":
                    try:
                        i, n = int(c.get("i", -1)), int(c.get("n", 0))
                    except Exception:
                        i, n = -1, 0
                    if 0 <= i < len(rel["lines"]) and n > 0:
                        rel["lines"][i]["qty"] = n; rel_save(); rel_rebuild()
                        logln(f"\U0001F4F1 Handheld: reloc line {i+1} qty = {n}")
                        _beep("ok")
                elif op == "rel_undo":
                    rel_remove_last()
                elif op == "pick_undo":
                    remove_last_pick()
                elif op == "pick_qty":
                    # BULK z reki: operator wskazuje linie i podaje ilosc (pudelko/paleta).
                    # The same posting path as a desktop bulk entry: counters, telemetry, sounds.
                    try:
                        i, n = int(c.get("i", -1)), int(c.get("n", 0))
                    except Exception:
                        i, n = -1, 0
                    cur = station.get("cur")
                    if cur and 0 <= i < len(cur["items"]) and n > 0:
                        line = cur["items"][i]
                        remaining = int(line["need"]) - int(line["scanned"])
                        if n > remaining:
                            logln(f"\u2715 Handheld BULK blocked: {line['sku']} +{n}, only {remaining} outstanding")
                            log_event("BULK_BLOCKED", sku=line["sku"], scanned_qty=n, remaining=remaining,
                                      pick=cur.get("title",""), doc=cur.get("hdr",""), src="handheld")
                            feedback(f"BLOCKED: +{n} > remaining {remaining} ({line['sku']})", "err"); _beep("err")
                        else:
                            line["scanned"] += n
                            try: learn_bulk(line["sku"], n)
                            except Exception as e:
                                logln(f"⚠ learn_bulk({line['sku']}): {str(e)[:60]}")
                            _now = time.time()
                            if pick_t["t0"] is None: pick_t["t0"] = _now
                            pick_t["scans"].append((round(_now - pick_t["t0"], 1), line["sku"], line.get("bin",""), n))
                            log_event("BULK_HANDHELD", sku=line["sku"], qty=n,
                                      pick=cur.get("title",""), doc=cur.get("hdr",""))
                            logln(f"\U0001F4F1 Handheld BULK: {line['sku']} +{n} ({line['scanned']}/{line['need']})")
                            feedback(f"BULK +{n}   \u00B7   {line['sku']}   ({line['scanned']}/{line['need']})", "ok")
                            _beep("line" if line["scanned"] >= line["need"] else "ok")
                            refresh_station(); check_complete()
                elif op == "ser_pick":
                    # selects the unit that subsequent serial scans belong to
                    try: i = int(c.get("i", -1))
                    except Exception: i = -1
                    if serial_mode.get("on") and 0 <= i < len(serial_mode.get("radios") or []):
                        serial_mode["active"] = i
                        refresh_station()
                        r = serial_mode["radios"][i]
                        logln(f"\U0001F4F1 Handheld: serial target \u2192 {r.get('sku','')} ({r.get('bin','')})")
                        _beep("ok")
                elif op == "ser_pack":
                    # pack size for the SELECTED unit (single / dual / quad / 6-pack)
                    try:
                        i, n = int(c.get("i", -1)), int(c.get("n", 1))
                    except Exception:
                        i, n = -1, 1
                    if serial_mode.get("on") and 0 <= i < len(serial_mode.get("radios") or []) and n >= 1:
                        serial_mode["radios"][i]["pack"] = n
                        serial_mode["active"] = i
                        refresh_station()
                        logln(f"\U0001F4F1 Handheld: pack size {n} \u2192 {serial_mode['radios'][i].get('sku','')}")
                        _beep("ok")
                elif op == "pick_serial":
                    nb.select(t1)
                    if station["cur"]: toggle_serial()
                elif op == "pick_scan":
                    code = (c.get("v") or "").strip()
                    if code:
                        scan_var.set(code); on_scan()
                cloud_state["dirty"] = True
                render_web_state()      # natychmiastowa swiezosc po akcji z handhelda (bez czekania na tick)
        except _queue.Empty:
            pass
        root.after(300, drain_web)
    root.after(700, drain_web)
    root.after(800, drain_bc)

    def _items_autoload():
        """Item catalogue at startup. Source order:
           1) plik wskazany w Settings, 2) katalog DOLACZONY DO PACZKI (obok exe).
           Dzieki (2) nowa stacja ma komplet opisow od pierwszego uruchomienia - zero konfiguracji."""
        try:
            f = (cfg.get("items_csv") or "").strip()
            if not (f and os.path.exists(f)):
                for cand in ("Items.csv", "items.csv", "ItemCatalogue.csv", "item_descriptions.csv"):
                    bundled = os.path.join(app_base_dir(), cand)
                    if os.path.exists(bundled):
                        f = bundled
                        if not cfg.get("items_csv"):
                            cfg["items_csv"] = f; save_cfg(cfg)
                            logln(f"\U0001F4D6 Item catalogue found in package: {cand}")
                        break
            if not (f and os.path.exists(f)): return
            reg_m = SKU_DESC_PATH.stat().st_mtime if SKU_DESC_PATH.exists() else 0
            if os.path.getmtime(f) <= reg_m: 
                logln(f"\U0001F4D6 Item catalogue: {len(load_sku_desc())} descriptions in registry")
                return
            n, info = import_item_descriptions(f)
            logln(f"\U0001F4D6 Item catalogue refreshed: {info}")
        except Exception: pass
    root.after(2200, _items_autoload)

    def _customers_autoload():
        """Customer file: the path from Settings first, then the copy SHIPPED IN THE PACKAGE next to the exe."""
        try:
            f = (cfg.get("customers_csv") or "").strip()
            if not (f and os.path.exists(f)):
                base = app_base_dir()
                cand = [os.path.join(base, n) for n in ("customers.csv", "Customers.csv")]
                cand += sorted(glob.glob(os.path.join(base, "Default*.csv")))
                f = next((c for c in cand if os.path.exists(c)), "")
                if f and not cfg.get("customers_csv"):
                    cfg["customers_csv"] = f; save_cfg(cfg)
            if not f: return
            n, info = load_customer_db(f)
            logln(f"\U0001F465 Customer database: {info}" if n else f"\u26A0 Customer database: {info}")
        except Exception as e:
            logln(f"\u26A0 Customer database: {str(e)[:110]}")
    root.after(2400, _customers_autoload)

    # ---------- AUTO-UPDATE: detects a new version on the share, never swapping mid-shift ----------
    def _update_check(announce=False):
        def run():
            newer, ver, info = check_update(cfg.get("update_dir"))
            upd_q.put((newer, ver, info, announce))
        threading.Thread(target=run, daemon=True).start()
    def drain_upd():
        try:
            while True:
                newer, ver, info, announce = upd_q.get_nowait()
                if newer:
                    upd_state.update(ver=ver, info=info)
                    hdr_upd.config(text=f"\u2191 {ver}")
                    if not hdr_upd.winfo_ismapped(): hdr_upd.pack(side="right", padx=6)
                    logln(f"\u2191 Update {ver} available \u2014 Settings \u2192 Install update")
                elif announce:
                    messagebox.showinfo("Update", f"You are up to date ({APP_VERSION}).\n{info}")
        except _queue.Empty:
            pass
        root.after(1000, drain_upd)
    upd_q = _queue.Queue()
    root.after(1500, drain_upd)
    if (cfg.get("update_dir") or "").strip():
        root.after(4000, _update_check)
        def _upd_loop():
            _update_check(); root.after(3600000, _upd_loop)   # co godzine, cicho
        root.after(3600000, _upd_loop)
    hdr_upd.bind("<Button-1>", lambda e: (nb.select(t3), None)[-1])

    HANDHELD_HTML = r"""<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>PickCore handheld</title><style>
*{box-sizing:border-box}body{margin:0;background:#1a1714;color:#f2ece1;font:14px/1.4 "Bahnschrift",Roboto,Arial}
header{background:#241f1a;padding:8px 12px;display:flex;align-items:center;gap:8px;position:sticky;top:0;border-bottom:1px solid #45392c}
header b{color:#2fb5a8}#ts{color:#7a6c58;font:11px monospace}
#snd{margin-left:auto;background:#383026;border:0;color:#a99a84;border-radius:6px;padding:6px 10px;font-size:15px;margin-right:8px}
#snd.on{background:#12301f;color:#27c66d}
nav{display:flex;background:#241f1a;border-bottom:1px solid #45392c}
nav button{flex:1;background:none;border:0;color:#a99a84;padding:11px 4px;font:600 13px "Bahnschrift";border-bottom:2px solid transparent}
nav button.on{color:#2fb5a8;border-bottom-color:#2fb5a8}
#grp{display:flex;background:#0b0d12;border-bottom:1px solid #2e2720}
#grp .g{flex:1;text-align:center;font:700 9px "Bahnschrift";letter-spacing:.9px;color:#4a5462;padding:5px 0}
#grp .g.on{color:#2fb5a8}
.dot{color:#ff5c5c;font-style:normal;font-size:11px;vertical-align:super}
main{padding:10px}.card{background:#241f1a;border:1px solid #45392c;border-radius:8px;padding:10px;margin-bottom:8px}
.hero{font-weight:600;font-size:15px}.muted{color:#a99a84;font-size:12px}
.row{display:flex;align-items:center;gap:8px;padding:7px 0;border-bottom:1px solid #2e2720}
.mini{background:#383026;border:0;color:#a99a84;border-radius:6px;padding:8px 10px;font:600 13px "Bahnschrift"}
.fb.ok .mini{background:#27c66d;color:#1a1714}
.row.tgt{background:#132a1c;margin:0 -6px;padding-left:6px;padding-right:6px;border-radius:6px;box-shadow:inset 3px 0 0 #27c66d}
.row:last-child{border:0}.sku{font:600 13px Consolas,monospace;flex:1}.qty{color:#f5d24a;font:12px monospace;width:44px;text-align:right}
input{background:#120f0d;border:1px solid #45392c;color:#f2ece1;border-radius:6px;padding:9px;font:600 14px Consolas,monospace;width:112px;text-transform:uppercase}
input.ok{border-color:#27c66d;color:#27c66d}
button.act{background:#2fb5a8;color:#fff;border:0;border-radius:8px;padding:12px;font:600 14px "Bahnschrift";width:100%;margin-top:6px}
button.grn{background:#28a745}button.amb{background:#b45309}button.gry{background:#383026;color:#a99a84}
table{width:100%;border-collapse:collapse;font:11px Consolas,monospace}td{padding:3px 4px;border-bottom:1px solid #2e2720}
.t{color:#7a6c58;width:62px}.err{color:#ff6b6b}
.bar{height:8px;background:#120f0d;border-radius:4px;overflow:hidden;margin:8px 0 4px}
.bar i{display:block;height:100%;background:#27c66d}
.cnt{font:700 15px Consolas,monospace;text-align:right}
.fb{border-radius:8px;padding:12px;margin-bottom:8px;font:700 15px "Bahnschrift";background:#241f1a;color:#a99a84}
.fb.ok{background:#12301f;color:#27c66d}.fb.err{background:#3a1418;color:#ff6b6b}.fb.warn{background:#3a2c12;color:#f5b342}
.prow{display:flex;align-items:center;gap:6px;padding:9px 0;border-bottom:1px solid #2e2720}
.prow:last-child{border:0}.prow .mk{width:14px;color:#27c66d}.prow .loc{font:12px Consolas,monospace;color:#a99a84;width:82px}
.prow.sel{box-shadow:inset 3px 0 0 #2fb5a8}
.qpan{display:flex;gap:6px;align-items:center;padding:8px 0 10px}
.qpan input{width:74px;text-align:center;text-transform:none}
.qpan button{background:#383026;border:0;color:#f2ece1;border-radius:8px;padding:12px 14px;font:700 15px "Bahnschrift"}
.qpan .qmax{background:#2a3140;font-size:13px}
.qpan .qok{background:#28a745;color:#fff;flex:1}
.prow.done{opacity:.45}.prow.part{background:#2a2410;margin:0 -6px;padding-left:6px;padding-right:6px;border-radius:6px}
.sug{padding:0 0 8px 2px}.sug button em{font-style:normal;opacity:.6;font-size:11px;margin-left:4px}
.sug button{background:#132a1c;color:#27c66d;border:1px solid #27c66d;border-radius:6px;
padding:7px 10px;margin:2px 4px 0 0;font:600 12px Consolas,monospace}
</style></head><body>
<header><b>PickCore</b><span class="muted" id="op"></span>
<button id="snd" onclick="snd()">&#128263;</button><span id="ts"></span></header>
<div id="grp"><span class="g on">OUTBOUND</span><span class="g">INBOUND</span><span class="g">OPERATIONS</span><span class="g">SYSTEM</span></div>
<nav><button id="b0" class="on" onclick="tab(0)">Pick</button>
<button id="b1" onclick="tab(1)">Put-away</button>
<button id="b2" onclick="tab(2)">Reloc</button>
<button id="b3" onclick="tab(3)">Status</button></nav>
<main><div id="v0"></div><div id="v1" hidden></div><div id="v2" hidden></div><div id="v3" hidden></div></main>
<script>
let T=0,S={},dirty=false;
var seen={pick:0,inbound:0};
var qLine=-1,qVal=1,rLine=-1,rVal=1;
function rsel(i,q){if(rLine===i){rLine=-1}else{rLine=i;rVal=q||1}forceDraw=true;draw()}
function rSet(v){v=parseInt(v||'1',10);if(isNaN(v)||v<1)v=1;rVal=v;forceDraw=true;draw()}
function radd(d){rSet(rVal+d)}
function rsend(i){var v=rVal;rLine=-1;post('rel_qty',{i:i,n:v})}
function serpick(i){post('ser_pick',{i:i})}
function serpack(i,n){post('ser_pack',{i:i,n:n})}
function qsel(i,rem){if(qLine===i){qLine=-1}else{qLine=i;qVal=Math.min(1,rem)||1}forceDraw=true;draw()}
function qSet(v){v=parseInt(v||'1',10);if(isNaN(v)||v<1)v=1;qVal=v;forceDraw=true;draw()}
function qadd(d){qSet(qVal+d)}
function qsend(i){var v=qVal;qLine=-1;qVal=1;post('pick_qty',{i:i,n:v})}
function msend(){var e=document.getElementById('mcode');if(!e)return;var v=(e.value||'').trim();
 if(v){e.value='';post('pick_scan',{v:v})}}
function tab(n){T=n;for(let i=0;i<4;i++){document.getElementById('v'+i).hidden=(i!=n);
document.getElementById('b'+i).className=(i==n?'on':'')}
 forceDraw=true;
 var A=S.alerts||{};
 if(n===0)seen.pick=A.pick||0; if(n===1)seen.inbound=A.inbound||0;
 post('view',{v:['pick','inbound','reloc','status'][n]});draw()}
function tgt(i){post('target',{i:i})}
function typing(){var a=document.activeElement;return a&&(a.tagName==='INPUT'||a.tagName==='TEXTAREA')}
function post(op,extra){dirty=true;forceDraw=true;fetch('/api/cmd',{method:'POST',headers:{'Content-Type':'application/json'},
body:JSON.stringify(Object.assign({op:op},extra||{}))}).then(()=>setTimeout(()=>{dirty=false;pull()},220))}
function esc(s){return (s||'').replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]))}
function draw(){if(!S.inb)return;
document.getElementById('op').textContent=S.op||'';document.getElementById('ts').textContent=S.ts||'';
// Nie ruszamy DOM, gdy operator pisze - przerysowanie zabija pole i zamyka klawiature.
// Dotyczy WSZYSTKICH widokow (wczesniej tylko Inbound - stad problem z wpisywaniem w Pick).
if(typing())return;
var A=S.alerts||{};
var gs=document.querySelectorAll('#grp .g');
for(var gi=0;gi<gs.length;gi++)gs[gi].className='g'+(gi===T?' on':'');
document.getElementById('b0').innerHTML='Pick'+((A.pick||0)>seen.pick&&T!==0?' <i class="dot">&#9679;</i>':'');
document.getElementById('b1').innerHTML='Put-away'+((A.inbound||0)>seen.inbound&&T!==1?' <i class="dot">&#9679;</i>':'');
if(T===0)seen.pick=A.pick||0; if(T===1)seen.inbound=A.inbound||0;
let P=S.pick,ph='';
if(!P.lines.length){ph='<div class="card muted">'+esc(P.title||'No active pick')+'</div>'}
else{let pc=P.need?Math.round(P.done/P.need*100):0;
ph='<div class="card"><div class="hero">'+esc(P.title)+'</div><div class="muted">'+esc(P.sub)+'</div>'+
'<div class="bar"><i style="width:'+pc+'%"></i></div><div class="cnt">'+P.done+' / '+P.need+
(P.queue?' &middot; '+P.queue+' queued':'')+'</div></div>'+
'<div class="fb '+esc(P.fbk)+'">'+esc(P.fb)+'</div><div class="card">';
P.lines.forEach(l=>{var rem=l.nd-l.sc;
ph+='<div class="prow '+l.st+(l.i===qLine?' sel':'')+'" onclick="qsel('+l.i+','+rem+')"><span class="mk">'+
(l.st=='done'?'&#10003;':(l.st=='part'?'&#9654;':''))+
'</span><span class="sku">'+(l.ser?'&#128278; ':'')+esc(l.sku)+'</span><span class="loc">'+esc(l.bin)+
'</span><span class="qty">'+l.sc+'/'+l.nd+'</span></div>';
if(l.i===qLine&&rem>0){ph+='<div class="qpan"><button onclick="event.stopPropagation();qadd(-1)">&minus;</button>'+
'<input id="qv" type="number" inputmode="numeric" value="'+qVal+'" onclick="event.stopPropagation()" onchange="qSet(this.value)">'+
'<button onclick="event.stopPropagation();qadd(1)">+</button>'+
'<button class="qmax" onclick="event.stopPropagation();qSet('+rem+')">all '+rem+'</button>'+
'<button class="qok" onclick="event.stopPropagation();qsend('+l.i+')">Add</button></div>'}});
ph+='</div>';
var SM=S.ser||{on:false,radios:[],packs:{}};
if(SM.on&&SM.radios.length){
 ph+='<div class="card"><div class="hero" style="font-size:14px">&#9889; SERIAL SESSION</div>'+
 '<div class="muted">Tap a radio to scan its serials \u00b7 then pick pack size</div></div><div class="card">';
 SM.radios.forEach(function(r){
  var done=r.n>=r.need;
  ph+='<div class="prow '+(done?'done':(r.n>0?'part':''))+(r.act?' sel':'')+'" onclick="serpick('+r.i+')">'+
  '<span class="mk">'+(done?'&#10003;':(r.act?'&#9654;':''))+'</span>'+
  '<span class="sku">'+esc(r.sku)+'</span><span class="loc">'+esc(r.bin)+'</span>'+
  '<span class="qty">'+r.n+'/'+r.need+'</span></div>';
  if(r.act){ph+='<div class="qpan">';
   Object.keys(SM.packs).forEach(function(name){var v=SM.packs[name];
    ph+='<button class="'+(r.pack===v?'qok':'')+'" onclick="event.stopPropagation();serpack('+r.i+','+v+')">'+esc(name)+'</button>'});
   ph+='</div>'}});
 ph+='</div>'}
ph+='<div class="card"><div class="muted">No barcode? Type the item code:</div>'+
'<div class="qpan"><input id="mcode" placeholder="ITEM CODE" onclick="event.stopPropagation()">'+
'<button class="qok" onclick="msend()">Scan</button></div></div>'+
'<button class="act gry" onclick="post(\'pick_undo\')">Remove last pick</button>'}
document.getElementById('v0').innerHTML=ph;
let i=S.inb,h='';
if(!i.lines.length){h='<div class="card muted">No active put-away.'+(i.pending?' '+i.pending+' pending.':'')+'</div>'}
var D=S.done||{};
if(D.doc){h=(h||'')+'<div class="card" style="border-color:#2a3f2f"><div class="muted">LAST CONFIRMED &middot; '+esc(D.ts)+'</div>'+
'<div class="hero" style="font-size:14px">'+esc(D.doc)+' &middot; PO '+esc(D.po)+'</div>'+
'<div class="muted">'+D.lines+' lines'+(D.missing?' &middot; <b style="color:#f5b342">'+D.missing+' without bin</b>':' &middot; all bins set')+
'</div><div class="muted">'+esc(D.file)+'</div></div>'}
else{h='<div class="card"><div class="hero">PO '+esc(i.po)+'</div><div class="muted">'+i.done+'/'+i.lines.length+' bins set'+(i.pending?' &middot; '+i.pending+' pending':'')+
'</div><div class="muted">Scan a location \u2014 it lands on the highlighted line. Tap a line to retarget.</div></div>';
var L=S.last||{};
if(L.id&&L.i>=0&&L.bin){h+='<div class="fb ok" style="display:flex;align-items:center;gap:8px">'+
'<span style="flex:1">line '+(L.i+1)+' &middot; '+esc(L.sku)+' &rarr; <b>'+esc(L.bin)+'</b></span>'+
'<button class="mini" onclick="post(\'bin_undo\')">Undo</button></div>'}
else if(L.id&&L.i>=0&&!L.bin){h+='<div class="fb warn">line '+(L.i+1)+' &middot; '+esc(L.sku)+' &mdash; cleared, rescan</div>'}
h+='<div class="card">';
var tg=(S.focus&&S.focus.i>=0)?S.focus.i:i.lines.findIndex(x=>!x.bin);
i.lines.forEach(l=>{h+='<div class="row'+(l.i===tg?' tgt':'')+'" onclick="tgt('+l.i+')"><span class="sku">'+esc(l.sku)+'</span><span class="qty">'+esc(''+l.qty)+'</span>'+
'<input value="'+esc(l.bin)+'" class="'+(l.bin?'ok':'')+'" placeholder="BIN" onchange="post(\'bin\',{i:'+l.i+',v:this.value})">'+
(l.bin?'<button class="mini" onclick="event.stopPropagation();post(\'bin_clear\',{i:'+l.i+'})">&#10005;</button>':'')+'</div>';
if(l.sug&&l.sug.length){h+='<div class="sug">';l.sug.forEach(b=>{var bn=(typeof b==='string')?b:b.b, q=(typeof b==='string')?null:b.q;
h+='<button onclick="post(\'bin\',{i:'+l.i+',v:\''+esc(bn)+'\'})">&#8627; '+esc(bn)+(q!=null?' <em>'+esc(''+q)+'</em>':'')+'</button>'});h+='</div>'}});
h+='</div><button class="act grn" onclick="post(\'confirm\')">Confirm put-away</button>';
if(i.pending)h+='<button class="act amb" onclick="post(\'next_pa\')">Next put-away ('+i.pending+')</button>'}
document.getElementById('v1').innerHTML=h;
let r=S.rel,p=r.pend||{},g='<div class="card"><div class="hero">'+(p.sku?esc(p.sku)+(p.frm?' &rarr; FROM '+esc(p.frm):'')+' &hellip;':'Scan ITEM &rarr; FROM &rarr; TO')+'</div>'+
'<div class="muted">'+r.total+' moves this session</div></div>';
if(r.rows.length){g+='<div class="card">';
r.rows.forEach(x=>{g+='<div class="prow'+(x.i===rLine?' sel':'')+'" onclick="rsel('+x.i+','+x.qty+')">'+
'<span class="sku">'+esc(x.sku)+'</span><span class="loc">'+esc(x.frm)+' &rarr; '+esc(x.to)+'</span>'+
'<span class="qty">'+x.qty+'</span></div>';
if(x.i===rLine){g+='<div class="qpan"><button onclick="event.stopPropagation();radd(-1)">&minus;</button>'+
'<input type="number" inputmode="numeric" value="'+rVal+'" onclick="event.stopPropagation()" onchange="rSet(this.value)">'+
'<button onclick="event.stopPropagation();radd(1)">+</button>'+
'<button class="qok" onclick="event.stopPropagation();rsend('+x.i+')">Set qty</button></div>'}});
g+='</div><button class="act gry" onclick="post(\'rel_undo\')">Remove last</button>'}
document.getElementById('v2').innerHTML=g;
var C=S.scn||{},sok=(C.ago>=0&&C.ago<120);
let s='<div class="card" style="border-color:'+(sok?'#2a3f2f':(C.on?'#4a3a1a':'#4a1f22'))+'">'+
'<div class="hero" style="font-size:14px">'+(sok?'&#128225; Scanner connected':(C.on?'&#128225; Waiting for scans':'&#9888; Listener OFF'))+'</div>'+
'<div class="muted">DataWedge IP output must point to <b>'+esc(C.ip||'?')+':'+(C.port||0)+'</b></div>'+
'<div class="muted">'+(C.ago>=0?('last scan '+C.ago+' s ago'):'no scan received yet')+
(C.peers&&C.peers.length?(' &middot; from '+esc(C.peers.join(', '))):'')+'</div>'+
(C.err?'<div class="muted err">'+esc(C.err)+'</div>':'')+
(C.ago<0?'<div class="muted" style="margin-top:6px;line-height:1.6">Nothing received yet. Check on the scanner:<br>'+
'1. DataWedge &rarr; Settings &rarr; <b>DataWedge enabled</b><br>'+
'2. Profile <b>PickCore</b> &rarr; <b>Associated apps</b> must contain <b>Chrome</b> (activity *)<br>'+
'3. Profile &rarr; <b>Barcode input</b> enabled, <b>Keystroke output</b> disabled<br>'+
'4. Profile &rarr; <b>IP output</b> enabled, TCP, address above<br>'+
'<i>Quick test: open the address bar in Chrome and scan. If the code types itself in, '+
'the PickCore profile is asleep &mdash; fix step 2.</i></div>':'')+'</div>'+
'<div class="card"><div class="hero">'+esc(S.pick.title)+'</div><div class="muted">'+esc(S.pick.badge)+'</div></div><div class="card"><table>';
(S.log||[]).forEach(e=>{s+='<tr><td class="t">'+esc(e.t)+'</td><td>'+esc(e.m)+'</td></tr>'});
document.getElementById('v3').innerHTML=s+'</table></div><div class="muted">PickCore '+esc(S.ver)+'</div>'}
var AC=null,lastSnd=-1,lastSig='',forceDraw=false;
function snd(){
 try{ if(!AC){AC=new (window.AudioContext||window.webkitAudioContext)();}
  AC.resume().then(()=>{document.getElementById('snd').className='on';
   document.getElementById('snd').innerHTML='&#128266;';beepSeq([[1200,70],[1700,130]]);});
 }catch(e){}}
function beepSeq(seq){
 if(!AC||AC.state!=='running'||!seq)return;
 var t=AC.currentTime;
 seq.forEach(function(p){
  var f=p[0],d=p[1]/1000;
  var o=AC.createOscillator(),g=AC.createGain();
  o.type='square';o.frequency.value=f;
  g.gain.setValueAtTime(0.0001,t);g.gain.exponentialRampToValueAtTime(0.28,t+0.008);
  g.gain.exponentialRampToValueAtTime(0.0001,t+d);
  o.connect(g);g.connect(AC.destination);o.start(t);o.stop(t+d+0.01);
  t+=d+0.02;});}
var audioCache={};
function playFile(kind){
 try{var a=audioCache[kind];
  if(!a){a=new Audio('/sound/'+kind);a.preload='auto';audioCache[kind]=a;}
  a.currentTime=0;var pr=a.play();if(pr&&pr.catch)pr.catch(function(){beepSeq((S.tones||{})[kind]);});
  return true;}catch(e){return false;}}
function playIfNew(){
 if(!S.snd)return;
 if(lastSnd<0){lastSnd=S.snd.id;return;}          // pierwszy poll nie gra historii
 if(S.snd.id!==lastSnd){lastSnd=S.snd.id;
  var k=S.snd.k,files=S.snd.files||[];
  // wlasny plik operatora ma pierwszenstwo; ton generowany zostaje jako zapasowy
  if(files.indexOf(k)>=0){if(playFile(k))return;}
  beepSeq((S.tones||{})[k]);}}
function pull(){if(dirty)return;
 var v=['pick','inbound','reloc','status'][T];
 var tg=(S.focus&&S.focus.i!=null)?S.focus.i:-1;
 fetch('/api/state?view='+v+'&i='+tg).then(r=>r.json()).then(d=>{S=d;draw();playIfNew()}).catch(()=>{})}
pull();setInterval(pull,1500);
</script></body></html>"""

    def _http_srv(port):
        from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
        import urllib.parse
        class _H(BaseHTTPRequestHandler):
            def _send(self, body, ctype="text/html; charset=utf-8", code=200):
                b = body.encode("utf-8") if isinstance(body, str) else body
                self.send_response(code)
                self.send_header("Content-Type", ctype)
                self.send_header("Cache-Control", "no-store")
                self.send_header("Content-Length", str(len(b)))
                self.end_headers()
                try: self.wfile.write(b)
                except Exception: pass
            def do_GET(self):
                path = self.path.split("?")[0]
                if path == "/setup":
                    self._send(build_setup_page(bool(find_chrome())), "text/html; charset=utf-8")
                    return
                if path.startswith("/sound/"):
                    kind = path.split("/")[-1]
                    f = find_sound(cfg, kind, SOUND_EXT_WEB) if kind in SOUND_KINDS else ""
                    if not f:
                        self.send_response(404); self.end_headers(); return
                    ct = {"mp3": "audio/mpeg", "ogg": "audio/ogg",
                          "wav": "audio/wav", "m4a": "audio/mp4"}.get(f.rsplit(".", 1)[-1].lower(), "audio/mpeg")
                    try:
                        data = open(f, "rb").read()
                    except Exception:
                        self.send_response(404); self.end_headers(); return
                    self.send_response(200)
                    self.send_header("Content-Type", ct)
                    self.send_header("Content-Length", str(len(data)))
                    self.send_header("Cache-Control", "max-age=3600")
                    self.end_headers(); self.wfile.write(data); return
                if path == "/api/state":
                    # Heartbeat: on every poll the console reports where the operator is standing.
                    # Without it the focus expired after 300 s: tap a tab at the desk,
                    # walk to the rack, scan, and the scan landed in the wrong view.
                    try:
                        q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
                        v = (q.get("view", [""])[0] or "").strip()
                        if v in ("pick", "inbound", "reloc", "status"):
                            web_focus["view"] = v; web_focus["ts"] = time.time()
                            iv = q.get("i", [""])[0]
                            if iv.lstrip("-").isdigit(): web_focus["i"] = int(iv)
                    except Exception:
                        pass
                    self._send(web_state.get("json") or "{}", "application/json; charset=utf-8")
                elif path in ("/board", "/live"):
                    self._send(cloud_state.get("html") or "<html><body>warming up</body></html>")
                else:
                    self._send(HANDHELD_HTML)
            def do_POST(self):
                if self.path.split("?")[0] != "/api/cmd":
                    self._send("not found", "text/plain", 404); return
                try:
                    n = int(self.headers.get("Content-Length") or 0)
                    cmd = json.loads(self.rfile.read(n).decode("utf-8") or "{}")
                    if isinstance(cmd, dict) and cmd.get("op"):
                        web_cmd_q.put(cmd)
                    self._send('{"ok":true}', "application/json")
                except Exception as e:
                    self._send(json.dumps({"ok": False, "err": str(e)}), "application/json", 400)
            def log_message(self, *a): pass
        try:
            srv = ThreadingHTTPServer(("0.0.0.0", port), _H)
            http_state["on"] = True
            srv.serve_forever()
        except Exception as e:
            http_state["err"] = str(e)
    if cfg.get("http_serve"):
        threading.Thread(target=_http_srv, args=(int(cfg.get("http_port") or 8080),), daemon=True).start()

    def _net_boot_report():
        """An explicit network-services report in Logs, so a failed start is never silent."""
        if cfg.get("scanner_listen"):
            if scanner_state.get("on"):
                logln(f"\U0001F4E1 Scanner listener ON \u00B7 DataWedge IP output \u2192 {local_ip()}:{cfg.get('scanner_port') or 5577}")
            else:
                logln(f"\u2715 Scanner listener FAILED: {scanner_state.get('err') or 'not started'}")
        if cfg.get("http_serve"):
            if http_state.get("on"):
                logln(f"\U0001F310 Handheld console ON \u00B7 http://{local_ip()}:{cfg.get('http_port') or 8080}")
            else:
                logln(f"\u2715 Live view FAILED: {http_state.get('err') or 'not started'}")
    root.after(1800, _net_boot_report)

    _TEST_HOOKS.update(pa_q=pa_q, pa=pa, pa_pending=pa_pending, load_next_pa=load_next_pa,
                       confirm_putaway=confirm_putaway, drain_pa=drain_pa,
                       pa_cv=_cv, pa_in=_in, pa_load_pdf=pa_load_pdf, refresh_stats=refresh_stats,
                       scan_net_q=scan_net_q, scanner_state=scanner_state,
                       drain_scanner=drain_scanner, cloud_tick=cloud_tick, live_ring=live_ring,
                       sb=sb, sb_load=_sb_load, sb_copy=_sb_copy, sb_setup=_sb_setup, sb_collect=_sb_collect, sbv=sbv,
                       sb_weight=sb_weight, sb_parcels=sb_parcels,
                       rel=rel, rel_feed=rel_feed, rel_copy_tsv=rel_copy_tsv, trel=trel, nb=nb,
                       web_cmd_q=web_cmd_q, web_state=web_state, drain_web=drain_web, web_focus=web_focus,
                       render_web_state=render_web_state, http_state=http_state,
                       bc_q=bc_q, drain_bc=drain_bc, cfg=cfg,
                       station=station, refresh_station=refresh_station, on_scan=on_scan,
                       save_settings=save_settings,
                       serial_mode=serial_mode,
                       scan_var=scan_var, feedback=feedback)
    try:
        _sp = _sess_path()
        if os.path.exists(_sp):
            _d = json.loads(Path(_sp).read_text(encoding="utf-8"))
            if _d.get("items") and messagebox.askyesno("Resume session",
                    f"Unfinished serial session found ({len(_d['items'])} item(s)). Resume?"):
                session["items"]=_d["items"]; session["active"]=_d.get("active")
                session["await"]=_d.get("await") or "sku"; session["on"]=True
                session["seen"]={s for it in _d["items"] for s in it.get("serials",[])}
                btn_session.config(text="✕  End session", bg="#B02E0C", fg="white")
                pack_fr.pack(fill="x", padx=16, pady=(0,2))
                feedback("⚡ Session resumed.", "info"); refresh_station()
            else:
                sess_clear_file()
    except Exception: pass
    try:
        _pp=_pa_path()
        if os.path.exists(_pp):
            _pd=json.loads(Path(_pp).read_text(encoding="utf-8"))
            if (_pd.get("lines") or _pd.get("po")) and messagebox.askyesno("Resume inbound",
                    f"Unfinished INBOUND session (PO {_pd.get('po') or '—'}, {len(_pd.get('lines',[]))} lines). Resume?"):
                pa.update(po=_pd.get("po",""), doc=_pd.get("doc",""), lines=_pd.get("lines",[]),
                          active=_pd.get("active"), await_=None)
                pa["await"]=_pd.get("await") or ("sku" if pa["po"] else "po")
                pa_po_var.set(pa["po"]); rebuild_rows(); pa_feedback("\U0001F4E6 Session restored (PO " + (pa["po"] or "\u2014") + ").","info")
            else:
                pa_clear_file()
    except Exception: pass
    scan_entry.focus_set()
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()

if __name__ == "__main__":
    run_gui()
