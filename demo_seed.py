# -*- coding: utf-8 -*-
"""PickCore demo seeder.

Generates a synthetic event log so the Impact dashboard and the warehouse heat
maps have something to render. Nothing here comes from a real operation: the
location codes follow the demo layout in ISO_BLOCKS, the SKUs use the example
catalogue prefixes, and the volumes are made up.

    python demo_seed.py        # writes Logs/pickcore_events.jsonl
    python pickcore.py         # then open Impact (F6) and Warehouse (F7)
"""
import json, os, random
from datetime import datetime, timedelta

random.seed(20260906)
BASE = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(BASE, "Logs")
os.makedirs(LOG_DIR, exist_ok=True)
OUT = os.path.join(LOG_DIR, "pickcore_events.jsonl")

RADIO = ["UNA", "UNB", "UNC", "TRX", "TRM", "HRT", "HRP"]
ACC = ["ACB", "ACC", "ACD", "ACK", "ACL", "AUD", "ANT", "BAT", "CBL", "CHG"]

def sku(fast=False):
    fam = random.choice(RADIO if random.random() < 0.3 else ACC)
    n = random.randint(1, 40) if fast else random.randint(1, 400)
    return f"{fam}{n:04d}A"

# Demo layout: zone 01 shelves (aisles A-D), zone 02 pallets (A-B), zone 03 pick face.
def bin_code():
    r = random.random()
    if r < 0.62:                                   # zone 01 carries most of the traffic
        z, a, b = "01", random.choice("ABCD"), random.randint(2, 23)
    elif r < 0.88:
        z, a, b = "02", random.choice("AB"), random.choice([2, 3, 4, 5, 6, 7, 9, 10, 12, 14, 16])
    else:
        z, a, b = "03", "A", random.choice([2, 4, 6, 8, 10])
    return f"{z}-{a}{b:02d}-A{random.randint(1, 4)}"

FAST = [sku(fast=True) for _ in range(18)]         # a small head of fast movers drives ABC

BIN_OF = {}
def bin_for(s):
    """A SKU keeps its home bin, so the heat map shows structure rather than noise."""
    if s not in BIN_OF:
        BIN_OF[s] = bin_code()
    return BIN_OF[s]

rows = []
def ev(ts, etype, **d):
    rec = {"ts": ts.isoformat(timespec="seconds"), "type": etype, "app": "1.0"}
    if etype != "PICK_TELEMETRY":
        rec["op"] = "DEMO"
    rec.update(d)
    rows.append(rec)

start = datetime.now() - timedelta(days=13)
pick_no = 4100

for day in range(14):
    d0 = start + timedelta(days=day)
    if d0.weekday() >= 5:                          # weekends are quiet
        picks = random.randint(2, 5)
    else:
        picks = random.randint(14, 26)
    for _ in range(picks):
        pick_no += 1
        t = d0.replace(hour=random.randint(7, 16), minute=random.randint(0, 59),
                       second=random.randint(0, 59))
        n_lines = random.randint(2, 11)
        lines, units = [], 0
        for i in range(n_lines):
            s = random.choice(FAST) if random.random() < 0.55 else sku()
            q = random.choice([1, 1, 1, 2, 2, 3, 5, 6, 10])
            units += q
            lines.append({"t": round(i * random.uniform(9, 34), 1),
                          "sku": s, "bin": bin_for(s), "qty": q})
        title = f"SO-{pick_no}"
        ev(t, "PICK_VERIFIED", pick=title, doc=f"D{pick_no}", units=units, lines=n_lines)
        ev(t, "PICK_TELEMETRY", pick=title, doc=f"D{pick_no}", n_lines=n_lines, units=units,
           cycle_s=round(random.uniform(180, 900), 1),
           validation_s=round(sum(l["t"] for l in lines) + random.uniform(20, 90), 1),
           queued=random.random() < 0.2, lines=lines)

    # Errors caught: the whole point of the validation station.
    for _ in range(random.randint(0, 4) if d0.weekday() < 5 else random.randint(0, 1)):
        t = d0.replace(hour=random.randint(7, 16), minute=random.randint(0, 59), second=0)
        kind = random.choices(["WRONG_ITEM", "OVERPICK", "BULK_BLOCKED"], [6, 3, 2])[0]
        ev(t, kind, pick=f"SO-{random.randint(4101, pick_no)}", code=sku(),
           reason={"WRONG_ITEM": "scanned item not on this pick",
                   "OVERPICK": "line already complete",
                   "BULK_BLOCKED": "qty outside the learned bulk profile"}[kind])

    if random.random() < 0.12:
        ev(d0.replace(hour=15, minute=20), "OVERRIDE", pick=f"SO-{pick_no}",
           reason="supervisor release")

    # Inbound side: put-away sessions.
    for _ in range(random.randint(1, 4) if d0.weekday() < 5 else 0):
        t = d0.replace(hour=random.randint(8, 15), minute=random.randint(0, 59), second=0)
        po = f"PO-{random.randint(50000, 59999)}"
        for _ in range(random.randint(3, 12)):
            s = sku()
            ev(t, "PA_LINE", po=po, sku=s, qty=random.choice([1, 2, 4, 6, 10, 24]),
               bin=bin_for(s))
        ev(t, "PA_CONFIRMED", po=po)
    if d0.weekday() < 5:
        ev(d0.replace(hour=16, minute=45), "PA_EXPORTED", file=f"putaway_{d0:%Y%m%d}.txt")

    # Relocations feed the location index.
    for _ in range(random.randint(0, 3)):
        t = d0.replace(hour=random.randint(9, 15), minute=random.randint(0, 59), second=0)
        s = sku()
        ev(t, "RELOC_LINE", sku=s, frm=bin_code(), to=bin_for(s), qty=1)

rows.sort(key=lambda r: r["ts"])
with open(OUT, "w", encoding="utf-8") as f:
    for r in rows:
        f.write(json.dumps(r, ensure_ascii=False) + "\n")

from collections import Counter
c = Counter(r["type"] for r in rows)
print(f"{OUT}")
print(f"{len(rows)} events over 14 days")
for k, v in sorted(c.items()):
    print(f"  {k:<16} {v}")
