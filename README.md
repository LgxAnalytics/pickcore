# PickCore

A warehouse picking cockpit. One Python file, Tkinter GUI, no server-side
dependencies. The tool answers a specific bottleneck: order picking verified
by eye against a paper printout, with no error control and no data on where
the time actually goes.

**Status:** portfolio project. The code is generic, input data is synthetic,
and no customer, supplier or operational identifiers are part of this
repository.

## What it does

| Module | Function |
|---|---|
| Document pipeline | `pdfplumber` parses the picking list into a structure, renders HTML, prints through headless Edge |
| Scan validation | validates scanned SKUs with variant tolerance and rejection of false matches |
| Speed-reject | inter-keystroke gap analysis separates scanner input from operator typing |
| Serial capture | captures serial numbers during picking, with undo |
| Put-away | records inbound receipts and bin relocations |
| Handheld console | HTTP console for Wi-Fi terminals, IP allowlist and token |
| Scanner listener | receives scans over TCP from DataWedge terminals |
| Pick heat map | heat map of rack visits, 2D elevation and isometric 3D views |
| Star schema export | exports to a star model for slotting analytics |
| Label generator | ZPL labels for Zebra printers, generated natively |

## Architecture

A single process with worker threads for anything that blocks: directory
watcher, PDF rendering, scanner listener, HTTP server. All communication with
the GUI layer goes through queues, because Tkinter is not thread safe.

State is persisted with atomic writes (temp file plus replace), after a silent
write failure once cost a full customer file. A mutex guards the instance so
two windows cannot write to the same files.

## Interface

The navigation rail starts collapsed to icons and expands on click or `Ctrl+B`.
A context bar above the content area carries the active view name and its
function key, so a collapsed rail never leaves the operator guessing. Views
register through a Notebook-compatible API, which keeps the shell replaceable
without touching any tab content.

## Browser bookmarklet

`bookmarklet.js` fills a third-party forwarder portal form from a shipment
payload held on the clipboard. The portal offers no API and no stable markup,
so fields are located by their visible label first, then by the text of the
surrounding cell, then by name/id/placeholder as a last resort. Nothing is ever
submitted: the operator reviews the form and presses Submit.

Two details make it survive real portals. Every filled field is tagged through
`dataset`, so a second run is idempotent and a later pass cannot overwrite an
earlier one. And because the form re-renders after country and parcel-count
changes, dropping values as it goes, applied fields are verified and re-applied
once before the summary is shown.

That file is the source of truth. `pickcore.py` carries a percent-encoded copy
generated from it, so the readable version is the one to edit.

## Running it

```
pip install pdfplumber pywin32
python pickcore.py
```

Building the executable:

```
python -m PyInstaller --onedir --windowed --name PickCore ^
    --collect-all pdfplumber --collect-all pdfminer pickcore.py
```

Onedir, not onefile. Onefile unpacked about 100 MB into `Temp\_MEIxxxxx` on
every start, and antivirus held handles during shutdown, which ended in a
failure to remove the temporary directory.

## Security

The HTTP console binds to the LAN because handheld terminals connect over
Wi-Fi. Access is gated twice: an IP allowlist with prefix matching (an entry
without the last octet admits a whole subnet) and a token passed once in the
URL and carried afterwards in a cookie. Integration secrets never reach the
config file or the source, only an environment variable or Windows Credential
Manager.

## License

MIT
