#!/usr/bin/env python3
"""
generate_report.py
==================
Smart CSV → PDF + auto‑print (v3.3 – 13‑07‑2025)
------------------------------------------------

**Novità v3.3**
* **Log persistente**: tutto ciò che prima andava su `stdout` ora finisce anche
  in `generate_report.log` nella stessa cartella dell’EXE.
* Se la stampa fallisce, un **popup** di Windows avvisa l’utente e apre il file
  di log in Notepad.
* Argomento `--log-console` per forzare la stampa dei log su console (utile se
  ricompili SENZA `--noconsole`).

Il resto (auto‑print 60 s, colonne, retry, titolo senza `.csv`) rimane invariato.
"""

from __future__ import annotations

import argparse
import glob
import logging
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import List, Tuple

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------
LOG_PATH = Path(__file__).with_suffix(".log")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),  # gets ignored if --noconsole build
    ],
)
log = logging.getLogger("report")

# ---------------------------------------------------------------------------
# Config costanti
# ---------------------------------------------------------------------------
NARROW_COLS = {"date", "time"}
NARROW_W_MM = 25
TEMP_COLS = ["TEMP_AIR_IN", "TEMP_PRODUCT_1", "TEMP_PRODUCT_2", "TEMP_PRODUCT_3"]
MODE_COLS = {
    "alarm": ["Date", "Time", "Alarm Message", "Alarm Status"],
    "operlog": [
        "Date",
        "Time",
        "User",
        "Object_Action",
        "Trigger",
        "PreviousValue",
        "ChangedValue",
    ],
    "other": ["Date", "Time", "USER", *TEMP_COLS],
}

# ---------------------------------------------------------------------------
# Helper CSV
# ---------------------------------------------------------------------------

def _find_header(csv_path: Path, prefix: str = "Date\t") -> int:
    with csv_path.open("r", encoding="utf-8", errors="ignore") as fp:
        for i, line in enumerate(fp):
            if line.startswith(prefix):
                return i
    raise ValueError("Intestazione non trovata (Date\t…)")


def _detect_mode(csv_path: Path) -> str:
    stem = csv_path.stem.upper()
    if stem.startswith("ALARM"):
        return "alarm"
    if stem.startswith("OPERLOG"):
        return "operlog"
    return "other"


def _latest_csv() -> Path:
    files = [Path(p) for p in glob.glob("*.csv")]
    if not files:
        raise FileNotFoundError("Nessun file CSV trovato nella cartella corrente.")
    return max(files, key=lambda p: p.stat().st_mtime)


def _load_df(csv_path: Path, mode: str) -> pd.DataFrame:
    skip = _find_header(csv_path)
    df = pd.read_csv(csv_path, sep="\t", skiprows=skip)
    df = df.loc[:, ~df.columns.str.match(r"^Unnamed")]

    cmap = {c.lower(): c for c in df.columns}
    cols = []
    for want in MODE_COLS[mode]:
        key = want.lower()
        if key not in cmap:
            raise ValueError(f"Colonna '{want}' mancante in {csv_path.name}")
        cols.append(cmap[key])
    df = df[cols]

    if mode == "other":
        for col in TEMP_COLS:
            if col.lower() in cmap:
                df[cmap[col.lower()]] = pd.to_numeric(df[cmap[col.lower()]], errors="coerce").round(1)

    if {"Date", "Time"}.issubset(df.columns):
        df["_TS"] = pd.to_datetime(df["Date"] + " " + df["Time"], errors="coerce")
        df = df.sort_values("_TS").drop(columns="_TS")
    return df

# ---------------------------------------------------------------------------
# Titolo / sottotitolo
# ---------------------------------------------------------------------------

def _title_sub(df: pd.DataFrame, csv_path: Path, mode: str) -> Tuple[str, str]:
    if mode in {"alarm", "operlog"}:
        return csv_path.stem, ""

    batch = df.get("BATCH_NAME", pd.Series([csv_path.stem])).iloc[0]
    if "Date" in df.columns:
        ts = pd.to_datetime(df["Date"], errors="coerce")
        st, en = ts.min(), ts.max()
        sub = f"Starting Date: {st:%d/%m/%Y %H:%M:%S} – Ending Date: {en:%d/%m/%Y %H:%M:%S}"
    else:
        sub = ""
    return batch, sub

# ---------------------------------------------------------------------------
# PDF utilities
# ---------------------------------------------------------------------------

def _col_widths(df: pd.DataFrame, page_w: float) -> List[float]:
    avail = page_w - 20 * mm
    nar = [c for c in df.columns if c.lower() in NARROW_COLS]
    wid = [c for c in df.columns if c.lower() not in NARROW_COLS]
    nw = NARROW_W_MM * mm
    rem = max(avail - len(nar) * nw, 10 * mm)
    ww = rem / max(len(wid), 1)
    return [nw if c.lower() in NARROW_COLS else ww for c in df.columns]

# ---------------------------------------------------------------------------
# Build PDF & print
# ---------------------------------------------------------------------------

def _build_pdf(df: pd.DataFrame, csv_path: Path, pdf_path: Path, mode: str, font: int) -> None:
    from reportlab.platypus import BaseDocTemplate  # keep PyInstaller happy

    title, subtitle = _title_sub(df, csv_path, mode)
    styles = getSampleStyleSheet()
    cell = ParagraphStyle("cell", parent=styles["BodyText"], fontSize=font, leading=font+1, alignment=1, wordWrap="LTR")
    head = ParagraphStyle("head", parent=styles["Heading4"], fontSize=font+1, leading=font+2, alignment=1)

    elems = [Paragraph(title, styles["Title"])]
    if subtitle:
        elems.append(Paragraph(subtitle, styles["Heading2"]))
    elems.append(Spacer(1, 12))

    data = [[Paragraph(str(c), head) for c in df.columns]]
    for _, row in df.iterrows():
        data.append([Paragraph(str(v), cell) for v in row])

    pw, _ = landscape(A4)
    table = Table(data, colWidths=_col_widths(df, pw), repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("GRID", (0,0), (-1,-1), 0.25, colors.black),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
    ]))
    elems.append(table)

    doc = SimpleDocTemplate(str(pdf_path), pagesize=landscape(A4), leftMargin=10*mm, rightMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
    doc.build(elems)


def _popup(msg: str) -> None:
    if os.name != "nt":
        return
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "generate_report", 0x30)  # MB_ICONWARNING
    except Exception:
        pass


def _print_after_delay(pdf_path: Path, delay: int = 20) -> None:
    if os.name != "nt":
        log.info("Stampa saltata (non Windows)")
        return
    log.info("Attendo %ss prima di stampare…", delay)
    time.sleep(delay)
    try:
        os.startfile(str(pdf_path), "print")  # type: ignore[attr-defined]
        log.info("PDF inviato alla stampante predefinita")
    except Exception as e:
        log.error("Errore stampa: %s", e)
        _popup("Impossibile stampare il report. Controlla generate_report.log per dettagli.")
        subprocess.Popen(["notepad.exe", str(LOG_PATH)])

# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _run(csv_path: Path, font: int, no_print: bool) -> None:
    mode = _detect_mode(csv_path)
    df = _load_df(csv_path, mode)
    pdf_path = csv_path.with_suffix(".pdf")
    _build_pdf(df, csv_path, pdf_path, mode, font)
    log.info("Report creato: %s", pdf_path.name)
    if not no_print:
        _print_after
