#!/usr/bin/env python3
"""
generate_report.py
==================
Smart CSV → PDF + auto‑print (v4.0 – 14‑07‑2025)
------------------------------------------------

Eseguito da doppio‑click (nessun parametro necessario):
1. Trova il **file .csv più recente** nella stessa cartella.
2. Genera il PDF con layout dedicato (ALARM / OPERLOG / Other).
3. Scrive un log persistente `generate_report.log` accanto all’EXE.
4. Dopo 60 s invia il PDF alla **stampante predefinita**.
   * Se la stampa fallisce compare un popup di Windows e si apre Notepad col log.

Dipendenze (windows): pandas, reportlab, pywin32.
"""

from __future__ import annotations

import glob
import logging
import os
import sys
import time
import subprocess
from pathlib import Path
from typing import List, Tuple

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

try:
    import win32api, win32print  # type: ignore
except ImportError:
    win32api = win32print = None  # fallback su os.startfile

# ---------------------------------------------------------------------------
# Config costanti
# ---------------------------------------------------------------------------
FONT_SIZE = 7
PRINT_DELAY_S = 60
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
# Logging (file accanto all’EXE)
# ---------------------------------------------------------------------------
EXE_DIR = Path(getattr(sys, "frozen", False) and sys.executable or __file__).parent
LOG_PATH = (EXE_DIR / "generate_report.log").resolve()
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_PATH, encoding="utf-8")],
)
log = logging.getLogger("report")

# ---------------------------------------------------------------------------
# Utility CSV
# ---------------------------------------------------------------------------

def _find_header(csv: Path, prefix: str = "Date\t") -> int:
    with csv.open("r", encoding="utf-8", errors="ignore") as fp:
        for i, line in enumerate(fp):
            if line.startswith(prefix):
                return i
    raise ValueError("Intestazione non trovata (Date\t…)")


def _latest_csv() -> Path:
    files = [Path(p) for p in glob.glob(str(EXE_DIR / "*.csv"))]
    if not files:
        raise FileNotFoundError("Nessun file CSV trovato nella cartella.")
    return max(files, key=lambda p: p.stat().st_mtime)


def _detect_mode(csv: Path) -> str:
    s = csv.stem.upper()
    if s.startswith("ALARM"):
        return "alarm"
    if s.startswith("OPERLOG"):
        return "operlog"
    return "other"


def _load_df(csv: Path, mode: str) -> pd.DataFrame:
    skip = _find_header(csv)
    df = pd.read_csv(csv, sep="\t", skiprows=skip).loc[:, lambda d: ~d.columns.str.match(r"^Unnamed")]
    cmap = {c.lower(): c for c in df.columns}
    cols = [cmap[w.lower()] for w in MODE_COLS[mode] if w.lower() in cmap]
    if len(cols) != len(MODE_COLS[mode]):
        missing = set(MODE_COLS[mode]) - {c.title() for c in cols}
        raise ValueError(f"Colonne mancanti: {missing}")
    df = df[cols]
    if mode == "other":
        for t in TEMP_COLS:
            if t.lower() in cmap:
                df[cmap[t.lower()]] = pd.to_numeric(df[cmap[t.lower()]], errors="coerce").round(1)
    if {"Date", "Time"}.issubset(df.columns):
        df["_TS"] = pd.to_datetime(df["Date"] + " " + df["Time"], errors="coerce")
        df = df.sort_values("_TS").drop(columns="_TS")
    return df

# ---------------------------------------------------------------------------
# Titolo / sottotitolo
# ---------------------------------------------------------------------------

def _title_sub(df: pd.DataFrame, csv: Path, mode: str) -> Tuple[str, str]:
    if mode in {"alarm", "operlog"}:
        return csv.stem, ""
    batch = df.get("BATCH_NAME", pd.Series([csv.stem])).iloc[0]
    if "Date" in df.columns:
        ts = pd.to_datetime(df["Date"], errors="coerce")
        st, en = ts.min(), ts.max()
        sub = f"Starting Date: {st:%d/%m/%Y %H:%M:%S} – Ending Date: {en:%d/%m/%Y %H:%M:%S}"
    else:
        sub = ""
    return batch, sub

# ---------------------------------------------------------------------------
# PDF builder
# ---------------------------------------------------------------------------

def _col_widths(df: pd.DataFrame, pw: float) -> List[float]:
    avail = pw - 20 * mm
    nar = [c for c in df.columns if c.lower() in NARROW_COLS]
    nw = NARROW_W_MM * mm
    wid = len(df.columns) - len(nar)
    rem = max(avail - len(nar) * nw, 10 * mm)
    ww = rem / max(wid, 1)
    return [nw if c.lower() in NARROW_COLS else ww for c in df.columns]


def _build_pdf(df: pd.DataFrame, csv: Path, pdf: Path, mode: str) -> None:
    title, sub = _title_sub(df, csv, mode)
    styles = getSampleStyleSheet()
    cell = ParagraphStyle("cell", parent=styles["BodyText"], fontSize=FONT_SIZE, leading=FONT_SIZE+1, alignment=1, wordWrap="LTR")
    head = ParagraphStyle("head", parent=styles["Heading4"], fontSize=FONT_SIZE+1, leading=FONT_SIZE+2, alignment=1)
    elems = [Paragraph(title, styles["Title"])]
    if sub:
        elems.append(Paragraph(sub, styles["Heading2"]))
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
    doc = SimpleDocTemplate(str(pdf), pagesize=landscape(A4), leftMargin=10*mm, rightMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm)
    doc.build(elems)

# ---------------------------------------------------------------------------
# Stampa
# ---------------------------------------------------------------------------

def _popup(msg: str):
    if os.name != "nt":
        return
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, msg, "generate_report", 0x30)
    except Exception:
        pass


def _print_pdf(pdf: Path):
    if os.name != "nt":
        log.info("Stampa saltata (non Windows)")
        return
    time.sleep(PRINT_DELAY_S)
    try:
        if win32api and win32print:
            win32api.ShellExecute(0, "print", str(pdf), None, str(pdf.parent), 0)
        else:
            os.startfile(str(pdf), "print")  # type: ignore[attr-defined]
        log.info("PDF inviato alla stampante predefinita")
    except Exception as e:
        log.error("Errore stampa: %s", e)
        _popup("Impossibile stampare il report. Aprendo il file di log.")
        subprocess.Popen(["notepad.exe", str(LOG_PATH)])

# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    try:
        csv = _latest_csv()
        mode = _detect_mode(csv)
        log.info("CSV selezionato: %s (mode=%s)", csv.name, mode)
        df = _load_df(csv, mode)
        pdf = csv.with_suffix(".pdf")
        _build_pdf(df, csv, pdf, mode)
        log.info("PDF creato: %s", pdf.name)
        _print_pdf(pdf)
    except Exception as err:
        log.exception("Errore irreversibile: %s", err)
        _popup("Errore durante la generazione del report. Controlla il log.")
        raise

if __name__ == "__main__":
    main()
