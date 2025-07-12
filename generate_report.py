#!/usr/bin/env python3
"""
generate_report.py
==================
Smart CSV → PDF report generator (v3.1 – 13-07-2025)
---------------------------------------------------

**Novità v3.1**
* **Titolo** ora è il nome file **senza .csv** (es. `ALARM_MC-BATCH3120725144147`).
* Sottotitolo (solo per modalità *other*) diventa:
  `Starting Date: dd/mm/yyyy HH:MM:SS – Ending Date: dd/mm/yyyy HH:MM:SS`.
* **Retry automatico**: se il primo tentativo di generazione fallisce (file appena
  creato, lock temporaneo, ecc.) il programma aspetta 1 s e riprova una volta.

Restano invariati i tre layout (ALARM, OPERLOG, other) e le colonne con larghezza
fissa per *Date* e *Time*.
"""

from __future__ import annotations

import argparse
import glob
import time
from pathlib import Path
from typing import List, Tuple

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)

# ---------------------------------------------------------------------------
# Config costanti
# ---------------------------------------------------------------------------
NARROW_COLS = {"date", "time"}
NARROW_W_MM = 25  # larghezza fissa Date & Time
TEMP_COLS = [
    "TEMP_AIR_IN",
    "TEMP_PRODUCT_1",
    "TEMP_PRODUCT_2",
    "TEMP_PRODUCT_3",
]
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
    s = csv_path.stem.upper()
    if s.startswith("ALARM"):
        return "alarm"
    if s.startswith("OPERLOG"):
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
    df = df.loc[:, ~df.columns.str.match(r"^Unnamed")]  # cleanup

    cmap = {c.lower(): c for c in df.columns}
    selected = []
    for want in MODE_COLS[mode]:
        key = want.lower()
        if key not in cmap:
            raise ValueError(f"Colonna '{want}' mancante in {csv_path.name}")
        selected.append(cmap[key])
    df = df[selected]

    if mode == "other":
        for col in TEMP_COLS:
            key = col.lower()
            if key in cmap:
                df[cmap[key]] = pd.to_numeric(df[cmap[key]], errors="coerce").round(1)

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

    # other mode
    batch = df.get("BATCH_NAME", pd.Series([csv_path.stem])).iloc[0]
    if "Date" in df.columns:
        ts = pd.to_datetime(df["Date"], errors="coerce")
        st, en = ts.min(), ts.max()
        subtitle = (
            f"Starting Date: {st:%d/%m/%Y %H:%M:%S} – "
            f"Ending Date: {en:%d/%m/%Y %H:%M:%S}"
        )
    else:
        subtitle = ""
    return batch, subtitle

# ---------------------------------------------------------------------------
# PDF builder utils
# ---------------------------------------------------------------------------

def _col_widths(df: pd.DataFrame, page_w: float) -> List[float]:
    avail = page_w - 20 * mm  # 10 mm per lato
    narrow = [c for c in df.columns if c.lower() in NARROW_COLS]
    wide = [c for c in df.columns if c.lower() not in NARROW_COLS]
    narrow_w = NARROW_W_MM * mm
    rem = max(avail - len(narrow) * narrow_w, 10 * mm)
    wide_w = rem / max(len(wide), 1)
    return [narrow_w if c.lower() in NARROW_COLS else wide_w for c in df.columns]

# ---------------------------------------------------------------------------
# Build PDF
# ---------------------------------------------------------------------------

def _build_pdf(df: pd.DataFrame, csv_path: Path, pdf_path: Path, mode: str, font: int) -> None:
    title, subtitle = _title_sub(df, csv_path, mode)

    styles = getSampleStyleSheet()
    cell = ParagraphStyle("cell", parent=styles["BodyText"], fontSize=font, leading=font+1, alignment=1, wordWrap="LTR")
    head = ParagraphStyle("head", parent=styles["Heading4"], fontSize=font+1, leading=font+2, alignment=1)

    elems: List = [Paragraph(title, styles["Title"])]
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

    doc = SimpleDocTemplate(
        str(pdf_path), pagesize=landscape(A4),
        leftMargin=10*mm, rightMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm
    )
    doc.build(elems)

# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _run(csv_path: Path, font_size: int) -> None:
    mode = _detect_mode(csv_path)
    df = _load_df(csv_path, mode)
    pdf_path = csv_path.with_suffix(".pdf")
    _build_pdf(df, csv_path, pdf_path, mode, font_size)
    print(f"✅ Report creato: {pdf_path}")


def main() -> None:
    ap = argparse.ArgumentParser(description="Genera PDF dal CSV più recente o specificato.")
    ap.add_argument("csvfile", nargs="?", type=Path, help="CSV di origine (opzionale)")
    ap.add_argument("--font-size", type=int, default=7, help="Dimensione font tabella")
    args = ap.parse_args()

    csv_path = args.csvfile.expanduser().resolve() if args.csvfile else _latest_csv()
    if not csv_path.exists():
        ap.error(f"File non trovato: {csv_path}")

    # Retry (per doppio-click che a volte fallisce per lock temporanei)
    for attempt in (1, 2):
        try:
            _run(csv_path, args.font_size)
            break
        except Exception as err:
            if attempt == 2:
                raise
            print("⚠️  Errore:", err, "– ritento tra 1 s…")
            time.sleep(1)


if __name__ == "__main__":
    main()
