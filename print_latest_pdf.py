#!/usr/bin/env python3
"""
print_latest_pdf.py

Stampa il PDF più recente nella cartella corrente usando la stampante predefinita di Windows.
"""

import os
import sys
import glob
from pathlib import Path
from time import sleep


def find_latest_pdf(folder="."):
    pdf_files = list(Path(folder).glob("*.pdf"))
    if not pdf_files:
        return None
    # Ordina per data di modifica (decrescente)
    pdf_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return pdf_files[0]


def print_pdf_windows(pdf_path):
    try:
        print(f"Invio alla stampante: {pdf_path}")
        os.startfile(str(pdf_path), "print")
        # Attendi qualche secondo per evitare che il processo termini troppo presto
        sleep(2)
        print("Stampa inviata con successo.")
    except Exception as e:
        print(f"ERRORE durante la stampa: {e}", file=sys.stderr)
        sys.exit(2)


def main():
    # Cerca il PDF più recente
    latest_pdf = find_latest_pdf()
    if not latest_pdf:
        print("Nessun file PDF trovato nella cartella corrente.", file=sys.stderr)
        sys.exit(1)

    # Stampa il PDF
    print_pdf_windows(latest_pdf)


if __name__ == "__main__":
    main() 