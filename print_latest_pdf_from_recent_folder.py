#!/usr/bin/env python3
"""
print_latest_pdf_from_recent_folder.py

Stampa il PDF più recente nella cartella DDMMYY più recente usando la stampante predefinita di Windows.

Copyright © 2024 Filippo Caliò
Version: 1.0.0
"""

import os
import sys
import glob
import re
from pathlib import Path
from time import sleep
from datetime import datetime


def parse_directory_date(dirname):
    """Converte nome directory DDMMYY in oggetto datetime."""
    match = re.match(r'^(\d{2})(\d{2})(\d{2})$', dirname)
    if not match:
        return None
    
    day, month, year = match.groups()
    
    # Assume secolo 20YY
    full_year = 2000 + int(year)
    
    try:
        return datetime(full_year, int(month), int(day))
    except ValueError:
        # Data non valida
        return None


def find_latest_pdf_in_recent_folder(base_directory="."):
    """Trova il PDF più recente nella cartella DDMMYY più recente."""
    print(f"Ricerca cartelle DDMMYY in: {base_directory}")
    
    # Trova tutte le sottocartelle con pattern DDMMYY
    date_dirs = []
    for item in os.listdir(base_directory):
        item_path = os.path.join(base_directory, item)
        if os.path.isdir(item_path):
            parsed_date = parse_directory_date(item)
            if parsed_date:
                date_dirs.append((parsed_date, item_path))
    
    if not date_dirs:
        print("Nessuna cartella DDMMYY trovata", file=sys.stderr)
        return None
    
    # Ordina per data (più recente prima)
    date_dirs.sort(key=lambda x: x[0], reverse=True)
    most_recent_dir = date_dirs[0][1]
    
    print(f"Cartella più recente: {most_recent_dir}")
    
    # Cerca PDF nella cartella
    pdf_files = list(Path(most_recent_dir).glob("*.pdf"))
    if not pdf_files:
        print(f"Nessun file PDF trovato in: {most_recent_dir}", file=sys.stderr)
        return None
    
    # Ordina per data di modifica (decrescente)
    pdf_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    latest_pdf = pdf_files[0]
    
    print(f"PDF più recente trovato: {latest_pdf}")
    return latest_pdf


def print_pdf_windows(pdf_path):
    """Stampa il PDF usando la stampante predefinita di Windows."""
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
    # Cerca il PDF più recente nella cartella DDMMYY più recente
    latest_pdf = find_latest_pdf_in_recent_folder()
    if not latest_pdf:
        print("Nessun file PDF trovato nelle cartelle DDMMYY.", file=sys.stderr)
        sys.exit(1)

    # Stampa il PDF
    print_pdf_windows(latest_pdf)


if __name__ == "__main__":
    main()
