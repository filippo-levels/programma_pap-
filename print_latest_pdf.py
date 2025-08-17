#!/usr/bin/env python3
"""
print_latest_pdf.py

Stampa i PDF più recenti nella cartella corrente usando la stampante predefinita di Windows.
Se disponibili, stampa prima il PDF principale e poi il PDF temperature trend.

Copyright © 2024 Filippo Caliò
Version: 1.0.0
"""

import os
import sys
import glob
from pathlib import Path
from time import sleep


def find_latest_pdfs(folder="."):
    """Trova i PDF più recenti separando quelli principali dai temperature trend."""
    pdf_files = list(Path(folder).glob("*.pdf"))
    if not pdf_files:
        return None, None
    
    # Separa PDF principali e temperature trend
    main_pdfs = [f for f in pdf_files if not f.name.endswith('_temperature_trend.pdf')]
    trend_pdfs = [f for f in pdf_files if f.name.endswith('_temperature_trend.pdf')]
    
    # Trova il più recente di ogni tipo
    latest_main = max(main_pdfs, key=lambda f: f.stat().st_mtime) if main_pdfs else None
    latest_trend = max(trend_pdfs, key=lambda f: f.stat().st_mtime) if trend_pdfs else None
    
    return latest_main, latest_trend


def print_pdfs_windows(pdf_paths):
    """Stampa una lista di PDF in sequenza."""
    for pdf_path in pdf_paths:
        if pdf_path:
            try:
                print(f"Invio alla stampante: {pdf_path}")
                os.startfile(str(pdf_path), "print")
                # Attendi qualche secondo per evitare che il processo termini troppo presto
                sleep(2)
                print("Stampa inviata con successo.")
            except Exception as e:
                print(f"ERRORE durante la stampa di {pdf_path}: {e}", file=sys.stderr)
                # Continua con il prossimo PDF invece di terminare


def main():
    # Cerca i PDF più recenti
    latest_main, latest_trend = find_latest_pdfs()
    
    if not latest_main and not latest_trend:
        print("Nessun file PDF trovato nella cartella corrente.", file=sys.stderr)
        sys.exit(1)
    
    # Prepara la lista dei PDF da stampare
    pdfs_to_print = []
    if latest_main:
        pdfs_to_print.append(latest_main)
        print(f"PDF principale trovato: {latest_main}")
    
    if latest_trend:
        pdfs_to_print.append(latest_trend)
        print(f"PDF temperature trend trovato: {latest_trend}")
    
    # Stampa prima il principale, poi il temperature trend
    print(f"\nAvvio stampa sequenziale di {len(pdfs_to_print)} PDF...")
    print_pdfs_windows(pdfs_to_print)
    
    print("Stampa completata.")


if __name__ == "__main__":
    main() 