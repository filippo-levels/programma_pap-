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
import logging
from pathlib import Path
from time import sleep
from datetime import datetime


def setup_logging():
    """Configura il sistema di logging per file e console."""
    # Crea logger
    logger = logging.getLogger('pdf_printer')
    logger.setLevel(logging.DEBUG)
    
    # Handler per file
    log_file = f"pdf_print_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    
    # Handler per console
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    
    # Formatter
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # Aggiungi handler
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger


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


def find_latest_pdf_in_recent_folder(base_directory=".", logger=None):
    """Trova il PDF più recente nella cartella DDMMYY più recente."""
    if logger:
        logger.info(f"Ricerca cartelle DDMMYY in: {os.path.abspath(base_directory)}")
    
    # Verifica che la directory base esista
    if not os.path.exists(base_directory):
        if logger:
            logger.error(f"Directory base non trovata: {base_directory}")
        return None
    
    # Trova tutte le sottocartelle con pattern DDMMYY
    date_dirs = []
    try:
        for item in os.listdir(base_directory):
            item_path = os.path.join(base_directory, item)
            if os.path.isdir(item_path):
                if logger:
                    logger.debug(f"Controllo directory: {item}")
                parsed_date = parse_directory_date(item)
                if parsed_date:
                    date_dirs.append((parsed_date, item_path))
                    if logger:
                        logger.debug(f"Directory DDMMYY valida trovata: {item} -> {parsed_date}")
                else:
                    if logger:
                        logger.debug(f"Directory non DDMMYY: {item}")
    except PermissionError as e:
        if logger:
            logger.error(f"Errore di permessi nell'accesso a {base_directory}: {e}")
        return None
    except Exception as e:
        if logger:
            logger.error(f"Errore imprevisto nella ricerca directory: {e}")
        return None
    
    if not date_dirs:
        if logger:
            logger.warning("Nessuna cartella DDMMYY trovata")
        return None
    
    # Ordina per data (più recente prima)
    date_dirs.sort(key=lambda x: x[0], reverse=True)
    most_recent_dir = date_dirs[0][1]
    
    if logger:
        logger.info(f"Cartella più recente: {most_recent_dir}")
        logger.info(f"Contenuto della cartella {most_recent_dir}:")
        try:
            for item in os.listdir(most_recent_dir):
                item_path = os.path.join(most_recent_dir, item)
                if os.path.isfile(item_path):
                    stat = os.stat(item_path)
                    logger.info(f"  File: {item} - Dimensione: {stat.st_size} bytes - Modificato: {datetime.fromtimestamp(stat.st_mtime)}")
                else:
                    logger.info(f"  Directory: {item}")
        except Exception as e:
            logger.error(f"Errore nel leggere contenuto directory: {e}")
    
    # Cerca PDF nella cartella
    try:
        pdf_files = list(Path(most_recent_dir).glob("*.pdf"))
        if logger:
            logger.info(f"Trovati {len(pdf_files)} file PDF")
            for pdf in pdf_files:
                stat = pdf.stat()
                logger.info(f"  PDF: {pdf.name} - Dimensione: {stat.st_size} bytes - Modificato: {datetime.fromtimestamp(stat.st_mtime)}")
    except Exception as e:
        if logger:
            logger.error(f"Errore nella ricerca PDF: {e}")
        return None
    
    if not pdf_files:
        if logger:
            logger.error(f"Nessun file PDF trovato in: {most_recent_dir}")
        return None
    
    # Ordina per data di modifica (decrescente)
    try:
        pdf_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
        latest_pdf = pdf_files[0]
        
        if logger:
            stat = latest_pdf.stat()
            logger.info(f"PDF più recente selezionato: {latest_pdf}")
            logger.info(f"  Dimensione: {stat.st_size} bytes")
            logger.info(f"  Modificato: {datetime.fromtimestamp(stat.st_mtime)}")
            logger.info(f"  Percorso completo: {os.path.abspath(latest_pdf)}")
        
        return latest_pdf
    except Exception as e:
        if logger:
            logger.error(f"Errore nell'ordinamento PDF: {e}")
        return None


def print_pdf_windows(pdf_path, logger=None):
    """Stampa il PDF usando la stampante predefinita di Windows."""
    if logger:
        logger.info(f"Tentativo di stampa: {pdf_path}")
    
    # Verifica che il file esista
    if not os.path.exists(pdf_path):
        if logger:
            logger.error(f"File PDF non trovato: {pdf_path}")
        return False
    
    # Verifica che sia un file PDF
    if not str(pdf_path).lower().endswith('.pdf'):
        if logger:
            logger.error(f"File non è un PDF: {pdf_path}")
        return False
    
    try:
        if logger:
            logger.info(f"Invio alla stampante: {pdf_path}")
        
        # Usa os.startfile per Windows
        os.startfile(str(pdf_path), "print")
        
        # Attendi qualche secondo per evitare che il processo termini troppo presto
        sleep(3)
        
        if logger:
            logger.info("Stampa inviata con successo alla coda di stampa")
        return True
        
    except Exception as e:
        if logger:
            logger.error(f"ERRORE durante la stampa: {e}")
        return False


def main():
    # Setup logging
    logger = setup_logging()
    logger.info("=== AVVIO STAMPA PDF ===")
    logger.info(f"Directory di lavoro: {os.getcwd()}")
    
    try:
        # Cerca il PDF più recente nella cartella DDMMYY più recente
        latest_pdf = find_latest_pdf_in_recent_folder(logger=logger)
        
        if not latest_pdf:
            logger.error("Nessun file PDF trovato nelle cartelle DDMMYY.")
            sys.exit(1)

        # Stampa il PDF
        success = print_pdf_windows(latest_pdf, logger=logger)
        
        if success:
            logger.info("=== STAMPA COMPLETATA CON SUCCESSO ===")
        else:
            logger.error("=== ERRORE DURANTE LA STAMPA ===")
            sys.exit(2)
            
    except Exception as e:
        logger.error(f"Errore imprevisto nel main: {e}")
        sys.exit(3)


if __name__ == "__main__":
    main()
