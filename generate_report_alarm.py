#!/usr/bin/env python3
"""
generate_report_alarm.py

Script per generare report PDF da file CSV ALARM.
Target: Python >= 3.9
Compatible: Python 3.11

Copyright © 2024 Filippo Caliò
Version: 1.0.0

Uso:
python generate_report_alarm.py [--csv <path_csv>] [out <path_pdf>] [--logo <path_logo>] [--limit-rows N] [--dry-run]
"""

import sys
import os
import argparse
import glob
from pathlib import Path
import warnings

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas non trovato. Installare con: pip install pandas", file=sys.stderr)
    sys.exit(1)

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, Paragraph, Spacer
    from reportlab.lib.units import cm
except ImportError:
    print("ERROR: reportlab non trovato. Installare con: pip install reportlab", file=sys.stderr)
    sys.exit(1)

# Importa utilità comuni
from report_utils import (
    convert_date_format, find_header_row, clean_dataframe_columns, 
    clean_dataframe_data, create_logo_header, get_common_styles, 
    add_page_number, create_missing_columns_note, get_common_table_style,
    setup_logging_for_pyinstaller, get_logo_path
)

# Importazioni opzionali
try:
    import pyarrow
    HAS_PYARROW = True
except ImportError:
    HAS_PYARROW = False


def find_csv_alarm(directory="."):
    """Trova il file CSV ALARM più recente nella directory."""
    pattern = os.path.join(directory, "*alarm*.csv")
    files = glob.glob(pattern, recursive=False)
    
    # Cerca anche maiuscolo
    pattern_upper = os.path.join(directory, "*ALARM*.csv")
    files.extend(glob.glob(pattern_upper, recursive=False))
    
    if not files:
        return None
    
    # Seleziona il più recente per mtime, poi alfabeticamente
    files.sort(key=lambda x: (os.path.getmtime(x), x), reverse=True)
    return files[0]





def load_alarm_data(filepath, limit_rows=None):
    """Carica i dati ALARM dal CSV."""
    print(f"Caricamento dati da: {filepath}")
    
    # Trova la riga header
    header_row = find_header_row(filepath)
    print(f"Header trovato alla riga: {header_row + 1}")
    
    # Colonne richieste
    required_cols = ['Date', 'Time', 'Alarm Message', 'Alarm Status']
    
    try:
        # Prova con engine pyarrow se disponibile
        engine = 'pyarrow' if HAS_PYARROW else 'python'
        
        # Carica con separatore auto-detect
        df = pd.read_csv(
            filepath,
            skiprows=header_row,
            sep=None,  # Auto-detect separator
            engine='python',  # Necessario per sep=None
            quotechar='"',
            skipinitialspace=True,
            nrows=limit_rows
        )
        
        print(f"Colonne trovate: {list(df.columns)}")
        
        # Pulizia nomi colonne e controllo colonne richieste
        missing_cols = clean_dataframe_columns(df, required_cols)
        
        if missing_cols:
            print(f"WARNING: Colonne mancanti: {missing_cols}", file=sys.stderr)
        
        # Seleziona solo le colonne richieste (quelle disponibili)
        available_cols = [col for col in required_cols if col in df.columns]
        df = df[available_cols].copy()
        
        # Pulizia dati
        clean_dataframe_data(df)
        
        # Pulisci Alarm Status se presente
        if 'Alarm Status' in df.columns:
            df['Alarm Status'] = df['Alarm Status'].str.title()  # Active, Return, Ack
        
        # Converti formato date
        if 'Date' in df.columns:
            df['Date'] = df['Date'].apply(convert_date_format)
        
        print(f"Dati caricati: {len(df)} righe, {len(df.columns)} colonne")
        return df, missing_cols
        
    except Exception as e:
        print(f"ERRORE durante caricamento CSV: {e}", file=sys.stderr)
        return None, []


def create_pdf_report(df, output_path, source_filename, logo_path=None, missing_cols=None):
    """Genera il report PDF."""
    print(f"Generazione PDF: {output_path}")
    
    # Setup documento
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=2*cm,
        rightMargin=2*cm,
        topMargin=1.5*cm,
        bottomMargin=1.5*cm
    )
    
    # Stili comuni
    styles, title_style, cell_style, header_style = get_common_styles()
    story = []
    
    # Header con logo e titolo
    title = Path(source_filename).stem
    header = create_logo_header(logo_path, title, title_style)
    story.append(header)
    story.append(Spacer(1, 12))
    
    # Note su colonne mancanti
    missing_note = create_missing_columns_note(missing_cols, styles)
    if missing_note:
        story.append(missing_note)
        story.append(Spacer(1, 6))
    
    # Tabella dati
    if len(df) == 0:
        no_data = Paragraph("Nessun dato disponibile", styles['Normal'])
        story.append(no_data)
    else:
        # Prepara dati tabella con Paragraph per word wrap
        headers = list(df.columns)
        table_data = [headers]  # Header senza Paragraph per mantenere il bold
        
        for _, row in df.iterrows():
            row_data = []
            for col in headers:
                cell_value = str(row[col]) if pd.notna(row[col]) else ""
                # Usa Paragraph per celle lunghe (Alarm Message)
                if col == 'Alarm Message' and len(cell_value) > 30:
                    row_data.append(Paragraph(cell_value, cell_style))
                else:
                    row_data.append(cell_value)
            table_data.append(row_data)
        
        # Calcola larghezze colonne
        page_width = A4[0] - 4*cm  # Margini
        if len(headers) == 4:  # Date, Time, Alarm Message, Alarm Status
            # Allarga colonna Alarm Status
            col_widths = [1.8*cm, 1.8*cm, page_width-6.8*cm, 3.2*cm]
        else:
            col_widths = [page_width/len(headers)] * len(headers)
        
        # Crea tabella
        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        
        # Applica stile comune e personalizzazioni specifiche
        table_style = get_common_table_style()
        table_style.add('FONTSIZE', (0, 0), (-1, 0), 10)  # Header font size
        table_style.add('FONTSIZE', (0, 1), (-1, -1), 9)  # Data font size
        table.setStyle(table_style)
        
        story.append(table)
    
    # Genera PDF
    try:
        doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
        print(f"PDF generato con successo: {output_path}")
        return True
    except Exception as e:
        print(f"ERRORE durante generazione PDF: {e}", file=sys.stderr)
        return False


def main():
    # Setup logging per PyInstaller
    setup_logging_for_pyinstaller('alarm_report')
    
    parser = argparse.ArgumentParser(description='Genera report PDF da file CSV ALARM')
    parser.add_argument('--csv', help='Path del file CSV (auto-detect se omesso)')
    parser.add_argument('--out', help='Path del PDF output (auto-generato se omesso)')
    parser.add_argument('--logo', help='Path del logo (default: logo.png nella directory corrente)')
    parser.add_argument('--limit-rows', type=int, help='Limita numero righe per debug')
    parser.add_argument('--dry-run', action='store_true', help='Mostra info senza generare PDF')
    
    args = parser.parse_args()
    
    # Trova CSV se non specificato
    if args.csv:
        csv_path = args.csv
        if not os.path.exists(csv_path):
            print(f"ERRORE: File CSV non trovato: {csv_path}", file=sys.stderr)
            sys.exit(1)
    else:
        csv_path = find_csv_alarm()
        if not csv_path:
            print("ERRORE: Nessun file CSV ALARM trovato nella directory corrente", file=sys.stderr)
            sys.exit(1)
    
    print(f"File CSV selezionato: {csv_path}")
    
    # Carica dati
    df, missing_cols = load_alarm_data(csv_path, args.limit_rows)
    if df is None:
        print("ERRORE: Impossibile caricare i dati", file=sys.stderr)
        sys.exit(1)
    
    if args.dry_run:
        print("=== DRY RUN ===")
        print(f"File CSV: {csv_path}")
        print(f"Righe caricate: {len(df)}")
        print(f"Colonne: {list(df.columns)}")
        if missing_cols:
            print(f"Colonne mancanti: {missing_cols}")
        print("Anteprima dati:")
        print(df.head())
        return
    
    # Determina output path
    if args.out:
        output_path = args.out
    else:
        base_name = Path(csv_path).stem
        output_path = f"{base_name}_report.pdf"
    
    # Logo path
    logo_path = get_logo_path(args.logo)
    
    # Genera PDF
    success = create_pdf_report(df, output_path, csv_path, logo_path, missing_cols)
    
    if success:
        print(f"Report generato: {output_path}")
        sys.exit(0)
    else:
        sys.exit(1)


if __name__ == "__main__":
    main() 