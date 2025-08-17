#!/usr/bin/env python3
"""
generate_report_operlog.py

Script per generare report PDF da file CSV OPERLOG.
Target: Python >= 3.9
Compatible: Python 3.11

Copyright © 2024 Filippo Caliò
Version: 1.0.0

Posizionamento: Script a livello delle cartelle DDMMYY contenenti i CSV OPERLOG.

Uso:
python generate_report_operlog.py [--csv <path_csv>] [--out <path_pdf>] [--logo <path_logo>] [--limit-rows N] [--dry-run]
"""

import sys
import os
import argparse
import glob
import re
from datetime import datetime

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


def find_csv_operlog(base_directory="."):
    """Trova il file CSV OPERLOG più recente nella cartella DDMMYY più recente."""
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
    
    # Cerca CSV OPERLOG nella cartella
    pattern = os.path.join(most_recent_dir, "*operlog*.csv")
    files = glob.glob(pattern, recursive=False)
    
    # Cerca anche maiuscolo
    pattern_upper = os.path.join(most_recent_dir, "*OPERLOG*.csv")
    files.extend(glob.glob(pattern_upper, recursive=False))
    
    if not files:
        print(f"Nessun file CSV OPERLOG trovato in: {most_recent_dir}", file=sys.stderr)
        return None
    
    # Seleziona il più recente per mtime
    files.sort(key=os.path.getmtime, reverse=True)
    return files[0]





def load_operlog_data(filepath, limit_rows=None):
    """Carica i dati OPERLOG dal CSV."""
    print(f"Caricamento dati da: {filepath}")
    
    # Trova la riga header
    header_row = find_header_row(filepath)
    print(f"Header trovato alla riga: {header_row + 1}")
    
    # Colonne richieste
    required_cols = ['Date', 'Time', 'User', 'Object_Action', 'Trigger', 'PreviousValue', 'ChangedValue']
    
    try:
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
        
        # Logica speciale per Object_Action solo se mancante
        if 'Object_Action' in missing_cols:
            if 'Screen' in df.columns and 'Object_Action' in df.columns:
                # Se entrambi esistono, combina
                df['Object_Action'] = df['Screen'].astype(str) + ':' + df['Object_Action'].astype(str)
                missing_cols.remove('Object_Action')
            elif 'Screen' in df.columns:
                # Rinomina Screen in Object_Action
                df.rename(columns={'Screen': 'Object_Action'}, inplace=True)
                missing_cols.remove('Object_Action')
        
        if missing_cols:
            print(f"WARNING: Colonne mancanti: {missing_cols}", file=sys.stderr)
        
        # Seleziona solo le colonne richieste (quelle disponibili)
        available_cols = [col for col in required_cols if col in df.columns]
        df = df[available_cols].copy()
        
        # Pulizia dati
        clean_dataframe_data(df)
        
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
    title = os.path.splitext(os.path.basename(source_filename))[0]
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
        # Formatta gli header per migliorare la leggibilità
        formatted_headers = []
        for col in df.columns:
            if col == 'PreviousValue':
                header_text = "Previous<br/>Value"
                formatted_headers.append(Paragraph(header_text, header_style))
            elif col == 'ChangedValue':
                header_text = "Changed<br/>Value"
                formatted_headers.append(Paragraph(header_text, header_style))
            elif col == 'Object_Action':
                header_text = "Object<br/>Action"
                formatted_headers.append(Paragraph(header_text, header_style))
            else:
                formatted_headers.append(Paragraph(col, header_style))
        
        table_data = [formatted_headers]
        
        for _, row in df.iterrows():
            row_data = []
            for col in df.columns:
                cell_value = str(row[col]) if pd.notna(row[col]) else ""
                # Usa Paragraph per celle lunghe con word wrapping migliorato
                if col in ['Object_Action', 'PreviousValue', 'ChangedValue', 'User', 'Trigger'] and len(cell_value) > 15:
                    # Migliora il word wrapping per testi lunghi
                    if len(cell_value) > 40:
                        # Inserisci spazi per facilitare il wrapping
                        words = cell_value.split()
                        wrapped_lines = []
                        current_line = ""
                        
                        for word in words:
                            if len(current_line) + len(word) > 25:
                                if current_line:
                                    wrapped_lines.append(current_line.strip())
                                current_line = word
                            else:
                                if current_line:
                                    current_line += " " + word
                                else:
                                    current_line = word
                        
                        if current_line:
                            wrapped_lines.append(current_line.strip())
                        
                        cell_value = "<br/>".join(wrapped_lines)
                    
                    row_data.append(Paragraph(cell_value, cell_style))
                else:
                    row_data.append(cell_value)
            table_data.append(row_data)
        
        # Calcola larghezze colonne (gestisce variabili colonne)
        page_width = A4[0] - 4*cm  # Margini - Aumentata larghezza disponibile
        if len(df.columns) == 7:  # Tutte le colonne richieste
            # Date e Time mantengono dimensione fissa, User allargato, altre colonne più larghe
            col_widths = [1.6*cm, 1.6*cm, 2.2*cm, 3.2*cm, 1.8*cm, 2.8*cm, 2.8*cm]
        else:
            col_widths = [page_width/len(df.columns)] * len(df.columns)
        
        # Crea tabella
        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        
        # Applica stile comune e personalizzazioni specifiche
        table_style = get_common_table_style()
        table_style.add('FONTSIZE', (0, 0), (-1, 0), 9)  # Header font size
        table_style.add('FONTSIZE', (0, 1), (-1, -1), 8)  # Data font size
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
    setup_logging_for_pyinstaller('operlog_report')
    
    parser = argparse.ArgumentParser(description='Genera report PDF da file CSV OPERLOG')
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
        csv_path = find_csv_operlog()
        if not csv_path:
            print("ERRORE: Nessun file CSV OPERLOG trovato nelle cartelle DDMMYY", file=sys.stderr)
            sys.exit(1)
    
    print(f"File CSV selezionato: {csv_path}")
    
    # Carica dati
    df, missing_cols = load_operlog_data(csv_path, args.limit_rows)
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
        # Usa os.path per gestire correttamente i separatori di percorso su Windows
        base_name = os.path.splitext(os.path.basename(csv_path))[0]
        csv_dir = os.path.dirname(csv_path)
        output_path = os.path.join(csv_dir, f"{base_name}_report.pdf")
    
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