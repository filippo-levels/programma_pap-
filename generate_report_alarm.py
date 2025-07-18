#!/usr/bin/env python3
"""
generate_report_alarm.py

Script per generare report PDF da file CSV ALARM.
Target: Python >= 3.9
Compatible: Python 3.11

Uso:
python generate_report_alarm.py [--csv <path_csv>] [--out <path_pdf>] [--logo <path_logo>] [--limit-rows N] [--dry-run]
"""

import sys
import os
import argparse
import glob
from pathlib import Path
import warnings
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas non trovato. Installare con: pip install pandas", file=sys.stderr)
    sys.exit(1)

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import cm, mm
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
except ImportError:
    print("ERROR: reportlab non trovato. Installare con: pip install reportlab", file=sys.stderr)
    sys.exit(1)

# Importazioni opzionali
try:
    import pyarrow
    HAS_PYARROW = True
except ImportError:
    HAS_PYARROW = False


def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)


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


def find_header_row(filepath, max_scan_rows=10):
    """Trova la riga header che inizia con 'Date'."""
    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            for i, line in enumerate(f):
                if i >= max_scan_rows:
                    break
                # Pulisci e controlla se inizia con Date
                clean_line = line.strip().replace('"', '').replace("'", "")
                if clean_line.lower().startswith('date'):
                    return i
    except Exception as e:
        print(f"WARNING: Errore durante ricerca header: {e}", file=sys.stderr)
    
    return 3  # Default: riga 4 (0-indexed)


def convert_date_format(date_str):
    """Converte formato data da MM/DD/YYYY a DD/MM/YY."""
    try:
        if pd.isna(date_str) or date_str == '' or str(date_str).strip() == '':
            return date_str
        
        date_str = str(date_str).strip()
        # Parse MM/DD/YYYY
        dt = datetime.strptime(date_str, '%m/%d/%Y')
        # Return DD/MM/YY
        return dt.strftime('%d/%m/%y')
    except:
        # Se il parsing fallisce, ritorna il valore originale
        return date_str


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
        
        # Pulizia nomi colonne
        df.columns = df.columns.str.strip().str.replace('"', '').str.replace("'", "")
        
        # Controlla colonne richieste
        missing_cols = []
        for col in required_cols:
            if col not in df.columns:
                # Prova match case-insensitive
                matches = [c for c in df.columns if c.lower() == col.lower()]
                if matches:
                    df.rename(columns={matches[0]: col}, inplace=True)
                else:
                    missing_cols.append(col)
        
        if missing_cols:
            print(f"WARNING: Colonne mancanti: {missing_cols}", file=sys.stderr)
        
        # Seleziona solo le colonne richieste (quelle disponibili)
        available_cols = [col for col in required_cols if col in df.columns]
        df = df[available_cols].copy()
        
        # Pulizia dati
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.strip().str.replace('"', '').str.replace("'", "")
                # Sostituisci 'nan' string con stringa vuota
                df[col] = df[col].replace('nan', '')
        
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
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=14,
        spaceAfter=12,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    story = []
    
    # Header con logo e titolo
    header_data = []
    
    # Logo (se presente)
    logo_cell = ""
    if logo_path and os.path.exists(logo_path):
        try:
            logo = Image(logo_path, width=25*mm, height=25*mm, kind='proportional')
            logo_cell = logo
        except Exception as e:
            print(f"WARNING: Impossibile caricare logo {logo_path}: {e}", file=sys.stderr)
            logo_cell = "Logo non disponibile"
    elif logo_path:
        print(f"WARNING: Logo non trovato: {logo_path}", file=sys.stderr)
        logo_cell = ""
    
    # Titolo
    title = Path(source_filename).stem
    title_para = Paragraph(title, title_style)
    
    if logo_cell:
        header_table = Table([[logo_cell, title_para]], colWidths=[3*cm, None])
        header_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (0, 0), 'LEFT'),
            ('ALIGN', (1, 0), (1, 0), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        story.append(header_table)
    else:
        story.append(title_para)
    
    story.append(Spacer(1, 12))
    
    # Note su colonne mancanti
    if missing_cols:
        missing_note = Paragraph(
            f"<i>Note: Colonne mancanti nel file sorgente: {', '.join(missing_cols)}</i>",
            styles['Normal']
        )
        story.append(missing_note)
        story.append(Spacer(1, 6))
    
    # Tabella dati
    if len(df) == 0:
        no_data = Paragraph("Nessun dato disponibile", styles['Normal'])
        story.append(no_data)
    else:
        # Stile per celle con word wrap
        cell_style = ParagraphStyle(
            'CellStyle',
            parent=styles['Normal'],
            fontSize=8,
            leading=9,
            wordWrap='LTR'
        )
        
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
        
        # Stile tabella
        table.setStyle(TableStyle([
            # Header
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            
            # Data rows
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
            
            # Borders
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))
        
        story.append(table)
    
    # Footer con numerazione pagine
    def add_page_number(canvas, doc):
        page_num = canvas.getPageNumber()
        text = f"Page {page_num}/1"
        canvas.drawRightString(A4[0] - 2*cm, 1*cm, text)
    
    # Genera PDF
    try:
        doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
        print(f"PDF generato con successo: {output_path}")
        return True
    except Exception as e:
        print(f"ERRORE durante generazione PDF: {e}", file=sys.stderr)
        return False


def main():
    # Redirect output when running as PyInstaller bundle (windowed mode)
    if getattr(sys, 'frozen', False):
        log_file = os.path.join(os.path.dirname(sys.executable), 'alarm_report.log')
        try:
            sys.stdout = open(log_file, 'w', encoding='utf-8')
            sys.stderr = sys.stdout
        except:
            pass  # If can't create log, continue without redirection
    
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
    
    # Logo path - handle PyInstaller bundled resources
    if args.logo:
        logo_path = args.logo
    else:
        # Try bundled logo first (for PyInstaller), then fallback to local
        bundled_logo = get_resource_path("logo.png")
        if os.path.exists(bundled_logo):
            logo_path = bundled_logo
        else:
            logo_path = "logo.png"
    
    # Genera PDF
    success = create_pdf_report(df, output_path, csv_path, logo_path, missing_cols)
    
    if success:
        print(f"Report generato: {output_path}")
        sys.exit(0)
    else:
        sys.exit(1)


if __name__ == "__main__":
    main() 