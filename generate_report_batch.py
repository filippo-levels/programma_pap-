#!/usr/bin/env python3
"""
generate_report_batch.py

Script per generare report PDF da file CSV BATCH.
Target: Python >= 3.9
Compatible: Python 3.11

Uso:
python generate_report_batch.py [--csv <path_csv>] [--out <path_pdf>] [--logo <path_logo>] [--limit-rows N] [--dry-run]
"""

import sys
import os
import argparse
import glob
from pathlib import Path
import tempfile
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas non trovato. Installare con: pip install pandas", file=sys.stderr)
    sys.exit(1)

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
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


def find_csv_batch(directory="."):
    """Trova il file CSV BATCH più recente nella directory."""
    pattern = os.path.join(directory, "*batch*.csv")
    files = glob.glob(pattern, recursive=False)
    
    # Cerca anche maiuscolo
    pattern_upper = os.path.join(directory, "*BATCH*.csv")
    files.extend(glob.glob(pattern_upper, recursive=False))
    
    if not files:
        return None
    
    # Seleziona il più recente per mtime, poi alfabeticamente
    files.sort(key=lambda x: (os.path.getmtime(x), x), reverse=True)
    return files[0]


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


def load_batch_data(filepath, limit_rows=None):
    """Carica i dati BATCH dal CSV."""
    print(f"Caricamento dati da: {filepath}")
    
    # Trova la riga header
    header_row = find_header_row(filepath)
    print(f"Header trovato alla riga: {header_row + 1}")
    
    # Colonne richieste (ignorando QF)
    required_cols = ['Date', 'Time', 'USER', 'TEMP_AIR_IN', 'TEMP_PRODUCT_1', 'TEMP_PRODUCT_2', 'TEMP_PRODUCT_3']
    
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
        
        # Pulizia nomi colonne
        df.columns = df.columns.str.strip().str.replace('"', '').str.replace("'", "")
        
        # Rimuovi colonne QF (Quality Flag)
        qf_cols = [col for col in df.columns if col == 'QF' or col.endswith('_QF')]
        if qf_cols:
            print(f"Rimozione colonne QF: {qf_cols}")
            df = df.drop(columns=qf_cols)
        
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
        
        # Conversione e arrotondamento temperature
        temp_cols = ['TEMP_AIR_IN', 'TEMP_PRODUCT_1', 'TEMP_PRODUCT_2', 'TEMP_PRODUCT_3']
        for col in temp_cols:
            if col in df.columns:
                try:
                    # Sostituisci virgola con punto per separatore decimale
                    df[col] = df[col].astype(str).str.replace(',', '.')
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                    # Arrotonda a 1 cifra decimale
                    df[col] = df[col].round(1)
                except Exception as e:
                    print(f"WARNING: Errore conversione {col}: {e}", file=sys.stderr)
        
        # Crea datetime per sort e calcoli (prima della conversione formato)
        if 'Date' in df.columns and 'Time' in df.columns:
            try:
                df['DateTime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'], errors='coerce')
                # Ordina per datetime
                df = df.sort_values('DateTime')
            except Exception as e:
                print(f"WARNING: Errore parsing datetime: {e}", file=sys.stderr)
        
        # Converti formato date DOPO aver creato DateTime
        if 'Date' in df.columns:
            df['Date'] = df['Date'].apply(convert_date_format)
        
        print(f"Dati caricati: {len(df)} righe, {len(df.columns)} colonne")
        return df, missing_cols
        
    except Exception as e:
        print(f"ERRORE durante caricamento CSV: {e}", file=sys.stderr)
        return None, []


def calculate_period(df):
    """Calcola il periodo inizio-fine dai dati in inglese."""
    if len(df) == 0:
        return "No data"
    
    try:
        if 'DateTime' in df.columns:
            # Usa datetime se disponibile
            valid_dates = df['DateTime'].dropna()
            if len(valid_dates) > 0:
                start_date = valid_dates.min()
                end_date = valid_dates.max()
                return f"Starting time: {start_date.strftime('%d/%m/%Y %H:%M')} - Ending time: {end_date.strftime('%d/%m/%Y %H:%M')}"
        
        # Fallback: usa Date + Time come stringhe
        if 'Date' in df.columns:
            first_date = df.iloc[0]['Date'] if not pd.isna(df.iloc[0]['Date']) else "N/A"
            last_date = df.iloc[-1]['Date'] if not pd.isna(df.iloc[-1]['Date']) else "N/A"
            
            if 'Time' in df.columns:
                first_time = df.iloc[0]['Time'] if not pd.isna(df.iloc[0]['Time']) else ""
                last_time = df.iloc[-1]['Time'] if not pd.isna(df.iloc[-1]['Time']) else ""
                return f"Starting time: {first_date} {first_time} - Ending time: {last_date} {last_time}"
            else:
                return f"Starting time: {first_date} - Ending time: {last_date}"
                
    except Exception as e:
        print(f"WARNING: Errore calcolo periodo: {e}", file=sys.stderr)
    
    return "Period not determinable"


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
        spaceAfter=6,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    subtitle_style = ParagraphStyle(
        'Subtitle',
        parent=styles['Normal'],
        fontSize=11,
        spaceAfter=12,
        alignment=TA_CENTER,
        fontName='Helvetica-Oblique'
    )
    
    story = []
    
    # Header con logo e titolo
    logo_cell = ""
    if logo_path and os.path.exists(logo_path):
        try:
            logo = Image(logo_path, width=25*mm, height=25*mm, kind='proportional')
            logo_cell = logo
        except Exception as e:
            print(f"WARNING: Impossibile caricare logo {logo_path}: {e}", file=sys.stderr)
            logo_cell = ""
    elif logo_path:
        print(f"WARNING: Logo non trovato: {logo_path}", file=sys.stderr)
        logo_cell = ""
    
    # Titolo
    title = f"Batch Report - {Path(source_filename).stem}"
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
    
    # Sottotitolo con periodo
    period = calculate_period(df)
    subtitle_para = Paragraph(f"{period}", subtitle_style)
    story.append(subtitle_para)
    
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
        
        header_style = ParagraphStyle(
            'HeaderStyle',
            parent=styles['Normal'],
            fontSize=9,
            leading=10,
            wordWrap='LTR',
            fontName='Helvetica-Bold'
        )
        
        # Prepara dati tabella (esclude DateTime se presente)
        display_cols = [col for col in df.columns if col != 'DateTime']
        
        # Formatta gli header per migliorare la leggibilità (in inglese)
        formatted_headers = []
        for col in display_cols:
            # Inserisci spazi prima delle maiuscole per le colonne temperatura
            if col.startswith('TEMP_'):
                parts = col.split('_')
                if len(parts) > 1:
                    # Formatta "TEMP_AIR_IN" come "Temp.<br/>Air<br/>Inlet"
                    if parts[1] == 'AIR' and len(parts) > 2 and parts[2] == 'IN':
                        header_text = f"Temp.<br/>Air<br/>Inlet"
                    # Formatta "TEMP_PRODUCT_1" come "Temp.<br/>Product<br/>1"
                    elif parts[1] == 'PRODUCT' and len(parts) > 2:
                        header_text = f"Temp.<br/>Product<br/>{parts[2]}"
                    else:
                        # Fallback per altri formati
                        header_text = "<br/>".join(parts)
                else:
                    header_text = col
                formatted_headers.append(Paragraph(header_text, header_style))
            else:
                formatted_headers.append(Paragraph(col, header_style))
        
        table_data = [formatted_headers]
        
        for _, row in df.iterrows():
            row_data = []
            for col in display_cols:
                cell_value = str(row[col]) if pd.notna(row[col]) else ""
                # Usa Paragraph per tutte le celle per supportare word-wrap
                if len(cell_value) > 10:
                    # Cerca di dividere testi lunghi inserendo <br/> ogni ~15 caratteri
                    if len(cell_value) > 30:
                        # Trova spazi dove inserire break
                        words = cell_value.split()
                        lines = []
                        current_line = ""
                        
                        for word in words:
                            if len(current_line) + len(word) > 15:
                                lines.append(current_line)
                                current_line = word
                            else:
                                if current_line:
                                    current_line += " " + word
                                else:
                                    current_line = word
                        
                        if current_line:
                            lines.append(current_line)
                        
                        cell_value = "<br/>".join(lines)
                    
                    row_data.append(Paragraph(cell_value, cell_style))
                else:
                    row_data.append(cell_value)
            table_data.append(row_data)
        
        # Calcola larghezze colonne (allargate)
        page_width = A4[0] - 3*cm  # Margini ridotti per più spazio
        if len(display_cols) == 7:  # Date, Time, USER, 4 temperature
            col_widths = [2*cm, 2*cm, 3.5*cm, 2*cm, 2*cm, 2*cm, 2*cm]
        else:
            col_widths = [page_width/len(display_cols)] * len(display_cols)
        
        # Crea tabella
        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        
        # Stile tabella
        table.setStyle(TableStyle([
            # Header
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Centra gli header
            ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('TOPPADDING', (0, 0), (-1, 0), 6),
            
            # Data rows
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
            ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            
            # Borders
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
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
        log_file = os.path.join(os.path.dirname(sys.executable), 'batch_report.log')
        try:
            sys.stdout = open(log_file, 'w', encoding='utf-8')
            sys.stderr = sys.stdout
        except:
            pass  # If can't create log, continue without redirection
    
    parser = argparse.ArgumentParser(description='Genera report PDF da file CSV BATCH')
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
        csv_path = find_csv_batch()
        if not csv_path:
            print("ERRORE: Nessun file CSV BATCH trovato nella directory corrente", file=sys.stderr)
            sys.exit(1)
    
    print(f"File CSV selezionato: {csv_path}")
    
    # Carica dati
    df, missing_cols = load_batch_data(csv_path, args.limit_rows)
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
        print(f"Periodo: {calculate_period(df)}")
        print("Anteprima dati:")
        print(df.head())
        return
    
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
    
    # Output path
    if args.out:
        output_path = args.out
    else:
        base_name = Path(csv_path).stem
        output_path = f"{base_name}_report.pdf"
    
    # Genera PDF report
    success = create_pdf_report(df, output_path, csv_path, logo_path, missing_cols)
    
    if success:
        print(f"✅ PDF report generato: {output_path}")
        sys.exit(0)
    else:
        print("❌ Errore generazione PDF report")
        sys.exit(1)


if __name__ == "__main__":
    main() 