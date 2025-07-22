#!/usr/bin/env python3
"""
report_utils.py

Utilità comuni per i generatori di report PDF.
Contiene funzioni condivise dai tre script: generate_report_alarm.py, generate_report_batch.py, generate_report_operlog.py
"""

import sys
import os
from datetime import datetime

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas non trovato. Installare con: pip install pandas", file=sys.stderr)
    sys.exit(1)

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import Table, TableStyle, Paragraph, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import cm, mm
    from reportlab.lib.enums import TA_CENTER
except ImportError:
    print("ERROR: reportlab non trovato. Installare con: pip install reportlab", file=sys.stderr)
    sys.exit(1)


def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)


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


def clean_dataframe_columns(df, required_cols):
    """Pulizia comune delle colonne del dataframe."""
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
    
    return missing_cols


def clean_dataframe_data(df):
    """Pulizia comune dei dati del dataframe."""
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str).str.strip().str.replace('"', '').str.replace("'", "")
            # Sostituisci 'nan' string con stringa vuota
            df[col] = df[col].replace('nan', '')
            df[col] = df[col].replace("'-", '')  # Rimuovi segnaposto "'-"


def create_logo_header(logo_path, title, title_style):
    """Crea l'header con logo e titolo. Ritorna la tabella header o il titolo semplice."""
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
    
    title_para = Paragraph(title, title_style)
    
    if logo_cell:
        header_table = Table([[logo_cell, title_para]], colWidths=[3*cm, None])
        header_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (0, 0), 'LEFT'),
            ('ALIGN', (1, 0), (1, 0), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        return header_table
    else:
        return title_para


def get_common_styles():
    """Ritorna gli stili comuni per i PDF."""
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=14,
        spaceAfter=12,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
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
    
    return styles, title_style, cell_style, header_style


def add_page_number(canvas):
    """Aggiunge numerazione pagine corretta (non più hardcoded /1)."""
    page_num = canvas.getPageNumber()
    text = f"Page {page_num}"
    canvas.drawRightString(A4[0] - 2*cm, 1*cm, text)


def create_missing_columns_note(missing_cols, styles):
    """Crea la nota per le colonne mancanti."""
    if missing_cols:
        return Paragraph(
            f"<i>Note: Colonne mancanti nel file sorgente: {', '.join(missing_cols)}</i>",
            styles['Normal']
        )
    return None


def get_common_table_style():
    """Ritorna lo stile comune per le tabelle."""
    return TableStyle([
        # Header
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        
        # Data rows
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
        
        # Borders
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ])


def setup_logging_for_pyinstaller(script_name):
    """Setup logging per PyInstaller (windowed mode)."""
    if getattr(sys, 'frozen', False):
        log_file = os.path.join(os.path.dirname(sys.executable), f'{script_name}.log')
        try:
            sys.stdout = open(log_file, 'w', encoding='utf-8')
            sys.stderr = sys.stdout
        except:
            pass  # If can't create log, continue without redirection


def get_logo_path(args_logo):
    """Determina il path del logo, gestendo PyInstaller bundled resources."""
    if args_logo:
        return args_logo
    else:
        # Try bundled logo first (for PyInstaller), then fallback to local
        bundled_logo = get_resource_path("logo.png")
        if os.path.exists(bundled_logo):
            return bundled_logo
        else:
            return "logo.png" 