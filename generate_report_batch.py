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

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas non trovato. Installare con: pip install pandas", file=sys.stderr)
    sys.exit(1)

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, Paragraph, Spacer, PageBreak, Image
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_CENTER
except ImportError:
    print("ERROR: reportlab non trovato. Installare con: pip install reportlab", file=sys.stderr)
    sys.exit(1)

try:
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
    import io
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False
    print("WARNING: matplotlib non trovato. I grafici non saranno disponibili.", file=sys.stderr)

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
        
        # Rimuovi colonne QF (Quality Flag) prima della pulizia
        qf_cols = [col for col in df.columns if col == 'QF' or col.endswith('_QF')]
        if qf_cols:
            print(f"Rimozione colonne QF: {qf_cols}")
            df = df.drop(columns=qf_cols)
        
        # Pulizia nomi colonne e controllo colonne richieste
        missing_cols = clean_dataframe_columns(df, required_cols)
        
        if missing_cols:
            print(f"WARNING: Colonne mancanti: {missing_cols}", file=sys.stderr)
        
        # Seleziona solo le colonne richieste (quelle disponibili)
        available_cols = [col for col in required_cols if col in df.columns]
        df = df[available_cols].copy()
        
        # Pulizia dati
        clean_dataframe_data(df)
        
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


def create_temperature_chart(df, output_path=None):
    """Crea un grafico dell'andamento delle temperature nel tempo."""
    if not HAS_MATPLOTLIB:
        print("WARNING: matplotlib non disponibile, impossibile creare grafico", file=sys.stderr)
        return None
    
    if len(df) == 0 or 'DateTime' not in df.columns:
        print("WARNING: Dati insufficienti per creare grafico", file=sys.stderr)
        return None
    
    try:
        # Configura il grafico
        plt.style.use('default')
        fig, ax = plt.subplots(figsize=(16, 10))
        
        # Colori specificati
        colors = {
            'TEMP_AIR_IN': 'green',
            'TEMP_AIR_OUT': 'orange',  # Se presente
            'TEMP_PRODUCT_3': 'blue',
            'TEMP_PRODUCT_2': 'yellow',
            'TEMP_PRODUCT_1': 'red'
        }
        
        # Nomi in inglese per la legenda
        legend_names = {
            'TEMP_AIR_IN': 'Air Inlet Temperature',
            'TEMP_AIR_OUT': 'Air Outlet Temperature',
            'TEMP_PRODUCT_3': 'Product 3 Temperature',
            'TEMP_PRODUCT_2': 'Product 2 Temperature',
            'TEMP_PRODUCT_1': 'Product 1 Temperature'
        }
        
        # Filtra dati validi
        valid_data = df.dropna(subset=['DateTime'])
        
        if len(valid_data) == 0:
            print("WARNING: Nessun dato valido per il grafico", file=sys.stderr)
            return None
        
        # Traccia ogni temperatura disponibile
        for col in ['TEMP_AIR_IN', 'TEMP_PRODUCT_1', 'TEMP_PRODUCT_2', 'TEMP_PRODUCT_3']:
            if col in valid_data.columns:
                # Rimuovi valori NaN per questa colonna
                temp_data = valid_data.dropna(subset=[col])
                if len(temp_data) > 0:
                    ax.plot(temp_data['DateTime'], temp_data[col], 
                           color=colors.get(col, 'black'), 
                           linewidth=2, 
                           label=legend_names.get(col, col),
                           marker='o', markersize=3)
        
        # Personalizza il grafico
        ax.set_xlabel('Time', fontsize=12, fontweight='bold')
        ax.set_ylabel('Temperature (°C)', fontsize=12, fontweight='bold')
        ax.set_title('Temperature Trend Over Time', fontsize=14, fontweight='bold')
        
        # Formatta l'asse X per mostrare le date/ore
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
        ax.xaxis.set_major_locator(mdates.MinuteLocator(interval=max(1, len(valid_data)//10)))
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
        
        # Aggiungi griglia
        ax.grid(True, alpha=0.3)
        
        # Aggiungi legenda
        if ax.get_legend_handles_labels()[0]:
            ax.legend(loc='upper right', fontsize=10)
        
        # Ottimizza layout
        plt.tight_layout()
        
        # Salva il grafico in un buffer
        buffer = io.BytesIO()
        plt.savefig(buffer, format='png', dpi=300, bbox_inches='tight')
        buffer.seek(0)
        plt.close()
        
        return buffer
        
    except Exception as e:
        print(f"ERRORE durante creazione grafico: {e}", file=sys.stderr)
        return None


def create_pdf_report(df, output_path, source_filename, logo_path=None, missing_cols=None):
    """Genera il report PDF."""
    print(f"Generazione PDF: {output_path}")
    
    # Setup documento principale (portrait)
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=2*cm,
        rightMargin=2*cm,
        topMargin=1.5*cm,
        bottomMargin=1.5*cm
    )
    
    # Stili comuni e sottotitolo specifico per batch
    styles, title_style, cell_style, header_style = get_common_styles()
    
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
    title = Path(source_filename).stem
    header = create_logo_header(logo_path, title, title_style)
    story.append(header)
    
    # Sottotitolo con periodo
    period = calculate_period(df)
    subtitle_para = Paragraph(f"{period}", subtitle_style)
    story.append(subtitle_para)
    
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
        
        # Applica stile comune e personalizzazioni specifiche
        table_style = get_common_table_style()
        table_style.add('ALIGN', (0, 0), (-1, 0), 'CENTER')  # Centra gli header
        table_style.add('VALIGN', (0, 0), (-1, 0), 'MIDDLE')
        table_style.add('BOTTOMPADDING', (0, 0), (-1, 0), 6)
        table_style.add('TOPPADDING', (0, 0), (-1, 0), 6)
        table_style.add('FONTSIZE', (0, 1), (-1, -1), 8)  # Data font size
        table_style.add('ALIGN', (0, 1), (-1, -1), 'LEFT')
        table.setStyle(table_style)
        
        story.append(table)
    
    # Genera PDF principale
    try:
        doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
        print(f"PDF principale generato con successo")
        
        # Aggiungi grafico delle temperature se disponibile
        if HAS_MATPLOTLIB and len(df) > 0:
            chart_buffer = create_temperature_chart(df)
            if chart_buffer:
                # Crea documento landscape per il grafico
                chart_output = output_path.replace('.pdf', '_chart.pdf')
                chart_doc = SimpleDocTemplate(
                    chart_output,
                    pagesize=landscape(A4),
                    leftMargin=1.5*cm,
                    rightMargin=1.5*cm,
                    topMargin=1*cm,
                    bottomMargin=1*cm
                )
                
                chart_story = []
                
                # Header con logo e titolo per la pagina del grafico
                chart_title = f"{Path(source_filename).stem} - Temperature Trend"
                chart_header = create_logo_header(logo_path, chart_title, title_style)
                chart_story.append(chart_header)
                chart_story.append(Spacer(1, 12))
                
                # Aggiungi il grafico (dimensioni ottimizzate per landscape)
                chart_image = Image(chart_buffer, width=24*cm, height=15*cm)
                chart_story.append(chart_image)
                
                # Aggiungi note esplicative
                note_style = ParagraphStyle(
                    'Note',
                    parent=styles['Normal'],
                    fontSize=9,
                    spaceAfter=6,
                    alignment=TA_CENTER,
                    fontName='Helvetica-Oblique'
                )
                
                note_text = "Chart shows temperature trends over time. Each line represents a different temperature sensor."
                note_para = Paragraph(note_text, note_style)
                chart_story.append(Spacer(1, 12))
                chart_story.append(note_para)
                
                # Genera PDF del grafico
                chart_doc.build(chart_story, onFirstPage=add_page_number, onLaterPages=add_page_number)
                print(f"PDF grafico generato: {chart_output}")
        
        return True
    except Exception as e:
        print(f"ERRORE durante generazione PDF: {e}", file=sys.stderr)
        return False


def main():
    # Setup logging per PyInstaller
    setup_logging_for_pyinstaller('batch_report')
    
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
    
    # Logo path
    logo_path = get_logo_path(args.logo)
    
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