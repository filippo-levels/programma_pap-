#!/usr/bin/env python3
"""
run_reports.py

Script per eseguire tutti i generatori di report (batch, alarm, operlog) in un unico comando.
Target: Python >= 3.9
Compatible: Python 3.11

Uso:
python run_reports.py [--data-dir <directory>] [--output-dir <directory>] [--logo <path_logo>] 
                      [--type <batch|alarm|operlog|all>] [--limit-rows N] [--dry-run]
"""

import os
import sys
import argparse
import subprocess
from pathlib import Path
import glob


def find_csv_files(data_dir, report_type):
    """Trova i file CSV del tipo specificato nella directory."""
    pattern_map = {
        'batch': ['BATCH*.csv', 'batch*.csv'],
        'alarm': ['ALARM*.csv', 'alarm*.csv'],
        'operlog': ['OPERLOG*.csv', 'operlog*.csv'],
    }
    
    if report_type not in pattern_map:
        return []
    
    files = []
    for pattern in pattern_map[report_type]:
        files.extend(glob.glob(os.path.join(data_dir, pattern)))
    
    # Filtra per evitare falsi positivi (ad esempio OPERLOG_BATCH per il tipo batch)
    filtered_files = []
    for file in files:
        filename = os.path.basename(file).upper()
        if report_type == 'batch':
            # Esclude file che contengono OPERLOG o ALARM
            if not ('OPERLOG' in filename or 'ALARM' in filename):
                filtered_files.append(file)
        elif report_type == 'alarm':
            # Esclude file che contengono OPERLOG o file che iniziano con BATCH
            if not ('OPERLOG' in filename or filename.startswith('BATCH')):
                filtered_files.append(file)
        elif report_type == 'operlog':
            # Solo file che iniziano con OPERLOG
            if filename.startswith('OPERLOG'):
                filtered_files.append(file)
        else:
            filtered_files.append(file)
    
    # Ordina per data di modifica (più recente prima)
    filtered_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return filtered_files


def run_report_script(script_name, csv_file, output_dir, logo_path=None, limit_rows=None, dry_run=False):
    """Esegue uno script di generazione report con i parametri specificati."""
    cmd = [sys.executable, script_name, '--csv', csv_file]
    
    # Se è specificata una directory di output, crea il nome file di output
    if output_dir:
        base_name = Path(csv_file).stem
        output_file = os.path.join(output_dir, f"{base_name}_report.pdf")
        cmd.extend(['--out', output_file])
    
    # Aggiungi logo se specificato
    if logo_path:
        cmd.extend(['--logo', logo_path])
    
    # Aggiungi limit-rows se specificato
    if limit_rows:
        cmd.extend(['--limit-rows', str(limit_rows)])
    
    # Aggiungi dry-run se specificato
    if dry_run:
        cmd.append('--dry-run')
    
    print(f"Esecuzione: {' '.join(cmd)}")
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print(result.stdout)
        if result.stderr:
            print(f"STDERR: {result.stderr}", file=sys.stderr)
        return True
    except subprocess.CalledProcessError as e:
        print(f"ERRORE durante l'esecuzione di {script_name}: {e}", file=sys.stderr)
        if e.stdout:
            print(f"STDOUT: {e.stdout}")
        if e.stderr:
            print(f"STDERR: {e.stderr}", file=sys.stderr)
        return False


def main():
    parser = argparse.ArgumentParser(description='Esegue tutti i generatori di report in un unico comando')
    parser.add_argument('--data-dir', default='data', help='Directory contenente i file CSV (default: ./data)')
    parser.add_argument('--output-dir', help='Directory per i file PDF di output (default: stessa dei CSV)')
    parser.add_argument('--logo', help='Path del logo (default: logo.png nella directory data)')
    parser.add_argument('--type', choices=['batch', 'alarm', 'operlog', 'all'], default='all',
                       help='Tipo di report da generare (default: all)')
    parser.add_argument('--limit-rows', type=int, help='Limita numero righe per debug')
    parser.add_argument('--dry-run', action='store_true', help='Mostra info senza generare PDF')
    
    args = parser.parse_args()
    
    # Verifica directory dati
    data_dir = args.data_dir
    if not os.path.isdir(data_dir):
        print(f"ERRORE: Directory dati non trovata: {data_dir}", file=sys.stderr)
        sys.exit(1)
    
    # Verifica directory output
    if args.output_dir:
        output_dir = args.output_dir
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
                print(f"Directory output creata: {output_dir}")
            except Exception as e:
                print(f"ERRORE: Impossibile creare directory output: {e}", file=sys.stderr)
                sys.exit(1)
    else:
        output_dir = None
    
    # Logo path
    logo_path = args.logo or os.path.join(data_dir, "logo.png")
    if not os.path.exists(logo_path):
        print(f"AVVISO: Logo non trovato: {logo_path}", file=sys.stderr)
        logo_path = None
    
    # Mappa script per tipo
    script_map = {
        'batch': 'generate_report_batch.py',
        'alarm': 'generate_report_alarm.py',
        'operlog': 'generate_report_operlog.py'
    }
    
    # Determina quali tipi di report generare
    report_types = list(script_map.keys()) if args.type == 'all' else [args.type]
    
    success_count = 0
    failure_count = 0
    
    for report_type in report_types:
        script_name = script_map[report_type]
        
        # Verifica che lo script esista
        if not os.path.exists(script_name):
            print(f"ERRORE: Script {script_name} non trovato", file=sys.stderr)
            continue
        
        # Trova i file CSV per questo tipo
        csv_files = find_csv_files(data_dir, report_type)
        
        if not csv_files:
            print(f"AVVISO: Nessun file CSV trovato per il tipo {report_type}", file=sys.stderr)
            continue
        
        # Usa solo il file più recente per ogni tipo
        csv_file = csv_files[0]
        print(f"\n=== Elaborazione {report_type.upper()} ===")
        print(f"File CSV: {csv_file}")
        
        # Esegui lo script
        success = run_report_script(
            script_name, 
            csv_file, 
            output_dir, 
            logo_path, 
            args.limit_rows, 
            args.dry_run
        )
        
        if success:
            success_count += 1
        else:
            failure_count += 1
    
    # Riepilogo
    print(f"\n=== RIEPILOGO ===")
    print(f"Report completati con successo: {success_count}")
    if failure_count > 0:
        print(f"Report falliti: {failure_count}")
        sys.exit(1)
    else:
        print("Tutti i report sono stati generati correttamente.")
        sys.exit(0)


if __name__ == "__main__":
    main() 