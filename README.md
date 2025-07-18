# Progetto Papà - Generatori Report PDF

Questo progetto contiene tre script Python per generare report PDF professionali da file CSV esportati da sistemi HMI/SCADA EcoStruxure.

## Script Disponibili

### 1. `generate_report_alarm.py` - Report Allarmi
Genera report PDF da file CSV contenenti dati di allarmi.

**Posizionamento**: Lo script deve trovarsi nella stessa cartella dei file CSV ALARM.

**Auto-selezione**: Cerca automaticamente il file CSV più recente che contiene "ALARM" nel nome.

**Colonne elaborate**: Date, Time, Alarm Message, Alarm Status

**Uso**:
```bash
python generate_report_alarm.py [--csv <file>] [--out <pdf>] [--logo <logo>] [--limit-rows N] [--dry-run]
```

### 2. `generate_report_operlog.py` - Report Log Operazioni
Genera report PDF da file CSV contenenti log delle operazioni utente.

**Posizionamento**: Lo script deve trovarsi a livello delle cartelle DDMMYY che contengono i CSV OPERLOG.

**Auto-selezione**: Cerca la cartella con data più recente (formato DDMMYY) e all'interno il file CSV OPERLOG più recente.

**Colonne elaborate**: Date, Time, User, Object_Action, Trigger, PreviousValue, ChangedValue

**Uso**:
```bash
python generate_report_operlog.py [--csv <file>] [--out <pdf>] [--logo <logo>] [--limit-rows N] [--dry-run]
```

### 3. `generate_report_batch.py` - Report Batch con Grafici Temperature
Genera report PDF da file CSV contenenti dati di batch con grafico delle temperature.

**Posizionamento**: Lo script deve trovarsi nella stessa cartella dei file CSV BATCH.

**Auto-selezione**: Cerca automaticamente il file CSV più recente che contiene "BATCH" nel nome.

**Colonne elaborate**: Date, Time, USER, TEMP_AIR_IN, TEMP_PRODUCT_1, TEMP_PRODUCT_2, TEMP_PRODUCT_3 (ignora colonne QF)

**Caratteristiche speciali**:
- Arrotonda le temperature a 1 cifra decimale
- Calcola e mostra il periodo di inizio-fine
- Genera grafico lineare delle temperature (richiede matplotlib)
- Include sottotitolo con range temporale

**Uso**:
```bash
python generate_report_batch.py [--csv <file>] [--out <pdf>] [--logo <logo>] [--limit-rows N] [--dry-run] [--chart-first] [--separate-files]
```

## Opzioni Comuni

- `--csv <file>`: Specifica il file CSV da processare (opzionale, auto-detect se omesso)
- `--out <file>`: Specifica il nome del PDF output (default: `<nome_csv>_report.pdf`)
- `--logo <file>`: Specifica il logo da includere (default: `logo.png`)
- `--limit-rows N`: Limita il numero di righe processate (utile per debug)
- `--dry-run`: Mostra informazioni senza generare il PDF
- `--chart-first` (solo BATCH): Posiziona il grafico prima della tabella
- `--separate-files` (solo BATCH): Genera due PDF separati invece di uno unico

## Caratteristiche PDF

- Formato A4 verticale
- Logo in alto a sinistra (se presente)
- Titolo centrato basato sul nome del file
- Numerazione pagine in basso a destra
- Tabelle con righe alternate grigie/bianche
- **Conversione automatica formato date**: MM/DD/YYYY → DD/MM/YY (es. 06/25/2025 → 25/06/25)
- **Word-wrap intelligente**: Le celle lunghe vanno a capo automaticamente all'interno della stessa cella
- Layout colonne ottimizzato per evitare sforamento testo
- Margini ottimizzati per la stampa (2cm laterali, 1.5cm top/bottom)

## Gestione Errori

- **File CSV non trovato**: Exit con codice di errore e messaggio chiaro
- **Logo mancante**: Continua l'esecuzione con warning, non interrompe il processo
- **Colonne mancanti**: Include nota nel PDF e processa le colonne disponibili
- **Dati vuoti**: Genera PDF con messaggio "Nessun dato disponibile"
- **Parsing fallito**: Mantiene valori raw quando la conversione automatica fallisce

## Requisiti

Installare le dipendenze con:
```bash
pip install -r requirements.txt
```

### Dipendenze principali:
- `pandas`: Parsing veloce e manipolazione CSV
- `reportlab`: Generazione PDF professionale
- `matplotlib`: Grafici temperature (solo per script BATCH)
- `pyarrow`: Parsing CSV ottimizzato (opzionale, migliora performance)

## Compatibilità

- **Python**: >= 3.9 (testato con 3.11)
- **Separatori CSV**: Auto-detect (tab, virgola, punto e virgola)
- **Encoding**: UTF-8 con fallback graceful per caratteri non standard
- **Formato dati EcoStruxure**: Riconosce automaticamente header metadata e li salta

## Esempi di Utilizzo

```bash
# Genera report alarm automatico
python generate_report_alarm.py

# Genera report batch con grafico specifico
python generate_report_batch.py --csv data/BATCH120725155609.csv --chart-first

# Test veloce senza generare PDF
python generate_report_operlog.py --dry-run --limit-rows 10

# Report personalizzato con logo specifico
python generate_report_alarm.py --logo company_logo.png --out alarm_report_custom.pdf
```

## Struttura File di Output

I PDF generati includono:
1. **Header**: Logo + Titolo del report
2. **Sottotitolo**: Periodo dati (solo script BATCH)
3. **Note**: Informazioni su colonne mancanti (se applicabile)
4. **Contenuto principale**: Tabella dati o grafico + tabella
5. **Footer**: Numerazione pagine

I file PDF vengono salvati nella stessa directory dello script con nome `<basename_input>_report.pdf`.

## Miglioramenti Implementati

### ✅ Conversione Formato Date
Le date nei CSV (formato americano MM/DD/YYYY) vengono automaticamente convertite nel formato europeo DD/MM/YY:
- **Prima**: `06/25/2025` 
- **Dopo**: `25/06/25`

### ✅ Word Wrap Intelligente
Le celle con testo lungo utilizzano word wrap automatico per evitare sforamento:
- **ALARM**: Messaggi di allarme lunghi vanno a capo nella colonna "Alarm Message"
- **OPERLOG**: Campi "Object_Action", "PreviousValue", "ChangedValue" con word wrap automatico
- **BATCH**: Campo "USER" con gestione celle lunghe

### ✅ Layout Tabelle Ottimizzato
- Larghezze colonne ridistribuite per migliore leggibilità
- Font size ridotto a 8pt per massimizzare spazio disponibile
- Leading ottimizzato per densità informazioni senza perdere chiarezza

### ✅ PDF Separati per Report Batch
Il script `generate_report_batch.py` supporta la generazione di **due PDF separati** tramite l'opzione `--separate-files`:

**File generati**:
- `<nome_file>_chart.pdf` (112 KB) - Solo grafico temperature con logo e titolo
- `<nome_file>_report.pdf` (56 KB) - Solo tabella dati con periodo e informazioni complete

**Vantaggi**:
- **Grafico ad alta risoluzione**: PDF dedicato ottimizzato per stampa o proiezione
- **Report dati compatto**: Tabella concentrata senza grafico per analisi rapida
- **Flessibilità d'uso**: Condivisione selettiva di grafico o dati secondo necessità
- **Qualità**: Ogni PDF ottimizzato per il suo contenuto specifico

**Esempio utilizzo**:
```bash
python generate_report_batch.py --separate-files --logo logo.png
# Genera: BATCH120725155609_chart.pdf + BATCH120725155609_report.pdf
```
