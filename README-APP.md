# Contaibil - Applicazione Web per Riconciliazione Fatture

Applicazione web moderna per confrontare i dati delle fatture ADE con i dati gestionali, evidenziando le discrepanze e fornendo KPI utili.

## Caratteristiche

- 📤 **Caricamento File**: Supporta file XML e CSV per fatture e dati gestionali
- 📊 **Dashboard KPI**: Visualizza metriche chiave come totale fatture, corrispondenze, discrepanze e percentuale
- 🔍 **Confronto Intelligente**: Confronta automaticamente i dati e identifica le discrepanze
- 🎯 **Filtri Avanzati**: Filtra i risultati per visualizzare tutte le fatture, solo corrispondenze o solo discrepanze
- 🎨 **UI Moderna**: Design responsive e moderno con evidenziazione visiva delle discrepanze
- 💬 **Codice Commentato**: Codice ben organizzato e commentato per facile manutenzione

## Struttura del Progetto

```
Contaibil/
├── index.html              # Pagina principale dell'applicazione
├── css/
│   └── styles.css         # Fogli di stile con design responsive
├── js/
│   ├── main.js            # Modulo principale e coordinazione
│   ├── fileParser.js      # Parsing di file XML/CSV
│   ├── dataComparator.js  # Logica di confronto dati
│   └── uiController.js    # Gestione interfaccia utente
└── sample-data/           # Dati di esempio per test
    ├── fatture-ade.csv
    ├── fatture-ade.xml
    ├── gestionale.csv
    └── gestionale.xml
```

## Come Utilizzare

### 1. Avviare l'Applicazione

Aprire il file `index.html` in un browser web moderno. L'applicazione è completamente client-side e non richiede un server.

**Opzione 1**: Aprire direttamente il file
```bash
# Su Mac
open index.html

# Su Linux
xdg-open index.html

# Su Windows
start index.html
```

**Opzione 2**: Utilizzare un server locale (consigliato per sviluppo)
```bash
# Con Python 3
python -m http.server 8000

# Con Node.js (se hai http-server installato)
npx http-server

# Poi aprire http://localhost:8000
```

### 2. Caricare i File

1. Cliccare sulla prima casella per caricare il file delle **Fatture ADE** (formato XML o CSV)
2. Cliccare sulla seconda casella per caricare il file dei **Dati Gestionali** (formato XML o CSV)
3. Il pulsante "Confronta Dati" si attiverà automaticamente quando entrambi i file sono caricati

### 3. Visualizzare i Risultati

- **Dashboard KPI**: Visualizza automaticamente il riepilogo dei risultati
- **Tabella Confronto**: Mostra tutte le fatture con evidenziazione delle discrepanze
- **Filtri**: Usa i pulsanti radio per filtrare i risultati
  - **Tutti**: Mostra tutte le fatture
  - **Solo Corrispondenze**: Mostra solo le fatture che corrispondono
  - **Solo Discrepanze**: Mostra solo le fatture con differenze

## Formato dei File

### CSV
Il file CSV deve avere la seguente struttura:
```csv
invoice_number,date,customer,amount
FAT-2025-001,2025-01-15,Azienda Alpha SRL,1250.00
FAT-2025-002,2025-01-16,Beta Company SpA,890.50
```

### XML
Il file XML deve avere la seguente struttura:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<invoices>
    <invoice>
        <invoice_number>FAT-2025-001</invoice_number>
        <date>2025-01-15</date>
        <customer>Azienda Alpha SRL</customer>
        <amount>1250.00</amount>
    </invoice>
</invoices>
```

## Test con Dati di Esempio

Nella cartella `sample-data/` sono presenti file di esempio che puoi utilizzare per testare l'applicazione:

- `fatture-ade.csv` / `fatture-ade.xml`: File fatture ADE di esempio
- `gestionale.csv` / `gestionale.xml`: File gestionale di esempio

I dati di esempio includono:
- ✓ Corrispondenze esatte (FAT-2025-001, FAT-2025-003, FAT-2025-004, FAT-2025-006)
- ⚠ Discrepanze negli importi (FAT-2025-002: €5 di differenza, FAT-2025-005: €100 di differenza)
- ⚠ Fatture mancanti (FAT-2025-007 solo in ADE, FAT-2025-008 solo in gestionale)

## Tecnologie Utilizzate

- **HTML5**: Struttura semantica dell'applicazione
- **CSS3**: Design moderno con variabili CSS, Grid e Flexbox
- **JavaScript ES6+**: Moduli ES6 per una struttura pulita e modulare
- **Web APIs**: FileReader, DOMParser per il parsing dei file

## Browser Supportati

- Chrome/Edge (versione 90+)
- Firefox (versione 88+)
- Safari (versione 14+)

## Funzionalità Chiave

### Confronto Intelligente
L'applicazione confronta le fatture utilizzando il numero di fattura come chiave e identifica:
- ✓ Corrispondenze esatte (differenza < €0.01)
- ⚠ Differenze negli importi
- ⚠ Fatture presenti solo in un dataset

### Dashboard KPI
- **Fatture Totali**: Numero totale di fatture analizzate
- **Corrispondenze**: Fatture che corrispondono perfettamente
- **Discrepanze**: Fatture con differenze o mancanti
- **Percentuale**: Percentuale di corrispondenze sul totale

### Evidenziazione Visiva
- Righe verdi: Corrispondenze
- Righe rosse: Discrepanze
- Valori evidenziati: Importi con differenze

## Sviluppo e Personalizzazione

Il codice è organizzato in moduli separati per facilitare la manutenzione:

- **main.js**: Coordina l'applicazione e gestisce gli eventi
- **fileParser.js**: Parsing flessibile di XML/CSV con supporto per diversi formati
- **dataComparator.js**: Logica di business per il confronto
- **uiController.js**: Gestione dell'interfaccia utente

## Licenza

© 2025 Contaibil - Sistema di Riconciliazione
