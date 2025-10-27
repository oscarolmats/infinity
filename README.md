# INFINITY - Excel/CSV Viewer med Skiktfunktionalitet

En webbapplikation för att visa och manipulera Excel- och CSV-filer med avancerad funktionalitet för att dela objekt i skikt.

## Funktioner

-  **Excel & CSV-support** - Läs både .xlsx och .csv/.tsv filer
-  **Filtrering** - Global och per-kolumn filtrering
-  **Gruppering** - Gruppera data efter valfri kolumn
-  **Skiktfunktionalitet** - Dela objekt i konfigurerbara skikt
-  **Automatisk skalning** - Net Area, Volume och Count skalas automatiskt
-  **Visuella indikatorer** - Nya skikt markeras med badges

## Installation

1. Klona repot:
```bash
git clone https://github.com/oscarolmats/infinity.git
cd infinity
```

2. Installera dependencies:
```bash
npm install
```

3. Starta servern:
```bash
npm start
```

4. Öppna http://localhost:3000 i din webbläsare

## Användning

1. **Ladda fil** - Välj en Excel (.xlsx) eller CSV-fil
2. **Filtrera** - Använd globalt filter eller per-kolumn filter
3. **Gruppera** - Välj en kolumn för gruppering eller "(ingen)"
4. **Skikta** - Klicka "Skikta" på rader eller "Skikta grupp" på grupper
5. **Konfigurera** - Ange antal skikt och eventuellt tjocklekar

## Teknisk stack

- **Backend**: Node.js, Express
- **Excel-hantering**: ExcelJS
- **Frontend**: Vanilla JavaScript, HTML5, CSS3
- **File upload**: Multer

## Licens

MIT
