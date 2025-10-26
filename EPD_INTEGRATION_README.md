# EPD Integration Guide

## Översikt
Programmet stöder nu val av standard EPD:er (Environmental Product Declarations) i alternativ klimatresurs-mappning.

## Hur det fungerar

### 1. EPD-filformat
EPD-filer ska vara i CSV-format med semikolon som separator, enligt följande struktur:
- Kolumn `Name (en)`: Produktnamn
- Kolumn `Ref. unit`: Referensenhet (t.ex. kg, m², m³)
- Kolumn `Module`: LCA-modul (A1-A3, A4, A5)
- Kolumn `GWPtotal (A2)`: Global Warming Potential (kg CO₂e)

### 2. Lägga till EPD-filer
1. Skapa en mapp `epd` i programmets rotmapp
2. Lägg till CSV-filer i denna mapp
3. Uppdatera `epd/index.json` med filnamnen
4. Starta om programmet

### 3. Användning
1. Öppna alternativ klimatresurs-modal
2. Välj "Välj från EPD-fil" 
3. Välj önskad EPD från dropdown-listan
4. Granska förhandsvisningen (inkl. URL, deklarationsägare, etc.)
5. Klicka "Lägg till resurs"

## Exempel EPD-fil
Se `98bf1f1c-275a-428f-9f27-6e07ec19f810.csv` för exempel på korrekt format.

## Teknisk implementation
- **Dynamisk filläsning:** EPD-filer läses från `epd/index.json`
- **Modulär parsing:** CSV-parsning hanteras av `epd-parser.js`
- **Automatisk konvertering:** Klimatpåverkan-värden konverteras till rätt enheter
- **Intelligent omräkningsfaktor:** Sätts automatiskt baserat på enhet (m² → kg/m²)

## Filstruktur
```
epd/
├── index.json                    # Lista över tillgängliga EPD-filer
├── 98bf1f1c-275a-428f-9f27-6e07ec19f810.csv
├── bddf46b3-86a7-47d0-913f-38437ddda3ff.csv
└── c898b335-b39b-4d55-bd38-ac5bf864ebf2.csv
```

## Fallback
Om `epd/index.json` inte kan läsas, används hårdkodad lista med befintliga EPD-filer.
