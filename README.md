# access2csv — Exporteer Microsoft Access (.mdb/.accdb) tabellen naar CSV via ODBC

Een eenvoudige, betrouwbare CLI-tool in Python om alle of geselecteerde tabellen uit een Access‑database naar CSV te schrijven met de **Microsoft Access ODBC‑driver**.

## Kenmerken
- Werkt met `.accdb` en `.mdb`
- Exporteert **alle** tabellen of een opgegeven subset
- Optioneel: ook **views** meenemen
- Instelbare `delimiter` (`,`, `;`, `\t`), `encoding`, en `line-terminator`
- Veilige bestandsnamen en automatische map‑aanmaak
- Grote tabellen worden in batches gelezen (lage memory‑footprint)
- Heldere foutmeldingen en exit-codes

## Vereisten
- Windows met de **Microsoft Access Database Engine (ODBC)**
  - Driver naam: `Microsoft Access Driver (*.mdb, *.accdb)`
  - Tip: Installeer *Access Database Engine 2016 Redistributable* (32‑ of 64‑bit passend bij je Python).
- Python 3.9+  
- `pyodbc` (`pip install pyodbc`)

> macOS/Linux? Dit script focust op ODBC met de Microsoft‑driver (meestal Windows). Voor andere platforms kun je alternatieven zoals `mdbtools` overwegen.

## Installatie
```bash
pip install pyodbc

## Gebruik

### Basis
```bash
python access2csv.py "C:\pad\naar\db.accdb"

python access2csv.py "C:\dbs\mijn.accdb" \
  --out "C:\export" \
  --delimiter ";" \
  --encoding "utf-8" \
  --include-views \
  --tables "Customers" "Orders"

### Alle opties

positional:
  db_path               Pad naar .accdb of .mdb

optional:
  -o, --out PATH        Uitvoormap (default: ./export)
  -t, --tables ...      Specifieke tabellen exporteren (meerdere namen)
  --include-views       Views ook exporteren
  -d, --delimiter STR   CSV delimiter (default: ,)  [specials: \t voor tab]
  -e, --encoding STR    Tekencodering (default: utf-8)
  --lineterm STR        Regelscheiding (default: system; bv. \n of \r\n)
  --batch-size N        Aantal rijen per fetch (default: 10000)
  --dsn NAME            i.p.v. pad, gebruik een DSN (ODBC) + optioneel --uid/--pwd
  --uid USER            Gebruikersnaam voor DSN
  --pwd PASS            Wachtwoord voor DSN
  -q, --quiet           Minder output
  --dry-run             Toon wat er zou gebeuren, zonder te schrijven
  --version             Toon versie

### Voorbeelden
# Alle tabellen + views, ; als delimiter (Excel-vriendelijk in NL)
python access2csv.py "C:\data\sales.accdb" --include-views --delimiter ";" --out "C:\exports"

# Alleen twee tabellen
python access2csv.py "D:\db\legacy.mdb" --tables Artikelen Orders

# Via vooraf ingestelde ODBC-DSN
python access2csv.py --dsn MijnAccessDSN --uid sa --pwd geheim --out out\

### Exit codes
0 = succes
2 = database of driver niet gevonden
3 = connectie/ODBC-fout
4 = geen tabellen gevonden
5 = schrijf-/bestandsfout

