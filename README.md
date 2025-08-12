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
