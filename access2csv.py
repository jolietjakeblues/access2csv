import argparse
import csv
import os
import re
import sys
from typing import Iterable, List, Tuple, Optional

try:
    import pyodbc
except ImportError:
    print("pyodbc niet geïnstalleerd. Installeer met: pip install pyodbc", file=sys.stderr)
    sys.exit(2)

VERSION = "1.0.0"

def log(msg: str, quiet: bool = False):
    if not quiet:
        print(msg)

def sanitize_filename(name: str) -> str:
    # Vervang ongeldige tekens en trim
    name = re.sub(r'[\\/:"*?<>|]+', "_", name)
    name = name.strip().strip(".")
    return name or "untitled"

def resolve_lineterminator(s: Optional[str]) -> str:
    if not s:
        return os.linesep
    mapping = {"\\n": "\n", "\\r\\n": "\r\n", "\\r": "\r"}
    return mapping.get(s, s)

def detect_access_driver() -> Optional[str]:
    # Zoek een geschikte Access-driver
    drivers = [d for d in pyodbc.drivers()]
    # Prefer 64-bit modern driver naam
    for needle in [
        "Microsoft Access Driver (*.mdb, *.accdb)",
        "Microsoft Access Driver (*.mdb)"
    ]:
        for d in drivers:
            if needle.lower() in d.lower():
                return d
    return None

def connect_via_path(db_path: str) -> pyodbc.Connection:
    driver = detect_access_driver()
    if not driver:
        print("Kon de Microsoft Access ODBC-driver niet vinden. Installeer Access Database Engine.", file=sys.stderr)
        sys.exit(2)
    if not os.path.exists(db_path):
        print(f"Database niet gevonden: {db_path}", file=sys.stderr)
        sys.exit(2)
    # DBQ wijst naar bestand; met standaard opties
    conn_str = f'DRIVER={{{driver}}};DBQ={db_path};'
    try:
        return pyodbc.connect(conn_str, autocommit=False)
    except pyodbc.Error as e:
        print(f"ODBC connectiefout: {e}", file=sys.stderr)
        sys.exit(3)

def connect_via_dsn(dsn: str, uid: Optional[str], pwd: Optional[str]) -> pyodbc.Connection:
    parts = [f"DSN={dsn}"]
    if uid:
        parts.append(f"UID={uid}")
    if pwd:
        parts.append(f"PWD={pwd}")
    conn_str = ";".join(parts) + ";"
    try:
        return pyodbc.connect(conn_str, autocommit=False)
    except pyodbc.Error as e:
        print(f"ODBC connectiefout (DSN): {e}", file=sys.stderr)
        sys.exit(3)

def list_objects(conn: pyodbc.Connection, include_views: bool) -> Tuple[List[str], List[str]]:
    # Retourneer (tables, views)
    tables, views = [], []
    cursor = conn.cursor()
    for row in cursor.tables(tableType="TABLE"):
        if row.table_name and not row.table_name.startswith("MSys"):
            tables.append(row.table_name)
    if include_views:
        for row in cursor.tables(tableType="VIEW"):
            if row.table_name:
                views.append(row.table_name)
    return (tables, views)

def export_table(conn: pyodbc.Connection,
                 name: str,
                 out_dir: str,
                 delimiter: str,
                 encoding: str,
                 lineterminator: str,
                 batch_size: int,
                 quiet: bool) -> Tuple[str, int]:
    safe = sanitize_filename(name)
    out_path = os.path.join(out_dir, f"{safe}.csv")

    # Query met expliciete kolomnamen voor nette header volgorde
    cursor = conn.cursor()
    try:
        # Haal kolomnamen op
        col_cursor = conn.cursor()
        col_cursor.execute(f"SELECT * FROM [{name}] WHERE 1=0")
        columns = [desc[0] for desc in col_cursor.description]
    except pyodbc.Error:
        # Fallback: direct SELECT * en pak description van cursor
        cursor.execute(f"SELECT * FROM [{name}] WHERE 1=0")
        columns = [desc[0] for desc in cursor.description]

    # Open writer
    try:
        os.makedirs(out_dir, exist_ok=True)
        newline_arg = ""  # verplicht voor csv-module in py3
        with open(out_path, "w", encoding=encoding, newline=newline_arg) as f:
            writer = csv.writer(f, delimiter=delimiter, quoting=csv.QUOTE_MINIMAL, lineterminator=lineterminator)
            writer.writerow(columns)

            # Lees in batches
            cursor.execute(f"SELECT * FROM [{name}]")
            total = 0
            while True:
                rows = cursor.fetchmany(batch_size)
                if not rows:
                    break
                for row in rows:
                    # pyodbc row -> tuple
                    writer.writerow(tuple(row))
                total += len(rows)
        return out_path, total
    except (OSError, IOError) as e:
        print(f"Schrijffout voor {out_path}: {e}", file=sys.stderr)
        sys.exit(5)
    except pyodbc.Error as e:
        print(f"Leesfout uit tabel [{name}]: {e}", file=sys.stderr)
        sys.exit(3)

def parse_args(argv: Optional[Iterable[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Exporteer Access (.mdb/.accdb) tabellen naar CSV via ODBC."
    )
    p.add_argument("db_path", nargs="?", help="Pad naar .accdb of .mdb (niet nodig als --dsn gebruikt wordt)")
    p.add_argument("-o", "--out", default="export", help="Uitvoermap (default: ./export)")
    p.add_argument("-t", "--tables", nargs="*", help="Specifieke tabellen exporteren (meerdere namen)")
    p.add_argument("--include-views", action="store_true", help="Neem ook views mee")
    p.add_argument("-d", "--delimiter", default=",", help="CSV delimiter (default: ,) — gebruik \\t voor tab")
    p.add_argument("-e", "--encoding", default="utf-8", help="Tekencodering (default: utf-8)")
    p.add_argument("--lineterm", default=None, help="Regelscheiding, bv. \\n of \\r\\n (default: systeem)")
    p.add_argument("--batch-size", type=int, default=10000, help="Rijen per fetch (default: 10000)")
    p.add_argument("--dsn", help="Gebruik ODBC-DSN i.p.v. pad")
    p.add_argument("--uid", help="Gebruikersnaam voor DSN")
    p.add_argument("--pwd", help="Wachtwoord voor DSN")
    p.add_argument("-q", "--quiet", action="store_true", help="Minder output")
    p.add_argument("--dry-run", action="store_true", help="Alleen tonen wat er zou gebeuren")
    p.add_argument("--version", action="store_true", help="Toon versie en stop")
    return p.parse_args(argv)

def main(argv: Optional[Iterable[str]] = None) -> int:
    args = parse_args(argv)

    if args.version:
        print(f"access2csv {VERSION}")
        return 0

    # delimiter specials
    delim = "\t" if args.delimiter == "\\t" else args.delimiter
    lineterm = resolve_lineterminator(args.lineterm)

    if args.dsn:
        conn = connect_via_dsn(args.dsn, args.uid, args.pwd)
        db_label = f"DSN={args.dsn}"
    else:
        if not args.db_path:
            print("Geef een pad naar de database of gebruik --dsn.", file=sys.stderr)
            return 2
        conn = connect_via_path(args.db_path)
        db_label = args.db_path

    try:
        tables, views = list_objects(conn, include_views=args.include_views)
    except pyodbc.Error as e:
        print(f"ODBC-fout bij ophalen van tabellen: {e}", file=sys.stderr)
        return 3

    selected: List[str]
    if args.tables:
        # Filter: alleen bestaande namen exporteren
        want = set(args.tables)
        available = set(tables + (views if args.include_views else []))
        missing = [t for t in args.tables if t not in available]
        if missing:
            print(f"Waarschuwing: niet gevonden en worden overgeslagen: {', '.join(missing)}", file=sys.stderr)
        selected = [t for t in args.tables if t in available]
    else:
        selected = tables + (views if args.include_views else [])

    if not selected:
        print("Geen tabellen/views gevonden om te exporteren.", file=sys.stderr)
        return 4

    if args.dry_run:
        print("DRY RUN — er wordt niets geschreven.")
        print(f"Bron: {db_label}")
        print(f"Zal exporteren ({len(selected)}): {', '.join(selected)}")
        print(f"Uitvoermap: {args.out} | delimiter='{delim}' | encoding='{args.encoding}'")
        return 0

    log(f"Verbonden met: {db_label}", args.quiet)
    log(f"Te exporteren objecten: {len(selected)}", args.quiet)
    os.makedirs(args.out, exist_ok=True)

    total_rows = 0
    for name in selected:
        log(f"- Export [{name}] ...", args.quiet)
        out_path, count = export_table(
            conn=conn,
            name=name,
            out_dir=args.out,
            delimiter=delim,
            encoding=args.encoding,
            lineterminator=lineterm,
            batch_size=args.batch_size,
            quiet=args.quiet,
        )
        total_rows += count
        log(f"  -> {out_path} ({count} rijen)", args.quiet)

    log(f"Klaar. {len(selected)} bestanden geschreven, totaal {total_rows} rijen.", args.quiet)
    return 0

if __name__ == "__main__":
    sys.exit(main())
