#!/usr/bin/env python3
import os
import sys
import re
import subprocess
import platform
from pathlib import Path

import pandas as pd

def is_windows():
    return platform.system().lower().startswith("win")

def has_cmd(cmd):
    from shutil import which
    return which(cmd) is not None

def ask_file():
    while True:
        p = input("Pad naar Access-bestand (.mdb of .accdb): ").strip('" ').strip()
        if not p:
            continue
        path = Path(p).expanduser().resolve()
        if path.exists() and path.suffix.lower() in [".mdb", ".accdb"]:
            return path
        print("Bestand niet gevonden of geen .mdb/.accdb. Probeer opnieuw.")

def ask_output_dir(default):
    p = input(f"Outputmap [{default}]: ").strip()
    if not p:
        p = default
    out = Path(p).expanduser().resolve()
    out.mkdir(parents=True, exist_ok=True)
    return out

def ask_format():
    while True:
        f = input("Formaat: 1) CSV  2) Excel [1/2]: ").strip()
        if f in ("1", "2"):
            return "csv" if f == "1" else "xlsx"

def ask_selection(table_names):
    print("\nGevonden tabellen:")
    for i, t in enumerate(table_names, 1):
        print(f"{i:>2}. {t}")
    print("Kies tabellen met komma's, bereik of 'alles'. Voorbeeld: 1,3-5")
    while True:
        s = input("Uw keuze: ").strip().lower()
        if s in ("alles", "all", "a"):
            return table_names
        idxs = set()
        ok = True
        for part in s.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                try:
                    a, b = [int(x) for x in part.split("-", 1)]
                    for j in range(min(a, b), max(a, b) + 1):
                        idxs.add(j)
                except Exception:
                    ok = False
            else:
                try:
                    idxs.add(int(part))
                except Exception:
                    ok = False
        if ok and idxs and max(idxs) <= len(table_names) and min(idxs) >= 1:
            return [table_names[i - 1] for i in sorted(idxs)]
        print("Ongeldige selectie. Probeer opnieuw.")

# ---------- Backends ----------
def list_tables_pyodbc(db_path):
    import pyodbc
    driver_candidates = [
        "{Microsoft Access Driver (*.mdb, *.accdb)}",
        "{Microsoft Access Driver (*.mdb)}",
        "{Microsoft Access Driver (*.accdb)}",
    ]
    last_err = None
    for drv in driver_candidates:
        try:
            conn_str = f"DRIVER={drv};DBQ={db_path};"
            with pyodbc.connect(conn_str, autocommit=True) as cn:
                cur = cn.cursor()
                tables = []
                for row in cur.tables(tableType="TABLE"):
                    name = row.table_name
                    if not name.lower().startswith("msys"):
                        tables.append(name)
                # ook zichtbare query's als views ophalen
                for row in cur.tables(tableType="VIEW"):
                    name = row.table_name
                    if not name.lower().startswith("msys"):
                        tables.append(name)
                return drv, tables
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f"Kon geen Access ODBC-driver gebruiken. Laatste fout: {last_err}")

def read_table_pyodbc(db_path, table_name):
    import pyodbc
    drv, _ = list_tables_pyodbc(db_path)
    conn_str = f"DRIVER={drv};DBQ={db_path};"
    with pyodbc.connect(conn_str, autocommit=True) as cn:
        q = f"SELECT * FROM [{table_name}]"
        return pd.read_sql(q, cn)

def list_tables_mdbtools(db_path):
    # Alleen .mdb betrouwbaar
    if db_path.suffix.lower() != ".mdb":
        raise RuntimeError("mdbtools ondersteunt meestal geen .accdb. Gebruik Windows met ODBC voor .accdb.")
    if not has_cmd("mdb-tables"):
        raise RuntimeError("mdb-tools niet gevonden. Installeer mdbtools via uw pakketbeheer.")
    res = subprocess.run(["mdb-tables", "-1", str(db_path)], capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError(f"mdb-tables fout: {res.stderr.strip()}")
    tables = [t.strip() for t in res.stdout.splitlines() if t.strip()]
    tables = [t for t in tables if not t.lower().startswith("msys")]
    return tables

def read_table_mdbtools(db_path, table_name):
    if not has_cmd("mdb-export"):
        raise RuntimeError("mdb-export niet gevonden. Installeer mdbtools.")
    res = subprocess.run(["mdb-export", str(db_path), table_name], capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError(f"mdb-export fout voor {table_name}: {res.stderr.strip()}")
    # mdb-export geeft CSV terug; inlezen naar DataFrame
    from io import StringIO
    return pd.read_csv(StringIO(res.stdout))

# ---------- Main ----------
def main():
    db_path = ask_file()
    out_dir = ask_output_dir(default=str(Path(db_path).with_suffix("").name + "_export"))
    out_fmt = ask_format()

    tables = []
    backend = None

    try:
        if is_windows():
            backend = "pyodbc"
            _, tables = list_tables_pyodbc(db_path)
        else:
            # probeer mdbtools eerst
            if db_path.suffix.lower() == ".mdb":
                backend = "mdbtools"
                tables = list_tables_mdbtools(db_path)
            else:
                # laatste redmiddel: gebruiker op niet-Windows met .accdb
                # eventueel werkt een ODBC-driver via wijn of externe driver, maar dat is buiten scope
                raise RuntimeError("Op macOS/Linux is .accdb zonder ODBC-driver niet ondersteund. Gebruik Windows of converteer eerst naar .mdb.")
    except Exception as e:
        print(f"Kon tabellen niet ophalen: {e}")
        sys.exit(1)

    if not tables:
        print("Geen tabellen gevonden.")
        sys.exit(1)

    selection = ask_selection(tables)

    for t in selection:
        try:
            if backend == "pyodbc":
                df = read_table_pyodbc(db_path, t)
            else:
                df = read_table_mdbtools(db_path, t)

            safe_name = re.sub(r'[\\/:*?"<>|]+', "_", t).strip()
            if out_fmt == "csv":
                out_file = out_dir / f"{safe_name}.csv"
                df.to_csv(out_file, index=False)
            else:
                out_file = out_dir / f"{safe_name}.xlsx"
                df.to_excel(out_file, index=False)
            print(f"OK: {t} -> {out_file}")
        except Exception as e:
            print(f"Fout bij {t}: {e}")

    print("\nKlaar.")

if __name__ == "__main__":
    main()
    