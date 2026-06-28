#!/usr/bin/env python3
"""Model electric – calcul metan si energie electrica din co-digestie.
Versiune COMPLETA cu Ec.9 (temperatura), Ec.10 (cinetica), YCH4 FINAL (Editia 3).
TOTUL se citeste din CSV.
"""

from __future__ import annotations

import csv
import math
from pathlib import Path
from typing import Dict, List

import matplotlib.pyplot as plt
from openpyxl import Workbook


# Aliasuri pentru antete CSV (fara underscore)
HEADER_ALIASES = {
    # Data
    "data": "data",
    "Data": "data",
    "date": "data",
    "Date": "data",
    "zi": "data",
    "Zi": "data",

    # Parametri generali
    "M": "M",
    "m": "M",

    "etaB": "etaB",
    "etab": "etaB",
    "EtaB": "etaB",
    "eta_b": "etaB",

    "etaEl": "etaEl",
    "etael": "etaEl",
    "EtaEl": "etaEl",
    "eta_el": "etaEl",

    "LHV": "LHV",
    "lhv": "LHV",

    # Proportii
    "aFV": "aFV",
    "afv": "aFV",
    "AFV": "aFV",

    "aRM": "aRM",
    "arm": "aRM",
    "ARM": "aRM",

    "aDM": "aDM",
    "adm": "aDM",
    "ADM": "aDM",

    # Materie uscata
    "xsuFV": "xsuFV",
    "xsufv": "xsuFV",

    "xsuRM": "xsuRM",
    "xsurm": "xsuRM",

    "xsuDM": "xsuDM",
    "xsudm": "xsuDM",

    # Materie organica
    "xcoFV": "xcoFV",
    "xcofv": "xcoFV",

    "xcoRM": "xcoRM",
    "xcorm": "xcoRM",

    "xcoDM": "xcoDM",
    "xcodm": "xcoDM",

    # Potențial biochimic
    "VboFV": "VboFV",
    "vbofv": "VboFV",

    "VboRM": "VboRM",
    "vborm": "VboRM",

    "VboDM": "VboDM",
    "vbodm": "VboDM",

    # Temperatura
    "T": "T",
    "Temp": "T",
    "temp": "T",
    "Temperatura": "T",
    "temperatura": "T",

    # Timp
    "t": "t",
    "Durata": "t",
    "durata": "t",
    "Timp": "t",
    "timp": "t",

    # Model
    "theta": "theta",
    "Theta": "theta",

    "k1": "k1",
    "K1": "k1",
}


def normalize_header(name: str) -> str:
    return "".join(ch for ch in name.strip() if ch.isalnum())

def normalize_input_row(row: Dict[str, str]) -> Dict[str, str]:
    normalized = {}
    for key, value in row.items():
        mapped = HEADER_ALIASES.get(normalize_header(key))
        if mapped:
            normalized[mapped] = value
    return normalized


def parse_float(value: str) -> float:
    cleaned = value.strip().replace(" ", "")
    if not cleaned:
        raise ValueError("Valoare numerica lipsa")
    return float(cleaned.replace(",", "."))


def detect_delimiter(sample: str) -> str:
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;")
        return dialect.delimiter
    except csv.Error:
        return ";"


# Campuri obligatorii (fara underscore)
REQUIRED_FIELDS = [
    "M", "etaB", "etaEl", "LHV",
    "aFV", "aRM", "aDM",
    "xsuFV", "xsuRM", "xsuDM",
    "xcoFV", "xcoRM", "xcoDM",
    "VboFV", "VboRM", "VboDM",
    "T", "t", "theta", "k1",
]


def parse_row(row: Dict[str, str], row_index: int) -> Dict[str, float]:
    parsed = {}

    for field in REQUIRED_FIELDS:
        if field not in row or not row[field].strip():
            raise KeyError(
                f"Eroare la randul {row_index}: lipseste coloana '{field}' sau valoarea este goala."
            )
        parsed[field] = parse_float(row[field])

    return parsed


# ============================================================
#  CALCUL COMPLET CU TEMPERATURA + CINETICA (Ec.9 + Ec.10)
# ============================================================
def calculate(values: Dict[str, float]) -> Dict[str, float]:

    # Date de intrare
    M = values["M"]
    etaB = values["etaB"]
    etaEl = values["etaEl"]
    LHV = values["LHV"]

    aFV = values["aFV"]
    aRM = values["aRM"]
    aDM = values["aDM"]

    xsuFV = values["xsuFV"]
    xsuRM = values["xsuRM"]
    xsuDM = values["xsuDM"]

    xcoFV = values["xcoFV"]
    xcoRM = values["xcoRM"]
    xcoDM = values["xcoDM"]

    VboFV = values["VboFV"]
    VboRM = values["VboRM"]
    VboDM = values["VboDM"]

    # NOILE date pentru temperatura si cinetica
    T = values["T"]       # temperatura essaiului (35,40,45)
    t = values["t"]       # durata essaiului (14,21,28)
    theta = values["theta"]
    k1 = values["k1"]

    # -------------------------
    # Ec.2 – kCH4,mix
    # -------------------------
    k_CH4 = (
        aFV * xsuFV * xcoFV * VboFV +
        aRM * xsuRM * xcoRM * VboRM +
        aDM * xsuDM * xcoDM * VboDM
    )

    # Ec.3 – VCH4,mix
    VCH4_mix = M * k_CH4

    # Ec.4 – VCH4,real
    VCH4_real = VCH4_mix * etaB

    # Ec.6 – SV
    SV = M * (
        aFV * xsuFV * xcoFV +
        aRM * xsuRM * xcoRM +
        aDM * xsuDM * xcoDM
    )

    # -------------------------
    # Ec.9 – Corecția termică
    # -------------------------
    Tref = 35
    VCH4_T = VCH4_real * math.exp(theta * (T - Tref))

    # -------------------------
    # Ec.10 – Model cinetic
    # -------------------------
    factor_cinetic = 1 - math.exp(-k1 * t)
    VCH4_t = VCH4_T * factor_cinetic

    # -------------------------
    # Ec.13 – Energie chimică
    # -------------------------
    Qcomb = VCH4_t * LHV

    # Ec.14 – Energie electrică
    Eel = Qcomb * etaEl

    # YCH4 FINAL (Editia 3)
    YCH4_final = VCH4_real / SV if SV > 0 else 0

    return {
        "k_CH4": k_CH4,
        "VCH4_mix": VCH4_mix,
        "VCH4_real": VCH4_real,
        "VCH4_T": VCH4_T,
        "VCH4_t": VCH4_t,
        "SV": SV,
        "YCH4": YCH4_final,
        "Qcomb": Qcomb,
        "Eel": Eel,
    }


OUTPUT_COLUMNS = [
    ("data", "Data"),
    ("M", "M,kg SF"),
    ("k_CH4", "k_CH4,Nm3/kgSF"),
    ("VCH4_mix", "VCH4_mix,Nm3"),
    ("VCH4_real", "VCH4_real,Nm3"),
    ("VCH4_T", "VCH4(T),Nm3"),
    ("VCH4_t", "VCH4(t),Nm3"),
    ("SV", "SV,kg"),
    ("YCH4", "YCH4,Nm3/kgSV"),
    ("Qcomb", "Qcomb,kWh"),
    ("Eel", "Eel,kWh"),
]


def read_input(path: Path) -> List[Dict[str, str]]:
    sample = path.read_text(encoding="utf-8-sig")
    delimiter = detect_delimiter(sample[:2048])
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=delimiter)
        return [normalize_input_row(row) for row in reader]


def build_output_rows(input_rows: List[Dict[str, str]]) -> List[Dict[str, object]]:
    output_rows = []
    for idx, row in enumerate(input_rows, start=1):
        if not any(str(v).strip() for v in row.values()):
            continue

        values = parse_row(row, idx)
        result = calculate(values)

        output_rows.append({
            "data": row.get("data", "").strip(),
            "M": values["M"],
            "k_CH4": result["k_CH4"],
            "VCH4_mix": result["VCH4_mix"],
            "VCH4_real": result["VCH4_real"],
            "VCH4_T": result["VCH4_T"],
            "VCH4_t": result["VCH4_t"],
            "SV": result["SV"],
            "YCH4": result["YCH4"],
            "Qcomb": result["Qcomb"],
            "Eel": result["Eel"],
        })

    return output_rows


def write_output(path: Path, rows: List[Dict[str, object]]) -> Path:
    workbook = Workbook()
    ws = workbook.active
    ws.title = "Energie_electrica"

    ws.append([display for _, display in OUTPUT_COLUMNS])

    for row in rows:
        ws.append([
            row.get(key, "") if not isinstance(row.get(key), float)
            else round(row.get(key), 6)
            for key, _ in OUTPUT_COLUMNS
        ])

    workbook.save(path)
    return path


def print_pretty_results(rows: List[Dict[str, object]]) -> None:
    print("\n" + "=" * 72)
    print("REZULTATE – METAN SI ENERGIE ELECTRICA (cu temperatura + cinetica)")
    print("=" * 72)

    for i, row in enumerate(rows, start=1):
        label = row.get("data", f"Rand {i}")
        print(f"\n[{label}]")
        print("-" * 72)
        print(f"VCH4_real [Nm3]       : {row['VCH4_real']:.6f}")
        print(f"VCH4(T) [Nm3]         : {row['VCH4_T']:.6f}")
        print(f"VCH4(t) [Nm3]         : {row['VCH4_t']:.6f}")
        print(f"Qcomb [kWh]           : {row['Qcomb']:.6f}")
        print(f"Eel [kWh]             : {row['Eel']:.6f}")
        print(f"YCH4 FINAL [Nm3/kgSV] : {row['YCH4']:.6f}")

    print("=" * 72)


def plot_results(rows: List[Dict[str, object]], output_dir: Path) -> None:
    labels = [row.get("data") or f"Rand {i+1}" for i, row in enumerate(rows)]
    eel_values = [row["Eel"] for row in rows]
    y_values = [row["YCH4"] for row in rows]

    output_dir.mkdir(parents=True, exist_ok=True)

    # Eel
    plt.figure(figsize=(12, 4))
    plt.plot(labels, eel_values, marker="o")
    plt.title("Eel – energie electrica [kWh]")
    plt.grid(True)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(output_dir / "Eel.png", dpi=150)

    # YCH4
    plt.figure(figsize=(12, 4))
    plt.plot(labels, y_values, marker="o")
    plt.title("YCH4 FINAL – randament specific [Nm3/kgSV]")
    plt.grid(True)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(output_dir / "YCH4.png", dpi=150)

    plt.show()


def main():
    input_path = Path("Data.csv")
    output_path = Path("rezultate_electrica.xlsx")

    if not input_path.exists():
        raise FileNotFoundError("Fisierul Data.csv nu exista in folderul programului!")

    input_rows = read_input(input_path)
    print(input_rows[0].keys()) #Aici am pus
    output_rows = build_output_rows(input_rows)
    saved = write_output(output_path, output_rows)
    print_pretty_results(output_rows)
    plot_results(output_rows, output_path.parent)
    print(f"Rezultate salvate in: {saved}")


if __name__ == "__main__":
    main()
