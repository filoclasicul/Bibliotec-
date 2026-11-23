import json
from pathlib import Path

import pandas as pd

PROJECT_DIRECTORY = Path(__file__).resolve().parents[1]
DIRECTORY = PROJECT_DIRECTORY / "data"
EXCEL_FILE = "Inventar_Biblioteci_Compilat.xlsx"
JSON_OUT = DIRECTORY / "carti.json"

def clasifica_sectiune(titlu):
    t = str(titlu).lower()
    if any(k in t for k in ["dicționar", "dictionnaire", "lexicon", "lexic", "glossary", "vocabulario", "concordance"]):
        return "Dicționare & Lexicoane"
    if any(k in t for k in ["latin", "cicero", "vergilius", "ovidiu", "tacit", "plaut", "seneca", "horatius", "suetonius", "caesar", "roma", "roman"]):
        return "Filologie Latină & Roma"
    if any(k in t for k in ["grec", "greek", "platon", "homer", "herodot", "bizan", "elen", "septuagint", "delphi", "oracl", "athens"]):
        return "Filologie Greacă & Bizanț"
    if any(k in t for k in ["biblia", "testament", "patristic", "părinți", "creștin", "teolog", "gnostic", "nyssa"]):
        return "Teologie & Patristică"
    if any(k in t for k in ["român", "dacia", "iași", "moldov", "cantemir", "rosetti", "densușianu"]):
        return "Studii Românești & Istorie"
    return "Altele"

def curata_valoare(val):
    if pd.isna(val):
        return "-"
    if isinstance(val, (int, float)) and not isinstance(val, bool) and float(val).is_integer():
        return str(int(val))
    s = str(val).strip()
    return s if s else "-"


def gaseste_fisier_excel():
    fisier_preferat = DIRECTORY / EXCEL_FILE
    if fisier_preferat.exists():
        return fisier_preferat

    fisier_plan = DIRECTORY / "Inventar_Biblioteca_Organizat_Pe_Rafturi.xlsx"
    fisiere_excel = sorted(
        fisier for fisier in DIRECTORY.glob("*.xlsx")
        if fisier.resolve() != fisier_plan.resolve()
    )
    if len(fisiere_excel) == 1:
        return fisiere_excel[0]
    if not fisiere_excel:
        raise FileNotFoundError(
            f"Nu am găsit fișierul {EXCEL_FILE} sau vreun fișier .xlsx în {DIRECTORY}"
        )

    nume_fisiere = ", ".join(fisier.name for fisier in fisiere_excel)
    raise RuntimeError(
        f"Am găsit mai multe fișiere .xlsx în {DIRECTORY}: {nume_fisiere}. "
        f"Redenumește registrul dorit în {EXCEL_FILE}."
    )


def main():
    fisier_excel = gaseste_fisier_excel()
    df = pd.read_excel(fisier_excel)

    if "Titlu" not in df.columns:
        raise ValueError(
            f"Registrul {fisier_excel.name} trebuie să conțină o coloană «Titlu»."
        )

    # Normalizare nume coloane în funcție de ce există în fișier
    col_nr_inv = "Număr de Inventar" if "Număr de Inventar" in df.columns else "Cod de Înregistrare"
    col_autor = "Autor / Coordonator" if "Autor / Coordonator" in df.columns else "Autor"
    col_an = "An / Ediție" if "An / Ediție" in df.columns else "Anul Apariției"
    col_ex = "Număr exemplare (adiacente)" if "Număr exemplare (adiacente)" in df.columns else "Număr Exemplare"

    carti = []
    for idx, row in df.iterrows():
        titlu = curata_valoare(row.get("Titlu", "-"))
        autor = curata_valoare(row.get(col_autor, "-"))
        editura = curata_valoare(row.get("Editură", "-"))
        an = curata_valoare(row.get(col_an, "-"))
        raft = curata_valoare(row.get("Raft", "-"))
        nr_inv = curata_valoare(row.get(col_nr_inv, f"INV-{idx+1:04d}"))
        exemplare = curata_valoare(row.get(col_ex, "1"))
        obs = curata_valoare(row.get("Observații", "-"))

        carti.append({
            "nr_inv": nr_inv,
            "titlu": titlu,
            "autor": autor,
            "editura": editura,
            "an": an,
            "raft": raft,
            "exemplare": exemplare,
            "observatii": obs,
            "sectiune": clasifica_sectiune(titlu)
        })

    with JSON_OUT.open("w", encoding="utf-8") as f:
        json.dump(carti, f, ensure_ascii=False, indent=2)

    print(f"Am exportat {len(carti)} intrări din {fisier_excel.name} în {JSON_OUT}.")


if __name__ == "__main__":
    main()