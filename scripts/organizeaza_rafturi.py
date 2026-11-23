from pathlib import Path

PROJECT_DIRECTORY = Path(__file__).resolve().parents[1]
DIRECTORY = PROJECT_DIRECTORY / "data"
EXCEL_IN = "Inventar_Biblioteca_Complet_Cu_Numere_Inventar_Noi.xlsx"
EXCEL_OUT = "Inventar_Biblioteca_Organizat_Pe_Rafturi.xlsx"

CARTI_PER_RAFT = 20
RAFTURI_PER_DULAP = 5  # Ex: A1..A5, B1..B5

# 1. Separarea pe categorii principale
def determina_domeniu(row):
    text = f"{row.get('Titlu', '')} {row.get('Autor / Coordonator', '')}".lower()
    
    # Cuvinte-cheie pentru Greacă
    kw_greaca = [
        "grec", "greek", "platon", "homer", "herodot", "eschil", "sofocle", "euripide",
        "aristotel", "eschine", "demostene", "plutarh", "plutarch", "lucian", "josephus",
        "septuagint", "elen", "bizan", "athens", "atena", "delphi", "oracl", "sappho",
        "neogreac", "neoellinik", "chantraine", "liddell", "bailly", "montanari", "muraoka"
    ]
    # Cuvinte-cheie pentru Latină
    kw_latina = [
        "latin", "cicero", "vergilius", "virgil", "ovidiu", "tacit", "plaut", "seneca",
        "horatius", "suetonius", "caesar", "cezar", "roma", "roman", "terentiu", "terențiu",
        "juvenal", "petroniu", "apuleius", "lucanus", "statius", "quintus curtius",
        "ammianus", "gellius", "fedru", "ernout", "forcellini", "niermeyer", "guțu", "gutu"
    ]
    # Cuvinte-cheie pentru Teologie / Patristică
    kw_teologie = [
        "biblia", "testament", "patristic", "părinți", "creștin", "teolog", "gnostic",
        "nyssa", "dumnezeu", "evanghel", "biseric", "scriban", "anselm", "damian"
    ]
    # Cuvinte-cheie pentru Studii Românești & Istorie
    kw_romana = [
        "român", "dacia", "iași", "moldov", "cantemir", "rosetti", "densușianu",
        "philippide", "ureche", "hasdeu", "caragiale", "eminescu", "despot", "vasile lupu"
    ]

    if any(k in text for k in kw_greaca):
        return "1. Greacă & Antichitate Elenă"
    if any(k in text for k in kw_latina):
        return "2. Latină & Roma Antică"
    if any(k in text for k in kw_teologie):
        return "3. Teologie & Patristică"
    if any(k in text for k in kw_romana):
        return "4. Studii Românești & Istorie"
    return "5. Varia & Literatură Universală"

def gaseste_fisier_excel():
    fisier_preferat = DIRECTORY / EXCEL_IN
    if fisier_preferat.exists():
        return fisier_preferat

    fisier_iesire = DIRECTORY / EXCEL_OUT
    fisiere_excel = sorted(
        fisier for fisier in DIRECTORY.glob("*.xlsx")
        if fisier.resolve() != fisier_iesire.resolve()
    )
    if len(fisiere_excel) == 1:
        return fisiere_excel[0]
    if not fisiere_excel:
        raise FileNotFoundError(
            f"Nu am găsit fișierul {EXCEL_IN} sau vreun fișier .xlsx în {DIRECTORY}"
        )

    nume_fisiere = ", ".join(fisier.name for fisier in fisiere_excel)
    raise RuntimeError(
        f"Am găsit mai multe fișiere .xlsx în {DIRECTORY}: {nume_fisiere}. "
        f"Redenumește registrul dorit în {EXCEL_IN}."
    )


def parse_exemplare(val):
    try:
        v = int(float(val))
        return v if v > 0 else 1
    except (TypeError, ValueError, OverflowError):
        return 1


def nume_litera(index):
    """Returnează etichete de raft de tip A..Z, AA..AZ, BA.."""
    eticheta = ""
    while index >= 0:
        index, rest = divmod(index, 26)
        eticheta = chr(ord("A") + rest) + eticheta
        index -= 1
    return eticheta


def main(fisier_intrare=None):
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils.dataframe import dataframe_to_rows

    fisier_intrare = Path(fisier_intrare) if fisier_intrare else gaseste_fisier_excel()
    df = pd.read_excel(fisier_intrare)

    if "Titlu" not in df.columns:
        raise ValueError(
            f"Registrul {fisier_intrare.name} trebuie să conțină o coloană «Titlu»."
        )

    coloane_autor = ["Autor / Coordonator", "Autor"]
    col_autor = next((c for c in coloane_autor if c in df.columns), None)
    if col_autor is None:
        col_autor = "Autor / Coordonator"
        df[col_autor] = "-"

    coloane_exemplare = ["Număr exemplare (adiacente)", "Număr Exemplare"]
    col_ex = next((c for c in coloane_exemplare if c in df.columns), None)
    if col_ex is None:
        df["_nr_vol"] = 1
    else:
        df["_nr_vol"] = df[col_ex].apply(parse_exemplare)

    df["Domeniu"] = df.apply(determina_domeniu, axis=1)
    df["_sort_autor"] = df[col_autor].apply(
        lambda x: "ZZZ" if str(x).strip() in ["-", "", "nan"] else str(x).strip().lower()
    )
    df["_sort_titlu"] = df["Titlu"].apply(lambda x: str(x).strip().lower())
    df = df.sort_values(
        by=["Domeniu", "_sort_autor", "_sort_titlu"]
    ).reset_index(drop=True)

    start_litere = {
        "1. Greacă & Antichitate Elenă": 0,
        "2. Latină & Roma Antică": 5,
        "3. Teologie & Patristică": 11,
        "4. Studii Românești & Istorie": 14,
        "5. Varia & Literatură Universală": 18,
    }
    cote_alocate = []

    for domeniu, grup in df.groupby("Domeniu", sort=True):
        idx_litera = start_litere.get(domeniu, 0)
        nr_raft = 1
        vol_pe_raft_curent = 0

        for _, row in grup.iterrows():
            vol_carte = row["_nr_vol"]
            if vol_pe_raft_curent > 0 and (
                vol_pe_raft_curent + vol_carte
            ) > CARTI_PER_RAFT:
                nr_raft += 1
                vol_pe_raft_curent = 0
                if nr_raft > RAFTURI_PER_DULAP:
                    nr_raft = 1
                    idx_litera += 1

            cote_alocate.append(f"{nume_litera(idx_litera)}{nr_raft}")
            vol_pe_raft_curent += vol_carte

    df["Cotă Raft Nou"] = cote_alocate
    df = df.drop(columns=["_nr_vol", "_sort_autor", "_sort_titlu"])

    coloane_prioritare = [
        "Nr. Crt. Global", "Număr de Inventar", "Cod de Înregistrare",
        "Domeniu", "Cotă Raft Nou", "Raft",
    ]
    alte_coloane = [c for c in df.columns if c not in coloane_prioritare]
    df = df[[c for c in coloane_prioritare if c in df.columns] + alte_coloane]

    wb = Workbook()
    ws = wb.active
    ws.title = "Plan Așezare Rafturi"

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_header = Font(color="FFFFFF", bold=True)
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF")
    )
    fill_zebra = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    for cell in ws[1]:
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = align_center

    for i, row in enumerate(
        ws.iter_rows(min_row=2, max_col=len(df.columns), max_row=ws.max_row)
    ):
        for j, cell in enumerate(row):
            col_name = df.columns[j]
            if col_name in [
                "Nr. Crt. Global", "Număr de Inventar", "Cod de Înregistrare",
                "Cotă Raft Nou", "Raft", "An / Ediție",
                "Număr exemplare (adiacente)",
            ]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
            cell.border = thin_border
            if i % 2 == 0:
                cell.fill = fill_zebra

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    fisier_iesire = fisier_intrare.parent / EXCEL_OUT
    wb.save(fisier_iesire)

    print("=== REZUMAT DISTRIBUȚIE PE RAFTURI ===")
    rezumat = df.groupby(["Domeniu", "Cotă Raft Nou"]).size().reset_index(name="Intrări")
    for dom, grup in rezumat.groupby("Domeniu"):
        print(
            f"\n{dom}: de la {grup['Cotă Raft Nou'].iloc[0]} "
            f"la {grup['Cotă Raft Nou'].iloc[-1]} ({len(grup)} rafturi folosite)"
        )
    print(f"\nFișier salvat: {fisier_iesire}")
    return fisier_iesire


def start_interfata():
    import tkinter as tk
    from tkinter import filedialog, messagebox

    fereastra = tk.Tk()
    fereastra.title("Organizare rafturi bibliotecă")
    fereastra.geometry("620x220")
    fereastra.minsize(500, 200)

    fisier_selectat = tk.StringVar()
    stare = tk.StringVar(value="Selectează registrul Excel sau folosește-l pe cel detectat.")

    try:
        fisier_selectat.set(str(gaseste_fisier_excel()))
    except (FileNotFoundError, RuntimeError) as eroare:
        stare.set(str(eroare))

    cadru = tk.Frame(fereastra, padx=20, pady=20)
    cadru.pack(fill="both", expand=True)

    tk.Label(
        cadru,
        text="Alege registrul bibliotecii, apoi apasă Start:",
        anchor="w",
    ).pack(fill="x")

    rand_fisier = tk.Frame(cadru)
    rand_fisier.pack(fill="x", pady=(8, 14))
    tk.Entry(rand_fisier, textvariable=fisier_selectat).pack(
        side="left", fill="x", expand=True
    )

    def alege_fisier():
        cale = filedialog.askopenfilename(
            title="Alege registrul bibliotecii",
            initialdir=DIRECTORY,
            filetypes=[("Fișiere Excel", "*.xlsx"), ("Toate fișierele", "*.*")],
        )
        if cale:
            fisier_selectat.set(cale)
            stare.set("Registrul este pregătit pentru organizare.")

    tk.Button(rand_fisier, text="Alege fișier...", command=alege_fisier).pack(
        side="left", padx=(8, 0)
    )

    def porneste():
        cale = fisier_selectat.get().strip()
        if not cale or not Path(cale).is_file():
            messagebox.showerror(
                "Registru invalid",
                "Alege un fișier Excel existent înainte de a porni.",
                parent=fereastra,
            )
            return

        buton_start.config(state="disabled")
        stare.set("Se organizează registrul. Te rog așteaptă...")
        fereastra.update_idletasks()
        try:
            fisier_iesire = main(Path(cale))
        except ImportError as eroare:
            stare.set("Lipsesc biblioteci Python necesare.")
            messagebox.showerror(
                "Dependențe lipsă",
                "Instalează dependențele din requirements.txt în interpretorul "
                f"folosit pentru aplicație.\n\nDetalii: {eroare}",
                parent=fereastra,
            )
        except Exception as eroare:
            stare.set("Organizarea nu a reușit.")
            messagebox.showerror(
                "Eroare la organizare",
                str(eroare),
                parent=fereastra,
            )
        else:
            stare.set(f"Organizare terminată: {fisier_iesire.name}")
            messagebox.showinfo(
                "Gata",
                f"Planul de rafturi a fost salvat aici:\n{fisier_iesire}",
                parent=fereastra,
            )
        finally:
            buton_start.config(state="normal")

    buton_start = tk.Button(
        cadru,
        text="Start",
        command=porneste,
        bg="#1F4E78",
        fg="white",
        padx=24,
        pady=8,
    )
    buton_start.pack()
    tk.Label(cadru, textvariable=stare, anchor="w", wraplength=570).pack(
        fill="x", pady=(12, 0)
    )

    fereastra.mainloop()


if __name__ == "__main__":
    start_interfata()