import argparse
import getpass
import json
import os
import re
import secrets
import tempfile
from functools import wraps
from pathlib import Path
from threading import Lock

import pandas as pd
from flask import Flask, jsonify, redirect, request, send_from_directory, session
from werkzeug.security import check_password_hash, generate_password_hash


DIRECTORY = Path(__file__).resolve().parent
DATA_DIRECTORY = DIRECTORY / "data"
SITE_DIRECTORY = DIRECTORY
EXCEL_FILE = DATA_DIRECTORY / "Inventar_Biblioteci_Compilat.xlsx"
SHELF_PLAN_FILE = DATA_DIRECTORY / "Inventar_Biblioteca_Organizat_Pe_Rafturi.xlsx"
SHEET_NAME = "Inventar Reunit"
ADMIN_CONFIG_FILE = Path.home() / ".biblioteca_admin.json"
DATA_LOCK = Lock()
ADMIN_CONFIG = None
EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

app = Flask(__name__, static_folder=None)
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("ADMIN_HTTPS") == "1",
    MAX_CONTENT_LENGTH=1024 * 1024,
)

FIELD_COLUMNS = {
    "titlu": ("Titlu",),
    "autor": ("Autor / Coordonator", "Autor"),
    "editura": ("Editură",),
    "an": ("An / Ediție", "Anul Apariției"),
    "raft": ("Raft",),
    "exemplare": ("Număr exemplare (adiacente)", "Număr Exemplare"),
    "observatii": ("Observații",),
}


def admin_required(handler):
    @wraps(handler)
    def wrapped(*args, **kwargs):
        if not session.get("admin_email"):
            return jsonify({"error": "Autentifică-te ca administrator."}), 401
        token = request.headers.get("X-CSRF-Token", "")
        saved_token = session.get("csrf_token", "")
        if not saved_token or not secrets.compare_digest(token, saved_token):
            return jsonify({"error": "Sesiune invalidă. Reautentifică-te."}), 403
        return handler(*args, **kwargs)
    return wrapped


def creeaza_cont_admin():
    if ADMIN_CONFIG_FILE.exists():
        raise FileExistsError(
            f"Există deja un cont în {ADMIN_CONFIG_FILE.name}; nu îl voi suprascrie."
        )

    email = input("E-mail administrator: ").strip().lower()
    if len(email) > 254 or not EMAIL_PATTERN.fullmatch(email):
        raise ValueError("Introdu o adresă de e-mail validă.")

    password = getpass.getpass("Parolă administrator (minimum 12 caractere): ")
    if len(password) < 12:
        raise ValueError("Parola trebuie să aibă cel puțin 12 caractere.")
    if password != getpass.getpass("Reintrodu parola: "):
        raise ValueError("Parolele nu coincid.")

    config = {
        "email": email,
        "password_hash": generate_password_hash(password),
        "secret_key": secrets.token_hex(32),
    }
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            prefix=".admin_credentials_",
            suffix=".tmp",
            dir=ADMIN_CONFIG_FILE.parent,
            delete=False,
        ) as temp_file:
            temp_path = Path(temp_file.name)
            json.dump(config, temp_file)
        os.replace(temp_path, ADMIN_CONFIG_FILE)
    finally:
        if temp_path and temp_path.exists():
            temp_path.unlink()
    print(f"Contul pentru «{email}» a fost creat. Parola nu este salvată în clar.")


def incarca_config_admin():
    global ADMIN_CONFIG
    if not ADMIN_CONFIG_FILE.is_file():
        return False
    with ADMIN_CONFIG_FILE.open("r", encoding="utf-8") as config_file:
        config = json.load(config_file)
    required = ("email", "password_hash", "secret_key")
    if any(not isinstance(config.get(key), str) or not config[key] for key in required):
        raise ValueError(f"Fișierul {ADMIN_CONFIG_FILE.name} nu are configurația necesară.")
    ADMIN_CONFIG = config
    app.secret_key = config["secret_key"]
    return True


def clasifica_sectiune(titlu):
    text = str(titlu).lower()
    if any(k in text for k in ["dicționar", "dictionnaire", "lexicon", "lexic", "glossary", "vocabulario", "concordance"]):
        return "Dicționare & Lexicoane"
    if any(k in text for k in ["latin", "cicero", "vergilius", "ovidiu", "tacit", "plaut", "seneca", "horatius", "suetonius", "caesar", "roma", "roman"]):
        return "Filologie Latină & Roma"
    if any(k in text for k in ["grec", "greek", "platon", "homer", "herodot", "bizan", "elen", "septuagint", "delphi", "oracl", "athens"]):
        return "Filologie Greacă & Bizanț"
    if any(k in text for k in ["biblia", "testament", "patristic", "părinți", "creștin", "teolog", "gnostic", "nyssa"]):
        return "Teologie & Patristică"
    if any(k in text for k in ["român", "dacia", "iași", "moldov", "cantemir", "rosetti", "densușianu"]):
        return "Studii Românești & Istorie"
    return "Altele"


def incarca_df():
    if not EXCEL_FILE.is_file():
        raise FileNotFoundError(f"Nu am găsit registrul Excel: {EXCEL_FILE}")

    with pd.ExcelFile(EXCEL_FILE) as xls:
        if not xls.sheet_names:
            raise ValueError(f"Registrul {EXCEL_FILE.name} nu conține foi de calcul.")
        sheet = SHEET_NAME if SHEET_NAME in xls.sheet_names else xls.sheet_names[0]
        df = xls.parse(sheet)
    if "Titlu" not in df.columns:
        raise ValueError(f"Foaia «{sheet}» trebuie să conțină coloana «Titlu».")

    for field, column_names in FIELD_COLUMNS.items():
        if not any(column in df.columns for column in column_names):
            df[column_names[0]] = "1" if field == "exemplare" else "-"
    return df


def coloana_camp(df, field):
    for column in FIELD_COLUMNS[field]:
        if column in df.columns:
            return column
    column = FIELD_COLUMNS[field][0]
    df[column] = "1" if field == "exemplare" else "-"
    return column


def text_celula(row, columns, default="-"):
    for column in columns:
        if column in row.index:
            value = row[column]
            if pd.notna(value):
                text = str(value).strip()
                return text if text else default
    return default


def urmatorul_inv(df):
    max_num = 0
    for column in ("Număr de Inventar", "Cod de Înregistrare"):
        if column not in df.columns:
            continue
        for value in df[column].dropna().astype(str):
            match = re.search(r"(\d+)", value)
            if match:
                max_num = max(max_num, int(match.group(1)))
    return f"INV-{max_num + 1:04d}"


def valoare_text(data, field, default="-"):
    value = data.get(field, default)
    if value is None:
        return default
    if isinstance(value, (dict, list, bool)):
        raise ValueError(f"Câmpul «{field}» trebuie să fie text.")
    text = str(value).strip()
    return text if text else default


def valideaza_date(data):
    titlu = valoare_text(data, "titlu", "")
    if not titlu:
        raise ValueError("Titlul este obligatoriu.")

    exemplare_text = valoare_text(data, "exemplare", "1")
    try:
        exemplare = int(exemplare_text)
    except ValueError as error:
        raise ValueError("Numărul de exemplare trebuie să fie un număr întreg pozitiv.") from error
    if exemplare < 1:
        raise ValueError("Numărul de exemplare trebuie să fie cel puțin 1.")

    return {
        "titlu": titlu,
        "autor": valoare_text(data, "autor"),
        "editura": valoare_text(data, "editura"),
        "an": valoare_text(data, "an"),
        "raft": valoare_text(data, "raft"),
        "exemplare": str(exemplare),
        "observatii": valoare_text(data, "observatii"),
    }


def salveaza_df(df):
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f"{EXCEL_FILE.stem}_",
            suffix=".tmp.xlsx",
            dir=DATA_DIRECTORY,
            delete=False,
        ) as temp_file:
            temp_path = Path(temp_file.name)

        with pd.ExcelWriter(temp_path, engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
        os.replace(temp_path, EXCEL_FILE)
    finally:
        if temp_path and temp_path.exists():
            temp_path.unlink()


def raspuns_eroare(mesaj, status):
    return jsonify({"error": mesaj}), status


@app.route("/")
@app.route("/catalog")
def catalog():
    return send_from_directory(SITE_DIRECTORY, "biblioteca.html")


@app.route("/login")
def login_page():
    if session.get("admin_email"):
        return redirect("/admin")
    return send_from_directory(SITE_DIRECTORY, "login.html")


@app.route("/admin")
def home():
    if not session.get("admin_email"):
        return redirect("/login")
    return send_from_directory(SITE_DIRECTORY, "admin.html")


@app.route("/api/login", methods=["POST"])
def autentifica_admin():
    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        return raspuns_eroare("Trimite numele și parola în format JSON.", 400)
    if ADMIN_CONFIG is None:
        return raspuns_eroare("Contul de administrator nu este configurat.", 503)

    email = data.get("email")
    password = data.get("password")
    if not isinstance(email, str) or not isinstance(password, str):
        return raspuns_eroare("E-mailul și parola sunt obligatorii.", 400)
    email = email.strip().lower()

    if (
        email != ADMIN_CONFIG["email"]
        or not check_password_hash(ADMIN_CONFIG["password_hash"], password)
    ):
        return raspuns_eroare("E-mail sau parolă incorecte.", 401)

    session.clear()
    session["admin_email"] = ADMIN_CONFIG["email"]
    session["csrf_token"] = secrets.token_urlsafe(32)
    return jsonify({
        "message": "Autentificare reușită.",
        "email": ADMIN_CONFIG["email"],
        "csrf_token": session["csrf_token"],
    })


@app.route("/api/session")
def stare_sesiune():
    return jsonify({
        "admin_available": True,
        "authenticated": bool(session.get("admin_email")),
    })


@app.route("/api/logout", methods=["POST"])
@admin_required
def deconecteaza_admin():
    session.clear()
    return jsonify({"message": "Ai fost deconectat."})


@app.route("/api/rafturi", methods=["POST"])
@admin_required
def genereaza_plan_rafturi():
    try:
        from scripts import organizeaza_rafturi

        fisier_plan = organizeaza_rafturi.main(EXCEL_FILE)
    except (FileNotFoundError, ImportError, OSError, RuntimeError, ValueError) as error:
        return raspuns_eroare(str(error), 500)
    return jsonify({
        "message": "Planul de rafturi a fost generat.",
        "filename": fisier_plan.name,
        "download_url": "/api/rafturi/plan",
    })


@app.route("/api/rafturi/plan")
def descarca_plan_rafturi():
    if not session.get("admin_email"):
        return raspuns_eroare("Autentifică-te ca administrator.", 401)
    if not SHELF_PLAN_FILE.is_file():
        return raspuns_eroare("Generează mai întâi planul de rafturi.", 404)
    return send_from_directory(
        DATA_DIRECTORY,
        SHELF_PLAN_FILE.name,
        as_attachment=True,
    )


@app.route("/api/carti", methods=["GET"])
def get_carti():
    try:
        df = incarca_df()
    except (FileNotFoundError, OSError, ValueError) as error:
        return raspuns_eroare(str(error), 500)

    rezultat = []
    for index, row in df.iterrows():
        cod = text_celula(row, ("Număr de Inventar", "Cod de Înregistrare"))
        titlu = text_celula(row, ("Titlu",))
        rezultat.append({
            "row_index": int(index),
            "nr_inv": cod,
            "titlu": titlu,
            "autor": text_celula(row, FIELD_COLUMNS["autor"]),
            "editura": text_celula(row, FIELD_COLUMNS["editura"]),
            "an": text_celula(row, FIELD_COLUMNS["an"]),
            "raft": text_celula(row, FIELD_COLUMNS["raft"]),
            "exemplare": text_celula(row, FIELD_COLUMNS["exemplare"], "1"),
            "observatii": text_celula(row, FIELD_COLUMNS["observatii"]),
            "sectiune": clasifica_sectiune(titlu),
        })
    return jsonify(rezultat)


@app.route("/api/carti", methods=["POST"])
@admin_required
def adauga_carte():
    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        return raspuns_eroare("Trimite datele cărții în format JSON.", 400)
    try:
        valori = valideaza_date(data)
    except ValueError as error:
        return raspuns_eroare(str(error), 400)

    with DATA_LOCK:
        try:
            df = incarca_df()
            nou_cod = urmatorul_inv(df)
            rand_nou = {column: "-" for column in df.columns}
            if "Nr. Crt. Global" in df.columns:
                rand_nou["Nr. Crt. Global"] = len(df) + 1
            if "Cod de Înregistrare" in df.columns:
                rand_nou["Cod de Înregistrare"] = nou_cod
            elif "Număr de Inventar" in df.columns:
                rand_nou["Număr de Inventar"] = nou_cod

            for field, value in valori.items():
                rand_nou[coloana_camp(df, field)] = value
            df = pd.concat([df, pd.DataFrame([rand_nou])], ignore_index=True)
            salveaza_df(df)
        except (FileNotFoundError, OSError, ValueError) as error:
            return raspuns_eroare(str(error), 500)

    return jsonify({"message": "Cartea a fost adăugată.", "nr_inv": nou_cod}), 201


@app.route("/api/carti/<int:row_index>", methods=["PUT"])
@admin_required
def editeaza_carte(row_index):
    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        return raspuns_eroare("Trimite datele cărții în format JSON.", 400)
    try:
        valori = valideaza_date(data)
    except ValueError as error:
        return raspuns_eroare(str(error), 400)

    with DATA_LOCK:
        try:
            df = incarca_df()
            if row_index < 0 or row_index >= len(df):
                return raspuns_eroare("Cartea solicitată nu există.", 404)
            for field, value in valori.items():
                column = coloana_camp(df, field)
                df[column] = df[column].astype("object")
                df.at[row_index, column] = value
            salveaza_df(df)
            cod = text_celula(
                df.iloc[row_index],
                ("Număr de Inventar", "Cod de Înregistrare"),
            )
        except (FileNotFoundError, OSError, ValueError) as error:
            return raspuns_eroare(str(error), 500)

    return jsonify({"message": "Modificările au fost salvate.", "nr_inv": cod})


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Server local pentru catalogul bibliotecii.")
    parser.add_argument(
        "--create-admin",
        action="store_true",
        help="creează contul inițial de administrator",
    )
    arguments = parser.parse_args()

    if arguments.create_admin or not ADMIN_CONFIG_FILE.exists():
        creeaza_cont_admin()
    if not incarca_config_admin():
        raise SystemExit("Nu s-a putut încărca configurația administratorului.")
    app.run(host="127.0.0.1", port=5000, debug=False)
