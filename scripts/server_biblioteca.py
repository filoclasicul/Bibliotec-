import csv
import json
import os
import secrets
from pathlib import Path
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

from werkzeug.security import check_password_hash

PROJECT_DIRECTORY = Path(__file__).resolve().parents[1]
DATA_DIRECTORY = PROJECT_DIRECTORY / "data"
SITE_DIRECTORY = PROJECT_DIRECTORY / "site"
CSV_FILE = DATA_DIRECTORY / "inventar.csv"
STATE_FILE = DATA_DIRECTORY / "biblioteca_stare.json"
ADMIN_CONFIG_FILE = Path.home() / ".biblioteca_admin.json"
ADMIN_CONFIG = None
ADMIN_EMAIL = ""
ACTIVE_TOKENS = set()


def incarca_config_admin() -> bool:
    global ADMIN_CONFIG, ADMIN_EMAIL
    if not ADMIN_CONFIG_FILE.is_file():
        return False
    with ADMIN_CONFIG_FILE.open("r", encoding="utf-8") as config_file:
        config = json.load(config_file)
    if any(
        not isinstance(config.get(key), str) or not config[key]
        for key in ("email", "password_hash", "secret_key")
    ):
        raise ValueError(f"Fișierul {ADMIN_CONFIG_FILE.name} nu este valid.")
    ADMIN_CONFIG = config
    ADMIN_EMAIL = config["email"]
    return True


def clasifica_sectiune(titlu: str) -> str:
    t = (titlu or "").lower()
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


def incarca_carti_din_csv():
    if not os.path.exists(CSV_FILE):
        return []
    with open(CSV_FILE, "r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(2048)
        f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;")
        except csv.Error:
            dialect = csv.excel
            dialect.delimiter = ";" if ";" in sample else ","

        reader = csv.DictReader(f, dialect=dialect)
        carti = []
        for i, row in enumerate(reader, start=1):
            titlu = (row.get("Titlu") or "-").strip() or "-"
            cod_inv = (row.get("Număr de Inventar") or row.get("Cod de Înregistrare") or f"INV-{i:04d}").strip() or f"INV-{i:04d}"
            autor = (row.get("Autor / Coordonator") or row.get("Autor") or "-").strip() or "-"
            editura = (row.get("Editură") or "-").strip() or "-"
            an = (row.get("An / Ediție") or row.get("Anul Apariției") or "-").strip() or "-"
            raft = (row.get("Raft") or "-").strip() or "-"
            ex_raw = (row.get("Număr exemplare (adiacente)") or row.get("Număr Exemplare") or "1").strip()
            try:
                ex_num = max(1, int(float(ex_raw)))
            except ValueError:
                ex_num = 1
            obs = (row.get("Observații") or "-").strip() or "-"

            carti.append({
                "nr_inv": cod_inv,
                "titlu": titlu,
                "autor": autor,
                "editura": editura,
                "an": an,
                "raft": raft,
                "exemplare": ex_num,
                "observatii": obs,
                "sectiune": clasifica_sectiune(titlu)
            })
        return carti


def incarca_stare():
    if not os.path.exists(STATE_FILE):
        return {"imprumuturi": []}
    with open(STATE_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    if "imprumuturi" not in data or not isinstance(data["imprumuturi"], list):
        data["imprumuturi"] = []
    return data


def salveaza_stare(stare):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(stare, f, ensure_ascii=False, indent=2)


def construieste_catalog():
    carti = incarca_carti_din_csv()
    stare = incarca_stare()
    imprumuturi_active = [x for x in stare["imprumuturi"] if not x.get("returnat", False)]

    imprumutate_pe_cod = {}
    for imp in imprumuturi_active:
        cod = imp.get("nr_inv", "")
        imprumutate_pe_cod.setdefault(cod, []).append(imp)

    rezultat = []
    for c in carti:
        cod = c["nr_inv"]
        active_pt_carte = imprumutate_pe_cod.get(cod, [])
        nr_imprumutate = len(active_pt_carte)
        disponibile = max(0, c["exemplare"] - nr_imprumutate)

        item = dict(c)
        item["imprumutate_count"] = nr_imprumutate
        item["disponibile_count"] = disponibile
        item["imprumuturi_active"] = active_pt_carte
        rezultat.append(item)

    return rezultat, imprumuturi_active


class BibliotecaHandler(SimpleHTTPRequestHandler):
    def _get_token(self):
        auth = self.headers.get("Authorization", "")
        if auth.startswith("Bearer "):
            return auth[7:].strip()
        return ""

    def _este_admin(self) -> bool:
        token = self._get_token()
        return bool(token and token in ACTIVE_TOKENS)

    def _send_json(self, payload, status=200):
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def _read_json_body(self):
        length = int(self.headers.get("Content-Length", 0))
        if length <= 0:
            return {}
        raw = self.rfile.read(length).decode("utf-8")
        return json.loads(raw)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/status":
            catalog, imprumuturi_active = construieste_catalog()
            self._send_json({
                "authenticated": self._este_admin(),
                "admin_email": ADMIN_EMAIL if self._este_admin() else None,
                "total_carti": len(catalog),
                "imprumuturi_active_count": len(imprumuturi_active)
            })
            return

        if parsed.path == "/api/carti":
            catalog, imprumuturi_active = construieste_catalog()
            self._send_json({
                "carti": catalog,
                "imprumuturi_active": imprumuturi_active
            })
            return

        if parsed.path == "/":
            self.path = "/biblioteca.html"
        return super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)

        if parsed.path == "/api/login":
            body = self._read_json_body()
            email = (body.get("email") or "").strip().lower()
            pwd = body.get("password") or ""
            if (
                ADMIN_CONFIG
                and email == ADMIN_EMAIL
                and isinstance(pwd, str)
                and check_password_hash(ADMIN_CONFIG["password_hash"], pwd)
            ):
                token = secrets.token_urlsafe(24)
                ACTIVE_TOKENS.add(token)
                self._send_json({"ok": True, "token": token, "email": ADMIN_EMAIL})
            else:
                self._send_json({"ok": False, "error": "Utilizator sau parolă incorecte."}, 401)
            return

        if parsed.path == "/api/logout":
            token = self._get_token()
            ACTIVE_TOKENS.discard(token)
            self._send_json({"ok": True})
            return

        if not self._este_admin():
            self._send_json({"ok": False, "error": "Trebuie să fii conectat ca administrator."}, 403)
            return

        if parsed.path == "/api/imprumuta":
            body = self._read_json_body()
            nr_inv = (body.get("nr_inv") or "").strip()
            persoana = (body.get("persoana") or "").strip()
            data_limita = (body.get("data_limita") or "-").strip() or "-"
            confirmat = bool(body.get("confirmat", False))

            if not nr_inv or not persoana:
                self._send_json({"ok": False, "error": "Selectează o carte și introdu numele persoanei."}, 400)
                return

            catalog, _ = construieste_catalog()
            carte = next((c for c in catalog if c["nr_inv"] == nr_inv), None)
            if not carte:
                self._send_json({"ok": False, "error": "Cartea nu a fost găsită în inventar."}, 404)
                return

            if carte["disponibile_count"] <= 0:
                self._send_json({
                    "ok": False,
                    "error": f"Cartea '{carte['titlu']}' ({nr_inv}) nu mai are exemplare disponibile pentru împrumut."
                }, 409)
                return

            if not confirmat:
                self._send_json({
                    "ok": True,
                    "needs_confirmation": True,
                    "preview": {
                        "nr_inv": carte["nr_inv"],
                        "titlu": carte["titlu"],
                        "autor": carte["autor"],
                        "persoana": persoana,
                        "data_limita": data_limita,
                        "disponibile_acum": carte["disponibile_count"],
                        "disponibile_dupa_imprumut": carte["disponibile_count"] - 1
                    }
                })
                return

            stare = incarca_stare()
            nou_id = f"IMP-{len(stare['imprumuturi']) + 1:04d}"
            inregistrare = {
                "id": nou_id,
                "nr_inv": carte["nr_inv"],
                "titlu": carte["titlu"],
                "autor": carte["autor"],
                "persoana": persoana,
                "data_limita": data_limita,
                "returnat": False
            }
            stare["imprumuturi"].append(inregistrare)
            salveaza_stare(stare)

            self._send_json({"ok": True, "needs_confirmation": False, "imprumut": inregistrare})
            return

        if parsed.path == "/api/returneaza":
            body = self._read_json_body()
            imp_id = (body.get("id") or "").strip()
            if not imp_id:
                self._send_json({"ok": False, "error": "Lipsește ID-ul împrumutului."}, 400)
                return

            stare = incarca_stare()
            gasit = False
            for item in stare["imprumuturi"]:
                if item.get("id") == imp_id and not item.get("returnat", False):
                    item["returnat"] = True
                    gasit = True
                    break

            if not gasit:
                self._send_json({"ok": False, "error": "Împrumutul nu a fost găsit sau este deja returnat."}, 404)
                return

            salveaza_stare(stare)
            self._send_json({"ok": True, "id": imp_id})
            return

        self._send_json({"ok": False, "error": "Rută inexistentă."}, 404)


if __name__ == "__main__":
    os.chdir(SITE_DIRECTORY)
    if not incarca_config_admin():
        raise SystemExit(
            "Contul administratorului nu există. Creează-l rulând app_admin.py "
            "și urmează pașii de configurare."
        )
    server = ThreadingHTTPServer(("127.0.0.1", 5000), BibliotecaHandler)
    print("Server pornit pe http://127.0.0.1:5000")
    print(f"Cont administrator: {ADMIN_EMAIL}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer oprit.")