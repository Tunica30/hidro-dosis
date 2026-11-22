"""Cifra src/app.html con tu contraseña y genera index.html.

Uso:
  python build.py            La primera vez pide las contraseñas. Después reutiliza las claves guardadas.
  python build.py --nueva    Pide contraseñas nuevas (para cambiarlas).
  python build.py --url URL  Guarda la dirección de la planilla de Google y regenera.

Las contraseñas no se guardan. Se guarda la clave derivada en .build-secrets.json
(que nunca se sube a GitHub). Mientras la clave no cambie, los celulares no tienen
que volver a poner la contraseña después de una actualización.
"""
import base64
import getpass
import hashlib
import json
import os
import sys
from pathlib import Path

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

ITER = 310_000
ROOT = Path(__file__).parent
SECRETS = ROOT / ".build-secrets.json"

b64 = lambda x: base64.b64encode(x).decode("ascii")


def ask_passwords():
    pw = getpass.getpass("Contraseña de la app (la que usan todos): ")
    if len(pw) < 6:
        sys.exit("Usá al menos 6 caracteres.")
    if getpass.getpass("Repetila: ") != pw:
        sys.exit("No coinciden. No se generó nada.")
    print("\nAhora la contraseña de AJUSTES (solo para vos, distinta de la anterior).")
    admin = getpass.getpass("Contraseña de ajustes: ")
    if len(admin) < 4:
        sys.exit("Usá al menos 4 caracteres.")
    if getpass.getpass("Repetila: ") != admin:
        sys.exit("No coinciden. No se generó nada.")
    if admin == pw:
        sys.exit("Tiene que ser distinta de la contraseña de la app. No se generó nada.")

    salt = os.urandom(16)
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(pw.encode("utf-8"))
    admin_salt = os.urandom(16).hex()
    admin_hash = hashlib.sha256((admin_salt + admin).encode("utf-8")).hexdigest()
    return {"salt": b64(salt), "key": b64(key), "admin_salt": admin_salt, "admin_hash": admin_hash}


def main():
    args = sys.argv[1:]
    secrets = json.loads(SECRETS.read_text(encoding="utf-8")) if SECRETS.exists() else {}

    if "--url" in args:
        url = args[args.index("--url") + 1].strip()
        if not url.startswith("https://script.google.com/"):
            sys.exit("La dirección tiene que empezar con https://script.google.com/")
        secrets["log_url"] = url

    if "--nueva" in args or "key" not in secrets:
        secrets.update(ask_passwords())

    SECRETS.write_text(json.dumps(secrets, indent=2), encoding="utf-8")

    app = (ROOT / "src" / "app.html").read_text(encoding="utf-8")
    for ph in ("__ADMIN__", "__LOG_URL__"):
        if ph not in app:
            sys.exit(f"Falta {ph} en src/app.html.")
    # La misma versión va en sw.js y adentro de la app: así la app sabe con certeza si quedó vieja
    version = f"hidro-{os.urandom(6).hex()}"
    app = app.replace("__BUILD__", version)
    # Cuándo se armó esta versión (se ve chiquito abajo a la derecha de la app)
    from datetime import datetime
    app = app.replace("__BUILT__", datetime.now().strftime("%d/%m/%Y %H:%M"))
    app = app.replace("__ADMIN__", json.dumps({"salt": secrets["admin_salt"], "hash": secrets["admin_hash"]}))
    app = app.replace("__LOG_URL__", secrets.get("log_url", ""))

    iv = os.urandom(12)
    ct = AESGCM(base64.b64decode(secrets["key"])).encrypt(iv, app.encode("utf-8"), None)
    payload = json.dumps({"salt": secrets["salt"], "iv": b64(iv), "iter": ITER, "ct": b64(ct)})
    gate = (ROOT / "src" / "gate.html").read_text(encoding="utf-8")
    (ROOT / "index.html").write_text(gate.replace("__PAYLOAD__", payload), encoding="utf-8", newline="\n")

    # Cambia la versión del caché para que el celular baje la nueva
    sw = ROOT / "sw.js"
    lines = sw.read_text(encoding="utf-8").splitlines()
    lines[0] = f"const VERSION = '{version}';"
    sw.write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")
    print("Listo: index.html generado y cifrado." + ("" if secrets.get("log_url") else " (Sin planilla conectada.)"))


if __name__ == "__main__":
    main()
