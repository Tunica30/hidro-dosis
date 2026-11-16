"""Cifra src/app.html con tu contraseña y genera index.html.

Uso:  python build.py
La contraseña se pide por teclado y no se guarda en ningún lado.
"""
import base64
import getpass
import json
import os
import sys
from pathlib import Path

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

ITER = 310_000
ROOT = Path(__file__).parent


def main():
    pw = getpass.getpass("Contraseña nueva para la app: ")
    if len(pw) < 6:
        sys.exit("Usá al menos 6 caracteres.")
    if getpass.getpass("Repetila: ") != pw:
        sys.exit("No coinciden. No se generó nada.")

    salt, iv = os.urandom(16), os.urandom(12)
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITER).derive(pw.encode("utf-8"))
    plain = (ROOT / "src" / "app.html").read_bytes()
    ct = AESGCM(key).encrypt(iv, plain, None)

    b = lambda x: base64.b64encode(x).decode("ascii")
    payload = json.dumps({"salt": b(salt), "iv": b(iv), "iter": ITER, "ct": b(ct)})
    gate = (ROOT / "src" / "gate.html").read_text(encoding="utf-8")
    (ROOT / "index.html").write_text(gate.replace("__PAYLOAD__", payload), encoding="utf-8", newline="\n")

    # Cambia la versión del caché para que el celular baje la nueva
    sw = ROOT / "sw.js"
    lines = sw.read_text(encoding="utf-8").splitlines()
    lines[0] = f"const VERSION = 'hidro-{b(os.urandom(6)).replace('/', '_').replace('+', '-')}';"
    sw.write_text("\n".join(lines) + "\n", encoding="utf-8", newline="\n")
    print("Listo: index.html generado y cifrado.")


if __name__ == "__main__":
    main()
