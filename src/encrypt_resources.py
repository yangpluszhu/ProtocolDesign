"""Build-time encryption tool for ProtocolDesign resources.

Reads plaintext resource files from resources/ and writes a single encrypted
blob to resources/resources.enc using AES-256-GCM.

Usage:
    python src/encrypt_resources.py
"""

from __future__ import annotations

import os
import struct
import sys
from pathlib import Path

from cryptography.hazmat.primitives.ciphers.aead import AESGCM

ROOT = Path(__file__).resolve().parents[1]
RESOURCES_DIR = ROOT / "resources"

RESOURCES = [
    "SKILL.md",
    "assets/corhortCRU-small.docx",
    "references/template-writing-guide.md",
]


def main() -> None:
    sys.path.insert(0, str(ROOT / "src"))
    from resource_crypto import _RESOURCE_KEY, _MAGIC

    aesgcm = AESGCM(_RESOURCE_KEY)
    blob = bytearray(_MAGIC)
    blob += struct.pack(">I", len(RESOURCES))

    for name in RESOURCES:
        plaintext_path = RESOURCES_DIR / name.replace("/", os.sep)
        if not plaintext_path.exists():
            raise SystemExit(f"Resource not found: {plaintext_path}")
        plaintext = plaintext_path.read_bytes()
        iv = os.urandom(12)
        ct_and_tag = aesgcm.encrypt(iv, plaintext, None)
        ciphertext = ct_and_tag[:-16]
        tag = ct_and_tag[-16:]
        encoded_name = name.encode("utf-8")
        blob += struct.pack(">H", len(encoded_name))
        blob += encoded_name
        blob += iv
        blob += tag
        blob += struct.pack(">I", len(ciphertext))
        blob += ciphertext
        size_kb = len(plaintext) / 1024
        print(f"  Encrypted: {name} ({size_kb:.1f} KB)")

    output_path = RESOURCES_DIR / "resources.enc"
    output_path.write_bytes(bytes(blob))
    total_kb = len(blob) / 1024
    print(f"  Written: {output_path} ({total_kb:.1f} KB, {len(RESOURCES)} entries)")


if __name__ == "__main__":
    main()
