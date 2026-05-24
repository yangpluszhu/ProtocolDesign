"""Runtime decryption of bundled encrypted resources.

In development mode (not frozen), plaintext files are read from the resources/
directory directly. In release mode (PyInstaller --onefile), the encrypted blob
resources/resources.enc is decrypted in memory. Plaintext is never written to disk.
"""

from __future__ import annotations

import io
import os
import struct
import sys
from pathlib import Path

os.environ.setdefault("CRYPTOGRAPHY_OPENSSL_NO_LEGACY", "1")

from cryptography.hazmat.primitives.ciphers.aead import AESGCM


_BLOB_NAME = "resources.enc"
_MAGIC = b"PDRS0001"

_PLAINTEXT_RESOURCES_DIR = "resources"

_RESOURCE_NAMES = frozenset({
    "SKILL.md",
    "assets/corhortCRU-small.docx",
    "references/template-writing-guide.md",
})

# --- Key material (XOR-sharded, scattered) ---

_S1 = bytes([0x7c, 0x3a, 0x91, 0x58, 0xe4, 0x2d, 0xb6, 0x0f])
_M1 = bytes([0x21, 0xdf, 0x4e, 0xa3, 0x87, 0x6b, 0xc5, 0x92])

_S2 = bytes([0xd8, 0x05, 0x4f, 0xe2, 0xa1, 0x6c, 0x3b, 0x97])
_M2 = bytes([0x5a, 0xb3, 0xc7, 0x1e, 0x49, 0x80, 0xf2, 0x63])

_S3 = bytes([0x42, 0x8d, 0xf6, 0x13, 0x5e, 0xa9, 0x70, 0xcb])
_M3 = bytes([0x9f, 0x30, 0x68, 0xb5, 0xe3, 0x07, 0xda, 0x5c])

_S4 = bytes([0xb1, 0xe7, 0x29, 0x84, 0x6d, 0x53, 0xf8, 0x3c])
_M4 = bytes([0x46, 0x9a, 0x0b, 0x71, 0x2c, 0xd5, 0x84, 0xef])


def _derive_key() -> bytes:
    return (
        bytes(a ^ b for a, b in zip(_S1, _M1))
        + bytes(a ^ b for a, b in zip(_S2, _M2))
        + bytes(a ^ b for a, b in zip(_S3, _M3))
        + bytes(a ^ b for a, b in zip(_S4, _M4))
    )


_RESOURCE_KEY = _derive_key()

# --- Blob cache ---

_blob_cache: dict[str, tuple[bytes, bytes, bytes]] | None = None


def _is_frozen() -> bool:
    return getattr(sys, "frozen", False)


def _project_root() -> Path:
    if _is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def _blob_path() -> Path:
    if _is_frozen():
        base = Path(getattr(sys, "_MEIPASS", _project_root()))
        return base / _PLAINTEXT_RESOURCES_DIR / _BLOB_NAME
    return _project_root() / _PLAINTEXT_RESOURCES_DIR / _BLOB_NAME


def _plaintext_path(name: str) -> Path:
    return _project_root() / _PLAINTEXT_RESOURCES_DIR / name.replace("/", os.sep)


def _parse_blob() -> dict[str, tuple[bytes, bytes, bytes]]:
    global _blob_cache
    if _blob_cache is not None:
        return _blob_cache

    blob_file = _blob_path()
    data = blob_file.read_bytes()

    offset = 0
    magic = data[offset:offset + 8]
    offset += 8
    if magic != _MAGIC:
        raise ValueError("Resource blob has invalid magic header")

    (entry_count,) = struct.unpack_from(">I", data, offset)
    offset += 4

    entries: dict[str, tuple[bytes, bytes, bytes]] = {}
    for _ in range(entry_count):
        (name_len,) = struct.unpack_from(">H", data, offset)
        offset += 2
        name = data[offset:offset + name_len].decode("utf-8")
        offset += name_len
        iv = data[offset:offset + 12]
        offset += 12
        tag = data[offset:offset + 16]
        offset += 16
        (cipher_len,) = struct.unpack_from(">I", data, offset)
        offset += 4
        ciphertext = data[offset:offset + cipher_len]
        offset += cipher_len
        entries[name] = (iv, tag, ciphertext)

    _blob_cache = entries
    return entries


def _decrypt_entry(iv: bytes, tag: bytes, ciphertext: bytes) -> bytes:
    aesgcm = AESGCM(_RESOURCE_KEY)
    return aesgcm.decrypt(iv, ciphertext + tag, None)


def decrypt_resource_bytes(name: str) -> bytes:
    if not _is_frozen():
        path = _plaintext_path(name)
        if path.exists():
            return path.read_bytes()
        raise FileNotFoundError(f"Resource not found: {name}")

    entries = _parse_blob()
    if name not in entries:
        raise FileNotFoundError(f"Resource not found: {name}")
    iv, tag, ciphertext = entries[name]
    return _decrypt_entry(iv, tag, ciphertext)


def decrypt_resource_text(name: str) -> str:
    return decrypt_resource_bytes(name).decode("utf-8")


def decrypt_resource_bytesio(name: str) -> io.BytesIO:
    return io.BytesIO(decrypt_resource_bytes(name))
