"""Криптографическое ядро Shifr.

Шифр сохраняет длину и «форму» текста: каждая буква заменяется буквой того же
алфавита и регистра, цифра — цифрой, остальные символы (пробелы, пунктуация)
остаются на месте. Ключевой поток — AES-256-CTR, своя последовательность на
каждый фрагмент (номер фрагмента в блоке счётчика). Целостность — усечённый
HMAC-SHA256 открытого текста.

Зашифрованный фрагмент на слайде выглядит так:  ⟦<номер><контроль>·<шифротекст>⟧
"""
from __future__ import annotations

import hashlib
import hmac
import json
import os
import re
import secrets
from dataclasses import dataclass

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.hkdf import HKDF

# --- алфавиты, внутри которых идёт замена ---------------------------------
_RU_UP = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ"
ALPHABETS = (
    _RU_UP,
    _RU_UP.lower(),
    "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
    "abcdefghijklmnopqrstuvwxyz",
    "0123456789",
)
_CHAR_CLASS: dict[str, tuple[str, int]] = {
    ch: (alpha, i) for alpha in ALPHABETS for i, ch in enumerate(alpha)
}

# --- служебные символы ------------------------------------------------------
OPEN, CLOSE, SEP = "⟦", "⟧", "·"
# base32 без похожих символов (нет 0, 1, O, I)
B32 = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ"
CHECK_LEN = 3  # 15 бит контроля на фрагмент
KEY_PREFIX = "SHIFR1"
FORBIDDEN_CHARS = (OPEN, CLOSE)

# Метка: ⟦ + номер(1+) и контроль(3) + разделитель + шифротекст + ⟧.
# Разделитель допускаем в нескольких вариантах на случай «нормализации».
MARKER_RE = re.compile(
    "%s([%s%s]{%d,14})[·•∙:|]([^%s%s]*)%s"
    % (OPEN, B32, B32.lower(), CHECK_LEN + 1, OPEN, CLOSE, CLOSE)
)


class KeyFormatError(ValueError):
    """Строка ключа набрана с ошибкой."""


def _b32_encode_int(n: int, width: int | None = None) -> str:
    out = ""
    while n:
        n, r = divmod(n, 32)
        out = B32[r] + out
    out = out or B32[0]
    if width is not None:
        out = out.rjust(width, B32[0])
    return out


def _b32_decode(s: str) -> int:
    n = 0
    for ch in s.upper():
        idx = B32.find(ch)
        if idx < 0:
            raise ValueError(ch)
        n = n * 32 + idx
    return n


# --- ключ -------------------------------------------------------------------
@dataclass(frozen=True)
class Key:
    master: bytes

    @classmethod
    def generate(cls) -> "Key":
        return cls(secrets.token_bytes(32))

    # производные ключи
    def _derive(self, info: bytes) -> bytes:
        return HKDF(algorithm=hashes.SHA256(), length=32, salt=None, info=info).derive(self.master)

    @property
    def stream_key(self) -> bytes:
        return self._derive(b"shifr1 stream")

    @property
    def mac_key(self) -> bytes:
        return self._derive(b"shifr1 mac")

    @property
    def vault_key(self) -> bytes:
        return self._derive(b"shifr1 vault")

    # строковое представление
    def to_string(self) -> str:
        body = _b32_encode_int(int.from_bytes(self.master, "big"), 52)
        chk = _b32_encode_int(int.from_bytes(hashlib.sha256(self.master).digest()[:2], "big") & 0x3FF, 2)
        s = body + chk  # 54 символа
        groups = [s[i : i + 6] for i in range(0, len(s), 6)]
        return KEY_PREFIX + "-" + "-".join(groups)

    @classmethod
    def from_string(cls, text: str) -> "Key":
        s = re.sub(r"[\s\-_]", "", text or "").upper()
        if s.startswith(KEY_PREFIX):
            s = s[len(KEY_PREFIX):]
        if len(s) != 54:
            raise KeyFormatError("Неверная длина ключа: проверьте, что строка скопирована целиком.")
        try:
            n = _b32_decode(s[:52])
            chk = _b32_decode(s[52:])
        except ValueError as e:
            raise KeyFormatError(f"Недопустимый символ в ключе: «{e.args[0]}».") from None
        if n >> 256:
            raise KeyFormatError("Ключ набран с ошибкой.")
        master = n.to_bytes(32, "big")
        if (int.from_bytes(hashlib.sha256(master).digest()[:2], "big") & 0x3FF) != chk:
            raise KeyFormatError("Ключ набран с ошибкой (не сходится контрольная сумма).")
        return cls(master)


# --- шифрование фрагмента ----------------------------------------------------
class _Keystream:
    """Равномерные числа 0..N-1 из AES-256-CTR (с отбрасыванием для отсутствия смещения)."""

    def __init__(self, key: bytes, frag_id: int):
        iv = frag_id.to_bytes(8, "big") + b"\x00" * 8
        self._enc = Cipher(algorithms.AES(key), modes.CTR(iv)).encryptor()
        self._buf = b""
        self._pos = 0

    def _byte(self) -> int:
        if self._pos >= len(self._buf):
            self._buf = self._enc.update(b"\x00" * 256)
            self._pos = 0
        b = self._buf[self._pos]
        self._pos += 1
        return b

    def below(self, n: int) -> int:
        limit = 256 - (256 % n)
        while True:
            b = self._byte()
            if b < limit:
                return b % n


def _transform(key: Key, frag_id: int, text: str, sign: int) -> str:
    ks = _Keystream(key.stream_key, frag_id)
    out = []
    for ch in text:
        cls = _CHAR_CLASS.get(ch)
        if cls is None:
            out.append(ch)
            continue
        alpha, idx = cls
        out.append(alpha[(idx + sign * ks.below(len(alpha))) % len(alpha)])
    return "".join(out)


def _check(key: Key, frag_id: int, plaintext: str) -> str:
    mac = hmac.new(key.mac_key, frag_id.to_bytes(8, "big") + plaintext.encode("utf-8"), hashlib.sha256).digest()
    return _b32_encode_int(int.from_bytes(mac[:2], "big") >> (16 - 5 * CHECK_LEN), CHECK_LEN)


def marker_parts(key: Key, frag_id: int, plaintext: str) -> tuple[str, str, str]:
    """Возвращает (префикс, шифротекст, суффикс). Шифротекст той же длины, что и текст."""
    for ch in FORBIDDEN_CHARS:
        if ch in plaintext:
            raise ValueError(f"Текст содержит служебный символ «{ch}».")
    head = OPEN + _b32_encode_int(frag_id) + _check(key, frag_id, plaintext) + SEP
    return head, _transform(key, frag_id, plaintext, +1), CLOSE


def encrypt_fragment(key: Key, frag_id: int, plaintext: str) -> str:
    return "".join(marker_parts(key, frag_id, plaintext))


@dataclass
class Decrypted:
    frag_id: int
    text: str  # расшифровка (при ok=False — «как получилось»)
    ok: bool


def parse_header(header: str) -> tuple[int, str]:
    header = header.upper()
    return _b32_decode(header[:-CHECK_LEN]), header[-CHECK_LEN:]


def decrypt_marker(key: Key, header: str, body: str) -> Decrypted:
    frag_id, chk = parse_header(header)
    plain = _transform(key, frag_id, body, -1)
    return Decrypted(frag_id, plain, hmac.compare_digest(_check(key, frag_id, plain), chk))


# --- резервная копия (.shifr) ------------------------------------------------
VAULT_MAGIC = b"SHIFRV1\n"


def vault_dump(key: Key, fragments: dict[int, dict]) -> bytes:
    data = json.dumps({str(k): v for k, v in fragments.items()}, ensure_ascii=False).encode("utf-8")
    nonce = os.urandom(12)
    return VAULT_MAGIC + nonce + AESGCM(key.vault_key).encrypt(nonce, data, VAULT_MAGIC)


def vault_load(key: Key, blob: bytes) -> dict[int, dict]:
    if not blob.startswith(VAULT_MAGIC):
        raise ValueError("Это не файл резервной копии Shifr.")
    nonce = blob[len(VAULT_MAGIC) : len(VAULT_MAGIC) + 12]
    try:
        data = AESGCM(key.vault_key).decrypt(nonce, blob[len(VAULT_MAGIC) + 12 :], VAULT_MAGIC)
    except Exception:
        raise ValueError("Резервная копия не подходит к этому ключу или повреждена.") from None
    return {int(k): v for k, v in json.loads(data).items()}
