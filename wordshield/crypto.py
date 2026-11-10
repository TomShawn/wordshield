"""仅在内存中序列化明文映射，磁盘上保存经过认证的密文。"""
from __future__ import annotations

import json
import os

from cryptography.exceptions import InvalidTag
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.scrypt import Scrypt

MAGIC = b"WORDSHIELD\x01"
MAX_MAP_SIZE = 16 * 1024 * 1024


class MappingError(ValueError):
    pass


def _key(password: str, salt: bytes) -> bytes:
    return Scrypt(salt=salt, length=32, n=2**15, r=8, p=1).derive(password.encode("utf-8"))


def encrypt_mapping(mapping: dict, password: str) -> bytes:
    if len(password) < 12:
        raise MappingError("映射表密码至少需要 12 个字符，请妥善离线保存。")
    salt, nonce = os.urandom(16), os.urandom(12)
    header = MAGIC + salt + nonce
    plain = json.dumps(mapping, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    if len(plain) > MAX_MAP_SIZE - 128:
        raise MappingError("映射表超出容量限制。")
    return header + AESGCM(_key(password, salt)).encrypt(nonce, plain, header)


def decrypt_mapping(data: bytes, password: str) -> dict:
    if len(data) > MAX_MAP_SIZE or len(data) < len(MAGIC) + 44 or not data.startswith(MAGIC):
        raise MappingError("映射表格式不支持或文件已损坏。")
    pos = len(MAGIC)
    salt, nonce = data[pos:pos + 16], data[pos + 16:pos + 28]
    header = data[:pos + 28]
    try:
        plain = AESGCM(_key(password, salt)).decrypt(nonce, data[pos + 28:], header)
        mapping = json.loads(plain)
    except (InvalidTag, ValueError, UnicodeError, RecursionError):
        raise MappingError("密码不正确，或映射表已损坏。") from None
    if not isinstance(mapping, dict) or mapping.get("version") != 1:
        raise MappingError("映射表版本不支持。")
    return mapping
