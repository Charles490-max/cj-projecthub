"""
reports.pkg를 메모리에 로드하여 페이지 이미지 제공
v1: Fernet 암호화 (기존 호환)
v2: AES-CTR 스트리밍 암호화 (대용량 패키지 지원)
"""
import io
import os
import json
import struct
import hmac as hmac_mod
import hashlib
import zipfile
import base64
import tempfile
from pathlib import Path
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from PIL import Image

from config import MASTER_PASSWORD

CHUNK_SIZE = 64 * 1024 * 1024  # 64MB chunks for streaming


class SecurePackageLoader:
    def __init__(self, pkg_path):
        self.pkg_path = Path(pkg_path)
        self._zip = None
        self.manifest = None
        if self.pkg_path.exists():
            self._load()

    def _derive_key(self, password, salt):
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(), length=32,
            salt=salt, iterations=480000
        )
        return base64.urlsafe_b64encode(kdf.derive(password.encode()))

    def _derive_raw_key(self, password, salt):
        """AES-CTR용 32바이트 원시 키 (base64 인코딩 없음)"""
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(), length=32,
            salt=salt, iterations=480000
        )
        return kdf.derive(password.encode())

    def _load(self):
        with open(self.pkg_path, 'rb') as f:
            magic = f.read(8)

        if magic == b'CJPHv1\x00\x00':
            self._load_v1()
        elif magic == b'CJPHv2\x00\x00':
            self._load_v2()
        else:
            raise ValueError("유효하지 않은 패키지 파일입니다.")

    def _load_v1(self):
        """기존 Fernet 방식 (소용량 패키지)"""
        with open(self.pkg_path, 'rb') as f:
            f.read(8)  # magic
            salt = f.read(16)
            encrypted = f.read()

        key = self._derive_key(MASTER_PASSWORD, salt)
        cipher = Fernet(key)
        decrypted = cipher.decrypt(encrypted)

        self._zip = zipfile.ZipFile(io.BytesIO(decrypted), 'r')
        with self._zip.open('manifest.json') as f:
            self.manifest = json.load(f)

    def _load_v2(self):
        """AES-CTR 스트리밍 방식 (대용량 패키지)"""
        with open(self.pkg_path, 'rb') as f:
            f.read(8)   # magic
            salt = f.read(16)
            nonce = f.read(16)
            stored_mac = f.read(32)
            encrypted_data = f.read()

        # 키 파생
        raw_key = self._derive_raw_key(MASTER_PASSWORD, salt)

        # HMAC 검증
        computed_mac = hmac_mod.new(raw_key, encrypted_data, hashlib.sha256).digest()
        if not hmac_mod.compare_digest(stored_mac, computed_mac):
            raise ValueError("패키지 무결성 검증 실패 (HMAC 불일치)")

        # AES-CTR 복호화 → 임시 파일에 기록
        cipher = Cipher(algorithms.AES(raw_key), modes.CTR(nonce))
        decryptor = cipher.decryptor()

        decrypted_buffer = io.BytesIO()
        offset = 0
        while offset < len(encrypted_data):
            chunk = encrypted_data[offset:offset + CHUNK_SIZE]
            decrypted_buffer.write(decryptor.update(chunk))
            offset += CHUNK_SIZE
        decrypted_buffer.write(decryptor.finalize())
        decrypted_buffer.seek(0)

        self._zip = zipfile.ZipFile(decrypted_buffer, 'r')
        with self._zip.open('manifest.json') as f:
            self.manifest = json.load(f)

    def is_loaded(self):
        return self._zip is not None

    def has_report(self, site_code):
        if not self.manifest:
            return False
        return str(site_code) in self.manifest['reports']

    def get_page_count(self, site_code):
        if not self.has_report(site_code):
            return 0
        return self.manifest['reports'][str(site_code)]['page_count']

    def get_original_name(self, site_code):
        if not self.has_report(site_code):
            return None
        return self.manifest['reports'][str(site_code)].get('original_name', '')

    def get_page_image(self, site_code, page_no):
        site_code = str(site_code)
        if not self.has_report(site_code):
            return None

        for prefix in ['slide_', 'page_']:
            try:
                path = f"{site_code}/{prefix}{page_no:03d}.png"
                with self._zip.open(path) as f:
                    return Image.open(io.BytesIO(f.read())).copy()
            except KeyError:
                continue
        return None

    def search_text(self, keyword):
        """전체 보고서에서 키워드 검색"""
        if not self.manifest:
            return []

        results = []
        keyword_lower = keyword.lower().strip()
        if not keyword_lower:
            return []

        for site_code, info in self.manifest['reports'].items():
            for entry in info.get('text_index', []):
                content = entry.get('content', '')
                if keyword_lower in content.lower():
                    idx = content.lower().find(keyword_lower)
                    start = max(0, idx - 30)
                    end = min(len(content), idx + len(keyword) + 30)
                    snippet = content[start:end].replace('\n', ' ')
                    if start > 0:
                        snippet = '…' + snippet
                    if end < len(content):
                        snippet = snippet + '…'

                    results.append({
                        'site_code': site_code,
                        'page': entry['page'],
                        'snippet': snippet,
                        'original_name': info.get('original_name', '')
                    })

        return results
