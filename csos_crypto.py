# -*- coding: utf-8 -*-
"""
csos_crypto.py — 고정 주소에 올리는 데이터를 잠근다 (설치 0 · 표준 라이브러리만)
================================================================================
고정 주소(GitHub Pages)는 **공개 저장소**다. 폰이 PC 없이 쓰려면 데이터를 거기 올려야
하는데, 그대로 올리면 캠프명·금액·기사 이름이 검색엔진에까지 걸린다. 그래서 잠가서 올린다.

  PC(파이썬)  : 잠근다   — 표준 라이브러리만 (외부 패키지 설치 없음이 이 프로젝트 원칙)
  폰(브라우저): 연다     — 브라우저 내장 WebCrypto(빠름). PIN을 넣으면 그 자리에서 풀린다.

방식: PBKDF2-HMAC-SHA256 → AES-256-CBC → HMAC-SHA256 (암호화 후 인증, encrypt-then-MAC)
  · 브라우저가 기본으로 지원하는 조합만 골랐다(WebCrypto: PBKDF2·AES-CBC·HMAC).
  · PIN이 4자리라 후보가 1만 개뿐이다. 그래서 **반복 60만 회**를 건다 —
    폰에서 여는 데는 1초 남짓이지만, 1만 개를 전부 대입하려면 같은 일을 1만 번 해야 한다.
  · AES는 여기 순수 파이썬으로 들어 있다. 올릴 때 한 번 돌리는 것이라 속도는 문제가 안 되고,
    맞게 짰는지는 FIPS-197 공식 시험벡터로 확인한다(self_test).

이 파일은 **암호를 담지 않는다**. PIN은 config/webapp.json(커밋 금지)에서 온다.
"""
import os, hmac, hashlib, base64, struct

ITERS = 600_000          # 폰에서 ~1초. 낮추면 PIN 전수대입이 그만큼 쉬워진다.
VERSION = 1

# ───────────────────────── AES-256 (FIPS-197) ─────────────────────────
_SBOX = bytes.fromhex(
    "637c777bf26b6fc53001672bfed7ab76ca82c97dfa5947f0add4a2af9ca472c0"
    "b7fd9326363ff7cc34a5e5f171d8311504c723c31896059a071280e2eb27b275"
    "09832c1a1b6e5aa0523bd6b329e32f8453d100ed20fcb15b6acbbe394a4c58cf"
    "d0efaafb434d338545f9027f503c9fa851a3408f929d38f5bcb6da2110fff3d2"
    "cd0c13ec5f974417c4a77e3d645d197360814fdc222a908846eeb814de5e0bdb"
    "e0323a0a4906245cc2d3ac629195e479e7c8376d8dd54ea96c56f4ea657aae08"
    "ba78252e1ca6b4c6e8dd741f4bbd8b8a703eb5664803f60e613557b986c11d9e"
    "e1f8981169d98e949b1e87e9ce5528df8ca1890dbfe6426841992d0fb054bb16")
_RCON = [0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80, 0x1B, 0x36,
         0x6C, 0xD8, 0xAB, 0x4D]


def _xt(a):
    """GF(2^8)에서 x를 한 번 곱한다(xtime)."""
    a <<= 1
    return (a ^ 0x11B) & 0xFF if a & 0x100 else a


def _expand(key):
    """AES-256 키 확장 → 15개 라운드키(각 16바이트). Nk=8, Nr=14."""
    nk, nr = 8, 14
    w = [list(key[4 * i:4 * i + 4]) for i in range(nk)]
    for i in range(nk, 4 * (nr + 1)):
        t = list(w[i - 1])
        if i % nk == 0:
            t = t[1:] + t[:1]
            t = [_SBOX[b] for b in t]
            t[0] ^= _RCON[i // nk - 1]
        elif i % nk == 4:
            t = [_SBOX[b] for b in t]
        w.append([w[i - nk][j] ^ t[j] for j in range(4)])
    return [bytes(b for word in w[4 * r:4 * r + 4] for b in word) for r in range(nr + 1)]


def _encrypt_block(rk, block):
    s = [b ^ k for b, k in zip(block, rk[0])]
    for rnd in range(1, len(rk)):
        s = [_SBOX[b] for b in s]
        # ShiftRows — 상태는 열 우선(column-major)이라 인덱스 r+4c 로 접근한다
        s = [s[(i + 4 * (i % 4)) % 16] for i in range(16)]
        if rnd != len(rk) - 1:
            out = []
            for c in range(4):
                col = s[4 * c:4 * c + 4]
                t = col[0] ^ col[1] ^ col[2] ^ col[3]
                out += [col[i] ^ t ^ _xt(col[i] ^ col[(i + 1) % 4]) for i in range(4)]
            s = out
        s = [b ^ k for b, k in zip(s, rk[rnd])]
    return bytes(s)


def _cbc_encrypt(key, iv, data):
    rk = _expand(key)
    pad = 16 - (len(data) % 16)
    data += bytes([pad]) * pad                      # PKCS#7
    out, prev = bytearray(), iv
    for i in range(0, len(data), 16):
        blk = bytes(a ^ b for a, b in zip(data[i:i + 16], prev))
        prev = _encrypt_block(rk, blk)
        out += prev
    return bytes(out)


# ───────────────────────── 봉인 ─────────────────────────
def derive(pin, salt, iters=ITERS):
    """PIN → 64바이트. 앞 32은 암호화 키, 뒤 32은 위·변조 확인용 키(용도 분리)."""
    return hashlib.pbkdf2_hmac("sha256", str(pin).encode(), salt, iters, 64)


def seal(plaintext, pin, iters=ITERS):
    """bytes → 폰이 열 수 있는 봉인 꾸러미(JSON으로 그대로 저장 가능한 dict)."""
    if isinstance(plaintext, str):
        plaintext = plaintext.encode("utf-8")
    salt, iv = os.urandom(16), os.urandom(16)
    k = derive(pin, salt, iters)
    ct = _cbc_encrypt(k[:32], iv, bytearray(plaintext))
    # 암호화한 **뒤에** 인증한다. 반대로 하면 위조된 데이터를 먼저 복호화하게 된다.
    tag = hmac.new(k[32:], iv + ct, hashlib.sha256).digest()
    b64 = lambda b: base64.b64encode(b).decode()
    return {"v": VERSION, "kdf": "PBKDF2-SHA256", "iter": iters, "cipher": "AES-256-CBC",
            "salt": b64(salt), "iv": b64(iv), "ct": b64(ct), "tag": b64(tag)}


# ───────────────────────── 자체 검증 ─────────────────────────
def self_test():
    """FIPS-197 부록 C.3(AES-256) 공식 시험벡터. 어긋나면 폰이 절대 못 연다."""
    key = bytes(range(32))
    pt = bytes.fromhex("00112233445566778899aabbccddeeff")
    want = "8ea2b7ca516745bfeafc49904b496089"
    got = _encrypt_block(_expand(key), pt).hex()
    assert got == want, f"AES-256 시험벡터 불일치: {got} != {want}"

    # CBC 왕복 — 복호화는 폰이 하므로 여기서는 '길이·패딩·태그' 규칙만 확인한다.
    # ★ 여기 PIN은 **시험용 0000**이다. 실 PIN을 적으면 공개 저장소에 그대로 남아
    #   올려 둔 사본의 잠금이 통째로 무의미해진다(config/webapp.json에만 둔다).
    s = seal(b"CSOS" * 40, "0000", iters=1000)
    ct = base64.b64decode(s["ct"])
    assert len(ct) % 16 == 0 and len(ct) > 160, len(ct)
    k = derive("0000", base64.b64decode(s["salt"]), 1000)
    assert hmac.compare_digest(
        base64.b64decode(s["tag"]),
        hmac.new(k[32:], base64.b64decode(s["iv"]) + ct, hashlib.sha256).digest())
    # 같은 평문이라도 매번 달라야 한다(salt·iv 재사용 금지)
    assert seal(b"x", "0000", 1000)["ct"] != seal(b"x", "0000", 1000)["ct"]
    return True


if __name__ == "__main__":
    print("AES-256 / PBKDF2 / HMAC 자체검증:", "통과" if self_test() else "실패")
