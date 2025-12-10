from __future__ import annotations
import io, zlib, struct
from typing import List, Tuple
import olefile

from server.core.normalize import normalization_text
from server.core.matching import find_sensitive_spans  

TAG_PARA_TEXT = 67
TAG_PICTURE = 0x04F
ENDOFCHAIN = 0xFFFFFFFE

# ─────────────────────────────
# 압축 관련 유틸리티
# ─────────────────────────────
def _decompress(raw: bytes) -> Tuple[bytes, int]:
    for w in (-15, +15):
        try:
            return zlib.decompress(raw, w), w
        except zlib.error:
            pass
    return raw, 0

def _recompress(buf: bytes, mode: int) -> bytes:
    if mode == 0:
        return buf
    c = zlib.compressobj(level=9, wbits=mode)
    return c.compress(buf) + c.flush()

def decomp_bin(raw: bytes, off: int, kind: str):
    data = raw[off:]
    try:
        if kind == "zlib":
            obj = zlib.decompressobj()
            dec = obj.decompress(data)
            consumed = len(data) - len(obj.unused_data)
        elif kind == "gzip":
            obj = zlib.decompressobj(16 + zlib.MAX_WBITS)
            dec = obj.decompress(data)
            consumed = len(data) - len(obj.unused_data)
        else:
            obj = zlib.decompressobj(-15)
            dec = obj.decompress(data)
            consumed = len(data) - len(obj.unused_data)
        if consumed <= 0 or len(dec) == 0:
            return None
        return dec, consumed
    except Exception:
        return None

def recomp_bin(kind: str, dec: bytes):
    if kind == "zlib":
        return zlib.compress(dec)
    if kind == "rawdef":
        co = zlib.compressobj(level=6, wbits=-15)
        return co.compress(dec) + co.flush()
    return None


# ─────────────────────────────
# OLE 내부 구조 유틸리티
# ─────────────────────────────

# HWP 파일 내부에서 section / BinData 스트림 찾기
def _direntry_for(ole: olefile.OleFileIO, path: Tuple[str, ...]):
    try:
        sid_or_entry = ole._find(path)
        if isinstance(sid_or_entry, int):
            i = sid_or_entry
            return ole.direntries[i] if 0 <= i < len(ole.direntries) else None
        return sid_or_entry
    except Exception:
        return None

# 섹션 레코드 파서
def iter_hwp_records(section_bytes: bytes):
    off = 0
    n = len(section_bytes)
    while off + 4 <= n:
        hdr = int.from_bytes(section_bytes[off:off+4], "little")
        tag   =  hdr        & 0x3FF
        level = (hdr >> 10) & 0x3FF
        size  = (hdr >> 20) & 0xFFF
        rec_start = off
        off += 4
        if size == 0xFFF:
            if off + 4 > n:
                break
            size = int.from_bytes(section_bytes[off:off+4], "little")
            off += 4
        if off + size > n:
            payload = section_bytes[off:n]
            yield tag, level, payload, rec_start, n
            break
        payload = section_bytes[off: off+size]
        rec_end = off + size
        yield tag, level, payload, rec_start, rec_end
        off = rec_end


# 미니스트림 오프셋 계산 (4096B 이하)
def _collect_ministream_offsets(ole: olefile.OleFileIO) -> List[int]:
    root = getattr(ole, "root", None)
    if root is None:
        return []
    sec_size = ole.sector_size
    fat = ole.fat
    s = root.isectStart
    out: List[int] = []
    while s not in (-1, olefile.ENDOFCHAIN) and 0 <= s < len(fat):
        out.append((s + 1) * sec_size) # 헤더 사이즈때문에 + 1, 해당 섹터의 offset 계산
        s = fat[s] # 다음 섹터로 이동
        if len(out) > 65536:
            break
    return out

# FAT (4096B 이상)에 바뀐 본문 덮어쓰기
def _overwrite_bigfat(ole: olefile.OleFileIO, container: bytearray, start_sector: int, new_raw: bytes) -> int:
    sec_size = ole.sector_size
    fat = ole.fat
    s = start_sector
    pos = wrote = 0
    while s not in (-1, olefile.ENDOFCHAIN) and 0 <= s < len(fat) and pos < len(new_raw):
        off = (s + 1) * sec_size
        chunk = new_raw[pos : pos + sec_size]
        container[off : off + len(chunk)] = chunk
        pos += len(chunk)
        wrote += len(chunk)
        s = fat[s]
    return wrote

# miniFAT(4096B 이하)에 바뀐 덮어쓰기
def _overwrite_minifat_chain(ole: olefile.OleFileIO, container: bytearray, mini_start: int, new_raw: bytes) -> int:
    ole.loadminifat()
    mini_size = ole.mini_sector_size
    minifat = getattr(ole, "minifat", [])
    ministream_offsets = _collect_ministream_offsets(ole)
    if not ministream_offsets or not minifat:
        return 0
    pos = wrote = 0
    s = mini_start
    while s not in (-1, olefile.ENDOFCHAIN) and 0 <= s < len(minifat) and pos < len(new_raw):
        mini_off = s * mini_size
        big_index = mini_off // ole.sector_size
        within = mini_off % ole.sector_size
        if big_index >= len(ministream_offsets):
            break
        file_off = ministream_offsets[big_index] + within
        chunk = new_raw[pos : pos + mini_size]
        container[file_off : file_off + len(chunk)] = chunk
        pos += len(chunk)
        wrote += len(chunk)
        s = minifat[s]
        if wrote > 64 * 1024 * 1024:
            break
    return wrote

def find_direntry_tail(ole, tail: str):
    for e in ole.direntries:
        if e and e.name == tail:
            return e
    return None



# ─────────────────────────────
# BINDATA 관련 함수
# ─────────────────────────────
HWPTAG_CTRL_HEADER = 0x0010
HWPTAG_CTRL_DATA   = 0x0011
MAKE_4CHID = lambda a,b,c,d: (a | (b<<8) | (c<<16) | (d<<24))
CTRLID_OLE = MAKE_4CHID(ord('$'), ord('o'), ord('l'), ord('e'))

CFB = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
PNG = b"\x89PNG\r\n\x1a\n"
GZ  = b"\x1F\x8B"
JPG = b"\xFF\xD8\xFF"
WMF = b"\xD7\xCD\xC6\x9A"

def parse_ctrl_header(payload: bytes):
    if len(payload) < 4:
        return None
    return int.from_bytes(payload[:4], "little")

def parse_bindata_id_from_ctrldata(payload: bytes):
    if len(payload) < 4:
        return None
    return int.from_bytes(payload[:4], "little")

# BodyText/Section 스트림에서 $ole 컨트롤이 참조하는 BinDataID 리스트 반환
def discover_ole_bindata_ids_strict(section_bytes: bytes):
    ids = []
    pending = None
    for tag, level, payload, rs, re in iter_hwp_records(section_bytes):
        if tag == HWPTAG_CTRL_HEADER:
            ctrlid = parse_ctrl_header(payload)
            pending = level if ctrlid == CTRLID_OLE else None
        elif pending is not None:
            if tag == HWPTAG_CTRL_DATA and level == pending:
                bid = parse_bindata_id_from_ctrldata(payload)
                if bid is not None:
                    ids.append(bid)
                pending = None
            elif level < pending:
                pending = None
    return ids

def magic_hits(raw: bytes):
    hits = []
    if raw.startswith(CFB): hits.append(("ole", 0))
    if raw.startswith(PNG): hits.append(("png", 0))
    if raw.startswith(GZ):  hits.append(("gzip", 0))
    if raw.startswith(JPG): hits.append(("jpeg", 0))
    if raw.startswith(WMF): hits.append(("wmf", 0))
    for sig, name in [(CFB, "ole"), (PNG, "png"), (GZ, "gzip")]:
        off = raw.find(sig, 1)
        if off != -1:
            hits.append((name, off))
    return hits

def is_zlib_head(b: bytes):
    return len(b) >= 2 and b[0] == 0x78 and b[1] in (0x01, 0x9C, 0xDA)

def scan_deflate(raw: bytes, limit: int = 64, step: int = 64):
    n = len(raw)
    cand = []
    for i in range(n - 1):
        if is_zlib_head(raw[i:i + 2]):
            cand.append(("zlib", i))
        if raw[i:i + 2] == GZ:
            cand.append(("gzip", i))
    for i in range(0, n, step):
        cand.append(("rawdef", i))
    out = []
    seen = set()
    for k, o in cand:
        if (k, o) in seen:
            continue
        seen.add((k, o))
        out.append((k, o))
        if len(out) >= limit:
            break
    return out

def patch_seg(raw: bytes, off: int, consumed: int, new_comp: bytes):
    seg = raw[off:off + consumed]
    if len(new_comp) > len(seg):
        return None
    if len(new_comp) < len(seg):
        new_comp = new_comp + b"\x00" * (len(seg) - len(new_comp))
    return raw[:off] + new_comp + raw[off + len(seg):]


# ─────────────────────────────
# 문자 추출 관련 함수
# ─────────────────────────────

# 파라미터 반복자 (본문 텍스트)
def _iter_para_text_records(section_dec: bytes):
    off, n = 0, len(section_dec)
    while off + 4 <= n:
        hdr = struct.unpack_from("<I", section_dec, off)[0]
        tag = hdr & 0x3FF
        size = (hdr >> 20) & 0xFFF
        off += 4
        if size < 0 or off + size > n:
            break
        payload = section_dec[off : off + size]
        if tag == TAG_PARA_TEXT:
            yield payload
        off += size

def _extract_bindata(raw:bytes) -> List[str]:
    out = []

    #  CFB (OLE Compound File)
    if raw.startswith(b"\xD0\xCF\x11\xE0"):
        try:
            sub = olefile.OleFileIO(io.BytesIO(raw))
            for p in sub.listdir(streams=True, storages=False):
                try:
                    sraw = sub.openstream(p).read()
                    for enc in ("utf-16le", "utf-8", "cp949"):
                        try:
                            txt = sraw.decode(enc)
                            out.append(txt)
                            break
                        except:
                            continue
                except:
                    continue
        except:
            pass

    #  ZLIB / raw deflate 블록 탐지
    for i in range(len(raw) - 2):
        # zlib 헤더 signature
        if raw[i] == 0x78:
            try:
                dec = zlib.decompress(raw[i:])
                for enc in ("utf-16le", "utf-8", "cp949"):
                    try:
                        txt = dec.decode(enc)
                        out.append(txt)
                    except:
                        continue
            except:
                continue

    return out



def extract_text(file_bytes: bytes) -> dict:
    texts: List[str] = []
    with olefile.OleFileIO(io.BytesIO(file_bytes)) as ole:
        for path in ole.listdir(streams=True, storages=False):

            # BodyText/Section
            if len(path) >= 2 and path[0] == "BodyText" and path[1].startswith("Section"):
                raw = ole.openstream(path).read()
                dec, _ = _decompress(raw)
                for payload in _iter_para_text_records(dec):
                    try:
                        texts.append(payload.decode("utf-16le", "ignore"))
                    except:
                        pass
                continue

            # BinData/*.OLE
            if len(path) >= 2 and path[0] == "BinData" and path[1].endswith(".OLE"):
                raw = ole.openstream(path).read()
                texts.extend(_extract_bindata(raw))
                continue

    full = "\n".join(texts)
    return {"full_text": full, "pages": [{"page": 1, "text": full}]}



# 정규식에 걸리는 문자열 collect
def _collect_targets_by_regex(text: str) -> List[str]:
    res = find_sensitive_spans(text) 
    targets: List[str] = []
    for s, e, _val, _rule in res:
        if isinstance(s, int) and isinstance(e, int) and e > s:
            frag = text[s:e]
            if frag and len(frag.strip()) >= 1:
                targets.append(frag)
    targets = sorted(set(targets), key=lambda x: (-len(x), x))
    return targets

# ParaText 전용 replace 함수 (utf-16 고정)
def _replace_utf16le_keep_len(buf: bytes, t: str) -> Tuple[bytes, int]:
    if not t:
        return buf, 0

    pat = t.encode("utf-16le", "ignore")

    rep_str = ""
    for ch in t:
        if ch == "-":
            rep_str += "-"
        else:
            rep_str += "*"

    rep = rep_str.encode("utf-16le")
    count = buf.count(pat)
    if count:
        buf = buf.replace(pat, rep)
    return buf, count

# BinData 내부 전용 replace 함수
def _replace_in_bindata(raw: bytes, t: str) -> Tuple[bytes, int]:
    total = 0
    out = raw
    for enc in ("utf-16le", "utf-8", "cp949"):
        try:
            pat = t.encode(enc, "ignore")
            if not pat:
                continue

            rep_str = ""
            for ch in t:
                if ch == "-":
                    rep_str += "-"
                else:
                    rep_str += "*"

            rep = (rep_str.encode("utf-16le") if enc == "utf-16le" else b"*" * len(pat))
            hits = out.count(pat)
            if hits:
                out = out.replace(pat, rep)
                total += hits
        except Exception:
            pass
    return out, total

# ─────────────────────────────
# 이미지 추출 관련 함수
# ─────────────────────────────

# 이미지 레코드 찾기 반복자
def _iter_picture_records(section_dec: bytes):
    off, n = 0, len(section_dec)
    while off + 4 <= n:
        hdr = struct.unpack_from("<I", section_dec, off)[0]
        tag = hdr & 0x3FF           
        size = (hdr >> 20) & 0xFFF  
        off += 4
        if off + size > n:
            break

        if tag == TAG_PICTURE:
            yield section_dec[off:off+size]

        off += size

def extract_picture(payload: bytes) -> int:
    # 개체 공통속성길이 파싱
    off = 0
    off += 4 * 7 # 제어ID, 속성, 세로off, 가로off, 폭, 높이, 층위, 
    off += 2 * 4 # 바깥 4방향 여백
    off += 4 * 2 # instance, 쪽 나눔 방지

    desc_len = struct.unpack_from("<H", payload, off)[0]
    off += 2

    off += desc_len * 2 # 개체 설명문 길이


# ─────────────────────────────
# 레닥션 함수
# ─────────────────────────────

# BinData 치환
def _redact_bindata(raw: bytes, targets: List[str]) -> bytes:
    if not targets:
        return raw

    rep = raw

    # 평문 치환
    for t in targets:
        rep, _ = _replace_in_bindata(rep, t)

    # zlib/rawdeflate 블록 치환
    for kind, off in scan_deflate(rep):
        res = decomp_bin(rep, off, kind)
        if not res:
            continue
        dec, consumed = res

        new_dec = dec
        total_hits = 0
        for t in targets:
            new_dec, c = _replace_in_bindata(new_dec, t)
            total_hits += c

        if total_hits <= 0:
            continue

        new_comp = recomp_bin(kind, new_dec)
        if not new_comp:
            continue

        patched = patch_seg(rep, off, consumed, new_comp)
        if patched is None:
            continue

        rep = patched

    return rep


# main 레닥션 함수
def redact(file_bytes: bytes) -> bytes:
    print("레닥션 시작")
    container = bytearray(file_bytes)
    full_raw = extract_text(file_bytes)["full_text"]
    full_norm = normalization_text(full_raw)
    targets = _collect_targets_by_regex(full_norm) 

    with olefile.OleFileIO(io.BytesIO(file_bytes)) as ole:
        streams = ole.listdir(streams=True, storages=False)
        print("[DEBUG streams] =>", streams)
        cutoff = getattr(ole, "minisector_cutoff", 4096)

        # BodyText/Section에서 OLE(차트)가 참조하는 BinDataID 수집
        ole_ids: set[int] = set()
        for path in streams:
            if len(path) >= 2 and path[0] == "BodyText" and path[1].startswith("Section"):
                raw = ole.openstream(path).read()
                dec, _ = _decompress(raw)
                ids = discover_ole_bindata_ids_strict(dec)
                if ids:
                    print(f"[DEBUG] Section {path} OLE BinDataIDs =", ids)
                    ole_ids.update(ids)

        # BodyText/Section* 본문 레닥션
        for path in streams:
            if not (len(path) >= 2 and path[0] == "BodyText" and path[1].startswith("Section")):
                continue

            raw = ole.openstream(path).read()
            dec, mode = _decompress(raw)
            buf = bytearray(dec)

            off, n = 0, len(buf)
            while off + 4 <= n:
                hdr = struct.unpack_from("<I", buf, off)[0]
                tag = hdr & 0x3FF
                size = (hdr >> 20) & 0xFFF
                off += 4
                if size < 0 or off + size > n:
                    break
                if tag == TAG_PARA_TEXT and size > 0 and targets:
                    seg = bytes(buf[off:off+size])
                    for t in targets:
                        seg, _ = _replace_utf16le_keep_len(seg, t)
                    buf[off:off+size] = seg
                off += size

            new_raw = _recompress(bytes(buf), mode)
            if len(new_raw) < len(raw):
                new_raw = new_raw + b"\x00" * (len(raw) - len(new_raw))
            elif len(new_raw) > len(raw):
                new_raw = new_raw[:len(raw)]

            entry = _direntry_for(ole, tuple(path))
            if not entry:
                continue
            if entry.size < cutoff:
                _overwrite_minifat_chain(ole, container, entry.isectStart, new_raw)
            else:
                _overwrite_bigfat(ole, container, entry.isectStart, new_raw)

        # BinData/*.OLE (차트 OLE만 필터링)
        for path in streams:
            if not (len(path) >= 2 and path[0] == "BinData" and path[1].endswith(".OLE")):
                continue

            # BINxxxx.OLE 에서 숫자 부분만 추출해 BinDataID로 사용
            name = path[1]
            num_part = "".join(ch for ch in name if ch.isdigit())
            try:
                bid = int(num_part) if num_part else None
            except ValueError:
                bid = None

            # 섹션에서 참조한 OLE가 아니면 스킵
            if not bid or (ole_ids and bid not in ole_ids):
                continue

            raw = ole.openstream(path).read()
            rep = _redact_bindata(raw, targets)

            if rep == raw or len(rep) != len(raw):
                continue

            entry = _direntry_for(ole, tuple(path))
            if not entry:
                continue

            # BinData는 일반적으로 Big FAT 영역
            _overwrite_bigfat(ole, container, entry.isectStart, rep)

        # PrvText
        for path in streams:
            if len(path) == 1:
                name = path[0].lower()
                if "prv" in name and "text" in name:
                    print("[DEBUG] PrvText detected:", path)

                    raw = ole.openstream(path).read()

                    new_raw = raw
                    for t in targets:
                        new_raw, _ = _replace_utf16le_keep_len(new_raw, t)

                    if len(new_raw) < len(raw):
                        new_raw = new_raw + b"\x00" * (len(raw) - len(new_raw))
                    elif len(new_raw) > len(raw):
                        new_raw = new_raw[:len(raw)]

                    entry = _direntry_for(ole, path)
                    if entry:
                        if entry.size < cutoff:
                            _overwrite_minifat_chain(ole, container, entry.isectStart, new_raw)
                        else:
                            _overwrite_bigfat(ole, container, entry.isectStart, new_raw)

                    print("[OK] prvtext 레닥션 완료!", path)
        
        # PrvImage
        for path in streams:
            if len(path) == 1:
                name = path[0].lower()
                if "prv" in name and "image" in name:
                    print("[DEBUG] PrvImage detected:", path)

                    raw = ole.openstream(path).read()
                    new_raw = b"\x00" * len(raw)

                    entry = _direntry_for(ole, path)
                    if entry:
                        if entry.size < cutoff:
                            _overwrite_minifat_chain(ole, container, entry.isectStart, new_raw)
                        else:
                            _overwrite_bigfat(ole, container, entry.isectStart, new_raw)

    return bytes(container)