from __future__ import annotations
import re, struct, unicodedata, zlib
from io import BytesIO
from typing import Any, Iterable, List, Optional, Tuple
from server.core.matching import find_sensitive_spans
from server.core.normalize import normalization_index

_ZW_CHARS = r"\u200B\u200C\u200D\uFEFF"

def _norm_line(t: str) -> str:
    t = unicodedata.normalize("NFKC", t or "")
    t = re.sub(f"[{_ZW_CHARS}]", "", t)
    t = t.replace("\r\n", "\n").replace("\r", "\n")
    t = re.sub(r"[ \t]+", " ", t)
    return t.strip()


_LINE_NOISE = re.compile(
    r"(?i)"  # case-insensitive
    r"(?:제목을 입력|텍스트를 입력|클릭하여|표지|발표자|날짜|부제목)"
    r"|마스터\s*(?:제목|텍스트)\s*스타일"
)

_RE_MASTER_LEVEL = re.compile(
    r"^(?:[•·\*\-\–\—◦●○◆◇▪▫▶▷■□]+\s*)?"
    r"(?:첫|두|둘|세|셋|네|넷|다섯|여섯|일곱|여덟|아홉|열|[0-9]+)\s*(?:번째)?\s*수준\s*$",
    re.IGNORECASE,
)

_RE_BULLET_ONLY = re.compile(r"^[\*\u2022•·\-\–\—○●◦■□]+$", re.IGNORECASE)


def _is_noise_line(line: str) -> bool:
    if not line:
        return True
    if "마스터" in line and "스타일" in line:
        return True
    if "편집하려면 클릭" in line:
        return True
    if _RE_MASTER_LEVEL.match(line):
        return True
    if _RE_BULLET_ONLY.match(line):
        return True

    return False

_HDR = struct.Struct("<HHI")

_TEXTCHARSATOM = 0x0FA0  # UTF-16LE 텍스트
_TEXTBYTESATOM = 0x0FA8  # 단일바이트 텍스트


try:  # olefile optional import
    import olefile  # type: ignore
    from olefile import OleFileIO  # type: ignore
except Exception:  # pragma: no cover
    olefile = None
    OleFileIO = Any  # type: ignore


# 기본 유틸
def _cleanup(text: str) -> str:
    out: List[str] = []
    for raw_line in (text or "").splitlines():
        line = _norm_line(raw_line)
        if not line:
            continue
        if _is_noise_line(line):
            continue
        out.append(line)
    return "\n".join(out)


def _walk_records(buf: bytes, base_off: int = 0) -> Iterable[Tuple[int, int, int, int]]:

    i, n = 0, len(buf)
    while i + _HDR.size <= n:
        try:
            verInst, rtype, rlen = _HDR.unpack_from(buf, i)
        except struct.error:
            break

        rec_ver = verInst & 0x000F
        i_hdr_end = i + _HDR.size
        i_data_end = i_hdr_end + rlen
        if i_data_end > n or rlen < 0:
            break

        data_off_abs = base_off + i_hdr_end

        # 컨테이너(하위 레코드 포함)
        if rec_ver == 0xF:
            child = buf[i_hdr_end:i_data_end]
            yield from _walk_records(child, base_off + i_hdr_end)
        else:
            # leaf record
            yield (rec_ver, rtype, rlen, data_off_abs)

        i = i_data_end


def _read_powerpoint_document(ole: OleFileIO) -> bytes:

    with ole.openstream("PowerPoint Document") as fp:  # type: ignore[attr-defined]
        return fp.read()


def _extract_text_from_records(doc_bytes: bytes) -> str:

    chunks: List[str] = []

    for _rec_ver, rtype, rlen, data_off in _walk_records(doc_bytes):
        if rlen <= 0:
            continue

        try:
            b = doc_bytes[data_off : data_off + rlen]
        except Exception:
            continue

        txt: Optional[str] = None
        if rtype == _TEXTCHARSATOM:
            # UTF-16LE 슬라이드 텍스트
            txt = b.decode("utf-16le", errors="ignore")
        elif rtype == _TEXTBYTESATOM:
            # 구버전/영문 슬라이드 텍스트
            try:
                txt = b.decode("cp949", errors="ignore")
            except Exception:
                txt = b.decode("latin1", errors="ignore")

        if not txt:
            continue

        txt = txt.replace("\r\n", "\n").replace("\r", "\n")
        chunks.append(txt)

    merged = "\n".join(chunks)
    return _cleanup(merged)


# 임베디드 OLE/차트 내부에서 휴리스틱으로 텍스트를 추출한다.
def _extract_text_from_ole_stream(raw: bytes) -> str:

    try:
        import olefile as _ole  # type: ignore
    except Exception:  # pragma: no cover
        return ""

    if not raw:
        return ""

    out: List[str] = []
    try:
        with _ole.OleFileIO(BytesIO(raw)) as sub:
            for entry in sub.listdir(streams=True, storages=False):
                try:
                    with sub.openstream(entry) as fp:
                        blob = fp.read()
                except Exception:
                    continue

                # UTF-8 / UTF-16 / cp949 정도만 시도
                for enc in ("utf-8", "utf-16le", "cp949", "latin1"):
                    try:
                        txt = blob.decode(enc)
                        txt = _cleanup(txt)
                        if txt:
                            out.append(txt)
                        break
                    except Exception:
                        continue
    except Exception:
        return ""

    return "\n".join(out)


def _extract_embedded_noise_prone(ole: OleFileIO) -> str:

    out: List[str] = []

    for entry in ole.listdir(streams=True, storages=False):
        path = "/".join(entry)

        # 핵심 Document 스트림은 여기서 건너뜀
        if path == "PowerPoint Document":
            continue

        try:
            with ole.openstream(entry) as fp:
                blob = fp.read()
        except Exception:
            continue

        # 압축 여부 heuristic (zlib 헤더)
        if len(blob) > 6 and blob[0] == 0x78 and blob[1] in (0x01, 0x9C, 0xDA):
            try:
                blob = zlib.decompress(blob)
            except Exception:
                pass

        txt = _extract_text_from_ole_stream(blob)
        if txt:
            out.append(txt)

    return _cleanup("\n".join(out))


def _extract_chart_ole_text_from_doc(doc_bytes: bytes) -> str:

    out: List[str] = []
    n = len(doc_bytes)
    i = 0

    while i + 8 <= n:
        try:
            verInst, rtype, rlen = _HDR.unpack_from(doc_bytes, i)
        except struct.error:
            break

        i_hdr_end = i + _HDR.size
        i_data_end = i_hdr_end + rlen
        if i_data_end > n or rlen < 0:
            break

        # ExOleObjStgCompressedAtom 같은 OLE 압축 blob 추정 영역만 본다.
        if rlen > 32:
            blob = doc_bytes[i_hdr_end:i_data_end]
            # zlib signature
            if len(blob) > 6 and blob[0] == 0x78 and blob[1] in (0x01, 0x9C, 0xDA):
                try:
                    decomp = zlib.decompress(blob)
                except Exception:
                    decomp = None
                if decomp:
                    txt = _extract_text_from_ole_stream(decomp)
                    if txt:
                        out.append(txt)

        i = i_data_end

    return _cleanup("\n".join(out))


# 외부 공개 함수: 텍스트 추출 + 레닥션
# 필요 시, 환경변수로 임베디드 텍스트 추출 on/off 가능하게
PPT_EXTRACT_EMBEDDED = True


def extract_text(file_bytes: bytes):

    if olefile is None:
        raise RuntimeError("olefile 모듈이 필요합니다. pip install olefile")
    with olefile.OleFileIO(BytesIO(file_bytes)) as ole:  # type: ignore[arg-type]
        doc = _read_powerpoint_document(ole)
        text_main = _extract_text_from_records(doc)

        if PPT_EXTRACT_EMBEDDED:
            extra_parts: List[str] = []

            # 1) 루트 OLE (embeddings/ObjectPool 등)에서 텍스트
            text_emb_root = _extract_embedded_noise_prone(ole)
            if text_emb_root:
                extra_parts.append(text_emb_root)

            # 2) PowerPoint Document 안의 차트 OLE 텍스트
            text_emb_chart = _extract_chart_ole_text_from_doc(doc)
            if text_emb_chart:
                extra_parts.append(text_emb_chart)

            if extra_parts:
                text_main = _cleanup(
                    (text_main or "") + "\n" + "\n".join(extra_parts)
                )

    pages = [{"index": 1, "text": text_main or ""}]
    return {"full_text": text_main or "", "pages": pages}


def redact(file_bytes: bytes) -> bytes:

    try:
        from .ole_redactor import redact_ole_bin_preserve_size  # type: ignore
    except Exception:  # pragma: no cover
        try:
            from server.modules.ole_redactor import redact_ole_bin_preserve_size  # type: ignore
        except Exception:
            redact_ole_bin_preserve_size = None  # type: ignore

    if redact_ole_bin_preserve_size is None:
        return file_bytes

    try:
        data = extract_text(file_bytes) or {}
        raw_text = data.get("full_text", "") or ""
    except Exception as e:
        print(f"[PPT] extract_text 예외: {e!r}")
        return file_bytes

    if not raw_text:
        print("[PPT] extract_text 결과가 비어 있음 → 레닥션 생략")
        return file_bytes

    # 정규화 텍스트 / 인덱스 맵 생성
    norm_text, index_map = normalization_index(raw_text)

    try:
        matches = find_sensitive_spans(norm_text)
    except Exception as e:
        print(f"[PPT] find_sensitive_spans 예외: {e!r}")
        return file_bytes

    if not matches:
        print("[PPT] 민감정보 매칭 0건 → 레닥션 생략")
        return file_bytes

    secrets: List[str] = []
    for s_idx, e_idx, _val, _name in matches:
        if not isinstance(s_idx, int) or not isinstance(e_idx, int) or e_idx <= s_idx:
            continue
        # 정규화 인덱스를 원본 인덱스로 역매핑
        if s_idx not in index_map or (e_idx - 1) not in index_map:
            continue
        start = index_map[s_idx]
        end = index_map.get(e_idx - 1, start) + 1
        if end <= start:
            continue
        frag = raw_text[start:end].strip()
        if len(frag) >= 2:
            secrets.append(frag)

    if not secrets:
        print("[PPT] 매칭은 있으나 실제 시크릿 문자열이 비어 있음 → 레닥션 생략")
        return file_bytes

    # 중복 제거 (긴 문장 우선)
    uniq: List[str] = []
    seen = set()
    for v in sorted(set(secrets), key=lambda x: (-len(x), x)):
        if v not in seen:
            seen.add(v)
            uniq.append(v)
    secrets = uniq

    try:
        return redact_ole_bin_preserve_size(file_bytes, secrets, mask_preview=False)
    except Exception as e:
        print(f"[PPT] redact_ole_bin_preserve_size 예외: {e!r}")
        return file_bytes
