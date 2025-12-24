from __future__ import annotations

import io
import re
import zipfile
import logging
from typing import Optional, List, Tuple

try:
    from .common import (
        cleanup_text,
        compile_rules,
        sub_text_nodes,
        chart_sanitize,
        redact_embedded_xlsx_bytes,
        HWPX_STRIP_PREVIEW,
        HWPX_DISABLE_CACHE,
        HWPX_BLANK_PREVIEW,
    )
except Exception:  # pragma: no cover
    from server.modules.common import (  # type: ignore
        cleanup_text,
        compile_rules,
        sub_text_nodes,
        chart_sanitize,
        redact_embedded_xlsx_bytes,
        HWPX_STRIP_PREVIEW,
        HWPX_DISABLE_CACHE,
        HWPX_BLANK_PREVIEW,
    )

try:
    from ..core.schemas import XmlMatch, XmlLocation
except Exception:
    try:
        from ..schemas import XmlMatch, XmlLocation   # 일부 브랜치/옛 구조
    except Exception:
        from server.core.schemas import XmlMatch, XmlLocation  # 절대경로 fallback

log = logging.getLogger("xml_redaction")
if not log.handlers:
    _h = logging.StreamHandler()
    _h.setFormatter(logging.Formatter("[%(levelname)s] xml_redaction: %(message)s"))
    log.addHandler(_h)
log.setLevel(logging.INFO)

_CURRENT_SECRETS: List[str] = []


def set_hwpx_secrets(values: List[str] | None):
    global _CURRENT_SECRETS
    _CURRENT_SECRETS = list(dict.fromkeys(v for v in (values or []) if v))

# HWPX 내부 텍스트 수집

def hwpx_text(zipf: zipfile.ZipFile) -> str:
    out: List[str] = []
    names = zipf.namelist()

    # 1) 본문 Contents/* 의 텍스트
    for name in sorted(names):
        low = name.lower()
        if not (low.startswith("contents/") and low.endswith(".xml")):
            continue
        try:
            xml = zipf.read(name).decode("utf-8", "ignore")
            out += [m.group(1) for m in re.finditer(r">([^<>]+)<", xml)]
        except Exception:
            continue

    # 2) 차트 Chart(s)/* 의 a:t, c:v 텍스트 (라벨/범주/제목 등)
    for name in sorted(names):
        low = name.lower()
        if not ((low.startswith("chart/") or low.startswith("charts/")) and low.endswith(".xml")):
            continue
        try:
            s = zipf.read(name).decode("utf-8", "ignore")
            for m in re.finditer(r"<a:t[^>]*>(.*?)</a:t>|<c:v[^>]*>(.*?)</c:v>", s, re.I | re.DOTALL):
                v = (m.group(1) or m.group(2) or "").strip()
                if v:
                    out.append(v)
        except Exception:
            pass

    # 3) BinData/*: ZIP(=내장 XLSX)이면 그 안에서도 텍스트 수집
    for name in names:
        low = name.lower()
        if not low.startswith("bindata/"):
            continue
        try:
            b = zipf.read(name)
        except KeyError:
            continue

        if len(b) >= 4 and b[:2] == b"PK":
            try:
                try:
                    from .common import xlsx_text_from_zip
                except Exception:
                    from server.modules.common import xlsx_text_from_zip
                with zipfile.ZipFile(io.BytesIO(b), "r") as ez:
                    out.append(xlsx_text_from_zip(ez))
            except Exception:
                pass

    return cleanup_text("\n".join(x for x in out if x))


# ─────────────────────────────────────────────────────────────────────────────
# /text/extract 용 텍스트 추출 (사람이 보기 좋게 정리)
# ─────────────────────────────────────────────────────────────────────────────
def extract_text(file_bytes: bytes) -> dict:
    with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as zipf:
        raw = hwpx_text(zipf)

    txt = re.sub(r"<[^>\n]+>", "", raw)

    lines = []
    for line in txt.splitlines():
        if re.fullmatch(r"\(?\^\d+[\).\s]*", line.strip()):
            continue
        lines.append(line)

    txt = "\n".join(lines)
    txt = re.sub(r"\(\^\d+\)", "", txt)

    # 3) 엑셀 시트/범위 토큰 제거
    #    예: "Sheet1!$B$1", "Sheet1!$B$2:$B$5"
    txt = re.sub(
        r"Sheet\d*!\$[A-Z]+\$\d+(?::\$[A-Z]+\$\d+)?",
        "",
        txt,
        flags=re.IGNORECASE,
    )

    # 4) "General4.3" 같은 포맷 문자열에서 General 제거 → "4.3"
    txt = re.sub(r"General(?=\s*\d)", "", txt, flags=re.IGNORECASE)

    # 5) 공백/줄바꿈 정리
    cleaned = cleanup_text(txt)

    return {
        "full_text": cleaned,
        "pages": [
            {"page": 1, "text": cleaned},
        ],
    }


# ─────────────────────────────────────────────────────────────────────────────
# 스캔: 정규식 규칙으로 텍스트에서 민감정보 후보를 추출
# ─────────────────────────────────────────────────────────────────────────────
def scan(zipf: zipfile.ZipFile) -> Tuple[List[XmlMatch], str, str]:
    text = hwpx_text(zipf)
    comp = compile_rules()

    try:
        from ..core.redaction_rules import RULES
    except Exception:
        from server.core.redaction_rules import RULES

    def _validator(rule):
        try:
            v = RULES.get(rule, {}).get("validator")
            return v if callable(v) else None
        except Exception:
            return None

    out: List[XmlMatch] = []

    for ent in comp:
        try:
            if isinstance(ent, (list, tuple)):
                rule, rx = ent[0], ent[1]
                need_valid = bool(ent[2]) if len(ent) >= 3 else True
            else:
                rule = getattr(ent, "name", getattr(ent, "rule", "unknown"))
                rx = getattr(ent, "rx", None)
                need_valid = bool(getattr(ent, "need_valid", True))
            if rx is None:
                continue
        except Exception:
            continue

        vfunc = _validator(rule)

        for m in rx.finditer(text):
            val = m.group(0)
            ok = True
            if need_valid and vfunc:
                try:
                    ok = bool(vfunc(val))
                except Exception:
                    ok = False

            out.append(
                XmlMatch(
                    rule=rule_name,
                    value=val,
                    valid=ok,
                    context=text[max(0, m.start() - 20): min(len(text), m.end() + 20)],
                    location=XmlLocation(
                        kind="hwpx",
                        part="*merged_text*",
                        start=m.start(),
                        end=m.end(),
                    ),
                )
            )

    return out, "hwpx", text


# ─────────────────────────────────────────────────────────────────────────────
# 파일 단위 레닥션
# ─────────────────────────────────────────────────────────────────────────────
def redact_item(filename: str, data: bytes, comp) -> Optional[bytes]:
    low = filename.lower()
    log.info("[HWPX][RED] entry=%s size=%d", filename, len(data))

    if low.startswith("preview/"):
        if low.endswith(IMAGE_EXTS):
            log.info("[HWPX][IMG] preview image=%s size=%d", filename, len(data))
        return b""

    # 2) settings.xml: 캐시/프리뷰 비활성화
    if HWPX_DISABLE_CACHE and low.endswith("settings.xml"):
        try:
            txt = data.decode("utf-8", "ignore")
            txt = re.sub(r'(?i)usepreview\s*=\s*"(?:true|1)"', 'usePreview="false"', txt)
            txt = re.sub(r"(?is)<preview>.*?</preview>", "<preview>0</preview>", txt)
            txt = re.sub(r"(?is)<cache>.*?</cache>", "<cache>0</cache>", txt)
            return txt.encode("utf-8", "ignore")
        except Exception:
            return data

    if low.startswith("contents/") and low.endswith(".xml"):
        return sub_text_nodes(data, comp)[0]

    if (low.startswith("chart/") or low.startswith("charts/")) and low.endswith(".xml"):
        b2, _ = chart_sanitize(data, comp)   # a:t, c:strCache
        masked, _ = sub_text_nodes(b2, comp)
        return masked

    # 5) BinData: 내장 XLSX 또는 OLE(CFBF)
    if low.startswith("bindata/"):
        # (a) ZIP(=PK..) → 내장 XLSX
        if len(data) >= 4 and data[:2] == b"PK":
            try:
                return redact_embedded_xlsx_bytes(data)
            except Exception:
                return data
        # (b) 그 외 → CFBF(OLE) 가능. 프리뷰는 무조건 블랭크 + 시크릿/이메일 동일길이 마스킹
        try:
            try:
                from .ole_redactor import redact_ole_bin_preserve_size
            except Exception:  # pragma: no cover
                from server.modules.ole_redactor import redact_ole_bin_preserve_size  # type: ignore

            return redact_ole_bin_preserve_size(data, _CURRENT_SECRETS, mask_preview=True)
        except Exception:
            return data

    if low.endswith(".xml") and not low.startswith("preview/"):
        return sub_text_nodes(data, comp)[0]

    return None
