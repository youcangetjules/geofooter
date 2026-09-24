#!/usr/bin/env python3
"""
AES attachment scanner — content-aware checks beyond file extensions.
"""

from __future__ import annotations

import hashlib
import logging
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

try:
    from aes.scanners.av import AvDirectoryResult, scan_directory as scan_with_av
except ImportError:
    scan_with_av = None  # type: ignore
    AvDirectoryResult = None  # type: ignore

logger = logging.getLogger(__name__)

DANGEROUS_EXTENSIONS = frozenset({
    "exe", "bat", "cmd", "com", "scr", "pif", "vbs", "vbe", "js", "jse", "wsf", "wsh",
    "ps1", "psm1", "msi", "msp", "hta", "dll", "lnk", "iso", "img", "jar", "reg", "inf",
    "docm", "xlsm", "pptm", "dotm", "xltm", "potm", "ppam", "sldm", "vb", "ws", "wsc",
    "application", "gadget", "msc", "cpl", "sys", "drv", "ocx", "apk", "deb", "rpm",
})

MACRO_ENABLED_EXTENSIONS = frozenset({
    "docm", "xlsm", "pptm", "dotm", "xltm", "potm", "ppam", "sldm",
})

SAFE_DOCUMENT_EXTENSIONS = frozenset({
    "pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx", "txt", "csv", "rtf", "odt", "ods", "odp",
})

SUSPICIOUS_IN_ARCHIVE = frozenset(DANGEROUS_EXTENSIONS | {"html", "htm", "js", "vbs", "ps1"})

READ_LIMIT = 512 * 1024  # 512 KB for content inspection
DEEP_READ_LIMIT = 2 * 1024 * 1024  # 2 MB for Deep Scan content inspection


@dataclass
class AttachmentVerdict:
    filename: str
    ok: bool
    risk_level: str  # low | medium | high
    reasons: List[str] = field(default_factory=list)
    declared_extension: str = ""
    detected_type: str = "unknown"
    sha256: str = ""
    size_bytes: int = 0
    defender_status: str = ""  # clean | threat | skipped | error
    defender_threat: str = ""
    av_provider: str = ""  # Windows Defender | Norton

    def to_metadata_fragment(self) -> str:
        status = "OK" if self.ok else "NOT_OK"
        reason = self.reasons[0] if self.reasons else "clean"
        reason = reason.replace(";", ",")[:120]
        return (
            f"{self.filename}|{status}|{self.detected_type}|{self.declared_extension}|{reason}"
        )


class AttachmentScanner:
    """Scan exported attachment files with signature, heuristic, and AV checks."""

    def __init__(
        self,
        use_defender: bool = True,
        defender_timeout: int = 180,
        deep_mode: bool = False,
    ):
        # use_defender retained for compatibility; triggers any supported AV backend.
        self.use_defender = use_defender
        self.defender_timeout = defender_timeout
        self.deep_mode = deep_mode
        self.read_limit = DEEP_READ_LIMIT if deep_mode else READ_LIMIT
        self.archive_depth = 2 if deep_mode else 1
        self.archive_list_limit = 80 if deep_mode else 25

    def scan_directory(self, directory: str | Path) -> List[AttachmentVerdict]:
        scan_path = Path(directory)
        if not scan_path.is_dir():
            logger.warning("Attachment scan directory missing: %s", directory)
            return []

        results: List[AttachmentVerdict] = []
        for file_path in sorted(p for p in scan_path.iterdir() if p.is_file()):
            try:
                results.append(self.scan_file(file_path))
            except Exception as exc:
                logger.error("Attachment scan failed for %s: %s", file_path.name, exc)
                results.append(
                    AttachmentVerdict(
                        filename=file_path.name,
                        ok=False,
                        risk_level="medium",
                        reasons=[f"Scan error: {exc}"],
                        declared_extension=_extension(file_path.name),
                        detected_type="error",
                        size_bytes=file_path.stat().st_size if file_path.exists() else 0,
                    )
                )

        if self.use_defender and results:
            self._apply_av_results(scan_path, results)
        return results

    def _apply_av_results(self, scan_path: Path, results: List[AttachmentVerdict]) -> None:
        if scan_with_av is None:
            logger.warning("av_scanner.py not available — skipping antivirus scan")
            for item in results:
                item.defender_status = "skipped"
            return

        av_result = scan_with_av(scan_path, timeout_seconds=self.defender_timeout)
        provider = av_result.provider or "Antivirus"
        logger.info(
            "%s directory scan (%s): %s — %s",
            provider,
            scan_path,
            av_result.status,
            av_result.detail,
        )

        if av_result.status in {"disabled", "unavailable"}:
            for item in results:
                item.defender_status = "skipped"
                item.av_provider = provider
            return

        if av_result.status == "error":
            for item in results:
                item.defender_status = "error"
                item.av_provider = provider
                item.reasons.append(f"{provider}: {av_result.detail}")
                item.ok = False
                item.risk_level = _max_risk(item.risk_level, "medium")
            return

        for item in results:
            item.av_provider = provider
            threat_name = av_result.threats_by_filename.get(item.filename)
            if threat_name:
                item.ok = False
                item.risk_level = "high"
                item.defender_status = "threat"
                item.defender_threat = threat_name
                item.reasons.append(f"{provider}: {threat_name}")
            elif av_result.status == "clean":
                item.defender_status = "clean"
            else:
                item.defender_status = "clean"

    def scan_file(self, file_path: str | Path) -> AttachmentVerdict:
        path = Path(file_path)
        full_data = path.read_bytes()
        data = full_data[: self.read_limit]
        size = len(full_data)
        sha256 = hashlib.sha256(full_data).hexdigest()
        declared_ext = _extension(path.name)
        detected = _detect_type(data)
        reasons: List[str] = []
        risk = "low"

        if size == 0:
            reasons.append("Empty attachment")
            risk = "medium"

        if _has_double_extension(path.name):
            reasons.append("Double extension filename")
            risk = "high"

        if declared_ext in DANGEROUS_EXTENSIONS:
            reasons.append(f"Blocked extension .{declared_ext}")
            risk = "high"

        if declared_ext in MACRO_ENABLED_EXTENSIONS:
            reasons.append("Macro-enabled Office document")
            risk = "high"

        mismatch = _extension_mismatch(declared_ext, detected)
        if mismatch:
            reasons.append(mismatch)
            risk = "high"

        if detected == "pe_executable":
            reasons.append("Windows executable content")
            risk = "high"

        if detected == "html" and declared_ext not in {"html", "htm"}:
            reasons.append("HTML/script content with misleading extension")
            risk = "high"

        if _contains_script_payload(data):
            reasons.append("Embedded script content detected")
            risk = _max_risk(risk, "high")

        if _legacy_office_macros(data, declared_ext):
            reasons.append("Legacy Office macro indicators")
            risk = "high"

        if _ooxml_macros(path, declared_ext):
            reasons.append("Office Open XML macro project present")
            risk = "high"

        if detected in {"zip_archive", "7z", "rar"} and declared_ext not in {
            "zip", "docx", "xlsx", "pptx", "xlsm", "docm", "pptm", "jar", "apk", "odt", "ods", "odp",
        }:
            nested = _scan_archive_contents(
                path,
                limit=self.archive_list_limit,
                max_depth=self.archive_depth,
            )
            if nested:
                reasons.append(f"Archive contains suspicious items: {', '.join(nested[:5])}")
                risk = "high"
        elif self.deep_mode and detected == "zip_archive":
            nested = _scan_archive_contents(
                path,
                limit=self.archive_list_limit,
                max_depth=self.archive_depth,
            )
            if nested:
                reasons.append(f"Archive contains suspicious items: {', '.join(nested[:5])}")
                risk = _max_risk(risk, "high")

        if declared_ext in SAFE_DOCUMENT_EXTENSIONS and detected == "zip_archive":
            if declared_ext in {"doc", "xls", "ppt"}:
                pass  # legacy OLE handled separately
            elif not _valid_ooxml(path, declared_ext):
                reasons.append("Office document extension but invalid package structure")
                risk = _max_risk(risk, "medium")

        ok = len(reasons) == 0
        return AttachmentVerdict(
            filename=path.name,
            ok=ok,
            risk_level=risk if not ok else "low",
            reasons=reasons,
            declared_extension=declared_ext or "none",
            detected_type=detected,
            sha256=sha256,
            size_bytes=size,
        )

    @staticmethod
    def summarize(results: Sequence[AttachmentVerdict]) -> Tuple[int, int, int]:
        total = len(results)
        ok_count = sum(1 for item in results if item.ok)
        return total, ok_count, total - ok_count


def _extension(name: str) -> str:
    if "." not in name:
        return ""
    return name.rsplit(".", 1)[-1].lower()


def _detect_type(data: bytes) -> str:
    if data.startswith(b"MZ"):
        return "pe_executable"
    if data.startswith(b"PK\x03\x04"):
        return "zip_archive"
    if data.startswith(b"%PDF"):
        return "pdf"
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if data.startswith(b"GIF87a") or data.startswith(b"GIF89a"):
        return "gif"
    if data.startswith(b"\xff\xd8\xff"):
        return "jpeg"
    if data.startswith(b"RIFF") and len(data) >= 12 and data[8:12] == b"WEBP":
        return "webp"
    if data.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return "ole_compound"
    if data.startswith(b"7z\xbc\xaf\x27\x1c"):
        return "7z"
    if data.startswith(b"Rar!\x1a\x07") or data.startswith(b"Rar!\x1a\x07\x00"):
        return "rar"
    head = data[:256].lstrip()
    lowered = head.lower()
    if lowered.startswith(b"<!doctype html") or lowered.startswith(b"<html") or lowered.startswith(b"<script"):
        return "html"
    if lowered.startswith(b"<?xml"):
        return "xml"
    if data.startswith(b"BM"):
        return "bmp"
    return "unknown"


def _has_double_extension(name: str) -> bool:
    lower = name.lower()
    parts = lower.split(".")
    if len(parts) < 3:
        return False
    return parts[-1] in DANGEROUS_EXTENSIONS or (
        parts[-2] in SAFE_DOCUMENT_EXTENSIONS and parts[-1] in DANGEROUS_EXTENSIONS
    )


def _extension_mismatch(declared_ext: str, detected: str) -> Optional[str]:
    if not declared_ext or detected == "unknown":
        return None

    expected = {
        "pdf": {"pdf"},
        "png": {"png"},
        "jpg": {"jpeg"},
        "jpeg": {"jpeg"},
        "gif": {"gif"},
        "webp": {"webp"},
        "bmp": {"bmp"},
        "zip": {"zip_archive"},
        "docx": {"zip_archive"},
        "xlsx": {"zip_archive"},
        "pptx": {"zip_archive"},
        "docm": {"zip_archive"},
        "xlsm": {"zip_archive"},
        "pptm": {"zip_archive"},
        "doc": {"ole_compound"},
        "xls": {"ole_compound"},
        "ppt": {"ole_compound"},
        "exe": {"pe_executable"},
        "dll": {"pe_executable"},
        "html": {"html"},
        "htm": {"html"},
    }
    allowed = expected.get(declared_ext)
    if allowed and detected not in allowed:
        return f"Extension .{declared_ext} does not match detected {detected}"
    if declared_ext in SAFE_DOCUMENT_EXTENSIONS and detected == "pe_executable":
        return f"Extension .{declared_ext} hides executable content"
    if declared_ext in {"pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx"} and detected == "html":
        return f"Extension .{declared_ext} hides HTML/script content"
    return None


def _contains_script_payload(data: bytes) -> bool:
    sample = data[:8192].lower()
    markers = (
        b"<script",
        b"javascript:",
        b"powershell",
        b"cmd.exe",
        b"wscript.shell",
        b"auto_open",
        b"vbaproject",
    )
    return any(marker in sample for marker in markers)


def _legacy_office_macros(data: bytes, declared_ext: str) -> bool:
    if declared_ext not in {"doc", "xls", "ppt", "rtf"} and _detect_type(data) != "ole_compound":
        return False
    markers = (b"vba", b"_vba_project", b"macros", b"autoopen", b"document_open")
    lowered = data.lower()
    return any(marker in lowered for marker in markers)


def _ooxml_macros(path: Path, declared_ext: str) -> bool:
    if _detect_type(path.read_bytes()[:4]) != "zip_archive":
        return False
    if declared_ext not in MACRO_ENABLED_EXTENSIONS and declared_ext not in {"docx", "xlsx", "pptx", "docm", "xlsm", "pptm"}:
        return False
    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            if "vbaProject.bin" in names:
                return True
            if any("vba" in name.lower() for name in names):
                return True
    except zipfile.BadZipFile:
        return False
    return False


def _valid_ooxml(path: Path, declared_ext: str) -> bool:
    mapping = {
        "docx": "word/",
        "docm": "word/",
        "xlsx": "xl/",
        "xlsm": "xl/",
        "pptx": "ppt/",
        "pptm": "ppt/",
    }
    expected_prefix = mapping.get(declared_ext)
    if not expected_prefix:
        return True
    try:
        with zipfile.ZipFile(path) as archive:
            return any(name.startswith(expected_prefix) for name in archive.namelist())
    except zipfile.BadZipFile:
        return False


def _scan_archive_contents(
    path: Path,
    limit: int = 25,
    max_depth: int = 1,
    _depth: int = 0,
) -> List[str]:
    suspicious: List[str] = []
    if _depth >= max_depth:
        return suspicious
    try:
        with zipfile.ZipFile(path) as archive:
            for info in archive.infolist()[:limit]:
                if info.is_dir():
                    continue
                inner_ext = _extension(info.filename)
                if inner_ext in SUSPICIOUS_IN_ARCHIVE:
                    suspicious.append(Path(info.filename).name)
                elif (
                    max_depth > 1
                    and _depth + 1 < max_depth
                    and inner_ext == "zip"
                ):
                    try:
                        nested_bytes = archive.read(info)
                        import tempfile
                        with tempfile.NamedTemporaryFile(suffix=".zip", delete=False) as tmp:
                            tmp.write(nested_bytes)
                            tmp_path = Path(tmp.name)
                        try:
                            nested_hits = _scan_archive_contents(
                                tmp_path, limit=limit, max_depth=max_depth, _depth=_depth + 1
                            )
                            for hit in nested_hits:
                                suspicious.append(f"{Path(info.filename).name}/{hit}")
                        finally:
                            try:
                                tmp_path.unlink(missing_ok=True)
                            except Exception:
                                pass
                    except Exception:
                        pass
    except zipfile.BadZipFile:
        return []
    return suspicious


def _max_risk(current: str, new: str) -> str:
    order = {"low": 0, "medium": 1, "high": 2}
    return new if order.get(new, 0) > order.get(current, 0) else current


def format_scan_metadata(results: Sequence[AttachmentVerdict]) -> str:
    """Compact metadata string for AES export preamble."""
    total, ok_count, not_ok_count = AttachmentScanner.summarize(results)
    details = ";".join(item.to_metadata_fragment() for item in results[:20])
    return f"count={total}; ok={ok_count}; not_ok={not_ok_count}; details={details}"
