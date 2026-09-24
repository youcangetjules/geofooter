#!/usr/bin/env python3
"""
Antivirus scanner facade for AES attachment scanning.

Prefers Windows Defender when enabled; otherwise uses Norton Auto-Protect on-access
scanning when Norton is the active AV product on the PC.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Optional, Protocol

logger = logging.getLogger(__name__)

try:
    from windows_defender import WindowsDefenderClient
except ImportError:
    WindowsDefenderClient = None  # type: ignore

try:
    from norton_scanner import NortonOnAccessClient
except ImportError:
    NortonOnAccessClient = None  # type: ignore


@dataclass
class AvDirectoryResult:
    provider: str
    status: str  # clean | threat | disabled | unavailable | error | ready
    detail: str = ""
    threats_by_filename: Dict[str, str] = field(default_factory=dict)


class _AvClient(Protocol):
    def scan_directory(self, directory: str | Path) -> object: ...


def scan_directory(directory: str | Path, timeout_seconds: int = 180) -> AvDirectoryResult:
    provider, client = _select_client(timeout_seconds)
    if client is None:
        return AvDirectoryResult(
            provider="",
            status="unavailable",
            detail="No supported antivirus scanner available",
        )

    raw = client.scan_directory(directory)
    return AvDirectoryResult(
        provider=provider,
        status=getattr(raw, "status", "error"),
        detail=getattr(raw, "detail", ""),
        threats_by_filename=dict(getattr(raw, "threats_by_filename", {}) or {}),
    )


def _select_client(timeout_seconds: int) -> tuple[str, Optional[_AvClient]]:
    defender = _defender_client_if_ready(timeout_seconds)
    if defender is not None:
        return "Windows Defender", defender

    norton = _norton_client_if_ready(timeout_seconds)
    if norton is not None:
        return "Norton", norton

    return "", None


def _defender_client_if_ready(timeout_seconds: int) -> Optional[_AvClient]:
    if WindowsDefenderClient is None:
        return None

    client = WindowsDefenderClient(timeout_seconds=timeout_seconds)
    availability = client._check_availability()  # noqa: SLF001 - shared integration probe
    if availability.status == "ready":
        logger.info("Using Windows Defender for attachment AV scan")
        return client

    logger.info(
        "Windows Defender unavailable for attachment scan: %s — %s",
        availability.status,
        availability.detail,
    )
    return None


def _norton_client_if_ready(timeout_seconds: int) -> Optional[_AvClient]:
    if NortonOnAccessClient is None:
        return None

    client = NortonOnAccessClient(timeout_seconds=timeout_seconds)
    availability = client._check_availability()  # noqa: SLF001 - shared integration probe
    if availability.status == "ready":
        logger.info("Using Norton Auto-Protect for attachment AV scan: %s", availability.detail)
        return client

    logger.info(
        "Norton unavailable for attachment scan: %s — %s",
        availability.status,
        availability.detail,
    )
    return None
