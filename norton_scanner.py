#!/usr/bin/env python3
"""
Norton integration for AES attachment scanning.

Modern Norton 360 builds (Avast engine) do not expose a reliable ashCmd.exe CLI.
This module uses Norton Auto-Protect on-access scanning: reading an exported file
triggers a real-time scan and blocked malware returns a distinctive OSError.
"""

from __future__ import annotations

import logging
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Optional

logger = logging.getLogger(__name__)

_CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000) if sys.platform == "win32" else 0


def _hidden_run(cmd, timeout: int) -> subprocess.CompletedProcess:
    kwargs = {
        "capture_output": True,
        "text": True,
        "timeout": timeout,
        "check": False,
    }
    if _CREATE_NO_WINDOW:
        kwargs["creationflags"] = _CREATE_NO_WINDOW
    return subprocess.run(cmd, **kwargs)


NORTON_INSTALL_CANDIDATES = (
    Path(r"C:\Program Files\Norton\Suite"),
    Path(r"C:\Program Files (x86)\Norton\Suite"),
)


@dataclass
class NortonDirectoryResult:
    status: str  # clean | threat | disabled | unavailable | error
    detail: str = ""
    threats_by_filename: Dict[str, str] = field(default_factory=dict)


class NortonOnAccessClient:
    """Scan exported attachments via Norton real-time (on-access) protection."""

    def __init__(self, timeout_seconds: int = 60):
        self.timeout_seconds = timeout_seconds

    def scan_directory(self, directory: str | Path) -> NortonDirectoryResult:
        scan_dir = Path(directory)
        if not scan_dir.is_dir():
            return NortonDirectoryResult("error", f"Scan directory not found: {scan_dir}")

        availability = self._check_availability()
        if availability.status != "ready":
            return availability

        threats: Dict[str, str] = {}
        errors: list[str] = []

        for file_path in sorted(p for p in scan_dir.iterdir() if p.is_file()):
            status, detail = self._scan_file_on_access(file_path)
            if status == "threat":
                threats[file_path.name] = detail
            elif status == "error":
                errors.append(f"{file_path.name}: {detail}")

        if threats:
            return NortonDirectoryResult(
                status="threat",
                detail=f"{len(threats)} threat(s) blocked by Norton Auto-Protect",
                threats_by_filename=threats,
            )

        if errors:
            return NortonDirectoryResult(
                status="error",
                detail="; ".join(errors[:3]),
            )

        return NortonDirectoryResult(
            status="clean",
            detail="Norton Auto-Protect allowed access to all exported attachments",
        )

    def _scan_file_on_access(self, file_path: Path) -> tuple[str, str]:
        # Norton may trust python.exe; probe each file via PowerShell ReadAllBytes.
        # Do not pass -NonInteractive — it changes Norton/quarantine timing on this host.
        if not file_path.exists():
            return "threat", "Norton removed file (quarantined)"

        escaped = str(file_path).replace("'", "''")
        command = (
            f"try {{ [void][System.IO.File]::ReadAllBytes('{escaped}'); "
            f"Write-Output 'CLEAN' }} catch {{ Write-Output $_.Exception.Message }}"
        )
        output, exit_code = _run_powershell_command(command, timeout=self.timeout_seconds)
        combined = output.lower()

        if "clean" in combined and "virus" not in combined and "potentially unwanted" not in combined:
            return "clean", ""
        if "virus" in combined or "potentially unwanted" in combined:
            return "threat", "Norton blocked file (virus or PUP)"
        if "could not find file" in combined:
            return "threat", "Norton removed file (quarantined)"
        if output.strip():
            return "error", output.strip()
        return "error", f"Norton scan produced no output (exit {exit_code})"

    def _check_availability(self) -> NortonDirectoryResult:
        if not _find_norton_install_root():
            return NortonDirectoryResult("unavailable", "Norton installation not found")

        display_name = _active_norton_product_name()
        if display_name:
            return NortonDirectoryResult("ready", f"{display_name} is active")

        return NortonDirectoryResult(
            "disabled",
            "Norton is installed but not reported as the active antivirus",
        )


def _find_norton_install_root() -> Optional[Path]:
    for candidate in NORTON_INSTALL_CANDIDATES:
        if candidate.is_dir():
            return candidate
    return None


def _active_norton_product_name() -> str:
    script = """
$ErrorActionPreference = 'SilentlyContinue'
$product = Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct |
    Where-Object { $_.displayName -like '*Norton*' -and ($_.productState -band 0x1000) } |
    Select-Object -First 1
if ($product) {
    Write-Output ('READY|' + [string]$product.displayName)
} else {
    Write-Output 'DISABLED|Norton not active'
}
"""
    output, _ = _run_powershell(script, timeout=20)
    for line in output.splitlines():
        if line.startswith("READY|"):
            return line.split("|", 1)[1].strip()
    return ""

def _run_powershell_command(command: str, timeout: int = 30) -> tuple[str, int]:
    cmd = [
        "powershell.exe",
        "-NoProfile",
        "-WindowStyle",
        "Hidden",
        "-Command",
        command,
    ]
    try:
        completed = _hidden_run(cmd, timeout=timeout)
        output = "\n".join(filter(None, [completed.stdout, completed.stderr])).strip()
        return output, completed.returncode
    except subprocess.TimeoutExpired:
        return "PowerShell timed out", 124
    except OSError as exc:
        return str(exc), 127


def _run_powershell(script: str, timeout: int = 30) -> tuple[str, int]:
    return _run_powershell_command(script, timeout=timeout)
