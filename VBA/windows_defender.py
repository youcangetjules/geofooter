#!/usr/bin/env python3
"""
Windows Defender integration for AES attachment scanning.
Uses MpCmdRun (primary) and Defender PowerShell cmdlets (status/threat lookup).
"""

from __future__ import annotations

import logging
import os
import re
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

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


MPCMDRUN_CANDIDATES = (
    Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Windows Defender" / "MpCmdRun.exe",
    Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Windows Defender" / "MpCmdRun.exe",
)


@dataclass
class DefenderDirectoryResult:
    status: str  # clean | threat | disabled | unavailable | error
    detail: str = ""
    threats_by_filename: Dict[str, str] = field(default_factory=dict)


class WindowsDefenderClient:
    """Scan a directory of exported attachments with Windows Defender."""

    def __init__(self, timeout_seconds: int = 180):
        self.timeout_seconds = timeout_seconds
        self._mpcmdrun_path: Optional[Path] = None

    def scan_directory(self, directory: str | Path) -> DefenderDirectoryResult:
        scan_dir = Path(directory)
        if not scan_dir.is_dir():
            return DefenderDirectoryResult("error", f"Scan directory not found: {scan_dir}")

        availability = self._check_availability()
        if availability.status != "ready":
            return availability

        scan_start = self._run_defender_scan(scan_dir)
        if scan_start.status in {"disabled", "unavailable", "error"}:
            return scan_start

        threats = self._collect_threats(scan_dir)
        if threats:
            return DefenderDirectoryResult(
                status="threat",
                detail=f"{len(threats)} threat(s) detected by Windows Defender",
                threats_by_filename=threats,
            )

        return DefenderDirectoryResult(
            status="clean",
            detail="Windows Defender scan completed — no threats detected",
        )

    def _check_availability(self) -> DefenderDirectoryResult:
        script = """
$ErrorActionPreference = 'SilentlyContinue'
try {
    $mp = Get-MpComputerStatus
    if (-not $mp) { Write-Output 'STATUS|unavailable|Defender status unavailable'; exit 10 }
    if (-not $mp.AMServiceEnabled -or -not $mp.AntivirusEnabled) {
        Write-Output 'STATUS|disabled|Windows Defender is not enabled on this PC'
        exit 10
    }
    Write-Output 'STATUS|ready|Windows Defender is enabled'
    exit 0
} catch {
    Write-Output ('STATUS|unavailable|' + $_.Exception.Message)
    exit 10
}
"""
        output, exit_code = self._run_powershell(script)
        for line in output.splitlines():
            if line.startswith("STATUS|"):
                _, status, detail = (line.split("|", 2) + [""])[:3]
                if status == "ready":
                    return DefenderDirectoryResult("ready", detail)
                return DefenderDirectoryResult(status, detail)
        return DefenderDirectoryResult("unavailable", "Could not determine Defender status")

    def _run_defender_scan(self, scan_dir: Path) -> DefenderDirectoryResult:
        mpcmd = self._find_mpcmdrun()
        if mpcmd:
            return self._run_mpcmdrun_scan(mpcmd, scan_dir)
        return self._run_powershell_scan(scan_dir)

    def _run_mpcmdrun_scan(self, mpcmd: Path, scan_dir: Path) -> DefenderDirectoryResult:
        cmd = [str(mpcmd), "-Scan", "-ScanType", "3", "-File", str(scan_dir)]
        logger.info("Windows Defender MpCmdRun scan: %s", scan_dir)
        try:
            completed = _hidden_run(
                cmd,
                timeout=self.timeout_seconds,
            )
        except subprocess.TimeoutExpired:
            return DefenderDirectoryResult("error", "Windows Defender scan timed out")
        except OSError as exc:
            return DefenderDirectoryResult("error", f"Windows Defender scan failed: {exc}")

        combined = "\n".join(filter(None, [completed.stdout, completed.stderr]))
        if "Product/Feature disabled" in combined or "disabled" in combined.lower():
            return DefenderDirectoryResult("disabled", "Windows Defender scanning is disabled")

        log_hint = self._read_recent_mpcmdrun_log()
        if log_hint:
            combined = combined + "\n" + log_hint

        if self._output_indicates_threat(combined):
            return DefenderDirectoryResult("ready", "MpCmdRun reported threats — verifying detections")

        if completed.returncode not in (0, 2):
            detail = combined.strip() or f"MpCmdRun exit code {completed.returncode}"
            if "0x80004005" in detail and "disabled" in detail.lower():
                return DefenderDirectoryResult("disabled", "Windows Defender scan unavailable (disabled)")
            logger.warning("Windows Defender MpCmdRun non-zero exit: %s", detail[:300])
        return DefenderDirectoryResult("ready", "MpCmdRun scan finished")

    def _run_powershell_scan(self, scan_dir: Path) -> DefenderDirectoryResult:
        escaped = str(scan_dir).replace("'", "''")
        script = f"""
$ErrorActionPreference = 'Stop'
$scanDir = '{escaped}'
try {{
    Start-MpScan -ScanType CustomScan -ScanPath $scanDir
    Write-Output 'STATUS|ready|Start-MpScan completed'
    exit 0
}} catch {{
    $msg = $_.Exception.Message
    if ($msg -match 'disabled') {{
        Write-Output 'STATUS|disabled|Windows Defender scan disabled'
        exit 10
    }}
    Write-Output ('STATUS|error|' + $msg)
    exit 11
}}
"""
        output, exit_code = self._run_powershell(script, timeout=self.timeout_seconds)
        for line in output.splitlines():
            if line.startswith("STATUS|"):
                _, status, detail = (line.split("|", 2) + [""])[:3]
                if status == "ready":
                    return DefenderDirectoryResult("ready", detail)
                return DefenderDirectoryResult(status, detail)
        return DefenderDirectoryResult("error", output.strip() or "Start-MpScan failed")

    def _collect_threats(self, scan_dir: Path) -> Dict[str, str]:
        escaped = str(scan_dir).replace("'", "''")
        script = f"""
$ErrorActionPreference = 'SilentlyContinue'
$scanDir = '{escaped}'.ToLower()
Get-MpThreatDetection | ForEach-Object {{
    $resource = [string]$_.Resources
    $threat = [string]$_.ThreatName
    if ([string]::IsNullOrWhiteSpace($resource)) {{ return }}
    if ($resource.ToLower().Contains($scanDir)) {{
        Write-Output ("THREAT|" + $resource + "|" + $threat)
    }}
}}
"""
        output, _ = self._run_powershell(script, timeout=30)
        threats: Dict[str, str] = {}
        scan_prefix = str(scan_dir).lower()

        for line in output.splitlines():
            if not line.startswith("THREAT|"):
                continue
            parts = line.split("|", 2)
            if len(parts) < 3:
                continue
            resource, threat_name = parts[1], parts[2]
            filename = self._filename_from_resource(resource, scan_dir)
            if filename:
                threats[filename] = threat_name or "Unknown threat"

        if not threats:
            threats.update(self._threats_from_mpcmdrun_log(scan_dir))
        return threats

    @staticmethod
    def _filename_from_resource(resource: str, scan_dir: Path) -> Optional[str]:
        normalized = resource.replace("file:_", "").replace("file:", "").strip()
        normalized = normalized.split("->")[-1].strip()
        try:
            path = Path(normalized)
            if scan_dir.resolve() in path.resolve().parents or path.parent.resolve() == scan_dir.resolve():
                return path.name
        except OSError:
            pass

        lower_resource = resource.lower()
        for candidate in scan_dir.iterdir():
            if candidate.name.lower() in lower_resource:
                return candidate.name
        return None

    def _find_mpcmdrun(self) -> Optional[Path]:
        if self._mpcmdrun_path and self._mpcmdrun_path.exists():
            return self._mpcmdrun_path

        for candidate in MPCMDRUN_CANDIDATES:
            if candidate.exists():
                self._mpcmdrun_path = candidate
                return candidate

        platform_root = Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Windows Defender Platform"
        if platform_root.is_dir():
            versions = sorted(platform_root.glob("*/MpCmdRun.exe"), reverse=True)
            if versions:
                self._mpcmdrun_path = versions[0]
                return versions[0]
        return None

    @staticmethod
    def _run_powershell(script: str, timeout: int = 60) -> tuple[str, int]:
        cmd = [
            "powershell.exe",
            "-NoProfile",
            "-NonInteractive",
            "-WindowStyle",
            "Hidden",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            script,
        ]
        try:
            completed = _hidden_run(cmd, timeout=timeout)
            output = "\n".join(filter(None, [completed.stdout, completed.stderr])).strip()
            return output, completed.returncode
        except subprocess.TimeoutExpired:
            return "PowerShell timed out", 124
        except OSError as exc:
            return str(exc), 127

    @staticmethod
    def _read_recent_mpcmdrun_log() -> str:
        log_path = Path(tempfile.gettempdir()) / "MpCmdRun.log"
        if not log_path.exists():
            return ""
        try:
            return log_path.read_text(encoding="utf-8", errors="ignore")[-4000:]
        except OSError:
            return ""

    @staticmethod
    def _output_indicates_threat(text: str) -> bool:
        markers = (
            "threat found",
            "threat(s) found",
            "detected threat",
            "malware",
            "0x80508007",
            "0x80508023",
        )
        lower = text.lower()
        return any(marker in lower for marker in markers)

    def _threats_from_mpcmdrun_log(self, scan_dir: Path) -> Dict[str, str]:
        log_text = self._read_recent_mpcmdrun_log()
        if not log_text or str(scan_dir).lower() not in log_text.lower():
            return {}

        threats: Dict[str, str] = {}
        threat_name = "Windows Defender detection"
        name_match = re.search(r"Threat\s+([^\r\n]+)", log_text, re.IGNORECASE)
        if name_match:
            threat_name = name_match.group(1).strip()

        for candidate in scan_dir.iterdir():
            if candidate.name.lower() in log_text.lower():
                threats[candidate.name] = threat_name
        return threats
