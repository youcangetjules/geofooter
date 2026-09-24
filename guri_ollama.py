#!/usr/bin/env python3
"""Ollama HTTP client + settings for GURI."""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

DEFAULT_BASE_URL = "http://localhost:11434"
DEFAULT_MODEL = ""
DEFAULT_SYSTEM = (
    "You are an assistant for Aliniant Smart Ass Email / GURI. "
    "Summarize emails, extract action items the user must take, "
    "and list deadlines clearly. Be concise and practical."
)


def _local_geofooter() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    if local:
        return Path(local) / "GeoFooter"
    return Path(r"C:\GeoFooter")


def config_path() -> Path:
    return _local_geofooter() / "guri_ollama.json"


def load_config() -> Dict[str, Any]:
    path = config_path()
    cfg = {
        "base_url": DEFAULT_BASE_URL,
        "model": DEFAULT_MODEL,
        "system": DEFAULT_SYSTEM,
        "timeout": 120,
    }
    try:
        if path.is_file():
            raw = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                cfg.update({k: raw[k] for k in cfg if k in raw})
    except Exception:
        pass
    cfg["base_url"] = str(cfg.get("base_url") or DEFAULT_BASE_URL).rstrip("/")
    return cfg


def save_config(cfg: Dict[str, Any]) -> None:
    path = config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "base_url": str(cfg.get("base_url") or DEFAULT_BASE_URL).rstrip("/"),
        "model": str(cfg.get("model") or ""),
        "system": str(cfg.get("system") or DEFAULT_SYSTEM),
        "timeout": int(cfg.get("timeout") or 120),
    }
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def _requests():
    import requests

    return requests


def ping(base_url: str = "", timeout: float = 5.0) -> Tuple[bool, str]:
    """Return (ok, message) for Ollama root / tags reachability."""
    url = (base_url or DEFAULT_BASE_URL).rstrip("/")
    try:
        r = _requests().get(f"{url}/api/tags", timeout=timeout)
        if r.status_code == 200:
            models = r.json().get("models") or []
            return True, f"Connected — {len(models)} model(s) available"
        return False, f"HTTP {r.status_code}: {r.text[:200]}"
    except Exception as exc:
        return False, f"Unreachable ({type(exc).__name__}: {exc})"


def list_models(base_url: str = "", timeout: float = 10.0) -> List[str]:
    url = (base_url or DEFAULT_BASE_URL).rstrip("/")
    r = _requests().get(f"{url}/api/tags", timeout=timeout)
    r.raise_for_status()
    names: List[str] = []
    for m in r.json().get("models") or []:
        name = str(m.get("name") or m.get("model") or "").strip()
        if name:
            names.append(name)
    return names


def chat(
    prompt: str,
    *,
    base_url: str = "",
    model: str = "",
    system: str = "",
    timeout: float = 120.0,
) -> str:
    """Non-streaming chat completion via /api/chat."""
    url = (base_url or DEFAULT_BASE_URL).rstrip("/")
    if not model:
        raise ValueError("No Ollama model selected")
    if not (prompt or "").strip():
        raise ValueError("Prompt is empty")

    messages: List[Dict[str, str]] = []
    sys_text = (system or "").strip()
    if sys_text:
        messages.append({"role": "system", "content": sys_text})
    messages.append({"role": "user", "content": prompt.strip()})

    r = _requests().post(
        f"{url}/api/chat",
        json={"model": model, "messages": messages, "stream": False},
        timeout=timeout,
    )
    r.raise_for_status()
    data = r.json()
    msg = data.get("message") or {}
    content = msg.get("content")
    if content:
        return str(content)
    # Fallback for generate-style responses
    if data.get("response"):
        return str(data["response"])
    return json.dumps(data, indent=2)


def build_email_analysis_prompt(
    mail: Dict[str, Any],
    reasons: Optional[List[str]] = None,
    *,
    user_rules: Optional[str] = None,
) -> str:
    """Build a prompt to extract actions from a scraped email."""
    subject = mail.get("subject") or ""
    sender = mail.get("sender") or ""
    when = mail.get("received") or ""
    body = (mail.get("body_preview") or "")[:4000]
    reason_txt = "\n".join(f"- {r}" for r in (reasons or [])) or "(none)"
    rules_block = (user_rules or "").strip()
    rules_section = f"\n{rules_block}\n\n" if rules_block else "\n"
    return (
        "Analyze this email and list concrete actions I need to take.\n"
        "Return:\n"
        "1) One-line summary\n"
        "2) Action items (bullets, with owner = me when applicable)\n"
        "3) Deadlines / urgency\n"
        "4) Suggested reply tone (one line)\n"
        "5) How your user-defined rules affected this analysis (if any)\n"
        f"{rules_section}"
        f"From: {sender}\n"
        f"When: {when}\n"
        f"Subject: {subject}\n"
        f"Heuristic flags:\n{reason_txt}\n\n"
        f"Body:\n{body}\n"
    )


def build_rule_compile_prompt(
    rule_text: str,
    *,
    scope: str = "always",
    sender: str = "",
    subject: str = "",
    body_preview: str = "",
) -> str:
    """Ask the LLM to compile a free-text rule into match + effects JSON."""
    return (
        "Compile this GURI triage rule into JSON only (no markdown fences).\n"
        "Schema:\n"
        "{\n"
        '  "match": {\n'
        '    "domains": ["example.com"],\n'
        '    "senders": ["a@b.com"],\n'
        '    "subject_contains": ["catch-up"],\n'
        '    "body_contains": ["let me know"],\n'
        '    "any_keywords": ["o-ran"]\n'
        "  },\n"
        '  "effects": {\n'
        '    "category": "actionable|relevant|less|deprecated|",\n'
        '    "importance": "high|medium|low|",\n'
        '    "boost_score": 0,\n'
        '    "boost_sender": 0,\n'
        '    "demote_actions": ["reply","approve","rsvp","review","deadline"],\n'
        '    "hide_from_actions": false,\n'
        '    "hide_from_important": false,\n'
        '    "skip_deadline": false,\n'
        '    "force_vip": false\n'
        "  },\n"
        '  "notes": "one-line rationale"\n'
        "}\n"
        "Use empty strings/arrays when not applicable. Prefer precise matchers.\n"
        "category=relevant means keep visible but not actionable.\n\n"
        f"Scope: {scope}\n"
        f"Sender context: {sender or '(none)'}\n"
        f"Subject context: {subject or '(none)'}\n"
        f"Body snippet: {(body_preview or '')[:800]}\n\n"
        f"Rule text:\n{rule_text.strip()}\n"
    )


def build_conversation_rank_prompt(
    query: Dict[str, Any],
    threads: List[Dict[str, Any]],
) -> str:
    """Ask the LLM to rank email *conversations* against a thematic query."""
    lines = [
        "Rank email CONVERSATIONS (whole threads), not isolated messages.",
        "The user query may describe sentiment, subject matter, or both.",
        "Score how well each thread as a whole matches the query.",
        "Return JSON only: an array of objects with keys:",
        '  id (int), score (0-100), why (short), sentiment (label), subject_matter (string)',
        "Do not invent messages that are not in the thread summaries.",
        "",
        f"Query: {query.get('raw') or ''}",
        f"Requested sentiment: {', '.join(query.get('sentiments') or []) or '(any)'}",
        f"Subject matter / topics: {', '.join(query.get('topics') or []) or '(open)'}",
        f"Keywords: {', '.join(query.get('keywords') or []) or '(none)'}",
        "",
        "Threads:",
    ]
    for th in threads:
        idx = th.get("id")
        n = th.get("message_count") or len(th.get("messages") or [])
        people = ", ".join(th.get("participants") or [])
        snippet = (th.get("summary") or th.get("snippet") or "")[:900]
        lines.append(
            f"[{idx}] topic={th.get('topic') or ''} | msgs={n} | "
            f"heuristic_sentiment={th.get('sentiment') or ''} | people={people}\n"
            f"{snippet}\n"
        )
    return "\n".join(lines)


def parse_json_array(text: str) -> List[Dict[str, Any]]:
    """Extract a JSON array of objects from an LLM response."""
    raw = (text or "").strip()
    if not raw:
        return []
    fence = re.search(r"```(?:json)?\s*(\[.*?\])\s*```", raw, re.DOTALL | re.IGNORECASE)
    if fence:
        raw = fence.group(1)
    else:
        start = raw.find("[")
        end = raw.rfind("]")
        if start >= 0 and end > start:
            raw = raw[start : end + 1]
    try:
        data = json.loads(raw)
    except Exception:
        return []
    if not isinstance(data, list):
        return []
    return [row for row in data if isinstance(row, dict)]


def parse_compiled_rule_json(text: str) -> Dict[str, Any]:
    """Extract {match, effects, notes} from an LLM response."""
    raw = (text or "").strip()
    if not raw:
        return {}
    # Strip common markdown fences
    fence = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", raw, re.DOTALL | re.IGNORECASE)
    if fence:
        raw = fence.group(1)
    else:
        start = raw.find("{")
        end = raw.rfind("}")
        if start >= 0 and end > start:
            raw = raw[start : end + 1]
    try:
        data = json.loads(raw)
    except Exception:
        return {}
    if not isinstance(data, dict):
        return {}
    return {
        "match": data.get("match") if isinstance(data.get("match"), dict) else {},
        "effects": data.get("effects") if isinstance(data.get("effects"), dict) else {},
        "notes": str(data.get("notes") or "").strip(),
    }


def _fmt_bytes(n: int) -> str:
    try:
        n = int(n)
    except Exception:
        return str(n)
    for unit, div in (("GB", 1024**3), ("MB", 1024**2), ("KB", 1024)):
        if n >= div:
            return f"{n / div:.2f} {unit}"
    return f"{n} B"


def _run_cmd(args: List[str], timeout: float = 8.0) -> Tuple[int, str, str]:
    try:
        proc = subprocess.run(
            args,
            capture_output=True,
            text=True,
            timeout=timeout,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
        )
        return proc.returncode, proc.stdout or "", proc.stderr or ""
    except FileNotFoundError:
        return 127, "", "not found"
    except subprocess.TimeoutExpired:
        return 124, "", "timeout"
    except Exception as exc:
        return 1, "", f"{type(exc).__name__}: {exc}"


def probe_nvidia_smi() -> Dict[str, Any]:
    """Query host NVIDIA GPUs + driver / CUDA version via nvidia-smi."""
    result: Dict[str, Any] = {
        "available": False,
        "gpus": [],
        "driver": "",
        "cuda_version": "",
        "error": "",
    }
    if not shutil.which("nvidia-smi"):
        result["error"] = "nvidia-smi not found on PATH (NVIDIA driver not installed?)"
        return result

    # CUDA Version appears in the human-readable header
    code, out, err = _run_cmd(["nvidia-smi"], timeout=10.0)
    if code != 0:
        result["error"] = (err or out or f"nvidia-smi exit {code}").strip()[:300]
        return result
    m = re.search(r"CUDA Version:\s*([\d.]+)", out)
    if m:
        result["cuda_version"] = m.group(1)

    code, out, err = _run_cmd(
        [
            "nvidia-smi",
            "--query-gpu=index,name,driver_version,memory.total,memory.used,memory.free,utilization.gpu,compute_cap",
            "--format=csv,noheader,nounits",
        ],
        timeout=10.0,
    )
    if code != 0:
        result["error"] = (err or out or f"nvidia-smi query exit {code}").strip()[:300]
        return result

    gpus = []
    for line in out.strip().splitlines():
        parts = [p.strip() for p in line.split(",")]
        if len(parts) < 8:
            continue
        gpus.append(
            {
                "index": parts[0],
                "name": parts[1],
                "driver": parts[2],
                "memory_total_mib": parts[3],
                "memory_used_mib": parts[4],
                "memory_free_mib": parts[5],
                "utilization_pct": parts[6],
                "compute_cap": parts[7],
            }
        )
        if not result["driver"] and parts[2]:
            result["driver"] = parts[2]
    result["gpus"] = gpus
    result["available"] = bool(gpus)
    if not gpus:
        result["error"] = "nvidia-smi ran but reported no GPUs"
    return result


def probe_nvcc() -> Dict[str, Any]:
    """Optional CUDA toolkit compiler check (nvcc)."""
    info: Dict[str, Any] = {"available": False, "version": "", "path": "", "error": ""}
    path = shutil.which("nvcc")
    if not path:
        # Common Windows CUDA install location
        for candidate in (
            r"C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA\v12.8\bin\nvcc.exe",
            r"C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA\v12.6\bin\nvcc.exe",
            r"C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA\v12.4\bin\nvcc.exe",
            r"C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA\v12.2\bin\nvcc.exe",
            r"C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA\v11.8\bin\nvcc.exe",
        ):
            if os.path.isfile(candidate):
                path = candidate
                break
    if not path:
        info["error"] = "nvcc not found (CUDA toolkit optional — driver CUDA is enough for Ollama)"
        return info
    info["path"] = path
    code, out, err = _run_cmd([path, "--version"], timeout=8.0)
    blob = (out or "") + "\n" + (err or "")
    m = re.search(r"release\s+([\d.]+)", blob, re.IGNORECASE)
    if m:
        info["version"] = m.group(1)
        info["available"] = True
    else:
        info["error"] = (blob.strip() or f"nvcc exit {code}")[:200]
    return info


def ollama_running_models(base_url: str = "", timeout: float = 8.0) -> List[Dict[str, Any]]:
    """Return models currently loaded in Ollama (/api/ps)."""
    url = (base_url or DEFAULT_BASE_URL).rstrip("/")
    r = _requests().get(f"{url}/api/ps", timeout=timeout)
    r.raise_for_status()
    return list(r.json().get("models") or [])


def _warm_model_for_gpu_check(
    base_url: str,
    model: str,
    timeout: float = 90.0,
) -> Tuple[bool, str]:
    """Tiny generate so the model loads; then /api/ps can report size_vram."""
    url = base_url.rstrip("/")
    try:
        r = _requests().post(
            f"{url}/api/generate",
            json={
                "model": model,
                "prompt": "ping",
                "stream": False,
                "keep_alive": "45s",
                "options": {"num_predict": 1},
            },
            timeout=timeout,
        )
        if r.status_code >= 400:
            return False, f"generate HTTP {r.status_code}: {r.text[:200]}"
        return True, "model warmed"
    except Exception as exc:
        return False, f"{type(exc).__name__}: {exc}"


def probe_gpu(
    base_url: str = "",
    model: str = "",
    *,
    warm: bool = True,
    timeout: float = 90.0,
) -> Dict[str, Any]:
    """Full GPU / CUDA / Ollama VRAM connection test.

    Checks:
      1) Host NVIDIA GPU via nvidia-smi (driver + CUDA runtime version)
      2) Optional CUDA toolkit (nvcc)
      3) Ollama reachability
      4) Whether a model is resident on GPU (size_vram > 0 via /api/ps)
         — optionally warms the selected model first
    """
    url = (base_url or DEFAULT_BASE_URL).rstrip("/")
    lines: List[str] = []
    report: Dict[str, Any] = {
        "ok": False,
        "gpu_ok": False,
        "cuda_ok": False,
        "ollama_gpu": False,
        "summary": "",
        "lines": lines,
        "nvidia": {},
        "nvcc": {},
        "ollama": {},
    }

    lines.append("=== GPU / CUDA connection test ===")
    nvidia = probe_nvidia_smi()
    report["nvidia"] = nvidia
    if nvidia.get("available"):
        report["gpu_ok"] = True
        report["cuda_ok"] = bool(nvidia.get("cuda_version"))
        lines.append(
            f"NVIDIA driver: {nvidia.get('driver') or '—'}  ·  "
            f"CUDA (driver): {nvidia.get('cuda_version') or '—'}"
        )
        for g in nvidia.get("gpus") or []:
            lines.append(
                f"  GPU {g.get('index')}: {g.get('name')}  ·  "
                f"VRAM {g.get('memory_used_mib')}/{g.get('memory_total_mib')} MiB used  ·  "
                f"util {g.get('utilization_pct')}%  ·  CC {g.get('compute_cap')}"
            )
    else:
        lines.append(f"NVIDIA: FAIL — {nvidia.get('error') or 'no GPU'}")

    nvcc = probe_nvcc()
    report["nvcc"] = nvcc
    if nvcc.get("available"):
        lines.append(f"CUDA toolkit (nvcc): {nvcc.get('version')}  ·  {nvcc.get('path')}")
    else:
        lines.append(f"CUDA toolkit (nvcc): not required — {nvcc.get('error') or 'absent'}")

    # Ollama reachability
    ok, ping_msg = ping(url, timeout=5.0)
    ollama_info: Dict[str, Any] = {
        "reachable": ok,
        "ping": ping_msg,
        "version": "",
        "models_loaded": [],
        "warm_error": "",
        "size_vram": 0,
        "processor": "",
    }
    if not ok:
        lines.append(f"Ollama: FAIL — {ping_msg}")
        report["ollama"] = ollama_info
        report["summary"] = "GPU host OK, but Ollama unreachable" if report["gpu_ok"] else "GPU / Ollama check failed"
        report["lines"] = lines
        return report

    try:
        vr = _requests().get(f"{url}/api/version", timeout=5.0)
        if vr.status_code == 200:
            ollama_info["version"] = str((vr.json() or {}).get("version") or "")
    except Exception:
        pass
    lines.append(
        f"Ollama: connected{(' v' + ollama_info['version']) if ollama_info['version'] else ''} — {ping_msg}"
    )

    # Choose model to warm (prefer a small tag when none selected)
    use_model = (model or "").strip()
    if not use_model:
        try:
            names = list_models(url, timeout=8.0)
            preferred = (
                "gemma3:4b",
                "llama3.2:latest",
                "llama3.2",
                "qwen2.5:latest",
                "llama3:latest",
            )
            use_model = next((n for n in preferred if n in names), "")
            if not use_model:
                # Avoid accidentally warming a 70B just for a connectivity probe
                light = [
                    n
                    for n in names
                    if not re.search(r"70b|65b|405b|32b", n, re.IGNORECASE)
                ]
                use_model = (light or names or [""])[0]
        except Exception:
            use_model = ""

    if warm and use_model:
        lines.append(f"Warming model on GPU: {use_model} …")
        warmed, warm_msg = _warm_model_for_gpu_check(url, use_model, timeout=timeout)
        if not warmed:
            ollama_info["warm_error"] = warm_msg
            lines.append(f"  Warm failed: {warm_msg}")
        else:
            lines.append(f"  {warm_msg}")
    elif warm and not use_model:
        lines.append("No model selected / installed — skipped VRAM load test.")

    # Inspect /api/ps for VRAM
    try:
        running = ollama_running_models(url, timeout=8.0)
        ollama_info["models_loaded"] = running
        total_vram = 0
        for m in running:
            name = m.get("name") or m.get("model") or "?"
            size = int(m.get("size") or 0)
            vram = int(m.get("size_vram") or 0)
            total_vram += vram
            if size > 0 and vram > 0:
                pct = min(100, int(round(100.0 * vram / size)))
                proc = f"{pct}% GPU" if vram >= size else f"{pct}% GPU / {100 - pct}% CPU"
            elif vram > 0:
                proc = "GPU"
            else:
                proc = "CPU only"
            ollama_info["processor"] = proc
            lines.append(
                f"  Loaded: {name}  ·  size {_fmt_bytes(size)}  ·  "
                f"VRAM {_fmt_bytes(vram)}  ·  {proc}"
            )
        ollama_info["size_vram"] = total_vram
        if total_vram > 0:
            report["ollama_gpu"] = True
            lines.append("Ollama GPU: PASS — model resident in VRAM")
        elif running:
            lines.append("Ollama GPU: FAIL — model loaded but size_vram=0 (CPU fallback)")
        else:
            lines.append(
                "Ollama GPU: no model currently loaded "
                "(select a model and re-run Test GPU to verify VRAM)"
            )
    except Exception as exc:
        lines.append(f"Ollama /api/ps failed: {exc}")

    report["ollama"] = ollama_info
    report["ok"] = bool(report["gpu_ok"] and ok and (report["ollama_gpu"] or not warm or not use_model))
    if report["ollama_gpu"]:
        report["summary"] = "GPU + CUDA driver OK — Ollama using VRAM"
    elif report["gpu_ok"] and ok:
        report["summary"] = "GPU + CUDA driver OK — Ollama reachable (VRAM not confirmed)"
    else:
        report["summary"] = "GPU / CUDA test incomplete — see details"
    report["lines"] = lines
    return report


def format_gpu_report(report: Dict[str, Any]) -> str:
    """Human-readable multi-line report for the UI."""
    lines = list(report.get("lines") or [])
    summary = str(report.get("summary") or "")
    if summary:
        lines.append("")
        lines.append(f"Summary: {summary}")
    return "\n".join(lines)
