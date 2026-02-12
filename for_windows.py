#!/usr/bin/env python3
import argparse
from datetime import datetime, timezone
import json
import os
import socket
import subprocess
import sys
import threading
import uuid
from typing import Any, Dict, List, Optional


_jobs: Dict[str, Dict[str, Any]] = {}
_jobs_lock = threading.Lock()


def _log(message: str) -> None:
    sys.stderr.write(f"[listener] {message}\n")
    sys.stderr.flush()


def _preview_command(command: str, limit: int = 160) -> str:
    one_line = " ".join(command.splitlines()).strip()
    if len(one_line) <= limit:
        return one_line
    return one_line[: limit - 3] + "..."


def _write_line(sock: socket.socket, obj: Dict[str, Any]) -> None:
    data = json.dumps(obj, ensure_ascii=True).encode("utf-8") + b"\n"
    sock.sendall(data)


def _read_line(sock_file) -> Dict[str, Any]:
    line = sock_file.readline()
    if not line:
        raise ConnectionError("client disconnected")
    return json.loads(line)


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _run_single_powershell(command: str) -> Dict[str, Any]:
    _log(f"run command: {_preview_command(command)}")
    pwsh = os.environ.get("POWERSHELL_EXE")
    candidates = []
    if pwsh:
        candidates.append(pwsh)
    candidates.extend(["pwsh", "powershell.exe"])
    try:
        last_err = None
        for exe in candidates:
            try:
                proc = subprocess.run(
                    [exe, "-NoProfile", "-NonInteractive", "-Command", command],
                    capture_output=True,
                    text=True,
                )
                _log(
                    f"command finished: code={proc.returncode}, stdout={len(proc.stdout or '')} bytes, stderr={len(proc.stderr or '')} bytes"
                )
                return {
                    "ok": proc.returncode == 0,
                    "stdout": proc.stdout or "",
                    "stderr": proc.stderr or "",
                    "code": proc.returncode,
                }
            except FileNotFoundError as exc:
                last_err = exc
                continue
        return {
            "ok": False,
            "stdout": "",
            "stderr": "No PowerShell executable found (tried pwsh, powershell.exe).",
            "code": 127,
        }
    except Exception as exc:
        return {
            "ok": False,
            "stdout": "",
            "stderr": f"PowerShell execution failed: {exc}",
            "code": 1,
        }


def _run_powershell_batch(commands: List[str]) -> Dict[str, Any]:
    _log(f"batch start: {len(commands)} command(s)")
    results: List[Dict[str, Any]] = []
    for idx, command in enumerate(commands):
        result = _run_single_powershell(command)
        result["index"] = idx
        results.append(result)

    ok = all(bool(item.get("ok")) for item in results)
    code = 0
    for item in results:
        item_code = item.get("code")
        if isinstance(item_code, int) and item_code != 0:
            code = item_code
            break

    if len(results) == 1:
        first = results[0]
        return {
            "ok": ok,
            "stdout": first.get("stdout", ""),
            "stderr": first.get("stderr", ""),
            "code": int(first.get("code", code)),
            "results": results,
        }

    stdout_parts: List[str] = []
    stderr_parts: List[str] = []
    for item in results:
        idx = int(item.get("index", 0))
        out = item.get("stdout", "")
        err = item.get("stderr", "")
        if out:
            stdout_parts.append(f"[command {idx} stdout]\n{out.rstrip()}")
        if err:
            stderr_parts.append(f"[command {idx} stderr]\n{err.rstrip()}")

    return {
        "ok": ok,
        "stdout": "\n\n".join(stdout_parts) + ("\n" if stdout_parts else ""),
        "stderr": "\n\n".join(stderr_parts) + ("\n" if stderr_parts else ""),
        "code": code,
        "results": results,
    }


def _extract_commands(msg: Dict[str, Any]) -> Optional[List[str]]:
    command = msg.get("command")
    commands = msg.get("commands")

    if isinstance(command, str):
        return [command]

    if isinstance(commands, list) and commands:
        extracted = []
        for item in commands:
            if not isinstance(item, str):
                return None
            extracted.append(item)
        return extracted

    return None


def _status_from_job(job: Dict[str, Any]) -> Dict[str, Any]:
    status = str(job.get("status", "running"))
    started_at = str(job.get("started_at", ""))
    finished_at = job.get("finished_at")

    response = {
        "ok": status != "failed",
        "status": status,
        "job_id": job["job_id"],
        "started_at": started_at,
        "finished_at": finished_at,
        "command_count": int(job.get("command_count", 0)),
    }

    if status == "running":
        elapsed_seconds = max(0.0, datetime.now(timezone.utc).timestamp() - float(job.get("start_ts", 0.0)))
        response["elapsed_seconds"] = elapsed_seconds
        return response

    response["result"] = job.get("result") or {}
    return response


def _run_job(job_id: str, commands: List[str]) -> None:
    _log(f"job {job_id} started ({len(commands)} command(s))")
    result = _run_powershell_batch(commands)
    finished_at = _now_iso()
    status = "completed" if bool(result.get("ok")) else "failed"
    with _jobs_lock:
        if job_id in _jobs:
            _jobs[job_id]["status"] = status
            _jobs[job_id]["finished_at"] = finished_at
            _jobs[job_id]["result"] = result
    _log(f"job {job_id} finished with status={status}")


def _start_async_job(commands: List[str]) -> Dict[str, Any]:
    job_id = str(uuid.uuid4())
    started_at = _now_iso()
    job = {
        "job_id": job_id,
        "status": "running",
        "started_at": started_at,
        "start_ts": datetime.now(timezone.utc).timestamp(),
        "finished_at": None,
        "command_count": len(commands),
        "result": None,
    }
    with _jobs_lock:
        _jobs[job_id] = job

    thread = threading.Thread(target=_run_job, args=(job_id, commands), daemon=True)
    thread.start()
    _log(f"job {job_id} queued ({len(commands)} command(s))")

    return {
        "ok": True,
        "status": "running",
        "job_id": job_id,
        "started_at": started_at,
        "finished_at": None,
        "command_count": len(commands),
    }


def _handle_request(msg: Dict[str, Any]) -> Dict[str, Any]:
    action = msg.get("action")
    _log(f"request received: action={action or 'run(default)'} keys={sorted(msg.keys())}")

    if action == "status":
        job_id = msg.get("job_id")
        if not isinstance(job_id, str) or not job_id:
            return {
                "ok": False,
                "status": "invalid",
                "stderr": "Missing or invalid 'job_id'",
                "code": 2,
            }
        with _jobs_lock:
            job = _jobs.get(job_id)
            if job is None:
                _log(f"status check: job {job_id} not found")
                return {
                    "ok": False,
                    "status": "not_found",
                    "stderr": f"Unknown job_id '{job_id}'",
                    "code": 3,
                }
            response = _status_from_job(job)
            _log(f"status check: job {job_id} -> {response.get('status')}")
            return response

    commands = _extract_commands(msg)
    if commands is None:
        return {
            "ok": False,
            "status": "invalid",
            "stdout": "",
            "stderr": "Missing or invalid 'command'/'commands'",
            "code": 2,
        }

    async_mode = bool(msg.get("async", False))
    if async_mode:
        return _start_async_job(commands)

    result = _run_powershell_batch(commands)
    _log(f"sync run finished: status={'completed' if bool(result.get('ok')) else 'failed'}")
    return {
        **result,
        "status": "completed" if bool(result.get("ok")) else "failed",
        "command_count": len(commands),
    }


def _handle_client(conn: socket.socket, idle_timeout_seconds: float) -> None:
    try:
        peer = conn.getpeername()
        _log(f"client connected: {peer[0]}:{peer[1]}")
    except Exception:
        _log("client connected: <unknown-peer>")

    conn.settimeout(idle_timeout_seconds)
    with conn:
        conn_file = conn.makefile("r", encoding="utf-8", newline="")
        while True:
            try:
                msg = _read_line(conn_file)
            except socket.timeout:
                _log(f"client idle timeout after {int(idle_timeout_seconds)}s; closing connection")
                break
            except ConnectionError:
                _log("client disconnected")
                break
            except json.JSONDecodeError as exc:
                _log(f"invalid JSON from client: {exc}")
                break
            except Exception as exc:
                if "timed out" in str(exc).lower():
                    _log(f"client idle timeout after {int(idle_timeout_seconds)}s; closing connection")
                    break
                _log(f"client read error: {exc}")
                break

            result = _handle_request(msg)
            _write_line(conn, result)
            _log(f"response sent: ok={result.get('ok')} status={result.get('status')}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Windows PowerShell listener for MCP bridge."
    )
    parser.add_argument(
        "--host",
        default=os.environ.get("PS_LISTEN_HOST", "0.0.0.0"),
        help="Host/IP to bind (default: 0.0.0.0).",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=int(os.environ.get("PS_LISTEN_PORT", "8765")),
        help="TCP port to bind (default: 8765).",
    )
    parser.add_argument(
        "--client-idle-timeout",
        type=float,
        default=float(os.environ.get("PS_CLIENT_IDLE_TIMEOUT", "300")),
        help="Seconds before an idle client connection is closed (default: 300).",
    )
    args = parser.parse_args()

    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    sock.bind((args.host, args.port))
    sock.listen(16)

    sys.stderr.write(f"PowerShell listener on {args.host}:{args.port}\n")
    sys.stderr.flush()

    while True:
        conn, _addr = sock.accept()
        thread = threading.Thread(
            target=_handle_client,
            args=(conn, args.client_idle_timeout),
            daemon=True,
        )
        thread.start()


if __name__ == "__main__":
    main()
