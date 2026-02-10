#!/usr/bin/env python3
import argparse
import json
import os
import socket
import subprocess
import sys
from typing import Any, Dict


def _write_line(sock: socket.socket, obj: Dict[str, Any]) -> None:
    data = json.dumps(obj, ensure_ascii=True).encode("utf-8") + b"\n"
    sock.sendall(data)


def _read_line(sock_file) -> Dict[str, Any]:
    line = sock_file.readline()
    if not line:
        raise ConnectionError("client disconnected")
    return json.loads(line)


def _run_powershell(command: str) -> Dict[str, Any]:
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


def _handle_client(conn: socket.socket) -> None:
    with conn:
        conn_file = conn.makefile("r", encoding="utf-8", newline="")
        while True:
            msg = _read_line(conn_file)
            command = msg.get("command")
            if not isinstance(command, str):
                _write_line(
                    conn,
                    {
                        "ok": False,
                        "stdout": "",
                        "stderr": "Missing or invalid 'command'",
                        "code": 2,
                    },
                )
                continue

            result = _run_powershell(command)
            _write_line(conn, result)


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
    args = parser.parse_args()

    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    sock.bind((args.host, args.port))
    sock.listen(1)

    sys.stderr.write(f"PowerShell listener on {args.host}:{args.port}\n")
    sys.stderr.flush()

    while True:
        conn, _addr = sock.accept()
        try:
            _handle_client(conn)
        except Exception as exc:
            sys.stderr.write(f"Client error: {exc}\n")
            sys.stderr.flush()


if __name__ == "__main__":
    main()
