#!/usr/bin/env python3
import json
import os
import socket
import sys
import time
from typing import Any, Dict, Optional, Tuple


def _write_message(message: Dict[str, Any], framing: str) -> None:
    payload = json.dumps(message, ensure_ascii=True).encode("utf-8")
    if framing == "content-length":
        header = f"Content-Length: {len(payload)}\r\n\r\n".encode("ascii")
        sys.stdout.buffer.write(header + payload)
        sys.stdout.buffer.flush()
    else:
        sys.stdout.write(payload.decode("utf-8") + "\n")
        sys.stdout.flush()


def _error_response(msg_id: Any, code: int, message: str, framing: str) -> None:
    if msg_id is None:
        return
    _write_message(
        {
            "jsonrpc": "2.0",
            "id": msg_id,
            "error": {"code": code, "message": message},
        },
        framing,
    )


def _text_result(msg_id: Any, text: str, is_error: bool = False, framing: str = "jsonl") -> None:
    if msg_id is None:
        return
    _write_message(
        {
            "jsonrpc": "2.0",
            "id": msg_id,
            "result": {
                "content": [{"type": "text", "text": text}],
                "isError": is_error,
            },
        },
        framing,
    )


def _connect(hosts: list, port: int) -> socket.socket:
    while True:
        for host in hosts:
            try:
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(2.0)
                sock.connect((host, port))
                sock.settimeout(None)
                return sock
            except OSError:
                try:
                    sock.close()
                except Exception:
                    pass
                continue
        time.sleep(1.0)


def _default_hosts() -> list:
    hosts = []
    env_host = os.environ.get("PS_LISTEN_HOST")
    if env_host:
        hosts.append(env_host)

    # In WSL, Windows host IP is usually the nameserver.
    try:
        with open("/etc/resolv.conf", "r", encoding="utf-8") as f:
            for line in f:
                if line.startswith("nameserver"):
                    parts = line.split()
                    if len(parts) >= 2:
                        hosts.append(parts[1])
                        break
    except OSError:
        pass

    # Common WSL vEthernet defaults.
    hosts.extend(["172.26.64.1", "172.27.48.1", "localhost"])
    # De-dupe while preserving order.
    seen = set()
    ordered = []
    for h in hosts:
        if h not in seen:
            seen.add(h)
            ordered.append(h)
    return ordered


def _send_command(
    sock: socket.socket, command: str
) -> Dict[str, Any]:
    payload = json.dumps({"command": command}, ensure_ascii=True) + "\n"
    sock.sendall(payload.encode("utf-8"))

    with sock.makefile("r", encoding="utf-8", newline="") as reader:
        line = reader.readline()
        if not line:
            raise ConnectionError("Windows listener disconnected")
        return json.loads(line)


def _handle_tools_list(msg_id: Any, framing: str) -> None:
    _write_message(
        {
            "jsonrpc": "2.0",
            "id": msg_id,
            "result": {
                "tools": [
                    {
                        "name": "powershell",
                        "description": "Run a Windows PowerShell command via network bridge.",
                        "inputSchema": {
                            "type": "object",
                            "properties": {
                                "command": {
                                    "type": "string",
                                    "description": "PowerShell command to execute.",
                                }
                            },
                            "required": ["command"],
                            "additionalProperties": False,
                        },
                    }
                ]
            },
        },
        framing,
    )


def _read_message() -> Tuple[Optional[Dict[str, Any]], str]:
    buf = sys.stdin.buffer
    line = buf.readline()
    if not line:
        return None, "jsonl"

    if line.startswith(b"Content-Length:"):
        framing = "content-length"
        try:
            length = int(line.split(b":", 1)[1].strip())
        except ValueError:
            return None, framing

        # Consume remaining headers until blank line
        while True:
            hdr = buf.readline()
            if not hdr or hdr in (b"\n", b"\r\n"):
                break

        body = buf.read(length)
        if not body:
            return None, framing
        try:
            return json.loads(body.decode("utf-8")), framing
        except json.JSONDecodeError:
            return None, framing

    # JSONL framing
    framing = "jsonl"
    line = line.strip()
    if not line:
        return None, framing
    try:
        return json.loads(line.decode("utf-8")), framing
    except json.JSONDecodeError:
        return None, framing


def main() -> None:
    hosts = _default_hosts()
    port = int(os.environ.get("PS_LISTEN_PORT", "8765"))

    sock: Optional[socket.socket] = None

    framing_mode = "jsonl"

    while True:
        msg, framing = _read_message()
        if msg is None:
            if sys.stdin.buffer.closed:
                break
            continue

        framing_mode = framing

        msg_id = msg.get("id")
        method = msg.get("method")

        if method == "initialize":
            _write_message(
                {
                    "jsonrpc": "2.0",
                    "id": msg_id,
                    "result": {
                        "protocolVersion": "2024-11-05",
                        "serverInfo": {
                            "name": "linux-mcp-powershell-bridge",
                            "version": "0.1.0",
                        },
                        "capabilities": {"tools": {}},
                    },
                },
                framing_mode,
            )
            continue

        if method == "tools/list":
            _handle_tools_list(msg_id, framing_mode)
            continue

        if method == "initialized":
            # Notification; no response expected.
            continue

        if method == "tools/call":
            params = msg.get("params") or {}
            name = params.get("name")
            arguments = params.get("arguments") or {}

            if name != "powershell":
                _error_response(msg_id, -32601, "Tool not found", framing_mode)
                continue

            command = arguments.get("command")
            if not isinstance(command, str):
                _error_response(
                    msg_id, -32602, "Missing or invalid 'command'", framing_mode
                )
                continue

            if sock is None:
                sock = _connect(hosts, port)

            try:
                result = _send_command(sock, command)
            except Exception:
                try:
                    sock.close()
                except Exception:
                    pass
                sock = _connect(hosts, port)
                result = _send_command(sock, command)

            stdout = (result.get("stdout") or "").rstrip("\n")
            stderr = (result.get("stderr") or "").rstrip("\n")
            ok = bool(result.get("ok"))

            text = stdout
            if stderr:
                text = (stdout + "\n" if stdout else "") + "[stderr]\n" + stderr
            if text == "":
                text = "(no output)"

            _text_result(msg_id, text, is_error=not ok, framing=framing_mode)
            continue

        if method is None:
            _error_response(msg_id, -32600, "Invalid Request", framing_mode)
            continue

        _error_response(msg_id, -32601, "Method not found", framing_mode)


if __name__ == "__main__":
    main()
