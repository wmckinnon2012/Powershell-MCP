#!/usr/bin/env python3
import argparse
import json
import socket
import sys
import time
from typing import Any, Dict, List, Optional, Tuple


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


def _connect(host: str, port: int) -> socket.socket:
    while True:
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
        time.sleep(1.0)


def _send_request(sock: socket.socket, payload: Dict[str, Any]) -> Dict[str, Any]:
    encoded = json.dumps(payload, ensure_ascii=True) + "\n"
    sock.sendall(encoded.encode("utf-8"))

    with sock.makefile("r", encoding="utf-8", newline="") as reader:
        line = reader.readline()
        if not line:
            raise ConnectionError("Windows listener disconnected")
        return json.loads(line)


def _extract_commands(arguments: Dict[str, Any]) -> Optional[List[str]]:
    command = arguments.get("command")
    commands = arguments.get("commands")

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


def _format_execution_text(result: Dict[str, Any]) -> str:
    status = str(result.get("status", "unknown"))
    command_count = result.get("command_count")
    header = f"status: {status}"
    if isinstance(command_count, int):
        header += f", commands: {command_count}"

    stdout = (result.get("stdout") or "").rstrip("\n")
    stderr = (result.get("stderr") or "").rstrip("\n")
    sections = [header]
    if stdout:
        sections.append(stdout)
    if stderr:
        sections.append("[stderr]\n" + stderr)
    if len(sections) == 1:
        sections.append("(no output)")
    return "\n".join(sections)


def _format_status_text(result: Dict[str, Any]) -> str:
    status = str(result.get("status", "unknown"))
    job_id = result.get("job_id") or "(unknown)"
    lines = [f"job_id: {job_id}", f"status: {status}"]

    if isinstance(result.get("command_count"), int):
        lines.append(f"commands: {result['command_count']}")
    if isinstance(result.get("started_at"), str):
        lines.append(f"started_at: {result['started_at']}")
    if isinstance(result.get("finished_at"), str):
        lines.append(f"finished_at: {result['finished_at']}")
    if isinstance(result.get("elapsed_seconds"), (int, float)):
        lines.append(f"elapsed_seconds: {result['elapsed_seconds']:.1f}")

    inner = result.get("result")
    if isinstance(inner, dict):
        lines.append(_format_execution_text(inner))
    elif result.get("stderr"):
        lines.append("[stderr]\n" + str(result.get("stderr")))

    return "\n".join(lines)


def _handle_tools_list(msg_id: Any, framing: str) -> None:
    _write_message(
        {
            "jsonrpc": "2.0",
            "id": msg_id,
            "result": {
                "tools": [
                    {
                        "name": "powershell",
                        "description": "Run one or more Windows PowerShell commands via network bridge.",
                        "inputSchema": {
                            "type": "object",
                            "properties": {
                                "command": {
                                    "type": "string",
                                    "description": "Single PowerShell command to execute.",
                                },
                                "commands": {
                                    "type": "array",
                                    "items": {"type": "string"},
                                    "minItems": 1,
                                    "description": "List of PowerShell commands to run in order.",
                                },
                                "async": {
                                    "type": "boolean",
                                    "description": "If true, return immediately with job_id for status polling.",
                                }
                            },
                            "additionalProperties": False,
                        },
                    },
                    {
                        "name": "powershell_status",
                        "description": "Get status/output for a long-running PowerShell job.",
                        "inputSchema": {
                            "type": "object",
                            "properties": {
                                "job_id": {
                                    "type": "string",
                                    "description": "Job ID returned by async powershell call.",
                                }
                            },
                            "required": ["job_id"],
                            "additionalProperties": False,
                        },
                    },
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
    parser = argparse.ArgumentParser(
        description="MCP tool server that forwards PowerShell commands to a Windows TCP listener."
    )
    parser.add_argument(
        "--host",
        required=True,
        help="Windows listener host/IP to connect to.",
    )
    parser.add_argument(
        "--port",
        type=int,
        required=True,
        help="Windows listener TCP port to connect to.",
    )
    args = parser.parse_args()

    host = args.host
    port = args.port

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

            if name == "powershell":
                commands = _extract_commands(arguments)
                if commands is None:
                    _error_response(
                        msg_id,
                        -32602,
                        "Provide 'command' string or non-empty 'commands' array of strings",
                        framing_mode,
                    )
                    continue

                async_mode = bool(arguments.get("async", False))
                payload: Dict[str, Any] = {"action": "run", "async": async_mode}
                if len(commands) == 1:
                    payload["command"] = commands[0]
                else:
                    payload["commands"] = commands

                if sock is None:
                    sock = _connect(host, port)

                try:
                    result = _send_request(sock, payload)
                except Exception:
                    try:
                        sock.close()
                    except Exception:
                        pass
                    sock = _connect(host, port)
                    result = _send_request(sock, payload)

                if async_mode:
                    status = str(result.get("status", "unknown"))
                    job_id = result.get("job_id") or "(unknown)"
                    text = f"status: {status}\njob_id: {job_id}"
                else:
                    text = _format_execution_text(result)
                _text_result(msg_id, text, is_error=not bool(result.get("ok")), framing=framing_mode)
                continue

            if name == "powershell_status":
                job_id = arguments.get("job_id")
                if not isinstance(job_id, str) or not job_id:
                    _error_response(msg_id, -32602, "Missing or invalid 'job_id'", framing_mode)
                    continue

                payload = {"action": "status", "job_id": job_id}
                if sock is None:
                    sock = _connect(host, port)

                try:
                    result = _send_request(sock, payload)
                except Exception:
                    try:
                        sock.close()
                    except Exception:
                        pass
                    sock = _connect(host, port)
                    result = _send_request(sock, payload)

                _text_result(
                    msg_id,
                    _format_status_text(result),
                    is_error=not bool(result.get("ok")),
                    framing=framing_mode,
                )
                continue

            _error_response(msg_id, -32601, "Tool not found", framing_mode)
            continue

        if method is None:
            _error_response(msg_id, -32600, "Invalid Request", framing_mode)
            continue

        _error_response(msg_id, -32601, "Method not found", framing_mode)


if __name__ == "__main__":
    main()
