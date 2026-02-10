# PSExtender (PowerShell MCP Bridge)

This repo contains a very small bridge that lets a Linux/WSL process (speaking MCP over stdin/stdout) execute PowerShell commands on Windows via a TCP listener.

It is intentionally minimal and **has no authentication and no encryption**. Read **Security / Risks** before running.

## Components

- `for_windows.py`
  - Runs on Windows.
  - Listens on a TCP port and executes incoming PowerShell commands.
  - Protocol: newline-delimited JSON request/response.
- `linux_mcp_powershell_bridge.py`
  - Runs on Linux/WSL.
  - Implements an MCP tool server with one tool: `powershell`.
  - Forwards tool calls to the Windows listener over TCP.

## Requirements

- Windows:
  - PowerShell (`powershell.exe`) or PowerShell 7 (`pwsh`)
  - Python 3
- Linux/WSL:
  - Python 3

## Quick Start (WSL talking to Windows)

### 1) Start the listener on Windows

Run on Windows in a terminal:

```powershell
python for_windows.py --host 127.0.0.1 --port 8765
```

Notes:

- `--host 127.0.0.1` binds to localhost only. This is the safest default.
- If you bind to `0.0.0.0` or a LAN IP, you are exposing remote command execution to anything that can reach that port.

Optional environment variables (Windows):

- `POWERSHELL_EXE`: full path to `pwsh` or `powershell.exe` if auto-detection is wrong
- `PS_LISTEN_HOST`, `PS_LISTEN_PORT`: defaults used by `for_windows.py` if you do not pass `--host/--port`

### 2) Run the MCP bridge inside WSL/Linux

Run on Linux/WSL:

```bash
python3 linux_mcp_powershell_bridge.py
```

Optional environment variables (Linux/WSL):

- `PS_LISTEN_HOST`: hostname/IP for the Windows listener (overrides auto-detection)
- `PS_LISTEN_PORT`: port for the Windows listener (default `8765`)

`linux_mcp_powershell_bridge.py` will try to find the Windows host automatically. In WSL, it typically uses the `nameserver` value from `/etc/resolv.conf`, then falls back to a small list of common vEthernet IPs plus `localhost`.

### 3) Use it from an MCP client

The bridge exposes one tool:

- Tool name: `powershell`
- Input: `{ "command": "..." }`

Example PowerShell command:

```powershell
Get-ComputerInfo | Select-Object -First 3
```

How you configure an MCP client to launch this server depends on the client. The server communicates via stdin/stdout and supports:

- JSONL framing (one JSON object per line)
- `Content-Length:` framing (for clients that use it)

## Testing Without MCP (direct TCP protocol)

The Windows listener accepts a single-line JSON object like:

```json
{"command":"Write-Output 'hello from powershell'"}
```

and replies with a single-line JSON object containing:

- `ok` (boolean)
- `stdout` (string)
- `stderr` (string)
- `code` (int process exit code)

## Security / Risks (read before running)

This project is a remote command execution bridge.

- No authentication:
  - Anyone who can connect to the Windows listener can run arbitrary PowerShell commands.
- No encryption:
  - Commands and outputs are sent in plaintext over TCP.
- Default binding is dangerous:
  - `for_windows.py` defaults to binding `0.0.0.0` if `PS_LISTEN_HOST` is not set and `--host` is not passed.
  - Binding to `0.0.0.0` exposes the service on all network interfaces.
- Privilege and lateral movement:
  - Commands run as the Windows user running `for_windows.py`. If that user is admin, the impact is full system compromise.

Minimum safe operating guidance:

1. Bind the Windows listener to localhost only:
   - `python for_windows.py --host 127.0.0.1 --port 8765`
2. If you must bind to a LAN IP:
   - Use Windows Firewall to restrict inbound connections to a specific trusted source.
   - Never expose the port to the public internet.
3. Run as a least-privileged user and assume full compromise if the port is reachable by an attacker.

## Project Notes

- The `Powershell-MCP/` folder exists in this repo and may contain related upstream/project material, but the functional bridge is implemented by the two Python files in the repo root.

