# MCP Client

A simple client for connecting to Model Context Protocol (MCP) servers, allowing you to interact with tools exposed by MCP servers.

## Installation

1. Ensure you have Python 3.6+ installed
2. Create and activate a virtual environment:

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# Linux/macOS
source .venv/bin/activate
```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

Run the client:

```bash
python mcp_client.py
```

The client will:
1. Connect to all servers defined in `config.json`
2. Start an interactive mode where you can run commands

## Interactive Commands

- `help` - Show available commands
- `servers` - List available servers from config
- `connect <server>` - Connect to a specific server
- `tools <server>` - List tools for a connected server
- `info <server> [tool]` - Get detailed info about server tools
- `call <server> <tool> [json_args]` - Call a specific tool with arguments
- `exit` or `quit` - Exit the client

## Configuration

Edit the `config.json` file to define your MCP servers. Example:

```json
{
  "mcpServers": {
    "file-system": {
      "command": "uv",
      "args": [
        "run",
        "--with",
        "mcp[cli]",
        "path/to/server.py"
      ]
    }
  }
}
```

Each server should have:
- `command` - The command to run the server
- `args` - Array of arguments to pass to the command
