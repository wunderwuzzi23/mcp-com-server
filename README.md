# mcp-com-server

A Windows COM adapter for the Model Context Protocol (MCP).

Created by Johann Rehberger (@wunderwuzzi23)

## Experimental and Security Caution!

This is a very quick first version that serves as a proof-of-concept, but it already works pretty well.

There is only one very basic security feature, which is that you can ALLOWLIST CLSID and/or ProgIDs. Otherwise this is all very dangerous obviously - so use with caution!

## Description

mcp-com-server provides a bridge between Windows COM objects and the Model Context Protocol, enabling language models to interact with COM-based applications and services. This adapter exposes COM functionality through a standardized MCP interface.

## Features

- Create and manage COM object instances
- Invoke methods on COM objects
- Get and set properties
- Query interfaces
- Retrieve type information
- Dispose of objects when no longer needed
- The server maintains a list of currently instantiated COM servers

## Installation

### Using uv (recommended)

```bash
uv pip install .
```

### Using pip

```bash
pip install .
```

## Usage

Start the mcp-com-server server:

```bash
python server.py
```

Or configure it with your MCP client application.

## Requirements

- Python 3.10 or higher
- Windows operating system
- pywin32
- fastmcp

## Sample json file

This is how your `mcp.json` or `claude_desktop_config.js` should look like:

{
  "mcpServers": {
    "mcp-com-server": {
      "command": "python",
      "args": [
        "E:\\projects\\mcp-com-server\\server.py"
      ]
    }
  }
}

## License

MIT