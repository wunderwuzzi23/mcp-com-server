# mcp-com-server

An experimental MCP Server that hosts and manages COM servers on Windows. I built this to learn about MCP - so, it's not a serious project at this time.

**WARNING**
This gives applications that can host MCP servers the power to instantiate and interact with any Windows COM object, like Excel, Word, Outlook, Shell, SAPI, WMI, and many many others.
Created by Johann Rehberger (@wunderwuzzi23)

## Experimental 

This is a quick first version that serves as a proof-of-concept, but it works pretty well.

## Security Caution!

There is only one very basic security feature, which is that you can ALLOWLIST CLSID and/or ProgIDs (by default all is allowed, so it's yolo mode)
**This is all very dangerous obviously - so use with caution!**

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

```
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
```

## License

MIT
