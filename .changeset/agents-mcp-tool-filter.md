---
'@betteroffice/agents': patch
---

Adds `--formats` and `--tools` filters (with `BETTEROFFICE_MCP_FORMATS` and `BETTEROFFICE_MCP_TOOLS` env equivalents) to `betteroffice-mcp` and `createOfficeMcpServer`, so `tools/list` only exposes the tools a job needs.
