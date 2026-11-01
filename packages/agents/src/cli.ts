#!/usr/bin/env node
import { parseArgs } from 'node:util';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createOfficeMcpServer } from './mcp';
import { renderDocxPage } from './render';
import { resolveFormats, resolveTools } from './tool-filter';

try {
  const { values } = parseArgs({ options: { root: { type: 'string' }, 'read-only': { type: 'boolean' }, formats: { type: 'string' }, tools: { type: 'string' }, help: { type: 'boolean', short: 'h' } } });
  if (values.help) {
    process.stdout.write('Usage: betteroffice-mcp --root <directory> [--read-only] [--formats <list>] [--tools <list>]\n\nA local stdio MCP server for DOCX, XLSX, and PPTX search, proposals, and export. Page previews are DOCX-only.\n\nOptions:\n  --root <directory>  Workspace root (required).\n  --read-only         Expose only inspection tools.\n  --formats <list>    Comma-separated formats to serve: docx, xlsx, pptx (env BETTEROFFICE_MCP_FORMATS).\n  --tools <list>      Comma-separated tool names to expose (env BETTEROFFICE_MCP_TOOLS).\n');
  } else {
    if (!values.root) throw new Error('--root <directory> is required.');
    const formats = resolveFormats(values.formats, process.env.BETTEROFFICE_MCP_FORMATS);
    const tools = resolveTools(values.tools, process.env.BETTEROFFICE_MCP_TOOLS);
    const server = await createOfficeMcpServer({ root: values.root, readOnly: values['read-only'], renderer: renderDocxPage, formats, tools });
    await server.connect(new StdioServerTransport(process.stdin, process.stdout, { maxBufferSize: 1024 * 1024 }));
  }
} catch (error) {
  process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
  process.exitCode = 1;
}
