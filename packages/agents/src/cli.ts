#!/usr/bin/env node
import { parseArgs } from 'node:util';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { createOfficeMcpServer } from './mcp';
import { renderDocxPage } from './render';

try {
  const { values } = parseArgs({ options: { root: { type: 'string' }, 'read-only': { type: 'boolean' }, help: { type: 'boolean', short: 'h' } } });
  if (values.help) {
    process.stdout.write('Usage: betteroffice-mcp --root <directory> [--read-only]\n\nA local stdio MCP server for DOCX search, proposals, page previews, and export.\n');
  } else {
    if (!values.root) throw new Error('--root <directory> is required.');
    const server = await createOfficeMcpServer({ root: values.root, readOnly: values['read-only'], renderer: renderDocxPage });
    await server.connect(new StdioServerTransport());
  }
} catch (error) {
  process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
  process.exitCode = 1;
}
