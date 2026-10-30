import { engineBackend } from '@betteroffice/xlsx';

export function writeStartupDiagnostics(backend = engineBackend()) {
  process.stderr.write(`betteroffice-agents: xlsx engine ${backend}\n`);
}
