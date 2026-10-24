import { useEffect } from 'react';
import type { DocxInput } from '@betteroffice/docx/utils';

interface FileAliasProps {
  documentBuffer?: DocxInput | null;
  documentName?: string;
  file?: DocxInput | null;
  fileName?: string;
}

const ALIASES = [
  ['documentBuffer', 'file'],
  ['documentName', 'fileName'],
] as const;

const warned = new Set<string>();

function isDev(): boolean {
  return typeof process !== 'undefined' && process.env?.NODE_ENV !== 'production';
}

/** Moves `file` and `fileName` onto `documentBuffer` and `documentName`; a set older name wins. */
export function resolveDocxFileAliases<P extends FileAliasProps>(
  props: P
): { props: P; conflicts: string[] } {
  if (props.file === undefined && props.fileName === undefined) return { props, conflicts: [] };
  const { file, fileName, ...rest } = props;
  const resolved = rest as P;
  const aliases = { file, fileName };
  const conflicts: string[] = [];
  for (const [name, alias] of ALIASES) {
    if (aliases[alias] === undefined) continue;
    if (resolved[name] === undefined) Object.assign(resolved, { [name]: aliases[alias] });
    else conflicts.push(`[DocxEditor] Both ${name} and ${alias} are set; using ${name}. Pass only one.`);
  }
  return { props: resolved, conflicts };
}

export function warnDocxFileAliasConflicts(conflicts: readonly string[]): void {
  if (!isDev()) return;
  for (const warning of conflicts) {
    if (warned.has(warning)) continue;
    warned.add(warning);
    console.warn(warning);
  }
}

export function useDocxFileAliases<P extends FileAliasProps>(props: P): P {
  const { props: resolved, conflicts } = resolveDocxFileAliases(props);
  const key = conflicts.join('\n');
  useEffect(() => {
    if (key) warnDocxFileAliasConflicts(key.split('\n'));
  }, [key]);
  return resolved;
}

export function resetDocxFileAliasWarningsForTests(): void {
  warned.clear();
}
