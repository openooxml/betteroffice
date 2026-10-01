import type { DocxPackage } from '../types/document';
import { RELATIONSHIP_TYPES } from './relsParser';
import { headerFooterFilename } from './rezip/parts';

interface HeaderFooterAlias {
  kind: 'headers' | 'footers';
  rId: string;
}

function resolvedPartPath(target: string): string | undefined {
  if (!target || target.includes('\0') || target[1] === ':') return undefined;
  const path: string[] = [];
  for (const segment of headerFooterFilename(target.replaceAll('\\', '/')).split('/')) {
    if (!segment || segment === '.') continue;
    if (['..', '%2e%2e', '%2e.', '.%2e'].includes(segment.toLowerCase())) {
      if (path.pop() === undefined) return undefined;
    } else {
      path.push(segment);
    }
  }
  return path.join('/');
}

export function headerFooterAliasGroups(pkg: DocxPackage): HeaderFooterAlias[][] {
  const groups = new Map<string, HeaderFooterAlias[]>();
  for (const [kind, type] of [
    ['headers', RELATIONSHIP_TYPES.header],
    ['footers', RELATIONSHIP_TYPES.footer],
  ] as const) {
    for (const rId of pkg[kind]?.keys() ?? []) {
      const relationship = pkg.relationships?.get(rId);
      if (!relationship || relationship.type !== type || relationship.targetMode === 'External') {
        continue;
      }
      const path = resolvedPartPath(relationship.target);
      if (!path) continue;
      const group = groups.get(path) ?? [];
      group.push({ kind, rId });
      groups.set(path, group);
    }
  }
  return [...groups.values()].filter((group) => group.length >= 2);
}
