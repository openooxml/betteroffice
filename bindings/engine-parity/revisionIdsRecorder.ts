import { mock } from 'bun:test';
import { AsyncLocalStorage } from 'node:async_hooks';
import { randomUUID } from 'node:crypto';
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

const out = process.env.DOCX_REVISION_IDS_DIR;
const recording = new AsyncLocalStorage<{ name: string; count: number }>();

export async function recordRevisionIdCases(name: string, save: () => Promise<void>): Promise<number> {
  const current = { name: name.replace(/[^a-zA-Z0-9_-]/g, '-').slice(0, 160), count: 0 };
  await recording.run(current, save);
  return current.count;
}

if (out) {
  const { isSessionRevisionInfo, markRevisionInfo, withSessionRevisionIds } =
    await import('../../packages/docx/src/docx/sessionRevisionIds');

  function visit(value: unknown, session: Set<number>): void {
    if (value === null || typeof value !== 'object' || ArrayBuffer.isView(value)) return;
    if (isSessionRevisionInfo(value)) {
      const { id } = value as { id: number };
      session.add(id);
      return;
    }
    for (const child of Object.values(value)) visit(child, session);
  }

  mock.module('../../packages/docx/src/docx/sessionRevisionIds', () => ({
    isSessionRevisionInfo,
    markRevisionInfo,
    withSessionRevisionIds<T>(tree: object, write: () => T): T {
      if (!('determinism' in tree)) return withSessionRevisionIds(tree, write);
      const session = new Set<number>();
      visit(tree, session);
      if (!session.size) return withSessionRevisionIds(tree, write);

      const request = JSON.stringify(tree);
      const wire = withSessionRevisionIds(tree, write);
      if (typeof wire !== 'string') throw new Error('S13 revision ID recording requires a JSON wire');
      const current = recording.getStore();
      mkdirSync(out, { recursive: true });
      writeFileSync(
        join(out, `${current?.name ?? 'save'}-${randomUUID()}.json`),
        JSON.stringify({ request, session: [...session], wire })
      );
      if (current) current.count += 1;
      return wire;
    },
  }));
}
