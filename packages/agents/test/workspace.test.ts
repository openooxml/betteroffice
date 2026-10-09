import { expect, test } from 'bun:test';
import { appendFile, mkdir, mkdtemp, readdir, rename, rm, symlink, truncate, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { FileWorkspace } from '../src/workspace';
import { fixture } from './fixture';

test('refuses oversized files and files that grow after the size check', async () => {
  const root = await mkdtemp(join(tmpdir(), 'betteroffice-open-'));
  const workspace = await FileWorkspace.create(root);
  try {
    await writeFile(join(root, 'large.docx'), new Uint8Array());
    await truncate(join(root, 'large.docx'), 64 * 1024 * 1024 + 1);
    await expect(workspace.open('large.docx')).rejects.toMatchObject({ code: 'FILE_TOO_LARGE' });
    await writeFile(join(root, 'source.docx'), await fixture());
    const internal = workspace as unknown as { validateHandle: (file: string, handle: unknown) => Promise<unknown> };
    const original = internal.validateHandle.bind(workspace);
    internal.validateHandle = async (file, handle) => {
      const info = await original(file, handle);
      await appendFile(file, 'unexpected growth');
      return info;
    };
    await expect(workspace.open('source.docx')).rejects.toMatchObject({ code: 'FILE_CHANGED' });
    expect((await workspace.files()).open).toEqual([]);
  } finally {
    workspace.dispose();
    await rm(root, { recursive: true, force: true });
  }
});

test('refuses an export directory replaced with an outside symlink during generation', async () => {
  const root = await mkdtemp(join(tmpdir(), 'betteroffice-export-'));
  const outside = await mkdtemp(join(tmpdir(), 'betteroffice-outside-'));
  const workspace = await FileWorkspace.create(root);
  try {
    await mkdir(join(root, 'output'));
    await writeFile(join(root, 'source.docx'), await fixture());
    const opened = await workspace.open('source.docx');
    const document = workspace.get(opened.document);
    const original = document.export.bind(document);
    document.export = async () => {
      await rename(join(root, 'output'), join(root, 'moved'));
      await symlink(outside, join(root, 'output'));
      return original();
    };
    await expect(workspace.export(opened.document, 'output/result.docx')).rejects.toThrow('inside the configured workspace');
    expect(await readdir(outside)).toEqual([]);
  } finally {
    workspace.dispose();
    await rm(root, { recursive: true, force: true });
    await rm(outside, { recursive: true, force: true });
  }
});

test('reports incomplete exports without deleting a destination another process may have replaced', async () => {
  const root = await mkdtemp(join(tmpdir(), 'betteroffice-export-'));
  const workspace = await FileWorkspace.create(root);
  try {
    await writeFile(join(root, 'source.docx'), await fixture());
    const opened = await workspace.open('source.docx');
    const internal = workspace as unknown as { validateHandle: (...args: unknown[]) => Promise<unknown> };
    const original = internal.validateHandle;
    internal.validateHandle = async () => { throw new Error('Injected validation failure'); };
    await expect(workspace.export(opened.document, 'result.docx')).rejects.toMatchObject({
      code: 'EXPORT_INCOMPLETE', details: { path: 'result.docx', cause: 'Injected validation failure' },
    });
    expect((await readdir(root)).sort()).toEqual(['result.docx', 'source.docx']);
    internal.validateHandle = original;
    expect((await workspace.export(opened.document, 'retry.docx')).bytes).toBeGreaterThan(0);
  } finally {
    workspace.dispose();
    await rm(root, { recursive: true, force: true });
  }
});
