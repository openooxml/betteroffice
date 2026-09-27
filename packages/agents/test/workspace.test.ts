import { expect, test } from 'bun:test';
import { mkdir, mkdtemp, readdir, rename, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { FileWorkspace } from '../src/workspace';
import { fixture } from './fixture';

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
