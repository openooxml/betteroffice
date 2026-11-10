import { expect, test } from 'bun:test';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { writeFontCdnTree } from '../../../scripts/fonts/cdn-tree';

const files = ['Synthetic-Regular.ttf.zst', 'Synthetic-Bold.ttf.zst', 'Synthetic/slices.json',
  'Synthetic/skeleton.zst', 'Synthetic/core.zst', 'Synthetic/c001.zst'];

async function fixture(root: string): Promise<void> {
  for (const pack of ['latin', 'intl', 'cjk']) {
    const source = join(root, `fonts-${pack}`);
    for (const file of files) {
      const path = join(source, 'assets', file);
      await mkdir(dirname(path), { recursive: true });
      await writeFile(path, file.endsWith('.json') ? JSON.stringify({
        skeleton: { file: files[3] }, core: { file: files[4] }, chunks: [{ file: files[5] }],
      }) : new Uint8Array([1, 2, 3]));
    }
    await writeFile(join(source, 'assets/manifest.json'), JSON.stringify({
      pack, version: '1.2.3', faces: [
        { file: files[0], dict: null, slices: { index: files[2] } },
        { file: files[1], dict: files[0] },
      ],
    }));
    await mkdir(join(source, 'LICENSES'));
    await writeFile(join(source, 'LICENSES/Synthetic.txt'), 'Synthetic license');
  }
}

test('the self-hosted tree copies every manifest asset and pack license', async () => {
  const root = await mkdtemp(join(tmpdir(), 'font-cdn-synthetic-'));
  try {
    const packages = join(root, 'packages');
    await fixture(packages);
    await writeFontCdnTree(join(root, 'output'), packages);
    for (const pack of ['latin', 'intl', 'cjk']) {
      for (const file of [...files, 'manifest.json']) {
        expect(await readFile(join(root, 'output', pack, '1.2.3', file)))
          .toEqual(await readFile(join(packages, `fonts-${pack}`, 'assets', file)));
      }
      expect(await readFile(join(root, 'output', pack, '1.2.3/LICENSES/Synthetic.txt'), 'utf8')).toBe('Synthetic license');
    }
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test('a missing face, dictionary, slice or license directory fails the export', async () => {
  for (const file of [...files.map((file) => `assets/${file}`), 'assets/manifest.json', 'LICENSES']) {
    const root = await mkdtemp(join(tmpdir(), 'font-cdn-missing-'));
    try {
      const packages = join(root, 'packages');
      await fixture(packages);
      await rm(join(packages, 'fonts-latin', file), { recursive: true });
      await expect(writeFontCdnTree(join(root, 'output'), packages)).rejects.toThrow();
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  }
});
