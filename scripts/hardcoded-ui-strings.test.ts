import { expect, test } from 'bun:test';
import { resolve } from 'node:path';
import { allowedStrings, allows, scanDirectories, scanSource } from './hardcoded-ui-strings';

const root = resolve(import.meta.dir, '..');

const EDITOR_SOURCES = [
  'packages/docx-react/src',
  'packages/xlsx-react/src',
  'packages/pptx-react/src',
  'packages/vsdx-react/src',
  'shared/react-toolbar',
  'shared/format-glyph.tsx',
];

const allowlist = allowedStrings(root);
let scanned: ReturnType<typeof scanDirectories> | undefined;
const scan = () => (scanned ??= scanDirectories(root, EDITOR_SOURCES));

test('React editors route every user-visible string through i18n', () => {
  const found = scan()
    .filter((item) => !allowlist.some((entry) => allows(entry, item)))
    .map((item) => `${item.file}:${item.line} ${item.kind} ${item.name} "${item.text}"`);
  expect(found).toEqual([]);
}, 60_000);

test('every allowlisted string has a reason and still occurs', () => {
  const stale = allowlist.filter((entry) => entry.reason.trim() === '' || !scan().some((item) => allows(entry, item)));
  expect(stale).toEqual([]);
}, 60_000);

test('flags hardcoded JSX text, accessible names and labels', () => {
  const found = scanSource(
    'probe.tsx',
    `export function Probe({ t }: { t: (key: string) => string }) {
  const label = 'Insert row above';
  return (
    <div>
      <button aria-label="Close dialog" title={t('common.close')}>Save changes</button>
      <Tooltip content="Add comment" />
      {count > 1 ? 'Many items' : t('one')}
      {label}
    </div>
  );
}
const menu = { pasteAsPlainText: 'Paste as plain text', mergeCells: 'Merge cells' };
`,
  ).map((item) => item.text);
  expect(found).toEqual([
    'Insert row above',
    'Close dialog',
    'Save changes',
    'Add comment',
    'Many items',
    'Paste as plain text',
    'Merge cells',
  ]);
});

test('ignores diagnostics, styles, identifiers and translated text', () => {
  const found = scanSource(
    'probe.tsx',
    `import type { Mode } from 'Some Module';
export function Probe({ t, mode }: { t: (key: string) => string; mode: Mode }) {
  if (mode === 'Read only') throw new Error('The editor is read-only');
  console.warn('Layout is not ready');
  const failure = { code: 'stale-version', message: 'The document changed after that version' };
  return (
    <div className="toolbar-row" data-testid="Toolbar Row" style={{ fontFamily: 'Times New Roman' }}>
      <style>{'.a { color: red; }'}</style>
      <code>Ctrl+S</code>
      {t('toolbar.save')}
      &times;
      {failure.code}
    </div>
  );
}
`,
  );
  expect(found).toEqual([]);
});
