import { describe, expect, test } from 'bun:test';
import { CellClipboard, clipboardMarker, newClipboardMarker, parseCellClipboardText, readCellClipboard } from './clipboard';

const owner = {};
const range = { top: 0, left: 2, bottom: 0, right: 2 };
const cells = [[{ input: '=A1+$B$1', isFormula: true }]];
const text = cells[0][0].input;

function captured(cut = false) {
  const clipboard = new CellClipboard();
  const token = clipboard.begin();
  clipboard.capture(token, { owner, sheet: 0, range, cells, cut }, true);
  const html = `<table data-betteroffice-clipboard="${token.marker}"></table>`;
  const paste = (target = range, sheet = 0) => clipboard.paste(owner, text, target, html, sheet);
  return { clipboard, token, html, paste };
}

async function withClipboard(clipboard: Partial<Clipboard>, operation: () => Promise<void>) {
  const original = Object.getOwnPropertyDescriptor(globalThis, 'navigator');
  Object.defineProperty(globalThis, 'navigator', { configurable: true, value: { clipboard } });
  try { await operation(); }
  finally {
    if (original) Object.defineProperty(globalThis, 'navigator', original);
    else Reflect.deleteProperty(globalThis, 'navigator');
  }
}

describe('cell clipboard', () => {
  test('copy retains source coordinates and tiles per destination cell', () => {
    const { paste } = captured();
    expect(paste({ top: 2, left: 2, bottom: 3, right: 3 })?.edits).toEqual([
      { row: 2, col: 2, input: text, copyFrom: { row: 0, col: 2 } },
      { row: 2, col: 3, input: text, copyFrom: { row: 0, col: 2 } },
      { row: 3, col: 2, input: text, copyFrom: { row: 0, col: 2 } },
      { row: 3, col: 3, input: text, copyFrom: { row: 0, col: 2 } },
    ]);
  });

  test('refused and no-op moves retain the cut until a committed move', () => {
    const { clipboard, paste } = captured(true);
    const move = paste()!;
    clipboard.consume(move, false);
    expect(paste()?.cut).toBe(true);
    clipboard.consume(move, false);
    expect(paste()?.cut).toBe(true);
    clipboard.consume(move, true);
    expect(paste()).toBeNull();
  });

  test('cancel discards pending cuts and prevents a late write from arming one', () => {
    const { clipboard, paste } = captured(true);
    clipboard.cancelCut();
    expect(paste()).toBeNull();
    const token = clipboard.begin();
    clipboard.cancelCut();
    clipboard.capture(token, { owner, sheet: 0, range, cells, cut: true }, true);
    expect(clipboard.paste(owner, text, range, token.marker)).toBeNull();
  });

  test('a marker minted for a copy that never begins keeps the earlier capture, and a later begin writes under it', () => {
    const { clipboard, paste } = captured();
    const marker = newClipboardMarker();
    expect(paste()?.edits[0].copyFrom).toEqual({ row: 0, col: 2 });
    const token = clipboard.begin(undefined, marker);
    expect(token.marker).toBe(marker);
    clipboard.capture(token, { owner, sheet: 0, range, cells, cut: false }, true);
    const html = `<table data-betteroffice-clipboard="${marker}"></table>`;
    expect(clipboard.paste(owner, text, range, html, 0)?.edits[0].copyFrom).toEqual({ row: 0, col: 2 });
  });

  test('external text and replaced workbooks do not use internal copy metadata', () => {
    const { clipboard, html } = captured();
    expect(clipboard.paste(owner, '=B1', range, html)).toBeNull();
    expect(clipboard.paste({}, text, range, html)).toBeNull();
  });

  test('an edit cancels a captured cut and a cut whose clipboard write is pending', () => {
    const { clipboard, paste } = captured(true);
    clipboard.changed();
    expect(paste()).toBeNull();
    const token = clipboard.begin();
    clipboard.watch(token);
    clipboard.changed();
    clipboard.capture(token, { owner, sheet: 0, range, cells, cut: true }, true);
    expect(clipboard.paste(owner, text, range, token.marker)).toBeNull();
  });

  test('escape after a cut is accepted keeps it from arming when its write starts later', () => {
    const clipboard = new CellClipboard();
    const accepted = clipboard.escaped;
    clipboard.changed();
    clipboard.escape();
    const token = clipboard.begin(accepted);
    clipboard.capture(token, { owner, sheet: 0, range, cells, cut: true }, true);
    expect(clipboard.paste(owner, text, range, token.marker)).toBeNull();
    const next = clipboard.begin(clipboard.escaped);
    clipboard.capture(next, { owner, sheet: 0, range, cells, cut: true }, true);
    expect(clipboard.paste(owner, text, range, next.marker)?.cut).toBe(true);
  });

  test('earlier edits settling before the source read do not cancel a new cut', () => {
    const clipboard = new CellClipboard();
    const token = clipboard.begin();
    clipboard.changed();
    clipboard.watch(token);
    clipboard.capture(token, { owner, sheet: 0, range, cells, cut: true }, true);
    expect(clipboard.paste(owner, text, range, token.marker)?.cut).toBe(true);
  });

  for (const cut of [false, true]) {
    for (const html of [undefined, null, '<table>foreign</table>']) {
      test(`matching external text cannot reuse ${cut ? 'cut' : 'copy'} metadata with ${html} HTML`, () => {
        const { clipboard } = captured(cut);
        expect(clipboard.paste(owner, text, range, html)).toBeNull();
      });
    }

    test(`unmarked ${cut ? 'cut' : 'copy'} snapshots cannot be armed`, () => {
      const clipboard = new CellClipboard();
      const token = clipboard.begin();
      clipboard.capture(token, { owner, sheet: 0, range, cells, cut });
      expect(clipboard.paste(owner, text, range, token.marker)).toBeNull();
    });
  }

  test('an overwritten clipboard disarms the cut even if the old HTML is restored', () => {
    const { clipboard, paste } = captured(true);
    expect(clipboard.paste(owner, text, range)).toBeNull();
    expect(paste()).toBeNull();
  });

  test('cancel keeps copied snapshots and a newer copy survives an older cut completion', () => {
    const { clipboard, paste } = captured();
    clipboard.cancelCut();
    expect(paste()?.cut).toBe(false);
    const cut = captured(true);
    const move = cut.paste()!;
    const token = cut.clipboard.begin();
    cut.clipboard.capture(token, { owner, sheet: 0, range, cells, cut: false }, true);
    cut.clipboard.consume(move, true);
    expect(cut.clipboard.paste(owner, text, range, token.marker)?.cut).toBe(false);
  });

  test('a refused cross-sheet cut is disarmed without changing its source snapshot', () => {
    const { paste } = captured(true);
    expect(() => paste(range, 1)).toThrow('Moving cells between worksheets is unavailable');
    expect(paste()).toBeNull();
    expect(cells).toEqual([[{ input: text, isFormula: true }]]);
    expect(captured().paste(range, 1)?.cut).toBe(false);
  });

  for (const input of ['=SUM(A1:A2)', '=SEQUENCE(2)', '']) {
    test(`refuses copying array and spill cells (${input || 'spill member'})`, () => {
      const clipboard = new CellClipboard();
      const token = clipboard.begin();
      clipboard.capture(token, { owner, sheet: 0, range, cut: false,
        cells: [[{ input, isFormula: input !== '', isArrayFormula: true }]] }, true);
      expect(() => clipboard.paste(owner, input, range, token.marker)).toThrow('Copying array or spill formula cells is unavailable');
    });
  }

  test('bounds tiled edits before materialising a huge target', () => {
    const { paste } = captured();
    expect(() => paste({ top: 0, left: 0, bottom: 1_048_575, right: 16_383 })).toThrow('100000-cell limit');
    expect(paste({ top: 0, left: 0, bottom: 99_999, right: 0 })?.edits).toHaveLength(100_000);
    expect(() => paste({ top: 0, left: 0, bottom: 100_000, right: 0 })).toThrow('100000-cell limit');
  });

  test('bounds source dimensions even for a single-cell destination', () => {
    const clipboard = new CellClipboard();
    const token = clipboard.begin();
    clipboard.capture(token, { owner, sheet: 0, cut: false, cells,
      range: { top: 0, left: 0, bottom: 100_000, right: 0 } }, true);
    expect(() => clipboard.paste(owner, text, range, token.marker)).toThrow('100000-cell limit');
  });

  test('bounds external text edits and clipboard bytes', () => {
    expect(() => parseCellClipboardText('x\t'.repeat(100_000))).toThrow('100000-cell limit');
    expect(() => parseCellClipboardText('x'.repeat(16 * 1024 * 1024 + 1))).toThrow('16 MiB limit');
    expect(() => parseCellClipboardText('€'.repeat(6 * 1024 * 1024))).toThrow('16 MiB limit');
    expect(parseCellClipboardText('"a\tb"\tc')).toEqual([['a\tb', 'c']]);
    const quoted = '\t'.repeat(100_001);
    expect(parseCellClipboardText(`"${quoted}"`)).toEqual([[quoted]]);
  });

  test('text-only clipboard reads keep formula text as ordinary input', async () => {
    await withClipboard({ readText: async () => text }, async () => {
      const data = await readCellClipboard();
      const { clipboard } = captured(true);
      expect(clipboard.paste(owner, data.text, range, data.html)).toBeNull();
      expect(parseCellClipboardText(data.text)).toEqual([[text]]);
    });
  });

  test('refuses oversized HTML before decoding its blob', async () => {
    let decoded = false;
    const blob = { size: 16 * 1024 * 1024 + 1, text: async () => { decoded = true; return ''; } } as Blob;
    const item = { types: ['text/plain', 'text/html'], getType: async () => blob } as unknown as ClipboardItem;
    await withClipboard({ readText: async () => text, read: async () => [item] }, async () => {
      await expect(readCellClipboard()).rejects.toThrow('16 MiB limit');
      expect(decoded).toBe(false);
    });
  });

  test('refuses oversized plain text even when HTML reading is unavailable', async () => {
    await withClipboard({ readText: async () => 'x'.repeat(16 * 1024 * 1024 + 1) }, async () => {
      await expect(readCellClipboard()).rejects.toThrow('16 MiB limit');
    });
  });

  test('clipboard marker escapes cell content', async () => {
    const { token } = captured();
    const html = await clipboardMarker(token, '<script>&').text();
    expect(html).toContain('&lt;script&gt;&amp;');
    expect(html).not.toContain('<script>');
  });
});
