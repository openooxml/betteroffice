import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { clipboardFragmentPlainText, parseClipboardHtml, plainTextClipboardFragment } from './clipboardHtml';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function fixture(name: string): string {
  return readFileSync(resolve(import.meta.dir, 'fixtures/clipboard', name), 'utf8');
}

test('word fixture parses formatted runs', () => {
  expect(parseClipboardHtml(fixture('word.html'))).toEqual({
    paragraphs: [
      { text: 'Quarterly report', runs: [
        { text: 'Quarterly', format: { bold: true, color: { rgb: 'C00000' }, fontSize: 14, fontFamily: { ascii: 'Calibri' } } },
        { text: ' report', format: { fontSize: 14, fontFamily: { ascii: 'Calibri' } } },
      ] },
      { text: 'Second line', runs: [{ text: 'Second line', format: { italic: true } }] },
    ],
  });
});

test('google docs fixture cancels wrapper bold and keeps span formatting', () => {
  expect(parseClipboardHtml(fixture('google-docs.html'))).toEqual({
    paragraphs: [
      { text: 'Rich heading', runs: [{ text: 'Rich heading', format: { bold: true, color: { rgb: 'FF0000' }, fontSize: 20, fontFamily: { ascii: 'Georgia' } } }] },
      { text: 'Body note', runs: [{ text: 'Body note', format: { italic: true, fontSize: 11, fontFamily: { ascii: 'Arial' } } }] },
    ],
  });
});

test('chrome fixture parses heading link and list', () => {
  expect(parseClipboardHtml(fixture('chrome.html'))).toEqual({
    paragraphs: [
      { text: 'Release notes', pStyle: 'Heading2' },
      { text: 'Read the guide and save your work.', runs: [
        { text: 'Read the ' },
        { text: 'guide', href: 'https://example.com/guide' },
        { text: ' and ' },
        { text: 'save', format: { bold: true } },
        { text: ' your work.' },
      ] },
      { text: 'First item', list: 'bullet' },
      { text: 'Second item', list: 'bullet' },
    ],
  });
});

test('safari fixture parses converted space and inline styles', () => {
  expect(parseClipboardHtml(fixture('safari.html'))).toEqual({
    paragraphs: [
      { text: 'Linked text bold tail', runs: [
        { text: 'Linked text', format: { underline: true, color: { rgb: '0000FF' }, fontSize: 12, fontFamily: { ascii: 'Helvetica' } } },
        { text: ' ' },
        { text: 'bold tail', format: { bold: true } },
      ] },
      { text: 'Plain second' },
    ],
  });
});

test('firefox fixture parses heading highlight strike and numbered list', () => {
  expect(parseClipboardHtml(fixture('firefox.html'))).toEqual({
    paragraphs: [
      { text: 'Firefox heading', pStyle: 'Heading3' },
      { text: 'Highlighted and removed text', runs: [
        { text: 'Highlighted', format: { highlight: 'FFFF00' } },
        { text: ' and ' },
        { text: 'removed', format: { strike: true } },
        { text: ' text' },
      ] },
      { text: 'Numbered one', list: 'numbered' },
    ],
  });
});

test('outlook fixture parses word-like formatted runs', () => {
  expect(parseClipboardHtml(fixture('outlook.html'))).toEqual({
    paragraphs: [
      { text: 'Outlook title', runs: [{ text: 'Outlook title', format: { bold: true, color: { rgb: '0070C0' }, fontSize: 12, fontFamily: { ascii: 'Aptos' } } }] },
      { text: 'Underlined note', runs: [{ text: 'Underlined note', format: { underline: true } }] },
    ],
  });
});

test('dangerous link schemes drop the link but keep the text', () => {
  for (const href of ['javascript:alert(1)', 'JaVaScRiPt:alert(1)', ' data:text/html;base64,SGk=', 'vbscript:msgbox', 'file:///etc/passwd', '/relative/path', 'java\tscript:alert(1)']) {
    const parsed = parseClipboardHtml(`<body><p><b>bold</b> <a href="${href}">click</a></p></body>`);
    expect(parsed).not.toBeNull();
    expect(parsed!.paragraphs[0]!.text).toBe('bold click');
    expect(parsed!.paragraphs[0]!.runs).toEqual([{ text: 'bold', format: { bold: true } }, { text: ' click' }]);
  }
});

test('allowed link schemes are kept case-insensitively', () => {
  const parsed = parseClipboardHtml(`<body><p><a href="HTTPS://example.com/x">site</a> <a href="MailTo:a@example.com">mail</a></p></body>`);
  expect(parsed!.paragraphs[0]!.runs).toEqual([
    { text: 'site', href: 'HTTPS://example.com/x' },
    { text: ' ' },
    { text: 'mail', href: 'MailTo:a@example.com' },
  ]);
});

test('script style template and iframe content contributes no text', () => {
  const parsed = parseClipboardHtml(`<body><p><b>keep</b></p><script>evil()</script><style>.x{color:red}</style><template><p>templ</p></template><iframe srcdoc="frame"></iframe><svg><text>svgtext</text></svg></body>`);
  expect(parsed).toEqual({ paragraphs: [{ text: 'keep', runs: [{ text: 'keep', format: { bold: true } }] }] });
});

test('oversized html returns null', () => {
  expect(parseClipboardHtml(`<body><p><b>${'x'.repeat(2_000_001)}</b></p></body>`)).toBeNull();
});

test('nesting deeper than 64 returns null', () => {
  const html = `<body>${'<div>'.repeat(70)}<p><b>deep</b></p>${'</div>'.repeat(70)}</body>`;
  expect(parseClipboardHtml(html)).toBeNull();
});

test('more than 50000 elements returns null', () => {
  const html = `<body><p><b>x</b></p>${'<span>y</span>'.repeat(50_001)}</body>`;
  expect(parseClipboardHtml(html)).toBeNull();
});

test('style attribute longer than 4096 returns null', () => {
  const html = `<body><p><span style="color:#ff0000;${'font-weight:bold;'.repeat(300)}">x</span></p></body>`;
  expect(parseClipboardHtml(html)).toBeNull();
});

test('css validators reject url expression and oversized values', () => {
  const parsed = parseClipboardHtml(`<body><p><span style="font-family:url(evil);font-size:9999pt;color:expression(alert(1));background:url(http://x)">plain</span> <span style="font-size:12pt">sized</span></p></body>`);
  expect(parsed).not.toBeNull();
  expect(parsed!.paragraphs[0]!.runs).toEqual([{ text: 'plain ' }, { text: 'sized', format: { fontSize: 12 } }]);
});

test('css validators reject escaped font family and bad colours', () => {
  const parsed = parseClipboardHtml(`<body><p><span style="font-family:Ar\\ial;color:#12;font-weight:bold">x</span></p></body>`);
  expect(parsed!.paragraphs[0]!.runs).toEqual([{ text: 'x', format: { bold: true } }]);
});

test('content with no allowlisted formatting returns null like main', () => {
  expect(parseClipboardHtml(`<body><p>Plain paragraph</p></body>`)).toBeNull();
  expect(parseClipboardHtml(`<body><div>Just text</div></body>`)).toBeNull();
  expect(parseClipboardHtml(`<body><p><span style="margin:0;position:absolute">Plain</span></p></body>`)).toBeNull();
  expect(parseClipboardHtml(`<body><p><font color="red">Plain</font></p></body>`)).toBeNull();
});

test('keep text only fragment equals the plain result for fixtures', () => {
  const cases: Array<[string, string]> = [
    ['word.html', 'Quarterly report\nSecond line'],
    ['chrome.html', 'Release notes\nRead the guide and save your work.\nFirst item\nSecond item'],
    ['firefox.html', 'Firefox heading\nHighlighted and removed text\nNumbered one'],
  ];
  for (const [name, plain] of cases) {
    const parsed = parseClipboardHtml(fixture(name));
    expect(parsed).not.toBeNull();
    expect(plainTextClipboardFragment(plain)).toEqual({
      paragraphs: plain.split('\n').map((text) => ({ text })),
    });
    expect(plainTextClipboardFragment(plain).paragraphs.map((p) => p.text)).toEqual(parsed!.paragraphs.map((p) => p.text));
  }
});

test('headings and list items inside an inline wrapper keep their paragraph kind', () => {
  const html = '<b id="docs-internal-guid-1" style="font-weight:normal"><h2>Heading</h2><ul><li>One</li></ul><ol><li>Two</li></ol></b>';
  expect(parseClipboardHtml(html)).toEqual({
    paragraphs: [
      { text: 'Heading', pStyle: 'Heading2' },
      { text: 'One', list: 'bullet' },
      { text: 'Two', list: 'numbered' },
    ],
  });
});

test('an editor payload carrying a script link is not trusted', () => {
  const payload = JSON.stringify({ v: 1, fragment: { paragraphs: [{ text: 'Click', runs: [{ text: 'Click', href: 'javascript:alert(1)' }] }] } });
  const html = `<div data-betteroffice-docx="${payload.replace(/"/g, '&quot;')}"><a href="javascript:alert(1)">Click</a></div>`;
  const parsed = parseClipboardHtml(html);
  expect(JSON.stringify(parsed ?? null)).not.toContain('javascript:');
});

test('keep text only splits on the same line separators as a plain paste', () => {
  expect(plainTextClipboardFragment('a\u000Bb\u2028c\u2029d\r\ne\rf')).toEqual({
    paragraphs: ['a', 'b', 'c', 'd', 'e', 'f'].map((text) => ({ text })),
  });
});

test('plain text of a table-only fragment keeps the cell text', () => {
  const cell = (text: string) => ({ tcPr: {}, content: [{ kind: 'text' as const, text }, { kind: 'pilcrow' as const }] });
  expect(clipboardFragmentPlainText({
    paragraphs: [{ text: '', tables: [{ tblPr: {}, grid: [], rows: [{ trPr: {}, cells: [cell('a'), cell('b')] }, { trPr: {}, cells: [cell('c'), cell('d')] }] }] }],
  })).toBe('a\tb\nc\td');
});
