import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, expect, test } from 'bun:test';
import { validClipboardFragment } from '@betteroffice/docx/yrs';
import { mergeFormattingFragment } from '../pasteOptions/clipboardFragmentTransforms';
import { clipboardFragmentPlainText, pasteOptionsOffered, plainTextClipboardFragment } from '../pasteOptions/pasteText';
import {
  clipboardFragmentHtml, MAX_FORMATTED_HTML_LENGTH, MAX_FORMATTED_HTML_DEPTH, MAX_FORMATTED_HTML_RUNS, MAX_FORMATTED_HTML_TEXT, MAX_FORMATTED_HTML_PARAGRAPHS,
  parseClipboardHtml,
} from './clipboardHtml';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const formatted = { formatted: true };
const table = '<table><tr><td>cell</td></tr></table>';

function imported(html: string, lines: string[]) {
  const fragment = parseClipboardHtml(html, formatted);
  expect(fragment?.paragraphs.map(p => p.text)).toEqual(lines);
  expect(fragment!.paragraphs.map(p => (p.runs ?? []).map(run => run.text).join(''))).toEqual(lines);
  expect(validClipboardFragment(fragment)).toBe(true);
  expect(validClipboardFragment(mergeFormattingFragment(fragment!))).toBe(true);
  expect(validClipboardFragment(plainTextClipboardFragment(clipboardFragmentPlainText(fragment!)))).toBe(true);
  return fragment!;
}

test('paragraph rules and inline declarations produce raw run attributes and all paste options', () => {
  const fragment = imported('<style>p.MsoNormal{font-family:"Calibri",sans-serif;font-size:11.0pt}'
    + 'h1{font-size:16.0pt;color:#2F5496}</style><h1>Heading</h1>'
    + '<p class="MsoNormal"><b>Bold</b> <span style="color:rgb(255,0,0)">red</span></p>', ['Heading', 'Bold red']);
  const base = { fontFamily: { ascii: 'Calibri', hAnsi: 'Calibri' }, fontSize: { size: 22 } };
  expect(fragment.paragraphs[0]!.runs![0]!.format?.other).toEqual({ fontSize: { size: 32 }, textColor: { rgb: '2F5496' } });
  expect(fragment.paragraphs[1]!.runs!.map(run => run.format?.other)).toEqual([
    { ...base, bold: true }, base, { ...base, textColor: { rgb: 'FF0000' } },
  ]);
  expect(pasteOptionsOffered(fragment)).toEqual(['keep-source', 'merge', 'text-only']);
});

test('nested blocks and list items each become their own paragraph', () => {
  imported('<div class="WordSection1"><p class="MsoNormal"><b>a</b></p><p class="MsoNormal">b</p>'
    + '<div><h2>c</h2>tail</div></div><ul><li><b>one</b></li><li><p>two</p><p>three</p></li></ul>',
  ['a', 'b', 'c', 'tail', 'one', 'two', 'three']);
});

test('browser links retain their title, nested emphasis and first usable font', () => {
  const fragment = imported('<p><a href="https://example.com/x" title="tip"><em>link</em></a>'
    + '<span style="font-size:16px;font-family:-apple-system, Arial">font</span></p>', ['linkfont']);
  expect(fragment.paragraphs[0]!.runs).toEqual([
    { text: 'link', href: 'https://example.com/x', tooltip: 'tip', format: { other: { italic: true } } },
    { text: 'font', format: { other: { fontSize: { size: 24 }, fontFamily: { ascii: 'Arial', hAnsi: 'Arial' } } } },
  ]);
});

test('unsupported links keep their text and merge into adjacent plain runs', () => {
  const fragment = imported('<p><b>ok</b><a href="javascript:void(0)" title="tip">one</a>'
    + '<a href="data:text/plain,two">two</a></p>', ['okonetwo']);
  expect(fragment.paragraphs[0]!.runs![1]).toEqual({ text: 'onetwo', format: { other: {} } });
});

test('plain paragraphs defer to plain paste and formatted imports require the option', () => {
  expect(parseClipboardHtml('<p>one</p><p>two</p>', formatted)).toBeNull();
  expect(parseClipboardHtml('<p><b>x</b></p>')).toBeNull();
});

test('spaces collapse across run boundaries while tabs and line breaks survive', () => {
  const fragment = imported('<p>a <b> b </b> c<br> &nbsp;<i>d\t e</i> </p>', ['a b c', 'd\t e']);
  expect(fragment.paragraphs[0]!.runs!.map(run => run.text)).toEqual(['a ', 'b ', 'c']);
});

test('list marker glyphs use ordinary bullets and leave their symbol font behind', () => {
  for (const [font, glyph, bullet] of [['Symbol', '·', '•'], ['Wingdings', '§', '▪']] as const) {
    const fragment = imported(`<p style="mso-list:l0 level1 lfo1"><span style="font-family:${font};mso-list:Ignore">${glyph}`
      + '<span style=\'font:7.0pt "Times New Roman"\'>&nbsp;&nbsp; </span></span>Item</p>', [`${bullet} Item`]);
    expect(fragment.paragraphs[0]!.runs![0]!.format?.other).toEqual({});
    expect(fragment.paragraphs[0]!.runs![1]).toEqual({ text: ' ', format: { other: { fontSize: { size: 14 } } } });
  }
});

test('simple rules cascade, nested at-rules are ignored, and equivalent runs merge', () => {
  const fragment = imported('<style>@media screen{@supports (display:grid){b{font-size:99pt}}}'
    + 'b:hover{color:#f00} B.mark,.alt{color:#abc} b{color:#def}</style>'
    + '<p><b style="font-weight:normal"><b class="mark">a</b><strong class="alt">b</strong>c</b></p>', ['abc']);
  expect(fragment.paragraphs[0]!.runs).toEqual([
    { text: 'ab', format: { other: { bold: true, textColor: { rgb: 'AABBCC' } } } },
    { text: 'c', format: { other: { textColor: { rgb: 'DDEEFF' } } } },
  ]);
});

test('the style text budget spans style elements and inline styles take precedence', () => {
  const fragment = imported('<style>.mark{color:#012}</style><style>/*' + ' '.repeat(200_000)
    + '*/p{font-size:40px}</style><p class="mark" style="color:#345">one</p>', ['one']);
  expect(fragment.paragraphs[0]!.runs![0]!.format?.other).toEqual({ textColor: { rgb: '334455' } });
});

test('link and title boundaries survive merging and overlong values leave the text intact', () => {
  const fragment = imported('<p><a href="#a" title="one">a</a><a href="#a" title="two">b</a>'
    + '<a href="mailto:a@example.com" title="' + 'x'.repeat(1025) + '">c</a>'
    + '<a href="#' + 'x'.repeat(8192) + '">d</a></p>', ['abcd']);
  expect(fragment.paragraphs[0]!.runs!.map(run => [run.href, run.tooltip])).toEqual([
    ['#a', 'one'], ['#a', 'two'], ['mailto:a@example.com', undefined], [undefined, undefined],
  ]);
});

test('decorations, baseline shifts and highlights survive in text surrounding a table', () => {
  const fragment = imported('<blockquote><ins>a</ins><del>b</del><sup>c</sup><sub>d</sub></blockquote>' + table
    + '<pre><span style="vertical-align:super;background:#abcd00">e</span>'
    + '<span style="vertical-align:sub;mso-highlight:darkblue">f</span></pre>', ['abcd', 'ef']);
  expect(fragment.paragraphs[1]!.tables).toEqual(parseClipboardHtml(table)!.paragraphs[0]!.tables);
  expect(fragment.paragraphs.flatMap(p => p.runs!.map(run => run.format?.other))).toEqual([
    { underline: { style: 'single' } }, { strike: { double: false } }, { superscript: true }, { subscript: true },
    { superscript: true, highlight: 'ABCD00' }, { subscript: true, highlight: 'darkBlue' },
  ]);
});

test('invalid sizes, fonts, and transparent or default white backgrounds add no attributes', () => {
  const fragment = imported('<p><i>x</i><span style="font-size:-12pt;font-family:ui-serif;background:#fff">a</span>'
    + '<span style="font-size:1639pt;font-family:&quot;a/b&quot;;background-color:rgba(1,2,3,0)">b</span></p>', ['xab']);
  expect(fragment.paragraphs[0]!.runs![1]).toEqual({ text: 'ab', format: { other: {} } });
});

test('the run budget counts merged runs', () => {
  const count = MAX_FORMATTED_HTML_RUNS + 1;
  const fragment = imported('<p>' + '<b>x</b>'.repeat(count) + '</p>', ['x'.repeat(count)]);
  expect(fragment.paragraphs[0]!.runs).toHaveLength(1);
});

test('each formatted import limit falls back to the original table or plain paste', () => {
  expect([MAX_FORMATTED_HTML_LENGTH, MAX_FORMATTED_HTML_DEPTH, MAX_FORMATTED_HTML_RUNS]).toEqual([2_000_000, 200, 1_000]);
  for (const html of [
    '<b>x</b><!--' + ' '.repeat(MAX_FORMATTED_HTML_LENGTH) + '-->',
    '<span>'.repeat(300) + '<b>x</b>' + '</span>'.repeat(300),
    '<p>' + '<b>x</b>y'.repeat(MAX_FORMATTED_HTML_RUNS / 2) + '<b>x</b></p>',
  ]) {
    expect(parseClipboardHtml(html, formatted)).toBeNull();
    const legacy = parseClipboardHtml(html + table);
    expect(legacy?.paragraphs.at(-1)?.tables).toHaveLength(1);
    expect(parseClipboardHtml(html + table, formatted)).toEqual(legacy);
  }
});

test('an internal payload has priority even in three megabytes of HTML', () => {
  const fragment = { paragraphs: [{ text: 'payload' }] };
  expect(parseClipboardHtml(clipboardFragmentHtml(fragment) + '<b>' + 'x'.repeat(3_000_000) + '</b>', formatted)).toEqual(fragment);
});

test('blank lines from line breaks and empty paragraphs are kept, but not at the ends', () => {
  imported('<p><b>A</b><br><br>B</p>', ['A', '', 'B']);
  imported('<p class="MsoNormal"><o:p>&nbsp;</o:p></p><p class="MsoNormal"><b>A</b></p>'
    + '<p class="MsoNormal"><o:p>&nbsp;</o:p></p><p><br></p><p class="MsoNormal">B</p><p><br></p>', ['A', '', '', 'B']);
});

test('preformatted text keeps its spaces and line breaks', () => {
  imported('<pre><b>A\n  B</b>\n\nC</pre>', ['A', '  B', '', 'C']);
  imported('<p style="white-space:pre-wrap"><i>x  y</i></p>', ['x  y']);
});

test('plain text around a table with no formatting pastes as before', () => {
  const html = '<p>Before</p>' + table + '<p>After</p>';
  expect(parseClipboardHtml(html, formatted)).toEqual(parseClipboardHtml(html));
});

test('the text budget falls back to plain paste', () => {
  expect(parseClipboardHtml('<p><b>' + 'x'.repeat(MAX_FORMATTED_HTML_TEXT + 1) + '</b></p>', formatted)).toBeNull();
  imported('<p><b>' + 'x'.repeat(MAX_FORMATTED_HTML_TEXT) + '</b></p>', ['x'.repeat(MAX_FORMATTED_HTML_TEXT)]);
});

test('hostile styles and run counts stay linear', () => {
  const timed = (html: string) => {
    const start = performance.now();
    const fragment = parseClipboardHtml(html, formatted);
    return { fragment, ms: performance.now() - start };
  };
  const shorthand = timed('<p><b style="font:12px' + ' '.repeat(150_000) + 'Arial\nBlack">x</b></p>');
  expect(shorthand.fragment?.paragraphs.map(p => p.text)).toEqual(['x']);
  expect(shorthand.ms).toBeLessThan(2_000);
  const fonts = timed('<style>b{font-family:' + 'serif,'.repeat(30_000) + 'Arial}</style>' + '<b>x</b>'.repeat(150_000));
  expect(fonts.fragment?.paragraphs.map(p => p.text)).toEqual(['x'.repeat(150_000)]);
  expect(fonts.ms).toBeLessThan(10_000);
  const runs = timed('<p>' + '<b>x</b>y'.repeat(210_000) + '</p>');
  expect(runs.fragment).toBeNull();
  expect(runs.ms).toBeLessThan(10_000);
}, 60_000);

test('a line break that opens a paragraph is a blank line', () => {
  imported('<p><b>A</b></p><p><br>B</p>', ['A', '', 'B']);
});

test('the paragraph budget counts blank lines', () => {
  const html = (count: number) => '<pre><b>A' + '\n'.repeat(count) + 'B</b></pre>';
  expect(parseClipboardHtml(html(MAX_FORMATTED_HTML_PARAGRAPHS), formatted)).toBeNull();
  imported(html(MAX_FORMATTED_HTML_PARAGRAPHS - 1), ['A', ...Array(MAX_FORMATTED_HTML_PARAGRAPHS - 2).fill(''), 'B']);
  const start = performance.now();
  expect(parseClipboardHtml(html(1_900_000), formatted)).toBeNull();
  expect(performance.now() - start).toBeLessThan(10_000);
}, 60_000);

test('a table and the text after it count as one paragraph of the budget', () => {
  const count = MAX_FORMATTED_HTML_PARAGRAPHS;
  const fragment = parseClipboardHtml(table + '<p><b>x</b></p>'.repeat(count), formatted);
  expect(fragment?.paragraphs).toHaveLength(count);
  expect(fragment!.paragraphs.every(p => p.runs?.[0]?.format?.other?.bold === true)).toBe(true);
});

test('an empty preformatted segment keeps the trailing space trim of the line before', () => {
  imported('<p><b>A </b><span style="white-space:pre-wrap">&#10;B</span></p>', ['A', 'B']);
});
