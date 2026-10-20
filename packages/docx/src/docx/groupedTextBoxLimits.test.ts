import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '.';
import { preloadEditWasm } from '../wasm/edit';
import { residentWorkerFactory } from '../yrs/__fixtures__/residentWorker';
import { ResidentEngineWorkerClient } from '../yrs/residentEngineWorkerClient';
import { rezipPartsToArrayBuffer, toBytes } from './rezip/parts';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const LEAF = 'Innermost story text';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))));

function textBox(id: number, story: string): string {
  return `<wps:wsp><wps:cNvPr id="${id}"/><wps:cNvSpPr txBox="1"/><wps:spPr>` +
    '<a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"/>' +
    `</wps:spPr><wps:txbx><w:txbxContent>${story}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>`;
}

function group(id: number, children: string): string {
  return `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="${id}" name="Group ${id}"/>` +
    '<a:graphic><a:graphicData><wpg:wgp><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/>' +
    `<a:chOff x="0" y="0"/><a:chExt cx="952500" cy="476250"/></a:xfrm></wpg:grpSpPr>${children}</wpg:wgp>` +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
}

function nestedTextBoxes(levels: number): string {
  let story = `<w:p><w:r><w:t>${LEAF}</w:t></w:r></w:p>`;
  for (let level = 1; level <= levels; level++) story = group(level, textBox(100 + level, story));
  return story;
}

function nestedGroups(levels: number): string {
  const story = `<w:p><w:r><w:t>${LEAF}</w:t></w:r></w:p>`;
  return group(1, '<wpg:grpSp>'.repeat(levels) + textBox(2, story) + '</wpg:grpSp>'.repeat(levels));
}

function docx(body: string): Uint8Array<ArrayBuffer> {
  const parts = new Map<string, Uint8Array>();
  const add = (name: string, xml: string) => parts.set(name, toBytes(xml));
  add('[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  add('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="office" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`);
  add('word/document.xml',
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}" ` +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
    `<w:body>${body}<w:sectPr/></w:body></w:document>`);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function storyTexts(value: unknown, texts: string[] = []): string[] {
  if (Array.isArray(value)) {
    for (const item of value) storyTexts(item, texts);
  } else if (value && typeof value === 'object') {
    for (const [key, child] of Object.entries(value)) {
      if (key === 'textBody') texts.push(JSON.stringify(child));
      storyTexts(child, texts);
    }
  }
  return texts;
}

async function openInWorker(bytes: Uint8Array): Promise<string> {
  const startWorker = await residentWorkerFactory();
  const client = new ResidentEngineWorkerClient(startWorker(74_100));
  try {
    return (await client.open(bytes)).hostJson;
  } finally {
    client.destroy();
  }
}

for (const [name, body, rendered, warning] of [
  ['four nested group text boxes', nestedTextBoxes(4), true, undefined],
  ['five nested group text boxes', nestedTextBoxes(5), false, 'nested more than 4 levels'],
  ['twenty nested group text boxes', nestedTextBoxes(20), false, 'nested more than 4 levels'],
  ['a text box under forty nested groups', nestedGroups(40), false, 'nested deeper than 32 levels'],
] as const) {
  test(`${name} open in the parser and the resident worker`, async () => {
    const bytes = docx(body);
    const document = await parseDocx(bytes.buffer, { preloadFonts: false });
    expect(storyTexts(document.package.document.content).some((text) => text.includes(LEAF))).toBe(rendered);
    const warnings = document.warnings ?? [];
    if (warning === undefined) expect(warnings).toEqual([]);
    else expect(warnings.some((entry) => entry.includes(warning))).toBe(true);
    expect(typeof await openInWorker(bytes)).toBe('string');
  });
}
