import { afterEach, expect, test } from 'bun:test';
import { closeCommandEngines, openCommandEngine, TIMEOUT } from './residentCommandTestSupport';

const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (content: string) => `<w:tc><w:tcPr/>${content}</w:tc>`;
const table = (rows: string[][]) => `<w:tbl><w:tblPr/><w:tblGrid>${'<w:gridCol w:w="1500"/>'.repeat(rows[0]!.length)}</w:tblGrid>${
  rows.map((row) => `<w:tr>${row.map(cell).join('')}</w:tr>`).join('')}</w:tbl>`;
const nested = table([[paragraph('A'), table([[paragraph('N00'), paragraph('N01')], [paragraph('N10'), paragraph('N11')]]) + paragraph('')]]);
const body = paragraph('Before') + nested + paragraph('After');

afterEach(closeCommandEngines);

async function packageBytes(bodyXml: string): Promise<Uint8Array> {
  const { rezipPartsToArrayBuffer, toBytes } = await import('../docx/rezip/parts');
  const parts = new Map<string, Uint8Array>();
  parts.set('[Content_Types].xml', toBytes('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
  parts.set('_rels/.rels', toBytes('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
  parts.set('word/document.xml', toBytes(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

const stack = (error: unknown) => String((error as Error)?.stack ?? error).split('\n').slice(0, 10).join(' | ');

test('a peer joining a document with a nested table addresses every story the host has', async () => {
  const bytes = await packageBytes(body);
  const host = await openCommandEngine([], { bytes, clientId: 8701 });
  const hostStories = host.engine.geometryReader.storyIds().sort();
  expect(hostStories).toContain('body:t0:r0c1:t0:r0c0');
  const state = host.engine.encodeState();
  const report: Record<string, unknown> = { hostStories: hostStories.length };

  try {
    const helperPeer = await openCommandEngine([], { bytes, clientId: 8702, checkpoint: state });
    const stories = helperPeer.engine.geometryReader.storyIds().sort();
    report.helperPath = { ok: true, missing: hostStories.filter((story) => !stories.includes(story)), extra: stories.filter((story) => !hostStories.includes(story)) };
  } catch (error) {
    report.helperPath = { ok: false, error: stack(error) };
  }

  const joiner = await openCommandEngine(['Seed'], { clientId: 8703 });
  try {
    joiner.engine.openSourceDocx(bytes, { state });
    const stories = joiner.engine.geometryReader.storyIds().sort();
    const missing = hostStories.filter((story) => !stories.includes(story));
    const unreadable = stories.filter((story) => { try { joiner.engine.paragraphs(story); return false; } catch { return true; } });
    report.productPath = { ok: true, missing, extra: stories.filter((story) => !hostStories.includes(story)), unreadable,
      nestedText: joiner.engine.paragraphs('body:t0:r0c1:t0:r1c1').map((p) => p.text) };
  } catch (error) {
    report.productPath = { ok: false, error: stack(error) };
  }
  console.log(`JOIN-PROBE ${JSON.stringify(report)}`);
  expect(report.productPath).toMatchObject({ ok: true, missing: [], unreadable: [] });
}, TIMEOUT);
