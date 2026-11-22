import { afterEach, beforeAll, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { rezipPartsToArrayBuffer, toBytes } from '../docx/rezip/parts';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm, unzipContainer } from '../wasm/opc';
import { smartArtFixture } from './__fixtures__/smartArt';
import { createYrsSession, saveYrsDocx, type YrsSession } from './index';

const sessions: YrsSession[] = [];

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(new URL('../wasm/generated/edit/docx_edit_bg.wasm', import.meta.url))));
  await preloadOpcWasm();
}, 30_000);
afterEach(() => { for (const session of sessions.splice(0)) session.destroy(); });

function breakAfterDiagram(kind: 'page' | 'column'): Uint8Array<ArrayBuffer> {
  const source = smartArtFixture(false);
  const document = source.parts.get('word/document.xml')!;
  const paragraph = document.slice(document.indexOf('<w:p w14:paraId="10000001"'), document.indexOf('<w:p w14:paraId="10000002"'));
  source.parts.set('word/document.xml', document.replace(paragraph,
    `<w:p w14:paraId="10000001">${source.drawingRun}<w:r><w:br w:type="${kind}"/><w:t>After the break</w:t></w:r></w:p>`));
  return new Uint8Array(rezipPartsToArrayBuffer(new Map([...source.parts].map(([name, xml]) => [name, toBytes(xml)]))));
}

const count = (xml: string, needle: string) => xml.split(needle).length - 1;

for (const kind of ['page', 'column'] as const) {
  it(`an edit after a ${kind} break that follows an inline SmartArt keeps the diagram, then the break, then the text, in their paragraph`, async () => {
    const session = await createYrsSession();
    sessions.push(session);
    session.seedFromDocx(breakAfterDiagram(kind));
    const target = session.paragraphs('body').find(({ text }) => text.includes('After the break'))!;
    let units = 0;
    let length = 0;
    for (const segment of session.storySegments('body')) {
      if (segment.kind === 'pilcrow') {
        if (segment.paraId === target.paraId) units = length;
        length = 0;
      } else {
        length += segment.kind === 'text' ? segment.text.length : 1;
      }
    }
    session.insertText({ story: 'body', paraId: target.paraId, offset: units }, ' edited');
    const xml = new TextDecoder().decode(unzipContainer((await saveYrsDocx(session, { updateModifiedDate: false })).bytes)['word/document.xml']);
    const start = xml.lastIndexOf('<w:p ', xml.indexOf('dgm:relIds'));
    const paragraph = xml.slice(start, xml.indexOf('</w:p>', start));
    const at = (needle: string) => paragraph.indexOf(needle);
    expect({ diagrams: count(xml, 'dgm:relIds'), breaks: count(xml, `w:type="${kind}"`), edits: count(xml, ' edited') })
      .toEqual({ diagrams: 1, breaks: 1, edits: 1 });
    expect({ breakAfterDiagram: at(`w:type="${kind}"`) > at('dgm:relIds'), textAfterBreak: at(' edited') > at(`w:type="${kind}"`) })
      .toEqual({ breakAfterDiagram: true, textAfterBreak: true });
  });
}
