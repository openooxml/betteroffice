import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, it } from 'bun:test';
import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import type { PartsMap } from '../../docx/rezip/parts';
import type { DisplayPage } from './displayList';
import { applyFrameDelta, decodeFrameDelta } from './frameDelta';
import { buildMirrorPage, displayPageHoldsMirrorId } from './mirrorDom';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const WASM = resolve(import.meta.dir, '../../wasm/generated/edit/docx_edit_bg.wasm');
const hasEditWasm = existsSync(WASM);
const FONT = resolve(
  import.meta.dir,
  '../../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';
const ENDNOTES = [1, 2, 3];

let createEditSession: typeof import('../../wasm/edit').createEditSession;
let rezipPartsToArrayBuffer: typeof import('../../docx/rezip/parts').rezipPartsToArrayBuffer;
let toBytes: typeof import('../../docx/rezip/parts').toBytes;
const sessions: ReturnType<typeof createEditSession>[] = [];

beforeAll(async () => {
  if (!hasEditWasm) return;
  const editWasm = await import('../../wasm/edit');
  createEditSession = editWasm.createEditSession;
  ({ rezipPartsToArrayBuffer, toBytes } = await import('../../docx/rezip/parts'));
  await editWasm.preloadEditWasm(new Uint8Array(readFileSync(WASM)));
}, 30_000);

afterEach(() => {
  for (const session of sessions) session.free();
  sessions.length = 0;
});

/** A short body whose endnotes together outgrow the room left on its page. */
function longEndnotesDocx(): Uint8Array {
  const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
  const words = (note: number) =>
    Array.from({ length: 60 }, (_, index) => `n${note}t${index}`).join(' ');
  const body = ENDNOTES.map(
    (id) => `<w:p>${run(`Body ${id}`)}<w:r><w:endnoteReference w:id="${id}"/></w:r></w:p>`
  ).join('');
  const endnotes =
    `<w:endnotes ${W_NS}>` +
    `<w:endnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>` +
    `<w:endnote w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>` +
    ENDNOTES.map((id) => `<w:endnote w:id="${id}"><w:p>${run(words(id))}</w:p></w:endnote>`).join('') +
    `</w:endnotes>`;
  const relationship = (id: string, type: string, target: string) =>
    `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
  const relationships = (entries: string) =>
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries}</Relationships>`;
  const parts: PartsMap = new Map();
  parts.set(
    '[Content_Types].xml',
    toBytes(
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/>` +
        `<Override PartName="/word/endnotes.xml" ContentType="${OFFICE_DOC}.wordprocessingml.endnotes+xml"/>` +
        `</Types>`
    )
  );
  parts.set('_rels/.rels', toBytes(relationships(relationship('rId1', 'officeDocument', 'word/document.xml'))));
  parts.set(
    'word/_rels/document.xml.rels',
    toBytes(relationships(relationship('rId2', 'endnotes', 'endnotes.xml')))
  );
  parts.set(
    'word/document.xml',
    toBytes(
      `<w:document ${W_NS}><w:body>${body}<w:sectPr><w:pgSz w:w="4320" w:h="4320"/>` +
        `<w:pgMar w:top="300" w:right="300" w:bottom="300" w:left="300"/></w:sectPr></w:body></w:document>`
    )
  );
  parts.set('word/endnotes.xml', toBytes(endnotes));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

function displayPages(): DisplayPage[] {
  const session = createEditSession(81);
  sessions.push(session);
  session.seed_from_docx(longEndnotesDocx(), undefined);
  const fontId = session.register_measure_font(new Uint8Array(readFileSync(FONT)));
  const base = {
    bodyStory: 'body',
    regions: {
      sections: [
        {
          properties: {
            pageWidth: 4320,
            pageHeight: 4320,
            marginTop: 300,
            marginRight: 300,
            marginBottom: 300,
            marginLeft: 300,
          },
        },
      ],
    },
    notes: { contents: ENDNOTES.map((id) => ({ id, noteKind: 'endnote', height: 0 })) },
    renderEnv: {},
  };
  const requirements = JSON.parse(session.layout_font_requirements_json(JSON.stringify(base))) as Array<{ key: string }>;
  const fontChains = Object.fromEntries(requirements.map(({ key }) => [key, [fontId]]));
  session.layout_document_with_regions_json(
    JSON.stringify({
      ...base,
      measurement: { fontChains, defaults: { fontSize: 11, fontFamily: 'Calibri' }, authoritativeShaping: true },
    })
  );
  const frame = session.build_display_list_frame(JSON.stringify({ fontChains }), 0);
  return applyFrameDelta(null, decodeFrameDelta(frame)).displayList.pages;
}

const squeezed = (text: string) => text.replace(/\s+/g, '');

it.skipIf(!hasEditWasm)('an endnote the engine splits across two pages is one note in the mirror', () => {
  const pages = displayPages();
  const partsOf = (id: number) =>
    pages.flatMap((page, index) => {
      const primitives = (page.noteAreas ?? [])
        .filter((area) => area.kind === 'endnote')
        .flatMap((area) => area.primitives ?? [])
        .filter((primitive) => primitive.groupId === `endnote-${id}`);
      return primitives.length > 0 ? [{ index, primitives }] : [];
    });
  const split = ENDNOTES.find((id) => partsOf(id).length > 1);
  expect(split).toBeDefined();
  const id = split!;
  const parts = partsOf(id);
  expect(parts).toHaveLength(2);

  const mirrors = pages.map((page) => buildMirrorPage(page));
  const all = (selector: string) => mirrors.flatMap((mirror) => Array.from(mirror.querySelectorAll(selector)));
  expect(all(`#oox-endnote-${id}`)).toHaveLength(1);
  const notes = all(`[data-note-id="${id}"]`);
  expect(notes).toHaveLength(2);
  expect(notes.filter((note) => note.getAttribute('role') === 'doc-endnote')).toHaveLength(1);
  expect(notes.flatMap((note) => Array.from(note.querySelectorAll('[role="doc-backlink"]')))).toHaveLength(1);

  const read = squeezed(mirrors.map((mirror) => mirror.textContent ?? '').join(''));
  for (const { primitives } of parts) {
    const text = squeezed(
      primitives.map((primitive) => ('text' in primitive ? String(primitive.text) : '')).join('')
    );
    expect(text.length).toBeGreaterThan(0);
    expect(read.split(text)).toHaveLength(2);
  }
  expect(pages.map((page) => displayPageHoldsMirrorId(page, `oox-endnote-${id}`))).toEqual(
    pages.map((_, index) => index === parts[0]!.index)
  );
});
