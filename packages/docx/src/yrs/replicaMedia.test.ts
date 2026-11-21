import { beforeAll, expect, test } from 'bun:test';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { preloadEditWasm } from '../wasm/edit';
import { preloadOpcWasm } from '../wasm/opc';
import { createYrsSession, type YrsSession } from './index';
import fixture from '../../../../crates/docx-edit/tests/fixtures/behind_objects.json';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';

const root = resolve(import.meta.dir, '../../../..');

beforeAll(async () => {
  await preloadEditWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm'))));
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(import.meta.dir, '../wasm/generated/opc/ooxml_opc_bg.wasm'))));
});

const NS = `${fixture.namespaces} xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"`;
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const picture = (id: number) => `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>`;
const inline = (graphic: string, uri = 'http://schemas.openxmlformats.org/drawingml/2006/picture') => `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="10" name="Holder"/><a:graphic><a:graphicData uri="${uri}">${graphic}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

/** Every way the document can hold the one picture; each row of the outcome table is a holder and a metafile. */
const holders: Record<string, { body: string; header?: string }> = {
  inline: { body: inline(picture(1)) },
  anchored: { body: fixture.image.replace('IMAGE_RANK', '0').replace('IMAGE_BEHIND', '0') },
  group: { body: inline(`<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/></a:xfrm></wpg:grpSpPr>${picture(2)}</wpg:wgp>`, 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup') },
  'shape fill': { body: inline('<wps:wsp><wps:cNvPr id="3" name="Shape"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></wps:spPr><wps:bodyPr/></wps:wsp>', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape') },
  header: { body: '<w:p><w:r><w:t>BODY</w:t></w:r></w:p>', header: inline(picture(4)) },
  'table cell': { body: `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${inline(picture(5))}</w:tc></w:tr></w:tbl><w:p/>` },
};
const metafiles = ['clip-bitmap.emf', 'text.emf'] as const;
const blank = new Uint8Array(fixture.png);

type Outcome = { png: string } | { rejects: string };
interface Row { holder: string; metafile: string; seeded: Outcome; replica: Outcome; blank: Outcome }

/** The replica exports what the seeded session exports, and where the seeded session draws, the picture is on the page. */
function brokenRules(rows: Row[]): string[] {
  return rows.flatMap(({ holder, metafile, seeded, replica, blank }) => {
    const broken: string[] = [];
    if ('png' in seeded && JSON.stringify(seeded) === JSON.stringify(blank)) broken.push(`${holder} / ${metafile}: the picture is not drawn`);
    if (JSON.stringify(seeded) !== JSON.stringify(replica)) broken.push(`${holder} / ${metafile}: seeded ${JSON.stringify(seeded)}, replica ${JSON.stringify(replica)}`);
    return broken;
  });
}

function docx(holder: { body: string; header?: string }, image: Uint8Array, unused?: Uint8Array): Uint8Array {
  const header = holder.header !== undefined;
  const sect = header ? '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr>' : '';
  const parts: PartsMap = new Map([
    ...(unused ? [['word/media/aaa.png', unused] as [string, Uint8Array]] : []),
    ['[Content_Types].xml', toBytes(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>${header ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : ''}</Types>`)],
    ['_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="${REL}/image" Target="media/image.png"/>${header ? `<Relationship Id="rIdHeader" Type="${REL}/header" Target="header1.xml"/>` : ''}</Relationships>`)],
    ['word/document.xml', toBytes(`<w:document ${NS}><w:body>${holder.body}${fixture.body}${sect}</w:body></w:document>`)],
    ...(header ? [
      ['word/header1.xml', toBytes(`<w:hdr ${NS}>${holder.header}</w:hdr>`)] as [string, Uint8Array],
      ['word/_rels/header1.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="${REL}/image" Target="media/image.png"/></Relationships>`)] as [string, Uint8Array],
    ] : []),
    ['word/media/image.png', image],
  ]);
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function outcome(session: YrsSession): Promise<Outcome> {
  try {
    const png = await session.renderPng({ page: 0 });
    return { png: createHash('sha256').update(png).digest('hex') };
  } catch (error) {
    return { rejects: String((error as { code?: unknown }).code ?? (error as Error).message) };
  }
}

async function seeded(bytes: Uint8Array): Promise<{ session: YrsSession; outcome: Outcome }> {
  const session = await createYrsSession();
  session.openDocx(bytes, true);
  return { session, outcome: await outcome(session) };
}

test('a replica bootstrapped from a session exports its metafile pictures as the seeded session does, however the picture is held', async () => {
  const rows: Row[] = [];
  for (const [holder, parts] of Object.entries(holders)) {
    const control = await seeded(docx(parts, blank));
    control.session.destroy();
    for (const metafile of metafiles) {
      const bytes = docx(parts, new Uint8Array(readFileSync(resolve(root, 'crates/ooxml-metafile/tests/fixtures', metafile))));
      const source = await seeded(bytes);
      const replica = await createYrsSession();
      try {
        replica.bootstrapPeer(source.session.encodeState(), source.session.encodePeerMetadata(), bytes, replica.openDocx(bytes, false));
        rows.push({ holder, metafile, seeded: source.outcome, replica: await outcome(replica), blank: control.outcome });
      } finally { source.session.destroy(); replica.destroy(); }
    }
  }
  for (const row of rows) console.log(`ROW ${row.holder} | ${row.metafile} | ${JSON.stringify(row.seeded) === JSON.stringify(row.replica) ? 'equal' : 'differs'} | ${'png' in row.seeded ? 'drawn' : 'rejected'}`);
  expect(brokenRules(rows)).toEqual([]);
}, 120_000);

test('a replica opened again from other bytes exports by the new package, not the map of the old one', async () => {
  const picture = new Uint8Array(readFileSync(resolve(root, 'crates/ooxml-metafile/tests/fixtures/clip-bitmap.emf')));
  const other = new Uint8Array(readFileSync(resolve(root, 'crates/ooxml-metafile/tests/fixtures/text.emf')));
  const before = docx(holders.anchored!, picture), after = docx(holders.anchored!, picture, other);
  const source = await seeded(before);
  const replica = await createYrsSession();
  const reference = await seeded(after);
  try {
    replica.bootstrapPeer(source.session.encodeState(), source.session.encodePeerMetadata(), before, replica.openDocx(before, false));
    await outcome(replica);
    replica.openDocx(after, false);
    expect(await outcome(replica)).toEqual(reference.outcome);
  } finally { source.session.destroy(); replica.destroy(); reference.session.destroy(); }
});
