import { beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { parseDocx } from '../docx';
import { repackDocx } from '../docx/rezip';
import { rezipPartsToArrayBuffer, toBytes, type PartsMap } from '../docx/rezip/parts';
import { readDocxContainer } from '../docx/zipContainer';
import type { Document } from '../types/document';
import { preloadEditWasm } from '../wasm/edit';
import { documentToYrs } from './documentToYrs';
import { createYrsSession, type YrsSession } from './index';
import { captureSessionSave, saveYrsDocx, writeSessionSave } from './saveYrsDocx';
import { paragraphAttrsToFormatting } from './saveFormatting';
import { yrsToDocument } from './yrsToDocument';

type Seeder = 'native' | 'projected';
type Save = 'editor' | 'repack' | 'public';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const OFFICE_DOC = 'application/vnd.openxmlformats-officedocument';

beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(
  resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm')
))));

function fixture(thirdProperties = '', normalProperties = '', extraStyles = ''): Uint8Array {
  const parts: PartsMap = new Map();
  parts.set('[Content_Types].xml', toBytes(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="${OFFICE_DOC}.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="${OFFICE_DOC}.wordprocessingml.styles+xml"/>
</Types>`));
  parts.set('_rels/.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="doc" Type="${R}/officeDocument" Target="word/document.xml"/>
</Relationships>`));
  parts.set('word/_rels/document.xml.rels', toBytes(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="styles" Type="${R}/styles" Target="styles.xml"/>
</Relationships>`));
  parts.set('word/styles.xml', toBytes(`<w:styles xmlns:w="${W}">
  <w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>${normalProperties ? `<w:pPr>${normalProperties}</w:pPr>` : ''}</w:style>
  ${extraStyles}
</w:styles>`));
  parts.set('word/document.xml', toBytes(`<w:document xmlns:w="${W}"><w:body>
  <w:p><w:r><w:t>Finding QA</w:t></w:r></w:p>
  <w:p><w:r><w:t>Conclusion QA</w:t></w:r></w:p>
  ${thirdProperties ? `<w:p><w:pPr>${thirdProperties}</w:pPr><w:r><w:t>Override QA</w:t></w:r></w:p>` : ''}
</w:body></w:document>`));
  return new Uint8Array(rezipPartsToArrayBuffer(parts));
}

async function saveOnce(
  bytes: Uint8Array,
  seeder: Seeder,
  save: Save,
  edit?: (session: YrsSession) => void
): Promise<Uint8Array> {
  const parsed = await parseDocx(bytes.buffer as ArrayBuffer, { preloadFonts: false });
  const session = await createYrsSession({ clientId: 66101 });
  try {
    let base: Document;
    if (seeder === 'native') {
      session.seedFromDocx(bytes);
      base = session.materializeDocx()!;
      expect(base.package.document.content.map((block) =>
        block.type === 'paragraph' ? block.formatting : undefined
      )).toEqual(parsed.package.document.content.map((block) =>
        block.type === 'paragraph' ? block.formatting : undefined
      ));
    } else {
      documentToYrs(session, parsed);
      base = parsed;
    }
    const first = session.paragraphs('body')[0]!;
    expect(first.properties.spaceAfter).toBe(200);
    expect(first.properties.lineSpacing).toBe(276);
    edit?.(session);
    if (save === 'public') {
      session.insertText({ story: 'body', paraId: first.paraId, offset: 0 }, 'Edited ');
      return (await saveYrsDocx(session)).bytes;
    }
    if (save === 'repack') return new Uint8Array(await repackDocx(yrsToDocument(session, base)));
    const capture = captureSessionSave(session);
    const saved = await writeSessionSave(
      session,
      yrsToDocument(session, base),
      capture,
      base.originalBuffer ?? (bytes.buffer as ArrayBuffer),
      {},
      () => false
    );
    return saved.bytes;
  } finally {
    session.destroy();
  }
}

function documentXml(bytes: Uint8Array): string {
  const xml = readDocxContainer(bytes.buffer as ArrayBuffer).text('word/document.xml');
  expect(xml).not.toBeNull();
  return xml!;
}

function spacingTags(xml: string): string[] {
  return xml.match(/<w:spacing\b[^>]*\/>/g) ?? [];
}

function paragraphs(xml: string): string[] {
  return xml.match(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g) ?? [];
}

for (const seeder of ['native', 'projected'] as const) {
  const saves: Save[] = seeder === 'native' ? ['editor', 'repack', 'public'] : ['editor', 'repack'];
  for (const save of saves) {
    test(`${seeder} ${save} keeps document-default spacing inherited through three reopen cycles`, async () => {
      let bytes = fixture();
      for (let cycle = 0; cycle < 3; cycle += 1) {
        bytes = await saveOnce(bytes, seeder, save);
        const xml = documentXml(bytes);
        expect(xml).not.toContain('<w:spacing');
        expect(paragraphs(xml)).toHaveLength(2);
        expect(xml).toContain('Finding QA');
        expect(xml).toContain('Conclusion QA');
        if (save === 'public') expect(xml).toContain('Edited '.repeat(cycle + 1));
      }
    });

    for (const spacing of [
      '<w:spacing w:after="0"/>',
      '<w:spacing w:after="200"/>',
      '<w:spacing w:after="120" w:line="276" w:lineRule="auto"/>',
    ]) {
      test(`${seeder} ${save} preserves only the authored ${spacing} through three reopen cycles`, async () => {
        let bytes = fixture(spacing);
        for (let cycle = 0; cycle < 3; cycle += 1) {
          bytes = await saveOnce(bytes, seeder, save);
          const xml = documentXml(bytes);
          const savedParagraphs = paragraphs(xml);
          expect(savedParagraphs).toHaveLength(3);
          expect(savedParagraphs[0]).not.toContain('<w:spacing');
          expect(savedParagraphs[1]).not.toContain('<w:spacing');
          expect(spacingTags(xml)).toEqual([spacing]);
          expect(spacingTags(savedParagraphs[2]!)).toEqual([spacing]);
        }
      });
    }

    test(`${seeder} ${save} preserves paragraph-style inheritance through three reopen cycles`, async () => {
      const styleProperties = '<w:spacing w:after="480" w:line="360" w:lineRule="auto"/><w:jc w:val="center"/><w:keepNext/><w:widowControl w:val="0"/><w:ind w:left="720"/>';
      const styles = `<w:style w:type="paragraph" w:styleId="Spaced"><w:basedOn w:val="Normal"/><w:pPr>${styleProperties}</w:pPr></w:style>`;
      let bytes = fixture('<w:pStyle w:val="Spaced"/>', '', styles);
      for (let cycle = 0; cycle < 3; cycle += 1) {
        bytes = await saveOnce(bytes, seeder, save, (session) => {
          expect(session.paragraphs('body')[2]!.properties).toMatchObject({
            spaceAfter: 480,
            lineSpacing: 360,
            alignment: 'center',
            keepNext: true,
            widowControl: false,
            indentLeft: 720,
          });
        });
        const xml = documentXml(bytes);
        expect(xml).toContain('<w:pStyle w:val="Spaced"/>');
        for (const tag of ['spacing', 'jc', 'keepNext', 'widowControl', 'ind']) {
          expect(xml).not.toContain(`<w:${tag}`);
        }
      }
    });

    test(`${seeder} ${save} keeps default-style properties inherited through three reopen cycles`, async () => {
      let bytes = fixture('', '<w:jc w:val="center"/><w:keepNext/><w:widowControl w:val="0"/>');
      for (let cycle = 0; cycle < 3; cycle += 1) {
        bytes = await saveOnce(bytes, seeder, save, (session) => {
          expect(session.paragraphs('body')[0]!.properties).toMatchObject({
            alignment: 'center', keepNext: true, widowControl: false,
          });
        });
        const xml = documentXml(bytes);
        for (const tag of ['spacing', 'jc', 'keepNext', 'widowControl']) {
          expect(xml).not.toContain(`<w:${tag}`);
        }
      }
    });

    test(`${seeder} ${save} writes edits that differ from inherited values through three reopen cycles`, async () => {
      let bytes = fixture();
      for (let cycle = 0; cycle < 3; cycle += 1) {
        bytes = await saveOnce(bytes, seeder, save, cycle === 0 ? (session) => {
          const first = session.paragraphs('body')[0]!;
          const position = { paraId: first.paraId, offset: 0 };
          session.setParagraphAttrs({ story: 'body', start: position, end: position }, {
            spaceAfter: 0,
            alignment: 'right',
            indentLeft: 720,
          });
        } : undefined);
        const [edited, untouched] = paragraphs(documentXml(bytes));
        expect(edited).toContain('w:after="0"');
        expect(edited).not.toContain('w:line=');
        expect(edited).toContain('<w:jc w:val="right"/>');
        expect(edited).toContain('w:left="720"');
        expect(untouched).not.toContain('<w:pPr');
      }
    });

    test(`${seeder} ${save} clears a direct spacing override through three reopen cycles`, async () => {
      let bytes = fixture('<w:spacing w:after="0"/>');
      for (let cycle = 0; cycle < 3; cycle += 1) {
        bytes = await saveOnce(bytes, seeder, save, cycle === 0 ? (session) => {
          const paragraph = session.paragraphs('body')[2]!;
          const position = { paraId: paragraph.paraId, offset: 0 };
          session.setParagraphAttrs({ story: 'body', start: position, end: position }, {
            spaceAfter: null,
          });
        } : undefined);
        expect(documentXml(bytes)).not.toContain('<w:spacing');
      }
    });
  }
}

test('the save projection leaves out values that only restate inherited ones', () => {
  const inherited = {
    spaceAfter: 200,
    lineSpacing: 276,
    lineSpacingRule: 'auto' as const,
    alignment: 'center' as const,
  };
  expect(paragraphAttrsToFormatting({ ...inherited }, inherited)).toBeUndefined();
  expect(paragraphAttrsToFormatting({ ...inherited, spaceAfter: 0, lineSpacingRule: 'exact' }, inherited))
    .toEqual(expect.objectContaining({ spaceAfter: 0, lineSpacing: 276, lineSpacingRule: 'exact' }));
  expect(paragraphAttrsToFormatting({ ...inherited, _originalFormatting: { spaceAfter: 200 } }, inherited))
    .toEqual({ spaceAfter: 200 });
  expect(paragraphAttrsToFormatting({ ...inherited, alignment: 'right', _originalFormatting: {} }, inherited))
    .toEqual({ alignment: 'right' });
});
