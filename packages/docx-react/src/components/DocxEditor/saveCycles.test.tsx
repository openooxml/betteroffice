import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, beforeAll, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createRef } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { preloadEditWasm } from '@betteroffice/docx/wasm/edit';
import { rezipPartsToArrayBuffer, toBytes } from '@betteroffice/docx/docx/rezip/parts';
import { unzipContainer } from '@betteroffice/docx/docx/wasm';
import { DocxEditor, type DocxEditorRef } from '../../index';

const { act, cleanup, render } = await import('@testing-library/react');
const quiet = { error: console.error, warn: console.warn };

beforeAll(async () => {
  if (!window.document.fonts) {
    Object.defineProperty(window.document, 'fonts', {
      value: { addEventListener: () => {}, removeEventListener: () => {}, ready: Promise.resolve() },
      configurable: true,
    });
  }
  await preloadEditWasm(
    new Uint8Array(
      readFileSync(resolve(import.meta.dir, '../../../../docx/src/wasm/generated/edit/docx_edit_bg.wasm'))
    )
  );
  console.error = () => {};
  console.warn = () => {};
});
afterEach(cleanup);
afterAll(async () => {
  console.error = quiet.error;
  console.warn = quiet.warn;
  if (ownsDom) await GlobalRegistrator.unregister();
});

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS =
  `xmlns:w="${W}" xmlns:r="${R}" ` +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const PNG = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='),
  (char) => char.charCodeAt(0)
);
const t = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const commented = (id: number, text: string) =>
  `<w:commentRangeStart w:id="${id}"/>${t(text)}<w:commentRangeEnd w:id="${id}"/>` +
  `<w:r><w:commentReference w:id="${id}"/></w:r>`;
const picture = (id: number) =>
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
  `<pic:nvPicPr><pic:cNvPr id="${id}" name="image${id}.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
  '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>';
const INLINE =
  '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  `<wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Picture 1"/>${picture(1)}</wp:inline></w:drawing></w:r></w:p>`;
const ANCHORED =
  '<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="2" ' +
  'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>1828800</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="914400" cy="457200"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="2" name="Picture 2"/>' +
  `${picture(2)}</wp:anchor></w:drawing></w:r>${t('Anchored image')}</w:p>`;

function fixture(): ArrayBuffer {
  const body =
    `<w:p>${t('QA Title')}</w:p>` +
    `<w:p>${commented(0, 'Finding QA')}</w:p>` +
    INLINE +
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
    `<w:tr><w:tc><w:p>${commented(1, 'Cell finding')}</w:p></w:tc><w:tc><w:p>${t('B1')}</w:p></w:tc></w:tr>` +
    `<w:tr><w:tc><w:p>${t('A2')}</w:p></w:tc><w:tc><w:p>${t('B2')}</w:p></w:tc></w:tr></w:tbl>` +
    `<w:p>${t('Page ')}<w:fldSimple w:instr=" PAGE ">${t('1')}</w:fldSimple></w:p>` +
    ANCHORED +
    `<w:p>${t('Conclusion QA')}</w:p>` +
    `<w:p>${t('Closing')}</w:p>` +
    '<w:sdt><w:sdtPr><w:id w:val="7"/><w:tag w:val="qa"/></w:sdtPr><w:sdtContent>' +
    `<w:p>${t('Controlled text')}</w:p></w:sdtContent></w:sdt>` +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/>' +
    '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
  const parts = new Map<string, Uint8Array>(
    Object.entries({
      '[Content_Types].xml':
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
        '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>',
      '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
      'word/_rels/document.xml.rels':
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rIdC" Type="${R}/comments" Target="comments.xml"/>` +
        `<Relationship Id="rIdH" Type="${R}/header" Target="header1.xml"/>` +
        `<Relationship Id="rIdF" Type="${R}/footer" Target="footer1.xml"/>` +
        `<Relationship Id="rIdImg" Type="${R}/image" Target="media/image1.png"/></Relationships>`,
      'word/document.xml': `<w:document ${NS}><w:body>${body}</w:body></w:document>`,
      'word/comments.xml':
        `<w:comments xmlns:w="${W}">` +
        `<w:comment w:id="0" w:author="QA"><w:p>${t('Body comment')}</w:p></w:comment>` +
        `<w:comment w:id="1" w:author="QA"><w:p>${t('Cell comment')}</w:p></w:comment></w:comments>`,
      'word/header1.xml': `<w:hdr xmlns:w="${W}"><w:p>${t('Header QA')}</w:p></w:hdr>`,
      'word/footer1.xml': `<w:ftr xmlns:w="${W}"><w:p>${t('Page ')}<w:fldSimple w:instr=" PAGE ">${t('1')}</w:fldSimple></w:p></w:ftr>`,
    }).map(([name, xml]) => [name, toBytes(xml)])
  );
  parts.set('word/media/image1.png', PNG);
  return rezipPartsToArrayBuffer(parts);
}

async function until(done: () => boolean) {
  for (let attempt = 0; attempt < 300 && !done(); attempt += 1) {
    await act(async () => {
      await new Promise((resolve) => setTimeout(resolve, 10));
    });
  }
  expect(done()).toBe(true);
}

const documentXml = (buffer: ArrayBuffer) =>
  new TextDecoder().decode(unzipContainer(new Uint8Array(buffer))['word/document.xml']);
const paragraphs = (xml: string) => xml.match(/<w:p[ >/]/g)?.length ?? 0;
const markers = (xml: string, id: number) =>
  [...xml.matchAll(new RegExp(`<w:comment(RangeStart|RangeEnd|Reference) w:id="${id}"/>`, 'g'))].map(
    (match) => match[1]
  );
const drawings = (xml: string) =>
  [...xml.matchAll(/<w:p\b(?:(?!<w:p\b).)*?<w:drawing>.*?<\/w:p>/g)].map((match) =>
    match[0].replace(/ w14:(paraId|textId)="[^"]*"/g, '')
  );

for (const typing of [false, true]) {
  test(`three editor saves ${typing ? 'with' : 'without'} typing keep comment ranges and paragraphs`, async () => {
    let buffer = fixture();
    const source = documentXml(buffer);
    for (let cycle = 0; cycle < 3; cycle += 1) {
      const ref = createRef<DocxEditorRef>();
      const view = render(<DocxEditor ref={ref} documentBuffer={buffer} />);
      await until(() => ref.current?.commands.getState('save').enabled === true);
      if (typing) {
        const session = ref.current!.getEditorRef()!.getYrsSession()!;
        const [title] = session.paragraphs('body');
        await act(async () => {
          session.insertText({ story: 'body', paraId: title!.paraId, offset: 0 }, 'QA ');
        });
      }
      let saved: ArrayBuffer | null = null;
      await act(async () => {
        saved = await ref.current!.save();
      });
      view.unmount();
      expect(saved).not.toBeNull();
      buffer = saved!;
      const xml = documentXml(buffer);
      console.log(`CYCLE typing=${typing} ${cycle}`, paragraphs(source), paragraphs(xml), markers(xml, 0), markers(xml, 1));
      console.log('DRAWINGS-SOURCE', JSON.stringify(drawings(source)));
      console.log('DRAWINGS-SAVED', JSON.stringify(drawings(xml)));
      if (cycle === 0) console.log('XML', xml);
      expect([cycle, markers(xml, 0), markers(xml, 1)]).toEqual([
        cycle,
        ['RangeStart', 'RangeEnd', 'Reference'],
        ['RangeStart', 'RangeEnd', 'Reference'],
      ]);
      expect(xml).toMatch(/<w:tc>.*<w:commentRangeStart w:id="1"\/>.*Cell finding.*<w:commentRangeEnd w:id="1"\/>.*<\/w:tc>/);
      expect([cycle, paragraphs(xml)]).toEqual([cycle, paragraphs(source)]);
      expect(drawings(xml)).toEqual(drawings(source));
    }
  });
}
