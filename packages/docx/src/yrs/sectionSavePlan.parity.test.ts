import { afterAll, expect, test } from 'bun:test';
import JSZip from 'jszip';
import type { SectionProperties } from '../types/document';
import { FLAVOURS, syntheticDocx } from './__fixtures__/previewChain';
import { createYrsSession, saveYrsDocx, type YrsSession } from './index';
import { partMetadata } from './residentParts';
import { finalSectionStartMetadata } from './sectionBreak';
import {
  checkProjectedSectionPlan, checkSectionMetadataPlan, referenceProjectedSectionPlan,
  sectionSavePlanParityStats, setSectionMetadata,
} from './sectionSavePlanParity';
import { DocxSectionError } from './sections';
import { sectionPackage } from './sectionTestPackage';

await import('./sections.test');
await import('./sectionBreak.test');
await import('./sectionSaveRegression.test');

afterAll(() => {
  expect(sectionSavePlanParityStats().checked).toBeGreaterThan(0);
});

const MAIN = 'word/document.xml';

async function open(boundary = false, rewrite?: (xml: string) => string): Promise<YrsSession> {
  let bytes = await sectionPackage(boundary);
  if (rewrite) {
    const zip = await JSZip.loadAsync(bytes);
    zip.file(MAIN, rewrite(await zip.file(MAIN)!.async('string')));
    bytes = await zip.generateAsync({ type: 'uint8array' });
  }
  const session = await createYrsSession();
  session.seedFromDocx(bytes);
  return session;
}

const withOpaqueSection = (xml: string) =>
  xml.replace('<w:sectPr>', '<w:sectPr><u:opaque xmlns:u="urn:unknown"/>');

const withHiddenParagraph = (xml: string) => xml.replace(
  '<w:t>First</w:t>',
  '<w:t>First</w:t><u:opaque xmlns:u="urn:unknown"><w:p><w:r><w:t>Hidden</w:t></w:r></w:p></u:opaque>'
);

for (const opaque of [false, true]) {
  test(`new boundary in ${opaque ? 'unmodeled' : 'modeled'} section XML`, async () => {
    const session = await open(false, opaque ? withOpaqueSection : undefined);
    try {
      expect(await checkProjectedSectionPlan(session)).toBeNull();
      session.insertSectionBreakAfter(session.paragraphs('body')[0].paraId, 'continuous');
      expect(await checkProjectedSectionPlan(session)).toMatchObject({
        plan: { edits: [{ paragraph: 0 }, { paragraph: null, properties: { sectionStart: 'continuous' } }] },
        lossy: opaque,
        strict: false,
      });
    } finally { session.destroy(); }
  });
}

test('plans the final start from stored section metadata', async () => {
  const session = await open();
  try {
    const stored = partMetadata(session.materializeDocx()!);
    const metadata = finalSectionStartMetadata(stored, undefined, [], 'nextPage')!;
    expect(await checkSectionMetadataPlan(session, metadata)).toMatchObject({
      plan: { edits: [{ paragraph: null, donor: null, properties: { sectionStart: 'nextPage' } }] },
      lossy: false,
      strict: false,
    });
  } finally { session.destroy(); }
});

test('plans the final start from saved metadata keyed by the closing paragraphs', async () => {
  const session = await open(true);
  try {
    const source = partMetadata(session.materializeDocx()!);
    const saved = JSON.stringify({ ...source, ...source.body });
    const closing = [session.listSections()[0].closingParagraphId];
    for (const prefixes of [closing, [...closing, 'extra']]) {
      const metadata = finalSectionStartMetadata({}, saved, prefixes, 'oddPage')!;
      expect(metadata.body.sections!.map(section => section.id)).toEqual([closing[0], undefined]);
      expect(await checkSectionMetadataPlan(session, metadata)).toMatchObject({
        plan: { edits: [{ paragraph: null, donor: null, properties: { sectionStart: 'oddPage' } }] },
        lossy: false,
        strict: false,
      });
    }
  } finally { session.destroy(); }
});

test('plans no change when there is no stored or saved section start metadata', async () => {
  const session = await open();
  try {
    expect(finalSectionStartMetadata({}, undefined, [], 'nextPage')).toBeUndefined();
    expect(await checkProjectedSectionPlan(session)).toBeNull();
  } finally { session.destroy(); }
});

const withPageBorder = (xml: string) => xml.replace(
  '</w:sectPr>',
  '<w:pgBorders><w:top w:val="single" w:color="ABCDEF"/></w:pgBorders></w:sectPr>'
);

function changeFinalProperties(
  session: YrsSession, edit: (properties: SectionProperties) => void
) {
  const metadata = partMetadata(session.materializeDocx()!);
  const finalSectionProperties = structuredClone(metadata.body.finalSectionProperties ?? {});
  edit(finalSectionProperties);
  return setSectionMetadata(session, { ...metadata, body: { ...metadata.body, finalSectionProperties } });
}

test('the TypeScript planner falls back when only a section watermark changes', async () => {
  const session = await open();
  try {
    const options = changeFinalProperties(session, properties => {
      properties.watermark = {
        kind: 'text', text: 'DRAFT', font: 'Calibri', color: '#C0C0C0',
        semitransparent: true, layout: 'diagonal',
      };
    });
    expect(await referenceProjectedSectionPlan(session, options)).toEqual({ lossy: false, strict: false });
  } finally { session.destroy(); }
});

for (const [key, value] of [['relationshipId', 'rIdBorder'], ['artName', 'apples']] as const) {
  test(`the TypeScript planner falls back when only page border ${key} changes`, async () => {
    const session = await open(false, withPageBorder);
    try {
      const options = changeFinalProperties(session, properties => {
        properties.pageBorders!.top![key] = value;
      });
      expect(await referenceProjectedSectionPlan(session, options)).toEqual({ lossy: false, strict: false });
    } finally { session.destroy(); }
  });
}

test('changing only a page border color transform plans an ordinary save', async () => {
  const session = await open(false, withPageBorder);
  try {
    const options = changeFinalProperties(session, properties => {
      properties.pageBorders!.top!.color!.transforms = [{ kind: 'lumMod', value: 0.75 }];
    });
    expect(await checkProjectedSectionPlan(session, options)).toEqual({ lossy: false, strict: false });
  } finally { session.destroy(); }
});

function changeFirstMargin(session: YrsSession): void {
  const first = session.listSections()[0];
  const left = first.authored.marginLeft ?? 1440;
  const next = left > 0 ? Math.ceil(left) - 1 : 1;
  session.setSectionProperties(first.closingParagraphId, { margins: { left: next } });
}

for (const [index, flavour] of FLAVOURS.entries()) {
  test(`public preview fixture ${flavour} plans section saves`, async () => {
    const session = await createYrsSession();
    const save = () => saveYrsDocx(session, { updateModifiedDate: false });
    try {
      const size = 12;
      session.seedFromDocx(syntheticDocx(flavour, size, index * 100 + size));
      changeFirstMargin(session);
      await save();

      const first = session.paragraphs('body')[0];
      if (first && !session.listSections().some(section => section.closingParagraphId === first.paraId)) {
        let inserted = false;
        try {
          session.insertSectionBreakAfter(first.paraId, 'continuous');
          inserted = true;
        } catch (error) {
          if (!(error instanceof DocxSectionError) || error.code !== 'unsupported') throw error;
        }
        if (inserted) await save();
      }

      const final = session.listSections().at(-1)!;
      const orientation = final.authored.orientation === 'landscape' ? 'portrait' : 'landscape';
      session.setSectionProperties('final', { orientation });
      await save();

      const paragraph = session.paragraphs('body')[0];
      session.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'Edited ');
      changeFirstMargin(session);
      await save();
    } finally { session.destroy(); }
  }, 30_000);
}

test('only section edits use source ordinals after a hidden paragraph', async () => {
  const session = await open(true, withHiddenParagraph);
  try {
    const boundary = session.paragraphs('body')[1].paraId;
    session.setSectionProperties(boundary, { margins: { left: 1800 } });
    const result = await checkProjectedSectionPlan(session, { useLocalProvenance: true });
    expect(result).toMatchObject({
      plan: {
        onlySections: true,
        edits: [{ paragraph: 2, donor: 2, properties: { marginLeft: 1800 } }],
      },
      lossy: false,
      strict: false,
    });
    if (result && 'plan' in result) expect(result.plan!.edits[0]).not.toHaveProperty('bodyParagraphs');
  } finally { session.destroy(); }
});

test('paragraph keys visit block controls before the next body paragraph', async () => {
  const session = await open(true, xml => xml.replace(
    /<w:p w14:paraId="10000001">.*?<\/w:p>/,
    paragraph => `<w:sdt><w:sdtPr><w:id w:val="42"/></w:sdtPr><w:sdtContent>${paragraph}</w:sdtContent></w:sdt>`
  ));
  try {
    session.setSectionProperties(session.listSections()[0].closingParagraphId, { margins: { left: 1800 } });
    expect(await checkProjectedSectionPlan(session)).toMatchObject({
      plan: { edits: [{ donor: 1, properties: { marginLeft: 1800 } }] },
      lossy: false,
      strict: false,
    });
  } finally { session.destroy(); }
});

for (const structural of ['insert', 'remove'] as const) {
  test(`structural ${structural} beside a hidden paragraph requires strict placement`, async () => {
    const session = await open(false, withHiddenParagraph);
    try {
      const [first, second] = session.paragraphs('body');
      if (structural === 'insert') {
        session.insertSectionBreak({ story: 'body', paraId: second.paraId, offset: 3 }, 'continuous');
      } else {
        session.insertSectionBreakAfter(first.paraId, 'continuous');
        session.mergeParagraphs('body', second.paraId);
      }
      const plan = session.paragraphSavePlan();
      const splices = [...(plan.splicedParts ?? []), ...(plan.sectionSplicedParts ?? [])];
      expect(splices.some(splice => structural === 'insert' ? splice.inserted?.length : splice.removed?.length)).toBe(true);
      expect(await checkProjectedSectionPlan(session)).toMatchObject({
        plan: { onlySections: false },
        lossy: false,
        strict: true,
      });
    } finally { session.destroy(); }
  });
}

test('a final section change composes, with ordinary fallback when the source part is missing', async () => {
  const session = await open();
  try {
    const base = session.materializeDocx()!;
    const zip = await JSZip.loadAsync(base.originalBuffer!);
    zip.remove(MAIN);
    const sourceBuffer = await zip.generateAsync({ type: 'arraybuffer' });
    session.setSectionProperties('final', { margins: { right: 1800 } });
    expect(await checkProjectedSectionPlan(session)).toMatchObject({
      plan: { edits: [{ paragraph: null, donor: null, properties: { marginRight: 1800 } }] },
      lossy: false,
      strict: false,
    });
    expect(await checkProjectedSectionPlan(session, { sourceBuffer })).toEqual({ lossy: false, strict: false });
    session.insertSectionBreakAfter(session.paragraphs('body')[0].paraId, 'oddPage');
    expect(await checkProjectedSectionPlan(session, { sourceBuffer })).toEqual({
      error: 'Section source part is unavailable',
    });
  } finally { session.destroy(); }
});

for (const opaque of [false, true]) {
  test(`legacy properties use ${opaque ? 'a refusal' : 'the ordinary save'}`, async () => {
    const session = await open(false, opaque ? withOpaqueSection : undefined);
    try {
      session.insertSectionBreakAfter(session.paragraphs('body')[0].paraId, 'continuous');
      const result = await checkProjectedSectionPlan(session, {
        editOutput: ({ package: { document } }) => {
          document.finalSectionProperties = { ...document.finalSectionProperties, columnCount: 2 };
        },
      });
      expect(result).toEqual(opaque
        ? { error: 'Cannot safely compose section patches with legacy section properties' }
        : { lossy: false, strict: false });
    } finally { session.destroy(); }
  });
}
