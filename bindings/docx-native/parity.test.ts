import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import * as wasmCore from '../../packages/docx/src/core';
import * as wasmYrs from '../../packages/docx/src/yrs';
import * as nativeCore from '@betteroffice/docx-native';
import * as nativeYrs from '@betteroffice/docx-native/yrs';
import { NATIVE_ENTRIES } from '../../scripts/native-engine';
import { expectSavedPartsEqual } from '../native-parity';
import { documentFixture, exercisePackage } from './package-script';

for (const [entry, source] of Object.entries(NATIVE_ENTRIES.docx)) {
  test(`matches the wasm source exports at ${entry}`, async () => {
    const specifier = '@betteroffice/docx-native' + (entry === '.' ? '' : entry.slice(1));
    const nativeKeys = Object.keys(await import(specifier));
    const wasmKeys = Object.keys(await import(`../../packages/docx/${source}`));
    expect(nativeKeys.sort()).toEqual(wasmKeys.sort());
  });
}

test('the same public package script returns equal results and saved parts', async () => {
  const bytes = await documentFixture();
  const actual = await exercisePackage(nativeYrs, nativeCore, bytes);
  const expected = await exercisePackage(wasmYrs, wasmCore, bytes);
  expect(actual.results).toEqual(expected.results);
  await expectSavedPartsEqual(actual.bytes, expected.bytes);
  const saved = await JSZip.loadAsync(actual.bytes);
  expect(await saved.file('custom/untouched.bin')!.async('uint8array')).toEqual(new Uint8Array([0, 127, 255]));
});

test('language edits share native and wasm state, paragraph patches and styles XML', async () => {
  const zip = await JSZip.loadAsync(await documentFixture());
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const properties = '<w:rFonts w:ascii="Arial"/><w:b/><w:sz w:val="24"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>';
  zip.file('word/document.xml', `<w:document xmlns:w="${w}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:pPr><w:rPr>${properties}</w:rPr></w:pPr><w:r><w:rPr>${properties}</w:rPr><w:t>Package text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`);
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string')).replace('</Types>', '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'));
  zip.file('word/styles.xml', `<w:styles xmlns:w="${w}"><w:docDefaults><w:rPrDefault><w:rPr>${properties}</w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr>${properties}</w:rPr></w:style></w:styles>`);
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  const exercise = async (api: typeof wasmYrs) => {
    const session = await api.createYrsSession({ clientId: 74 });
    try {
      session.seedFromDocx(bytes);
      const paraId = session.paragraphs('body')[0].paraId;
      expect(session.formatTextTarget({ kind: 'paragraph', story: 'body', paraId }, { language: { lang: 'de-DE' } }).ok).toBe(true);
      session.formatRange({ story: 'body', start: { paraId, offset: 8 }, end: { paraId, offset: 12 } }, { language: { eastAsia: null } });
      expect(session.setStyleLanguage('Normal', { lang: 'de-DE' })).toBe(true);
      expect(session.setStyleLanguage(null, { lang: 'de-DE', bidi: null })).toBe(true);
      const state = session.styleLanguages();
      const styles = session.patchStyleLanguagesXml(await zip.file('word/styles.xml')!.async('string'));
      const saved = await api.saveYrsDocx(session, { updateModifiedDate: false });
      return { bytes: saved.bytes, state, styles, segments: session.storySegments('body') };
    } finally { session.destroy(); }
  };
  const actual = await exercise(nativeYrs);
  const expected = await exercise(wasmYrs);
  expect(actual.state).toEqual({ Normal: { lang: 'de-DE' }, '': { lang: 'de-DE', bidi: null } });
  expect(actual.state).toEqual(expected.state);
  expect(actual.styles).toEqual(expected.styles);
  expect(actual.segments).toEqual(expected.segments);
  await expectSavedPartsEqual(actual.bytes, expected.bytes);
  const saved = await JSZip.loadAsync(actual.bytes);
  expect(await saved.file('word/styles.xml')!.async('string')).toBe(actual.styles);
  const document = await saved.file('word/document.xml')!.async('string');
  expect(document).not.toContain('w:val="en-US"');
  expect(document).toContain('<w:lang w:val="de-DE" w:bidi="ar-SA"/>');
});

test('proposal rounds, decisions and withdrawals match through the package API', async () => {
  const zip = await JSZip.loadAsync(await documentFixture());
  const paragraphs = ['alpha beta alpha', 'Second paragraph', 'Third text']
    .map((text, index) => `<w:p w14:paraId="0000000${index + 1}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`)
    .join('');
  zip.file('word/document.xml', `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${paragraphs}<w:sectPr/></w:body></w:document>`);
  const bytes = await zip.generateAsync({ type: 'uint8array' });
  const suggest = { author: 'Parity', date: '2026-01-01T00:00:00Z' };
  const exercise = async (api: typeof wasmYrs) => {
    const session = await api.createYrsSession({ clientId: 75 });
    try {
      session.seedFromDocx(bytes);
      const { sessionId } = session.paragraphIdentities();
      const [first, second, third] = session.paragraphs('body').map(({ paraId }) => ({ kind: 'session' as const, sessionId, story: 'body' as const, paraId }));
      const steps: unknown[] = [];
      const step = <T,>(result: T) => { steps.push(result); return result; };
      const version = session.version();
      expect(step(session.proposeChanges({
        expectVersion: version,
        proposals: [
          { id: 'replace', paragraph: first!, suggest, op: 'replaceText', search: 'alpha', replaceWith: 'gamma', occurrence: 'all' },
          { id: 'insert', paragraph: second!, suggest, op: 'insertText', at: 'end', text: ' added' },
          { id: 'delete', paragraph: third!, suggest, op: 'replaceText', search: 'Third ', replaceWith: '' },
        ],
      })).ok).toBe(true);
      const current = session.getProposals();
      step(current);
      const refusal = (result: wasmYrs.DocxProposalResult) => {
        step(result);
        return result.ok ? 'ok' : result.failure.code;
      };
      expect(refusal(session.proposeChanges({
        expectVersion: 'stale',
        proposals: [{ id: 'late', paragraph: third!, suggest, op: 'insertText', at: 'start', text: 'late ' }],
      }))).toBe('stale-version');
      expect(refusal(session.proposeChanges({
        expectVersion: current.version,
        proposals: [{ id: 'absent', paragraph: first!, suggest, op: 'replaceText', search: 'absent', replaceWith: 'x' }],
      }))).not.toBe('ok');
      expect(refusal(session.proposeChanges({
        expectVersion: current.version,
        proposals: [{ id: 'insert', paragraph: first!, suggest, op: 'insertText', at: 'start', text: 'other ' }],
      }))).toBe('proposal-id-conflict');
      expect(refusal(session.setProposalStates({
        expectVersion: current.version,
        expectPreviewVersion: current.previewVersion + 1,
        changes: [{ id: 'replace', state: 'accepted' }],
      }))).toBe('stale-preview');
      expect(refusal(session.setProposalStates({
        expectVersion: current.version,
        expectPreviewVersion: current.previewVersion,
        changes: [{ id: 'unknown', state: 'accepted' }],
      }))).toBe('unknown-proposal');
      expect(step(session.setProposalStates({
        expectVersion: current.version,
        expectPreviewVersion: current.previewVersion,
        changes: [{ id: 'replace', state: 'accepted' }, { id: 'delete', state: 'rejected' }],
      })).ok).toBe(true);
      const decided = session.getProposals();
      step(decided);
      expect(step(session.withdrawProposals({
        expectVersion: decided.version,
        expectPreviewVersion: decided.previewVersion,
        ids: ['replace', 'delete'],
      })).ok).toBe(true);
      step(session.getProposals());
      const texts = session.paragraphs('body').map(({ text }) => text);
      const saved = await api.saveYrsDocx(session, { updateModifiedDate: false });
      return { steps: JSON.parse(JSON.stringify(steps).replaceAll(sessionId, 'session')), texts, bytes: saved.bytes };
    } finally { session.destroy(); }
  };
  const actual = await exercise(nativeYrs);
  const expected = await exercise(wasmYrs);
  expect(actual.texts[0]).toBe('gamma beta gamma');
  expect(actual.texts[2]).toBe('Third text');
  expect(actual.steps).toEqual(expected.steps);
  expect(actual.texts).toEqual(expected.texts);
  await expectSavedPartsEqual(actual.bytes, expected.bytes);
});

test('structured exports and content controls match through the package API', async () => {
  for (const fixture of [
    '../../crates/docx-edit/tests/fixtures/structured-export/principal.docx',
    '../../packages/docx/src/yrs/__fixtures__/content-controls/template.docx',
  ]) {
    const bytes = new Uint8Array(readFileSync(new URL(fixture, import.meta.url)));
    for (const revisionView of ['accepted', 'original', 'markup'] as const) {
      const options = { revisionView, includeFormatting: false, maxBlocks: 10_000, maxBytes: 8_388_608 };
      expect(await nativeCore.exportDocxStructured(bytes, options)).toEqual(await wasmCore.exportDocxStructured(bytes, options));
      expect(await nativeCore.exportDocxMarkdown(bytes, options)).toEqual(await wasmCore.exportDocxMarkdown(bytes, options));
    }
    expect(await nativeCore.listDocxContentControls(bytes)).toEqual(await wasmCore.listDocxContentControls(bytes));
    const query = { kind: 'tag', tag: 'account.reference' } as const;
    expect(await nativeCore.findDocxContentControls(bytes, query)).toEqual(await wasmCore.findDocxContentControls(bytes, query));
    expect(await nativeCore.fontPreflight(bytes)).toEqual(await wasmCore.fontPreflight(bytes));
  }
});
