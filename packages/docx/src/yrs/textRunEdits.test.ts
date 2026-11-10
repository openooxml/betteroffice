import { describe, expect, test } from 'bun:test';
import * as yrs from './index';
import type { DocxEditStep } from './edits';
import {
  MARKED_RUN_CASES,
  PLAIN_EQUIVALENT_STEPS,
  SNIPPET_STEP,
  editSaveReopen,
  holdsProperty,
  savedRunProperties,
  textRunFixture,
} from './__fixtures__/textRuns';
import { SAVE_PATHS, setupStructuralSaves } from './structuralSaveTestSupport';

const save = setupStructuralSaves();
const insert = (paraId: string, runs: unknown): DocxEditStep =>
  ({ op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId }, at: 'end', runs }) as DocxEditStep;

describe('formatted runs in host edits', () => {
  test('a run without marks saves exactly what plain text saves', async () => {
    for (const { name, text, runs } of PLAIN_EQUIVALENT_STEPS) {
      const plain = await editSaveReopen(yrs, [text]);
      const formatted = await editSaveReopen(yrs, [runs]);
      expect(formatted.xml, name).toBe(plain.xml);
      expect(formatted.segments, name).toEqual(plain.segments);
    }
  });

  for (const path of SAVE_PATHS) {
    test(`plain-equivalent runs and the snippet save the same through the ${path} path`, async () => {
      const source = await textRunFixture();
      const edit = (steps: DocxEditStep[]) => (engine: Parameters<Parameters<typeof save>[2]>[0]) => {
        const result = engine.applyEdits({ expectVersion: engine.version(), steps });
        if (!result.ok) throw new Error(JSON.stringify(result.failure));
      };
      const ids = (xml: string) => xml.replace(/w14:(paraId|textId)="[0-9A-F]{8}"/g, 'w14:$1=""');
      for (const { name, text, runs } of PLAIN_EQUIVALENT_STEPS) {
        expect(ids(await save(path, source, edit([runs]))), name).toBe(ids(await save(path, source, edit([text]))));
      }
      const xml = await save(path, source, edit([SNIPPET_STEP]));
      expect(holdsProperty(savedRunProperties(xml, 'No changes needed.'), 'w:b', {})).toBe(true);
      expect(holdsProperty(savedRunProperties(xml, 'See the attached notes for details.'), 'w:i', {})).toBe(true);
      const plain = savedRunProperties(xml, 'checked and approved.');
      expect(holdsProperty(plain, 'w:b', {}) || holdsProperty(plain, 'w:i', {})).toBe(false);
    });
  }

  test('each mark saves its run property and survives a reopen and second save', async () => {
    for (const { marks, saved, absent = [] } of MARKED_RUN_CASES) {
      const { xml, resavedXml } = await editSaveReopen(yrs, [insert('00000001', [{ text: ' Z', marks }])]);
      for (const document of [xml, resavedXml]) {
        const properties = savedRunProperties(document, ' Z');
        for (const [name, attributes] of saved) {
          expect(holdsProperty(properties, name, attributes), `${JSON.stringify(marks)} ${name}`).toBe(true);
        }
        for (const name of absent) expect(holdsProperty(properties, name, {}), name).toBe(false);
      }
    }
  });

  test('false turns bold off whether it is direct or comes from a paragraph or character style', async () => {
    const cases = [
      { paraId: '00000002', off: null },
      { paraId: '00000003', off: '<w:b w:val="0"/>' },
      { paraId: '00000004', off: '<w:b w:val="0"/>' },
    ];
    for (const { paraId, off } of cases) {
      const { xml, resavedXml, reopened } = await editSaveReopen(yrs, [insert(paraId, [{ text: ' kept' }, { text: ' off', marks: { bold: false } }])]);
      for (const document of [xml, resavedXml]) {
        const properties = savedRunProperties(document, ' off') ?? '';
        if (off) expect(properties, paraId).toContain(off);
        else expect(properties, paraId).not.toContain('<w:b');
      }
      expect(resavedXml, paraId).toBe(xml);
      const segment = reopened.find((entry) => entry.kind === 'text' && entry.text.includes(' off'));
      expect(segment && segment.kind === 'text' && Boolean(segment.attributes.bold), paraId).toBe(false);
    }
  });

  test('adjacent runs with equal marks save as one run', async () => {
    const { xml } = await editSaveReopen(yrs, [insert('00000001', [{ text: ' aa', marks: { bold: true } }, { text: 'bb', marks: { bold: true } }])]);
    expect(xml).toMatch(/<w:t(?: [^>]*)?> aabb<\/w:t>/);
  });

  test('a combining sequence split across runs keeps its text through save and reopen', async () => {
    const { xml, reopened } = await editSaveReopen(yrs, [insert('00000001', [{ text: ' e', marks: { bold: true } }, { text: '́', marks: { italic: true } }])]);
    expect(holdsProperty(savedRunProperties(xml, '́'), 'w:i', {})).toBe(true);
    const text = reopened.map((segment) => (segment.kind === 'text' ? segment.text : '')).join('');
    expect(text).toContain('Plain text é');
  });

  test('a refused mark or shape changes nothing', async () => {
    const session = await yrs.createYrsSession({ clientId: 51 });
    try {
      session.seedFromDocx(await textRunFixture());
      const before = session.encodeState();
      const refused = session.applyEdits({
        expectVersion: session.version(),
        steps: [
          { op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId: '00000003' }, at: 'end', text: '!' },
          insert('00000001', [{ text: 'a', marks: { link: { href: 'https://example.com' } } }]),
        ],
      });
      expect(refused).toMatchObject({ ok: false, failure: { code: 'invalid-step', stepIndex: 1 } });
      expect(session.encodeState()).toEqual(before);
      const both = { op: 'insertText', target: { kind: 'paragraph', story: 'body', paraId: '00000001' }, at: 'end', text: 'a', runs: [{ text: 'a' }] };
      expect(() => session.applyEdits({ expectVersion: session.version(), steps: [both as unknown as DocxEditStep] })).toThrow();
      expect(session.encodeState()).toEqual(before);
    } finally {
      session.destroy();
    }
  });
});
