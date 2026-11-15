import JSZip from 'jszip';
import type * as Yrs from '../index';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const OFFICE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const TARGET = '0E100001';
const TEXT = 'word word';
/** Repeats the last word, so the edit has no unique token alignment and the paragraph is written from the model. */
const APPENDED = ' word';

export type OnOffScope = 'paragraph' | 'run';
export type OnOffCase = 'off' | 'absent' | 'on';
export type SavedValue = OnOffCase | 'missing' | `mixed:${string}`;

/** On/off properties set by a paragraph style, each as its OOXML element under `w:pPr` or a run's `w:rPr`. */
export const ON_OFF_PROPERTIES: ReadonlyArray<{ element: string; scope: OnOffScope }> = [
  { element: 'pageBreakBefore', scope: 'paragraph' },
  { element: 'bidi', scope: 'paragraph' },
  ...['caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'vanish', 'cs', 'snapToGrid']
    .map((element) => ({ element, scope: 'run' as const })),
];

/** Run toggles that a paragraph style and a character style both turning them on cancel out. */
export const TOGGLE_PROPERTIES = ['caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'vanish'];

export const ON_OFF_CASES: ReadonlyArray<{ value: OnOffCase; title: string }> = [
  { value: 'off', title: 'explicitly off under a style that turns it on is saved as an explicit off (changes the effective value)' },
  { value: 'absent', title: 'absent under a style that turns it on is saved absent (keeps source bytes)' },
  { value: 'on', title: 'explicitly on under a style that turns it on is saved as an explicit on (keeps source bytes)' },
];

const direct = (element: string, value: OnOffCase) =>
  value === 'off' ? `<w:${element} w:val="0"/>` : value === 'on' ? `<w:${element}/>` : '';

/** A package whose paragraph `TARGET` has `pPr` and one run with `rPr`, and whose styles part adds `styles`. */
async function packageOf(styles: string, pPr: string, rPr: string): Promise<Uint8Array> {
  const zip = new JSZip();
  const date = new Date('2026-01-01T00:00:00Z');
  const plain = (paraId: string, text: string) => `<w:p w14:paraId="${paraId}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${OFFICE}.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="${OFFICE}.styles+xml"/></Types>`, { date });
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>', { date });
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>', { date });
  zip.file('word/styles.xml', `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>${styles}</w:styles>`, { date });
  zip.file('word/document.xml', `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${plain('0E100000', 'Lead')}<w:p w14:paraId="${TARGET}">${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t>${TEXT}</w:t></w:r></w:p>${plain('0E100002', 'Tail')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`, { date });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

const paragraphStyle = (properties: string) =>
  `<w:style w:type="paragraph" w:styleId="On"><w:name w:val="On"/><w:basedOn w:val="Normal"/>${properties}</w:style>`;

/** One paragraph in style `On`, which turns `element` on, carrying `value` as its direct setting. */
export function onOffFixture(element: string, scope: OnOffScope, value: OnOffCase): Promise<Uint8Array> {
  const style = scope === 'paragraph' ? `<w:pPr><w:${element}/></w:pPr>` : `<w:rPr><w:${element}/></w:rPr>`;
  return packageOf(
    paragraphStyle(style),
    `<w:pStyle w:val="On"/>${scope === 'paragraph' ? direct(element, value) : ''}`,
    scope === 'run' ? direct(element, value) : '',
  );
}

/** A run explicitly off for `element` in character style `OnRun`, inside paragraph style `On`; both styles turn it on. */
export function bothStylesFixture(element: string): Promise<Uint8Array> {
  return packageOf(
    `${paragraphStyle(`<w:rPr><w:${element}/></w:rPr>`)}<w:style w:type="character" w:styleId="OnRun"><w:name w:val="On Run"/><w:rPr><w:${element}/></w:rPr></w:style>`,
    '<w:pStyle w:val="On"/>',
    `<w:rStyle w:val="OnRun"/>${direct(element, 'off')}`,
  );
}

/** A paragraph in no style whose run has the direct run properties `rPr`. */
export const unstyledRunFixture = (rPr: string) => packageOf('', '', rPr);

type Session = Awaited<ReturnType<typeof Yrs.createYrsSession>>;
type FormatDelta = Parameters<Session['formatRange']>[1];

/** `replaced` when the edited paragraph is written from the model instead of by a paragraph patch. */
function savePath(session: Session): 'replaced' | 'patched' | 'unchanged' {
  const plan = session.paragraphSavePlan();
  const body = (plan.splicedParts ?? []).find((part) => part.part.replace(/^\//, '') === 'word/document.xml');
  if (!body) {
    const whole = [...(plan.wholeParts ?? []), ...(plan.touchedParts ?? [])].some((part) => part.part.replace(/^\//, '') === 'word/document.xml');
    return whole ? 'replaced' : 'unchanged';
  }
  if (body.changed.length === 0) return 'unchanged';
  const patched = new Set((body.paragraphPatches ?? []).map((patch) => patch.ordinal));
  return body.changed.some((ordinal) => patched.has(ordinal)) ? 'patched' : 'replaced';
}

/** Applies `format` to the paragraph's text when given, appends a repeated word and saves. */
export async function appendAndSave(yrs: typeof Yrs, bytes: Uint8Array, format?: FormatDelta): Promise<{ path: string; xml: string }> {
  const session = await yrs.createYrsSession();
  try {
    session.openDocx(bytes, true);
    const target = session.paragraphs('body').find((entry) => entry.paraId.toUpperCase() === TARGET);
    if (!target || target.text !== TEXT) throw new Error(`paragraph ${TARGET} not found with its source text`);
    if (format) {
      session.formatRange({
        story: 'body',
        start: { paraId: target.paraId, offset: 0 },
        end: { paraId: target.paraId, offset: target.text.length },
      }, format);
    }
    session.insertText({ story: 'body', paraId: target.paraId, offset: target.text.length }, APPENDED);
    const path = savePath(session);
    const saved = await yrs.saveYrsDocx(session, { updateModifiedDate: false });
    const xml = await (await JSZip.loadAsync(saved.bytes)).file('word/document.xml')!.async('string');
    return { path, xml };
  } finally {
    session.destroy();
  }
}

const tagOf = (xml: string, element: string) => xml.match(new RegExp(`<w:${element}(?=[\\s/>])[^>]*>`))?.[0];
const valueOf = (tag: string | undefined): OnOffCase =>
  tag === undefined ? 'absent' : /\sw:val="(0|false|off)"/.test(tag) ? 'off' : 'on';

/** How the saved target paragraph states `element`: in its `w:pPr`, or the same on every text run. */
export function savedValue(xml: string, element: string, scope: OnOffScope): SavedValue {
  const paragraph = xml.match(new RegExp(`<w:p\\b[^>]*w14:paraId="${TARGET}"[^>]*>[\\s\\S]*?</w:p>`, 'i'))?.[0];
  if (!paragraph) return 'missing';
  if (scope === 'paragraph') {
    const pPr = paragraph.match(/<w:pPr>[\s\S]*?<\/w:pPr>/)?.[0] ?? '';
    return valueOf(tagOf(pPr.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, ''), element));
  }
  const runs = (paragraph.match(/<w:r(?=[\s>])[\s\S]*?<\/w:r>/g) ?? []).filter((run) => /<w:t(?=[\s>])/.test(run));
  const values = [...new Set(runs.map((run) => valueOf(tagOf(run.match(/<w:rPr>[\s\S]*?<\/w:rPr>/)?.[0] ?? '', element))))];
  return values.length === 0 ? 'missing' : values.length === 1 ? values[0]! : `mixed:${values.join('+')}`;
}

/** The effective value of a run property under no style: absent reads as off. */
const effectiveUnstyled = (value: SavedValue) => (value === 'absent' ? 'off' : value);

type Outcome = { path: string; savedAs: string; expected: string };
type SaveCase = { group: string; name: string; run: (api: typeof Yrs) => Promise<Outcome> };

const STYLED = 'on/off properties of a paragraph written from the model on save';
const CHANGED = 'run properties the user changes on a paragraph written from the model on save';

/** Every case: its save must take the replace path and state the expected value. */
export const EXPLICIT_OFF_SAVE_CASES: readonly SaveCase[] = [
  ...ON_OFF_PROPERTIES.flatMap(({ element, scope }) => ON_OFF_CASES.map(({ value, title }): SaveCase => ({
    group: STYLED,
    name: `w:${element} ${title}`,
    run: async (api) => {
      const { path, xml } = await appendAndSave(api, await onOffFixture(element, scope, value));
      return { path, savedAs: savedValue(xml, element, scope), expected: value };
    },
  }))),
  ...TOGGLE_PROPERTIES.map((element): SaveCase => ({
    group: STYLED,
    name: `w:${element} explicitly off under a paragraph style and a character style that both turn it on is saved as an explicit off (keeps source bytes)`,
    run: async (api) => {
      const { path, xml } = await appendAndSave(api, await bothStylesFixture(element));
      return { path, savedAs: savedValue(xml, element, 'run'), expected: 'off' };
    },
  })),
  {
    group: CHANGED,
    name: 'w:caps turned on by the user over an explicit off is saved on (changes the effective value)',
    run: async (api) => {
      const { path, xml } = await appendAndSave(api, await unstyledRunFixture('<w:caps w:val="0"/>'), { other: { allCaps: true } });
      return { path, savedAs: effectiveUnstyled(savedValue(xml, 'caps', 'run')), expected: 'on' };
    },
  },
  {
    group: CHANGED,
    name: 'w:caps turned off by the user over an explicit on is saved off (changes the effective value)',
    run: async (api) => {
      const { path, xml } = await appendAndSave(api, await unstyledRunFixture('<w:caps/>'), { other: { allCaps: null } });
      return { path, savedAs: effectiveUnstyled(savedValue(xml, 'caps', 'run')), expected: 'off' };
    },
  },
  {
    group: CHANGED,
    name: 'w:smallCaps the user never touched keeps its explicit off beside w:caps turned on (keeps source bytes)',
    run: async (api) => {
      const { path, xml } = await appendAndSave(api, await unstyledRunFixture('<w:caps w:val="0"/><w:smallCaps w:val="0"/>'), { other: { allCaps: true } });
      return { path, savedAs: savedValue(xml, 'smallCaps', 'run'), expected: 'off' };
    },
  },
];
