import { expect, spyOn } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import type { YrsLoc, YrsSession, YrsStoryRange } from './index';
import { preflightDocxCommand, applyDocxCommand } from './residentCommandActions';
import { projectDocxEditState } from './residentCommandEditState';
import type {
  DocxCommandApplyContext,
  DocxCommandApplyResult,
  DocxCommandEnvelope,
  DocxCommandKind,
  DocxCommandPayloads,
  DocxCommandSelection,
  ResidentDocxCommandEngine,
  ResidentDocxCommandState,
} from './residentCommandProtocol';
import type { EditSession } from './wasm/index';
import type { ResidentEngineSession } from './residentEngineSession';

export const TIMEOUT = 20_000;
const owned: ResidentDocxCommandEngine[] = [];

export function closeCommandEngines(): void {
  for (const engine of owned.splice(0)) engine.destroy();
}

export function selection(paraId: string, anchor: number, focus = anchor, story = 'body'): DocxCommandSelection {
  return {
    anchor: { story, paraId, offset: anchor },
    focus: { story, paraId, offset: focus },
  };
}

function stateAt(selected: DocxCommandSelection | null): ResidentDocxCommandState {
  return { readOnly: false, selection: selected, storedMarks: null, stickyPositions: new Map() };
}

function basis(version: number): DocxCommandApplyContext {
  return { basis: { epoch: 1, version }, currentVersion: version, replayPreviouslyApplied: false };
}

function envelope<K extends DocxCommandKind>(kind: K, payload: DocxCommandPayloads[K], seq: number, version: number): DocxCommandEnvelope {
  return { session: 'd1-worker', incarnation: 1, seq, basis: { epoch: 1, version }, kind, payload };
}

export interface CommandArm {
  engine: ResidentDocxCommandEngine;
  native: EditSession;
  state: ResidentDocxCommandState;
  version: number;
  seq: number;
}

export interface OpenedCommandArm extends CommandArm {
  resident: ResidentEngineSession;
}

export async function openCommandEngine(
  texts = ['Alpha beta', 'Second'],
  options: { editState?: boolean; bodyXml?: string; clientId?: number; checkpoint?: Uint8Array; preview?: number } = {}
): Promise<OpenedCommandArm> {
  const initial = stateAt(null);
  const unloaded = {} as ResidentDocxCommandEngine;
  expect(() => options.editState
    ? projectDocxEditState(unloaded, initial)
    : preflightDocxCommand(unloaded, initial, envelope('setReadOnly', { readOnly: false }, 1, 0), basis(0))
  ).not.toThrow('not implemented');

  const wasm = await import('./wasm/index');
  await wasm.preloadEditWasm(new Uint8Array(readFileSync(resolve(
    process.cwd(), 'packages/docx/src/wasm/generated/edit/docx_edit_bg.wasm'
  ))));
  const { preloadOpcWasm } = await import('../wasm/opc');
  await preloadOpcWasm(new Uint8Array(readFileSync(resolve(
    process.cwd(), 'packages/docx/src/wasm/generated/opc/ooxml_opc_bg.wasm'
  ))));
  const { rezipPartsToArrayBuffer, toBytes } = await import('../docx/rezip/parts');
  const parts = new Map<string, Uint8Array>();
  parts.set('[Content_Types].xml', toBytes(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
  ));
  parts.set('_rels/.rels', toBytes(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  ));
  parts.set('word/_rels/document.xml.rels', toBytes(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/d1" TargetMode="External"/></Relationships>'
  ));
  parts.set('word/styles.xml', toBytes(
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/></w:style></w:styles>'
  ));
  const escaped = (text: string) => text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
  parts.set('word/document.xml', toBytes(
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>' +
    (options.bodyXml ?? texts.map((text, index) =>
      `<w:p w14:paraId="${(index + 1).toString(16).padStart(8, '0')}"><w:r><w:t xml:space="preserve">${escaped(text)}</w:t></w:r></w:p>`
    ).join('')) + '<w:sectPr/></w:body></w:document>'
  ));
  const { createResidentEngineSession } = await import('./residentEngineSession');
  const captured: { native?: EditSession } = {};
  const create = wasm.createEditSession;
  const capture = spyOn(wasm, 'createEditSession').mockImplementation((clientId) => {
    captured.native = create(clientId);
    return captured.native;
  });
  let resident: Awaited<ReturnType<typeof createResidentEngineSession>>;
  try {
    resident = await createResidentEngineSession(undefined, options.clientId ?? 8101);
  } finally {
    capture.mockRestore();
  }
  if (!captured.native) throw new Error('Resident native session was not captured');
  const raw = captured.native;
  let trackingUndo = false;
  const args = (range: YrsStoryRange): [string, string, number, string, number] =>
    [range.story, range.start.paraId, range.start.offset, range.end.paraId, range.end.offset];
  const methods: Pick<YrsSession,
    'paragraphs' | 'selectionContext' | 'storySegments' | 'cellSelection' | 'tablePayload' |
    'insertText' | 'insertLineBreak' | 'replaceRange' | 'deleteRange' | 'splitParagraph' | 'toggleMark' |
    'formatRange' | 'clearFormatting' | 'setParagraphAttrs' | 'applyParagraphStyle' |
    'beginUndoCapture' | 'canUndo' | 'canRedo'
  > = {
    paragraphs: (story) => JSON.parse(raw.paragraphs(story)),
    selectionContext: (range) => JSON.parse(raw.selection_context(...args(range))),
    storySegments: (story) => JSON.parse(raw.story_segments(story)),
    cellSelection: () => JSON.parse(raw.cell_selection()),
    tablePayload: (story, index) => {
      const json = raw.table_payload(story, index);
      return json === undefined ? null : JSON.parse(json);
    },
    insertText: (at, text, author) => JSON.parse(raw.insert_text(at.story, at.paraId, at.offset, text, author?.name, author?.date)),
    insertLineBreak: (at, author) => JSON.parse(raw.insert_line_break(at.story, at.paraId, at.offset, author?.name, author?.date)),
    replaceRange: (range, text, author) => JSON.parse(raw.replace_range(...args(range), text, author?.name, author?.date)),
    deleteRange: (range, author) => JSON.parse(raw.delete_range(...args(range), author?.name, author?.date)),
    splitParagraph: (at, author) => JSON.parse(raw.split_paragraph(at.story, at.paraId, at.offset, author?.name, author?.date)),
    toggleMark: (range, mark) => raw.toggle_mark(...args(range), JSON.stringify(mark)),
    formatRange: (range, delta) => raw.format_range(...args(range), JSON.stringify(delta)),
    clearFormatting: (range) => raw.clear_formatting(...args(range)),
    setParagraphAttrs: (range, attrs, author) => raw.set_paragraph_attrs(...args(range), JSON.stringify(attrs), author?.name, author?.date),
    applyParagraphStyle: (range, styleId, author) => raw.apply_paragraph_style(...args(range), styleId, author?.name, author?.date),
    beginUndoCapture: () => {
      if (!trackingUndo) raw.track_undo();
      trackingUndo = true;
    },
    canUndo: () => raw.can_undo(),
    canRedo: () => raw.can_redo(),
  };
  const engine = Object.assign({}, methods, resident);
  owned.push(engine);
  const bytes = new Uint8Array(rezipPartsToArrayBuffer(parts));
  if (options.checkpoint) {
    raw.open_docx(bytes, false, undefined, undefined);
    engine.loadState(options.checkpoint);
  } else if (options.preview !== undefined) {
    if (engine.openDocxPreview(bytes, options.preview) === null) throw new Error('Synthetic preview was refused');
  } else {
    engine.openDocx(bytes, undefined, 'd1-synthetic');
  }
  const first = engine.paragraphs('body')[0];
  if (!first) throw new Error('Synthetic document has no body paragraph');
  const selected = selection(first.paraId, 0);
  engine.setSelection(selected.anchor, selected.focus);
  engine.registerFont(new Uint8Array(readFileSync(resolve(
    process.cwd(), 'crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
  ))));
  engine.layoutDocumentWithRegionsRetained(JSON.stringify({
    bodyStory: 'body',
    regions: { sections: [{ sectionId: 'main', properties: {} }] },
    measurement: { defaults: { fontSize: 11, fontFamily: 'Liberation Sans' } },
    renderEnv: {},
  }));
  return { engine, resident: engine, native: raw, state: stateAt(selected), version: 0, seq: 0 };
}

export function command<K extends DocxCommandKind>(
  arm: CommandArm, kind: K, payload: DocxCommandPayloads[K], extra: Partial<DocxCommandEnvelope> = {}
): DocxCommandEnvelope {
  return { ...envelope(kind, payload, ++arm.seq, arm.version), ...extra };
}

export function context(arm: CommandArm, intent: DocxCommandEnvelope, predecessorSelection?: DocxCommandSelection | null): DocxCommandApplyContext {
  return { ...basis(arm.version), basis: intent.basis, predecessorSelection };
}

export function execute(arm: CommandArm, intent: DocxCommandEnvelope, predecessorSelection?: DocxCommandSelection | null): DocxCommandApplyResult {
  const ctx = context(arm, intent, predecessorSelection);
  const checked = preflightDocxCommand(arm.engine, arm.state, intent, ctx);
  if (checked.status === 'refused') return checked;
  const result = applyDocxCommand(arm.engine, arm.state, checked.prepared, ctx);
  if (result.status === 'applied') arm.version += 1;
  return result;
}

export function texts(arm: CommandArm): string[] {
  return arm.engine.paragraphs('body').map((paragraph) => paragraph.text);
}

export function snapshot(arm: CommandArm) {
  return {
    texts: texts(arm), ids: arm.engine.paragraphIdentities(), version: arm.version,
    nativeVersion: arm.engine.proposalEngine.version(), vector: arm.engine.encodeStateVector(),
    canUndo: arm.native.can_undo(), canRedo: arm.native.can_redo(),
    selection: structuredClone(arm.state.selection), storedMarks: structuredClone(arm.state.storedMarks),
    readOnly: arm.state.readOnly,
  };
}

export function expectCaret(result: DocxCommandApplyResult, at: YrsLoc): DocxCommandSelection {
  expect(result.selection).toMatchObject({ anchor: at, focus: at });
  if (!result.selection) throw new Error('Command result lacks selection');
  return result.selection;
}

export function storyRange(selected: DocxCommandSelection): YrsStoryRange {
  return {
    story: selected.anchor.story,
    start: { paraId: selected.anchor.paraId, offset: selected.anchor.offset },
    end: { paraId: selected.focus.paraId, offset: selected.focus.offset },
  };
}
