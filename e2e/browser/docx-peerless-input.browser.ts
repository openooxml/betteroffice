import { expect, test, type Locator, type Page } from 'playwright/test';
import { pinnedEvaluate } from './pinned-evaluate';
import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import { applyFrameDelta, decodeFrameDelta } from '../../packages/docx/src/layout/render/frameDelta';
import type {
  DocxCommandAck,
  DocxCommandEnvelope,
  DocxCommandFrame,
} from '../../packages/docx/src/yrs/residentCommandProtocol';
import type {
  ResidentEngineWorkerRequest,
  ResidentEngineWorkerResponse,
} from '../../packages/docx/src/yrs/residentEngineWorkerProtocol';

type HeldMessage = { worker: Worker; message: ResidentEngineWorkerResponse };
type CommandConfirmation = Pick<DocxCommandAck, 'session' | 'incarnation' | 'epoch' | 'throughSeq' | 'version'>;

interface InputProbe {
  commands: DocxCommandEnvelope[];
  confirmedAtSend: Record<number, CommandConfirmation>;
  acks: DocxCommandAck[];
  frames: DocxCommandFrame[];
  initial: DocxCommandAck | null;
  holdAcks: boolean;
  holdFrames: boolean;
  contextBytes: number | null;
  refuseNext: boolean;
  editDelayMs: number;
  heldAcks: HeldMessage[];
  heldFrames: HeldMessage[];
  mainInstances: number;
  worker?: Worker;
  releaseAcks(): void;
  releaseFrames(): void;
}

interface HarnessCheckpoint {
  stage: 'installed' | 'reconciled' | 'microtask' | 'raf';
  transaction: number;
  revision: number;
  coversThroughSeq: number;
  retired: number[];
  remaining: number[];
  pixels: string[];
}

interface BrowserHarness {
  capabilities: { presentation: boolean; mode: boolean; export: boolean; viewport?: boolean };
  checkpoints: HarnessCheckpoint[];
  registerWorker(worker: Worker): void;
  armPresentation(coversThroughSeq: number): void;
  exportBytes(): Promise<number[]>;
  setViewport?(pages: number[]): Promise<void>;
}

declare global {
  interface Window {
    __docxPeerlessInput: InputProbe;
    __docxPeerlessHarness?: BrowserHarness;
    __docxEvictedCanvas?: HTMLCanvasElement;
  }
}

test.use({ launchOptions: { executablePath: process.env.CHROME || undefined } });

async function installProbe(page: Page) {
  await page.addInitScript(() => {
    const replayed = new WeakSet<Event>();
    let confirmation: CommandConfirmation | undefined;
    const confirm = ({ session, incarnation, epoch, throughSeq, version }: DocxCommandAck) => {
      confirmation = { session, incarnation, epoch, throughSeq, version };
    };
    const dispatch = ({ worker, message }: HeldMessage) => {
      if (message.ok && message.commandAck) confirm(message.commandAck);
      const event = new MessageEvent('message', { data: message });
      replayed.add(event);
      worker.dispatchEvent(event);
    };
    const probe: InputProbe = {
      commands: [],
      confirmedAtSend: {},
      acks: [],
      frames: [],
      initial: null,
      holdAcks: false,
      holdFrames: false,
      contextBytes: null,
      refuseNext: false,
      editDelayMs: 0,
      heldAcks: [],
      heldFrames: [],
      mainInstances: 0,
      releaseAcks() {
        probe.holdAcks = false;
        for (const held of probe.heldAcks.splice(0)) dispatch(held);
      },
      releaseFrames() {
        probe.holdFrames = false;
        for (const held of probe.heldFrames.splice(0)) dispatch(held);
      },
    };
    window.__docxPeerlessInput = probe;
    const instantiate = WebAssembly.instantiate;
    WebAssembly.instantiate = function (...args: Parameters<typeof instantiate>) {
      probe.mainInstances += 1;
      return Reflect.apply(instantiate, WebAssembly, args);
    } as typeof instantiate;
    const streaming = WebAssembly.instantiateStreaming;
    WebAssembly.instantiateStreaming = function (...args: Parameters<typeof streaming>) {
      probe.mainInstances += 1;
      return Reflect.apply(streaming, WebAssembly, args);
    } as typeof streaming;
    WebAssembly.Instance = new Proxy(WebAssembly.Instance, {
      construct(target, args, newTarget) {
        probe.mainInstances += 1;
        return Reflect.construct(target, args, newTarget);
      },
    });
    const adjustContext = (frame: DocxCommandFrame | undefined) => {
      const context = frame?.payload.echoContext;
      if (context && probe.contextBytes !== null) context.serializedBytes = probe.contextBytes;
    };
    const NativeWorker = window.Worker;
    let delivery = Promise.resolve();
    let delayed = 0;
    window.Worker = class extends NativeWorker {
      constructor(url: string | URL, options?: WorkerOptions) {
        super(url, options);
        this.addEventListener('message', (event: MessageEvent<ResidentEngineWorkerResponse>) => {
          if (replayed.has(event)) return;
          const message = event.data;
          if (!message.ok) return;
          if (message.commandInitial) {
            probe.worker = this;
            adjustContext(message.commandInitial.frame);
            probe.initial = message.commandInitial;
            if (message.commandInitial.frame) probe.frames.push(message.commandInitial.frame);
          }
          const ack = message.commandAck;
          if (ack) {
            adjustContext(ack.frame);
            probe.acks.push(ack);
            if (ack.frame) probe.frames.push(ack.frame);
          }
          if (message.commandFrame) {
            adjustContext(message.commandFrame.frame);
            probe.frames.push(message.commandFrame.frame);
          }
          if (ack && probe.holdAcks) {
            event.stopImmediatePropagation();
            probe.heldAcks.push({ worker: this, message });
          } else if (ack?.frame && probe.holdFrames) {
            event.stopImmediatePropagation();
            const { frame, ...bareAck } = ack;
            probe.heldFrames.push({
              worker: this,
              message: {
                id: 0,
                ok: true,
                commandFrame: {
                  session: ack.session,
                  incarnation: ack.incarnation,
                  epoch: ack.epoch,
                  frame,
                },
              },
            });
            dispatch({ worker: this, message: { ...message, commandAck: bareAck } });
          } else if (message.commandFrame && probe.holdFrames) {
            event.stopImmediatePropagation();
            probe.heldFrames.push({ worker: this, message });
          } else if (ack) confirm(ack);
        }, true);
      }

      postMessage(message: unknown, transferOrOptions?: Transferable[] | StructuredSerializeOptions) {
        let outgoing = message;
        const request = message as ResidentEngineWorkerRequest;
        if (request.type === 'edit') {
          if (probe.refuseNext) {
            probe.refuseNext = false;
            outgoing = {
              ...request,
              command: {
                ...request.command,
                basis: { ...request.command.basis, version: request.command.basis.version + 1 },
                expectVersion: true,
              },
            };
          }
          const command = structuredClone((outgoing as typeof request).command);
          probe.commands.push(command);
          if (confirmation) probe.confirmedAtSend[command.seq] = confirmation;
        }
        const post = () => {
          if (Array.isArray(transferOrOptions)) super.postMessage(outgoing, transferOrOptions);
          else super.postMessage(outgoing, transferOrOptions);
        };
        const delay = request.type === 'edit' ? probe.editDelayMs : 0;
        if (delay === 0 && delayed === 0) return post();
        delayed += 1;
        delivery = delivery.then(() => new Promise((resolve) => setTimeout(resolve, delay))).then(() => {
          delayed -= 1;
          post();
        });
      }
    };
  });
}

async function documentBytes(pageCount = 1, feature = 'plain') {
  const zip = new JSZip();
  zip.file('[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>');
  zip.file('_rels/.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>');
  zip.file('word/document.xml',
    '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:body>' +
    (feature === 'table' ? '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Table cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' : '') +
    (feature === 'PAGE' ? '<w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p>' : '') +
    Array.from({ length: pageCount }, (_, index) =>
      (feature === 'section' && index === 32 ? '<w:p><w:pPr><w:sectPr><w:pgSz w:w="9000" w:h="15840"/></w:sectPr></w:pPr></w:p>' : '') +
      '<w:p>' + (feature === 'keep' && (index === 0 || index === 31 || index === 32)
        ? '<w:pPr><w:pStyle w:val="Heading1"/><w:outlineLvl w:val="0"/><w:keepNext/><w:keepLines/></w:pPr>'
        : index ? '<w:pPr><w:pageBreakBefore/></w:pPr>' : '') +
      '<w:r><w:t>' + (index ? `page ${index + 1}` : 'abcd') + '</w:t></w:r></w:p>'
    ).join('') + '<w:sectPr>' +
    '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>' +
    '</w:sectPr></w:body></w:document>');
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function openHarness(page: Page, pageCount = 1, buffer?: Buffer) {
  await installProbe(page);
  await page.goto('/?format=docx&peerless=1');
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'peerless-input.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: buffer ?? await documentBytes(pageCount),
  });
  await expect.poll(() => page.evaluate(() => window.__docxPeerlessHarness?.capabilities), {
    message: 'Preview must be built with DOCX_PEERLESS_HARNESS=1', timeout: 5_000,
  }).toMatchObject({ presentation: true, mode: true });
  await expect.poll(() => page.evaluate(() => !!window.__docxPeerlessInput.initial), { timeout: 60_000 }).toBe(true);
  await page.evaluate(() => window.__docxPeerlessHarness!.registerWorker(window.__docxPeerlessInput.worker!));
  expect(await page.evaluate(() => window.__docxPeerlessHarness!.capabilities.export)).toBe(true);
  await expect(page.locator('[data-peerless-docx] [data-page-id] canvas').first()).toBeVisible();
}

async function focusHarness(page: Page) {
  await page.locator('[data-peerless-docx] [data-page-id] canvas').first().click({ position: { x: 150, y: 150 } });
  const input = page.locator('[data-peerless-docx] textarea');
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+Home');
  await settled(page);
  return input;
}

test('font picker prepares the selected family and retains shaped glyph widths', async ({ page }) => {
  test.skip(process.env.DOCX_PEERLESS_HARNESS !== '1', 'Font command probes require DOCX_PEERLESS_HARNESS=1');
  const zip = await JSZip.loadAsync(await documentBytes());
  const xml = await zip.file('word/document.xml')!.async('string');
  zip.file('word/document.xml', xml.replace('<w:r>', '<w:r><w:rPr><w:rFonts w:ascii="Liberation Sans" w:hAnsi="Liberation Sans"/></w:rPr>'));
  await openHarness(page, 1, await zip.generateAsync({ type: 'nodebuffer' }));
  const input = await focusHarness(page);
  await input.press('ControlOrMeta+End');
  await input.press('Shift+Home');
  await settled(page);
  await page.getByRole('combobox', { name: 'Select font family' }).click();
  await page.getByRole('option', { name: 'Times New Roman', exact: true }).click();
  await expect.poll(() => page.evaluate(() => window.__docxPeerlessInput.commands.some(command => command.kind === 'formatRange'))).toBe(true);
  await settled(page);
  const commands = await page.evaluate(() => window.__docxPeerlessInput.commands.map(command => ({ kind: command.kind,
    chains: command.kind === 'setFontEnvironment' ? (command.payload as { chains: Record<string, number[]> }).chains : undefined })));
  const formatIndex = commands.findIndex(command => command.kind === 'formatRange');
  const environment = commands.slice(0, formatIndex).find(command => command.kind === 'setFontEnvironment');
  expect(environment?.chains?.['times new roman|0|0']?.length).toBeGreaterThan(0);
  const frames = await page.evaluate(() => window.__docxPeerlessInput.frames.map(frame => Array.from(new Uint8Array(frame.payload.frame))));
  let retained = applyFrameDelta(null, decodeFrameDelta(Uint8Array.from(frames[0]!)));
  const original = retained.pages[0]!.page.primitives.flatMap(primitive => primitive.kind === 'glyphRun' && !primitive.listMarker ? [primitive] : []);
  for (const frame of frames.slice(1)) retained = applyFrameDelta(retained, decodeFrameDelta(Uint8Array.from(frame)));
  const selected = retained.pages[0]!.page.primitives.flatMap(primitive => primitive.kind === 'glyphRun' && !primitive.listMarker ? [primitive] : []);
  expect(original.map(run => run.text).join('')).toBe('abcd');
  expect(selected.map(run => run.text).join('')).toBe('abcd');
  expect(selected.flatMap(run => run.glyphs.map(glyph => glyph.advance))).not.toEqual(original.flatMap(run => run.glyphs.map(glyph => glyph.advance)));
});

async function harnessText(page: Page) {
  const bytes = await pinnedEvaluate(page, () => window.__docxPeerlessHarness!.exportBytes());
  const zip = await JSZip.loadAsync(Buffer.from(bytes));
  const xml = await zip.file('word/document.xml')!.async('string');
  return page.evaluate((xml) => Array.from(
    new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagNameNS('*', 't')
  ).map((node) => node.textContent ?? '').join(''), xml);
}

async function open(page: Page, peerless = true, contextBytes: number | null = null, pageCount = 1, buffer?: Buffer) {
  await installProbe(page);
  await page.goto(peerless ? '/?format=docx&peerless=1' : '/?format=docx&peerless=0');
  await page.evaluate((bytes) => { window.__docxPeerlessInput.contextBytes = bytes; }, contextBytes);
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'peerless-input.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: buffer ?? await documentBytes(pageCount),
  });
  if (peerless) {
    const root = page.locator('[data-peerless-docx]');
    await expect(root).toBeVisible({ timeout: 60_000 });
    await expect(root.locator('[data-page-id] canvas').first()).toBeVisible({ timeout: 60_000 });
    await expect.poll(() => page.evaluate(() => !!window.__docxPeerlessInput.initial)).toBe(true);
  } else {
    await expect(page.locator('.editor-stage canvas').first()).toBeVisible({ timeout: 60_000 });
    await expect(page.getByTestId('formatting-bar')).toBeVisible();
  }
}

function documentCanvas(page: Page, peerless = true) {
  return page.locator(peerless
    ? '[data-peerless-docx] [data-page-id] canvas'
    : '.canvas-page canvas').first();
}

async function setMode(page: Page, current: 'Editing' | 'Viewing', next: 'Editing' | 'Viewing') {
  const controls = page.getByTestId('peerless-harness-mode');
  if (await controls.isVisible()) await controls.getByRole('button', { name: next, exact: true }).click();
  else {
    await page.getByRole('button', { name: current, exact: true }).click();
    await page.getByRole('menuitemradio', { name: new RegExp(`^${next}\\b`) }).click();
  }
}

async function focusDocument(page: Page, peerless = true) {
  await documentCanvas(page, peerless).click({ position: { x: 150, y: 150 } });
  const input = peerless
    ? page.locator('.editor-stage textarea').last()
    : page.getByTestId('yrs-input');
  await expect(input).toBeEditable();
  await input.press('ControlOrMeta+Home');
  if (peerless) await settled(page);
  return input;
}

async function settled(page: Page) {
  await expect.poll(() => page.evaluate(() => {
    const probe = window.__docxPeerlessInput;
    const last = probe.commands.at(-1);
    return !last || probe.acks.some((ack) => ack.throughSeq >= last.seq);
  })).toBe(true);
  await paintTurns(page);
}

async function paintTurns(page: Page) {
  await pinnedEvaluate(page, () => new Promise<void>((resolve) => {
    requestAnimationFrame(() => requestAnimationFrame(() => resolve()));
  }));
}

async function commandsSince(page: Page, start: number) {
  return page.evaluate((start) => window.__docxPeerlessInput.commands.slice(start), start);
}

async function commandCount(page: Page) {
  return page.evaluate(() => window.__docxPeerlessInput.commands.length);
}

function expectChain(commands: DocxCommandEnvelope[], confirmedAtSend?: Record<number, CommandConfirmation>) {
  expect(commands.length).toBeGreaterThan(0);
  for (let index = 1; index < commands.length; index += 1) {
    const command = commands[index]!;
    const previous = commands[index - 1]!;
    if (command.afterSeq === undefined && confirmedAtSend) {
      expect(confirmedAtSend[command.seq]).toEqual({
        session: command.session,
        incarnation: command.incarnation,
        epoch: command.basis.epoch,
        throughSeq: previous.seq,
        version: command.basis.version,
      });
    } else expect(command.afterSeq).toBe(previous.seq);
  }
  for (const command of commands) {
    if (command.kind !== 'select') expect(command.payload).not.toHaveProperty('selection');
  }
}

async function beforeinput(input: Locator, text: string) {
  await input.evaluate((element, text) => {
    element.dispatchEvent(new InputEvent('beforeinput', {
      bubbles: true, cancelable: true, inputType: 'insertText', data: text,
    }));
  }, text);
}

async function paste(input: Locator, text: string) {
  await input.evaluate((element, text) => {
    const clipboardData = new DataTransfer();
    clipboardData.setData('text/plain', text);
    element.dispatchEvent(new ClipboardEvent('paste', {
      bubbles: true, cancelable: true, clipboardData,
    }));
  }, text);
}

async function savedXml(page: Page) {
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save', exact: true }).first().click();
  const stream = await (await downloading).createReadStream();
  if (!stream) throw new Error('Save did not produce a download');
  const chunks: Buffer[] = [];
  for await (const chunk of stream) chunks.push(Buffer.from(chunk));
  const zip = await JSZip.loadAsync(Buffer.concat(chunks));
  return zip.file('word/document.xml')!.async('string');
}

async function savedText(page: Page) {
  const xml = await savedXml(page);
  return page.evaluate((xml) => Array.from(
    new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagNameNS('*', 't')
  ).map((node) => node.textContent ?? '').join(''), xml);
}

const steadyCaret = '[data-peerless-docx] div[aria-hidden="true"][style*="currentcolor" i] { opacity: 1 !important; }';

async function pixels(page: Page, peerless = true) {
  return documentCanvas(page, peerless).screenshot({ animations: 'disabled', style: steadyCaret });
}

test('peerless Shift+Enter saves a soft break within one paragraph', async ({ page }) => {
  const zip = await JSZip.loadAsync(await documentBytes());
  zip.file('word/document.xml', (await zip.file('word/document.xml')!.async('string'))
    .replace('<w:r><w:t>abcd</w:t></w:r>', ''));
  await open(page, true, null, 1, await zip.generateAsync({ type: 'nodebuffer' }));
  const input = await focusDocument(page);
  const start = await commandCount(page);
  await input.pressSequentially('a');
  await input.press('Shift+Enter');
  await input.pressSequentially('b');
  await settled(page);
  const commands = await commandsSince(page, start);
  expect(commands.map(command => command.kind)).toEqual(['insertText', 'lineBreak', 'insertText']);
  expectChain(commands, await page.evaluate(() => window.__docxPeerlessInput.confirmedAtSend));
  const xml = await savedXml(page);
  const paragraphs = await page.evaluate(xml => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    return Array.from(document.getElementsByTagNameNS('*', 'p')).map(paragraph =>
      Array.from(paragraph.getElementsByTagNameNS('*', '*')).map(node =>
        node.localName === 't' ? node.textContent ?? '' : node.localName === 'br' ? '\n' : ''
      ).join('')
    );
  }, xml);
  expect(paragraphs).toEqual(['a\nb']);
  expect(xml).toMatch(/<w:t[^>]*>a<\/w:t>[\s\S]*<w:br(?: w:type="textWrapping")?\s*\/>[\s\S]*<w:t[^>]*>b<\/w:t>/);
});

test('d2d_pending_insert_click_arrow_backspace_order', async ({ page }) => {
  await open(page);
  const input = await focusDocument(page);
  const start = await commandCount(page);
  await page.evaluate(() => { window.__docxPeerlessInput.holdAcks = true; });
  await beforeinput(input, 'XY');
  await expect.poll(() => page.evaluate(() => window.__docxPeerlessInput.heldAcks.length)).toBeGreaterThan(0);
  await documentCanvas(page).click({ position: { x: 150, y: 150 } });
  await input.press('ControlOrMeta+End');
  await input.press('ArrowLeft');
  await input.press('Backspace');
  await page.evaluate(() => window.__docxPeerlessInput.releaseAcks());
  await expect.poll(async () => (await commandsSince(page, start)).length).toBe(5);
  await settled(page);
  const commands = await commandsSince(page, start);
  expect(commands.map((command) => command.kind)).toEqual([
    'insertText', 'selectPoint', 'navigate', 'navigate', 'deleteBackward',
  ]);
  expectChain(commands);
  expect(await savedText(page)).toBe('XYabd');
});

test('d2d_typing_paste_formatting_share_chain', async ({ page }) => {
  await open(page);
  const input = await focusDocument(page);
  const start = await commandCount(page);
  await page.evaluate(() => { window.__docxPeerlessInput.holdAcks = true; });
  await beforeinput(input, 'T');
  await paste(input, 'PASTE');
  const toolbar = page.getByTestId('formatting-bar');
  const bold = toolbar.getByRole('button', { name: 'Bold', exact: true });
  if (await bold.isVisible()) await bold.click();
  else {
    await toolbar.getByTestId('toolbar-more').click();
    await page.getByRole('menu', { name: 'More actions' }).locator('[data-label="Bold"]').click();
  }
  await beforeinput(input, 'B');
  await page.evaluate(() => window.__docxPeerlessInput.releaseAcks());
  await expect.poll(async () => (await commandsSince(page, start)).length).toBe(4);
  await settled(page);
  const commands = await commandsSince(page, start);
  expect(commands[0]!.kind).toBe('insertText');
  expect(['insertText', 'replaceSelection']).toContain(commands[1]!.kind);
  expect(commands[1]!.payload).toEqual({ text: 'PASTE' });
  expect(commands[2]!.kind).toBe('toggleMark');
  expect(commands[3]!.kind).toBe('insertText');
  expectChain(commands);
  const xml = await savedXml(page);
  expect(await page.evaluate((xml) => {
    const document = new DOMParser().parseFromString(xml, 'application/xml');
    const run = Array.from(document.getElementsByTagNameNS('*', 'r')).find((run) =>
      Array.from(run.getElementsByTagNameNS('*', 't')).some((text) => text.textContent === 'B')
    );
    return !!run?.getElementsByTagNameNS('*', 'b').length;
  }, xml)).toBe(true);
  expect(await savedText(page)).toBe('TPASTEBabcd');
});

test('typing_applied_late_undoes_in_one_step', async ({ page }) => {
  await open(page);
  const input = await focusDocument(page);
  await page.evaluate(() => { window.__docxPeerlessInput.editDelayMs = 700; });
  await input.pressSequentially('Beta');
  await settled(page);
  expect(await savedText(page)).toBe('Betaabcd');
  await input.press('ControlOrMeta+z');
  await settled(page);
  expect(await savedText(page)).toBe('abcd');
});

test('save_includes_typing_deferred_after_vertical_navigation', async ({ page }) => {
  await open(page);
  const input = await focusDocument(page);
  const start = await commandCount(page);
  await page.evaluate(() => { window.__docxPeerlessInput.holdAcks = true; });
  await beforeinput(input, 'x');
  await input.press('ArrowUp');
  await beforeinput(input, 'y');
  const saving = savedText(page);
  await paintTurns(page);
  await page.evaluate(() => window.__docxPeerlessInput.releaseAcks());
  expect(await saving).toContain('y');
  const commands = await commandsSince(page, start);
  expect(commands.map(({ kind }) => kind)).toEqual(['insertText', 'navigate', 'insertText']);
  expectChain(commands);
});

test('keyboard_and_drag_ranges_have_visible_highlights', async ({ page }) => {
  await open(page);
  const input = await focusDocument(page);
  await input.press('Shift+ArrowRight');
  await settled(page);
  const overlay = page.locator('[data-peerless-docx] div[aria-hidden="true"][style*="rgba(59, 130, 246"]');
  await expect(overlay.first()).toBeVisible();
  await input.press('ArrowRight');
  await settled(page);
  await expect(overlay).toHaveCount(0);
  const initial = applyFrameDelta(null, decodeFrameDelta(Uint8Array.from(await page.evaluate(() =>
    Array.from(new Uint8Array(window.__docxPeerlessInput.initial!.frame!.payload.frame))
  ))));
  const run = initial.pages[0]!.page.primitives.find((primitive) =>
    primitive.kind === 'glyphRun' && !primitive.listMarker && primitive.glyphs.length > 0);
  if (run?.kind !== 'glyphRun' || !run.glyphs.length) throw new Error('Document text has no rendered glyphs');
  const bounds = await documentCanvas(page).boundingBox();
  if (!bounds) throw new Error('Document canvas is unavailable');
  const first = run.glyphs[0]!;
  const last = run.glyphs.at(-1)!;
  const y = bounds.y + first.y - run.size / 2;
  await page.mouse.move(bounds.x + first.x, y);
  await page.mouse.down();
  await page.mouse.move(bounds.x + last.x + (last.advance ?? 1), y, { steps: 4 });
  await page.mouse.up();
  await settled(page);
  await expect(overlay.first()).toBeVisible();
});

test('d2d_echo_caps_preserve_all_input', async ({ page }) => {
  for (const contextBytes of [null, 16 * 1024 + 1]) {
    await open(page, true, contextBytes);
    const input = await focusDocument(page);
    const start = await commandCount(page);
    const before = await pixels(page);
    expect(await page.evaluate(() => !!window.__docxPeerlessInput.initial?.frame?.payload.echoContext)).toBe(true);
    await page.evaluate(() => { window.__docxPeerlessInput.holdAcks = true; });
    await beforeinput(input, 'x');
    await paintTurns(page);
    if (contextBytes === null) expect(await pixels(page)).not.toEqual(before);
    else expect(await pixels(page)).toEqual(before);
    await input.evaluate((element) => {
      for (let index = 1; index < 257; index += 1) {
        element.dispatchEvent(new InputEvent('beforeinput', {
          bubbles: true, cancelable: true, inputType: 'insertText', data: 'x',
        }));
      }
    });
    await paintTurns(page);
    const capped = await pixels(page);
    await input.evaluate((element) => {
      for (let index = 257; index < 300; index += 1) {
        element.dispatchEvent(new InputEvent('beforeinput', {
          bubbles: true, cancelable: true, inputType: 'insertText', data: 'x',
        }));
      }
    });
    await paintTurns(page);
    expect(await pixels(page)).toEqual(capped);
    if (contextBytes !== null) expect(capped).toEqual(before);
    await page.evaluate(() => window.__docxPeerlessInput.releaseAcks());
    await expect.poll(async () => (await commandsSince(page, start)).reduce((count, command) => {
      const payload = command.payload as { text?: string };
      return count + (command.kind === 'insertText' ? payload.text?.length ?? 0 : 0);
    }, 0)).toBe(300);
    await settled(page);
    expectChain(await commandsSince(page, start));
    expect(await savedText(page)).toBe('x'.repeat(300) + 'abcd');
  }
});

test('d2d_refusal_restores_pixels_and_selection', async ({ page }) => {
  await open(page);
  const input = await focusDocument(page);
  const before = await pixels(page);
  const selection = await page.evaluate(() => {
    const probe = window.__docxPeerlessInput;
    return [...probe.acks].reverse().find((ack) => ack.selection !== undefined)?.selection ?? probe.initial?.selection;
  });
  await page.evaluate(() => {
    window.__docxPeerlessInput.holdAcks = true;
    window.__docxPeerlessInput.refuseNext = true;
  });
  await beforeinput(input, 'REFUSED');
  await expect.poll(() => page.evaluate(() => window.__docxPeerlessInput.heldAcks.length)).toBeGreaterThan(0);
  await paintTurns(page);
  expect(await pixels(page)).not.toEqual(before);
  await page.evaluate(() => window.__docxPeerlessInput.releaseAcks());
  await settled(page);
  expect(await pixels(page)).toEqual(before);
  expect(await page.evaluate(() => window.__docxPeerlessInput.acks.at(-1)?.selection)).toEqual(selection);
  expect(await page.evaluate(() => window.__docxPeerlessInput.acks.at(-1)?.outcomes.at(-1)?.status)).toBe('refused');
  await beforeinput(input, 'OK');
  await settled(page);
  expect(await savedText(page)).toBe('OKabcd');
});

test('d2d_in_epoch_refusal_covers_sequence_and_accepts_later_input', async ({ page }) => {
  await open(page);
  const input = await focusDocument(page);
  const start = await commandCount(page);
  const baseline = await page.evaluate(() => {
    const probe = window.__docxPeerlessInput;
    return (probe.acks.at(-1) ?? probe.initial)!;
  });
  await page.evaluate(() => { window.__docxPeerlessInput.refuseNext = true; });
  await beforeinput(input, 'REFUSED');
  await expect.poll(async () => (await commandsSince(page, start)).length).toBe(1);
  await settled(page);
  const refused = (await commandsSince(page, start))[0]!;
  expect(refused).toMatchObject({
    basis: { epoch: baseline.epoch, version: baseline.version + 1 },
    expectVersion: true,
  });
  const refusal = await page.evaluate((seq) => window.__docxPeerlessInput.acks.find((ack) =>
    ack.outcomes.some((outcome) => outcome.seq === seq)
  ), refused.seq);
  expect(refusal).toMatchObject({
    session: baseline.session,
    incarnation: baseline.incarnation,
    epoch: baseline.epoch,
    version: baseline.version,
    outcomes: [{ seq: refused.seq, status: 'refused', reason: 'stale-version', committed: false }],
  });
  expect(refusal!.throughSeq).toBeGreaterThanOrEqual(refused.seq);
  await beforeinput(input, 'OK');
  await expect.poll(async () => (await commandsSince(page, start)).length).toBe(2);
  await settled(page);
  const accepted = (await commandsSince(page, start))[1]!;
  const ack = await page.evaluate((seq) => window.__docxPeerlessInput.acks.find((ack) =>
    ack.outcomes.some((outcome) => outcome.seq === seq)
  ), accepted.seq);
  expect(ack).toMatchObject({
    epoch: baseline.epoch,
    version: baseline.version + 1,
    outcomes: [{ seq: accepted.seq, status: 'applied' }],
  });
  expect(ack!.throughSeq).toBeGreaterThanOrEqual(accepted.seq);
  expect(await savedText(page)).toBe('OKabcd');
});

test('handoff_installs_A_then_redraws_pending_B', async ({ page }) => {
  test.skip(process.env.DOCX_PEERLESS_HARNESS !== '1', 'Presentation checkpoints require DOCX_PEERLESS_HARNESS=1');
  await openHarness(page);
  const input = await focusHarness(page);
  const canvas = page.locator('[data-peerless-docx] [data-page-id] canvas').first();
  const before = await canvas.evaluate((element) => (element as HTMLCanvasElement).toDataURL());
  await page.evaluate(() => { window.__docxPeerlessInput.holdFrames = true; });
  await beforeinput(input, 'A');
  await expect.poll(() => page.evaluate(() => window.__docxPeerlessInput.heldFrames.length)).toBeGreaterThan(0);
  await paintTurns(page);
  const echo = await canvas.evaluate((element) => (element as HTMLCanvasElement).toDataURL());
  expect(echo).not.toEqual(before);
  await paintTurns(page);
  expect(await canvas.evaluate((element) => (element as HTMLCanvasElement).toDataURL())).toEqual(echo);
  await page.evaluate(() => { window.__docxPeerlessInput.holdAcks = true; });
  await beforeinput(input, 'B');
  await expect.poll(() => page.evaluate(() => window.__docxPeerlessInput.heldAcks.length)).toBeGreaterThan(0);
  await paintTurns(page);
  const laterEcho = await canvas.evaluate((element) => (element as HTMLCanvasElement).toDataURL());
  expect(laterEcho).not.toEqual(echo);
  const [a, b] = await page.evaluate(() => window.__docxPeerlessInput.commands.filter((command) =>
    command.kind === 'insertText'
  ).slice(-2));
  expect(a!.payload).toEqual({ text: 'A' });
  expect(b!.payload).toEqual({ text: 'B' });
  const authoritativeA = await page.evaluate((seq) => {
    const frame = window.__docxPeerlessInput.heldFrames.map(({ message }) =>
      message.ok ? message.commandFrame?.frame : undefined
    ).find((frame) => frame?.coversThroughSeq === seq)!;
    return frame.payload.stagedPresentation!.pages.map(({ bitmap }) => {
      const canvas = document.createElement('canvas');
      canvas.width = bitmap.width;
      canvas.height = bitmap.height;
      canvas.getContext('2d')!.drawImage(bitmap, 0, 0);
      return canvas.toDataURL();
    });
  }, a!.seq);
  expect(authoritativeA.length).toBeGreaterThan(0);
  await page.evaluate((seq) => {
    window.__docxPeerlessHarness!.armPresentation(seq);
    window.__docxPeerlessInput.releaseFrames();
  }, a!.seq);
  await expect.poll(() => page.evaluate(() => window.__docxPeerlessHarness!.checkpoints.length)).toBe(4);
  const handoff = await page.evaluate(() => window.__docxPeerlessHarness!.checkpoints);
  expect(handoff.map((checkpoint) => checkpoint.stage)).toEqual(['installed', 'reconciled', 'microtask', 'raf']);
  expect(new Set(handoff.map((checkpoint) => checkpoint.transaction)).size).toBe(1);
  for (const checkpoint of handoff) {
    expect(checkpoint.coversThroughSeq).toBe(a!.seq);
    expect(checkpoint.retired).toEqual([a!.seq]);
    expect(checkpoint.remaining).toEqual([b!.seq]);
    expect(checkpoint.revision).toBe(handoff[0]!.revision);
    expect(checkpoint.pixels.length).toBeGreaterThan(0);
  }
  expect(handoff[0]!.pixels).toEqual(authoritativeA);
  expect(handoff[0]!.pixels).not.toEqual(handoff[1]!.pixels);
  expect(handoff[2]!.pixels).toEqual(handoff[1]!.pixels);
  expect(handoff[3]!.pixels).toEqual(handoff[1]!.pixels);
  await paintTurns(page);
  const covered = await canvas.evaluate((element) => (element as HTMLCanvasElement).toDataURL());
  expect(covered).not.toEqual(before);
  expect(covered).not.toEqual(echo);
  await paintTurns(page);
  expect(await canvas.evaluate((element) => (element as HTMLCanvasElement).toDataURL())).toEqual(covered);
  await page.evaluate((seq) => {
    window.__docxPeerlessHarness!.armPresentation(seq);
    window.__docxPeerlessInput.releaseAcks();
  }, b!.seq);
  await expect.poll(() => page.evaluate(() => window.__docxPeerlessHarness!.checkpoints.length)).toBe(4);
  const handoffB = await page.evaluate(() => window.__docxPeerlessHarness!.checkpoints);
  expect(handoffB.map((checkpoint) => checkpoint.stage)).toEqual(['installed', 'reconciled', 'microtask', 'raf']);
  for (const checkpoint of handoffB) {
    expect(checkpoint.coversThroughSeq).toBe(b!.seq);
    expect(checkpoint.retired).toEqual([b!.seq]);
    expect(checkpoint.remaining).toEqual([]);
    expect(checkpoint.pixels).toEqual(handoffB[0]!.pixels);
  }
  await settled(page);
  const finalPixels = await canvas.evaluate((element) => (element as HTMLCanvasElement).toDataURL());
  expect(finalPixels).not.toEqual(before);
  expect(finalPixels).not.toEqual(echo);
  expect(await harnessText(page)).toBe('ABabcd');
  await expect(canvas).toBeVisible();
  expect(await canvas.evaluate((element) => {
    const root = element.closest('[data-peerless-docx]');
    const stage = element.closest('.editor-stage');
    const caret = root?.querySelector<HTMLElement>('div[aria-hidden="true"][style*="currentcolor" i]');
    if (!root || !stage || !caret) return false;
    const rect = caret.getBoundingClientRect();
    const x = rect.left + rect.width / 2;
    const y = rect.top + rect.height / 2;
    const containsPoint = (bounds: DOMRect) => x >= bounds.left && x < bounds.right &&
      y >= bounds.top && y < bounds.bottom;
    const hit = document.elementFromPoint(x, y);
    return rect.width > 0 && rect.height > 0 && containsPoint(element.getBoundingClientRect()) &&
      containsPoint(stage.getBoundingClientRect()) && !!hit && root.contains(hit) &&
      (hit === element || caret.contains(hit));
  }), 'The caret must be on-screen within the page canvas and editor stage').toBe(true);
});

const ECHO_WRAPPED = 'Wrapped' + ' echo words'.repeat(12);
const ECHO_PARAGRAPHS = Array.from({ length: 120 }, (_, index) => index === 1 ? ECHO_WRAPPED : `Seed paragraph ${index}`);

async function echoDocumentBytes() {
  const zip = await JSZip.loadAsync(await documentBytes());
  const xml = await zip.file('word/document.xml')!.async('string');
  const body = ECHO_PARAGRAPHS.map((text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`).join('');
  zip.file('word/document.xml', xml.replace(/<w:body>.*?<w:sectPr>/, `<w:body>${body}<w:sectPr>`));
  return zip.generateAsync({ type: 'nodebuffer' });
}

test('echo paints the caret line of a wrapped paragraph before the reply and reconciles to the reply', async ({ page }) => {
  test.skip(process.env.DOCX_PEERLESS_HARNESS !== '1', 'Presentation checkpoints require DOCX_PEERLESS_HARNESS=1');
  await openHarness(page, 1, await echoDocumentBytes());
  const input = await focusHarness(page);
  for (const key of ['ArrowDown', 'ArrowDown', 'ArrowRight', 'ArrowRight', 'ArrowRight']) await input.press(key);
  await settled(page);
  const context = await page.evaluate(() => window.__docxPeerlessInput.frames.at(-1)?.payload.echoContext);
  expect(context, 'the caret line must carry an echo context').toBeDefined();
  expect(ECHO_WRAPPED.endsWith(context!.text) && context!.text.length < ECHO_WRAPPED.length).toBe(true);
  const canvas = page.locator('[data-peerless-docx] [data-page-id] canvas').first();
  const canvasPixels = () => canvas.evaluate((element) => (element as HTMLCanvasElement).toDataURL());
  const key = async (press: () => Promise<void>) => {
    const before = await canvasPixels();
    const start = await page.evaluate(() => {
      window.__docxPeerlessInput.holdAcks = true;
      window.__docxPeerlessInput.holdFrames = true;
      return window.__docxPeerlessInput.commands.length;
    });
    await press();
    await paintTurns(page);
    const early = await canvasPixels();
    await expect.poll(() => page.evaluate((start) => {
      const probe = window.__docxPeerlessInput;
      const seq = probe.commands[start]?.seq;
      return seq !== undefined && probe.heldAcks.some(({ message }) => message.ok && (message.commandAck?.throughSeq ?? -1) >= seq &&
        !!message.commandAck?.frame);
    }, start)).toBe(true);
    const authoritative = await page.evaluate((start) => {
      const probe = window.__docxPeerlessInput;
      const seq = probe.commands[start]!.seq;
      const frame = probe.heldAcks.map(({ message }) => message.ok ? message.commandAck : undefined)
        .find((ack) => ack?.frame && ack.throughSeq >= seq)!.frame!;
      const item = frame.payload.stagedPresentation!.pages.find(({ page }) => page.pageIndex === 0)!;
      const copy = document.createElement('canvas');
      copy.width = item.bitmap.width;
      copy.height = item.bitmap.height;
      copy.getContext('2d')!.drawImage(item.bitmap, 0, 0);
      return copy.toDataURL();
    }, start);
    expect(await canvasPixels(), 'nothing but the echo paints while the reply is held').toEqual(early);
    await page.evaluate(() => {
      window.__docxPeerlessInput.releaseAcks();
      window.__docxPeerlessInput.releaseFrames();
    });
    await settled(page);
    const after = await canvasPixels();
    expect(after, 'the reply replaces the echo with its own pixels').toEqual(authoritative);
    expect(after).not.toEqual(before);
    return { before, early };
  };
  const typed = await key(() => beforeinput(input, 'x'));
  expect(typed.early, 'the typed glyph paints before the reply').not.toEqual(typed.before);
  const deleted = await key(() => input.press('Backspace'));
  expect(deleted.early, 'the backspace paints before the reply').not.toEqual(deleted.before);
  const split = await key(() => input.press('Enter'));
  expect(split.early, 'a declined key leaves the pixels to the reply').toEqual(split.before);
  expect(await harnessText(page)).toBe(ECHO_PARAGRAPHS.join(''));
});

test('echo paints exactly the pixels of the reply for inserts and backspaces', async ({ page }) => {
  test.skip(process.env.DOCX_PEERLESS_HARNESS !== '1', 'Presentation checkpoints require DOCX_PEERLESS_HARNESS=1');
  await openHarness(page, 1, await echoDocumentBytes());
  const input = await focusHarness(page);
  const canvas = page.locator('[data-peerless-docx] [data-page-id] canvas').first();
  const key = async (label: string, press: () => Promise<void>) => {
    await settled(page);
    expect(await page.evaluate(() => window.__docxPeerlessInput.frames.at(-1)?.payload.echoContext),
      `${label}: the caret line must carry an echo context`).toBeDefined();
    const before = await canvas.evaluate((element) => (element as HTMLCanvasElement).toDataURL());
    const start = await page.evaluate(() => {
      window.__docxPeerlessInput.holdAcks = true;
      window.__docxPeerlessInput.holdFrames = true;
      return window.__docxPeerlessInput.commands.length;
    });
    await press();
    await paintTurns(page);
    await expect.poll(() => page.evaluate((start) => {
      const probe = window.__docxPeerlessInput;
      const seq = probe.commands[start]?.seq;
      return seq !== undefined && probe.heldAcks.some(({ message }) => message.ok &&
        (message.commandAck?.throughSeq ?? -1) >= seq && !!message.commandAck?.frame);
    }, start)).toBe(true);
    const comparison = await canvas.evaluate((element, start) => {
      const probe = window.__docxPeerlessInput;
      const seq = probe.commands[start]!.seq;
      const frame = probe.heldAcks.map(({ message }) => message.ok ? message.commandAck : undefined)
        .find((ack) => ack?.frame && ack.throughSeq >= seq)!.frame!;
      const item = frame.payload.stagedPresentation!.pages.find(({ page }) => page.pageIndex === 0)!;
      const target = element as HTMLCanvasElement;
      const reply = document.createElement('canvas');
      reply.width = item.bitmap.width;
      reply.height = item.bitmap.height;
      const replyContext = reply.getContext('2d', { willReadFrequently: true })!;
      replyContext.drawImage(item.bitmap, 0, 0);
      if (reply.width !== target.width || reply.height !== target.height) return { size: false, differing: -1, max: 0, bbox: null };
      const echo = target.getContext('2d')!.getImageData(0, 0, target.width, target.height).data;
      const expected = replyContext.getImageData(0, 0, reply.width, reply.height).data;
      let differing = 0;
      let max = 0;
      let bbox: number[] | null = null;
      for (let index = 0; index < echo.length; index += 4) {
        let delta = 0;
        for (let channel = 0; channel < 4; channel += 1) delta = Math.max(delta, Math.abs(echo[index + channel]! - expected[index + channel]!));
        if (!delta) continue;
        differing += 1;
        max = Math.max(max, delta);
        const x = (index / 4) % target.width;
        const y = Math.floor(index / 4 / target.width);
        bbox = bbox ? [Math.min(bbox[0]!, x), Math.min(bbox[1]!, y), Math.max(bbox[2]!, x), Math.max(bbox[3]!, y)] : [x, y, x, y];
      }
      return { size: true, differing, max, bbox };
    }, start);
    const early = await canvas.evaluate((element) => (element as HTMLCanvasElement).toDataURL());
    expect(early, `${label}: the echo paints before the reply`).not.toEqual(before);
    expect(comparison, `${label}: the echo paints exactly the reply's pixels`).toEqual({ size: true, differing: 0, max: 0, bbox: null });
    await page.evaluate(() => {
      window.__docxPeerlessInput.releaseAcks();
      window.__docxPeerlessInput.releaseFrames();
    });
    await settled(page);
  };
  for (const press of ['ArrowDown', 'ArrowDown', ...Array<string>(8).fill('ArrowRight')]) await input.press(press);
  await key('wrapped insert', () => beforeinput(input, 'x'));
  await key('wrapped backspace', () => input.press('Backspace'));
  await input.press('ControlOrMeta+Home');
  for (const press of ['ArrowDown', 'ArrowDown', 'ArrowDown', 'ArrowDown', 'End', 'ArrowLeft', 'ArrowLeft', 'ArrowLeft', 'ArrowLeft']) await input.press(press);
  await key('single-line insert', () => beforeinput(input, 'x'));
  await key('single-line backspace', () => input.press('Backspace'));
  expect(await harnessText(page)).toBe(ECHO_PARAGRAPHS.join(''));
});

test('viewing_refuses_typing_then_editing_applies', async ({ page }) => {
  await open(page);
  const input = await focusDocument(page);
  const start = await commandCount(page);
  await page.evaluate(() => { window.__docxPeerlessInput.holdAcks = true; });
  await setMode(page, 'Editing', 'Viewing');
  await expect.poll(async () => (await commandsSince(page, start)).some((command) =>
    command.kind === 'setReadOnly' && (command.payload as { readOnly: boolean }).readOnly
  )).toBe(true);
  await beforeinput(input, 'REFUSED');
  expect((await commandsSince(page, start)).some((command) => command.kind === 'insertText')).toBe(false);
  if (!await page.getByTestId('peerless-harness-mode').isVisible()) {
    await page.evaluate(() => window.__docxPeerlessInput.releaseAcks());
  }
  await setMode(page, 'Viewing', 'Editing');
  await expect(input).toBeEditable();
  await beforeinput(input, 'OK');
  await page.evaluate(() => window.__docxPeerlessInput.releaseAcks());
  await settled(page);
  const commands = await commandsSince(page, start);
  const accepted = commands.find((command) => command.kind === 'insertText' &&
    (command.payload as { text: string }).text === 'OK');
  expect(accepted).toBeDefined();
  expect(await page.evaluate((seq) => window.__docxPeerlessInput.acks.flatMap((ack) => ack.outcomes)
    .find((outcome) => outcome.seq === seq)?.status, accepted!.seq)).toBe('applied');
  expect(await savedText(page)).toBe('OKabcd');
});

test('evicted_page_repaints_after_scroll_back', async ({ page }) => {
  await open(page, true, null, 40);
  const input = await focusDocument(page);
  const first = page.locator('[data-peerless-docx] [data-page-id] canvas').first();
  const original = await first.evaluate((canvas) => {
    window.__docxEvictedCanvas = canvas as HTMLCanvasElement;
    return { pageId: canvas.parentElement!.dataset.pageId!, pixels: (canvas as HTMLCanvasElement).toDataURL() };
  });
  const pageCanvas = page.locator(`[data-peerless-docx] [data-page-id="${original.pageId}"] canvas`);
  const ink = async () => pageCanvas.evaluate((element) => {
    const canvas = element as HTMLCanvasElement;
    const data = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
    let count = 0;
    for (let index = 0; index < data.length; index += 4) {
      if (data[index + 3]! > 0 && Math.min(data[index]!, data[index + 1]!, data[index + 2]!) < 200) count += 1;
    }
    return count;
  });
  expect(await ink()).toBeGreaterThan(20);
  const scroller = page.locator('.editor-stage .docx-editor__scroll-container');
  await expect(scroller).toHaveCount(1);
  await input.press('ControlOrMeta+End');
  await settled(page);
  await scroller.evaluate((element) => { element.scrollTop = element.scrollHeight; });
  await expect.poll(() => scroller.evaluate((element) => element.scrollTop)).toBeGreaterThan(0);
  await paintTurns(page);
  await expect.poll(() => page.evaluate(() => {
    const canvas = window.__docxEvictedCanvas!;
    return !canvas.isConnected && canvas.width === 0 && canvas.height === 0;
  }), { message: 'Scrolling must evict the original page canvas' }).toBe(true);
  await scroller.evaluate((element) => { element.scrollTop = 0; });
  await expect.poll(() => scroller.evaluate((element) => element.scrollTop)).toBe(0);
  await paintTurns(page);
  await expect(pageCanvas).toBeVisible();
  expect(await pageCanvas.evaluate((canvas) => canvas !== window.__docxEvictedCanvas)).toBe(true);
  await expect.poll(ink, { message: 'The evicted page must regain its text pixels' }).toBeGreaterThan(20);
  await paintTurns(page);
  expect(await pageCanvas.evaluate((canvas) => (canvas as HTMLCanvasElement).toDataURL())).toBe(original.pixels);
});

test('peerless_document_scrolls_with_wheel_and_page_down', async ({ page }) => {
  await open(page, true, null, 40);
  const input = await focusDocument(page);
  const scroller = page.locator('.editor-stage .docx-editor__scroll-container');
  await expect(scroller).toHaveCount(1);
  expect(await scroller.evaluate((element) => getComputedStyle(element).overflowY)).toBe('auto');
  const later = await page.locator('[data-peerless-docx] [data-page-id] canvas').nth(1).evaluate((canvas) => {
    const container = canvas.closest('.docx-editor__scroll-container')!;
    return { pageId: canvas.parentElement!.dataset.pageId!,
      top: canvas.getBoundingClientRect().top - container.getBoundingClientRect().top + container.scrollTop };
  });
  const laterCanvas = page.locator(`[data-peerless-docx] [data-page-id="${later.pageId}"] canvas`);
  await expect(laterCanvas).not.toBeInViewport();
  const before = await scroller.evaluate((element) => element.scrollTop);
  await scroller.hover({ position: { x: 150, y: 150 } });
  await page.mouse.wheel(0, later.top - before);
  await expect.poll(() => scroller.evaluate((element) => element.scrollTop)).toBeGreaterThan(before);
  await expect.poll(() => scroller.evaluate((element) => element.scrollTop)).toBeGreaterThanOrEqual(later.top - 1);
  await paintTurns(page);
  await expect(laterCanvas).toBeInViewport();
  await laterCanvas.click({ position: { x: 150, y: 150 } });
  await settled(page);
  await expect(input).toBeFocused();
  const afterWheel = await scroller.evaluate((element) => element.scrollTop);
  await page.keyboard.press('PageDown');
  await expect.poll(() => scroller.evaluate((element) => element.scrollTop)).toBeGreaterThan(afterWheel);
  await paintTurns(page);
  await expect(input).toBeFocused();
  await expect(laterCanvas).toBeInViewport();
  const afterPageDown = await scroller.evaluate((element) => element.scrollTop);
  await page.keyboard.press('PageUp');
  await expect.poll(() => scroller.evaluate((element) => element.scrollTop)).toBeLessThan(afterPageDown);
  await paintTurns(page);
  await expect(input).toBeFocused();
  await expect(laterCanvas).toBeInViewport();
});

test('d2d_readonly_navigation_only', async ({ page }) => {
  await open(page);
  const input = await focusDocument(page);
  await setMode(page, 'Editing', 'Viewing');
  await expect.poll(() => page.evaluate(() => window.__docxPeerlessInput.acks.at(-1)?.editState?.readOnly)).toBe(true);
  await expect(input).not.toBeEditable();
  await settled(page);
  const start = await commandCount(page);
  await input.focus();
  await input.press('ArrowRight');
  await input.press('Shift+ArrowRight');
  await beforeinput(input, 'blocked');
  await input.press('Backspace');
  await paste(input, 'blocked');
  await input.press('ControlOrMeta+b');
  await settled(page);
  const commands = await commandsSince(page, start);
  expect(commands.map((command) => command.kind)).toEqual(['navigate', 'navigate']);
  expect(await savedText(page)).toBe('abcd');
});

test('d2d_peerless_mount_creates_no_main_document', async ({ page }) => {
  const workers: string[] = [];
  page.on('worker', (worker) => workers.push(worker.url()));
  await open(page);
  const input = await focusDocument(page);
  await beforeinput(input, 'worker');
  await settled(page);
  expect(await savedText(page)).toBe('workerabcd');
  expect(await page.evaluate(() => window.__docxPeerlessInput.mainInstances)).toBe(0);
  expect(workers.some((url) => url.includes('residentEngineWorker'))).toBe(true);
  expect(await page.evaluate(() => window.__docxPeerlessInput.commands.some((command) => command.kind === 'insertText'))).toBe(true);
});

test('d2d_legacy_input_regression', async ({ page }) => {
  await open(page, false);
  const input = await focusDocument(page, false);
  await input.pressSequentially('legacy');
  await input.press('ArrowLeft');
  await input.press('Backspace');
  await paste(input, 'P');
  expect(await savedText(page)).toBe('legaPyabcd');
  expect(await page.evaluate(() => window.__docxPeerlessInput.initial)).toBeNull();
  expect(await commandCount(page)).toBe(0);
});

test('d2d_composition_event_orders_commit_once', async ({ page }) => {
  await open(page);
  const input = await focusDocument(page);
  for (const order of ['end-beforeinput-input', 'beforeinput-end-input', 'input-end-beforeinput']) {
    const start = await commandCount(page);
    await page.evaluate(() => { window.__docxPeerlessInput.holdAcks = true; });
    await beforeinput(input, 'P');
    await input.evaluate((element) => {
      const textarea = element as HTMLTextAreaElement;
      textarea.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }));
      for (const text of ['か', 'かな']) {
        textarea.value = text;
        textarea.dispatchEvent(new CompositionEvent('compositionupdate', { bubbles: true, data: text }));
        textarea.dispatchEvent(new InputEvent('beforeinput', {
          bubbles: true, cancelable: true, inputType: 'insertCompositionText', data: text, isComposing: true,
        }));
        textarea.dispatchEvent(new InputEvent('input', {
          bubbles: true, inputType: 'insertCompositionText', data: text, isComposing: true,
        }));
      }
    });
    expect((await commandsSince(page, start)).map((command) => command.kind)).toEqual(['insertText']);
    await page.evaluate(() => window.__docxPeerlessInput.releaseAcks());
    await settled(page);
    await expect(input).toHaveValue('かな');
    let ended = false;
    for (const source of order.split('-')) {
      await input.evaluate((element, source) => {
        const textarea = element as HTMLTextAreaElement;
        if (source === 'end') {
          textarea.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '漢字' }));
        } else {
          textarea.value = '漢字';
          textarea.dispatchEvent(new InputEvent(source, {
            bubbles: true, cancelable: source === 'beforeinput', inputType: 'insertFromComposition', data: '漢字',
          }));
        }
      }, source);
      if (source === 'end') ended = true;
      if (ended) {
        await expect.poll(async () => (await commandsSince(page, start)).length).toBe(2);
      }
      await paintTurns(page);
      expect((await commandsSince(page, start)).map((command) => command.kind)).toEqual(
        ended ? ['insertText', 'replaceSelection'] : ['insertText']
      );
    }
    await expect.poll(async () => (await commandsSince(page, start)).length).toBe(2);
    await settled(page);
    const commands = await commandsSince(page, start);
    expect(commands.map((command) => command.kind)).toEqual(['insertText', 'replaceSelection']);
    expect(commands[1]!.payload).toEqual({ text: '漢字' });
    expect(commands[1]!.afterSeq).toBe(commands[0]!.seq);
  }
  expect(await savedText(page)).toBe('P漢字'.repeat(3) + 'abcd');
});

test('package_image_pixels_survive_authoritative_edit_frame', async ({ page }) => {
  const zip = await JSZip.loadAsync(await documentBytes());
  const fixture = new URL('../../packages/docx/src/yrs/__fixtures__/seed-parity/images/word/', import.meta.url);
  zip.file('word/document.xml', readFileSync(new URL('document.xml', fixture)));
  zip.file('word/_rels/document.xml.rels', readFileSync(new URL('_rels/document.xml.rels', fixture)));
  zip.file('word/media/image1.png', Buffer.from(readFileSync(new URL('media/image1.png.base64', fixture), 'utf8'), 'base64'));
  zip.file('[Content_Types].xml', (await zip.file('[Content_Types].xml')!.async('string'))
    .replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>'));
  await open(page, true, null, 1, await zip.generateAsync({ type: 'nodebuffer' }));
  const coloredPixels = () => documentCanvas(page).evaluate((element) => {
    const canvas = element as HTMLCanvasElement;
    const data = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
    let count = 0;
    for (let index = 0; index < data.length; index += 4) {
      if (data[index + 3]! > 0 && Math.max(data[index]!, data[index + 1]!, data[index + 2]!) -
        Math.min(data[index]!, data[index + 1]!, data[index + 2]!) > 60) count += 1;
    }
    return count;
  });
  await expect.poll(coloredPixels).toBeGreaterThan(100);
  const before = await coloredPixels();
  const input = await focusDocument(page);
  await beforeinput(input, 'edited ');
  await settled(page);
  expect(await savedText(page)).toContain('edited ');
  await expect.poll(coloredPixels).toBeGreaterThan(before * 0.9);
  expect(await page.evaluate(() => window.__docxPeerlessInput.mainInstances)).toBe(0);
});

test('table_tab_and_shift_tab_preserve_typing_with_acks_withheld', async ({ page }) => {
  const zip = await JSZip.loadAsync(await documentBytes());
  const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const table = '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
    `<w:tr>${cell('A')}${cell('B')}</w:tr></w:tbl>`;
  zip.file('word/document.xml', (await zip.file('word/document.xml')!.async('string'))
    .replace('<w:r><w:t>abcd</w:t></w:r>', '').replace('</w:p><w:sectPr>', `</w:p>${table}<w:sectPr>`));
  await open(page, true, null, 1, await zip.generateAsync({ type: 'nodebuffer' }));
  const initial = applyFrameDelta(null, decodeFrameDelta(Uint8Array.from(await page.evaluate(() =>
    Array.from(new Uint8Array(window.__docxPeerlessInput.initial!.frame!.payload.frame))
  ))));
  const run = initial.pages[0]!.page.primitives.find((primitive) =>
    primitive.kind === 'glyphRun' && primitive.cell?.row === 0 && primitive.cell.col === 0
  );
  if (run?.kind !== 'glyphRun' || !run.glyphs[0]) throw new Error('First cell has no rendered glyph');
  const glyph = run.glyphs[0];
  await documentCanvas(page).click({ position: { x: glyph.x + (glyph.advance ?? 1) / 2, y: glyph.y - run.size / 2 } });
  await settled(page);
  await expect.poll(() => page.evaluate(() => window.__docxPeerlessInput.acks.at(-1)?.editState?.tableContext))
    .toMatchObject({ rowIndex: 0, columnIndex: 0 });
  const input = page.locator('[data-peerless-docx] textarea');
  await input.press('Home');
  await settled(page);
  const start = await commandCount(page);
  await page.evaluate(() => { window.__docxPeerlessInput.holdAcks = true; });
  await beforeinput(input, 'X');
  await input.press('Tab');
  await beforeinput(input, 'Y');
  await input.press('Shift+Tab');
  await beforeinput(input, 'Z');
  await expect.poll(async () => (await commandsSince(page, start)).length).toBe(4);
  expect((await commandsSince(page, start)).map(({ kind }) => kind))
    .toEqual(['insertText', 'navigateCell', 'insertText', 'navigateCell']);
  await page.evaluate(() => window.__docxPeerlessInput.releaseAcks());
  await expect.poll(async () => (await commandsSince(page, start)).length).toBe(5);
  const commands = await commandsSince(page, start);
  expect(commands.map(({ kind }) => kind)).toEqual(['insertText', 'navigateCell', 'insertText', 'navigateCell', 'insertText']);
  expectChain(commands);
  await settled(page);
  expect(await savedText(page)).toBe('ZXAYB');
  expect(await page.evaluate(() => window.__docxPeerlessInput.mainInstances)).toBe(0);
});

async function openViewportFirst(page: Page, feature: string) {
  await installProbe(page);
  await page.goto('/?format=docx&peerless=1');
  await page.evaluate(() => { window.__docxPeerlessInput.holdFrames = true; });
  await page.locator('input[type=file]').last().setInputFiles({
    name: 'generated-pages.docx',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    buffer: await documentBytes(80, feature),
  });
  const canvas = documentCanvas(page);
  await expect(canvas).toBeVisible();
  await expect.poll(() => page.evaluate(() => !!window.__docxPeerlessInput.initial)).toBe(true);
  return canvas;
}

async function releaseCertifiedGeometry(page: Page) {
  const scroller = page.locator('.editor-stage .docx-editor__scroll-container');
  const prefixHeight = await scroller.evaluate((element) => element.scrollHeight);
  await expect.poll(() => page.evaluate(() => window.__docxPeerlessInput.heldFrames.some(({ message }) =>
    message.ok && message.commandFrame?.frame.payload.layoutProvisional !== true))).toBe(true);
  await page.evaluate(() => window.__docxPeerlessInput.releaseFrames());
  await expect.poll(() => scroller.evaluate((element) => element.scrollHeight)).toBeGreaterThan(prefixHeight);
}

for (const feature of ['plain', 'table', 'keep', 'section', 'PAGE']) {
test(`viewport-first ${feature} open preserves pixels and page position as certified geometry grows`, async ({ page }) => {
  const canvas = await openViewportFirst(page, feature);
  const initial = await page.evaluate(() => {
    const payload = window.__docxPeerlessInput.initial!.frame!.payload;
    return { provisional: payload.layoutProvisional, frame: Array.from(new Uint8Array(payload.frame)) };
  });
  expect(initial.provisional).toBe(true);
  expect(applyFrameDelta(null, decodeFrameDelta(Uint8Array.from(initial.frame))).pages).toHaveLength(3);
  const before = await canvas.evaluate((element) => {
    const rect = element.getBoundingClientRect();
    return { pixels: (element as HTMLCanvasElement).toDataURL(), x: rect.x, y: rect.y };
  });
  await releaseCertifiedGeometry(page);
  const after = await canvas.evaluate((element) => {
    const rect = element.getBoundingClientRect();
    return { pixels: (element as HTMLCanvasElement).toDataURL(), x: rect.x, y: rect.y };
  });
  expect(after).toEqual(before);
});

test(`viewport-first ${feature} open paints later pages after certified geometry grows`, async ({ page }) => {
  test.skip(process.env.DOCX_PEERLESS_HARNESS !== '1', 'Viewport probes require DOCX_PEERLESS_HARNESS=1');
  await openViewportFirst(page, feature);
  await releaseCertifiedGeometry(page);
  await expect.poll(() => page.evaluate(() => window.__docxPeerlessHarness?.capabilities)).toMatchObject({ presentation: true });
  await page.evaluate(() => window.__docxPeerlessHarness!.registerWorker(window.__docxPeerlessInput.worker!));
  await pinnedEvaluate(page, async () => { await window.__docxPeerlessHarness!.setViewport!([40, 41, 42]); });
  const later = page.locator('[data-peerless-docx] [data-page-index="40"] canvas');
  await expect(later).toBeAttached();
  await expect.poll(() => later.evaluate((element) => (element as HTMLCanvasElement).width)).toBeGreaterThan(0);
});
}
