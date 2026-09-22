import {
  type Latest,
  type PageScore,
  type Report,
  type Sample,
  hasRender,
  parseLatest,
  parseReport,
  rank,
  referenceUrl,
  renderUrl,
  reportUrl,
  score,
} from './report';
import { sameOrigin } from './origin';

const find = <T extends HTMLElement>(id: string): T => window.document.getElementById(id) as T;

const aside = window.document.querySelector('aside') as HTMLElement;
const main = window.document.querySelector('main') as HTMLElement;
const loader = find('loader');
const loaderMessage = find('loader-message');
const list = find<HTMLOListElement>('documents');
const filter = find<HTMLInputElement>('filter');
const revision = find('revision');
const title = find('title');
const scores = find('scores');
const note = find<HTMLParagraphElement>('note');
const strip = find('pages');
const stage = find('stage');
const sheet = find('sheet');
const handle = find('handle');
const reference = find<HTMLImageElement>('reference');
const render = find<HTMLImageElement>('render');

let report: Report | null = null;
let latest: Latest | null = null;
let current: Sample | null = null;
let page = 1;
let zoom: 'fit' | 'full' = 'fit';

function currentPage(): PageScore | null {
  return current?.pages.find((entry) => entry.page === page) ?? null;
}

const buttons = new Map<string, HTMLButtonElement>();

function drawList(): void {
  if (!report) return;
  const needle = filter.value.trim().toLowerCase();
  buttons.clear();
  list.replaceChildren(
    ...rank(report.documents)
      .filter((sample) => !needle || sample.id.toLowerCase().includes(needle))
      .map((sample) => {
        const button = window.document.createElement('button');
        button.type = 'button';
        button.append(
          span('format', sample.format),
          span('id', sample.id),
          span('num', score(sample.commit))
        );
        button.addEventListener('click', () => select(sample));
        buttons.set(sample.id, button);
        const item = window.document.createElement('li');
        item.append(button);
        return item;
      })
  );
  markCurrent();
}

function markCurrent(): void {
  for (const [id, button] of buttons) button.ariaCurrent = String(id === current?.id);
}

function span(className: string, text: string): HTMLSpanElement {
  const element = window.document.createElement('span');
  element.className = className;
  element.textContent = text;
  return element;
}

function drawStrip(): void {
  if (!current) return;
  strip.replaceChildren(
    ...current.pages.map((entry) => {
      const button = window.document.createElement('button');
      button.type = 'button';
      button.className = 'num';
      button.textContent = String(entry.page);
      button.ariaCurrent = String(entry.page === page);
      button.dataset.missing = String(!hasRender(latest, current!.id, entry.page));
      button.title =
        entry.ssim === null ? 'No common page' : `Page ${entry.page}: SSIM ${entry.ssim.toFixed(4)}`;
      button.addEventListener('click', () => {
        page = entry.page;
        drawStrip();
        drawStage();
        remember();
      });
      return button;
    })
  );
}

function drawHeader(): void {
  if (!report || !current) return;
  title.textContent = current.id;
  const entry = currentPage();
  const commit = current.commit;
  const parts = [
    ['Published', score(current.published)],
    ['Commit', score(current.commit)],
    ['Page', entry?.ssim === null || entry === null ? '—' : entry.ssim.toFixed(4)],
    ['Pages', commit?.status === 'ok' ? `${commit.referencePages} / ${commit.actualPages}` : '—'],
  ];
  scores.replaceChildren(
    ...parts.map(([label, value]) => {
      const element = window.document.createElement('span');
      element.append(`${label} `, bold(value!));
      return element;
    })
  );
}

function bold(text: string): HTMLElement {
  const element = window.document.createElement('b');
  element.className = 'num';
  element.textContent = text;
  return element;
}

function drawStage(): void {
  if (!report || !current) return;
  drawHeader();
  const entry = currentPage();
  const commit = current.commit;
  const messages: string[] = [];

  if (commit?.status === 'failed') messages.push(`${commit.stage} failed: ${commit.error}`);
  const hasReference = commit?.status !== 'ok' || page <= commit.referencePages;
  const hasOurs = hasRender(latest, current.id, page);

  reference.dataset.absent = String(!hasReference);
  if (hasReference) reference.src = referenceUrl(current.id, page);
  else reference.removeAttribute('src');

  const ours = render.parentElement as HTMLElement;
  ours.dataset.absent = String(!hasOurs);
  if (hasOurs && latest) render.src = renderUrl(latest.sha, current.id, page);
  else render.removeAttribute('src');

  if (!hasReference) messages.push('Office rendered fewer pages: this page exists only in our render.');
  if (!hasOurs)
    messages.push(
      latest
        ? 'The published run has no render for this page.'
        : 'No published renders are available, so only the Office reference is shown.'
    );
  stage.dataset.pair = String(hasReference && hasOurs);
  note.textContent = messages.join(' ');
  note.hidden = messages.length === 0;
  layout();
}

function absent(layer: HTMLElement, message: string): void {
  layer.dataset.absent = 'true';
  stage.dataset.pair = 'false';
  if (note.textContent?.includes(message)) return;
  note.textContent = `${note.textContent ?? ''} ${message}`.trim();
  note.hidden = false;
}

function layout(): void {
  const entry = currentPage();
  const width = entry?.width ?? 1275;
  const height = entry?.height ?? 1650;
  const scale =
    zoom === 'full'
      ? 1
      : Math.min((stage.clientWidth - 40) / width, (stage.clientHeight - 40) / height, 1);
  sheet.style.width = `${Math.round(width * scale)}px`;
  sheet.style.height = `${Math.round(height * scale)}px`;
}

function remember(): void {
  if (!current) return;
  const url = new URL(window.location.href);
  url.searchParams.set('doc', current.id);
  url.searchParams.set('page', String(page));
  window.history.replaceState(null, '', url);
}

function select(sample: Sample, requested?: number): void {
  current = sample;
  page = sample.pages.some((entry) => entry.page === requested)
    ? requested!
    : (sample.pages[0]?.page ?? 1);
  markCurrent();
  drawStrip();
  drawStage();
  remember();
}

function step(delta: number): void {
  if (!current) return;
  const index = current.pages.findIndex((entry) => entry.page === page);
  const next = current.pages[index + delta];
  if (!next) return;
  page = next.page;
  drawStrip();
  drawStage();
  remember();
}

function toggle(group: HTMLElement, attribute: string, value: string): void {
  for (const button of group.querySelectorAll('button'))
    button.ariaPressed = String(button.dataset[attribute] === value);
}

function start(loaded: Report, manifest: Latest | null): void {
  report = loaded;
  latest = manifest;
  const versions = Object.entries(loaded.versions)
    .map(([format, version]) => `${format} ${version}`)
    .join(' · ');
  revision.textContent = `${loaded.commit.slice(0, 8)} vs ${versions}`;
  loader.hidden = true;
  aside.hidden = false;
  main.hidden = false;
  drawList();
  const params = new URL(window.location.href).searchParams;
  const wanted = loaded.documents.find((sample) => sample.id === params.get('doc'));
  const first = wanted ?? rank(loaded.documents)[0];
  if (!first) return;
  select(first, Number(params.get('page')) || undefined);
  buttons.get(first.id)?.scrollIntoView({ block: 'center' });
}

async function json(url: string): Promise<unknown> {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`${response.status} ${url}`);
  return response.json();
}

async function boot(): Promise<void> {
  const override = sameOrigin(new URL(window.location.href).searchParams.get('report'), window.location.href);
  if (override) {
    start(parseReport(await json(override)), null);
    return;
  }
  const manifest = parseLatest(await json('/renders/latest.json'));
  start(parseReport(await json(reportUrl(manifest.sha))), manifest);
}

filter.addEventListener('input', drawList);

find('modes').addEventListener('click', (event) => {
  const mode = (event.target as HTMLElement).dataset.mode;
  if (!mode) return;
  stage.dataset.mode = mode;
  toggle(find('modes'), 'mode', mode);
});

find('zooms').addEventListener('click', (event) => {
  const next = (event.target as HTMLElement).dataset.zoom;
  if (next !== 'fit' && next !== 'full') return;
  zoom = next;
  toggle(find('zooms'), 'zoom', next);
  layout();
});

find<HTMLInputElement>('file').addEventListener('change', async (event) => {
  const file = (event.target as HTMLInputElement).files?.[0];
  if (!file) return;
  try {
    start(parseReport(JSON.parse(await file.text())), null);
  } catch (error) {
    loaderMessage.textContent = String(error instanceof Error ? error.message : error);
  }
});

handle.addEventListener('pointerdown', (event) => {
  handle.setPointerCapture(event.pointerId);
  const move = (moved: PointerEvent) => {
    const box = sheet.getBoundingClientRect();
    const ratio = Math.min(Math.max((moved.clientX - box.left) / box.width, 0), 1);
    sheet.style.setProperty('--split', `${(ratio * 100).toFixed(2)}%`);
  };
  move(event);
  handle.addEventListener('pointermove', move);
  handle.addEventListener(
    'pointerup',
    () => handle.removeEventListener('pointermove', move),
    { once: true }
  );
});

window.addEventListener('keydown', (event) => {
  if (event.target instanceof HTMLInputElement) return;
  if (event.key === 'b' || event.key === 'B') stage.dataset.blink = 'true';
  if (event.key === 'ArrowLeft') step(-1);
  if (event.key === 'ArrowRight') step(1);
});
window.addEventListener('keyup', (event) => {
  if (event.key === 'b' || event.key === 'B') delete stage.dataset.blink;
});
window.addEventListener('resize', layout);
reference.addEventListener('error', () => {
  if (reference.getAttribute('src')) absent(reference, 'The Office reference page could not be loaded.');
});
render.addEventListener('error', () => {
  if (render.getAttribute('src'))
    absent(render.parentElement as HTMLElement, 'The published render for this page could not be loaded.');
});

sheet.style.setProperty('--split', '50%');
boot().catch((error) => {
  loaderMessage.textContent = `${
    error instanceof Error ? error.message : error
  } — open a report.json to browse the measured scores and Office pages.`;
});
