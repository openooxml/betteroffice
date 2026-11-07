import { expect, mock, spyOn, test } from 'bun:test';
import { act, fireEvent, render, waitFor } from '@testing-library/react';
import type { SlideDisplayList, TextBoxPrimitive, TextLanguage, TextLinkTarget } from '@betteroffice/pptx';
import { LocaleProvider } from '../i18n';
import { PeerlessSlideMirror } from './PeerlessSlideMirror';
import type { PeerlessPresentation } from './usePeerlessPresentation';
import { open } from './peerlessTestHarness';

const mirror = (view: Awaited<ReturnType<typeof open>>['view']) => view.getByTestId('pptx-slide-mirror');
const live = (view: Awaited<ReturnType<typeof open>>['view']) => view.getByTestId('pptx-live-region');

test('the slide mirror is a labelled document holding the text as whole paragraphs', async () => {
  const { view } = await open();
  const document = mirror(view);
  expect(document.getAttribute('role')).toBe('document');
  expect(document.getAttribute('aria-label')).toBe('Slide 1 of 2');
  expect([...document.querySelectorAll('p')].map((node) => node.textContent)).toEqual(['alpha beta', 'x', 'alpha beta']);
  expect(document.closest('[role="application"]')).not.toBeNull();
  expect(view.getAllByRole('application')).toHaveLength(1);
});

test('slide thumbnails carry the slide position in their names', async () => {
  const { view } = await open();
  expect(view.getByRole('button', { name: 'Slide 2 of 2' })).toBeTruthy();
});

test('moving to another slide is announced, and only while focus is in the editor', async () => {
  const { view, stage } = await open();
  act(() => { (stage as HTMLElement).focus(); });
  fireEvent.keyDown(stage, { key: 'PageDown' });
  await waitFor(() => expect(mirror(view).getAttribute('aria-label')).toBe('Slide 2 of 2'));
  await waitFor(() => expect(live(view).textContent?.trim()).toBe('Slide 2 of 2'));
  act(() => { (stage as HTMLElement).blur(); });
  const before = live(view).textContent;
  fireEvent.click(view.getByRole('button', { name: 'Slide 1 of 2' }));
  await waitFor(() => expect(mirror(view).getAttribute('aria-label')).toBe('Slide 1 of 2'));
  act(() => { (document.activeElement as HTMLElement | null)?.blur(); });
  await new Promise((resolve) => setTimeout(resolve, 250));
  expect(live(view).textContent).toBe(before);
});

test('selecting a shape is announced by name', async () => {
  const { view, stage, select } = await open();
  act(() => { (stage as HTMLElement).focus(); });
  await select();
  await waitFor(() => expect(live(view).textContent?.trim()).toBe('Text selected'));
});

test('page keys in the mirror change slides, and other keys leave the selection alone', async () => {
  const { view, select, client } = await open();
  await select();
  const document = mirror(view);
  act(() => { document.focus(); });
  fireEvent.keyDown(document, { key: 'Delete' });
  await new Promise((resolve) => setTimeout(resolve, 50));
  expect(client.state.selection).toMatchObject({ kind: 'object' });
  fireEvent.keyDown(document, { key: 'PageDown' });
  await waitFor(() => expect(mirror(view).getAttribute('aria-label')).toBe('Slide 2 of 2'));
});

test('a host page document around the editor leaves the stage keys working', async () => {
  const { view, stage } = await open();
  view.container.setAttribute('role', 'document');
  act(() => { (stage as HTMLElement).focus(); });
  fireEvent.keyDown(stage, { key: 'PageDown' });
  await waitFor(() => expect(mirror(view).getAttribute('aria-label')).toBe('Slide 2 of 2'));
});

test('paragraphs and joined headings mark authored languages without splitting one language into text nodes', () => {
  const box = (shapeId: string, text: string, languages?: TextLanguage[]): TextBoxPrimitive => ({
    kind: 'textBox', objectId: 1, shapeId, x: 0, y: 0, w: 400, h: 100, anchor: 'top', lines: [],
    paragraphs: [{ level: 0, languages, runs: text.split('|').map((part) => ({
      text: part, fontFamily: 'Arial', fontSizePt: 12, color: '#000000',
    })) }],
  });
  const list: SlideDisplayList = {
    contractVersion: 2, width: 960, height: 540, titleShape: 'title', primitives: [
      box('title', 'Hello world', [{ start: 0, end: 11, lang: 'en-US' }]),
      box('title', 'Bonjour', [{ start: 0, end: 7, lang: 'fr-FR' }]),
      box('single', 'Whole |words', [{ start: 0, end: 11, lang: 'en-US' }]),
      box('mixed', 'Hello bonjour world', [
        { start: 0, end: 6, lang: 'en-US' }, { start: 6, end: 13, lang: 'fr-FR' }, { start: 13, end: 19, lang: 'en-US' },
      ]),
      box('unknown', 'Unmarked words'),
    ],
  };
  const { container } = render(<LocaleProvider i18n={{ _lang: 'de' }}>
    <PeerlessSlideMirror run={{} as PeerlessPresentation} list={list} label="Slide 1" scale={1} />
  </LocaleProvider>);
  const heading = container.querySelector('[role="heading"]')!;
  expect(heading.getAttribute('lang')).toBe('en-US');
  expect(heading.textContent).toBe('Hello world Bonjour');
  expect(heading.querySelector('span')?.getAttribute('lang')).toBe('fr-FR');
  expect(heading.querySelector('span')?.textContent).toBe('Bonjour');
  const [single, mixed, unknown] = [...container.querySelectorAll('p')];
  expect(single!.getAttribute('lang')).toBe('en-US');
  expect(single!.textContent).toBe('Whole words');
  expect(single!.childNodes.length).toBe(1);
  expect(single!.firstChild!.nodeType).toBe(3);
  expect(mixed!.getAttribute('lang')).toBe('en-US');
  expect(mixed!.textContent).toBe('Hello bonjour world');
  expect(mixed!.querySelectorAll('span')).toHaveLength(1);
  expect(mixed!.querySelector('span')?.getAttribute('lang')).toBe('fr-FR');
  expect(mixed!.querySelector('span')?.textContent).toBe('bonjour');
  expect(unknown!.hasAttribute('lang')).toBe(false);
  expect(unknown!.querySelector('[lang]')).toBeNull();
  expect(unknown!.childNodes.length).toBe(1);
});

function linkMirror(targets: { text: string; target?: TextLinkTarget }[], active = 's1') {
  const slides = [{ id: 's1' }, { id: 's2' }, { id: 's3' }];
  let slideId = active;
  const dispatch = mock(async (command: { kind: string; payload: { slideId: string } }) => {
    slideId = command.payload.slideId;
    return {};
  });
  const run = {
    current: true, client: { state: { slides } },
    store: { getSnapshot: () => ({ confirmed: { frame: { slideId } }, composition: null }), dispatch },
  } as unknown as PeerlessPresentation;
  const list: SlideDisplayList = {
    contractVersion: 2, width: 960, height: 540,
    primitives: targets.map(({ text, target }, index): TextBoxPrimitive => ({
      kind: 'textBox', objectId: index + 1, shapeId: `shape-${index}`, x: 0, y: index * 30, w: 400, h: 30, anchor: 'top', lines: [],
      paragraphs: [{ level: 0, links: target ? [{ start: 0, end: text.length, target }] : [],
        runs: [{ text, fontFamily: 'Arial', fontSizePt: 12, color: '#000000' }] }],
    })),
  };
  const view = render(<LocaleProvider><PeerlessSlideMirror run={run} list={list} label="Slide links" scale={1} /></LocaleProvider>);
  return { ...view, dispatch, run, list };
}

test('only safe absolute hrefs become anchors and retain their exact authored href', () => {
  const safe = ['https://example.test/path', 'HTTP://example.test', 'mailto:reader@example.test', 'tel:+123', ' \tHTTPS://example.test/\r'];
  const hostile = ['javascript:alert(1)', 'JaVaScRiPt:alert(1)', 'java\tscript:alert(1)', '\x01javascript:alert(1)', 'data:text/html,bad', 'vbscript:bad', 'file:///tmp/file', 'foo.html', '//evil', 'ftp://example.test', 'https:'];
  const view = linkMirror([...safe, ...hostile].map((href, index) => ({ text: `Link ${index}`, target: { kind: 'url', href } })));
  expect(view.getAllByRole('link')).toHaveLength(safe.length);
  safe.forEach((href, index) => expect(view.getByRole('link', { name: `Link ${index}` }).getAttribute('href')).toBe(href));
  hostile.forEach((_, index) => {
    const node = view.getByText(`Link ${safe.length + index}`, { exact: true });
    expect(node.closest('a')).toBeNull();
    expect(node.childNodes).toHaveLength(1);
  });
});

test('external clicks open an isolated tab and recheck hrefs injected into the DOM', () => {
  const opened = spyOn(window, 'open').mockImplementation(() => null);
  try {
    const view = linkMirror([{ text: 'Read more', target: { kind: 'url', href: 'https://example.test/path' } }]);
    const anchor = view.getByRole('link', { name: 'Read more' });
    expect(anchor.getAttribute('target')).toBe('_blank');
    expect(anchor.getAttribute('tabindex')).toBeNull();
    fireEvent.click(anchor);
    expect(opened).toHaveBeenCalledWith('https://example.test/path', '_blank', 'noopener,noreferrer');
    opened.mockClear();
    for (const href of ['javascript:alert(1)', 'java\tscript:alert(1)', '\x01javascript:alert(1)', 'foo.html', '//evil']) {
      anchor.setAttribute('href', href);
      const click = new MouseEvent('click', { bubbles: true, cancelable: true });
      anchor.dispatchEvent(click);
      expect(click.defaultPrevented).toBe(true);
    }
    expect(opened).not.toHaveBeenCalled();
    expect(view.dispatch).not.toHaveBeenCalled();
  } finally { opened.mockRestore(); }
});

test('slide and jump links route activation to deck ids and omit invalid slide indices', async () => {
  const view = linkMirror([
    { text: 'Third slide', target: { kind: 'slide', index: 2 } },
    { text: 'Missing slide', target: { kind: 'slide', index: 3 } },
    { text: 'Negative slide', target: { kind: 'slide', index: -1 } },
    { text: 'Fractional slide', target: { kind: 'slide', index: 0.5 } },
    ...(['next', 'previous', 'first', 'last'] as const).map((jump) => ({ text: jump, target: { kind: 'jump', jump } as const })),
  ]);
  expect(view.getAllByRole('link')).toHaveLength(5);
  for (const name of ['Missing slide', 'Negative slide', 'Fractional slide']) expect(view.getByText(name).closest('a')).toBeNull();
  const third = view.getByRole('link', { name: 'Third slide' });
  expect(third.getAttribute('href')).toBe('#');
  expect(third.getAttribute('data-slide-link')).toBe('slide:2');
  fireEvent.click(third);
  await waitFor(() => expect(view.dispatch).toHaveBeenLastCalledWith({ kind: 'activateSlide', payload: { slideId: 's3' } }));
  for (const [jump, id] of [['previous', 's2'], ['first', 's1'], ['next', 's2'], ['last', 's3']] as const) {
    const anchor = view.getByRole('link', { name: jump });
    expect(anchor.getAttribute('data-slide-link')).toBe(`jump:${jump}`);
    fireEvent.click(anchor);
    await waitFor(() => expect(view.dispatch).toHaveBeenLastCalledWith({ kind: 'activateSlide', payload: { slideId: id } }));
  }
  view.dispatch.mockClear();
  fireEvent.click(view.getByRole('link', { name: 'next' }));
  expect(view.dispatch).not.toHaveBeenCalled();
  third.setAttribute('href', 'javascript:alert(1)');
  fireEvent.click(third);
  expect(view.dispatch).not.toHaveBeenCalled();
});

test('one link spanning language boundaries has one accessible name', () => {
  const view = linkMirror([{ text: 'Read bonjour', target: { kind: 'url', href: 'https://example.test' } }]);
  const box = view.list.primitives[0] as TextBoxPrimitive;
  box.paragraphs[0]!.languages = [{ start: 0, end: 5, lang: 'en-US' }, { start: 5, end: 12, lang: 'fr-FR' }];
  view.rerender(<LocaleProvider><PeerlessSlideMirror run={view.run} list={{ ...view.list }} label="Slide links" scale={1} /></LocaleProvider>);
  expect(view.getAllByRole('link')).toHaveLength(1);
  expect(view.getByRole('link', { name: 'Read bonjour' }).querySelector('span')?.textContent).toBe('Read ');
});
