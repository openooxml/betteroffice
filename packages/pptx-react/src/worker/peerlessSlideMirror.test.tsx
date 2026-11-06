import { expect, test } from 'bun:test';
import { act, fireEvent, render, waitFor } from '@testing-library/react';
import type { SlideDisplayList, TextBoxPrimitive, TextLanguage } from '@betteroffice/pptx';
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
