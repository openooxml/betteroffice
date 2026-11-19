import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { pluginActivationId } from '../../../../shared/plugin-host/gesture';
import type { DocxCommandController } from '../commands/createDocxCommandStore';
import { createDocxPluginHost, type DocxPluginHost, type DocxPluginHostAccess } from './createDocxPluginHost';
import { defineDocxPlugin } from './defineDocxPlugin';
import type { DocxPlugin, DocxPluginContext } from './types';

const hosts: DocxPluginHost[] = [];
afterEach(() => {
  for (const host of hosts.splice(0)) host.close('unmounted');
  document.body.replaceChildren();
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

const settle = () => new Promise((done) => setTimeout(done, 0));

function editor(parent: ParentNode = document.body) {
  const root = document.createElement('div');
  const decoy = document.createElement('textarea');
  decoy.className = 'paged-editor__input';
  const input = document.createElement('textarea');
  input.className = 'paged-editor__input';
  root.append(decoy, input);
  parent.append(root);
  return { root, decoy, input };
}

function setup(input: HTMLElement | null, root: Element | null) {
  const commands = {
    focusEditor: () => {},
    scoped: () => undefined,
    shortcutAt: () => null,
    shortcutOwner: () => undefined,
    hostPluginChords: () => undefined,
    shortcutRevision: () => 0,
    setPluginCommands: () => {},
    refresh: () => {},
  } as unknown as DocxCommandController;
  const access = {
    pagedEditorRef: { current: null },
    writeMode: () => 'editing',
    commands: () => commands,
    layout: () => ({ queries: null, complete: false, failed: false }),
    subscribeLayout: () => () => {},
    geometry: () => null,
    translate: (key: string) => key,
    documentInput: () => input,
    editorRoot: () => root,
  };
  const host = createDocxPluginHost(access as unknown as DocxPluginHostAccess);
  host.setReporter(() => {});
  hosts.push(host);
  return host;
}

function plugin(id = 'acme.focus'): DocxPlugin {
  return defineDocxPlugin({ id, createState: () => null });
}

async function opened(input: HTMLElement | null = null, root: Element | null = null) {
  const host = setup(input, root);
  host.setPlugins([plugin()]);
  host.openWorker('v1');
  await settle();
  return host;
}

function contextOf(host: DocxPluginHost): DocxPluginContext<unknown> {
  return host.activations()[0]!.context;
}

function pluginButton(host: DocxPluginHost, parent: ParentNode = document.body) {
  const marker = document.createElement('div');
  marker.setAttribute('data-plugin-activation', pluginActivationId(host.activations()[0]!.activation));
  const button = document.createElement('button');
  marker.append(button);
  parent.append(marker);
  button.focus();
  return { marker, button };
}

describe('snapshot.inputFocused', () => {
  test('follows the input focus reports and republishes only on a change', async () => {
    const host = await opened();
    let notifications = 0;
    host.subscribe(() => { notifications += 1; });
    expect(contextOf(host).snapshot.inputFocused).toBe(false);
    host.inputFocusChanged(true);
    await settle();
    expect(contextOf(host).snapshot.inputFocused).toBe(true);
    expect(notifications).toBe(1);
    host.inputFocusChanged(true);
    await settle();
    expect(notifications).toBe(1);
    host.inputFocusChanged(false);
    await settle();
    expect(contextOf(host).snapshot.inputFocused).toBe(false);
    expect(notifications).toBe(2);
  });

  test('is independent of focus inside the document surface', async () => {
    const host = await opened();
    host.focusChanged(true);
    await settle();
    expect(contextOf(host).snapshot).toMatchObject({ focused: true, inputFocused: false });
    host.inputFocusChanged(true);
    await settle();
    expect(contextOf(host).snapshot).toMatchObject({ focused: true, inputFocused: true });
    host.inputFocusChanged(false);
    await settle();
    expect(contextOf(host).snapshot).toMatchObject({ focused: true, inputFocused: false });
  });

  test('a report before the document opens reaches the load snapshot', async () => {
    const host = setup(null, null);
    host.setPlugins([plugin()]);
    host.inputFocusChanged(true);
    host.openWorker('v1');
    await settle();
    expect(contextOf(host).snapshot.inputFocused).toBe(true);
  });
});

describe('context.focusDocument', () => {
  test('from the plugin’s own UI it focuses the editor’s input, not an element with its class', async () => {
    const { root, decoy, input } = editor();
    const host = await opened(input, root);
    const { button } = pluginButton(host);
    expect(document.activeElement).toBe(button);
    expect(contextOf(host).focusDocument()).toEqual({ ok: true });
    expect(document.activeElement).toBe(input);
    expect(document.activeElement).not.toBe(decoy);
  });

  test('an editor dialog in the editor root refuses it and focus stays in the plugin UI', async () => {
    const { root, input } = editor();
    const host = await opened(input, root);
    const dialog = document.createElement('div');
    dialog.setAttribute('role', 'dialog');
    root.append(dialog);
    const { button } = pluginButton(host);
    let answer: unknown;
    expect(() => { answer = contextOf(host).focusDocument(); }).not.toThrow();
    expect(answer).toMatchObject({ ok: false, failure: { code: 'unsupported-policy' } });
    expect(document.activeElement).toBe(button);
    dialog.setAttribute('aria-modal', 'true');
    expect(contextOf(host).focusDocument()).toMatchObject({ ok: false, failure: { code: 'unsupported-policy' } });
    expect(document.activeElement).toBe(button);
    dialog.remove();
    expect(contextOf(host).focusDocument()).toEqual({ ok: true });
    expect(document.activeElement).toBe(input);
  });

  test('a non-modal editor dialog or a dialog in plugin UI does not refuse it', async () => {
    const { root, input } = editor();
    const host = await opened(input, root);
    const find = document.createElement('div');
    find.setAttribute('role', 'dialog');
    find.setAttribute('aria-modal', 'false');
    root.append(find);
    const { marker, button } = pluginButton(host, root);
    const own = document.createElement('div');
    own.setAttribute('role', 'dialog');
    marker.append(own);
    expect(document.activeElement).toBe(button);
    expect(contextOf(host).focusDocument()).toEqual({ ok: true });
    expect(document.activeElement).toBe(input);
  });

  test('from a background callback it is refused and focus stays where it is', async () => {
    const { root, input } = editor();
    const host = await opened(input, root);
    let focused = 0;
    input.addEventListener('focus', () => { focused += 1; });
    const outside = document.createElement('input');
    document.body.append(outside);
    outside.focus();
    expect(contextOf(host).focusDocument()).toMatchObject({ ok: false, failure: { code: 'unsupported-policy' } });
    expect(focused).toBe(0);
    expect(document.activeElement).toBe(outside);
    pluginButton(host);
    expect(contextOf(host).focusDocument()).toEqual({ ok: true });
    expect(focused).toBe(1);
    expect(document.activeElement).toBe(input);
  });

  test('is allowed while read-only and refused once the plugin ended', async () => {
    const { root, input } = editor();
    const host = await opened(input, root);
    host.modeChanged('viewing', true);
    await settle();
    const { button } = pluginButton(host);
    const context = contextOf(host);
    expect(context.snapshot.readOnly).toBe(true);
    expect(context.focusDocument()).toEqual({ ok: true });
    expect(document.activeElement).toBe(input);
    button.focus();
    host.close('unmounted');
    expect(context.focusDocument()).toMatchObject({ ok: false, failure: { code: 'plugin-unavailable' } });
    expect(document.activeElement).toBe(button);
  });

  test('without a document input, or when focus throws or does not take, it is refused, never thrown', async () => {
    const throwing = document.createElement('textarea');
    throwing.focus = () => { throw new Error('focus failed'); };
    const inert = document.createElement('textarea');
    inert.focus = () => {};
    document.body.append(throwing, inert);
    for (const input of [null, throwing, inert]) {
      const host = await opened(input, document.body);
      const { button } = pluginButton(host);
      let answer: unknown;
      expect(() => { answer = contextOf(host).focusDocument(); }).not.toThrow();
      expect(answer).toMatchObject({ ok: false, failure: { code: 'unsupported-policy' } });
      expect(document.activeElement).toBe(button);
    }
  });

  test('inside a shadow root it reads focus from the editor’s own root', async () => {
    const shadowHost = document.createElement('div');
    document.body.append(shadowHost);
    const shadow = shadowHost.attachShadow({ mode: 'open' });
    const { root, input } = editor(shadow);
    const host = await opened(input, root);
    const { button } = pluginButton(host, shadow);
    expect(shadow.activeElement).toBe(button);
    expect(contextOf(host).focusDocument()).toEqual({ ok: true });
    expect(shadow.activeElement).toBe(input);
  });
});
