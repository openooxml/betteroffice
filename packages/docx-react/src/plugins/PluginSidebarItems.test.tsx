import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import { useState } from 'react';

const ownsDom = !GlobalRegistrator.isRegistered;
if (ownsDom) GlobalRegistrator.register();

import { UNAVAILABLE_DOCX_COMMANDS } from '../commands/createDocxCommandStore';
import type { DocxPluginActivation, DocxPluginHost } from './createDocxPluginHost';
import { UnifiedSidebar } from '../components/UnifiedSidebar';
import type { ReactSidebarItem } from '../plugin-api/types';
import { managedSidebarItems, mergeSidebarItems } from './PluginSidebarItems';
import type { DocxPluginSidebarItem } from './types';

const { cleanup, fireEvent, render } = await import('@testing-library/react');
const quiet = console.error;
afterEach(() => {
  cleanup();
  console.error = quiet;
});
afterAll(async () => {
  if (ownsDom) await GlobalRegistrator.unregister();
});

function harness(items: () => DocxPluginSidebarItem<unknown>[]) {
  const failures: string[] = [];
  const host = {
    guard(pluginId: string, phase: string, call: (context: unknown) => unknown, fallback: unknown) {
      try {
        return call({});
      } catch {
        failures.push(`${pluginId}:${phase}`);
        return fallback;
      }
    },
    fail: (pluginId: string, phase: string) => failures.push(`${pluginId}:${phase}`),
    commandStore: () => UNAVAILABLE_DOCX_COMMANDS,
  } as unknown as DocxPluginHost;
  const activation = {
    pluginId: 'acme',
    key: 'acme#1',
    plugin: { getSidebarItems: items },
    context: {},
  } as unknown as DocxPluginActivation;
  return { host, activation, failures };
}

const anchor = (version: string, paraId = '00000001') => ({ version, story: 'body', paraId });

describe('managed sidebar items', () => {
  test('namespace ids, place cards pre-zoom and hide unplaceable anchors', () => {
    const { host, activation } = harness(() => [
      { id: 'current', anchor: anchor('v2'), priority: 1, render: () => <p>current</p> },
      { id: 'stale', anchor: anchor('v1'), render: () => <p>stale</p> },
    ]);
    const items = managedSidebarItems(host, [activation], (where) =>
      where.version === 'v2' ? { position: 7, y: 40 } : null
    );
    expect(
      items.map(({ id, fixedY, priority, hidden }) => ({ id, fixedY, priority, hidden }))
    ).toEqual([
      { id: 'plugin:acme/current', fixedY: 40, priority: 1, hidden: undefined },
      { id: 'plugin:acme/stale', fixedY: undefined, priority: undefined, hidden: true },
    ]);
    expect(items[0].anchorPos).toBe(7);
  });

  test('one card component keeps its state until its id or activation changes', () => {
    function Card({ item }: { item: DocxPluginSidebarItem<unknown> }) {
      const [clicks, setClicks] = useState(0);
      return (
        <button
          type="button"
          data-testid="card"
          data-id={item.id}
          data-clicks={clicks}
          onClick={() => setClicks((count) => count + 1)}
        />
      );
    }
    let id = 'note';
    const { host, activation } = harness(() => [{ id, anchor: anchor('v'), render: Card }]);
    const sidebar = (current: DocxPluginActivation, placed = true) => (
      <UnifiedSidebar
        items={managedSidebarItems(host, [current], () => (placed ? { position: 0, y: 10 } : null))}
        anchorPositions={new Map()}
        renderedDomContext={null}
        pageWidth={800}
        zoom={1}
        editorContainerRef={{ current: null }}
      />
    );
    const view = render(sidebar(activation));
    const card = () => view.getByTestId('card');
    fireEvent.click(card());
    expect(card().dataset).toMatchObject({ id: 'note', clicks: '1' });

    view.rerender(sidebar({ ...activation, context: { ...activation.context } }, false));
    expect(card().parentElement!.style.visibility).toBe('hidden');
    view.rerender(sidebar({ ...activation, context: { ...activation.context } }));
    expect(card().parentElement!.style.visibility).toBe('');
    expect(card().dataset.clicks).toBe('1');

    id = 'other';
    view.rerender(sidebar(activation));
    expect(card().dataset).toMatchObject({ id: 'other', clicks: '0' });
    fireEvent.click(card());
    view.rerender(sidebar({ ...activation, key: 'acme#2' }));
    expect(card().dataset.clicks).toBe('0');
  });

  test('duplicate local ids fail the plugin; a throwing card reports a render failure', () => {
    const duplicate = harness(() => [
      { id: 'same', anchor: anchor('v'), render: () => null },
      { id: 'same', anchor: anchor('v'), render: () => null },
    ]);
    expect(
      managedSidebarItems(duplicate.host, [duplicate.activation], () => ({ position: 0, y: 0 }))
    ).toEqual([]);
    expect(duplicate.failures).toEqual(['acme:sidebar']);

    console.error = () => {};
    const throwing = harness(() => [
      {
        id: 'boom',
        anchor: anchor('v'),
        render: () => {
          throw new Error('card failed');
        },
      },
    ]);
    const [item] = managedSidebarItems(throwing.host, [throwing.activation], () => ({
      position: 0,
      y: 0,
    }));
    const view = render(
      <div>
        {item.render({ isExpanded: false, onToggleExpand: () => {}, measureRef: () => {} })}
        <span data-testid="sibling">still here</span>
      </div>
    );
    expect(throwing.failures).toEqual(['acme:render']);
    expect(view.getByTestId('sibling').textContent).toBe('still here');
  });
});

describe('sidebar merging and caches', () => {
  const item = (id: string, extra: Partial<ReactSidebarItem> = {}): ReactSidebarItem => ({
    id,
    anchorPos: 0,
    render: () => <p data-testid={`card-${id}`}>{id}</p>,
    ...extra,
  });

  test('host items never shadow comments or plugin cards', () => {
    const warn = console.warn;
    console.warn = () => {};
    try {
      const merged = mergeSidebarItems(
        [item('comment-1')],
        [item('plugin:acme/note'), item('comment-1'), item('host')],
        [item('plugin:acme/note')]
      );
      expect(merged.map((entry) => entry.id)).toEqual(['comment-1', 'host', 'plugin:acme/note']);
    } finally {
      console.warn = warn;
    }
  });

  test('a card that returns is not placed at the Y it had before it was removed', () => {
    const sidebar = (items: ReactSidebarItem[]) => (
      <UnifiedSidebar
        items={items}
        anchorPositions={new Map()}
        renderedDomContext={null}
        pageWidth={800}
        zoom={1}
        editorContainerRef={{ current: null }}
      />
    );
    const view = render(sidebar([item('plugin:acme/card', { fixedY: 100 })]));
    const top = () => view.getByTestId('card-plugin:acme/card').parentElement!.style.top;
    expect(top()).toBe('100px');
    view.rerender(sidebar([]));
    view.rerender(sidebar([item('plugin:acme/card')]));
    expect(top()).not.toBe('100px');
  });
});
