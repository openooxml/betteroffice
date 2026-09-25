import type { ReactSidebarItem } from '../plugin-api/types';
import type { DocxPluginActivation, DocxPluginHost } from './createDocxPluginHost';
import { PluginRenderScope } from './PluginRenderScope';
import type { DocxPluginSidebarItem } from './types';

/** Where a sidebar anchor renders: its display position and its pre-zoom Y. */
export type SidebarAnchorPlacement = { position: number; y: number };

function checkedItems(value: unknown): readonly DocxPluginSidebarItem<unknown>[] {
  if (!Array.isArray(value)) throw new TypeError('getSidebarItems must return an array');
  const ids = new Set<string>();
  for (const item of value as DocxPluginSidebarItem<unknown>[]) {
    if (typeof item?.id !== 'string' || item.id.length === 0) {
      throw new TypeError('Every sidebar item needs an id');
    }
    if (ids.has(item.id)) throw new TypeError(`Duplicate sidebar item id "${item.id}"`);
    ids.add(item.id);
    const anchor = item.anchor;
    if (
      typeof anchor?.version !== 'string' ||
      typeof anchor.story !== 'string' ||
      typeof anchor.paraId !== 'string' ||
      (typeof item.render !== 'function' &&
        (item.render === null || typeof item.render !== 'object'))
    ) {
      throw new TypeError(
        `Sidebar item "${item.id}" needs a versioned paragraph anchor and a renderer`
      );
    }
  }
  return value;
}

/**
 * The sidebar cards of every ready plugin, with namespaced ids and placements resolved now.
 * Items whose anchor is stale, ambiguous, missing or not rendered are left out.
 */
export function managedSidebarItems(
  host: DocxPluginHost,
  activations: readonly DocxPluginActivation[],
  place: (anchor: DocxPluginSidebarItem<unknown>['anchor']) => SidebarAnchorPlacement | null
): ReactSidebarItem[] {
  const items: ReactSidebarItem[] = [];
  for (const activation of activations) {
    const generate = activation.plugin.getSidebarItems;
    if (!generate) continue;
    const produced = host.guard(
      activation.pluginId,
      'sidebar',
      (context) => checkedItems(generate.call(activation.plugin, context)),
      null
    );
    for (const item of produced ?? []) {
      const placement = place(item.anchor);
      if (!placement) continue;
      const Render = item.render;
      items.push({
        id: `plugin:${activation.pluginId}/${item.id}`,
        anchorPos: placement.position,
        fixedY: placement.y,
        ...(item.priority === undefined ? {} : { priority: item.priority }),
        ...(item.estimatedHeight === undefined ? {} : { estimatedHeight: item.estimatedHeight }),
        render: (props) => (
          <PluginRenderScope key={activation.key} host={host} activation={activation}>
            <Render context={activation.context} {...props} />
          </PluginRenderScope>
        ),
      });
    }
  }
  return items;
}

/**
 * The sidebar's items: comments, then unmanaged host items, then managed plugin cards. A host
 * item whose id a plugin card already uses is dropped with a warning.
 */
export function mergeSidebarItems(
  comments: readonly ReactSidebarItem[],
  unmanaged: readonly ReactSidebarItem[],
  managed: readonly ReactSidebarItem[]
): ReactSidebarItem[] {
  const ids = new Set([...comments, ...managed].map((item) => item.id));
  const kept = unmanaged.filter((item) => {
    if (!ids.has(item.id)) return true;
    console.warn(`[DocxEditor] pluginSidebarItems id "${item.id}" collides with another item`);
    return false;
  });
  return [...comments, ...kept, ...managed];
}
