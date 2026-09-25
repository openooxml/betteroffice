/**
 * Core Plugin API — Framework-Agnostic
 *
 * Exports the core plugin interfaces and types that can be used
 * by any framework adapter (React, Vue, etc.).
 *
 * @experimental Plugin API is still evolving. Breaking changes may
 * happen in minor releases until plugin authors stabilize the contract.
 * The snapshot-based `EditorPluginCore`, `PluginPanelProps`, `PanelConfig`
 * and `SidebarItemContext` are deprecated; React hosts install plugins with
 * `defineDocxPlugin` and the `plugins` prop of `@betteroffice/docx-react`.
 * @packageDocumentation
 * @public
 */

export type {
  EditorPluginCore,
  PluginPanelProps,
  PanelConfig,
  RenderedDomContext,
  PositionCoordinates,
  SidebarItem,
  SidebarItemContext,
} from './types';
