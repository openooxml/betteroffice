/**
 * @betteroffice/docx-react/plugin-api
 *
 * Generic plugin interface and host component for integrating external
 * plugins with the editor. Pairs with the framework-agnostic plugin types
 * exported from `@betteroffice/docx/plugin-api`.
 *
 * @example
 * ```tsx
 * import { PluginHost, type EditorPlugin } from '@betteroffice/docx-react/plugin-api';
 *
 * function MyEditor() {
 *   return (
 *     <PluginHost plugins={[myPlugin]}>
 *       <DocxEditor document={doc} onChange={handleChange} />
 *     </PluginHost>
 *   );
 * }
 * ```
 *
 * @packageDocumentation
 * @public
 */

// Types (React-specific + re-exports from core)
export type {
  EditorPlugin,
  ReactEditorPlugin,
  PluginPanelProps,
  PanelConfig,
  PluginContext,
  PluginHostProps,
  PluginHostRef,
  RenderedDomContext,
  PositionCoordinates,
  SidebarItem,
  SidebarItemContext,
  ReactSidebarItem,
  SidebarItemRenderProps,
} from './types';

// Components
export { PluginHost, PLUGIN_HOST_STYLES } from './PluginHost';

// Rendered DOM Context
export {
  createCanvasHostProjector,
  createRenderedDomContext,
  RenderedDomContextImpl,
} from './RenderedDomContext';
