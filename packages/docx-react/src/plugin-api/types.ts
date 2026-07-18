/** React plugin contracts for the Yrs-backed editor. */

import type { ReactNode } from 'react';
import type { Document } from '@betteroffice/docx/types/document';

export type {
  PanelConfig,
  RenderedDomContext,
  PositionCoordinates,
  SidebarItem,
} from '@betteroffice/docx/plugin-api';

import type {
  PanelConfig,
  RenderedDomContext,
  SidebarItem,
} from '@betteroffice/docx/plugin-api';

export interface PluginPanelProps<TState = unknown> {
  /** Current Yrs-projected document snapshot. */
  document: Document | null;
  /** Scroll to a display position. */
  scrollToPosition: (position: number) => void;
  /** Select a display-position range. */
  selectRange: (from: number, to: number) => void;
  pluginState: TState;
  panelWidth: number;
  renderedDomContext: RenderedDomContext | null;
}

export interface SidebarItemRenderProps {
  isExpanded: boolean;
  onToggleExpand: () => void;
  measureRef: (element: HTMLDivElement | null) => void;
}

export interface ReactSidebarItem extends SidebarItem {
  render: (props: SidebarItemRenderProps) => ReactNode;
  estimatedHeight?: number;
}

export interface SidebarItemContext {
  document: Document | null;
  renderedDomContext: RenderedDomContext | null;
  anchorPositions: Map<string, number>;
  zoom: number;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export interface ReactEditorPlugin<TState = any> {
  id: string;
  name: string;
  panelConfig?: PanelConfig;
  /** Initialize state from the current projected document. */
  initialize?: (document: Document | null) => TState;
  /** Recompute state after a committed document change. */
  onStateChange?: (document: Document) => TState | undefined;
  destroy?: () => void;
  styles?: string;
  Panel?: React.ComponentType<PluginPanelProps<TState>>;
  renderOverlay?: (context: RenderedDomContext, state: TState) => ReactNode;
  getSidebarItems?: (state: TState, context: SidebarItemContext) => ReactSidebarItem[];
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type EditorPlugin<TState = any> = ReactEditorPlugin<TState>;

export interface PluginContext {
  plugins: EditorPlugin[];
  document: Document | null;
  getPluginState: <T>(pluginId: string) => T | undefined;
  setPluginState: <T>(pluginId: string, state: T) => void;
  scrollToPosition: (position: number) => void;
  selectRange: (from: number, to: number) => void;
}

export interface PluginHostProps {
  plugins: EditorPlugin[];
  children: React.ReactElement;
  className?: string;
}

export interface PluginHostRef {
  getPluginState: <T>(pluginId: string) => T | undefined;
  setPluginState: <T>(pluginId: string, state: T) => void;
  getDocument: () => Document | null;
  refreshPluginStates: () => void;
}
