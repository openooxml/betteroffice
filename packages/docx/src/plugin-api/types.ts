/**
 * Framework-Agnostic Plugin Interface for the DOCX Editor
 *
 * Core plugin types that can be used by any framework (React, Vue, etc.).
 * Framework-specific adapters extend EditorPluginCore with their own
 * UI rendering capabilities (e.g., ReactEditorPlugin, VueEditorPlugin).
 * @packageDocumentation
 * @public
 */

import type { Document } from '../types/document';

/**
 * Coordinates returned by position lookup in the rendered DOM.
 */
export interface PositionCoordinates {
  x: number;
  y: number;
  height: number;
}

/**
 * Context for accessing rendered-page geometry in the paged editor.
 *
 * Provides position mapping over the canvas renderer's output (Rust
 * display-list queries, with the accessibility mirror's DOM as fallback).
 * Use this for rendering overlays, annotations, and other visual elements
 * positioned relative to rendered content.
 *
 * The mirror DOM uses data-doc-start/data-doc-end attributes on spans
 * to map between display positions and DOM elements.
 */
export interface RenderedDomContext {
  /** The container element holding all rendered pages. */
  pagesContainer: HTMLElement;

  /**
   * Get pixel coordinates for a display position in the rendered DOM.
   * Returns null if the position cannot be found.
   */
  getCoordinatesForPosition(position: number): PositionCoordinates | null;

  /**
   * Find DOM elements that overlap with a display-position range.
   */
  findElementsForRange(from: number, to: number): Element[];

  /**
   * Get bounding rectangles for a range of text, accounting for line wraps.
   * Returns rects relative to the pages container.
   */
  getRectsForRange(
    from: number,
    to: number
  ): Array<{ x: number; y: number; width: number; height: number }>;

  /** Query-backed caret coordinates for one explicit header/footer page. */
  getCoordinatesForHfPosition(
    region: 'header' | 'footer',
    rId: string,
    position: number,
    pageIndex: number
  ): PositionCoordinates | null;

  /** Query-backed range rectangles for one explicit header/footer page. */
  getRectsForHfRange(
    region: 'header' | 'footer',
    rId: string,
    from: number,
    to: number,
    pageIndex: number
  ): Array<{ x: number; y: number; width: number; height: number }>;

  /** Bounds of one rendered page relative to the pages container. */
  getPageBounds(pageIndex: number): { x: number; y: number; width: number; height: number } | null;

  /** Current zoom level (1 = 100%). */
  zoom: number;

  /**
   * Offset of the pages container from its parent viewport.
   */
  getContainerOffset(): { x: number; y: number };
}

/**
 * Props passed to plugin panel components (framework-agnostic base).
 */
export interface PluginPanelProps<TState = unknown> {
  /** Current serializer-facing document snapshot. */
  document: Document | null;

  /** Scroll editor to a specific position */
  scrollToPosition: (pos: number) => void;

  /** Select a range in the editor */
  selectRange: (from: number, to: number) => void;

  /** Plugin-specific state (managed by the plugin) */
  pluginState: TState;

  /** Width of the panel in pixels */
  panelWidth: number;

  /**
   * Context for the rendered pages (canvas renderer output).
   * May be null if layout hasn't completed yet.
   */
  renderedDomContext: RenderedDomContext | null;
}

/**
 * Configuration for plugin panel rendering.
 */
export interface PanelConfig {
  /** Where to render the panel */
  position: 'left' | 'right' | 'bottom';

  /** Default width/height of the panel */
  defaultSize: number;

  /** Minimum size */
  minSize?: number;

  /** Maximum size */
  maxSize?: number;

  /** Whether the panel is resizable */
  resizable?: boolean;

  /** Whether the panel can be collapsed */
  collapsible?: boolean;

  /** Initial collapsed state */
  defaultCollapsed?: boolean;
}

/**
 * Framework-agnostic core plugin interface.
 *
 * Contains all non-UI plugin capabilities:
 * - State management (initialize, onStateChange, destroy)
 * - CSS injection
 * - Panel configuration
 *
 * Framework adapters (ReactEditorPlugin, VueEditorPlugin) extend this
 * with their own Panel component type and renderOverlay function.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export interface EditorPluginCore<TState = any> {
  /** Unique plugin identifier */
  id: string;

  /** Display name for the plugin */
  name: string;

  /**
   * Configuration for the panel (position, size, etc.)
   */
  panelConfig?: PanelConfig;

  /**
   * Called when the projected document changes.
   * Use this to update plugin-specific state based on document changes.
   */
  onStateChange?: (document: Document) => TState | undefined;

  /**
   * Initialize plugin state when the plugin is first loaded.
   */
  initialize?: (document: Document | null) => TState;

  /**
   * Called when the plugin is being destroyed.
   * Use this for cleanup (subscriptions, timers, etc.)
   */
  destroy?: () => void;

  /**
   * CSS styles to inject for this plugin.
   * Can be a string of CSS or a URL to a stylesheet.
   */
  styles?: string;
}

/**
 * A sidebar item anchored to a document position.
 * Framework adapters extend this with rendering capabilities.
 */
export interface SidebarItem {
  /** Unique ID for this item (used as React key and for overlap resolution). */
  id: string;

  /** Display position this item anchors to. */
  anchorPos: number;

  /** Optional key into the anchorPositions Map (e.g. "comment-42", "revision-7"). */
  anchorKey?: string;

  /** Sort priority within items at the same anchor Y. Lower = first. Default: 0. */
  priority?: number;

  /** Temporary items (e.g. "add comment" input) skip entrance animation. */
  isTemporary?: boolean;

  /** Pre-computed Y position (scroll-container coords, pre-zoom). Overrides anchor resolution. */
  fixedY?: number;
}

/**
 * Context provided to plugins when computing sidebar items.
 */
export interface SidebarItemContext {
  document: Document | null;
  renderedDomContext: RenderedDomContext | null;
  /** Pre-computed Y positions from layout engine (keys like "comment-{id}"). */
  anchorPositions: Map<string, number>;
  zoom: number;
}
