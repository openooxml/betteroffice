/** React plugin host for the Yrs-backed editor. */

import {
  cloneElement,
  forwardRef,
  useCallback,
  useEffect,
  useImperativeHandle,
  useMemo,
  useRef,
  useState,
} from 'react';
import { injectStyles } from '@betteroffice/docx';
import type { Document } from '@betteroffice/docx/types/document';
import type { DocxEditorRef } from '../components/DocxEditor';
import type {
  EditorPlugin,
  PanelConfig,
  PluginHostProps,
  PluginHostRef,
  ReactSidebarItem,
  RenderedDomContext,
} from './types';

const DEFAULT_PANEL_CONFIG: Required<PanelConfig> = {
  position: 'right',
  defaultSize: 280,
  minSize: 200,
  maxSize: 500,
  resizable: true,
  collapsible: true,
  defaultCollapsed: false,
};

const PLUGIN_HOST_STYLES = `
.plugin-host { display:flex; width:100%; height:100%; overflow:visible; position:relative; }
.plugin-host-editor { flex:1; display:flex; flex-direction:column; min-width:0; overflow:visible; }
.plugin-panels-left,.plugin-panels-bottom { display:flex; flex-direction:column; flex-shrink:0; background:#f8f9fa; border-color:#e9ecef; }
.plugin-panels-left { border-right:1px solid #e9ecef; }
.plugin-panels-bottom { border-top:1px solid #e9ecef; }
.plugin-panel { position:relative; display:flex; flex-direction:column; overflow:hidden; }
.plugin-panel.collapsed { overflow:visible; }
.plugin-panel-toggle { display:flex; align-items:center; gap:4px; padding:6px 8px; background:transparent; border:0; cursor:pointer; font-size:12px; color:#6c757d; }
.plugin-panel-toggle:hover { background:#e9ecef; color:#495057; }
.plugin-panel-content { flex:1; overflow:auto; }
.plugin-panel-in-viewport { position:absolute; top:0; width:220px; pointer-events:auto; z-index:10; overflow:visible; }
.plugin-panel-in-viewport-content { overflow:visible; position:relative; }
.plugin-overlays-container,.plugin-overlay { position:absolute; inset:0; pointer-events:none; overflow:visible; }
.plugin-overlays-container { z-index:5; }
`;

type InjectedEditorProps = {
  ref?: React.Ref<DocxEditorRef>;
  document?: Document | null;
  onChange?: (document: Document) => void;
  pluginOverlays?: React.ReactNode;
  pluginSidebarItems?: ReactSidebarItem[];
  pluginRenderedDomContext?: RenderedDomContext | null;
  onRenderedDomContextReady?: (context: RenderedDomContext) => void;
};

function assignRef(ref: React.Ref<DocxEditorRef> | undefined, value: DocxEditorRef | null): void {
  if (typeof ref === 'function') ref(value);
  else if (ref) (ref as React.MutableRefObject<DocxEditorRef | null>).current = value;
}

export const PluginHost = forwardRef<PluginHostRef, PluginHostProps>(function PluginHost(
  { plugins, children, className = '' },
  ref
) {
  const child = children as React.ReactElement<InjectedEditorProps>;
  const childPropsRef = useRef(child.props);
  childPropsRef.current = child.props;
  const editorRef = useRef<DocxEditorRef | null>(null);
  const [document, setDocument] = useState<Document | null>(child.props.document ?? null);
  const documentRef = useRef(document);
  documentRef.current = document;
  const [renderedDomContext, setRenderedDomContext] = useState<RenderedDomContext | null>(null);
  const [pluginStates, setPluginStates] = useState<Map<string, unknown>>(new Map());

  useEffect(() => setDocument(child.props.document ?? null), [child.props.document]);

  useEffect(() => {
    const states = new Map<string, unknown>();
    for (const plugin of plugins) {
      if (plugin.initialize) states.set(plugin.id, plugin.initialize(documentRef.current));
    }
    setPluginStates(states);
    return () => {
      for (const plugin of plugins) plugin.destroy?.();
    };
  }, [plugins]);

  useEffect(() => {
    const cleanups = plugins
      .filter((plugin) => plugin.styles)
      .map((plugin) => injectStyles(plugin.id, plugin.styles!));
    return () => cleanups.forEach((cleanup) => cleanup());
  }, [plugins]);

  useEffect(() => injectStyles('plugin-host-base', PLUGIN_HOST_STYLES), []);

  const updatePluginStates = useCallback(
    (nextDocument: Document) => {
      setPluginStates((previous) => {
        const next = new Map(previous);
        let changed = false;
        for (const plugin of plugins) {
          const state = plugin.onStateChange?.(nextDocument);
          if (state !== undefined) {
            next.set(plugin.id, state);
            changed = true;
          }
        }
        return changed ? next : previous;
      });
    },
    [plugins]
  );

  const getPluginState = useCallback(
    <T,>(pluginId: string): T | undefined => pluginStates.get(pluginId) as T | undefined,
    [pluginStates]
  );
  const setPluginState = useCallback(<T,>(pluginId: string, state: T) => {
    setPluginStates((previous) => new Map(previous).set(pluginId, state));
  }, []);
  const refreshPluginStates = useCallback(() => {
    if (documentRef.current) updatePluginStates(documentRef.current);
  }, [updatePluginStates]);
  const scrollToPosition = useCallback((position: number) => {
    editorRef.current?.scrollToPosition(position);
    editorRef.current?.focus();
  }, []);
  const selectRange = useCallback((from: number, to: number) => {
    editorRef.current?.highlightRange(from, to);
    editorRef.current?.focus();
  }, []);

  useImperativeHandle(
    ref,
    () => ({
      getPluginState,
      setPluginState,
      getDocument: () => documentRef.current,
      refreshPluginStates,
    }),
    [getPluginState, refreshPluginStates, setPluginState]
  );

  const [collapsedPanels, setCollapsedPanels] = useState<Set<string>>(
    () =>
      new Set(
        plugins
          .filter((plugin) => ({ ...DEFAULT_PANEL_CONFIG, ...plugin.panelConfig }).defaultCollapsed)
          .map((plugin) => plugin.id)
      )
  );
  const panelSizes = useMemo(
    () =>
      new Map(
        plugins.map((plugin) => [
          plugin.id,
          ({ ...DEFAULT_PANEL_CONFIG, ...plugin.panelConfig }).defaultSize,
        ])
      ),
    [plugins]
  );
  const togglePanelCollapsed = useCallback((pluginId: string) => {
    setCollapsedPanels((previous) => {
      const next = new Set(previous);
      if (next.has(pluginId)) next.delete(pluginId);
      else next.add(pluginId);
      return next;
    });
  }, []);

  const [panelLeftPosition, setPanelLeftPosition] = useState<number | null>(null);
  useEffect(() => {
    if (!renderedDomContext) {
      setPanelLeftPosition(null);
      return;
    }
    const calculate = () => {
      const bounds = renderedDomContext.getPageBounds(0);
      const offset = renderedDomContext.getContainerOffset();
      setPanelLeftPosition(bounds ? offset.x + bounds.x + bounds.width + 5 : null);
    };
    calculate();
    const observer = new ResizeObserver(() => requestAnimationFrame(calculate));
    observer.observe(renderedDomContext.pagesContainer);
    window.addEventListener('resize', calculate);
    return () => {
      observer.disconnect();
      window.removeEventListener('resize', calculate);
    };
  }, [renderedDomContext]);

  const panelProps = (plugin: EditorPlugin, width: number) => ({
    document,
    scrollToPosition,
    selectRange,
    pluginState: pluginStates.get(plugin.id),
    panelWidth: width,
    renderedDomContext,
  });

  const renderPanel = (plugin: EditorPlugin) => {
    if (!plugin.Panel) return null;
    const config = { ...DEFAULT_PANEL_CONFIG, ...plugin.panelConfig };
    const collapsed = collapsedPanels.has(plugin.id);
    const size = panelSizes.get(plugin.id) ?? config.defaultSize;
    const Panel = plugin.Panel;
    return (
      <div
        key={plugin.id}
        className={`plugin-panel plugin-panel-${config.position} ${collapsed ? 'collapsed' : ''}`}
        style={{
          [config.position === 'bottom' ? 'height' : 'width']: collapsed ? 32 : size,
          minWidth: config.position !== 'bottom' ? (collapsed ? 32 : config.minSize) : undefined,
          maxWidth: config.position !== 'bottom' ? config.maxSize : undefined,
          minHeight: config.position === 'bottom' ? (collapsed ? 32 : config.minSize) : undefined,
          maxHeight: config.position === 'bottom' ? config.maxSize : undefined,
        }}
        data-plugin-id={plugin.id}
      >
        {config.collapsible && (
          <button
            className="plugin-panel-toggle"
            onClick={() => togglePanelCollapsed(plugin.id)}
            aria-label={collapsed ? `Show ${plugin.name}` : `Hide ${plugin.name}`}
          >
            {collapsed ? '›' : '‹'} {collapsed ? plugin.name : ''}
          </button>
        )}
        {!collapsed && (
          <div className="plugin-panel-content">
            <Panel {...panelProps(plugin, size)} />
          </div>
        )}
      </div>
    );
  };

  const pluginOverlays = useMemo(() => {
    const overlays: React.ReactNode[] = [];
    if (renderedDomContext) {
      for (const plugin of plugins) {
        if (plugin.renderOverlay) {
          overlays.push(
            <div key={`overlay-${plugin.id}`} className="plugin-overlay" data-plugin-id={plugin.id}>
              {plugin.renderOverlay(renderedDomContext, pluginStates.get(plugin.id))}
            </div>
          );
        }
      }
    }
    for (const plugin of plugins) {
      if (!plugin.Panel || (plugin.panelConfig?.position ?? 'right') !== 'right') continue;
      const config = { ...DEFAULT_PANEL_CONFIG, ...plugin.panelConfig };
      const collapsed = collapsedPanels.has(plugin.id);
      const size = panelSizes.get(plugin.id) ?? config.defaultSize;
      const Panel = plugin.Panel;
      overlays.push(
        <div
          key={`panel-overlay-${plugin.id}`}
          className={`plugin-panel-in-viewport ${collapsed ? 'collapsed' : ''}`}
          style={{ width: collapsed ? 32 : size, left: panelLeftPosition ?? 'calc(50% + 428px)' }}
          data-plugin-id={plugin.id}
        >
          {config.collapsible && (
            <button
              className="plugin-panel-toggle"
              onClick={() => togglePanelCollapsed(plugin.id)}
              aria-label={collapsed ? `Show ${plugin.name}` : `Hide ${plugin.name}`}
            >
              {collapsed ? '‹' : '›'}
            </button>
          )}
          {!collapsed && (
            <div className="plugin-panel-in-viewport-content">
              <Panel {...panelProps(plugin, size)} />
            </div>
          )}
        </div>
      );
    }
    return overlays.length > 0 ? overlays : null;
  }, [collapsedPanels, document, panelLeftPosition, panelSizes, pluginStates, plugins, renderedDomContext, scrollToPosition, selectRange, togglePanelCollapsed]);

  const pluginSidebarItems = useMemo(() => {
    const items: ReactSidebarItem[] = [];
    for (const plugin of plugins) {
      if (!plugin.getSidebarItems) continue;
      items.push(
        ...plugin.getSidebarItems(pluginStates.get(plugin.id), {
          document,
          renderedDomContext,
          anchorPositions: new Map(),
          zoom: renderedDomContext?.zoom ?? 1,
        })
      );
    }
    return items;
  }, [document, pluginStates, plugins, renderedDomContext]);

  const handleRenderedDomContextReady = useCallback((context: RenderedDomContext) => {
    setRenderedDomContext(context);
    childPropsRef.current.onRenderedDomContextReady?.(context);
  }, []);
  const handleDocumentChange = useCallback(
    (nextDocument: Document) => {
      setDocument(nextDocument);
      updatePluginStates(nextDocument);
      childPropsRef.current.onChange?.(nextDocument);
    },
    [updatePluginStates]
  );
  const handleEditorRef = useCallback((value: DocxEditorRef | null) => {
    editorRef.current = value;
    assignRef(childPropsRef.current.ref, value);
  }, []);

  const editorElement = cloneElement(child, {
    ref: handleEditorRef,
    onChange: handleDocumentChange,
    pluginOverlays,
    pluginSidebarItems,
    pluginRenderedDomContext: renderedDomContext,
    onRenderedDomContextReady: handleRenderedDomContextReady,
  });

  const left = plugins.filter(
    (plugin) => plugin.Panel && (plugin.panelConfig?.position ?? 'right') === 'left'
  );
  const bottom = plugins.filter((plugin) => plugin.Panel && plugin.panelConfig?.position === 'bottom');

  return (
    <div className={`plugin-host ${className}`}>
      {left.length > 0 && <div className="plugin-panels-left">{left.map(renderPanel)}</div>}
      <div className="plugin-host-editor">
        {editorElement}
        {bottom.length > 0 && (
          <div className="plugin-panels-bottom">{bottom.map(renderPanel)}</div>
        )}
      </div>
    </div>
  );
});

export { PLUGIN_HOST_STYLES };
export default PluginHost;
