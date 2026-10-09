import { useEffect, useId, useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent, RefObject } from 'react';
import { PLUGIN_CHROME } from '../commands/pluginEvents';
import { useTranslation } from '../i18n';
import type { XlsxPluginActivation, XlsxPluginHost } from './createXlsxPluginHost';
import { PluginRenderScope } from './PluginRenderScope';

export type DockPlacement = 'left' | 'right' | 'bottom';

const DEFAULT_SIZE: Record<DockPlacement, number> = { left: 280, right: 280, bottom: 200 };
const MIN_SIZE: Record<DockPlacement, number> = { left: 160, right: 160, bottom: 120 };
/** Below this workspace width side docks show their tabs and open panels as drawers. */
export const NARROW_EDITOR_WIDTH = 640;

const colors = {
  surface: '#ffffff',
  border: '#e2e8f0',
  text: '#172033',
  muted: '#647087',
  selected: '#e6f4ea',
  primary: '#217346',
};

/** The panel size a dock gets: its preference, clamped to at most 40% of the workspace. */
export function dockSize(
  placement: DockPlacement,
  preferred: number | undefined,
  available: { width: number; height: number }
): number {
  const room = placement === 'bottom' ? available.height : available.width;
  const limit = Math.max(MIN_SIZE[placement], Math.floor(room * 0.4));
  return Math.max(MIN_SIZE[placement], Math.min(preferred ?? DEFAULT_SIZE[placement], limit));
}

function useElementSize(
  active = true
): [RefObject<HTMLDivElement | null>, { width: number; height: number }] {
  const ref = useRef<HTMLDivElement>(null);
  const [size, setSize] = useState({ width: 0, height: 0 });
  useEffect(() => {
    const element = ref.current;
    if (!active || !element) return;
    const measure = () =>
      setSize((previous) =>
        previous.width === element.clientWidth && previous.height === element.clientHeight
          ? previous
          : { width: element.clientWidth, height: element.clientHeight }
      );
    measure();
    if (typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(measure);
    observer.observe(element);
    return () => observer.disconnect();
  }, [active]);
  return [ref, size];
}

/** Tracks the size of the workspace the docks share with the grid while plugins are installed. */
export const useDockArea = useElementSize;

function PanelBody({
  host,
  activation,
  id,
  labelledBy,
}: {
  host: XlsxPluginHost;
  activation: XlsxPluginActivation;
  id: string;
  labelledBy: string;
}) {
  const [ref, size] = useElementSize();
  const Render = activation.plugin.panel!.render;
  return (
    <div
      ref={ref}
      id={id}
      role="tabpanel"
      aria-labelledby={labelledBy}
      className="xlsx-plugin-dock__panel"
      style={{ flex: 1, minHeight: 0, minWidth: 0, overflow: 'auto' }}
    >
      <PluginRenderScope host={host} activation={activation}>
        <Render context={activation.context} width={size.width} height={size.height} />
      </PluginRenderScope>
    </div>
  );
}

const tabStyle = (current: boolean): CSSProperties => ({
  maxWidth: 160,
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap',
  padding: '4px 8px',
  border: 0,
  borderRadius: 4,
  background: current ? colors.selected : 'transparent',
  color: current ? colors.primary : colors.muted,
  font: '12px ui-sans-serif, system-ui, sans-serif',
  cursor: 'pointer',
});

const iconButtonStyle: CSSProperties = {
  flex: '0 0 auto',
  padding: '0 6px',
  border: 0,
  borderRadius: 4,
  background: 'transparent',
  color: colors.muted,
  font: '13px ui-sans-serif, system-ui, sans-serif',
  cursor: 'pointer',
};

/**
 * One dock: accessible tabs over the panels of every plugin placed there. In a narrow workspace
 * a side dock shows only its tabs, and a tab opens its panel as a drawer over the grid that
 * Escape closes, returning focus to the tab.
 */
export function PluginDock({
  host,
  placement,
  activations,
  available,
}: {
  host: XlsxPluginHost;
  placement: DockPlacement;
  activations: readonly XlsxPluginActivation[];
  available: { width: number; height: number };
}) {
  const { t } = useTranslation();
  const baseId = useId();
  const [selected, setSelected] = useState<string | null>(null);
  const [collapsed, setCollapsed] = useState<Record<string, boolean>>({});
  const [drawer, setDrawer] = useState(false);
  const tabsRef = useRef<HTMLDivElement>(null);
  const drawerRef = useRef<HTMLDivElement>(null);
  const narrow =
    placement !== 'bottom' && available.width > 0 && available.width < NARROW_EDITOR_WIDTH;
  const drawerOpen = narrow && drawer && activations.length > 0;
  useEffect(() => {
    if (drawerOpen) drawerRef.current?.focus();
  }, [drawerOpen]);
  if (activations.length === 0) return null;

  const active = activations.find((entry) => entry.pluginId === selected) ?? activations[0];
  const panel = active.plugin.panel!;
  const isCollapsed = narrow
    ? !drawerOpen
    : collapsed[active.pluginId] ?? panel.defaultCollapsed ?? false;
  const size = dockSize(placement, panel.preferredSize, available);
  const tabId = (pluginId: string) => `${baseId}-tab-${pluginId}`;
  const panelId = `${baseId}-panel`;
  const focusTab = (pluginId: string) =>
    tabsRef.current?.querySelector<HTMLElement>(`[data-plugin-tab="${pluginId}"]`)?.focus();

  const choose = (pluginId: string) => {
    if (narrow) setDrawer(!(drawerOpen && pluginId === active.pluginId));
    else if (pluginId === active.pluginId && isCollapsed) {
      setCollapsed((previous) => ({ ...previous, [pluginId]: false }));
    }
    setSelected(pluginId);
  };

  const closeDrawer = () => {
    setDrawer(false);
    focusTab(active.pluginId);
  };

  const onKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    const index = activations.findIndex((entry) => entry.pluginId === active.pluginId);
    const last = activations.length - 1;
    const next =
      event.key === 'ArrowRight' || event.key === 'ArrowDown'
        ? index === last
          ? 0
          : index + 1
        : event.key === 'ArrowLeft' || event.key === 'ArrowUp'
        ? index === 0
          ? last
          : index - 1
        : event.key === 'Home'
        ? 0
        : event.key === 'End'
        ? last
        : null;
    if (next === null) return;
    event.preventDefault();
    const pluginId = activations[next].pluginId;
    setSelected(pluginId);
    focusTab(pluginId);
  };

  const toggleLabel = t(isCollapsed ? 'plugins.expandPanel' : 'plugins.collapsePanel', {
    title: panel.title,
  });
  const bottom = placement === 'bottom';
  const vertical = narrow || (!bottom && isCollapsed);
  const body = (
    <PanelBody
      key={active.key}
      host={host}
      activation={active}
      id={panelId}
      labelledBy={tabId(active.pluginId)}
    />
  );
  return (
    <section
      {...PLUGIN_CHROME}
      className={`xlsx-plugin-dock xlsx-plugin-dock--${placement}`}
      data-testid={`plugin-dock-${placement}`}
      aria-label={t('plugins.panelTabs')}
      style={{
        position: 'relative',
        display: 'flex',
        flexDirection: 'column',
        flexShrink: 0,
        minWidth: 0,
        minHeight: 0,
        background: colors.surface,
        color: colors.text,
        ...(bottom
          ? { height: isCollapsed ? undefined : size, borderTop: `1px solid ${colors.border}` }
          : {
              width: isCollapsed || narrow ? undefined : size,
              maxWidth: isCollapsed || narrow ? 96 : undefined,
              [placement === 'left' ? 'borderRight' : 'borderLeft']: `1px solid ${colors.border}`,
            }),
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: 4, padding: 4, flexWrap: 'wrap' }}>
        <div
          ref={tabsRef}
          role="tablist"
          aria-label={t('plugins.panelTabs')}
          aria-orientation={vertical ? 'vertical' : 'horizontal'}
          onKeyDown={onKeyDown}
          style={{
            display: 'flex',
            flexDirection: vertical ? 'column' : 'row',
            gap: 2,
            flex: 1,
            minWidth: 0,
            overflow: 'hidden',
          }}
        >
          {activations.map((entry) => {
            const current = entry.pluginId === active.pluginId;
            return (
              <button
                key={entry.pluginId}
                type="button"
                role="tab"
                id={tabId(entry.pluginId)}
                data-plugin-tab={entry.pluginId}
                aria-selected={current}
                aria-controls={current && !isCollapsed ? panelId : undefined}
                aria-expanded={narrow ? current && drawerOpen : undefined}
                tabIndex={current ? 0 : -1}
                title={entry.plugin.panel!.title}
                onClick={() => choose(entry.pluginId)}
                style={tabStyle(current)}
              >
                {entry.plugin.panel!.title}
              </button>
            );
          })}
        </div>
        {!narrow && (
          <button
            type="button"
            aria-expanded={!isCollapsed}
            aria-label={toggleLabel}
            title={toggleLabel}
            onClick={() =>
              setCollapsed((previous) => ({ ...previous, [active.pluginId]: !isCollapsed }))
            }
            style={iconButtonStyle}
          >
            {isCollapsed ? '+' : '−'}
          </button>
        )}
      </div>
      {!isCollapsed && !narrow && body}
      {drawerOpen && (
        <div
          ref={drawerRef}
          tabIndex={-1}
          data-testid={`plugin-drawer-${placement}`}
          onKeyDown={(event) => {
            if (event.key !== 'Escape') return;
            event.stopPropagation();
            closeDrawer();
          }}
          style={{
            position: 'absolute',
            top: 0,
            bottom: 0,
            [placement === 'left' ? 'left' : 'right']: '100%',
            width: Math.min(size, Math.max(MIN_SIZE[placement], available.width - 64)),
            zIndex: 60,
            display: 'flex',
            flexDirection: 'column',
            background: colors.surface,
            boxShadow: '0 4px 16px rgba(15, 23, 42, 0.18)',
            outline: 'none',
          }}
        >
          <div style={{ display: 'flex', justifyContent: 'flex-end', padding: 4 }}>
            <button
              type="button"
              aria-label={toggleLabel}
              title={toggleLabel}
              onClick={closeDrawer}
              style={iconButtonStyle}
            >
              ×
            </button>
          </div>
          {body}
        </div>
      )}
    </section>
  );
}
