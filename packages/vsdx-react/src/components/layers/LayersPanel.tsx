import type { CSSProperties } from 'react';
import type { TFunction } from '@betteroffice/vsdx-i18n';
import type { PageLayer } from '@betteroffice/vsdx';

export interface LayersPanelProps {
  layers: readonly PageLayer[];
  collapsed: boolean;
  onToggleCollapsed: () => void;
  onToggleLayer: (index: number, visible: boolean) => void;
  t: TFunction;
  className?: string;
}

const styles: Record<string, CSSProperties> = {
  root: { display: 'flex', flexDirection: 'column', minHeight: 0, maxHeight: '42%', background: '#fff', color: '#242424', font: '400 13px ui-sans-serif, system-ui, sans-serif', borderTop: '1px solid #e0e0e0', boxSizing: 'border-box' },
  header: { display: 'flex', alignItems: 'center', justifyContent: 'space-between', minHeight: 44, padding: '0 10px 0 14px', borderBottom: '1px solid #e5e5e5', fontWeight: 600, fontSize: 14 },
  toggle: { appearance: 'none', display: 'grid', placeItems: 'center', width: 28, height: 28, padding: 0, border: 0, borderRadius: 4, background: 'transparent', color: '#424242', cursor: 'pointer' },
  list: { margin: 0, padding: '6px 6px 10px', overflowY: 'auto', listStyle: 'none' },
  row: { display: 'flex', alignItems: 'center', minHeight: 30, padding: '0 8px', borderRadius: 3 },
  label: { display: 'flex', alignItems: 'center', gap: 8, width: '100%', cursor: 'pointer' },
  name: { display: 'block', overflow: 'hidden', whiteSpace: 'nowrap', textOverflow: 'ellipsis' },
  empty: { margin: '12px 14px', color: '#616161', textAlign: 'center' },
};

export function LayersPanel({ layers, collapsed, onToggleCollapsed, onToggleLayer, t, className }: LayersPanelProps) {
  return (
    <aside className={className} style={styles.root} aria-label={t('layersPanel.title')}>
      <header style={styles.header}>
        <span>{t('layersPanel.title')}</span>
        <button
          type="button"
          aria-label={collapsed ? t('layersPanel.expand') : t('layersPanel.collapse')}
          aria-expanded={!collapsed}
          title={collapsed ? t('layersPanel.expand') : t('layersPanel.collapse')}
          onClick={onToggleCollapsed}
          style={styles.toggle}
        >
          <svg aria-hidden="true" width="16" height="16" viewBox="0 0 16 16">
            <path d={collapsed ? 'M 6 3 L 11 8 L 6 13' : 'M 10 3 L 5 8 L 10 13'} fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        </button>
      </header>
      {!collapsed && (layers.length === 0 ? <p style={styles.empty}>{t('layersPanel.empty')}</p> : (
        <ul style={styles.list}>
          {layers.map((layer) => {
            const name = layer.name || t('layersPanel.unnamedLayer', { index: layer.index });
            return (
              <li key={layer.index} style={styles.row}>
                <label style={styles.label}>
                  <input
                    type="checkbox"
                    checked={layer.visible}
                    onChange={(event) => onToggleLayer(layer.index, event.target.checked)}
                    aria-label={layer.visible ? t('layersPanel.hideLayer', { name }) : t('layersPanel.showLayer', { name })}
                  />
                  <span title={name} style={styles.name}>{name}</span>
                </label>
              </li>
            );
          })}
        </ul>
      ))}
    </aside>
  );
}
