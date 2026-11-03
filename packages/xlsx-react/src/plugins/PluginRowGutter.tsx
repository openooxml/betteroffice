import { memo, useSyncExternalStore, type CSSProperties, type KeyboardEvent, type ReactElement } from 'react';
import { PLUGIN_CHROME } from '../commands/pluginEvents';
import { useTranslation } from '../i18n';
import type { XlsxPluginHost } from './createXlsxPluginHost';
import { ROW_GUTTER_WIDTH, type RowDecorationEntry, type RowGutterIndex } from './rowDecorations';
import type { XlsxPluginRowDecoration } from './types';
import type { PaintedGridStore } from './useXlsxPluginHost';

type Tone = NonNullable<XlsxPluginRowDecoration['tone']>;

const TONES: Record<Tone, { ink: string; fill: string }> = {
  neutral: { ink: '#3c4043', fill: '#e8eaed' },
  info: { ink: '#174ea6', fill: '#d2e3fc' },
  success: { ink: '#0d652d', fill: '#ceead6' },
  warning: { ink: '#7a4f01', fill: '#feefc3' },
  danger: { ink: '#a50e0e', fill: '#fad2cf' },
};

const BASE: CSSProperties = {
  flex: 'none',
  boxSizing: 'border-box',
  margin: 0,
  padding: 0,
  border: 'none',
  font: 'inherit',
};

function badgeStyle(tone: Tone, scale: number): CSSProperties {
  return {
    ...BASE,
    maxWidth: (ROW_GUTTER_WIDTH - 2) * scale,
    padding: `0 ${scale}px`,
    borderRadius: 3 * scale,
    background: TONES[tone].fill,
    color: TONES[tone].ink,
    fontSize: 8 * scale,
    lineHeight: `${10 * scale}px`,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
  };
}

function itemStyle(item: XlsxPluginRowDecoration, scale: number): CSSProperties {
  const tone = item.tone ?? 'neutral';
  if (item.kind === 'badge') return { ...badgeStyle(tone, scale), cursor: 'pointer' };
  const size = 6 * scale;
  return { ...BASE, width: size, height: size, borderRadius: '50%', background: TONES[tone].ink, cursor: 'pointer' };
}

/**
 * The row gutter at the left of the grid: plugin row decorations on the rows of the painted
 * frame, at its row offsets, so frozen, hidden and resized rows place them as the canvas does.
 * The editor renders it only while `rowGutter` has items and moves its grid right by
 * `ROW_GUTTER_WIDTH * zoom` meanwhile.
 */
export const PluginRowGutter = memo(function PluginRowGutter({
  host,
  index,
  painted: store,
  zoom,
}: {
  host: XlsxPluginHost;
  index: RowGutterIndex;
  painted: PaintedGridStore;
  zoom: number;
}) {
  const { t } = useTranslation();
  const painted = useSyncExternalStore(store.subscribe, store.get, () => null);
  const scale = painted?.zoom ?? zoom;
  const onKeyDown = (entry: RowDecorationEntry) => (event: KeyboardEvent) => {
    if (event.key !== 'Enter' && event.key !== ' ') return;
    event.preventDefault();
    event.stopPropagation();
    host.activateRowDecoration(entry);
  };
  const rows: ReactElement[] = [];
  const frame = painted && painted.sheetId === index.sheetId ? painted.frame : null;
  const grid = frame?.grid;
  for (let local = 0; frame && grid && local < grid.rowOffsets.length - 1; local++) {
    const row = grid.rowIndices?.[local] ?? grid.startRow + local;
    const line = index.rows.get(row);
    const height = (grid.rowOffsets[local + 1] - grid.rowOffsets[local]) * scale;
    if (!line || !(height > 0) || grid.rowOffsets[local] >= frame.height) continue;
    rows.push(
      <div
        key={row}
        data-row={row}
        style={{
          position: 'absolute',
          left: 0,
          top: grid.rowOffsets[local] * scale,
          width: '100%',
          height,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          gap: scale,
          overflow: 'hidden',
        }}
      >
        {line.shown.map(({ entry, label }) => (
          <button
            key={`${entry.key}\u0000${entry.item.id}`}
            type="button"
            data-plugin-id={entry.pluginId}
            data-decoration-id={entry.item.id}
            aria-label={label}
            title={entry.item.tooltip}
            onClick={() => host.activateRowDecoration(entry)}
            onKeyDown={onKeyDown(entry)}
            onKeyUp={(event) => { if (event.key === ' ') event.preventDefault(); }}
            style={itemStyle(entry.item, scale)}
          >
            {entry.item.kind === 'badge' ? entry.item.text : null}
          </button>
        ))}
        {line.more > 0 && (
          <span
            role="img"
            data-testid="xlsx-row-decoration-more"
            aria-label={line.tooltip}
            title={line.tooltip}
            style={badgeStyle('neutral', scale)}
          >
            {`+${line.more}`}
          </span>
        )}
      </div>
    );
  }
  return (
    <div style={{ position: 'relative', flex: 'none', width: 0, height: 0 }}>
      <div
        {...PLUGIN_CHROME}
        role="region"
        aria-label={t('plugins.rowGutter')}
        data-testid="xlsx-row-gutter"
        style={{
          position: 'absolute',
          top: 0,
          left: 0,
          width: ROW_GUTTER_WIDTH * zoom,
          height: painted ? painted.frame.height * scale : 0,
          overflow: 'hidden',
          boxSizing: 'border-box',
          borderRight: '1px solid #e0e0e0',
          background: '#f8f9fa',
        }}
      >
        {rows}
      </div>
    </div>
  );
});
