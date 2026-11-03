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

const GUTTER_BORDER = 1;

const SLOT: CSSProperties = {
  position: 'absolute',
  top: 0,
  boxSizing: 'border-box',
  margin: 0,
  border: 'none',
  background: 'transparent',
  font: 'inherit',
  cursor: 'pointer',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  overflow: 'hidden',
};

function markStyle(item: XlsxPluginRowDecoration, scale: number): CSSProperties {
  const tone = TONES[item.tone ?? 'neutral'];
  if (item.kind === 'dot') {
    const size = 6 * scale;
    return { flex: 'none', width: size, height: size, maxWidth: '100%', maxHeight: '100%', borderRadius: '50%',
      background: tone.ink };
  }
  return {
    display: 'block',
    minWidth: 0,
    maxWidth: '100%',
    maxHeight: '100%',
    boxSizing: 'border-box',
    padding: `0 ${scale}px`,
    borderRadius: 3 * scale,
    background: tone.fill,
    color: tone.ink,
    fontSize: 8 * scale,
    lineHeight: `${10 * scale}px`,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  };
}

function moreStyle(scale: number, width: number, height: number): CSSProperties {
  return {
    position: 'absolute',
    top: 0,
    right: 0,
    maxWidth: width,
    maxHeight: height,
    boxSizing: 'border-box',
    padding: `0 ${scale}px`,
    borderRadius: 2 * scale,
    background: TONES.neutral.fill,
    color: TONES.neutral.ink,
    fontSize: 6 * scale,
    lineHeight: `${7 * scale}px`,
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    pointerEvents: 'none',
  };
}

/**
 * The row gutter at the left of the grid: plugin row decorations on the rows of the painted
 * frame, at its row offsets, so frozen, hidden and resized rows place them as the canvas does.
 * The shown items split the gutter's width, each a full-height button; a `+n` count sits in the
 * top right corner, no wider than one item, so everything stays inside the gutter. The count
 * lets the pointer through to the item under it; the row's title names the hidden items. The editor
 * renders it only while `rowGutter` has items and moves its grid right by
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
  const inner = ROW_GUTTER_WIDTH * zoom - GUTTER_BORDER;
  const rows: ReactElement[] = [];
  const frame = painted && painted.sheetId === index.sheetId ? painted.frame : null;
  const grid = frame?.grid;
  for (let local = 0; frame && grid && local < grid.rowOffsets.length - 1; local++) {
    const row = grid.rowIndices?.[local] ?? grid.startRow + local;
    const line = index.rows.get(row);
    const height = (grid.rowOffsets[local + 1] - grid.rowOffsets[local]) * scale;
    if (!line || !(height > 0) || grid.rowOffsets[local] >= frame.height) continue;
    const slot = inner / line.shown.length;
    rows.push(
      <div
        key={row}
        data-row={row}
        title={line.more > 0 ? line.tooltip : undefined}
        style={{
          position: 'absolute',
          left: 0,
          top: grid.rowOffsets[local] * scale,
          width: inner,
          height,
          overflow: 'hidden',
        }}
      >
        {line.shown.map(({ entry, label }, position) => (
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
            style={{ ...SLOT, left: position * slot, width: slot, height, padding: `0 ${scale}px` }}
          >
            <span aria-hidden style={markStyle(entry.item, scale)}>
              {entry.item.kind === 'badge' ? entry.item.text : null}
            </span>
          </button>
        ))}
        {line.more > 0 && (
          <span
            role="img"
            data-testid="xlsx-row-decoration-more"
            aria-label={line.tooltip}
            title={line.tooltip}
            style={moreStyle(scale, slot, height)}
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
          width: inner + GUTTER_BORDER,
          height: painted ? painted.frame.height * scale : 0,
          overflow: 'hidden',
          boxSizing: 'border-box',
          borderRight: `${GUTTER_BORDER}px solid #e0e0e0`,
          background: '#f8f9fa',
        }}
      >
        {rows}
      </div>
    </div>
  );
});
