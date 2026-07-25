import { describe, expect, it } from 'bun:test';
import { renderToStaticMarkup } from 'react-dom/server';
import type { GridMeta, MergedRange } from '@betteroffice/xlsx';
import { colorForClientId, type AwarenessPeer } from '@betteroffice/xlsx/collaboration';
import { expandRangeToMergedCells, PresenceStrip, RemoteSelections } from './Presence';

const grid: GridMeta = {
  startRow: 0,
  startCol: 0,
  rowOffsets: [0, 10, 20],
  colOffsets: [0, 10, 20],
};

function peer(clientId: number, color = '#0B57D0'): AwarenessPeer {
  return {
    clientId,
    clock: 1,
    user: { name: `Peer ${clientId}`, color },
    cursor: {
      sheet: 'sheet:0',
      anchor: { row: 0, col: 0 },
      head: { row: 0, col: 0 },
    },
    lastSeen: 1,
    cursorMovedAt: 0,
  };
}

describe('xlsx presence rendering', () => {
  it('expands local and remote geometry through intersecting merged cells', () => {
    const chained: MergedRange[] = [
      { start: { row: 2, col: 0 }, end: { row: 3, col: 0 } },
      { start: { row: 1, col: 0 }, end: { row: 1, col: 1 } },
    ];
    expect(expandRangeToMergedCells({ top: 1, left: 1, bottom: 2, right: 2 }, chained)).toEqual({
      top: 1,
      left: 0,
      bottom: 3,
      right: 2,
    });

    const markup = renderToStaticMarkup(
      <RemoteSelections
        peers={[peer(2)]}
        grid={grid}
        sheetIds={['sheet:0']}
        activeSheet={0}
        zoom={1}
        mergedRanges={[{ start: { row: 0, col: 0 }, end: { row: 1, col: 1 } }]}
      />
    );
    expect(markup).toContain('width:20px');
    expect(markup).toContain('height:20px');
  });

  it('uses palette colors for unsafe peer input and specific CSS properties', () => {
    const unsafe = peer(42, 'url(https://example.invalid/peer)');
    const strip = renderToStaticMarkup(
      <PresenceStrip
        peers={[unsafe]}
        sheetIds={['sheet:0']}
        sheetNames={['Sheet 1']}
        activeSheet={0}
      />
    );
    const selection = renderToStaticMarkup(
      <RemoteSelections
        peers={[unsafe]}
        grid={grid}
        sheetIds={['sheet:0']}
        activeSheet={0}
        zoom={1}
        mergedRanges={[]}
      />
    );

    expect(strip).toContain(`background-color:${colorForClientId(42)}`);
    expect(selection).toContain(`border-color:${colorForClientId(42)}`);
    expect(strip).not.toContain('url(');
    expect(selection).not.toContain('url(');
  });

  it('caps peer visuals and preserves overflow counts', () => {
    const peers = Array.from({ length: 100 }, (_, index) => peer(index + 1));
    const strip = renderToStaticMarkup(
      <PresenceStrip
        peers={peers}
        sheetIds={['sheet:0']}
        sheetNames={['Sheet 1']}
        activeSheet={0}
      />
    );
    const selections = renderToStaticMarkup(
      <RemoteSelections
        peers={peers}
        grid={grid}
        sheetIds={['sheet:0']}
        activeSheet={0}
        zoom={1}
        mergedRanges={[]}
      />
    );

    expect(strip.match(/data-testid="xlsx-presence-chip-\d+"/g)).toHaveLength(8);
    expect(strip).toContain('+92');
    expect(selections.match(/data-testid="xlsx-remote-selection-\d+"/g)).toHaveLength(32);
    expect(selections).toContain('+68');
  });
});
