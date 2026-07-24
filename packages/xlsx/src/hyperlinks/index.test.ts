import { describe, expect, it } from 'bun:test';
import type { DisplayList, HyperlinkRegion } from '../display-list/types';
import {
  hyperlinkAtCell,
  parseHyperlinkLocation,
  safeExternalHyperlink,
} from './index';

const link: HyperlinkRegion = {
  top: 1,
  left: 2,
  bottom: 3,
  right: 4,
  externalTarget: 'https://example.com/report',
  location: 'detail',
};

describe('hyperlinks', () => {
  it('finds a visible hyperlink by sheet address', () => {
    const displayList: DisplayList = {
      width: 100,
      height: 100,
      commands: [],
      hyperlinks: [link],
    };
    expect(hyperlinkAtCell(displayList, 2, 3)).toEqual(link);
    expect(hyperlinkAtCell(displayList, 0, 0)).toBeNull();
  });

  it('allows browser-safe external protocols and rejects script URLs', () => {
    expect(safeExternalHyperlink(link)).toBe('https://example.com/report#detail');
    expect(safeExternalHyperlink({ ...link, externalTarget: 'mailto:user@example.com' })).toBe(
      'mailto:user@example.com#detail'
    );
    expect(safeExternalHyperlink({ ...link, externalTarget: 'javascript:alert(1)' })).toBeNull();
    expect(safeExternalHyperlink({ ...link, externalTarget: '../relative' })).toBeNull();
  });

  it('parses local and quoted cross-sheet cell destinations', () => {
    expect(parseHyperlinkLocation('$B$4', 'Current')).toEqual({
      sheetName: 'Current',
      row: 3,
      col: 1,
    });
    expect(parseHyperlinkLocation("#'Other '' Sheet'!$XFD$1048576", 'Current')).toEqual({
      sheetName: "Other ' Sheet",
      row: 1_048_575,
      col: 16_383,
    });
    expect(parseHyperlinkLocation('NamedDestination', 'Current')).toBeNull();
    expect(parseHyperlinkLocation('XFE1', 'Current')).toBeNull();
  });
});
