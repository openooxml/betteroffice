import type { YrsLoc, YrsSession, YrsStoryRange } from './index';
import { findAllMatches, type FindOptions } from '../utils/findReplace';

/** @internal */
export interface DocxFindDisplayMatch {
  paragraphIndex: number;
  contentIndex: number;
  startOffset: number;
  endOffset: number;
  text: string;
  displayFrom: number;
  displayTo: number;
  yrsRange: YrsStoryRange;
}

/** @internal */
export function findBodyMatches(
  reader: Pick<YrsSession, 'paragraphs'>,
  locToDisplayPosition: (loc: YrsLoc) => number | null,
  searchText: string,
  options: FindOptions
): DocxFindDisplayMatch[] {
  const matches: DocxFindDisplayMatch[] = [];
  const paragraphs = reader.paragraphs('body');
  for (let paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex += 1) {
    const paragraph = paragraphs[paragraphIndex];
    if (!paragraph.text) continue;
    for (const match of findAllMatches(paragraph.text, searchText, options)) {
      const startLoc = { story: 'body', paraId: paragraph.paraId, offset: match.start };
      const endLoc = { story: 'body', paraId: paragraph.paraId, offset: match.end };
      const displayFrom = locToDisplayPosition(startLoc);
      const displayTo = locToDisplayPosition(endLoc);
      if (displayFrom == null || displayTo == null || displayFrom >= displayTo) continue;
      matches.push({
        paragraphIndex,
        contentIndex: 0,
        startOffset: match.start,
        endOffset: match.end,
        text: paragraph.text.slice(match.start, match.end),
        displayFrom,
        displayTo,
        yrsRange: {
          story: 'body',
          start: { paraId: paragraph.paraId, offset: match.start },
          end: { paraId: paragraph.paraId, offset: match.end },
        },
      });
    }
  }
  return matches;
}

