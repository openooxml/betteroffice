/** A paragraph the legacy paragraph finder matched, with the handle the paragraph helpers take. */
export interface DocxParagraphMatch {
  paraId: string;
  match: string;
  before: string;
  after: string;
}

/** @internal */
export interface DocxFindParagraphsOptions {
  caseSensitive?: boolean;
  limit?: number;
}

/** @internal */
export function findParagraphs(
  reader: { storyIds(): string[]; paragraphs(story: string): Array<{ paraId: string; text: string }> },
  query: string,
  options?: DocxFindParagraphsOptions
): DocxParagraphMatch[] {
  if (!query) return [];
  const caseSensitive = options?.caseSensitive ?? false;
  const needle = caseSensitive ? query : query.toLowerCase();
  const limit = options?.limit ?? 20;
  const results: DocxParagraphMatch[] = [];
  for (const story of reader.storyIds().filter((id) => id === 'body' || id.startsWith('body:'))) {
    for (const paragraph of reader.paragraphs(story)) {
      if (results.length >= limit) return results;
      const haystack = caseSensitive ? paragraph.text : paragraph.text.toLowerCase();
      const offset = haystack.indexOf(needle);
      if (offset < 0 || haystack.indexOf(needle, offset + 1) >= 0) continue;
      results.push({
        paraId: paragraph.paraId,
        match: paragraph.text.slice(offset, offset + query.length),
        before: paragraph.text.slice(Math.max(0, offset - 40), offset),
        after: paragraph.text.slice(offset + query.length, offset + query.length + 40),
      });
    }
  }
  return results;
}
