import type { YrsSession, YrsStoryRange } from '@betteroffice/docx/yrs';

export function overlapsTextRevision(session: YrsSession, range: YrsStoryRange): boolean {
  const offset = (at: YrsStoryRange['start']) =>
    session.locateParagraph(range.story, at.paraId).start + at.offset;
  const start = offset(range.start);
  const end = offset(range.end);
  return session.listRevisions().some((revision) => {
    if (revision.story !== range.story) return false;
    if (revision.kind !== 'insertion' && revision.kind !== 'deletion') return false;
    const revisionStart = offset(revision.range.start);
    const revisionEnd = offset(revision.range.end);
    return start === end
      ? revisionStart <= start && start <= revisionEnd
      : start < revisionEnd && revisionStart < end;
  });
}
