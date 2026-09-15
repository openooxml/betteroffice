# BetterOffice redaction

Structure-preserving redaction for DOCX, XLSX, and PPTX packages. Text becomes
`x` runs (or independent random letters with `RedactionOptions {
random_characters: true }`), numbers become `8` runs, formulas become `0`,
dates become the epoch, booleans become `false`, and error values become
`#N/A`. Element counts, part names, and relationship targets are preserved.

## XLSX coverage

- Cells: shared strings, inline strings, string results, formulas, error
  values, numbers; sheet names; header/footer text.
- Comments: legacy authors and text; threaded-comment `text`, `dT`, and
  `xl/persons/person.xml` display names and user IDs. Person GUID links,
  `personId`, and `providerId` are kept so threads still resolve.
- Pivot caches: field names and captions, worksheet-source sheet names,
  shared-item and record string values; numbers, dates, booleans, and errors
  follow the placeholder policy above. Shared-item indexes are kept.
- Pivot tables and slicers: table, data-field, page-field, calculated-member,
  slicer, timeline, and cache names and captions; calculated-item formulas
  become `0`. Field and item indexes are kept.
- Connections: names, descriptions, connection strings, commands, parameter
  names, prompts, and values, text-import source files, web URLs and posts.
  Connection IDs are kept.
- External links: cached values, sheet names, defined names and `refersTo`,
  DDE service/topic/item names, OLE program IDs and item names.
- Tables and queries: table, column, query-table, and query-field names,
  table comments, totals-row labels, calculated-column formulas.
- Worksheets: hyperlinks, validation prompts, filter values, conditional-format
  match text, scenario names and comments, web-publish titles and destinations,
  OLE program IDs and links, control names.
- Rich data and metadata: rich-value strings and key names, metadata values.
- Document properties and custom properties; external-relationship targets
  become `https://example.com`.
