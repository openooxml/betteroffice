# fidelity-scorecard

## ADDED Requirements

### Requirement: Six axes with fixed vocabularies

Every criterion SHALL be scored on six axes with these statuses, ordered weakest to strongest. `n/a` is allowed where an axis does not apply to a criterion and SHALL be justified in the entry.

| Axis | Question | Statuses |
| --- | --- | --- |
| preserve | Does it survive open→edit→save? | `none` → `partial` → `gated` (corpus fixtures through both oracles and the census) |
| model | Is it typed, or carried generically? | `generic` → `partial` → `typed` |
| edit | Can a user edit it the way Word edits it? | `none` → `partial` → `word-parity` (structural-key, isolation, and undo gates) |
| layout | Is its geometry Word's geometry? | `none` → `partial` → `pinned` (fixture-pinned deciding geometry) |
| paint | Does it paint correctly and deterministically? | `none` → `partial` → `golden` (byte-exact render golden) |
| evidence | What backs the claim? | `none` → `spec` (ECMA-376 clause cited) → `word` (paired probe or Word-authored fixture pinned) |

`generic` is not a defect on the model axis — carrying unknown markup losslessly is the design — but `preserve` SHALL target `gated` for every criterion without exception: nothing is allowed to be lost, modelled or not.

#### Scenario: Preservation has no exemptions

- **WHEN** any criterion's preserve target is set below `gated`
- **THEN** the ledger test fails; preservation is the one axis with a universal target

### Requirement: The ledger is the measurement

`crates/betteroffice-docx/tests/scorecard/ledger.json` SHALL hold one entry per criterion: current status per axis, target per axis, the gates (test names) proving each status above the axis floor, and the enumerated known defects. A test SHALL verify: the entry id set equals the catalogue exactly; every claimed status names at least one gate and every named gate exists in the workspace; every measurable number (corpus size, pass rates, digest reach, required-scenario count, ceiling values, defect counts) recomputes to exactly the recorded value. The ledger regenerates only under `GOLDEN_UPDATE=1`.

#### Scenario: A claim without a gate fails

- **WHEN** an entry claims `layout: pinned` and names no existing test
- **THEN** the ledger test fails naming the criterion and the axis

#### Scenario: A stale number fails

- **WHEN** the corpus gains a fixture and the ledger's recorded corpus size is not updated
- **THEN** the exact-equality recomputation fails

### Requirement: The ledger only ratchets

Statuses and targets SHALL only move up the axis vocabularies; measured numbers SHALL only improve. A change that lowers a status, shrinks a required list, or raises a defect ceiling SHALL fail unless it is an explicit, reviewed regression acknowledgment recorded in the entry with the issue that tracks it.

#### Scenario: A silent regression cannot land

- **WHEN** a change makes a `gated` criterion fail its gate
- **THEN** either the change is fixed, or the ledger records the regression explicitly with its tracking issue — the gate never just disappears

### Requirement: The gap is a list, not an estimate

The generated `scorecard.md` SHALL report: per-axis criteria-at-target over total; corpus round-trip pass rate; census-loss rate; digest reach; and the defect list. **The gap to 100% Word fidelity is defined as: every (criterion, axis) pair below its target, plus every ledger defect.** 100% is reached when both lists are empty. The report is generated from the ledger and drift-checked against it; nobody edits the report by hand.

#### Scenario: The gap is enumerable

- **WHEN** someone asks how far the engine is from 100% Word fidelity
- **THEN** `scorecard.md` answers with the exact pair list and defect list, not a percentage estimated in prose

### Requirement: The criteria catalogue is normative and complete

The ledger SHALL contain exactly one entry for each criterion below, and only these. Adding a capability the catalogue does not cover SHALL first add its criterion here. Removing a criterion is a spec change, not a ledger edit.

**pkg — Package and preservation**

| id | covers |
| --- | --- |
| pkg.opc | Part inventory, content types, and relationship graphs survive open→save |
| pkg.unknown-parts | XML parts the engine does not model are byte-identical through save |
| pkg.media | Image and media part bytes identical through save |
| pkg.fonts | Embedded font parts, including obfuscated fonts, identical through save |
| pkg.vba | Macro and OLE binaries identical through save |
| pkg.custom-xml | customXml parts, their properties, and data-store references survive |
| pkg.doc-props | Core, extended, and custom document properties survive |
| pkg.alternate-content | `mc:AlternateContent` choices and fallbacks, `mc:Ignorable` declarations survive in position |
| pkg.unknown-xml | Unknown elements and attributes inside modelled parts survive in position |
| pkg.whitespace | Authored whitespace under `xml:space="preserve"` survives verbatim |
| pkg.settings | `settings.xml` including compatibility flags survives; flags reach layout where they change it |
| pkg.theme | Theme parts survive; theme font and color references resolve |
| pkg.fixed-point | save→reopen→save is byte-identical |

**run — Text and runs**

| id | covers |
| --- | --- |
| run.text | Character content, including surrogate pairs and symbol runs |
| run.breaks | Tabs, line/page/column breaks, soft and non-breaking hyphens, as characters and in layout |
| run.toggles | b, i, caps, smallCaps, strike, dstrike, vanish, emboss, imprint, outline, shadow, with toggle semantics |
| run.underline | Every underline style, with color |
| run.color-shading | Color, highlight, character shading |
| run.fonts | ascii/hAnsi/eastAsia/cs font slots with hint resolution |
| run.size | sz and szCs |
| run.vert-align | Superscript and subscript |
| run.metrics | spacing, kern, position, horizontal scale (w) |
| run.lang-rtl | lang, rtl, complex-script marking |
| run.extended | Extended effects (glow, reflection, ligatures, number forms) preserved; painted where modelled |
| run.field-text | instrText and cached field-result runs |

**para — Paragraphs**

| id | covers |
| --- | --- |
| para.alignment | jc including distributed |
| para.indents | left/right/firstLine/hanging, character-unit variants, negative values |
| para.spacing | before/after including auto; line rules auto/exact/atLeast |
| para.borders | pBdr edges, between-borders, bar |
| para.shading | shd fills and patterns |
| para.tabs | Custom stops, alignments, leaders; default interval from settings |
| para.keeps | keepNext, keepLines, widowControl, pageBreakBefore |
| para.frames | framePr including drop caps |
| para.numpr | numId/ilvl binding |
| para.outline-bidi | outlineLvl; paragraph direction |
| para.mark-rpr | Paragraph-mark run properties |
| para.misc | contextualSpacing, suppressLineNumbers, suppressAutoHyphens, mirrorIndents |

**style — Styles**

| id | covers |
| --- | --- |
| style.doc-defaults | docDefaults as the base of resolution |
| style.hierarchy | basedOn chains; resolution order docDefaults → table → numbering → paragraph → character → direct |
| style.linked | Linked paragraph/character style pairs |
| style.toggles | Toggle-property resolution across layers |
| style.table-conditional | tblLook with firstRow/lastRow/firstCol/lastCol/banding conditional formats |
| style.latent | Latent styles survive |
| style.duplicates | Duplicate style ids preserved as authored |
| style.defaults | Default-style attributes; fallback when a referenced style is absent |

**num — Numbering**

| id | covers |
| --- | --- |
| num.definitions | abstractNum, num, the full lvl vocabulary |
| num.overrides | lvlOverride and startOverride |
| num.formats | numFmt values including bullets, roman, letters, ordinal; picture bullets |
| num.lvltext | Multi-level %n composition |
| num.restart | Restart and continuation semantics |
| num.style-links | numStyleLink/styleLink; ilvl via paragraph style |
| num.suffix | suff, lvlJc, and marker indent interaction |

**tbl — Tables**

| id | covers |
| --- | --- |
| tbl.grid | tblGrid and grid-column claims |
| tbl.widths | tblW/tcW auto/pct/dxa; fixed vs autofit parity with Word |
| tbl.spans | gridSpan; vMerge continuation |
| tbl.nesting | Nested tables at Word's practical depths |
| tbl.borders | Per-edge borders; explicit nil vs inherited; Word's conflict resolution |
| tbl.shading | Table, row, cell shading precedence |
| tbl.margins | Cell margins, tblCellSpacing |
| tbl.alignment | Row jc, cell vAlign, cell textDirection |
| tbl.header-rows | tblHeader repetition across pages |
| tbl.row-splitting | cantSplit; row height rules; rows splitting across pages |
| tbl.floating | tblpPr positioning and text wrapping around tables |
| tbl.bidi | bidiVisual |
| tbl.cell-fit | noWrap, tcFitText |

**sect — Sections**

| id | covers |
| --- | --- |
| sect.page-setup | pgSz size and orientation, pgMar, gutter |
| sect.types | nextPage/continuous/evenPage/oddPage/nextColumn semantics |
| sect.columns | Equal and unequal columns, spacing, separator, balancing |
| sect.hf-binding | header/footerReference default/first/even; inheritance; titlePg; evenAndOddHeaders |
| sect.page-numbering | pgNumType start and format |
| sect.line-numbering | lnNumType |
| sect.page-borders | pgBorders with display and offset options |
| sect.valign | Section vertical alignment |
| sect.notes-config | footnotePr/endnotePr per section and document |

**hf — Headers and footers**

| id | covers |
| --- | --- |
| hf.content | Full block vocabulary inside header and footer stories |
| hf.inheritance | Reference inheritance across sections |
| hf.geometry | Header and footer distances; band growth against body content |
| hf.page-fields | PAGE/NUMPAGES resolving per rendered page |

**note — Footnotes and endnotes**

| id | covers |
| --- | --- |
| note.content | Full block vocabulary inside notes |
| note.refs | Reference marks, custom marks |
| note.separators | Separator, continuation separator, continuation notice |
| note.placement | Bottom-of-page reservation stealing body space; endnote placement |
| note.numbering | Restart per page and section; formats |

**fld — Fields, links, bookmarks**

| id | covers |
| --- | --- |
| fld.simple | fldSimple with cached result |
| fld.complex | fldChar begin/separate/end chains, nesting |
| fld.instructions | Instruction text verbatim; results never recomputed on save |
| fld.page | PAGE/NUMPAGES/SECTIONPAGES live in layout |
| fld.toc | TOC structure, hyperlink behavior, preserved cached results |
| fld.ref | REF/PAGEREF against bookmarks |
| fld.date-doc | DATE/AUTHOR/FILENAME cached-result policy |
| fld.form | Legacy ffData form fields |
| fld.hyperlinks | r:id and anchor links; nested run content |
| fld.bookmarks | bookmarkStart/End spanning arbitrary structures |

**drw — Drawings and objects**

| id | covers |
| --- | --- |
| drw.inline | Inline drawings: extent, effectExtent, docPr |
| drw.anchor-position | Anchor bases (page/margin/column/character/paragraph), align and posOffset |
| drw.wrap | square, tight, through, topAndBottom, none; behind and in-front |
| drw.zorder | relativeHeight ordering and overlap |
| drw.formats | png/jpeg/gif/bmp/tiff/emf/wmf/svg with fallbacks |
| drw.transforms | Crop, rotation, flip; effects preserved |
| drw.textboxes | Text-box body content, anchoring, writing direction |
| drw.shapes | Preset geometry, adjust values, fills, outlines, theme colors |
| drw.groups | Group shapes with child transforms |
| drw.charts | Chart parts preserved; rendered |
| drw.smartart | Diagram parts preserved |
| drw.vml | Legacy VML (w:pict), watermarks |
| drw.ole | OLE objects with icon presentation |
| drw.ink | Ink annotations preserved |

**trk — Tracked changes**

| id | covers |
| --- | --- |
| trk.runs | w:ins and w:del with delText |
| trk.moves | moveFrom/moveTo with range bookmarks |
| trk.format | rPrChange, pPrChange, sectPrChange, tblPrChange, tcPrChange, trPrChange, numberingChange |
| trk.structure | Tracked row insert and delete, cell merge |
| trk.metadata | Authors, dates, rsids preserved |
| trk.render | All-markup rendering with attribution and changed lines |
| trk.accept-reject | Accept and reject produce Word's outcome |

**cmt — Comments**

| id | covers |
| --- | --- |
| cmt.ranges | commentRangeStart/End and references |
| cmt.threads | Replies and resolution state |
| cmt.content | Full block vocabulary in comment stories |
| cmt.render | Anchor highlighting and margin presentation |

**sdt — Content controls**

| id | covers |
| --- | --- |
| sdt.levels | Block, inline, row, and cell controls preserved |
| sdt.types | richText, text, dropdown, combo, date, checkbox, picture, group, repeating |
| sdt.locks | Lock semantics with ancestor union |
| sdt.placeholder | showingPlcHdr as a state, not a string |
| sdt.binding | dataBinding preserved; bound controls refuse edits |

**math — Math**

| id | covers |
| --- | --- |
| math.preserve | OMML inline and display preserved |
| math.render | OMML layout and paint |

**lay — Layout engine**

| id | covers |
| --- | --- |
| lay.shaping | Glyph shaping and font metrics parity |
| lay.line-breaking | Word-wrap points; no unauthorized hyphenation |
| lay.justification | Left, right, center, both |
| lay.line-height | Auto spacing largest-run rule; exact and atLeast |
| lay.tabs | Default and custom stop resolution with leaders |
| lay.page-breaking | Widow/orphan, keeps, explicit breaks |
| lay.columns | Balancing and column breaks |
| lay.tables | Autofit algorithm, row splitting, header repetition |
| lay.floats | Wrap geometry, multiple floats, overlap |
| lay.notes | Footnote space reservation |
| lay.hf | Band geometry and growth |
| lay.fields | Page-dependent field results |
| lay.rtl | Bidi reordering and mirroring |
| lay.cjk | East Asian line rules and CJK metrics |
| lay.vertical | Vertical text in cells and text boxes |
| lay.hit | Hit testing and caret geometry agree with paint |

**paint — Paint**

| id | covers |
| --- | --- |
| paint.text | Glyphs and text decorations |
| paint.borders | Border styles and shading patterns |
| paint.images | Decode, crop, transform |
| paint.shapes | Fills, outlines, geometry |
| paint.marks | Tracked-change and comment marks |
| paint.determinism | Byte-identical output for identical input |

**edit — Editing machinery**

| id | covers |
| --- | --- |
| edit.reach | The caret enters every story Word's caret enters |
| edit.selection | Selection across boundaries; cell and block selection |
| edit.typing | Insertion at every position class, including link, field, and note-reference boundaries |
| edit.keys | Enter, Tab, Shift+Tab, Backspace, Delete structural parity |
| edit.lists | Continuation, level change, numbering removal |
| edit.tables | Cell navigation; row and column insert/delete; merge and split |
| edit.formatting | Apply and clear formatting with Word's toggle and range semantics |
| edit.clipboard | Cut, copy, paste within a document preserves structure and formatting |
| edit.undo | Exact undo per operation family |
| edit.isolation | Edit-scope fingerprint per operation family |
| edit.tracked | Edits recorded correctly under tracking; reject restores the original |
| edit.collab | Concurrent edits converge to a faithful state |

#### Scenario: The catalogue and the ledger cannot diverge

- **WHEN** the ledger contains an id absent from this catalogue, or misses one present
- **THEN** the ledger test fails listing the difference
