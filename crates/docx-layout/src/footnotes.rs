//! Footnote layout — Rust port of
//! `packages/core/src/layout/regions/footnoteLayout.ts`.
//!
//! Self-contained: all input structs are LOCAL mirrors of the TS parameter
//! shapes (serde camelCase, unknown fields preserved via `#[serde(flatten)]`
//! so blocks pass through `apply_footnote_presentation` losslessly). Maps are
//! insertion-ordered (`OrderedMap`) to mirror TS `Map` iteration semantics —
//! iteration order is load-bearing for byte-identity of reservation maps.
//!
//! Ported exports (TS name → Rust name, 1:1 semantics):
//! - `FOOTNOTE_SEPARATOR_HEIGHT`, `FOOTNOTE_COLUMN_GAP_PX`,
//!   `MAX_FOOTNOTE_LAYOUT_PASSES` — same names
//! - `footnoteReservedHeightsEqual` → `footnote_reserved_heights_equal`
//! - `collectFootnoteRefs` → `collect_footnote_refs`
//! - `mapFootnotesToPages` → `map_footnotes_to_pages`
//! - `applyFootnotePresentation` → `apply_footnote_presentation`
//! - `distributeFootnotesIntoColumns` → `distribute_footnotes_into_columns`
//! - `calculateFootnoteReservedHeights` → `calculate_footnote_reserved_heights`
//! - `buildFootnoteContentMap` → `build_footnote_content_map` (the TS
//!   `contentWidth` + `ConvertFootnoteOptions` args feed only
//!   `convertFootnoteToContent`, which stays JS-side, so here they are folded
//!   into the injected conversion callback)
//! - `stabilizeFootnoteLayout` → `stabilize_footnote_layout` (the TS
//!   `blocks`/`measures`/`layoutOpts` args feed only `layoutDocument`; here
//!   the relayout is an injected callback standing in for
//!   `layoutDocument(measured, { ...layoutOpts, footnoteReservedHeights })`)
//!
//! NOT ported — these cross the ProseMirror/Canvas/DOM seam and stay in TS:
//! `convertFootnoteToContent` (footnoteToProseDoc → toLayoutBlocks → adapter
//! `measureBlocks`), `MeasureBlocksFn`, `ConvertFootnoteOptions`, and
//! `buildFootnoteRenderItems` (paint-side payload; needs `Document` +
//! `getFootnoteText`). The Rust layout core only ever needs footnote content
//! HEIGHTS (the `HasHeight` trait); the JS host measures and passes them in.
//!
//! # INTEGRATION — reserved-height feedback (hook for the place/spine side)
//!
//! Footnote presence reserves page height, and reserving height can move a
//! reference to another page, so layout re-enters. The TS control flow:
//!
//! 1. `LayoutOptions.footnoteReservedHeights?: Map<pageNumber, px>`
//!    (`layout/pagination/types.ts:1049`). `layoutDocument` does nothing with
//!    it except forward it into `createPageFlow`
//!    (`layout/pagination/index.ts:168`).
//! 2. The ONLY in-pass consumer is page creation
//!    (`layout/pagination/pageFlow.ts:158-169`, option declared at `:53`):
//!    for the page being created, `pageNumber = pages.length + 1 +
//!    pageNumberOffset`, `footnoteHeight =
//!    footnoteReservedHeights?.get(pageNumber) ?? 0`, and the page state's
//!    `contentLimit` becomes `(pageSize.h - margins.bottom) - footnoteHeight`
//!    (`getContentBottom` is `pageFlow.ts:106-108`). Every subsequent
//!    fits/overflow/break decision on that page reads the reduced limit.
//!    When `footnoteHeight > 0` it is also stamped on the output page as
//!    `page.footnoteReservedHeight` (`types.ts:945`). THIS is the whole hook
//!    the Rust spine paginator must implement — it is what the golden
//!    `footnotes-force-content-up` exercises (`__golden__/corpus.ts:260-266`
//!    reserves `Map{1 → 200}`; the golden asserts
//!    `"footnoteReservedHeight": 200` on page 1 and the 8th paragraph pushed
//!    to page 2).
//! 3. A single layout pass never recomputes reservations. The re-entry loop
//!    lives OUTSIDE `layoutDocument`, in `stabilizeFootnoteLayout`
//!    (`footnoteLayout.ts:555-642`, ported below): map refs → pages, compute
//!    per-page reservations, re-run the FULL layout with them, repeat until
//!    the reservation map reaches a fixpoint (≤ `MAX_FOOTNOTE_LAYOUT_PASSES`);
//!    if it oscillates, a fallback loop max-merges reservations monotonically
//!    (≤ another `MAX_FOOTNOTE_LAYOUT_PASSES`) until every page's requirement
//!    is covered, then settles. Finally `page.footnoteIds` (and
//!    `page.footnoteColumns` when > 1) are written onto the result pages.
//! 4. Incremental relayout must BAIL to a full layout whenever
//!    `footnoteReservedHeights` is non-empty — reservation couples pages
//!    globally (`layout/pagination/incremental.ts:191`, forwarded at `:250`).

#![allow(dead_code)]

use serde::{Deserialize, Serialize};
use serde_json::{Map as JsonMap, Value};

/// Separator line height + vertical padding in pixels.
pub const FOOTNOTE_SEPARATOR_HEIGHT: f64 = 12.0;

/// Gutter between footnote columns when `w15:footnoteColumns` > 1, in pixels
/// (~0.25in). Shared by the reserved-height/measurement path and the footnote
/// painter so a footnote measured at column width paints into a column of
/// exactly that width. Single-column footnotes never consult it.
pub const FOOTNOTE_COLUMN_GAP_PX: f64 = 24.0;

/// Hard cap on the multi-pass footnote layout loop. Reserving footnote space
/// can move a reference to another page, so callers keep remapping until the
/// page→height contract is stable. Dense layouts converge in 2-3 passes in
/// practice; 6 is a safe ceiling.
pub const MAX_FOOTNOTE_LAYOUT_PASSES: usize = 6;

/// Default footnote font size in points. Word's built-in "Footnote Text"
/// style sets 8pt; applied only when the footnote's runs don't already
/// specify a fontSize (avoids overriding authored sizes).
const FOOTNOTE_FONT_SIZE_PT: f64 = 8.0;

// ============================================================================
// insertion-ordered map (mirrors TS `Map` semantics)
// ============================================================================

/// Insertion-ordered key→value map mirroring TS `Map`: `set` on an existing
/// key updates in place (original position kept), iteration is insertion
/// order. Reservation/page maps are tiny (one entry per page), so a Vec scan
/// beats hashing and — unlike `HashMap` — keeps TS iteration order exactly.
#[derive(Clone, Debug, PartialEq)]
pub struct OrderedMap<K, V> {
    entries: Vec<(K, V)>,
}

impl<K: PartialEq, V> OrderedMap<K, V> {
    pub fn new() -> Self {
        Self {
            entries: Vec::new(),
        }
    }

    pub fn len(&self) -> usize {
        self.entries.len()
    }

    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    pub fn get(&self, key: &K) -> Option<&V> {
        self.entries.iter().find(|(k, _)| k == key).map(|(_, v)| v)
    }

    pub fn get_mut(&mut self, key: &K) -> Option<&mut V> {
        self.entries
            .iter_mut()
            .find(|(k, _)| k == key)
            .map(|(_, v)| v)
    }

    /// TS `Map.set`: replace in place when the key exists, append otherwise.
    pub fn set(&mut self, key: K, value: V) {
        match self.get_mut(&key) {
            Some(slot) => *slot = value,
            None => self.entries.push((key, value)),
        }
    }

    pub fn iter(&self) -> std::slice::Iter<'_, (K, V)> {
        self.entries.iter()
    }
}

impl<K: PartialEq, V> Default for OrderedMap<K, V> {
    fn default() -> Self {
        Self::new()
    }
}

impl<K: PartialEq, V> FromIterator<(K, V)> for OrderedMap<K, V> {
    fn from_iter<I: IntoIterator<Item = (K, V)>>(iter: I) -> Self {
        let mut map = Self::new();
        for (k, v) in iter {
            map.set(k, v);
        }
        map
    }
}

// ============================================================================
// local input mirrors (duck-typed like the TS structural types; unknown
// fields survive round-trips through `rest`)
// ============================================================================

/// TS `BlockId = string | number`.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(untagged)]
pub enum BlockId {
    Str(String),
    Num(f64),
}

/// One flow run; only the fields this module's logic touches are declared
/// (`kind`, `text`, `fontSize`, `fontFamily`, `superscript`, `footnoteRefId`,
/// `pmStart`), the rest ride along untouched.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Run {
    pub kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_size: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub superscript: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footnote_ref_id: Option<i64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(flatten)]
    pub rest: JsonMap<String, Value>,
}

/// One flow block, duck-typed on `kind` exactly like the TS union:
/// `runs` when `kind == "paragraph"`, `rows` when `kind == "table"`,
/// `content` when `kind == "textBox"`. Unknown kinds pass through untouched.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct LayoutBlock {
    pub kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub id: Option<BlockId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub runs: Option<Vec<Run>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub rows: Option<Vec<TableRow>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub content: Option<Vec<LayoutBlock>>,
    #[serde(flatten)]
    pub rest: JsonMap<String, Value>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableRow {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub id: Option<BlockId>,
    pub cells: Vec<TableCell>,
    #[serde(flatten)]
    pub rest: JsonMap<String, Value>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableCell {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub id: Option<BlockId>,
    pub blocks: Vec<LayoutBlock>,
    #[serde(flatten)]
    pub rest: JsonMap<String, Value>,
}

/// One laid-out fragment; only the fields the page-mapping logic reads are
/// declared. `rowStart`/`rowEnd` are present on table fragments.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Fragment {
    pub kind: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub block_id: Option<BlockId>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub pm_end: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub row_start: Option<f64>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub row_end: Option<f64>,
    #[serde(flatten)]
    pub rest: JsonMap<String, Value>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Page {
    pub number: u32,
    pub fragments: Vec<Fragment>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footnote_ids: Option<Vec<i64>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub footnote_columns: Option<f64>,
    #[serde(flatten)]
    pub rest: JsonMap<String, Value>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Layout {
    pub pages: Vec<Page>,
    #[serde(flatten)]
    pub rest: JsonMap<String, Value>,
}

/// TS `Footnote` — only `id` + `noteType` are read here.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Footnote {
    pub id: i64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub note_type: Option<String>,
    #[serde(flatten)]
    pub rest: JsonMap<String, Value>,
}

/// TS `FootnoteContent`, minus `blocks`/`measures` — those stay JS-side with
/// the measurement pipeline; the layout core only consumes the total height.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FootnoteContent {
    pub id: i64,
    pub display_number: i64,
    pub height: f64,
    #[serde(flatten)]
    pub rest: JsonMap<String, Value>,
}

/// Stand-in for the TS structural constraint `T extends { height: number }`.
pub trait HasHeight {
    fn height(&self) -> f64;
}

impl HasHeight for FootnoteContent {
    fn height(&self) -> f64 {
        self.height
    }
}

/// Anonymous `{ height }` object used inside the reservation calculation.
#[derive(Clone, Debug, PartialEq)]
struct HeightItem {
    height: f64,
}

impl HasHeight for HeightItem {
    fn height(&self) -> f64 {
        self.height
    }
}

// ============================================================================
// reserved-height map helpers
// ============================================================================

/// Compare two per-page footnote reservation maps. Used by the multi-pass
/// loop to detect when it has converged.
pub fn footnote_reserved_heights_equal(a: &OrderedMap<u32, f64>, b: &OrderedMap<u32, f64>) -> bool {
    if a.len() != b.len() {
        return false;
    }
    for (page_number, height) in a.iter() {
        if b.get(page_number).copied() != Some(*height) {
            return false;
        }
    }
    true
}

fn footnote_reserved_heights_cover(
    reserved: &OrderedMap<u32, f64>,
    required: &OrderedMap<u32, f64>,
) -> bool {
    for (page_number, height) in required.iter() {
        if reserved.get(page_number).copied().unwrap_or(0.0) < *height {
            return false;
        }
    }
    true
}

fn merge_footnote_reserved_heights(
    a: &OrderedMap<u32, f64>,
    b: &OrderedMap<u32, f64>,
) -> OrderedMap<u32, f64> {
    let mut merged = a.clone();
    for (page_number, height) in b.iter() {
        let current = merged.get(page_number).copied().unwrap_or(0.0);
        merged.set(*page_number, current.max(*height));
    }
    merged
}

// ============================================================================
// 1. Scan FlowBlocks for footnote references
// ============================================================================

/// Where a footnote reference lives, as found by [`collect_footnote_refs`].
///
/// `pm_pos` alone is enough to attribute a reference to a page for ordinary
/// (paragraph) content, whose fragments carry a per-page pm sub-range. A table
/// is different: it splits across pages by ROW, but every table fragment keeps
/// the whole table's `pmStart`/`pmEnd`. So for a reference authored inside a
/// table cell the OUTERMOST table's id and the index of the row that contains
/// it are also recorded, letting [`map_footnotes_to_pages`] attribute the
/// reference to the page that actually laid out that row.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FootnoteRefLocation {
    pub footnote_id: i64,
    pub pm_pos: f64,
    /// Id of the outermost enclosing table block, when the ref is in a table cell.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub table_block_id: Option<BlockId>,
    /// Index (into the outermost table's `rows`) of the row holding the ref.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub row_index: Option<u32>,
}

/// Scan FlowBlocks for runs with `footnoteRefId` set. Returns the refs in
/// document order. Recurses into container blocks (table cells, text boxes);
/// for refs inside a table, the OUTERMOST table's id and row index are
/// recorded (a nested table keeps the outer context, since the outer row is
/// what the paginator splits into per-page fragments).
pub fn collect_footnote_refs(blocks: &[LayoutBlock]) -> Vec<FootnoteRefLocation> {
    struct TableCtx {
        table_block_id: Option<BlockId>,
        row_index: u32,
    }

    fn walk(
        input: &[LayoutBlock],
        table_ctx: Option<&TableCtx>,
        refs: &mut Vec<FootnoteRefLocation>,
    ) {
        for block in input {
            match block.kind.as_str() {
                "paragraph" => {
                    let Some(runs) = &block.runs else { continue };
                    for run in runs {
                        if run.kind == "text" {
                            if let Some(footnote_id) = run.footnote_ref_id {
                                refs.push(FootnoteRefLocation {
                                    footnote_id,
                                    pm_pos: run.pm_start.unwrap_or(0.0),
                                    table_block_id: table_ctx
                                        .and_then(|c| c.table_block_id.clone()),
                                    row_index: table_ctx.map(|c| c.row_index),
                                });
                            }
                        }
                    }
                }
                "table" => {
                    let Some(rows) = &block.rows else { continue };
                    for (row_index, row) in rows.iter().enumerate() {
                        // nested tables keep the OUTER table's context —
                        // pagination decisions happen at the granularity of
                        // the outermost row.
                        let own_ctx;
                        let ctx = match table_ctx {
                            Some(ctx) => ctx,
                            None => {
                                own_ctx = TableCtx {
                                    table_block_id: block.id.clone(),
                                    row_index: row_index as u32,
                                };
                                &own_ctx
                            }
                        };
                        for cell in &row.cells {
                            walk(&cell.blocks, Some(ctx), refs);
                        }
                    }
                }
                "textBox" => {
                    if let Some(content) = &block.content {
                        walk(content, table_ctx, refs);
                    }
                }
                _ => {}
            }
        }
    }

    let mut refs = Vec::new();
    walk(blocks, None, &mut refs);
    refs
}

// ============================================================================
// 2. Map footnote references to pages
// ============================================================================

/// After layout, determine which footnotes appear on which pages. Checks each
/// page's fragments to see if any footnote-ref PM positions fall within.
/// Returns `pageNumber → footnoteId[]` in document order.
pub fn map_footnotes_to_pages(
    pages: &[Page],
    footnote_refs: &[FootnoteRefLocation],
) -> OrderedMap<u32, Vec<i64>> {
    let mut page_footnotes: OrderedMap<u32, Vec<i64>> = OrderedMap::new();

    if footnote_refs.is_empty() {
        return page_footnotes;
    }

    // For each footnote ref, find which page it lands on.
    for r in footnote_refs {
        'pages: for page in pages {
            for fragment in &page.fragments {
                let matched = match (&r.table_block_id, r.row_index) {
                    (Some(table_block_id), Some(row_index)) => {
                        // In-table ref: a table splits across pages by row, but
                        // every fragment keeps the whole table's pm range, so a
                        // pm-position match would land every ref on the first
                        // table page. Attribute the ref to the fragment whose
                        // [rowStart, rowEnd) slice contains its row.
                        fragment.kind == "table"
                            && fragment.block_id.as_ref() == Some(table_block_id)
                            && fragment
                                .row_start
                                .is_some_and(|rs| f64::from(row_index) >= rs)
                            && fragment.row_end.is_some_and(|re| f64::from(row_index) < re)
                    }
                    _ => {
                        let pm_start = fragment.pm_start.unwrap_or(-1.0);
                        let pm_end = fragment.pm_end.unwrap_or(-1.0);
                        pm_start >= 0.0
                            && pm_end >= 0.0
                            && r.pm_pos >= pm_start
                            && r.pm_pos < pm_end
                    }
                };
                if matched {
                    // Avoid duplicates (same footnote shouldn't appear twice on
                    // the same page).
                    match page_footnotes.get_mut(&page.number) {
                        Some(existing) => {
                            if !existing.contains(&r.footnote_id) {
                                existing.push(r.footnote_id);
                            }
                        }
                        None => page_footnotes.set(page.number, vec![r.footnote_id]),
                    }
                    break 'pages;
                }
            }
        }
    }

    page_footnotes
}

// ============================================================================
// 3. Footnote presentation + content map
// ============================================================================

/// Footnote-specific block normalization: post-process the body-pipeline
/// output for a single footnote so it carries the correct visual prefix (its
/// display number, rendered as a superscript) and a default 8pt font for any
/// run that didn't specify a size. The display number is prepended onto the
/// FIRST paragraph as a fresh superscript text run.
pub fn apply_footnote_presentation(
    blocks: Vec<LayoutBlock>,
    display_number: i64,
) -> Vec<LayoutBlock> {
    if blocks.is_empty() {
        return vec![LayoutBlock {
            kind: "paragraph".to_string(),
            id: Some(BlockId::Str(format!("fn-empty-{display_number}"))),
            runs: Some(vec![Run {
                kind: "text".to_string(),
                text: Some(format!("{display_number}  ")),
                font_size: Some(FOOTNOTE_FONT_SIZE_PT),
                superscript: Some(true),
                ..Default::default()
            }]),
            ..Default::default()
        }];
    }

    // Apply default 8pt to every text/tab run that didn't specify a fontSize.
    let mut out: Vec<LayoutBlock> = blocks
        .into_iter()
        .map(|mut b| {
            if b.kind != "paragraph" {
                return b;
            }
            if let Some(runs) = b.runs.as_mut() {
                for r in runs.iter_mut() {
                    if (r.kind == "text" || r.kind == "tab") && r.font_size.is_none() {
                        r.font_size = Some(FOOTNOTE_FONT_SIZE_PT);
                    }
                }
            }
            b
        })
        .collect();

    // Prepend display number on the first paragraph. Match the marker's font
    // to the note text it precedes: Word renders the footnote number in the
    // FootnoteText paragraph font; the FootnoteReference char style only adds
    // superscript, not a face. When the note text itself has no explicit font
    // the marker stays unset too (both then inherit the same container font).
    let first = &mut out[0];
    if first.kind == "paragraph" {
        let first_text_font = first
            .runs
            .as_ref()
            .and_then(|runs| runs.iter().find(|r| r.kind == "text"))
            .and_then(|r| r.font_family.clone())
            // mirror the TS truthiness check — an empty string is falsy
            .filter(|f| !f.is_empty());
        let number_run = Run {
            kind: "text".to_string(),
            text: Some(format!("{display_number}  ")),
            font_size: Some(FOOTNOTE_FONT_SIZE_PT),
            superscript: Some(true),
            font_family: first_text_font,
            ..Default::default()
        };
        match first.runs.as_mut() {
            Some(runs) => runs.insert(0, number_run),
            None => first.runs = Some(vec![number_run]),
        }
    }

    out
}

/// Build footnote content for all footnotes referenced in the document.
/// Display numbers are assigned by first-appearance order (the same way Word
/// renders them). The conversion callback stands in for the TS
/// `convertFootnoteToContent(footnote, displayNumber, contentWidth, options)`
/// — `contentWidth`/`options` are captured by the closure, since the actual
/// conversion (ProseMirror bridge + adapter measurement) stays JS-side.
pub fn build_footnote_content_map<F>(
    footnotes: &[Footnote],
    footnote_refs: &[FootnoteRefLocation],
    mut convert_footnote_to_content: F,
) -> OrderedMap<i64, FootnoteContent>
where
    F: FnMut(&Footnote, i64) -> FootnoteContent,
{
    let mut content_map: OrderedMap<i64, FootnoteContent> = OrderedMap::new();
    let mut footnote_by_id: OrderedMap<i64, &Footnote> = OrderedMap::new();

    for f in footnotes {
        if f.note_type.as_deref() == Some("normal") || f.note_type.is_none() {
            footnote_by_id.set(f.id, f);
        }
    }

    let mut display_number: i64 = 1;
    let mut seen: Vec<i64> = Vec::new();

    for r in footnote_refs {
        if seen.contains(&r.footnote_id) {
            continue;
        }
        seen.push(r.footnote_id);

        let Some(footnote) = footnote_by_id.get(&r.footnote_id).copied() else {
            continue;
        };

        content_map.set(
            r.footnote_id,
            convert_footnote_to_content(footnote, display_number),
        );
        display_number += 1;
    }

    content_map
}

// ============================================================================
// 4. Per-page footnote area height reservation
// ============================================================================

/// Distribute footnote items across `columns` balanced columns, preserving
/// document order (footnotes must still read in numeric sequence). Items fill
/// the first column until it reaches the balanced target height (~ total / N),
/// then spill into the next column — the same order-preserving balance Word
/// applies to its footnote columns, not a greedy shortest-column packing
/// (which would scramble the reading order).
///
/// `columns <= 1` (the default for ordinary single-column footnotes) returns a
/// single column unchanged.
pub fn distribute_footnotes_into_columns<T: HasHeight>(items: Vec<T>, columns: f64) -> Vec<Vec<T>> {
    let n = columns.floor().max(1.0);
    if n <= 1.0 || items.len() <= 1 {
        return vec![items];
    }
    let n = n as usize;

    let total = items.iter().fold(0.0_f64, |sum, item| sum + item.height());
    let target = total / n as f64;

    let mut result: Vec<Vec<T>> = vec![Vec::new()];
    let mut column_height = 0.0_f64;
    for item in items {
        // Move to the next column once the current one has passed the balanced
        // target (measured at the item's midpoint to avoid lopsided splits)
        // and columns remain. Never leave a column empty.
        if result.len() < n && column_height > 0.0 && column_height + item.height() / 2.0 > target {
            result.push(Vec::new());
            column_height = 0.0;
        }
        column_height += item.height();
        result
            .last_mut()
            .expect("result starts non-empty")
            .push(item);
    }

    result
}

/// Calculate per-page footnote reserved heights. Returns
/// `pageNumber → reservedHeight`.
///
/// With `columns > 1` the footnotes are balanced across that many columns and
/// the reserved height is the tallest column (plus the separator), since the
/// columns sit side by side — not the sum of every footnote height.
pub fn calculate_footnote_reserved_heights<V: HasHeight>(
    page_footnote_map: &OrderedMap<u32, Vec<i64>>,
    footnote_content_map: &OrderedMap<i64, V>,
    columns: f64,
) -> OrderedMap<u32, f64> {
    let mut reserved: OrderedMap<u32, f64> = OrderedMap::new();

    for (page_number, footnote_ids) in page_footnote_map.iter() {
        let heights: Vec<HeightItem> = footnote_ids
            .iter()
            .map(|fn_id| {
                footnote_content_map
                    .get(fn_id)
                    .map(|c| c.height())
                    .unwrap_or(0.0)
            })
            .filter(|h| *h > 0.0)
            .map(|height| HeightItem { height })
            .collect();

        if heights.is_empty() {
            continue;
        }

        let cols = distribute_footnotes_into_columns(heights, columns);
        let tallest_column = cols.iter().fold(0.0_f64, |max, col| {
            max.max(col.iter().fold(0.0_f64, |sum, item| sum + item.height))
        });

        if tallest_column > 0.0 {
            // Add separator height
            reserved.set(*page_number, tallest_column + FOOTNOTE_SEPARATOR_HEIGHT);
        }
    }

    reserved
}

// ============================================================================
// 4b. Multi-pass footnote layout convergence
// ============================================================================

/// Result of [`stabilize_footnote_layout`].
pub struct StabilizeFootnoteLayoutResult {
    pub layout: Layout,
    pub page_footnote_map: OrderedMap<u32, Vec<i64>>,
    /// True if the loop converged before hitting `MAX_FOOTNOTE_LAYOUT_PASSES`.
    pub converged: bool,
}

/// Run the multi-pass footnote layout loop. Reserving footnote space on a
/// page can move a reference to another page, which changes the reservation,
/// which can move references again. Iterate until the page→height contract is
/// the same one used by the latest layout, or `MAX_FOOTNOTE_LAYOUT_PASSES`
/// passes have run. Writes `page.footnote_ids` (and `page.footnote_columns`
/// when > 1) onto each page in the returned layout.
///
/// `layout_with_reserved` stands in for the TS
/// `layoutDocument(measured, { ...layoutOpts, footnoteReservedHeights })` —
/// the spine paginator owns that call; it must reduce each page's content
/// bottom by its reserved height (see the module INTEGRATION note).
/// `footnote_columns` is `w15:footnoteColumns` (`None` = 1). The TS version
/// logs a console warning when it fails to stabilize; here the caller reads
/// `converged` instead.
pub fn stabilize_footnote_layout<F, V>(
    mut layout_with_reserved: F,
    footnote_refs: &[FootnoteRefLocation],
    footnote_content_map: &OrderedMap<i64, V>,
    initial_layout: Layout,
    footnote_columns: Option<f64>,
) -> StabilizeFootnoteLayoutResult
where
    F: FnMut(&OrderedMap<u32, f64>) -> Layout,
    V: HasHeight,
{
    let footnote_columns = footnote_columns.unwrap_or(1.0).max(1.0);

    let mut page_footnote_map = map_footnotes_to_pages(&initial_layout.pages, footnote_refs);
    let mut footnote_reserved_heights = calculate_footnote_reserved_heights(
        &page_footnote_map,
        footnote_content_map,
        footnote_columns,
    );

    if footnote_reserved_heights.is_empty() {
        return StabilizeFootnoteLayoutResult {
            layout: initial_layout,
            page_footnote_map,
            converged: true,
        };
    }

    let mut new_layout = initial_layout;
    let mut converged = false;
    for _pass in 0..MAX_FOOTNOTE_LAYOUT_PASSES {
        new_layout = layout_with_reserved(&footnote_reserved_heights);

        let next_page_footnote_map = map_footnotes_to_pages(&new_layout.pages, footnote_refs);
        let next_footnote_reserved_heights = calculate_footnote_reserved_heights(
            &next_page_footnote_map,
            footnote_content_map,
            footnote_columns,
        );

        page_footnote_map = next_page_footnote_map;
        if footnote_reserved_heights_equal(
            &footnote_reserved_heights,
            &next_footnote_reserved_heights,
        ) {
            footnote_reserved_heights = next_footnote_reserved_heights;
            converged = true;
            break;
        }
        footnote_reserved_heights = next_footnote_reserved_heights;
    }

    if !converged {
        // Oscillating layouts settle with conservative (max-merged, monotone)
        // page reservations that cover every observed requirement.
        let mut fallback_reserved_heights = footnote_reserved_heights.clone();
        let mut fallback_covered = false;
        for _pass in 0..MAX_FOOTNOTE_LAYOUT_PASSES {
            new_layout = layout_with_reserved(&fallback_reserved_heights);
            page_footnote_map = map_footnotes_to_pages(&new_layout.pages, footnote_refs);
            let required_heights = calculate_footnote_reserved_heights(
                &page_footnote_map,
                footnote_content_map,
                footnote_columns,
            );
            if footnote_reserved_heights_cover(&fallback_reserved_heights, &required_heights) {
                fallback_covered = true;
                break;
            }
            fallback_reserved_heights =
                merge_footnote_reserved_heights(&fallback_reserved_heights, &required_heights);
        }
        if !fallback_covered {
            new_layout = layout_with_reserved(&fallback_reserved_heights);
            page_footnote_map = map_footnotes_to_pages(&new_layout.pages, footnote_refs);
        }
    }

    for (page_num, fn_ids) in page_footnote_map.iter() {
        if let Some(page) = new_layout.pages.iter_mut().find(|p| p.number == *page_num) {
            page.footnote_ids = Some(fn_ids.clone());
            if footnote_columns > 1.0 {
                page.footnote_columns = Some(footnote_columns);
            }
        }
    }

    StabilizeFootnoteLayoutResult {
        layout: new_layout,
        page_footnote_map,
        converged,
    }
}

// ============================================================================
// tests — ported 1:1 from
// packages/core/src/layout/regions/__tests__/footnoteLayout.test.ts and
// packages/core/src/layout/regions/__tests__/footnote-columns-distribute.test.ts
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    fn paragraph_with_footnote(id: &str, footnote_id: i64, pm_start: f64) -> LayoutBlock {
        LayoutBlock {
            kind: "paragraph".to_string(),
            id: Some(BlockId::Str(id.to_string())),
            runs: Some(vec![Run {
                kind: "text".to_string(),
                text: Some("x".to_string()),
                footnote_ref_id: Some(footnote_id),
                pm_start: Some(pm_start),
                ..Default::default()
            }]),
            ..Default::default()
        }
    }

    fn body_ref(footnote_id: i64, pm_pos: f64) -> FootnoteRefLocation {
        FootnoteRefLocation {
            footnote_id,
            pm_pos,
            table_block_id: None,
            row_index: None,
        }
    }

    fn table_ref(
        footnote_id: i64,
        pm_pos: f64,
        table_id: &str,
        row_index: u32,
    ) -> FootnoteRefLocation {
        FootnoteRefLocation {
            footnote_id,
            pm_pos,
            table_block_id: Some(BlockId::Str(table_id.to_string())),
            row_index: Some(row_index),
        }
    }

    // ---- footnote layout reservation (footnoteLayout.test.ts) -------------

    #[test]
    fn adds_the_shared_separator_height_to_each_page_reservation() {
        let page_map: OrderedMap<u32, Vec<i64>> =
            [(1, vec![10, 11]), (3, vec![12])].into_iter().collect();
        let content_map: OrderedMap<i64, HeightItem> = [
            (10, HeightItem { height: 14.0 }),
            (11, HeightItem { height: 18.0 }),
            (12, HeightItem { height: 9.0 }),
        ]
        .into_iter()
        .collect();

        let reserved = calculate_footnote_reserved_heights(&page_map, &content_map, 1.0);

        assert_eq!(
            reserved.get(&1).copied(),
            Some(14.0 + 18.0 + FOOTNOTE_SEPARATOR_HEIGHT)
        );
        assert_eq!(
            reserved.get(&3).copied(),
            Some(9.0 + FOOTNOTE_SEPARATOR_HEIGHT)
        );
    }

    // ---- applyFootnotePresentation -----------------------------------------

    #[test]
    fn the_synthetic_marker_run_inherits_the_footnote_text_font() {
        // Regression: the prepended number run carried no fontFamily, so the
        // painter fell back to the inherited container default and the footnote
        // number rendered in a different font than the note text. The marker
        // must match the note's font (Word renders the number in the
        // FootnoteText face; the FootnoteReference char style only adds
        // superscript, not a face).
        let blocks = vec![LayoutBlock {
            kind: "paragraph".to_string(),
            id: Some(BlockId::Str("fn1".to_string())),
            runs: Some(vec![Run {
                kind: "text".to_string(),
                text: Some("See note.".to_string()),
                font_family: Some("Cambria".to_string()),
                font_size: Some(10.0),
                ..Default::default()
            }]),
            ..Default::default()
        }];

        let out = apply_footnote_presentation(blocks, 3);
        let marker = &out[0].runs.as_ref().unwrap()[0];

        assert_eq!(marker.text.as_deref(), Some("3  "));
        assert_eq!(marker.superscript, Some(true));
        assert_eq!(marker.font_family.as_deref(), Some("Cambria"));
    }

    #[test]
    fn leaves_the_marker_font_unset_when_the_note_text_has_no_explicit_font() {
        // Both marker and note text then inherit the same container font, so
        // they still match; we must not invent a divergent family.
        let blocks = vec![LayoutBlock {
            kind: "paragraph".to_string(),
            id: Some(BlockId::Str("fn1".to_string())),
            runs: Some(vec![Run {
                kind: "text".to_string(),
                text: Some("See note.".to_string()),
                ..Default::default()
            }]),
            ..Default::default()
        }];

        let out = apply_footnote_presentation(blocks, 1);
        let marker = &out[0].runs.as_ref().unwrap()[0];

        assert_eq!(marker.font_family, None);
    }

    // ---- collectFootnoteRefs -----------------------------------------------

    #[test]
    fn collects_refs_from_top_level_paragraphs() {
        let blocks = vec![
            paragraph_with_footnote("p1", 1, 10.0),
            paragraph_with_footnote("p2", 2, 20.0),
        ];

        assert_eq!(
            collect_footnote_refs(&blocks),
            vec![body_ref(1, 10.0), body_ref(2, 20.0)]
        );
    }

    #[test]
    fn recurses_into_table_cells_so_cell_authored_refs_reach_the_page_reservation_pass() {
        // Regression: previously the collector iterated only top-level blocks
        // and skipped `kind: "table"` entirely, so any footnote authored inside
        // a table cell never made it into pageFootnoteMap. The body still
        // rendered the in-line ref marker, but the per-page footnote area
        // dropped the entry — leaving readers with a dangling superscript
        // number.
        let nested_table = LayoutBlock {
            kind: "table".to_string(),
            id: Some(BlockId::Str("t-nested".to_string())),
            rows: Some(vec![TableRow {
                id: Some(BlockId::Str("r-nested".to_string())),
                cells: vec![TableCell {
                    id: Some(BlockId::Str("c-nested".to_string())),
                    blocks: vec![paragraph_with_footnote("nested-p", 8, 200.0)],
                    ..Default::default()
                }],
                ..Default::default()
            }]),
            ..Default::default()
        };
        let table = LayoutBlock {
            kind: "table".to_string(),
            id: Some(BlockId::Str("t1".to_string())),
            rows: Some(vec![TableRow {
                id: Some(BlockId::Str("r1".to_string())),
                cells: vec![
                    TableCell {
                        id: Some(BlockId::Str("c1".to_string())),
                        blocks: vec![paragraph_with_footnote("cell-p1", 7, 100.0)],
                        ..Default::default()
                    },
                    TableCell {
                        id: Some(BlockId::Str("c2".to_string())),
                        blocks: vec![nested_table],
                        ..Default::default()
                    },
                ],
                ..Default::default()
            }]),
            ..Default::default()
        };

        let blocks = vec![
            paragraph_with_footnote("body-p", 1, 10.0),
            table,
            paragraph_with_footnote("trailing-p", 2, 300.0),
        ];

        assert_eq!(
            collect_footnote_refs(&blocks),
            vec![
                body_ref(1, 10.0),
                // In-table refs carry the outermost table id + row index so a
                // split table can distribute its footnotes per page. The
                // nested-table ref (8) keeps the OUTER table's context
                // (t1, row 0), not the inner table's.
                table_ref(7, 100.0, "t1", 0),
                table_ref(8, 200.0, "t1", 0),
                body_ref(2, 300.0),
            ]
        );
    }

    #[test]
    fn recurses_into_text_box_content_blocks() {
        let text_box = LayoutBlock {
            kind: "textBox".to_string(),
            id: Some(BlockId::Str("tb1".to_string())),
            content: Some(vec![paragraph_with_footnote("tb-p", 9, 50.0)]),
            ..Default::default()
        };

        assert_eq!(collect_footnote_refs(&[text_box]), vec![body_ref(9, 50.0)]);
    }

    // ---- mapFootnotesToPages -----------------------------------------------

    #[test]
    fn uses_split_paragraph_fragment_ranges_instead_of_the_whole_paragraph_range() {
        let paragraph_fragment = |pm_start: f64, pm_end: f64| Fragment {
            kind: "paragraph".to_string(),
            block_id: Some(BlockId::Str("p1".to_string())),
            pm_start: Some(pm_start),
            pm_end: Some(pm_end),
            ..Default::default()
        };
        let pages = vec![
            Page {
                number: 1,
                fragments: vec![paragraph_fragment(9.0, 22.0)],
                ..Default::default()
            },
            Page {
                number: 2,
                fragments: vec![paragraph_fragment(22.0, 30.0)],
                ..Default::default()
            },
        ];

        let expected: OrderedMap<u32, Vec<i64>> =
            [(1, vec![1]), (2, vec![2])].into_iter().collect();
        assert_eq!(
            map_footnotes_to_pages(&pages, &[body_ref(1, 16.0), body_ref(2, 22.0)]),
            expected
        );
    }

    #[test]
    fn distributes_a_multi_page_tables_footnotes_by_the_page_holding_each_row() {
        // Regression: a table split across pages keeps the WHOLE table's pm
        // range on every fragment, so pm-position matching dumped all footnote
        // refs on the first table page. Row-index attribution sends each ref to
        // the page that actually laid out its row. Both fragments below
        // deliberately carry the same pm range (5..80) to prove the fix does
        // not rely on it.
        let table_fragment = |row_start: f64, row_end: f64| Fragment {
            kind: "table".to_string(),
            block_id: Some(BlockId::Str("t1".to_string())),
            row_start: Some(row_start),
            row_end: Some(row_end),
            pm_start: Some(5.0),
            pm_end: Some(80.0),
            ..Default::default()
        };

        let pages = vec![
            Page {
                number: 1,
                fragments: vec![table_fragment(0.0, 2.0)],
                ..Default::default()
            },
            Page {
                number: 2,
                fragments: vec![table_fragment(2.0, 4.0)],
                ..Default::default()
            },
        ];

        let refs = vec![
            table_ref(1, 10.0, "t1", 0),
            table_ref(2, 12.0, "t1", 1),
            table_ref(3, 40.0, "t1", 2),
            table_ref(4, 60.0, "t1", 3),
        ];

        let expected: OrderedMap<u32, Vec<i64>> =
            [(1, vec![1, 2]), (2, vec![3, 4])].into_iter().collect();
        assert_eq!(map_footnotes_to_pages(&pages, &refs), expected);
    }

    // ---- distributeFootnotesIntoColumns (footnote-columns-distribute.test.ts)

    #[derive(Clone, Debug, PartialEq)]
    struct Item {
        id: &'static str,
        height: f64,
    }

    impl HasHeight for Item {
        fn height(&self) -> f64 {
            self.height
        }
    }

    fn item(id: &'static str, height: f64) -> Item {
        Item { id, height }
    }

    #[test]
    fn columns_lte_1_returns_a_single_column_unchanged() {
        let items = vec![item("a", 10.0), item("b", 10.0)];
        assert_eq!(
            distribute_footnotes_into_columns(items.clone(), 1.0),
            vec![items.clone()]
        );
        assert_eq!(
            distribute_footnotes_into_columns(items.clone(), 0.0),
            vec![items]
        );
    }

    #[test]
    fn balances_four_equal_items_2_and_2_across_two_columns() {
        let items = vec![
            item("a", 10.0),
            item("b", 10.0),
            item("c", 10.0),
            item("d", 10.0),
        ];
        let cols = distribute_footnotes_into_columns(items, 2.0);
        assert_eq!(cols.len(), 2);
        assert_eq!(
            cols[0].iter().map(|i| i.id).collect::<Vec<_>>(),
            vec!["a", "b"]
        );
        assert_eq!(
            cols[1].iter().map(|i| i.id).collect::<Vec<_>>(),
            vec!["c", "d"]
        );
    }

    #[test]
    fn preserves_document_order_within_and_across_columns() {
        let ids = ["0", "1", "2", "3", "4", "5"];
        let items: Vec<Item> = ids.iter().map(|id| item(id, 10.0)).collect();
        let cols = distribute_footnotes_into_columns(items, 3.0);
        let flattened: Vec<&str> = cols.iter().flatten().map(|i| i.id).collect();
        assert_eq!(flattened, ids);
    }

    #[test]
    fn a_single_tall_note_followed_by_short_notes_does_not_scramble_order() {
        let items = vec![item("tall", 100.0), item("s1", 10.0), item("s2", 10.0)];
        let cols = distribute_footnotes_into_columns(items, 2.0);
        // 'tall' must stay first; short notes follow in order.
        let flattened: Vec<&str> = cols.iter().flatten().map(|i| i.id).collect();
        assert_eq!(flattened, vec!["tall", "s1", "s2"]);
    }

    // ---- calculateFootnoteReservedHeights with columns ----------------------

    fn columns_content_map() -> OrderedMap<i64, HeightItem> {
        [
            (1, HeightItem { height: 10.0 }),
            (2, HeightItem { height: 10.0 }),
            (3, HeightItem { height: 10.0 }),
            (4, HeightItem { height: 10.0 }),
        ]
        .into_iter()
        .collect()
    }

    fn columns_page_map() -> OrderedMap<u32, Vec<i64>> {
        [(1, vec![1, 2, 3, 4])].into_iter().collect()
    }

    #[test]
    fn single_column_reserves_the_full_sum_plus_separator() {
        let reserved =
            calculate_footnote_reserved_heights(&columns_page_map(), &columns_content_map(), 1.0);
        assert_eq!(
            reserved.get(&1).copied(),
            Some(40.0 + FOOTNOTE_SEPARATOR_HEIGHT)
        );
    }

    #[test]
    fn two_columns_reserve_only_the_tallest_balanced_column_plus_separator() {
        let reserved =
            calculate_footnote_reserved_heights(&columns_page_map(), &columns_content_map(), 2.0);
        // 4 x 10 balanced 2-up -> tallest column is 20, not the 40 sum.
        assert_eq!(
            reserved.get(&1).copied(),
            Some(20.0 + FOOTNOTE_SEPARATOR_HEIGHT)
        );
    }
}
