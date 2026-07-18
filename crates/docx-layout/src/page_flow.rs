//! Page-flow state machine — port of
//! `packages/core/src/layout/pagination/pageFlow.ts` (`createPageFlow`).
//!
//! Tracks the page being built, the pen position, and available space, and
//! creates new pages/columns when content doesn't fit. All numeric behavior is
//! f64, in the same operation order as the TS closure. Two deliberate
//! omissions, both output-invariant:
//! - checkpoint capture (`setOnNewPage` / `snapshotGeometry`): resume
//!   bookmarks are derived data, omitted from golden serialization, and the
//!   `resume` option is only exercised by incremental pagination which stays
//!   on the TS side for now;
//! - the oversized-fragment `console.warn` (log only).

use crate::LayoutError;
use crate::types::{ColumnLayout, Fragment, Page, PageMargins, Size};

/// Complete page-to-page geometry needed to restart placement at a clean
/// page boundary. Cursor/spacing state is intentionally absent: checkpoints
/// are captured only at column zero on a pristine page, where both are fixed
/// by the page geometry.
#[derive(Debug, Clone, PartialEq)]
pub struct PageFlowGeometry {
    pub page_size: Size,
    pub margins: PageMargins,
    pub columns: ColumnLayout,
    pub pending_page_size: Option<Size>,
    pub pending_margins: Option<PageMargins>,
}

/// TS `PageFlow` — current state of a page being laid out. `page_index`
/// replaces the TS object reference into `pages`.
#[derive(Debug, Clone)]
pub struct FlowState {
    pub page_index: usize,
    /// Current Y position (cursor) from page top.
    pub pen_y: f64,
    /// Current column index (0-based).
    pub column_index: usize,
    /// Top margin of content area.
    pub content_top: f64,
    /// Bottom boundary of content area (minus any footnote reservation).
    pub content_limit: f64,
    /// Accumulated trailing spacing (space after previous block).
    pub deferred_spacing: f64,
}

/// TS `calculateColumnWidth`.
fn calculate_column_width(
    page_width: f64,
    left_margin: f64,
    right_margin: f64,
    columns: &ColumnLayout,
) -> f64 {
    let content_width = page_width - left_margin - right_margin;
    let total_gaps = (columns.count - 1.0) * columns.gap;
    (content_width - total_gaps) / columns.count
}

/// TS `createPageFlow` return value, as a struct. Methods mirror the closure's
/// functions one-to-one.
pub struct Paginator {
    pub pages: Vec<Page>,
    states: Vec<FlowState>,
    page_size: Size,
    margins: PageMargins,
    columns: ColumnLayout,
    pending_page_size: Option<Size>,
    pending_margins: Option<PageMargins>,
    column_width: f64,
    column_region_top: f64,
    footnote_reserved_heights: Option<std::collections::BTreeMap<String, f64>>,
    start_page_number: u32,
}

impl Paginator {
    /// Construct with the initial section geometry (TS `createPageFlow`
    /// body up to the content-area guard).
    pub fn new(
        page_size: Size,
        margins: PageMargins,
        columns: ColumnLayout,
        footnote_reserved_heights: Option<std::collections::BTreeMap<String, f64>>,
    ) -> Result<Self, LayoutError> {
        let content_height = (page_size.h - margins.bottom) - margins.top;
        if content_height <= 0.0 {
            return Err(LayoutError::Invalid(
                "Paginator: page size and margins yield no content area".into(),
            ));
        }
        let column_width =
            calculate_column_width(page_size.w, margins.left, margins.right, &columns);
        let column_region_top = margins.top;
        Ok(Paginator {
            pages: Vec::new(),
            states: Vec::new(),
            page_size,
            margins,
            columns,
            pending_page_size: None,
            pending_margins: None,
            column_width,
            column_region_top,
            footnote_reserved_heights,
            start_page_number: 1,
        })
    }

    /// Restore a paginator at a clean page start. The first lazily-created
    /// page uses `start_page_number`; no prefix pages are copied into this
    /// instance, so callers can splice the resulting suffix onto retained
    /// pages without walking them again.
    pub fn resume(
        geometry: &PageFlowGeometry,
        start_page_number: u32,
        footnote_reserved_heights: Option<std::collections::BTreeMap<String, f64>>,
    ) -> Result<Self, LayoutError> {
        let mut paginator = Self::new(
            geometry.page_size.clone(),
            geometry.margins.clone(),
            geometry.columns.clone(),
            footnote_reserved_heights,
        )?;
        paginator.pending_page_size = geometry.pending_page_size.clone();
        paginator.pending_margins = geometry.pending_margins.clone();
        paginator.start_page_number = start_page_number;
        Ok(paginator)
    }

    /// Snapshot the page-to-page state. This is sound as a resume bookmark
    /// only when [`Self::clean_page_start`] returns a page.
    pub fn snapshot_geometry(&self) -> PageFlowGeometry {
        PageFlowGeometry {
            page_size: self.page_size.clone(),
            margins: self.margins.clone(),
            columns: self.columns.clone(),
            pending_page_size: self.pending_page_size.clone(),
            pending_margins: self.pending_margins.clone(),
        }
    }

    /// Ensure a current page exists and return its local index and number when
    /// placement is exactly at a resumable page start.
    pub fn clean_page_start(&mut self) -> Option<(usize, u32, PageFlowGeometry)> {
        let state_index = self.get_current();
        let state = &self.states[state_index];
        let page = &self.pages[state.page_index];
        (state.column_index == 0
            && state.pen_y == state.content_top
            && state.deferred_spacing == 0.0
            && page.fragments.is_empty())
        .then(|| (state.page_index, page.number, self.snapshot_geometry()))
    }

    /// Fragment counts used by the placement walk to recognize a block that
    /// was moved wholesale onto a newly-created clean page.
    pub fn page_fragment_counts(&self) -> Vec<usize> {
        self.pages.iter().map(|page| page.fragments.len()).collect()
    }

    /// Geometry of the current page after it was created during placement.
    /// Checkpoint discovery only calls this for the current page.
    pub fn current_page_start(&self) -> Option<(usize, u32, PageFlowGeometry)> {
        let state = self.states.last()?;
        let page = self.pages.get(state.page_index)?;
        Some((state.page_index, page.number, self.snapshot_geometry()))
    }

    fn get_content_bottom(&self) -> f64 {
        self.page_size.h - self.margins.bottom
    }

    /// TS `getContentWidth` — content width for the active section.
    pub fn get_content_width(&self) -> f64 {
        self.page_size.w - self.margins.left - self.margins.right
    }

    /// Current column width (TS `columnWidth` getter).
    pub fn column_width(&self) -> f64 {
        self.column_width
    }

    /// TS `getColumnX`.
    pub fn get_column_x(&self, column_index: usize) -> f64 {
        self.margins.left + column_index as f64 * (self.column_width + self.columns.gap)
    }

    /// TS `createNewPage`. Returns the new state's index.
    fn create_new_page(&mut self) -> usize {
        // apply any geometry queued by a continuous section break before
        // computing the new page's size / margins
        if self.pending_page_size.is_some() || self.pending_margins.is_some() {
            if let Some(size) = self.pending_page_size.take() {
                self.page_size = size;
            }
            if let Some(margins) = self.pending_margins.take() {
                self.margins = margins;
            }
            self.column_width = calculate_column_width(
                self.page_size.w,
                self.margins.left,
                self.margins.right,
                &self.columns,
            );
        }
        let page_number = self.start_page_number + self.pages.len() as u32;
        let content_top = self.margins.top;
        let content_limit = self.get_content_bottom();

        // reduce content bottom by the footnote reserved height for this page
        let footnote_height = self
            .footnote_reserved_heights
            .as_ref()
            .and_then(|m| m.get(&page_number.to_string()).copied())
            .unwrap_or(0.0);
        let page_content_bottom = content_limit - footnote_height;

        let page = Page {
            number: page_number,
            fragments: Vec::new(),
            margins: self.margins.clone(),
            size: self.page_size.clone(),
            orientation: None,
            section_index: None,
            header_footer_refs: None,
            footnote_ids: None,
            footnote_reserved_height: if footnote_height > 0.0 {
                Some(footnote_height)
            } else {
                None
            },
            footnote_columns: None,
            // initial columns; may be overwritten by update_columns() for
            // continuous section breaks
            columns: if self.columns.count > 1.0 {
                Some(self.columns.clone())
            } else {
                None
            },
        };

        let state = FlowState {
            page_index: self.pages.len(),
            pen_y: content_top,
            column_index: 0,
            content_top,
            content_limit: page_content_bottom,
            deferred_spacing: 0.0,
        };

        self.pages.push(page);
        self.states.push(state);

        // reset column region to page top on new page
        self.column_region_top = content_top;

        self.states.len() - 1
    }

    /// TS `getCurrentState` — index of the current state, creating page 1 if
    /// none exists.
    pub fn get_current(&mut self) -> usize {
        if self.states.is_empty() {
            return self.create_new_page();
        }
        self.states.len() - 1
    }

    /// Read a state by index.
    pub fn state(&self, idx: usize) -> &FlowState {
        &self.states[idx]
    }

    /// Number of fragments already on the state's page.
    pub fn page_fragment_count(&self, idx: usize) -> usize {
        self.pages[self.states[idx].page_index].fragments.len()
    }

    fn available_height_of(&self, idx: usize) -> f64 {
        let s = &self.states[idx];
        s.content_limit - s.pen_y
    }

    /// TS `getAvailableHeight()` on the current state.
    pub fn get_available_height(&mut self) -> f64 {
        let idx = self.get_current();
        self.available_height_of(idx)
    }

    fn fits(&self, height: f64, idx: usize) -> bool {
        self.available_height_of(idx) >= height
    }

    /// TS `advanceColumn` — next column, or a new page when columns are spent.
    fn advance_column(&mut self, idx: usize) -> usize {
        if (self.states[idx].column_index as f64) < self.columns.count - 1.0 {
            let region_top = self.column_region_top;
            let state = &mut self.states[idx];
            state.column_index += 1;
            state.pen_y = region_top;
            state.deferred_spacing = 0.0;
            return idx;
        }
        self.create_new_page()
    }

    /// TS `ensureFits` — advance column/page until `height` fits; oversized
    /// fragments are placed with overflow rather than looping forever.
    pub fn ensure_fits(&mut self, height: f64) -> usize {
        let mut idx = self.get_current();
        let safe_height = if height.is_finite() && height > 0.0 {
            height
        } else {
            0.0
        };

        while !self.fits(safe_height, idx) {
            // oversized-fragment guard, re-checked per iteration because a
            // queued continuous-section geometry can change page capacity
            let column_capacity = self.states[idx].content_limit - self.states[idx].content_top;
            if safe_height > column_capacity {
                if self.states[idx].pen_y != self.states[idx].content_top {
                    idx = self.advance_column(idx);
                }
                return idx;
            }
            idx = self.advance_column(idx);
        }

        idx
    }

    /// TS `addFragment` — position the fragment at the cursor (collapsing
    /// adjacent spacing to the larger of the two), push it, advance the pen.
    /// Returns `(x, y)`.
    pub fn add_fragment(
        &mut self,
        mut fragment: Fragment,
        height: f64,
        space_before: f64,
        space_after: f64,
    ) -> (f64, f64) {
        // Word collapses spaceAfter / next.spaceBefore to the larger of the
        // two (CSS-style margin-collapse), not the sum. NOTE: read from the
        // CURRENT state before ensureFits, exactly like the TS.
        let cur = self.get_current();
        let effective_space_before = space_before.max(self.states[cur].deferred_spacing);
        let total_height = effective_space_before + height;

        let idx = self.ensure_fits(total_height);

        // Word 2013+ honors an explicit w:before at the top of a page/column;
        // deferred spacing was already reset when the page/column started.
        let actual_space_before = effective_space_before;

        let x = self.get_column_x(self.states[idx].column_index);
        let y = self.states[idx].pen_y + actual_space_before;

        fragment.set_xy(x, y);
        let page_index = self.states[idx].page_index;
        self.pages[page_index].fragments.push(fragment);

        let state = &mut self.states[idx];
        state.pen_y = y + height;
        state.deferred_spacing = space_after;

        (x, y)
    }

    /// TS `forcePageBreak` — idempotent on a pristine page.
    pub fn force_page_break(&mut self) -> usize {
        if let Some(idx) = self.states.len().checked_sub(1) {
            let current = &self.states[idx];
            if self.pages[current.page_index].fragments.is_empty()
                && current.pen_y == current.content_top
            {
                return idx;
            }
        }
        self.create_new_page()
    }

    /// Non-idempotent page creation for the truly blank sheet required by an
    /// evenPage/oddPage section start.
    pub fn insert_blank_page(&mut self) -> usize {
        self.create_new_page()
    }

    /// TS `forceColumnBreak`.
    pub fn force_column_break(&mut self) -> usize {
        let idx = self.get_current();
        self.advance_column(idx)
    }

    /// TS `updateColumns` — swap the column layout mid-document. The new
    /// column region starts at the current pen so continuous breaks keep the
    /// band below existing content.
    pub fn update_columns(&mut self, new_columns: ColumnLayout) {
        self.columns = new_columns;
        self.column_width = calculate_column_width(
            self.page_size.w,
            self.margins.left,
            self.margins.right,
            &self.columns,
        );

        let idx = self.get_current();
        let page_index = self.states[idx].page_index;
        self.pages[page_index].columns = if self.columns.count > 1.0 {
            Some(self.columns.clone())
        } else {
            None
        };

        self.column_region_top = self.states[idx].pen_y;
        self.states[idx].column_index = 0;
    }

    /// TS `updatePageLayout` — swap (or queue, for continuous breaks) the page
    /// geometry used by subsequently created pages.
    pub fn update_page_layout(
        &mut self,
        new_page_size: Option<Size>,
        new_margins: Option<PageMargins>,
        apply_immediately: bool,
    ) -> Result<(), LayoutError> {
        if !apply_immediately {
            if let Some(size) = new_page_size {
                self.pending_page_size = Some(size);
            }
            if let Some(margins) = new_margins {
                self.pending_margins = Some(margins);
            }
            return Ok(());
        }
        if let Some(size) = new_page_size {
            self.page_size = size;
        }
        if let Some(margins) = new_margins {
            self.margins = margins;
        }
        if (self.page_size.h - self.margins.bottom) - self.margins.top <= 0.0 {
            return Err(LayoutError::Invalid(
                "Paginator: section page size and margins yield no content area".into(),
            ));
        }
        self.column_width = calculate_column_width(
            self.page_size.w,
            self.margins.left,
            self.margins.right,
            &self.columns,
        );
        // a pending swap is superseded by this immediate swap
        self.pending_page_size = None;
        self.pending_margins = None;
        Ok(())
    }

    /// Push a fragment straight onto the current page without moving the pen
    /// (TS sites do `state.page.fragments.push(fragment)` for anchored /
    /// floating content).
    pub fn push_fragment_direct(&mut self, fragment: Fragment) {
        let idx = self.get_current();
        let page_index = self.states[idx].page_index;
        self.pages[page_index].fragments.push(fragment);
    }

    /// Raise the current pen to `y` (TS sites assign `state.penY` directly).
    #[allow(dead_code)] // reached once the floating-table hook is swapped in
    pub fn set_pen_y(&mut self, idx: usize, y: f64) {
        self.states[idx].pen_y = y;
    }
}

/// The section-break slice of the paginator (`sectionBreaks.ts` calls exactly
/// these four `pageFlow.ts` methods).
impl crate::section_breaks::SectionBreakPaginator for Paginator {
    fn update_page_layout(
        &mut self,
        page_size: Option<&Size>,
        margins: Option<&PageMargins>,
        apply_immediately: bool,
    ) -> Result<(), LayoutError> {
        Paginator::update_page_layout(
            self,
            page_size.cloned(),
            margins.cloned(),
            apply_immediately,
        )
    }

    fn force_page_break(&mut self) -> u32 {
        let idx = Paginator::force_page_break(self);
        self.pages[self.states[idx].page_index].number
    }

    fn insert_blank_page(&mut self) -> u32 {
        let idx = Paginator::insert_blank_page(self);
        self.pages[self.states[idx].page_index].number
    }

    fn current_page_size(&mut self) -> Size {
        let idx = self.get_current();
        self.pages[self.states[idx].page_index].size.clone()
    }

    fn update_columns(&mut self, columns: &ColumnLayout) {
        Paginator::update_columns(self, columns.clone());
    }
}

/// The column-balancing slice of the paginator (`columnBalancing.ts` reads
/// `columns` and the current state's `penY`/`contentLimit`, and writes
/// `contentLimit` back).
impl crate::column_balancing::ColumnBalancePaginator for Paginator {
    fn columns(&self) -> ColumnLayout {
        self.columns.clone()
    }

    fn pen_y(&mut self) -> f64 {
        let idx = self.get_current();
        self.states[idx].pen_y
    }

    fn content_limit(&mut self) -> f64 {
        let idx = self.get_current();
        self.states[idx].content_limit
    }

    fn set_content_limit(&mut self, value: f64) {
        let idx = self.get_current();
        self.states[idx].content_limit = value;
    }
}
