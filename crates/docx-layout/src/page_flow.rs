//! Page and column flow state: where the pen is, what room is left, and when
//! to open the next column or page.
//!
//! [`Paginator`] owns the growing `pages` output and one [`FlowState`] per page
//! created; the last state is the current one. No page exists until something
//! asks for the current state, so an empty document produces no pages. Page
//! numbers run `start_page_number + pages.len()`.
//!
//! Vertical room is `content_limit - pen_y`. `content_limit` is the content
//! bottom already reduced by the page's footnote reservation, so reserved note
//! space is simply invisible to body flow. [`Paginator::ensure_fits`] walks to
//! the next column, then the next page, until the height fits; a fragment
//! taller than a whole column is placed with overflow — after moving off a
//! partly used column — rather than looping forever.
//!
//! Word collapses adjacent vertical spacing to the larger of the two instead of
//! summing, so [`Paginator::add_fragment`] takes `max(space_before,
//! deferred_spacing)` and leaves `space_after` deferred for the next fragment.
//! Word spends that collapsed gap from the bottom up: the previous fragment's
//! `space_after` sits below it and only the remainder, `max(0, space_before -
//! space_after)`, sits above the next one. A break discards whatever was below
//! it, so `leading_spacing_spent` carries the discarded `space_after` onto the
//! new page or column and an automatic break spends everything.
//!
//! Columns live in a *region* starting at `column_region_top`. A new page
//! resets that to the content top; [`Paginator::update_columns`] sets it to the
//! deepest column and returns to column zero, so a continuous section break stacks
//! its new column band below content already on the page. Geometry that cannot
//! change mid-sheet is deferred until the next page.

use crate::LayoutError;
use crate::types::{ColumnLayout, Fragment, Page, PageMargins, SectionPageMargins, Size};

/// Complete page-to-page geometry needed to restart placement at a clean
/// page boundary. Cursor/spacing state is intentionally absent: checkpoints
/// are captured only at column zero on a pristine page, where both are fixed
/// by the page geometry.
#[derive(Debug, Clone, PartialEq)]
pub struct PageFlowGeometry {
    pub leading_spacing_spent: f64,
    pub numbering_parity_offset: bool,
    pub page_size: Size,
    pub margins: PageMargins,
    pub columns: ColumnLayout,
    pub pending_page_size: Option<Size>,
    pub pending_margins: Option<PageMargins>,
    pub pending_columns: Option<ColumnLayout>,
    /// Whether the section in force already opened a page before this one.
    pub section_started: bool,
    /// Whether the displayed number this page would continue from the page
    /// before it differs in parity from its physical one.
    pub continued_parity_offset: bool,
    /// Whether this page opened a column region that placement balances.
    pub balanced_region: bool,
}

/// Current state of a page being laid out.
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

/// Splits the content width evenly after subtracting the inter-column gaps.
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

fn effective_margins(margins: PageMargins) -> PageMargins {
    PageMargins {
        top: margins.top.abs(),
        bottom: margins.bottom.abs(),
        ..margins
    }
}

/// The page/column cursor and the pages it has produced so far.
pub struct Paginator {
    /// Leading spacing already accounted for below the break the cursor just
    /// crossed; `f64::INFINITY` for a break that spends all of it.
    leading_spacing_spent: f64,
    /// `leading_spacing_spent` as of the current page's first fragment. Field
    /// resets to 0 once the fragment lands, so checkpoints taken after it are
    /// recorded against this preserved value.
    page_start_spacing_spent: f64,
    /// The column the current page's first fragment landed in.
    page_start_column: usize,
    numbering_parity_offset: bool,
    pub pages: Vec<Page>,
    states: Vec<FlowState>,
    page_size: Size,
    margins: PageMargins,
    columns: ColumnLayout,
    pending_page_size: Option<Size>,
    pending_margins: Option<PageMargins>,
    pending_columns: Option<ColumnLayout>,
    column_width: f64,
    column_region_top: f64,
    column_region_bottom: f64,
    footnote_reserved_heights: Option<std::collections::BTreeMap<String, f64>>,
    start_page_number: u32,
    section_index: usize,
    section_page_margins: Vec<SectionPageMargins>,
    /// Whether the section in force has opened a page yet.
    section_started: bool,
    /// Per state, whether its page is the first of its section.
    opens_section: Vec<bool>,
    /// Whether the current page's displayed number differs in parity from
    /// its physical one, which selects its even band.
    displayed_parity_offset: bool,
    /// [`Self::displayed_parity_offset`] had the current page continued the
    /// previous page's numbering.
    continued_parity_offset: bool,
    /// The page whose column region placement balanced; on a resumed
    /// paginator, the checkpoint's page until placement balances it again.
    balanced_page: Option<usize>,
    /// A resumed checkpoint's deferred geometry, which applies from the page
    /// after the checkpoint's own.
    resumed_pending: Option<(Option<Size>, Option<PageMargins>, Option<ColumnLayout>)>,
}

impl Paginator {
    /// Constructs a paginator with the initial section geometry.
    pub fn new(
        page_size: Size,
        margins: PageMargins,
        columns: ColumnLayout,
        footnote_reserved_heights: Option<std::collections::BTreeMap<String, f64>>,
    ) -> Result<Self, LayoutError> {
        let margins = effective_margins(margins);
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
            leading_spacing_spent: 0.0,
            page_start_spacing_spent: 0.0,
            page_start_column: 0,
            numbering_parity_offset: false,
            pages: Vec::new(),
            states: Vec::new(),
            page_size,
            margins,
            columns,
            pending_page_size: None,
            pending_margins: None,
            pending_columns: None,
            column_width,
            column_region_top,
            column_region_bottom: column_region_top,
            footnote_reserved_heights,
            start_page_number: 1,
            section_index: 0,
            section_page_margins: Vec::new(),
            section_started: false,
            opens_section: Vec::new(),
            displayed_parity_offset: false,
            continued_parity_offset: false,
            balanced_page: None,
            resumed_pending: None,
        })
    }

    /// The margins of the pages that show a first-page or even-page band,
    /// per section; see [`SectionPageMargins`].
    pub fn set_section_page_margins(&mut self, margins: Vec<SectionPageMargins>) {
        self.section_page_margins = margins;
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
        Self::resume_in_section(geometry, start_page_number, 0, footnote_reserved_heights)
    }

    /// [`Paginator::resume`] with section `section_index` already in force.
    pub fn resume_in_section(
        geometry: &PageFlowGeometry,
        start_page_number: u32,
        section_index: usize,
        footnote_reserved_heights: Option<std::collections::BTreeMap<String, f64>>,
    ) -> Result<Self, LayoutError> {
        let mut paginator = Self::new(
            geometry.page_size.clone(),
            geometry.margins.clone(),
            geometry.columns.clone(),
            footnote_reserved_heights,
        )?;
        paginator.resumed_pending = Some((
            geometry.pending_page_size.clone(),
            geometry.pending_margins.clone(),
            geometry.pending_columns.clone(),
        ));
        paginator.start_page_number = start_page_number;
        paginator.leading_spacing_spent = geometry.leading_spacing_spent;
        paginator.numbering_parity_offset = geometry.numbering_parity_offset;
        paginator.section_index = section_index;
        paginator.section_started = geometry.section_started;
        paginator.displayed_parity_offset = geometry.continued_parity_offset;
        paginator.balanced_page = geometry.balanced_region.then_some(0);
        Ok(paginator)
    }

    /// Records that placement balances the column region on the current page.
    pub fn mark_balanced_region(&mut self) {
        let idx = self.get_current();
        self.balanced_page = Some(self.states[idx].page_index);
    }

    /// Whether the current page opened a column region placement balances.
    pub fn balances_region(&mut self) -> bool {
        let idx = self.get_current();
        self.balanced_page == Some(self.states[idx].page_index)
    }

    pub fn restart_page_numbering(&mut self, start: u64) {
        let idx = self.get_current();
        let number = self.pages[self.states[idx].page_index].number;
        self.numbering_parity_offset = start % 2 != u64::from(number % 2);
    }

    pub fn physical_parity_is_odd(&self, number: u64) -> bool {
        (number % 2 != 0) != self.numbering_parity_offset
    }

    /// Sets ownership for future pages and the untouched current page, which
    /// then opens the section.
    pub fn set_section_index(&mut self, section_index: usize) {
        let entering = section_index != self.section_index;
        self.section_index = section_index;
        if entering {
            self.section_started = false;
        }
        if let Some(idx) = self.pristine_page() {
            let page_index = self.states[idx].page_index;
            self.pages[page_index].region_section_index = section_index;
            if entering {
                self.opens_section[idx] = true;
                self.section_started = true;
                self.displayed_parity_offset =
                    self.page_parity_offset(true, self.pages[page_index].number);
                if self.states[idx].column_index == 0 && self.promote_pending_geometry() {
                    self.restamp_pristine_page();
                } else {
                    self.restamp_pristine_band();
                }
            }
        }
    }

    /// Snapshot the page-to-page state. This is sound as a resume bookmark
    /// only when [`Self::clean_page_start`] returns a page.
    pub fn snapshot_geometry(&self) -> PageFlowGeometry {
        PageFlowGeometry {
            leading_spacing_spent: self.leading_spacing_spent,
            numbering_parity_offset: self.numbering_parity_offset,
            page_size: self.page_size.clone(),
            margins: self.margins.clone(),
            columns: self.columns.clone(),
            pending_page_size: self.pending_page_size.clone(),
            pending_margins: self.pending_margins.clone(),
            pending_columns: self.pending_columns.clone(),
            section_started: self
                .opens_section
                .last()
                .map_or(self.section_started, |opens| !opens),
            continued_parity_offset: if self.states.is_empty() {
                self.displayed_parity_offset
            } else {
                self.continued_parity_offset
            },
            balanced_region: self
                .states
                .last()
                .is_some_and(|state| self.balanced_page == Some(state.page_index)),
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
        let mut flow = self.snapshot_geometry();
        if !page.fragments.is_empty() {
            if self.page_start_column != 0 {
                return None;
            }
            flow.leading_spacing_spent = self.page_start_spacing_spent;
        }
        Some((state.page_index, page.number, flow))
    }

    /// Parity offset of a page's displayed number: its section's numbering
    /// restart when it opens the section, else the previous page's plus one.
    fn page_parity_offset(&self, opens_section: bool, page_number: u32) -> bool {
        self.section_page_margins
            .get(self.section_index)
            .and_then(|margins| margins.restart)
            .filter(|_| opens_section)
            .map_or(self.continued_parity_offset, |start| {
                start % 2 != u64::from(page_number % 2)
            })
    }

    /// The margins of a page: its section's, unless the page shows the
    /// first-page or even-page band and that band reserves its own extent.
    fn page_margins(&self, opens_section: bool, page_number: u32) -> PageMargins {
        let variant = self
            .section_page_margins
            .get(self.section_index)
            .and_then(|variants| {
                if opens_section && variants.first.is_some() {
                    variants.first.as_ref()
                } else if (page_number % 2 == 0) != self.displayed_parity_offset {
                    variants.even.as_ref()
                } else {
                    None
                }
            });
        variant.map_or_else(
            || self.margins.clone(),
            |variant| effective_margins(variant.clone()),
        )
    }

    /// Returns the active section's content width.
    pub fn get_content_width(&self) -> f64 {
        self.page_size.w - self.margins.left - self.margins.right
    }

    /// Returns the current column width.
    pub fn column_width(&self) -> f64 {
        self.column_width
    }

    /// Returns the horizontal origin of a column.
    pub fn get_column_x(&self, column_index: usize) -> f64 {
        self.margins.left + column_index as f64 * (self.column_width + self.columns.gap)
    }

    fn footnote_reservation(&self, page_number: u32) -> f64 {
        self.footnote_reserved_heights
            .as_ref()
            .and_then(|m| m.get(&page_number.to_string()).copied())
            .unwrap_or(0.0)
    }

    /// Returns the untouched current page. Column index is excluded because a
    /// standalone column break can leave a blank sheet for a section to reclaim.
    fn pristine_page(&self) -> Option<usize> {
        let idx = self.states.len().checked_sub(1)?;
        let state = &self.states[idx];
        (self.pages[state.page_index].fragments.is_empty() && state.pen_y == state.content_top)
            .then_some(idx)
    }

    /// Re-forms a pristine page when a promoted break remains on the blank sheet.
    /// Callers must next call `update_columns` to refresh its columns and index.
    fn restamp_pristine_page(&mut self) {
        let Some(idx) = self.pristine_page() else {
            return;
        };
        let page_index = self.states[idx].page_index;
        let number = self.pages[page_index].number;
        let margins = self.page_margins(self.opens_section[idx], number);
        let content_top = margins.top;
        let content_limit = self.page_size.h - margins.bottom - self.footnote_reservation(number);
        self.pages[page_index].size = self.page_size.clone();
        self.pages[page_index].margins = margins;
        self.pages[page_index].columns = (self.columns.count > 1.0).then(|| self.columns.clone());
        self.pages[page_index].region_section_index = self.section_index;
        let state = &mut self.states[idx];
        state.pen_y = content_top;
        state.content_top = content_top;
        state.content_limit = content_limit;
        self.column_region_top = content_top;
        self.column_region_bottom = content_top;
    }

    /// Puts deferred page geometry in force; false when none was deferred.
    fn promote_pending_geometry(&mut self) -> bool {
        if self.pending_page_size.is_none()
            && self.pending_margins.is_none()
            && self.pending_columns.is_none()
        {
            return false;
        }
        if let Some(size) = self.pending_page_size.take() {
            self.page_size = size;
        }
        if let Some(margins) = self.pending_margins.take() {
            self.margins = margins;
        }
        if let Some(columns) = self.pending_columns.take() {
            self.columns = columns;
        }
        self.column_width = calculate_column_width(
            self.page_size.w,
            self.margins.left,
            self.margins.right,
            &self.columns,
        );
        true
    }

    /// Re-forms a pristine page whose band changed after it opened.
    fn restamp_pristine_band(&mut self) {
        let Some(idx) = self.pristine_page() else {
            return;
        };
        let page = &self.pages[self.states[idx].page_index];
        if self.page_margins(self.opens_section[idx], page.number) != page.margins {
            self.restamp_pristine_page();
        }
    }

    /// Opens the next page, promoting any deferred geometry first, and returns
    /// its state index.
    fn create_new_page(&mut self) -> usize {
        self.promote_pending_geometry();
        let page_number = self.start_page_number + self.pages.len() as u32;
        let opens_section = !self.section_started;
        self.section_started = true;
        self.continued_parity_offset = self.displayed_parity_offset;
        self.displayed_parity_offset = self.page_parity_offset(opens_section, page_number);
        let margins = self.page_margins(opens_section, page_number);
        let content_top = margins.top;
        let footnote_height = self.footnote_reservation(page_number);
        let page_content_bottom = self.page_size.h - margins.bottom - footnote_height;

        let page = Page {
            number: page_number,
            fragments: Vec::new(),
            margins,
            size: self.page_size.clone(),
            orientation: None,
            section_index: None,
            region_section_index: self.section_index,
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
            section_id: None,
            section_page_index: None,
            section_page_number: None,
            page_label: None,
            page_numbering: None,
            header_distance: None,
            footer_distance: None,
            page_borders: None,
            watermark: None,
            vertical_align: None,
            note_areas: None,
            parity_filler: None,
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
        self.opens_section.push(opens_section);
        if let Some((size, margins, columns)) = self.resumed_pending.take() {
            self.pending_page_size = size;
            self.pending_margins = margins;
            self.pending_columns = columns;
        }

        // reset column region to page top on new page
        self.column_region_top = content_top;
        self.column_region_bottom = content_top;

        self.states.len() - 1
    }

    /// Returns the current state index, creating the first page if needed.
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

    /// Returns the current state's available height.
    pub fn get_available_height(&mut self) -> f64 {
        let idx = self.get_current();
        self.available_height_of(idx)
    }

    fn fits(&self, height: f64, idx: usize) -> bool {
        self.available_height_of(idx) >= height
    }

    /// Moves to the next column of the current region, or opens a new page once
    /// the region's columns are spent, reporting which of the two it did.
    fn advance_column(&mut self, idx: usize) -> (usize, bool) {
        self.leading_spacing_spent = f64::INFINITY;
        if (self.states[idx].column_index as f64) < self.columns.count - 1.0 {
            self.column_region_bottom = self.column_region_bottom.max(self.states[idx].pen_y);
            let region_top = self.column_region_top;
            let state = &mut self.states[idx];
            state.column_index += 1;
            state.pen_y = region_top;
            state.deferred_spacing = 0.0;
            return (idx, false);
        }
        (self.create_new_page(), true)
    }

    /// Advances until a height fits or an oversized fragment can overflow.
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
                    idx = self.advance_column(idx).0;
                }
                return idx;
            }
            idx = self.advance_column(idx).0;
        }

        idx
    }

    /// Places a fragment at the cursor, collapsing adjacent spacing, and
    /// returns its resolved `(x, y)`.
    pub fn add_fragment(
        &mut self,
        mut fragment: Fragment,
        height: f64,
        space_before: f64,
        space_after: f64,
    ) -> (f64, f64) {
        // Read deferred spacing before fitting can advance the state.
        let cur = self.get_current();
        let effective_space_before = self
            .leading_spacing(space_before)
            .max(self.states[cur].deferred_spacing);
        let total_height = effective_space_before + height;

        let idx = self.ensure_fits(total_height);

        let actual_space_before = self
            .leading_spacing(space_before)
            .max(self.states[idx].deferred_spacing);

        let x = self.get_column_x(self.states[idx].column_index);
        let y = self.states[idx].pen_y + actual_space_before;

        fragment.set_xy(x, y);
        let page_index = self.states[idx].page_index;
        self.pages[page_index].fragments.push(fragment);
        if self.pages[page_index].fragments.len() == 1 {
            self.page_start_spacing_spent = self.leading_spacing_spent;
            self.page_start_column = self.states[idx].column_index;
        }

        let state = &mut self.states[idx];
        state.pen_y = y + height;
        state.deferred_spacing = space_after;
        self.leading_spacing_spent = 0.0;

        (x, y)
    }

    /// Forces a page break and is idempotent on a pristine page.
    pub fn force_page_break(&mut self) -> usize {
        self.spend_deferred_spacing();
        match self.pristine_page() {
            Some(idx) => idx,
            None => self.create_new_page(),
        }
    }

    /// Forces an authored page. `keep_leading_spacing` is Word's break rule:
    /// a break authored on the paragraph itself carries its space-before to
    /// the new page, an automatic one spends it.
    pub fn force_authored_page_break(&mut self, keep_leading_spacing: bool) -> usize {
        let index = self.force_page_break();
        if !keep_leading_spacing {
            self.leading_spacing_spent = f64::INFINITY;
        }
        index
    }

    /// Charges the pending space-after against the space-before that follows,
    /// so a break keeps only what the collapsed gap left above it.
    fn spend_deferred_spacing(&mut self) {
        if let Some(state) = self.states.last() {
            self.leading_spacing_spent = self.leading_spacing_spent.max(state.deferred_spacing);
        }
    }

    pub fn leading_spacing(&self, spacing: f64) -> f64 {
        (spacing - self.leading_spacing_spent).max(0.0)
    }

    /// Non-idempotent page creation for the truly blank sheet required by an
    /// evenPage/oddPage section start.
    pub fn insert_blank_page(&mut self) -> usize {
        self.create_new_page()
    }

    /// Marks the current page as an automatic parity filler.
    pub fn mark_parity_filler(&mut self) {
        let idx = self.get_current();
        let page_index = self.states[idx].page_index;
        self.pages[page_index].parity_filler = Some(true);
    }

    /// Moves to the next column, or the next page from the last column.
    pub fn force_column_break(&mut self) -> usize {
        let idx = self.get_current();
        let spent = self.states[idx].deferred_spacing;
        let next = self.advance_column(idx).0;
        self.leading_spacing_spent = spent;
        next
    }

    pub fn advance_for_overflow(&mut self) -> usize {
        let idx = self.get_current();
        self.advance_column(idx).0
    }

    /// Applies a column layout below content already placed in the region.
    pub fn update_columns(&mut self, new_columns: ColumnLayout) {
        self.pending_columns = None;
        self.balanced_page = None;
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

        let page = &self.pages[page_index];
        let content_limit =
            page.size.h - page.margins.bottom - self.footnote_reservation(page.number);
        self.column_region_top = self.column_region_bottom.max(self.states[idx].pen_y);
        self.column_region_bottom = self.column_region_top;
        let state = &mut self.states[idx];
        state.pen_y = self.column_region_top;
        state.column_index = 0;
        state.content_limit = content_limit;
    }

    /// Queues a column layout for the next page, leaving the band in force to
    /// finish the sheet in progress.
    pub fn queue_columns(&mut self, new_columns: ColumnLayout) {
        self.pending_columns = Some(new_columns);
    }

    /// Applies or queues page geometry for subsequently created pages.
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
                self.pending_margins = Some(effective_margins(margins));
            }
            return Ok(());
        }
        if let Some(size) = new_page_size {
            self.page_size = size;
        }
        if let Some(margins) = new_margins {
            self.margins = effective_margins(margins);
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
        self.restamp_pristine_page();
        Ok(())
    }

    /// Pushes a fragment onto the current page without moving the pen.
    pub fn push_fragment_direct(&mut self, fragment: Fragment) {
        let idx = self.get_current();
        let page_index = self.states[idx].page_index;
        self.pages[page_index].fragments.push(fragment);
        if self.pages[page_index].fragments.len() == 1 {
            self.page_start_spacing_spent = self.leading_spacing_spent;
            self.page_start_column = self.states[idx].column_index;
        }
    }

    #[allow(dead_code)] // reached once the floating-table hook is swapped in
    pub fn set_pen_y(&mut self, idx: usize, y: f64) {
        self.states[idx].pen_y = y;
    }

    /// Restarts flow meeting a floating table's band below it, since Word never
    /// paints a page-anchored float over flow content. Declines when the first
    /// fragment's lead clears the band, needing a split this cannot do, or when
    /// the shift would pass the content limit.
    pub fn clear_float_band(&mut self, idx: usize, top: f64, bottom: f64) -> Option<f64> {
        let page_index = self.states[idx].page_index;
        let limit = self.states[idx].content_limit;
        let boxes: Vec<(f64, f64, f64)> = self.pages[page_index]
            .fragments
            .iter()
            .filter_map(Fragment::flow_box)
            .collect();
        let (first, lead) = boxes
            .iter()
            .filter(|(y, height, _)| *y < bottom && y + height > top)
            .map(|(y, _, lead)| (*y, *lead))
            .reduce(|a, b| if b.0 < a.0 { b } else { a })?;
        let delta = bottom - first;
        if delta <= 0.0 || first + lead <= top {
            return None;
        }
        let deepest = boxes
            .iter()
            .filter(|(y, _, _)| *y >= first)
            .map(|(y, height, _)| y + height)
            .fold(f64::NEG_INFINITY, f64::max);
        if deepest + delta > limit || self.states[idx].pen_y + delta > limit {
            return None;
        }
        for fragment in &mut self.pages[page_index].fragments {
            if fragment.flow_box().is_some_and(|(y, _, _)| y >= first) {
                fragment.shift_y(delta);
            }
        }
        self.states[idx].pen_y += delta;
        Some(delta)
    }
}

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

    fn advance_to_next_column(&mut self) -> bool {
        let idx = self.get_current();
        // already at a fresh region start: advancing would only add a blank
        // page, exactly what force_page_break declines to do
        if self.pristine_page() == Some(idx) && self.states[idx].column_index == 0 {
            return true;
        }
        self.advance_column(idx).1
    }

    fn insert_blank_page(&mut self) -> u32 {
        let idx = Paginator::insert_blank_page(self);
        self.pages[self.states[idx].page_index].number
    }

    fn mark_parity_filler(&mut self) {
        Paginator::mark_parity_filler(self);
    }

    fn current_page_size(&mut self) -> Size {
        let idx = self.get_current();
        self.pages[self.states[idx].page_index].size.clone()
    }

    fn current_columns(&self) -> ColumnLayout {
        self.columns.clone()
    }

    fn update_columns(&mut self, columns: &ColumnLayout) {
        Paginator::update_columns(self, columns.clone());
    }

    fn queue_columns(&mut self, columns: &ColumnLayout) {
        Paginator::queue_columns(self, columns.clone());
    }
}

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

#[cfg(test)]
mod tests {
    use super::*;

    fn margins(top: f64, bottom: f64) -> PageMargins {
        PageMargins {
            top,
            right: 96.0,
            bottom,
            left: 96.0,
            header: Some(48.0),
            footer: Some(48.0),
        }
    }

    fn columns() -> ColumnLayout {
        ColumnLayout {
            count: 1.0,
            gap: 0.0,
            equal_width: None,
            separator: None,
            columns: None,
        }
    }

    #[test]
    fn a_pristine_page_entering_a_section_takes_its_geometry_as_a_resume_does() {
        let next = PageMargins {
            left: 144.0,
            ..margins(96.0, 96.0)
        };
        let bands = vec![
            SectionPageMargins::default(),
            SectionPageMargins {
                first: Some(PageMargins {
                    top: 180.0,
                    ..next.clone()
                }),
                even: None,
                restart: None,
            },
        ];
        let size = Size {
            w: 816.0,
            h: 1056.0,
        };
        let mut paginator = Paginator::new(size, margins(96.0, 96.0), columns(), None).unwrap();
        paginator.set_section_page_margins(bands.clone());
        paginator.get_current();
        paginator.force_authored_page_break(false);
        paginator
            .update_page_layout(None, Some(next), false)
            .unwrap();
        paginator.set_section_index(1);
        let (page_index, number, flow) = paginator.clean_page_start().unwrap();
        let full = (
            paginator.pages[page_index].margins.clone(),
            paginator.column_width(),
        );
        assert_eq!((full.0.top, full.0.left), (180.0, 144.0));

        let mut resumed = Paginator::resume_in_section(&flow, number, 1, None).unwrap();
        resumed.set_section_page_margins(bands);
        let idx = resumed.get_current();
        assert_eq!(
            (resumed.pages[idx].margins.clone(), resumed.column_width()),
            full
        );
        assert_eq!(resumed.state(idx).content_top, 180.0);
    }

    fn even_band_paginator() -> Paginator {
        let even = |restart| SectionPageMargins {
            first: None,
            even: Some(margins(180.0, 96.0)),
            restart,
        };
        let size = Size {
            w: 816.0,
            h: 1056.0,
        };
        let mut paginator = Paginator::new(size, margins(96.0, 96.0), columns(), None).unwrap();
        paginator.set_section_page_margins(vec![even(None), even(Some(3)), even(None)]);
        let first = paginator.get_current();
        paginator.set_pen_y(first, 200.0);
        paginator.set_section_index(1);
        paginator
    }

    #[test]
    fn a_restart_decides_even_bands_from_the_first_page_its_section_owns() {
        let mut paginator = even_band_paginator();
        let second = paginator.force_authored_page_break(false);
        assert_eq!(paginator.pages[second].margins.top, 96.0);
        paginator.set_pen_y(second, 200.0);
        let third = paginator.force_authored_page_break(false);
        assert_eq!(paginator.pages[third].margins.top, 180.0);
    }

    #[test]
    fn a_section_that_owns_no_page_does_not_restart_the_numbering() {
        let mut paginator = even_band_paginator();
        let second = paginator.force_authored_page_break(false);
        paginator.set_section_index(2);
        assert_eq!(paginator.pages[second].margins.top, 180.0);
    }

    #[test]
    fn paginator_normalizes_negative_margins_to_effective_origins() {
        let size = Size {
            w: 816.0,
            h: 1056.0,
        };
        let mut paginator = Paginator::new(
            size.clone(),
            margins(-1438.0 / 15.0, -1440.0 / 15.0),
            columns(),
            None,
        )
        .unwrap();
        let idx = paginator.get_current();
        assert_eq!(paginator.pages[idx].margins.top, 1438.0 / 15.0);
        assert_eq!(paginator.pages[idx].margins.bottom, 1440.0 / 15.0);
        assert_eq!(paginator.state(idx).content_top, 1438.0 / 15.0);

        paginator
            .update_page_layout(None, Some(margins(-60.0, 96.0)), true)
            .unwrap();
        let idx = paginator.get_current();
        assert_eq!(paginator.pages[idx].margins.top, 60.0);

        paginator
            .update_page_layout(None, Some(margins(96.0, -70.0)), false)
            .unwrap();
        assert_eq!(
            paginator
                .snapshot_geometry()
                .pending_margins
                .unwrap()
                .bottom,
            70.0
        );
    }
}
