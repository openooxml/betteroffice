use docx_layout::skipped_effects::{self, CANVAS};

use super::{EditSessionCore, EngineError, EngineResult};

impl EditSessionCore {
    /// The canvas backend's skipped effects on resident page `page_index`, as effect bits.
    pub fn display_canvas_skipped_effects(
        &self,
        page_index: u32,
        unavailable: &[u32],
    ) -> EngineResult<u32> {
        let unavailable = skipped_effects::unpack_positions(unavailable);
        let effects = self.engine.with_display_list(|list| {
            let page = list
                .pages
                .iter()
                .find(|page| page.page_index == u64::from(page_index))?;
            let effects = skipped_effects::skipped_effects(page, &unavailable, &CANVAS);
            Some(skipped_effects::effect_bits(&effects))
        });
        effects
            .flatten()
            .ok_or_else(|| EngineError::message("no resident display page"))
    }

    /// Region-aware hit test against the resident display list, so no
    /// display-list JSON crosses the boundary. `x`/`y` are page-local px.
    /// Returns
    /// `{"region":"body"|"header"|"footer"|"footnote"|"endnote","rId"?,`
    /// `"noteId"?,"pos":n|null,"target":"text"|"image"|"none"}`, or `"null"`
    /// for a page index outside the frame. A header/footer `pos` refers to
    /// that part's own document and a note `pos` to the `fn:{id}` / `en:{id}`
    /// story, never the body; `target` is what sits under the point, which is
    /// what a pointer cursor needs.
    pub fn display_hit_test_regions_json(
        &self,
        page_index: u32,
        x: f64,
        y: f64,
    ) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        self.engine
            .display_hit_test_regions_json(page_index as usize, x, y)
            .map_err(EngineError::message)
    }

    /// Nearest caret position on the adjacent visual line. `direction` is
    /// `"up"` or `"down"`; `goal_x` is the page-local x the caret is trying to
    /// hold across successive moves (a non-finite value falls back to the
    /// caret's own x). Returns `{"position","goalX"}`, or `"null"` when there
    /// is no line in that direction. Errors on an unknown `direction` and when
    /// no display list is resident.
    pub fn display_vertical_move_json(
        &self,
        position: f64,
        direction: &str,
        goal_x: f64,
    ) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        self.engine
            .display_vertical_move_json(position as i64, direction, goal_x)
            .map_err(EngineError::message)
    }

    /// Highlight rectangles for a body document range against the resident
    /// display list: a JSON array of `{pageIndex,x,y,width,height}` in
    /// page-local px, one entry per page the range touches. Body positions
    /// only. Errors when no display list is resident.
    pub fn display_range_rects_json(&self, from: f64, to: f64) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        self.engine
            .display_range_rects_json(from as i64, to as i64)
            .map_err(EngineError::message)
    }

    /// @internal
    pub fn display_range_rects_on_pages_json(
        &self,
        from: f64,
        to: f64,
        first_page: f64,
        last_page: f64,
    ) -> EngineResult<String> {
        let Some((first_page, last_page)) = docx_layout::hit::page_window(first_page, last_page)
        else {
            return Ok("[]".to_string());
        };
        let _fonts = self.fonts.enter();
        self.engine
            .display_range_rects_on_pages_json(from as i64, to as i64, first_page, last_page)
            .map_err(EngineError::message)
    }

    /// Same rectangles as [`EditSessionCore::display_range_rects_json`], scoped to
    /// a region. `region` is `"body"`, `"header"`, `"footer"`, `"footnote"` or
    /// `"endnote"`; `part_id` names one header/footer part (an empty string
    /// matches any) or the note whose story the positions belong to.
    /// `from`/`to` are positions in THAT region's document, and a
    /// header/footer part paints on every page carrying it, so the result
    /// holds one rect set per such page, each tagged with its `pageIndex`.
    /// Errors on an unknown `region` and when no display list is resident.
    pub fn display_range_rects_region_json(
        &self,
        region: &str,
        part_id: &str,
        from: f64,
        to: f64,
    ) -> EngineResult<String> {
        let _fonts = self.fonts.enter();
        self.engine
            .display_range_rects_region_json(region, part_id, from as i64, to as i64)
            .map_err(EngineError::message)
    }

    /// [`EditSessionCore::display_range_rects_region_json`] for many ranges, given as flat
    /// `[from, to, from, to, ...]` pairs. Per range, in order: its rect count, then
    /// `pageIndex, x, y, width, height` for each rect.
    pub fn display_range_rects_batch(
        &self,
        region: &str,
        part_id: &str,
        ranges: Vec<f64>,
    ) -> EngineResult<Vec<f64>> {
        let ranges: Vec<(i64, i64)> = ranges
            .as_chunks::<2>()
            .0
            .iter()
            .map(|&[from, to]| (from as i64, to as i64))
            .collect();
        let batches = self
            .engine
            .display_range_rects_batch(region, part_id, &ranges)
            .map_err(EngineError::message)?;
        let mut flat = Vec::new();
        for rects in batches {
            flat.push(rects.len() as f64);
            for rect in rects {
                flat.extend([
                    rect.page_index as f64,
                    rect.x,
                    rect.y,
                    rect.width,
                    rect.height,
                ]);
            }
        }
        Ok(flat)
    }
}
