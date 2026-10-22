use super::*;

#[derive(Default)]
struct SheetDelta {
    rows: BTreeSet<u32>,
    columns: BTreeSet<u32>,
    merges: bool,
    full: bool,
}

#[derive(Default)]
pub(super) struct GeometryDelta {
    sheets: HashMap<SheetId, SheetDelta>,
    full: bool,
}

pub(super) struct GeometryState {
    geometry: Arc<GridGeometry>,
    pending: SheetDelta,
    generation: u64,
}

impl Workbook {
    pub(super) fn geometry_delta(&self, ops: &[Op]) -> GeometryDelta {
        let mut delta = GeometryDelta::default();
        for op in ops {
            match op {
                Op::SetCell { sheet, at, .. } => {
                    delta.sheets.entry(*sheet).or_default().rows.insert(at.row);
                }
                Op::SetRowHeight { sheet, row, .. }
                | Op::SetRowVisibility { sheet, row, .. }
                | Op::RestoreRowDimension { sheet, row, .. } => {
                    delta.sheets.entry(*sheet).or_default().rows.insert(*row);
                }
                Op::SetColWidth { sheet, col, .. }
                | Op::SetColVisibility { sheet, col, .. }
                | Op::RestoreColDimension { sheet, col, .. } => {
                    delta.sheets.entry(*sheet).or_default().columns.insert(*col);
                }
                Op::PatchRangeStyle { sheet, range, .. }
                | Op::SetRangeNumberFormat { sheet, range, .. }
                | Op::ApplyRangeFormat { sheet, range, .. } => {
                    delta
                        .sheets
                        .entry(*sheet)
                        .or_default()
                        .rows
                        .extend(range.start.row..=range.end.row);
                }
                Op::MergeCells { sheet, range } | Op::UnmergeCells { sheet, range } => {
                    let sheet_delta = delta.sheets.entry(*sheet).or_default();
                    sheet_delta.merges = true;
                    let rows = &mut sheet_delta.rows;
                    rows.insert(range.start.row);
                    if let Some(sheet) = self.model.sheet(*sheet) {
                        rows.extend(
                            sheet
                                .merges
                                .iter()
                                .filter(|merged| ranges_intersect(**merged, *range))
                                .map(|merged| merged.start.row),
                        );
                    }
                }
                Op::RestoreRowStyles { sheet, .. } => {
                    delta.sheets.entry(*sheet).or_default().full = true
                }
                Op::RestoreColStyles { .. } => {}
                Op::SetFreezePane { .. }
                | Op::SetHyperlinks { .. }
                | Op::SetHyperlink { .. }
                | Op::RemoveHyperlink { .. }
                | Op::AddConditionalFormat { .. }
                | Op::RemoveConditionalFormat { .. }
                | Op::SetConditionalFormats { .. }
                | Op::SetCharts { .. }
                | Op::SetChartAnchor { .. }
                | Op::SetDefinedName { .. }
                | Op::DeleteDefinedName { .. }
                | Op::SetDefinedNames { .. } => {}
                Op::InsertRows { sheet, .. }
                | Op::DeleteRows { sheet, .. }
                | Op::InsertCols { sheet, .. }
                | Op::DeleteCols { sheet, .. } => {
                    delta.sheets.entry(*sheet).or_default().full = true
                }
                Op::AddSheet { .. }
                | Op::RemoveSheet { .. }
                | Op::MoveSheet { .. }
                | Op::RenameSheet { .. }
                | Op::RestoreSheet { .. } => delta.full = true,
            }
        }
        delta
    }

    pub(super) fn queue_geometry_delta(&mut self, delta: GeometryDelta) {
        let cache = self
            .geometry_cache
            .get_mut()
            .unwrap_or_else(|error| error.into_inner());
        if delta.full {
            cache.clear();
            return;
        }
        for (sheet, delta) in delta.sheets {
            if delta.full {
                cache.remove(&sheet);
                continue;
            }
            if let Some(state) = cache.get_mut(&sheet) {
                state.pending.rows.extend(delta.rows);
                state.pending.columns.extend(delta.columns);
                state.pending.merges |= delta.merges;
            }
        }
    }

    pub(super) fn dirty_geometry_cells(&mut self, cells: &[(SheetId, CellRef)]) {
        let cache = self
            .geometry_cache
            .get_mut()
            .unwrap_or_else(|error| error.into_inner());
        for (sheet, at) in cells {
            if let Some(state) = cache.get_mut(sheet) {
                state.pending.rows.insert(at.row);
            }
        }
        if self
            .model
            .sheet(self.active_sheet)
            .is_some_and(|sheet| sheet.array_formulas().next().is_some())
        {
            self.invalidate_sheet_info();
            return;
        }
        let mut slot = self
            .sheet_info_cache
            .lock()
            .unwrap_or_else(|error| error.into_inner());
        let Some(cache) = slot.as_mut() else {
            return;
        };
        for (sheet, at) in cells {
            if *sheet != self.active_sheet {
                continue;
            }
            if self
                .model
                .sheet(*sheet)
                .and_then(|sheet| sheet.cell(*at))
                .is_some()
            {
                cache.extend_bounds(*at);
            } else if on_used_edge(cache.bounds, *at) {
                *slot = None;
                return;
            }
        }
    }

    pub(super) fn install_model_with_ops(
        &mut self,
        model: WorkbookModel,
        ops: &[Op],
    ) -> Result<()> {
        let mut delta = self.geometry_delta(ops);
        if self.is_collaborative() {
            delta.full = true;
        }
        let old = &self.model.styles;
        let new = &model.styles;
        let normal =
            |styles: &Stylesheet| styles.font_for(0).or_else(|| styles.fonts.first()).cloned();
        delta.full |= !new.fonts.starts_with(&old.fonts)
            || !new.cell_xfs.starts_with(&old.cell_xfs)
            || normal(old) != normal(new)
            || self.model.sheets.len() != model.sheets.len()
            || self
                .model
                .sheets
                .iter()
                .zip(&model.sheets)
                .any(|(old, new)| old.format != new.format);
        self.bump_model_epoch();
        self.model = model;
        self.invalidate_sheet_info();
        self.queue_geometry_delta(delta);
        Ok(())
    }

    pub(super) fn clear_geometry(&mut self) {
        self.geometry_cache
            .get_mut()
            .unwrap_or_else(|error| error.into_inner())
            .clear();
    }

    pub(super) fn sheet_geometry(&self, sheet: SheetId) -> Result<Arc<GridGeometry>> {
        self.sheet_geometry_with_generation(sheet)
            .map(|(geometry, _)| geometry)
    }

    pub(super) fn sheet_geometry_with_generation(
        &self,
        sheet: SheetId,
    ) -> Result<(Arc<GridGeometry>, u64)> {
        let sheet_ref = self.sheet(sheet)?;
        let mut cache = self
            .geometry_cache
            .lock()
            .unwrap_or_else(|error| error.into_inner());
        if let Some(state) = cache.get_mut(&sheet) {
            if !state.pending.rows.is_empty()
                || !state.pending.columns.is_empty()
                || state.pending.merges
            {
                let geometry = Arc::make_mut(&mut state.geometry);
                if std::mem::take(&mut state.pending.merges) {
                    geometry.update_merges(sheet_ref);
                }
                geometry.update_rows(
                    sheet_ref,
                    &self.model.styles,
                    std::mem::take(&mut state.pending.rows),
                );
                geometry.update_columns(sheet_ref, std::mem::take(&mut state.pending.columns));
                state.generation = state.generation.wrapping_add(1);
            }
            return Ok((Arc::clone(&state.geometry), state.generation));
        }
        let geometry = Arc::new(GridGeometry::new(sheet_ref, &self.model.styles));
        cache.insert(
            sheet,
            GeometryState {
                geometry: Arc::clone(&geometry),
                pending: SheetDelta::default(),
                generation: 0,
            },
        );
        Ok((geometry, 0))
    }
}
