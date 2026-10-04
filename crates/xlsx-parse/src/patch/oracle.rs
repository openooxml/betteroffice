use super::*;

impl SheetPatch<'_> {
    pub(super) fn sheet_data_oracle(&self, source: &[u8]) -> Result<Option<Vec<u8>>, ParseError> {
        let Some((element, rows)) = scan_sheet_data(source)? else {
            return Ok(None);
        };
        let dirty = self.dirty_formulas_oracle(&rows);
        let mut out = Vec::with_capacity(source.len());
        out.extend_from_slice(&element.prefix);
        let mut source_rows = rows
            .iter()
            .filter_map(|row| {
                self.axes
                    .rows
                    .current(row.index)
                    .map(|current| (current, row))
            })
            .peekable();
        let mut cells = self.sheet.iter_cells().peekable();
        let mut heights = self.sheet.row_heights.keys().copied().peekable();
        loop {
            let next_source = source_rows.peek().map(|(current, _)| *current);
            let next_cell = cells.peek().map(|(at, _)| at.row);
            let next_height = heights.peek().copied();
            let Some(row) = [next_source, next_cell, next_height]
                .into_iter()
                .flatten()
                .min()
            else {
                break;
            };
            if next_height == Some(row) {
                heights.next();
            }
            if next_source == Some(row) {
                let (_, source_row) = source_rows.next().expect("peeked");
                self.emit_source_row_oracle(&mut out, source, source_row, row, &mut cells, &dirty)?;
            } else {
                self.emit_generated_row_oracle(&mut out, row, &mut cells)?;
            }
        }
        let tail = rows
            .last()
            .map_or(element.content.start, |row| row.span.end);
        out.extend_from_slice(&source[tail..element.content.end]);
        out.extend_from_slice(&element.suffix);
        Ok(Some(out))
    }

    fn emit_generated_row_oracle<'c, I>(
        &self,
        out: &mut Vec<u8>,
        row: u32,
        cells: &mut Peekable<I>,
    ) -> Result<(), ParseError>
    where
        I: Iterator<Item = (CellRef, &'c Cell)>,
    {
        let mut writer = Writer::new(std::mem::take(out));
        write_row(
            &mut writer,
            self.sheet,
            row,
            cells,
            self.workbook,
            self.sst_index,
            self.retained,
            self.plan,
        )
        .map_err(xml_err)?;
        *out = writer.into_inner();
        Ok(())
    }

    pub(super) fn emit_source_row_oracle<'c, I>(
        &self,
        out: &mut Vec<u8>,
        data: &[u8],
        source: &SourceRow,
        row: u32,
        cells: &mut Peekable<I>,
        dirty: &DirtyFormulas,
    ) -> Result<(), ParseError>
    where
        I: Iterator<Item = (CellRef, &'c Cell)>,
    {
        let mut body = Vec::new();
        let mut columns: Option<(u32, u32)> = None;
        let mut source_cells = source
            .cells
            .iter()
            .filter_map(|cell| {
                self.axes
                    .cols
                    .current(cell.at.col)
                    .map(|current| (current, cell))
            })
            .peekable();
        loop {
            let next_source = source_cells.peek().map(|(current, _)| *current);
            let next_model = cells
                .peek()
                .filter(|(at, _)| at.row == row)
                .map(|(at, _)| at.col);
            let Some(col) = [next_source, next_model].into_iter().flatten().min() else {
                break;
            };
            let source_cell =
                (next_source == Some(col)).then(|| source_cells.next().expect("peeked").1);
            let model_cell = (next_model == Some(col)).then(|| cells.next().expect("peeked").1);
            let at = CellRef::new(row, col);
            match source_cell {
                Some(source_cell) if self.verbatim_oracle(source_cell, at, dirty) => {
                    body.extend_from_slice(&data[source_cell.before.clone()]);
                    self.emit_source_cell_oracle(&mut body, data, source_cell, at)?;
                }
                Some(source_cell) => {
                    let Some(cell) = model_cell else {
                        continue;
                    };
                    body.extend_from_slice(&data[source_cell.before.clone()]);
                    self.emit_cell_oracle(&mut body, at, cell)?;
                }
                None => {
                    self.emit_cell_oracle(&mut body, at, model_cell.expect("one side is present"))?
                }
            }
            columns = Some(match columns {
                Some((min, max)) => (min, max.max(col + 1)),
                None => (col + 1, col + 1),
            });
        }

        let moved = source.index != row;
        let height = self.sheet.row_heights.get(&row).copied();
        let height_changed = height != self.original.row_heights.get(&source.index).copied();
        let (name, mut attributes) = start_tag(&data[source.tag.clone()])?;
        let mut rewrite = moved || height_changed;
        if moved {
            set_attribute(&mut attributes, "r", "r", (u64::from(row) + 1).to_string());
        }
        if height_changed {
            set_row_height(&mut attributes, height);
        }
        if let Some((min, max)) = columns
            && spans_exclude(&attributes, min, max)
        {
            set_attribute(&mut attributes, "spans", "spans", format!("{min}:{max}"));
            rewrite = true;
        }

        out.extend_from_slice(&data[source.before.clone()]);
        let empty = source.empty && body.is_empty();
        if rewrite || (source.empty && !empty) {
            write_start_tag(out, &name, &attributes, empty)?;
        } else {
            out.extend_from_slice(&data[source.tag.clone()]);
        }
        if empty {
            return Ok(());
        }
        out.extend_from_slice(&body);
        if source.empty {
            write_end_tag(out, &name)?;
        } else {
            out.extend_from_slice(&data[source.content_start()..source.span.end]);
        }
        Ok(())
    }

    fn verbatim_oracle(&self, cell: &SourceCell, at: CellRef, dirty: &DirtyFormulas) -> bool {
        EMISSION_LOOKUPS.with(|count| count.set(count.get() + 2));
        let model = self.sheet.cell(at);
        if model != self.original.cell(cell.at) {
            return false;
        }
        if cell.shared_string
            && let Some(model) = model
        {
            let Some(source) = self.retained.get(&(at.row, at.col)).copied() else {
                return false;
            };
            let written = shared_string_index(
                model,
                at,
                self.workbook,
                self.sst_index,
                self.retained,
                self.plan,
            );
            if written != Some(source) {
                return false;
            }
        }
        let Some(formula) = &cell.formula else {
            return true;
        };
        if cell.at != at && formula.positional {
            return false;
        }
        match formula.group {
            Some(group) => !dirty.groups.contains(&group),
            None => {
                formula.reference.is_none() || !dirty.masters.contains(&(cell.at.row, cell.at.col))
            }
        }
    }

    fn emit_source_cell_oracle(
        &self,
        out: &mut Vec<u8>,
        data: &[u8],
        cell: &SourceCell,
        at: CellRef,
    ) -> Result<(), ParseError> {
        if cell.at == at {
            out.extend_from_slice(&data[cell.span.clone()]);
            return Ok(());
        }
        let (name, mut attributes) = start_tag(&data[cell.tag.clone()])?;
        set_attribute(&mut attributes, "r", "r", at.to_a1());
        write_start_tag(out, &name, &attributes, cell.empty)?;
        let mut cursor = cell.tag.end;
        if let Some(formula) = &cell.formula
            && let Some(reference) = formula.reference
            && let Some(remapped) = self.remap_range(reference)
            && remapped != reference
        {
            out.extend_from_slice(&data[cursor..formula.tag.start]);
            let (name, mut attributes) = start_tag(&data[formula.tag.clone()])?;
            set_attribute(&mut attributes, "ref", "ref", remapped.to_a1());
            write_start_tag(out, &name, &attributes, formula.empty)?;
            cursor = formula.tag.end;
        }
        out.extend_from_slice(&data[cursor..cell.span.end]);
        Ok(())
    }

    fn emit_cell_oracle(
        &self,
        out: &mut Vec<u8>,
        at: CellRef,
        cell: &Cell,
    ) -> Result<(), ParseError> {
        let retained = shared_string_index(
            cell,
            at,
            self.workbook,
            self.sst_index,
            self.retained,
            self.plan,
        );
        let mut writer = Writer::new(std::mem::take(out));
        write_cell(
            &mut writer,
            at,
            cell,
            self.sst_index,
            retained,
            self.sheet.array_formula(at),
        )
        .map_err(xml_err)?;
        *out = writer.into_inner();
        Ok(())
    }

    fn dirty_formulas_oracle(&self, rows: &[SourceRow]) -> DirtyFormulas {
        let changed = self.changed_source_cells_oracle();
        let mut groups: HashMap<u32, (bool, Option<CellRange>)> = HashMap::new();
        let mut masters = Vec::new();
        for cell in rows.iter().flat_map(|row| &row.cells) {
            let Some(formula) = &cell.formula else {
                continue;
            };
            let key = (cell.at.row, cell.at.col);
            match formula.group {
                Some(group) => {
                    let entry = groups.entry(group).or_insert((false, None));
                    entry.0 |= changed.contains(&key);
                    if let Some(reference) = formula.reference
                        && entry.1.replace(reference).is_some()
                    {
                        entry.0 = true;
                    }
                }
                None => {
                    if let Some(reference) = formula.reference {
                        masters.push((key, reference));
                    }
                }
            }
        }
        DirtyFormulas {
            groups: groups
                .into_iter()
                .filter(|(_, (dirty, master))| {
                    *dirty || !master.is_some_and(|reference| self.moves_uniformly(reference))
                })
                .map(|(group, _)| group)
                .collect(),
            masters: masters
                .into_iter()
                .filter(|(key, reference)| {
                    changed.contains(key)
                        || !self.moves_uniformly(*reference)
                        || range_changed(*reference, &changed)
                })
                .map(|(key, _)| key)
                .collect(),
        }
    }

    pub(super) fn changed_source_cells_oracle(&self) -> BTreeSet<(u32, u32)> {
        let mut changed = BTreeSet::new();
        for (at, cell) in self.original.iter_cells() {
            if self.mapped(at).is_some() {
                DETECTOR_LOOKUPS.with(|count| count.set(count.get() + 1));
            }
            if self.mapped(at).and_then(|mapped| self.sheet.cell(mapped)) != Some(cell) {
                changed.insert((at.row, at.col));
            }
        }
        for (at, cell) in self.sheet.iter_cells() {
            if self.inverse(at).is_some() {
                DETECTOR_LOOKUPS.with(|count| count.set(count.get() + 1));
            }
            if let Some(source) = self.inverse(at)
                && self.original.cell(source) != Some(cell)
            {
                changed.insert((source.row, source.col));
            }
        }
        changed
    }
}
