use super::*;

impl SheetPatch<'_> {
    fn equivalent_style_oracle(&self, source: Option<u32>, current: Option<u32>) -> bool {
        if source == current {
            return true;
        }
        let (Some(source), Some(current)) = (source, current) else {
            return false;
        };
        let original = self.styles.original;
        let written = &self.workbook.styles;
        let Some(source_xf) = original.cell_xfs.get(source as usize) else {
            return false;
        };
        if written.cell_xfs.get(source as usize) != Some(source_xf)
            || written.cell_xfs.get(current as usize).is_none()
        {
            return false;
        }
        let format = original.resolved_format(Some(source));
        format == written.resolved_format(Some(source))
            && format == written.resolved_format(Some(current))
    }

    fn same_cell_oracle(&self, source: Option<&Cell>, current: Option<&Cell>) -> bool {
        match (source, current) {
            (Some(source), Some(current)) => {
                source.value == current.value
                    && source.formula == current.formula
                    && self.equivalent_style_oracle(source.style, current.style)
            }
            (None, None) => true,
            _ => false,
        }
    }

    pub(super) fn sheet_data_oracle(&self, source: &[u8]) -> Result<Option<Vec<u8>>, ParseError> {
        let Some((element, rows)) = scan_sheet_data(source)? else {
            return Ok(None);
        };
        let dirty = self.dirty_formulas_oracle(source, &rows)?;
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
            None,
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
            let original = source_cell.and_then(|cell| self.original.cell(cell.at));
            match source_cell {
                Some(source_cell) if self.verbatim_oracle(source_cell, at, original, dirty) => {
                    body.extend_from_slice(&data[source_cell.before.clone()]);
                    self.emit_source_cell_oracle(&mut body, data, source_cell, at)?;
                }
                Some(source_cell) => {
                    let Some(cell) = model_cell else {
                        continue;
                    };
                    body.extend_from_slice(&data[source_cell.before.clone()]);
                    if !self.emit_cached_cell_oracle(
                        &mut body,
                        data,
                        source_cell,
                        at,
                        cell,
                        dirty,
                    )? {
                        self.emit_cell_oracle(&mut body, at, cell, original)?;
                    }
                }
                None => self.emit_cell_oracle(
                    &mut body,
                    at,
                    model_cell.expect("one side is present"),
                    None,
                )?,
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

    fn verbatim_oracle(
        &self,
        cell: &SourceCell,
        at: CellRef,
        original: Option<&Cell>,
        dirty: &DirtyFormulas,
    ) -> bool {
        EMISSION_LOOKUPS.with(|count| count.set(count.get() + 2));
        let model = self.sheet.cell(at);
        if !self.same_cell_oracle(original, model) {
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
        self.formula_clean_oracle(cell, at, dirty)
    }

    fn formula_clean_oracle(&self, cell: &SourceCell, at: CellRef, dirty: &DirtyFormulas) -> bool {
        let Some(formula) = &cell.formula else {
            return true;
        };
        if cell.at != at && formula.positional {
            return false;
        }
        match formula.group {
            Some(group) => !dirty.groups.contains(&group),
            None => !dirty.masters.contains(&(cell.at.row, cell.at.col)),
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
        self.emit_source_content_oracle(out, data, cell, at, cell.span.end)
    }

    fn emit_source_content_oracle(
        &self,
        out: &mut Vec<u8>,
        data: &[u8],
        cell: &SourceCell,
        at: CellRef,
        end: usize,
    ) -> Result<(), ParseError> {
        let mut cursor = cell.tag.end;
        if cell.at != at
            && let Some(formula) = &cell.formula
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
        out.extend_from_slice(&data[cursor..end]);
        Ok(())
    }

    fn emit_cached_cell_oracle(
        &self,
        out: &mut Vec<u8>,
        data: &[u8],
        source: &SourceCell,
        at: CellRef,
        cell: &Cell,
        dirty: &DirtyFormulas,
    ) -> Result<bool, ParseError> {
        let (Some(_), Some(original)) = (&source.formula, self.original.cell(source.at)) else {
            return Ok(false);
        };
        if cell.formula.is_none()
            || cell.formula != original.formula
            || !self.formula_clean_oracle(source, at, dirty)
        {
            return Ok(false);
        }
        let (name, mut attributes) = start_tag(&data[source.tag.clone()])?;
        let Some(slot) = self.cache_slot_oracle(data, source, at)? else {
            return Ok(false);
        };

        let mut writer = Writer::new(Vec::new());
        write_cell(
            &mut writer,
            at,
            cell,
            cell.style,
            self.sst_index,
            None,
            None,
        )
        .map_err(xml_err)?;
        let written = writer.into_inner();
        let mut reader = Reader::from_reader(written.as_slice());
        let ty = match reader.read_event().map_err(xml_err)? {
            Event::Start(element) => attr(&element, b"t")?,
            _ => unreachable!("write_cell emits a cell start tag"),
        };
        let cached =
            cache_span_oracle(&written)?.expect("write_cell emits one formula and cache slot");
        if source.at != at {
            set_attribute(&mut attributes, "r", "r", at.to_a1());
        }
        if !self.equivalent_style_oracle(original.style, cell.style) {
            attributes.retain(|attribute| attribute.name != "s");
            if let Some(style) = cell.style {
                attributes.push(XmlAttribute {
                    name: "s".to_owned(),
                    value: style.to_string(),
                });
            }
        }
        if let Some(ty) = ty {
            match attributes
                .iter_mut()
                .find(|attribute| attribute.name == "t")
            {
                Some(attribute) => attribute.value = ty,
                None => attributes.push(XmlAttribute {
                    name: "t".to_owned(),
                    value: ty,
                }),
            }
        } else {
            attributes.retain(|attribute| attribute.name != "t");
        }
        write_start_tag(out, &name, &attributes, false)?;
        self.emit_source_content_oracle(out, data, source, at, slot.start)?;
        out.extend_from_slice(&written[cached]);
        out.extend_from_slice(&data[slot.end..source.span.end]);
        Ok(true)
    }

    fn cache_slot_oracle(
        &self,
        data: &[u8],
        source: &SourceCell,
        at: CellRef,
    ) -> Result<Option<Range<usize>>, ParseError> {
        let Some(formula) = &source.formula else {
            return Ok(None);
        };
        let (name, cell_attributes) = start_tag(&data[source.tag.clone()])?;
        if name != "c"
            || cell_attributes
                .iter()
                .any(|attribute| attribute.name == "xmlns" || attribute.name.contains(':'))
        {
            return Ok(None);
        }
        let (_, formula_attributes) = start_tag(&data[formula.tag.clone()])?;
        if cell_attributes
            .iter()
            .any(|attribute| matches!(attribute.local_name(), "cm" | "vm"))
            || formula_attributes
                .iter()
                .any(|attribute| attribute.local_name() == "t" && attribute.value == "array")
        {
            let rectangle = self.original.array_formula(source.at);
            if rectangle.is_none()
                || rectangle != formula.reference
                || rectangle != self.sheet.array_formula(at)
            {
                return Ok(None);
            }
        }
        Ok(cache_span_oracle(&data[source.span.clone()])?
            .map(|span| source.span.start + span.start..source.span.start + span.end))
    }

    fn emit_cell_oracle(
        &self,
        out: &mut Vec<u8>,
        at: CellRef,
        cell: &Cell,
        original: Option<&Cell>,
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
        let style = original
            .filter(|source| self.equivalent_style_oracle(source.style, cell.style))
            .map_or(cell.style, |source| source.style);
        write_cell(
            &mut writer,
            at,
            cell,
            style,
            self.sst_index,
            retained,
            self.sheet.array_formula(at),
        )
        .map_err(xml_err)?;
        *out = writer.into_inner();
        Ok(())
    }

    fn dirty_formulas_oracle(
        &self,
        data: &[u8],
        rows: &[SourceRow],
    ) -> Result<DirtyFormulas, ParseError> {
        let mut changed = self.changed_source_cells_oracle();
        let mut arrays = BTreeSet::new();
        for cell in rows.iter().flat_map(|row| &row.cells) {
            if cell.formula.is_none() {
                continue;
            }
            let source_array = self.original.array_formula(cell.at);
            let current_array = self
                .mapped(cell.at)
                .and_then(|at| self.sheet.array_formula(at));
            if source_array.and_then(|range| self.remap_range(range)) != current_array {
                changed.insert((cell.at.row, cell.at.col));
                arrays.insert((cell.at.row, cell.at.col));
            }
        }
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
                    entry.0 |= self
                        .mapped(cell.at)
                        .is_none_or(|at| self.inverse(at) != Some(cell.at))
                        || (changed.contains(&key) && self.source_formula_changed_oracle(cell.at));
                    entry.0 |= arrays.contains(&key);
                    if !entry.0
                        && formula.reference.is_some()
                        && let Some(at) = self.mapped(cell.at)
                    {
                        let moved_positional = cell.at != at && formula.positional;
                        let cache_fallback = changed.contains(&key)
                            && self.cache_slot_oracle(data, cell, at)?.is_none();
                        entry.0 |= moved_positional || cache_fallback;
                    }
                    if let Some(reference) = formula.reference
                        && entry.1.replace(reference).is_some()
                    {
                        entry.0 = true;
                    }
                }
                None => {
                    masters.push((key, formula.reference));
                }
            }
        }
        Ok(DirtyFormulas {
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
                    arrays.contains(key)
                        || (reference.is_some() && changed.contains(key))
                        || reference.is_some_and(|reference| {
                            !self.moves_uniformly(reference) || range_changed(reference, &changed)
                        })
                })
                .map(|(key, _)| key)
                .collect(),
        })
    }

    fn source_formula_changed_oracle(&self, source: CellRef) -> bool {
        let current = self
            .mapped(source)
            .filter(|&at| self.inverse(at) == Some(source))
            .and_then(|at| self.sheet.cell(at));
        match (self.original.cell(source), current) {
            (Some(original), Some(current)) => original.formula != current.formula,
            _ => true,
        }
    }

    pub(super) fn changed_source_cells_oracle(&self) -> BTreeSet<(u32, u32)> {
        let mut changed = BTreeSet::new();
        for (at, cell) in self.original.iter_cells() {
            if self.mapped(at).is_some() {
                DETECTOR_LOOKUPS.with(|count| count.set(count.get() + 1));
            }
            if !self.same_cell_oracle(
                Some(cell),
                self.mapped(at).and_then(|mapped| self.sheet.cell(mapped)),
            ) {
                changed.insert((at.row, at.col));
            }
        }
        for (at, cell) in self.sheet.iter_cells() {
            if self.inverse(at).is_some() {
                DETECTOR_LOOKUPS.with(|count| count.set(count.get() + 1));
            }
            if let Some(source) = self.inverse(at)
                && !self.same_cell_oracle(self.original.cell(source), Some(cell))
            {
                changed.insert((source.row, source.col));
            }
        }
        changed
    }
}

fn cache_span_oracle(cell: &[u8]) -> Result<Option<Range<usize>>, ParseError> {
    let mut reader = Reader::from_reader(cell);
    reader.config_mut().expand_empty_elements = false;
    if !matches!(reader.read_event().map_err(xml_err)?, Event::Start(_)) {
        return Ok(None);
    }
    let mut formulas = Vec::new();
    let mut values = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        let (element, empty) = match reader.read_event().map_err(xml_err)? {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            Event::End(_) | Event::Eof => break,
            _ => continue,
        };
        if !matches!(element.name().as_ref(), b"f" | b"v" | b"is")
            || attributes(&element)?
                .iter()
                .any(|attribute| attribute.name == "xmlns" || attribute.name.starts_with("xmlns:"))
        {
            return Ok(None);
        }
        if element.name().as_ref() == b"is" {
            return Ok(None);
        }
        if !empty {
            loop {
                match reader.read_event().map_err(xml_err)? {
                    Event::Start(_) | Event::Empty(_) => return Ok(None),
                    Event::End(_) => break,
                    Event::Eof => return Ok(None),
                    _ => {}
                }
            }
        }
        let span = start..reader.buffer_position() as usize;
        match element.name().as_ref() {
            b"f" => formulas.push(span),
            b"v" => values.push(span),
            _ => {}
        }
    }
    Ok(match (formulas.as_slice(), values.as_slice()) {
        ([formula], []) => Some(formula.end..formula.end),
        ([formula], [value]) if value.start >= formula.end => Some(value.clone()),
        _ => None,
    })
}
