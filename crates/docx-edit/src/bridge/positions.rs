use super::*;

pub(super) struct PositionShift {
    boundary: u64,
    delta: i64,
    apply: bool,
}

impl PositionShift {
    pub(super) fn new(boundary: u64, delta: i64, apply: bool) -> Self {
        Self {
            boundary,
            delta,
            apply,
        }
    }

    fn position(&self, position: &mut f64) -> Option<()> {
        if !position.is_finite() || position.fract() != 0.0 || *position < 0.0 {
            return None;
        }
        if *position >= self.boundary as f64 {
            let next = *position + self.delta as f64;
            if !(0.0..=9_007_199_254_740_991.0).contains(&next) {
                return None;
            }
            if self.apply {
                *position = next;
            }
        }
        Some(())
    }

    fn pair(&self, start: &mut Option<f64>, end: &mut Option<f64>) -> Option<()> {
        for position in [start, end].into_iter().flatten() {
            self.position(position)?;
        }
        Some(())
    }

    fn id(&self, id: &mut String, prefix: &str, suffix: bool) -> Option<()> {
        let tail = id.strip_prefix(prefix)?;
        let (number, rest) = tail
            .split_once(':')
            .map_or((tail, ""), |(number, _)| (number, &tail[number.len()..]));
        if !suffix && !rest.is_empty() {
            return None;
        }
        let position: u64 = number.parse().ok()?;
        if position >= self.boundary {
            let next = position.checked_add_signed(self.delta)?;
            if self.apply {
                *id = format!("{prefix}{next}{rest}");
            }
        }
        Some(())
    }

    fn shape_id(&self, id: &mut BlockId, suffix: bool) -> Option<()> {
        let BlockId::Str(id) = id else {
            return None;
        };
        self.id(id, "shape:", suffix)
    }

    fn groups(&self, groups: &mut Option<Vec<SdtGroup>>) -> Option<()> {
        for group in groups.iter_mut().flatten() {
            self.id(&mut group.id, "sdt@", false)?;
            if let Some(position) = &mut group.pos {
                let value = u64::try_from(*position).ok()?;
                if value >= self.boundary {
                    let next = value.checked_add_signed(self.delta)?;
                    let next = i64::try_from(next).ok()?;
                    if self.apply {
                        *position = next;
                    }
                }
            }
        }
        Some(())
    }

    fn widget(&self, widget: &mut Option<Value>) -> Option<()> {
        let Some(widget) = widget else {
            return Some(());
        };
        let object = widget.as_object_mut()?;
        if object.get("kind")?.as_str()? != "checkbox" {
            return None;
        }
        let position = object.get("pos")?.as_u64()?;
        if object.get("groupId")?.as_str()? != format!("sdt@{position}") {
            return None;
        }
        if position >= self.boundary {
            let next = position.checked_add_signed(self.delta)?;
            if self.apply {
                object.insert("pos".to_owned(), Value::from(next));
                object.insert("groupId".to_owned(), Value::String(format!("sdt@{next}")));
            }
        }
        Some(())
    }

    fn paragraph(&self, paragraph: &mut ParagraphBlock) -> Option<()> {
        self.groups(&mut paragraph.sdt_groups)?;
        self.pair(&mut paragraph.pm_start, &mut paragraph.pm_end)?;
        if let Some(attrs) = &mut paragraph.attrs {
            for rule in &mut attrs.horizontal_rules {
                self.position(&mut rule.pm_start)?;
                self.position(&mut rule.pm_end)?;
            }
        }
        for run in &mut paragraph.runs {
            match run {
                Run::Text(run) => {
                    self.pair(&mut run.pm_start, &mut run.pm_end)?;
                    self.widget(&mut run.inline_sdt_widget)?;
                }
                Run::Tab(run) => self.pair(&mut run.pm_start, &mut run.pm_end)?,
                Run::LineBreak(run) => self.pair(&mut run.pm_start, &mut run.pm_end)?,
                Run::Field(run) => self.pair(&mut run.pm_start, &mut run.pm_end)?,
                Run::Image(run) => {
                    self.pair(&mut run.pm_start, &mut run.pm_end)?;
                    if let Some(shape) = &mut run.inline_shape {
                        self.shape(shape, false)?;
                    }
                }
                Run::Unsupported => return None,
            }
        }
        Some(())
    }

    fn shape(&self, shape: &mut ShapeBlock, child: bool) -> Option<()> {
        self.shape_id(&mut shape.id, child)?;
        self.groups(&mut shape.sdt_groups)?;
        self.pair(&mut shape.pm_start, &mut shape.pm_end)?;
        self.pair(&mut shape.doc_start, &mut shape.doc_end)?;
        for paragraph in shape.inner_text.iter_mut().flatten() {
            self.shape_id(&mut paragraph.id, true)?;
            self.paragraph(paragraph)?;
        }
        for child in &mut shape.children {
            self.shape(child, true)?;
        }
        Some(())
    }

    pub(super) fn map(&self, map: &mut LoweringMap) -> Option<()> {
        let start = map
            .paragraph_blocks
            .partition_point(|(pm, _)| *pm < self.boundary);
        let table_start = map.tables.partition_point(|(pm, ..)| *pm < self.boundary);
        for pm in map.paragraph_blocks[start..]
            .iter_mut()
            .map(|(pm, _)| pm)
            .chain(map.tables[table_start..].iter_mut().map(|(pm, ..)| pm))
        {
            let next = pm.checked_add_signed(self.delta)?;
            if self.apply {
                *pm = next;
            }
        }
        let start = map
            .spans
            .partition_point(|span| span.pm_start < self.boundary);
        for span in &mut map.spans[start..] {
            let pm_start = span.pm_start.checked_add_signed(self.delta)?;
            let pm_end = span.pm_end.checked_add_signed(self.delta)?;
            if map.paragraphs[span.paragraph as usize].0 == 0 {
                let raw_start = u32::try_from(i64::from(span.raw_start) + self.delta).ok()?;
                let raw_end = u32::try_from(i64::from(span.raw_end) + self.delta).ok()?;
                if self.apply {
                    span.raw_start = raw_start;
                    span.raw_end = raw_end;
                }
            }
            if self.apply {
                span.pm_start = pm_start;
                span.pm_end = pm_end;
            }
        }
        Some(())
    }

    pub(super) fn block(&self, block: &mut LayoutBlock) -> Option<()> {
        match block {
            LayoutBlock::Paragraph(paragraph) => self.paragraph(paragraph)?,
            LayoutBlock::Table(table) => {
                self.groups(&mut table.sdt_groups)?;
                self.pair(&mut table.pm_start, &mut table.pm_end)?;
                for row in &mut table.rows {
                    for cell in &mut row.cells {
                        for block in &mut cell.blocks {
                            self.block(block)?;
                        }
                    }
                }
            }
            LayoutBlock::Image(image) => {
                self.groups(&mut image.sdt_groups)?;
                self.pair(&mut image.pm_start, &mut image.pm_end)?;
            }
            LayoutBlock::Shape(shape) => self.shape(shape, false)?,
            LayoutBlock::Chart(chart) => {
                let BlockId::Str(id) = &mut chart.id else {
                    return None;
                };
                self.id(id, "chart:", false)?;
                self.groups(&mut chart.sdt_groups)?;
                self.pair(&mut chart.pm_start, &mut chart.pm_end)?;
                self.pair(&mut chart.doc_start, &mut chart.doc_end)?;
            }
            LayoutBlock::TextBox(text_box) => {
                self.groups(&mut text_box.sdt_groups)?;
                self.pair(&mut text_box.pm_start, &mut text_box.pm_end)?;
                for paragraph in &mut text_box.content {
                    self.paragraph(paragraph)?;
                }
            }
            LayoutBlock::SectionBreak(section) => self.groups(&mut section.sdt_groups)?,
            LayoutBlock::PageBreak(page) => {
                self.groups(&mut page.sdt_groups)?;
                self.pair(&mut page.pm_start, &mut page.pm_end)?;
            }
            LayoutBlock::ColumnBreak(column) => {
                self.groups(&mut column.sdt_groups)?;
                self.pair(&mut column.pm_start, &mut column.pm_end)?;
            }
            LayoutBlock::Unsupported => return None,
        }
        Some(())
    }
}
