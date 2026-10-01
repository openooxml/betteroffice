use super::*;

#[derive(Debug, Default)]
pub(crate) struct LocalLowering {
    pub(super) blocked: bool,
    pub(super) source: std::sync::Weak<crate::seed::SourceMetadata>,
    pub(super) seeds: BTreeMap<String, ParagraphSeed>,
    pub(crate) edit: Option<TextEdit>,
    excluded: BTreeSet<String>,
    position_ids: BTreeMap<String, (String, u64)>,
    table_ids: BTreeMap<String, (String, u32, u64)>,
}

#[derive(Debug, Default)]
pub(super) struct ParagraphSeed {
    ineligible: bool,
    raw_start: u32,
    slot: usize,
    source: u32,
    segments: Vec<TextSegment>,
    pilcrow: Option<MapRef>,
    mark_attrs: Option<Attrs>,
    pm_start: u64,
}

#[derive(Clone, Debug)]
struct TextSegment {
    text: String,
    attrs: Attrs,
}

#[derive(Debug)]
pub(crate) struct TextEdit {
    pub(crate) paragraph: String,
    pub(crate) offset: u32,
    pub(crate) removed: u32,
    pub(crate) text: String,
    pub(crate) attributes: Option<Attrs>,
    pub(crate) epochs: (u64, u64),
}

fn attributes(attrs: Option<&Attrs>) -> Attrs {
    let mut attrs = attrs.cloned().unwrap_or_default();
    attrs.retain(|_, value| *value != Any::Null);
    attrs
}

fn same_attributes(attrs: &Attrs, other: Option<&Attrs>) -> bool {
    let other = || {
        other
            .into_iter()
            .flatten()
            .filter(|(_, value)| **value != Any::Null)
    };
    other().count() == attrs.len() && other().all(|(key, value)| attrs.get(key) == Some(value))
}

fn unsafe_value(key: &str, value: &Any) -> bool {
    if matches!(value, Any::Null | Any::Undefined | Any::Bool(false)) {
        return false;
    }
    [INS, DEL, "pPrIns", "pPrDel", "pPrChange"].contains(&key)
        || ["trPrChange", "tcPrChange"].contains(&key)
        || ["trackedIns", "trackedDel", "trIns", "trDel"].contains(&key)
        || ["hyperlink", "bookmarks", "comment", "cellMarker"].contains(&key)
        || ["hidden", "vanish", "contextualSpacing", "floating"].contains(&key)
        || match value {
            Any::Map(map) => map.iter().any(|(key, value)| unsafe_value(key, value)),
            Any::Array(values) => values.iter().any(|value| unsafe_value("", value)),
            _ => false,
        }
}

impl LocalLowering {
    pub(crate) fn new(enabled: bool) -> Self {
        Self {
            blocked: !enabled,
            ..Default::default()
        }
    }

    pub(super) fn observe<T: ReadTxn>(
        &mut self,
        paragraph: &mut ParagraphSeed,
        diff: &yrs::types::text::Diff<YChange>,
        txn: &T,
        story: &str,
    ) {
        if self.blocked {
            return;
        }
        let attrs = diff.attributes.as_deref();
        paragraph.ineligible |= attrs
            .into_iter()
            .flatten()
            .any(|(key, value)| unsafe_value(key, value));
        match &diff.insert {
            Out::Any(Any::String(_)) if story != "body" => {}
            Out::Any(Any::String(text)) => {
                if let Some(last) = paragraph.segments.last_mut()
                    && same_attributes(&last.attrs, attrs)
                {
                    last.text.push_str(text);
                } else {
                    paragraph.segments.push(TextSegment {
                        text: text.to_string(),
                        attrs: attributes(attrs),
                    });
                }
            }
            Out::YMap(mark) if is_pilcrow(mark, txn) => {}
            Out::YMap(mark) => {
                let values = pilcrow_values(mark, txn);
                if matches!(
                    value_string(values.get("_kind")).as_deref(),
                    Some("table" | "blockSdt" | "pageBreak" | "columnBreak")
                ) {
                    *paragraph = ParagraphSeed::default();
                } else {
                    paragraph.ineligible = true;
                }
            }
            _ => paragraph.ineligible = true,
        }
    }

    #[allow(clippy::too_many_arguments)]
    pub(super) fn observe_pilcrow(
        &mut self,
        paragraph: &mut ParagraphSeed,
        mark: &MapRef,
        values: &BTreeMap<String, Any>,
        attrs: Option<&Attrs>,
        story: &str,
        (start, pm_start, slot, source): (u32, u64, usize, u32),
        comments: &[CommentInterval],
        end: u32,
    ) {
        if self.blocked {
            return;
        }
        let id = value_string(values.get("paraId")).unwrap_or_default();
        paragraph.ineligible |= values.iter().any(|(key, value)| unsafe_value(key, value))
            || values.contains_key("sectPr")
            || values.contains_key("sectionBreakType")
            || comments.iter().any(|interval| {
                interval.start.min(interval.end) <= end && interval.start.max(interval.end) >= start
            })
            || self
                .source
                .upgrade()
                .is_some_and(|source| source.run_revision(story, &id));
        if story == "body" && !paragraph.ineligible {
            paragraph.raw_start = start;
            paragraph.pm_start = pm_start;
            paragraph.slot = slot;
            paragraph.source = source;
            paragraph.pilcrow = Some(mark.clone());
            paragraph.mark_attrs = attrs.cloned();
            let seed = std::mem::take(paragraph);
            self.seeds.insert(id, seed);
        }
        *paragraph = ParagraphSeed::default();
    }

    pub(super) fn finish(&mut self, blocks: &[LayoutBlock], map: &LoweringMap) {
        if self.blocked {
            self.seeds.clear();
            return;
        }
        let mut identities = BTreeSet::new();
        for (_, id) in &map.paragraphs {
            if !identities.insert(id) {
                self.excluded.insert(id.clone());
            }
        }
        self.seeds.retain(|id, seed| {
            if id.is_empty() || self.excluded.contains(id) {
                return false;
            }
            let Some(LayoutBlock::Paragraph(paragraph)) = blocks.get(seed.slot) else {
                return false;
            };
            paragraph.pm_start == Some(seed.pm_start as f64)
                && paragraph.id == BlockId::Str(id.clone())
                && paragraph
                    .attrs
                    .as_ref()
                    .is_some_and(|attrs| attrs.contextual_spacing != Some(true))
        });
    }

    pub(super) fn exclude(&mut self, id: String) {
        if !self.blocked {
            self.excluded.insert(id);
        }
    }

    pub(super) fn position_id(&mut self, id: String, prefix: &str, position: u64) {
        if !self.blocked {
            self.position_ids.insert(id, (prefix.to_owned(), position));
        }
    }

    pub(super) fn table_id(&mut self, id: String, story: &str, raw: u32, pm: u64) {
        if !self.blocked {
            self.table_ids.insert(id, (story.to_owned(), raw, pm));
        }
    }

    pub(crate) fn matches_source(&self, doc: &EditingDoc) -> bool {
        let source = doc.source_metadata();
        self.source
            .ptr_eq(&source.as_ref().map(Arc::downgrade).unwrap_or_default())
    }

    pub(crate) fn offset(&self, paragraph: &str, raw: u32) -> Option<u32> {
        raw.checked_sub(self.seeds.get(paragraph)?.raw_start)
    }

    pub(crate) fn patch<T: ReadTxn>(
        &mut self,
        blocks: &mut [LayoutBlock],
        map: &mut LoweringMap,
        revealable: &mut [LayoutBlock],
        txn: &T,
        env: &RenderEnv,
        edit: &TextEdit,
    ) -> Option<()> {
        if edit
            .attributes
            .as_ref()
            .is_some_and(|attrs| attrs.iter().any(|(key, value)| unsafe_value(key, value)))
        {
            return None;
        }
        let seed = self.seeds.get(&edit.paragraph)?;
        let pilcrow = seed.pilcrow.as_ref()?;
        let (raw, slot, source) = (seed.raw_start, seed.slot, seed.source);
        let pm_start = seed.pm_start;
        let old_units: u32 = seed
            .segments
            .iter()
            .map(|segment| utf16_len(&segment.text))
            .sum();
        let old_end = pm_start + u64::from(old_units) + 2;
        let segments = patch_segments(&seed.segments, edit)?;
        let delta = i64::from(utf16_len(&edit.text)) - i64::from(edit.removed);
        let mut runs = Vec::new();
        let mut units = 0;
        for segment in &segments {
            push_text_chunks(
                &mut runs,
                &segment.text,
                raw + units,
                Some(&segment.attrs),
                &[],
                env,
                units,
            );
            units += utf16_len(&segment.text);
        }
        let mut replacement = LoweringMap::default();
        let mut paragraph = flush_paragraph(
            runs,
            pilcrow,
            seed.mark_attrs.as_ref(),
            txn,
            "body",
            env,
            pm_start,
            units,
            &mut ListState::default(),
            (&mut replacement, source),
            Vec::new(),
        );
        let LayoutBlock::Paragraph(old) = blocks.get(slot)? else {
            return None;
        };
        if old
            .attrs
            .as_ref()
            .is_some_and(|attrs| attrs.num_pr.is_some() || attrs.list_marker.is_some())
        {
            paragraph.attrs = old.attrs.clone();
        }
        if !blocks[slot + 1..].iter().all(shiftable) || !revealable.iter().all(shiftable) {
            return None;
        }
        let mut ids = BTreeMap::new();
        let mut position_ids = BTreeMap::new();
        for (id, (prefix, position)) in &self.position_ids {
            let position = if *position >= old_end {
                position.checked_add_signed(delta)?
            } else {
                *position
            };
            let shifted = format!("{prefix}{position}");
            ids.insert(id.clone(), shifted.clone());
            position_ids.insert(shifted, (prefix.clone(), position));
        }
        let mut table_ids = BTreeMap::new();
        for (id, (story, raw, pm)) in &self.table_ids {
            let shifted = *pm >= old_end;
            let raw = if shifted && story == "body" {
                raw.checked_add_signed(delta.try_into().ok()?)?
            } else {
                *raw
            };
            let pm = if shifted {
                pm.checked_add_signed(delta)?
            } else {
                *pm
            };
            let new = format!("{story}:table:{raw}");
            ids.insert(id.clone(), new.clone());
            table_ids.insert(new, (story.clone(), raw, pm));
        }
        blocks[slot] = LayoutBlock::Paragraph(paragraph);
        for block in &mut blocks[slot + 1..] {
            shift_block(block, delta, old_end, &ids);
        }
        for block in revealable {
            if matches!(block, LayoutBlock::SectionBreak(_))
                || block_start(block).is_some_and(|start| start >= old_end as f64)
            {
                shift_block(block, delta, old_end, &ids);
            }
        }
        self.position_ids = position_ids;
        self.table_ids = table_ids;
        self.seeds.get_mut(&edit.paragraph).unwrap().segments = segments;
        shift_map(map, replacement, pm_start, old_end, source, delta);
        for seed in self.seeds.values_mut() {
            if seed.raw_start > raw {
                seed.raw_start = (i64::from(seed.raw_start) + delta) as u32;
                seed.pm_start = (seed.pm_start as i64 + delta) as u64;
            }
        }
        Some(())
    }
}

fn shift_map(
    map: &mut LoweringMap,
    replacement: LoweringMap,
    pm_start: u64,
    old_end: u64,
    source: u32,
    delta: i64,
) {
    for pm in map
        .paragraph_blocks
        .iter_mut()
        .map(|(pm, _)| pm)
        .chain(map.tables.iter_mut().map(|(pm, ..)| pm))
        .filter(|pm| **pm > pm_start)
    {
        *pm = (*pm as i64 + delta) as u64;
    }
    let span_start = map
        .spans
        .partition_point(|span| span.pm_start < pm_start + 1);
    let span_end = map.spans.partition_point(|span| span.pm_start < old_end);
    for span in &mut map.spans[span_end..] {
        span.pm_start = (span.pm_start as i64 + delta) as u64;
        span.pm_end = (span.pm_end as i64 + delta) as u64;
        if map.paragraphs[span.paragraph as usize].0 == map.paragraphs[source as usize].0 {
            span.raw_start = (i64::from(span.raw_start) + delta) as u32;
            span.raw_end = (i64::from(span.raw_end) + delta) as u32;
        }
    }
    map.spans.splice(span_start..span_end, replacement.spans);
}

fn patch_segments(segments: &[TextSegment], edit: &TextEdit) -> Option<Vec<TextSegment>> {
    let mut result = segments.to_vec();
    if result.is_empty() {
        if edit.offset != 0 || edit.removed != 0 {
            return None;
        }
        result.push(TextSegment {
            text: edit.text.clone(),
            attrs: attributes(edit.attributes.as_ref()),
        });
        return Some(result);
    }
    let mut before = 0;
    let slot = result.iter().position(|segment| {
        let end = before + utf16_len(&segment.text);
        if edit.offset < end || (edit.removed == 0 && edit.offset == end) {
            true
        } else {
            before = end;
            false
        }
    })?;
    let segment = &result[slot];
    let start = edit.offset.checked_sub(before)? as usize;
    let end = start.checked_add(edit.removed as usize)?;
    let units: Vec<_> = segment.text.encode_utf16().collect();
    let left = String::from_utf16(units.get(..start)?).ok()?;
    let removed = String::from_utf16(units.get(start..end)?).ok()?;
    let right = String::from_utf16(units.get(end..)?).ok()?;
    if !removed.is_empty() && removed.chars().count() != 1 {
        return None;
    }
    if edit.removed != 0 && left.is_empty() && right.is_empty() && edit.text.is_empty() {
        if result.len() > 1 {
            return None;
        }
        return Some(Vec::new());
    }
    let attrs = edit
        .attributes
        .as_ref()
        .map(|attrs| attributes(Some(attrs)))
        .unwrap_or_else(|| segment.attrs.clone());
    let replacement = [
        TextSegment {
            text: left,
            attrs: segment.attrs.clone(),
        },
        TextSegment {
            text: edit.text.clone(),
            attrs,
        },
        TextSegment {
            text: right,
            attrs: segment.attrs.clone(),
        },
    ];
    result.splice(slot..=slot, replacement);
    let mut merged: Vec<TextSegment> = Vec::new();
    for segment in result
        .into_iter()
        .filter(|segment| !segment.text.is_empty())
    {
        if let Some(last) = merged.last_mut()
            && same_attributes(&last.attrs, Some(&segment.attrs))
        {
            last.text.push_str(&segment.text);
        } else {
            merged.push(segment);
        }
    }
    Some(merged)
}

fn shift_pair(start: &mut Option<f64>, end: &mut Option<f64>, delta: i64) {
    for position in [start, end].into_iter().flatten() {
        *position += delta as f64;
    }
}

fn shiftable(block: &LayoutBlock) -> bool {
    match block {
        LayoutBlock::Table(table) => table
            .rows
            .iter()
            .flat_map(|row| &row.cells)
            .flat_map(|cell| &cell.blocks)
            .all(shiftable),
        LayoutBlock::Unsupported => false,
        _ => true,
    }
}

fn block_start(block: &LayoutBlock) -> Option<f64> {
    match block {
        LayoutBlock::Paragraph(block) => block.pm_start,
        LayoutBlock::Table(block) => block.pm_start,
        LayoutBlock::Image(block) => block.pm_start,
        LayoutBlock::Shape(block) => block.pm_start,
        LayoutBlock::Chart(block) => block.pm_start,
        LayoutBlock::TextBox(block) => block.pm_start,
        LayoutBlock::PageBreak(block) => block.pm_start,
        LayoutBlock::ColumnBreak(block) => block.pm_start,
        _ => None,
    }
}

fn shift_groups(
    groups: &mut Option<Vec<SdtGroup>>,
    delta: i64,
    cutoff: u64,
    ids: &BTreeMap<String, String>,
) {
    for group in groups.iter_mut().flatten() {
        if let Some(id) = ids.get(&group.id) {
            group.id = id.clone();
        }
        if let Some(pos) = &mut group.pos
            && *pos >= cutoff as i64
        {
            *pos += delta;
        }
    }
}

fn shift_id(id: &mut BlockId, ids: &BTreeMap<String, String>) {
    if let BlockId::Str(value) = id
        && let Some(shifted) = ids.get(value)
    {
        *value = shifted.clone();
    }
}

fn rename_shape(shape: &mut ShapeBlock, old: &str, new: &str) {
    if let BlockId::Str(id) = &mut shape.id
        && let Some(suffix) = id.strip_prefix(old)
    {
        *id = format!("{new}{suffix}");
    }
    for paragraph in shape.inner_text.iter_mut().flatten() {
        if let BlockId::Str(id) = &mut paragraph.id
            && let Some(suffix) = id.strip_prefix(old)
        {
            *id = format!("{new}{suffix}");
        }
    }
    for child in &mut shape.children {
        rename_shape(child, old, new);
    }
}

fn shift_shape(shape: &mut ShapeBlock, delta: i64, cutoff: u64, ids: &BTreeMap<String, String>) {
    if let BlockId::Str(old) = &shape.id
        && let Some(new) = ids.get(old)
    {
        let old = old.clone();
        rename_shape(shape, &old, new);
    }
    shift_pair(&mut shape.pm_start, &mut shape.pm_end, delta);
    shift_pair(&mut shape.doc_start, &mut shape.doc_end, delta);
    shift_groups(&mut shape.sdt_groups, delta, cutoff, ids);
    for paragraph in shape.inner_text.iter_mut().flatten() {
        shift_paragraph(paragraph, delta, cutoff, ids);
    }
    for child in &mut shape.children {
        shift_shape(child, delta, cutoff, ids);
    }
}

fn shift_widget(widget: &mut Option<Value>, delta: i64, cutoff: u64) {
    if let Some(widget) = widget.as_mut().and_then(Value::as_object_mut)
        && let Some(pos) = widget.get("pos").and_then(Value::as_i64)
        && pos >= cutoff as i64
    {
        let pos = pos + delta;
        widget.insert("pos".to_owned(), Value::from(pos));
        widget.insert("groupId".to_owned(), Value::String(format!("sdt@{pos}")));
    }
}

fn shift_paragraph(
    paragraph: &mut ParagraphBlock,
    delta: i64,
    cutoff: u64,
    ids: &BTreeMap<String, String>,
) {
    shift_pair(&mut paragraph.pm_start, &mut paragraph.pm_end, delta);
    shift_groups(&mut paragraph.sdt_groups, delta, cutoff, ids);
    if let Some(attrs) = &mut paragraph.attrs {
        for rule in &mut attrs.horizontal_rules {
            rule.pm_start += delta as f64;
            rule.pm_end += delta as f64;
        }
    }
    for run in &mut paragraph.runs {
        match run {
            Run::Text(run) => {
                shift_pair(&mut run.pm_start, &mut run.pm_end, delta);
                shift_widget(&mut run.inline_sdt_widget, delta, cutoff);
            }
            Run::Tab(run) => shift_pair(&mut run.pm_start, &mut run.pm_end, delta),
            Run::LineBreak(run) => shift_pair(&mut run.pm_start, &mut run.pm_end, delta),
            Run::Field(run) => shift_pair(&mut run.pm_start, &mut run.pm_end, delta),
            Run::Image(run) => {
                shift_pair(&mut run.pm_start, &mut run.pm_end, delta);
                if let Some(shape) = &mut run.inline_shape {
                    shift_shape(shape, delta, cutoff, ids);
                }
            }
        }
    }
}

fn shift_block(block: &mut LayoutBlock, delta: i64, cutoff: u64, ids: &BTreeMap<String, String>) {
    match block {
        LayoutBlock::Paragraph(paragraph) => shift_paragraph(paragraph, delta, cutoff, ids),
        LayoutBlock::Table(table) => {
            shift_pair(&mut table.pm_start, &mut table.pm_end, delta);
            shift_groups(&mut table.sdt_groups, delta, cutoff, ids);
            let new_id = match &table.id {
                BlockId::Str(id) => ids.get(id).cloned(),
                _ => None,
            };
            shift_id(&mut table.id, ids);
            for (index, row) in table.rows.iter_mut().enumerate() {
                if let Some(id) = &new_id {
                    row.id = BlockId::Str(format!("{id}:r{index}"));
                }
                for cell in &mut row.cells {
                    for block in &mut cell.blocks {
                        shift_block(block, delta, cutoff, ids);
                    }
                }
            }
        }
        LayoutBlock::Image(block) => {
            shift_pair(&mut block.pm_start, &mut block.pm_end, delta);
            shift_groups(&mut block.sdt_groups, delta, cutoff, ids);
            shift_id(&mut block.id, ids);
        }
        LayoutBlock::Shape(block) => shift_shape(block, delta, cutoff, ids),
        LayoutBlock::Chart(block) => {
            shift_pair(&mut block.pm_start, &mut block.pm_end, delta);
            shift_pair(&mut block.doc_start, &mut block.doc_end, delta);
            shift_groups(&mut block.sdt_groups, delta, cutoff, ids);
            shift_id(&mut block.id, ids);
        }
        LayoutBlock::TextBox(block) => {
            shift_pair(&mut block.pm_start, &mut block.pm_end, delta);
            shift_groups(&mut block.sdt_groups, delta, cutoff, ids);
            for paragraph in &mut block.content {
                shift_paragraph(paragraph, delta, cutoff, ids);
            }
        }
        LayoutBlock::PageBreak(block) => {
            shift_pair(&mut block.pm_start, &mut block.pm_end, delta);
            shift_groups(&mut block.sdt_groups, delta, cutoff, ids);
        }
        LayoutBlock::ColumnBreak(block) => {
            shift_pair(&mut block.pm_start, &mut block.pm_end, delta);
            shift_groups(&mut block.sdt_groups, delta, cutoff, ids);
        }
        LayoutBlock::SectionBreak(block) => shift_groups(&mut block.sdt_groups, delta, cutoff, ids),
        LayoutBlock::Unsupported => unreachable!("certified suffix"),
    }
}
