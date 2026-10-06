use std::rc::Rc;

use super::*;

#[derive(Debug, Default)]
pub(crate) struct LocalLowering {
    pub(crate) blocked: bool,
    pub(crate) preview_blocked: bool,
    pub(super) enabled: bool,
    pub(super) source: std::sync::Weak<crate::seed::SourceMetadata>,
    pub(super) seeds: BTreeMap<String, ParagraphSeed>,
    dependent: BTreeSet<String>,
    pub(crate) edit: Option<TextEdit>,
}

#[derive(Debug, Default)]
pub(super) struct ParagraphSeed {
    tainted: bool,
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

impl ParagraphSeed {
    pub(super) fn start(&mut self, safe: bool) {
        *self = Self {
            tainted: !safe,
            ..Self::default()
        };
    }
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

pub(super) fn preview_touches_state<T: ReadTxn>(
    diff: &yrs::types::text::Diff<YChange>,
    txn: &T,
) -> bool {
    match &diff.insert {
        Out::Any(Any::String(_)) => false,
        Out::YMap(mark) if is_pilcrow(mark, txn) => false,
        Out::YMap(mark) => {
            let values = pilcrow_values(mark, txn);
            let seed_only = [
                "break",
                "pageBreak",
                "columnBreak",
                "image",
                "shape",
                "noteRef",
                "math",
                "horizontalRule",
            ]
            .contains(
                &value_string(values.get("_kind"))
                    .unwrap_or_default()
                    .as_str(),
            );
            !seed_only
                || ["fieldCodeMarks", "fieldResultBlocks"]
                    .iter()
                    .any(|key| !any_strings(values.get(*key)).is_empty())
                || value_string(values.get("fieldCodeTarget")).is_some()
        }
        _ => true,
    }
}

impl LocalLowering {
    #[cfg(test)]
    pub(crate) fn has_dependencies(&self) -> bool {
        !self.dependent.is_empty()
    }

    #[cfg(test)]
    pub(crate) fn snapshot(&self, doc: &EditingDoc) -> impl PartialEq + std::fmt::Debug + use<> {
        use super::preview::AnySnapshot;
        use yrs::types::ToJson;

        let txn = doc.yrs_doc().transact();
        let attrs = |attrs: &Attrs| {
            attrs
                .iter()
                .map(|(key, value)| (key.to_string(), AnySnapshot::from(value)))
                .collect::<BTreeMap<_, _>>()
        };
        let seeds = self
            .seeds
            .iter()
            .map(|(id, seed)| {
                (
                    id.clone(),
                    seed.tainted,
                    seed.raw_start,
                    seed.slot,
                    seed.source,
                    seed.pm_start,
                    seed.segments
                        .iter()
                        .map(|segment| (segment.text.clone(), attrs(&segment.attrs)))
                        .collect::<Vec<_>>(),
                    seed.pilcrow
                        .as_ref()
                        .map(|mark| (mark.clone(), AnySnapshot::from(&mark.to_json(&txn)))),
                    seed.mark_attrs.as_ref().map(attrs),
                )
            })
            .collect::<Vec<_>>();
        (
            self.blocked,
            self.preview_blocked,
            self.enabled,
            self.dependent.clone(),
            seeds,
        )
    }

    pub(crate) fn new(enabled: bool) -> Self {
        Self {
            blocked: !enabled,
            preview_blocked: !enabled,
            enabled,
            ..Default::default()
        }
    }

    pub(super) fn block(&mut self, blocked: bool) {
        self.blocked |= blocked;
        self.preview_blocked |= blocked;
    }

    pub(super) fn replace_seeds(&mut self, pm: &std::ops::Range<u64>, replacement: Self) {
        self.seeds.retain(|_, seed| !pm.contains(&seed.pm_start));
        self.seeds.extend(replacement.seeds);
    }

    pub(super) fn replay_preserves_state(&self, replacement: &Self) -> bool {
        self.blocked == replacement.blocked
            && (self.preview_blocked || !replacement.preview_blocked)
            && replacement.dependent.is_empty()
    }

    pub(crate) fn refresh_seeds(&mut self, blocks: &[Rc<LayoutBlock>], map: &LoweringMap) {
        self.edit = None;
        if self.seeds.is_empty() {
            return;
        }
        let mut identities = BTreeSet::new();
        let duplicates: BTreeSet<_> = map
            .paragraphs
            .iter()
            .filter_map(|(_, id)| (!identities.insert(id.clone())).then_some(id.clone()))
            .collect();
        let sources: BTreeMap<_, _> = map
            .paragraphs
            .iter()
            .enumerate()
            .filter(|(_, (story, _))| *story == 0)
            .map(|(source, (_, id))| (id.as_str(), source as u32))
            .collect();
        let positions: BTreeMap<_, _> = map
            .paragraph_blocks
            .iter()
            .map(|&(pm, source)| (source, pm))
            .collect();
        let slots: BTreeMap<_, _> = blocks
            .iter()
            .enumerate()
            .filter_map(|(slot, block)| match block.as_ref() {
                LayoutBlock::Paragraph(paragraph) => paragraph.pm_start.map(|pm| (pm as u64, slot)),
                _ => None,
            })
            .collect();
        self.seeds.retain(|id, seed| {
            if self.dependent.contains(id) || duplicates.contains(id) {
                return false;
            }
            let Some((&source, &pm)) = sources
                .get(id.as_str())
                .and_then(|source| Some((source, positions.get(source)?)))
            else {
                return false;
            };
            let Some(&slot) = slots.get(&pm) else {
                return false;
            };
            let LayoutBlock::Paragraph(paragraph) = blocks[slot].as_ref() else {
                return false;
            };
            if paragraph.id != BlockId::Str(id.clone())
                || paragraph.attrs.is_none()
                || !shiftable(&blocks[slot])
                || page_break_changes_marker(paragraph, blocks.get(slot + 1).map(Rc::as_ref))
            {
                return false;
            }
            seed.source = source;
            seed.pm_start = pm;
            seed.slot = slot;
            true
        });
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
        let unsafe_attrs = attrs
            .into_iter()
            .flatten()
            .any(|(key, value)| unsafe_value(key, value));
        self.preview_blocked |= unsafe_attrs;
        paragraph.tainted |= unsafe_attrs;
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
                for key in ["fieldCodeMarks", "fieldResultBlocks"] {
                    self.dependent.extend(any_strings(values.get(key)));
                }
                if let Some(target) = value_string(values.get("fieldCodeTarget")) {
                    self.dependent.insert(target);
                }
                self.preview_blocked |= values.iter().any(|(key, value)| unsafe_value(key, value))
                    || value_string(values.get("_kind")).as_deref() != Some("table");
                if value_string(values.get("_kind")).as_deref() == Some("table") {
                    *paragraph = ParagraphSeed::default();
                } else {
                    paragraph.tainted = true;
                }
            }
            _ => {
                self.preview_blocked = true;
                paragraph.tainted = true;
            }
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
        last: bool,
        safe: bool,
    ) {
        if self.blocked {
            return;
        }
        let unsafe_values = values.iter().any(|(key, value)| unsafe_value(key, value));
        let sectioned = values.contains_key("sectPr") || values.contains_key("sectionBreakType");
        self.preview_blocked |= unsafe_values || sectioned && (story != "body" || !last);
        self.block(sectioned && story != "body");
        paragraph.tainted |= unsafe_values || !safe;
        if story == "body" && !sectioned && !paragraph.tainted && !self.blocked {
            paragraph.raw_start = start;
            paragraph.pm_start = pm_start;
            paragraph.slot = slot;
            paragraph.source = source;
            paragraph.pilcrow = Some(mark.clone());
            paragraph.mark_attrs = attrs.cloned();
            let seed = std::mem::take(paragraph);
            self.seeds
                .insert(value_string(values.get("paraId")).unwrap_or_default(), seed);
        }
        *paragraph = ParagraphSeed::default();
    }

    pub(super) fn finish(&mut self, blocks: &[LayoutBlock], map: &LoweringMap) {
        if self.blocked {
            self.seeds.clear();
            return;
        }
        let mut identities = BTreeSet::new();
        let duplicates: BTreeSet<_> = map
            .paragraphs
            .iter()
            .filter_map(|(_, id)| (!identities.insert(id.clone())).then_some(id.clone()))
            .collect();
        self.preview_blocked |= !duplicates.is_empty() || !blocks.iter().all(shiftable);
        self.seeds.retain(|id, seed| {
            if duplicates.contains(id) || self.dependent.contains(id) {
                return false;
            }
            let Some(LayoutBlock::Paragraph(paragraph)) = blocks.get(seed.slot) else {
                return false;
            };
            paragraph.pm_start == Some(seed.pm_start as f64)
                && paragraph.attrs.is_some()
                && shiftable(&blocks[seed.slot])
                && !page_break_changes_marker(paragraph, blocks.get(seed.slot + 1))
        });
    }

    pub(crate) fn matches_source(&self, doc: &EditingDoc) -> bool {
        let source = doc.source_metadata();
        self.source
            .ptr_eq(&source.as_ref().map(Arc::downgrade).unwrap_or_default())
    }

    pub(crate) fn offset(&self, paragraph: &str, raw: u32) -> Option<u32> {
        if self.blocked {
            return None;
        }
        raw.checked_sub(self.seeds.get(paragraph)?.raw_start)
    }

    pub(crate) fn edit_slot(&self, edit: &TextEdit) -> Option<usize> {
        if self.blocked {
            return None;
        }
        Some(self.seeds.get(&edit.paragraph)?.slot)
    }

    pub(crate) fn patch<T: ReadTxn>(
        &mut self,
        blocks: &mut [Rc<LayoutBlock>],
        map: &mut LoweringMap,
        revealable: &mut [LayoutBlock],
        txn: &T,
        env: &RenderEnv,
        edit: &TextEdit,
    ) -> Option<()> {
        if self.blocked
            || edit
                .attributes
                .as_ref()
                .is_some_and(|attrs| attrs.iter().any(|(key, value)| unsafe_value(key, value)))
        {
            return None;
        }
        let seed = self.seeds.get_mut(&edit.paragraph)?;
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
        let mut replacement = super::preview::LoweringOutput::default();
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
        let LayoutBlock::Paragraph(old) = blocks.get(slot)?.as_ref() else {
            return None;
        };
        if old
            .attrs
            .as_ref()
            .is_some_and(|attrs| attrs.num_pr.is_some() || attrs.list_marker.is_some())
        {
            paragraph.attrs = old.attrs.clone();
        }
        seed.segments = segments;
        blocks[slot] = Rc::new(LayoutBlock::Paragraph(paragraph));
        for block in &mut blocks[slot + 1..] {
            shift_block(Rc::make_mut(block), delta);
        }
        for block in revealable {
            if block
                .pm_start()
                .is_some_and(|start| start >= old_end as f64)
            {
                shift_block(block, delta);
            }
        }
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
            if map.paragraphs[span.paragraph as usize].0 == 0 {
                span.raw_start = (i64::from(span.raw_start) + delta) as u32;
                span.raw_end = (i64::from(span.raw_end) + delta) as u32;
            }
        }
        map.spans
            .splice(span_start..span_end, replacement.map.spans);
        for seed in self.seeds.values_mut() {
            if seed.raw_start > raw {
                seed.raw_start = (i64::from(seed.raw_start) + delta) as u32;
                seed.pm_start = (seed.pm_start as i64 + delta) as u64;
            }
        }
        Some(())
    }
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

fn page_break_changes_marker(paragraph: &ParagraphBlock, next: Option<&LayoutBlock>) -> bool {
    paragraph
        .attrs
        .as_ref()
        .is_some_and(|attrs| attrs.list_marker.is_some())
        && matches!(next, Some(LayoutBlock::PageBreak(page_break))
            if page_break.pm_start == paragraph.pm_end)
}

fn shiftable(block: &LayoutBlock) -> bool {
    match block {
        LayoutBlock::Paragraph(paragraph) => {
            paragraph
                .attrs
                .as_ref()
                .is_none_or(|attrs| attrs.horizontal_rules.is_empty())
                && paragraph
                    .runs
                    .iter()
                    .all(|run| matches!(run, Run::Text(_) | Run::Tab(_) | Run::LineBreak(_)))
        }
        LayoutBlock::Table(table) => table
            .rows
            .iter()
            .flat_map(|row| &row.cells)
            .flat_map(|cell| &cell.blocks)
            .all(shiftable),
        LayoutBlock::SectionBreak(_) => true,
        LayoutBlock::Image(_)
        | LayoutBlock::Shape(_)
        | LayoutBlock::Chart(_)
        | LayoutBlock::TextBox(_)
        | LayoutBlock::PageBreak(_)
        | LayoutBlock::ColumnBreak(_)
        | LayoutBlock::Unsupported => false,
    }
}

fn shift_block(block: &mut LayoutBlock, delta: i64) {
    match block {
        LayoutBlock::Paragraph(paragraph) => shift_paragraph(paragraph, delta),
        LayoutBlock::Table(table) => {
            shift_pair(&mut table.pm_start, &mut table.pm_end, delta);
            shift_groups(&mut table.sdt_groups, delta);
            for row in &mut table.rows {
                for cell in &mut row.cells {
                    for block in &mut cell.blocks {
                        shift_block(block, delta);
                    }
                }
            }
        }
        LayoutBlock::Image(image) => {
            shift_pair(&mut image.pm_start, &mut image.pm_end, delta);
            shift_groups(&mut image.sdt_groups, delta);
        }
        LayoutBlock::Shape(shape) => shift_shape(shape, delta),
        LayoutBlock::Chart(chart) => {
            shift_pair(&mut chart.pm_start, &mut chart.pm_end, delta);
            shift_pair(&mut chart.doc_start, &mut chart.doc_end, delta);
            shift_drawing_id(&mut chart.id, delta);
            shift_groups(&mut chart.sdt_groups, delta);
        }
        LayoutBlock::TextBox(text_box) => {
            shift_pair(&mut text_box.pm_start, &mut text_box.pm_end, delta);
            shift_groups(&mut text_box.sdt_groups, delta);
            for paragraph in &mut text_box.content {
                shift_paragraph(paragraph, delta);
            }
        }
        LayoutBlock::SectionBreak(section) => shift_groups(&mut section.sdt_groups, delta),
        LayoutBlock::PageBreak(page_break) => {
            shift_pair(&mut page_break.pm_start, &mut page_break.pm_end, delta);
            shift_groups(&mut page_break.sdt_groups, delta);
        }
        LayoutBlock::ColumnBreak(column_break) => {
            shift_pair(&mut column_break.pm_start, &mut column_break.pm_end, delta);
            shift_groups(&mut column_break.sdt_groups, delta);
        }
        LayoutBlock::Unsupported => {}
    }
}

fn shift_paragraph(paragraph: &mut ParagraphBlock, delta: i64) {
    if paragraph.pm_start.is_none() {
        shift_drawing_id(&mut paragraph.id, delta);
    }
    shift_pair(&mut paragraph.pm_start, &mut paragraph.pm_end, delta);
    shift_groups(&mut paragraph.sdt_groups, delta);
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
                if let Some(widget) = run
                    .inline_sdt_widget
                    .as_mut()
                    .and_then(Value::as_object_mut)
                {
                    if let Some(pos) = widget.get_mut("pos")
                        && let Some(value) = pos.as_i64()
                    {
                        *pos = Value::from(value + delta);
                    }
                    if let Some(Value::String(id)) = widget.get_mut("groupId") {
                        shift_group_id(id, delta);
                    }
                }
            }
            Run::Tab(run) => shift_pair(&mut run.pm_start, &mut run.pm_end, delta),
            Run::LineBreak(run) => shift_pair(&mut run.pm_start, &mut run.pm_end, delta),
            Run::Image(run) => {
                shift_pair(&mut run.pm_start, &mut run.pm_end, delta);
                if let Some(shape) = &mut run.inline_shape {
                    shift_shape(shape, delta);
                }
            }
            Run::Field(run) => shift_pair(&mut run.pm_start, &mut run.pm_end, delta),
            Run::Unsupported => {}
        }
    }
}

fn shift_shape(shape: &mut ShapeBlock, delta: i64) {
    shift_pair(&mut shape.pm_start, &mut shape.pm_end, delta);
    shift_pair(&mut shape.doc_start, &mut shape.doc_end, delta);
    shift_drawing_id(&mut shape.id, delta);
    shift_groups(&mut shape.sdt_groups, delta);
    for paragraph in shape.inner_text.iter_mut().flatten() {
        shift_paragraph(paragraph, delta);
    }
    for child in &mut shape.children {
        shift_shape(child, delta);
    }
}

fn shift_drawing_id(id: &mut BlockId, delta: i64) {
    let BlockId::Str(value) = id else {
        return;
    };
    let Some((kind @ ("shape" | "chart"), rest)) = value.split_once(':') else {
        return;
    };
    let end = rest.find(':').unwrap_or(rest.len());
    if let Ok(position) = rest[..end].parse::<i64>() {
        *value = format!("{kind}:{}{}", position + delta, &rest[end..]);
    }
}

fn shift_group_id(id: &mut String, delta: i64) {
    if let Some(position) = id
        .strip_prefix("sdt@")
        .and_then(|value| value.parse::<i64>().ok())
    {
        *id = format!("sdt@{}", position + delta);
    }
}

fn shift_groups(groups: &mut Option<Vec<SdtGroup>>, delta: i64) {
    for group in groups.iter_mut().flatten() {
        shift_group_id(&mut group.id, delta);
        if let Some(pos) = &mut group.pos {
            *pos += delta;
        }
    }
}
