use super::*;

#[derive(Debug, Default)]
pub(crate) struct LocalLowering {
    pub(super) blocked: bool,
    pub(super) source: std::sync::Weak<crate::seed::SourceMetadata>,
    pub(super) seeds: BTreeMap<String, ParagraphSeed>,
    pub(crate) edit: Option<TextEdit>,
}

#[derive(Debug, Default)]
pub(super) struct ParagraphSeed {
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
        self.blocked |= attrs
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
                self.blocked |= values.iter().any(|(key, value)| unsafe_value(key, value))
                    || value_string(values.get("_kind")).as_deref() != Some("table");
                *paragraph = ParagraphSeed::default();
            }
            _ => self.blocked = true,
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
    ) {
        if self.blocked {
            return;
        }
        self.blocked |= values.iter().any(|(key, value)| unsafe_value(key, value));
        let sectioned = values.contains_key("sectPr") || values.contains_key("sectionBreakType");
        self.blocked |= sectioned && (story != "body" || !last);
        if story == "body" && !sectioned && !self.blocked {
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
        self.blocked |= map.paragraphs.iter().any(|(_, id)| !identities.insert(id));
        self.blocked |= !blocks.iter().all(shiftable);
        if self.blocked {
            self.seeds.clear();
            return;
        }
        self.seeds.retain(|_, seed| {
            let Some(LayoutBlock::Paragraph(paragraph)) = blocks.get(seed.slot) else {
                return false;
            };
            paragraph.pm_start == Some(seed.pm_start as f64) && paragraph.attrs.is_some()
        });
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
        txn: &T,
        env: &RenderEnv,
        edit: &TextEdit,
    ) -> Option<()> {
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
        seed.segments = segments;
        blocks[slot] = LayoutBlock::Paragraph(paragraph);
        for block in &mut blocks[slot + 1..] {
            shift_block(block, delta);
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
        map.spans.splice(span_start..span_end, replacement.spans);
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
        _ => false,
    }
}

fn shift_block(block: &mut LayoutBlock, delta: i64) {
    match block {
        LayoutBlock::Paragraph(paragraph) => {
            shift_pair(&mut paragraph.pm_start, &mut paragraph.pm_end, delta);
            for run in &mut paragraph.runs {
                match run {
                    Run::Text(run) => shift_pair(&mut run.pm_start, &mut run.pm_end, delta),
                    Run::Tab(run) => shift_pair(&mut run.pm_start, &mut run.pm_end, delta),
                    Run::LineBreak(run) => shift_pair(&mut run.pm_start, &mut run.pm_end, delta),
                    _ => unreachable!("certified plain text"),
                }
            }
        }
        LayoutBlock::Table(table) => {
            shift_pair(&mut table.pm_start, &mut table.pm_end, delta);
            for row in &mut table.rows {
                for cell in &mut row.cells {
                    for block in &mut cell.blocks {
                        shift_block(block, delta);
                    }
                }
            }
        }
        LayoutBlock::SectionBreak(_) => {}
        _ => unreachable!("certified body blocks"),
    }
}
