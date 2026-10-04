use super::*;

#[derive(Debug, Default)]
pub(crate) struct LocalLowering {
    pub(crate) blocked: bool,
    pub(super) source: std::sync::Weak<crate::seed::SourceMetadata>,
    pub(super) seeds: BTreeMap<String, ParagraphSeed>,
    identities: BTreeSet<String>,
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
    start_safe: bool,
    rejected: bool,
    values: BTreeMap<String, Any>,
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
    [INS, DEL, "pPrIns", "pPrDel", "pPrChange", "rPrChange"].contains(&key)
        || ["trPrChange", "tcPrChange"].contains(&key)
        || ["trackedIns", "trackedDel", "trIns", "trDel"].contains(&key)
        || ["hyperlink", "bookmarks", "comment", "cellMarker"].contains(&key)
        || [
            "hidden",
            "vanish",
            "floating",
            "footnoteRefId",
            "endnoteRefId",
        ]
        .contains(&key)
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
        boundary: Option<bool>,
    ) {
        if self.blocked || story != "body" {
            return;
        }
        if let Some(safe) = boundary {
            paragraph.start_safe = safe;
        }
        let attrs = diff.attributes.as_deref();
        paragraph.rejected |= attrs
            .into_iter()
            .flatten()
            .any(|(key, value)| unsafe_value(key, value));
        match &diff.insert {
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
            _ => paragraph.rejected = true,
        }
    }

    #[allow(clippy::too_many_arguments)]
    pub(super) fn observe_pilcrow<T: ReadTxn>(
        &mut self,
        paragraph: &mut ParagraphSeed,
        mark: &MapRef,
        values: &BTreeMap<String, Any>,
        attrs: Option<&Attrs>,
        txn: &T,
        story: &str,
        (start, pm_start, slot, source): (u32, u64, usize, u32),
        end_safe: bool,
    ) {
        if self.blocked {
            return;
        }
        let id = value_string(values.get("paraId")).unwrap_or_default();
        self.blocked |= !self.identities.insert(id.clone());
        if self.blocked || story != "body" {
            return;
        }
        paragraph.rejected |= values.iter().any(|(key, value)| unsafe_value(key, value))
            || values.contains_key("sectPr")
            || values.contains_key("sectionBreakType")
            || self
                .source
                .upgrade()
                .is_some_and(|source| source.run_revision(story, &id));
        if paragraph.start_safe
            && end_safe
            && !paragraph.rejected
            && !id.is_empty()
            && comments_clear(txn, start, start + paragraph.units())
        {
            paragraph.raw_start = start;
            paragraph.pm_start = pm_start;
            paragraph.slot = slot;
            paragraph.source = source;
            paragraph.pilcrow = Some(mark.clone());
            paragraph.mark_attrs = attrs.cloned();
            paragraph.values = values.clone();
            self.seeds.insert(id, std::mem::take(paragraph));
        }
        *paragraph = ParagraphSeed::default();
    }

    pub(crate) fn finish(
        &mut self,
        blocks: &mut [LayoutBlock],
        map: &LoweringMap,
        revealable: &mut [LayoutBlock],
    ) {
        self.identities.clear();
        if self.blocked {
            self.seeds.clear();
            return;
        }
        let mut identities = BTreeSet::new();
        self.blocked |= map.paragraphs.iter().any(|(_, id)| !identities.insert(id));
        let shift = super::positions::PositionShift::new(0, 0, false);
        self.blocked |= blocks
            .iter_mut()
            .chain(revealable)
            .any(|block| shift.block(block).is_none());
        if self.blocked {
            self.seeds.clear();
            return;
        }
        let mut ownership = vec![0_u32; map.paragraphs.len()];
        for (_, source) in &map.paragraph_blocks {
            ownership[*source as usize] += 1;
        }
        for (slot, block) in blocks.iter().enumerate() {
            if let LayoutBlock::Paragraph(paragraph) = block
                && let BlockId::Str(id) = &paragraph.id
                && let Some(seed) = self.seeds.get_mut(id)
            {
                seed.slot = slot;
            }
        }
        self.seeds.retain(|id, seed| {
            let Some(LayoutBlock::Paragraph(paragraph)) = blocks.get(seed.slot) else {
                return false;
            };
            let start = map
                .spans
                .partition_point(|span| span.pm_start < seed.pm_start + 1);
            let end = map.spans.partition_point(|span| {
                span.pm_start < seed.pm_start + u64::from(seed.units()) + 2
            });
            ordinary(paragraph, seed.pm_start, seed.units())
                && matches!(&paragraph.id, BlockId::Str(block_id) if block_id == id)
                && map.spans[start..end]
                    .iter()
                    .all(|span| !span.atom && span.paragraph == seed.source)
                && map
                    .paragraphs
                    .get(seed.source as usize)
                    .is_some_and(|(story, paragraph)| *story == 0 && paragraph == id)
                && ownership[seed.source as usize] == 1
        });
    }

    pub(crate) fn accepts_range<T: ReadTxn>(
        &self,
        txn: &T,
        paragraph: &str,
        start: u32,
        end: u32,
    ) -> bool {
        self.seeds.get(paragraph).is_some_and(|seed| {
            let paragraph_end = seed.raw_start + seed.units();
            start >= seed.raw_start
                && end >= start
                && end <= paragraph_end
                && comments_clear(txn, seed.raw_start, paragraph_end)
        })
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
        let seed = self.seeds.get_mut(&edit.paragraph)?;
        let pilcrow = seed.pilcrow.as_ref()?;
        let (raw, slot, source) = (seed.raw_start, seed.slot, seed.source);
        let pm_start = seed.pm_start;
        let old_units = seed.units();
        if pilcrow_values(pilcrow, txn) != seed.values
            || !edit.text.is_empty() && edit.attributes.is_none()
        {
            return None;
        }
        let old_end = pm_start + u64::from(old_units) + 2;
        let segments = patch_segments(&seed.segments, edit)?;
        let new_units: u32 = segments
            .iter()
            .map(|segment| utf16_len(&segment.text))
            .sum();
        let expected_units = old_units
            .checked_sub(edit.removed)?
            .checked_add(utf16_len(&edit.text))?;
        if new_units != expected_units
            || (old_units == 0) != (new_units == 0)
            || !effective_attributes_match(&segments, edit)
            || segments.iter().any(|segment| {
                segment
                    .attrs
                    .iter()
                    .any(|(key, value)| unsafe_value(key, value))
            })
            || !comments_clear(txn, raw, raw + new_units)
        {
            return None;
        }
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
        if !ordinary(&paragraph, pm_start, new_units)
            || replacement
                .spans
                .iter()
                .any(|span| span.atom || span.paragraph != source)
        {
            return None;
        }
        let shift = super::positions::PositionShift::new(old_end, delta, false);
        for block in blocks[slot + 1..].iter_mut().chain(revealable.iter_mut()) {
            shift.block(block)?;
        }
        shift.map(map)?;
        for later in self.seeds.values().filter(|seed| seed.raw_start > raw) {
            u32::try_from(i64::from(later.raw_start) + delta).ok()?;
            later.pm_start.checked_add_signed(delta)?;
        }
        self.seeds.get_mut(&edit.paragraph)?.segments = segments;
        blocks[slot] = LayoutBlock::Paragraph(paragraph);
        let shift = super::positions::PositionShift::new(old_end, delta, true);
        for block in blocks[slot + 1..].iter_mut().chain(revealable) {
            shift.block(block).expect("validated position shift");
        }
        let span_start = map
            .spans
            .partition_point(|span| span.pm_start < pm_start + 1);
        let span_end = map.spans.partition_point(|span| span.pm_start < old_end);
        shift.map(map).expect("validated map shift");
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

fn effective_attributes_match(segments: &[TextSegment], edit: &TextEdit) -> bool {
    if edit.text.is_empty() {
        return true;
    }
    let Some(attrs) = edit.attributes.as_ref() else {
        return false;
    };
    let Some(end) = edit.offset.checked_add(utf16_len(&edit.text)) else {
        return false;
    };
    let mut offset = 0;
    for segment in segments {
        let next = offset + utf16_len(&segment.text);
        if offset < end && edit.offset < next && !same_attributes(&segment.attrs, Some(attrs)) {
            return false;
        }
        offset = next;
    }
    end <= offset
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

impl ParagraphSeed {
    fn units(&self) -> u32 {
        self.segments
            .iter()
            .map(|segment| utf16_len(&segment.text))
            .sum()
    }
}

fn ordinary(paragraph: &ParagraphBlock, start: u64, units: u32) -> bool {
    paragraph.pm_start == Some(start as f64)
        && paragraph.pm_end == Some((start + u64::from(units) + 2) as f64)
        && paragraph.sdt_groups.is_none()
        && paragraph
            .attrs
            .as_ref()
            .is_some_and(|attrs| attrs.horizontal_rules.is_empty())
        && paragraph.runs.iter().all(|run| match run {
            Run::Text(run) => run.inline_sdt_widget.is_none() && ordinary_formatting(&run.fmt),
            Run::Tab(run) => ordinary_formatting(&run.fmt),
            Run::LineBreak(_) => true,
            Run::Image(_) | Run::Field(_) | Run::Unsupported => false,
        })
}

fn ordinary_formatting(formatting: &RunFormatting) -> bool {
    formatting.hyperlink.is_none()
        && formatting.comment_ids.is_none()
        && formatting.footnote_ref_id.is_none()
        && formatting.endnote_ref_id.is_none()
        && formatting.hidden != Some(true)
        && formatting.is_insertion != Some(true)
        && formatting.is_deletion != Some(true)
        && formatting.change_revision_id.is_none()
}

fn comments_clear<T: ReadTxn>(txn: &T, start: u32, end: u32) -> bool {
    let Some(comments) = txn.get_map(COMMENTS) else {
        return false;
    };
    for (_, value) in comments.iter(txn) {
        let Out::YMap(comment) = value else {
            continue;
        };
        let Some(Out::Any(Any::Array(anchors))) = comment.get(txn, "anchors") else {
            return false;
        };
        for encoded in anchors.iter() {
            let Ok(anchor) = decode_anchor(encoded) else {
                return false;
            };
            if anchor.story != "body" {
                continue;
            }
            let (Some(left), Some(right)) =
                (anchor.start.get_offset(txn), anchor.end.get_offset(txn))
            else {
                return false;
            };
            if left.index.min(right.index) <= end && left.index.max(right.index) >= start {
                return false;
            }
        }
    }
    true
}
