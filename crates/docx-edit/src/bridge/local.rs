use std::borrow::Cow;
use std::collections::HashSet;
use std::rc::Rc;

use super::*;

#[derive(Debug, Default)]
pub(crate) struct LocalLowering {
    pub(crate) blocked: bool,
    pub(crate) preview_blocked: bool,
    pub(super) enabled: bool,
    pub(crate) deferred: bool,
    pub(crate) legacy: bool,
    pub(super) source: std::sync::Weak<crate::seed::SourceMetadata>,
    pub(super) seeds: BTreeMap<String, ParagraphSeed>,
    pub(super) chunks: Option<Rc<SeedChunks>>,
    validation_epoch: u64,
    dependent: BTreeSet<String>,
    pub(crate) edit: Option<TextEdit>,
    #[cfg(test)]
    pub(crate) materialized_text_units: u32,
    #[cfg(test)]
    discarded: BTreeSet<String>,
}

#[derive(Debug, Default)]
pub(super) struct ParagraphSeed {
    tainted: bool,
    raw_start: u32,
    slot: usize,
    source: u32,
    segments: Vec<TextSegment>,
    chunks: Option<(Rc<SeedChunks>, std::ops::Range<usize>)>,
    pilcrow: Option<MapRef>,
    mark_attrs: Option<Attrs>,
    pm_start: u64,
    validated: Option<(u64, usize)>,
}

#[derive(Debug)]
pub(super) struct SeedChunks {
    pub(super) diffs: Vec<yrs::types::text::Diff<YChange>>,
    pub(super) units: Vec<std::cell::Cell<u32>>,
}

impl SeedChunks {
    pub(super) fn new(diffs: Vec<yrs::types::text::Diff<YChange>>) -> Self {
        Self {
            units: vec![std::cell::Cell::new(0); diffs.len()],
            diffs,
        }
    }
}

#[cfg(test)]
#[test]
fn paragraph_seeds_share_chunk_widths() {
    let doc = EditingDoc::new(9643);
    doc.create_story("body", "α😀β", "Normal", "left").unwrap();
    doc.split_paragraph(
        &crate::EditCtx::local("", ""),
        crate::Position::new("body", 3),
        None,
    )
    .unwrap();
    let env = RenderEnv::default();
    let mut local = LocalLowering::new(true);
    let lowered =
        yrs_doc_to_mapped_layout_blocks_with_revealable(&doc, "body", &env, &mut local).unwrap();
    let mut disabled = LocalLowering::new(false);
    let cold =
        yrs_doc_to_mapped_layout_blocks_with_revealable(&doc, "body", &env, &mut disabled).unwrap();
    assert_eq!(lowered, cold);
    assert_eq!(local.seeds.len(), 2);
    assert_eq!(local.materialized_text_units, 0);
    let mut seeds = local.seeds.values();
    let (chunks, first) = seeds.next().unwrap().chunks.as_ref().unwrap();
    let (shared, second) = seeds.next().unwrap().chunks.as_ref().unwrap();
    assert!(Rc::ptr_eq(chunks, shared));
    assert_ne!(first, second);
    assert_eq!(chunks.diffs.len(), chunks.units.len());
    for seed in local.seeds.values() {
        assert!(seed.segments.is_empty());
        for part in seed.parts() {
            assert_eq!(part.units, utf16_len(part.text));
        }
    }
}

#[derive(Clone, Debug)]
struct TextSegment {
    text: String,
    units: u32,
    attrs: Attrs,
}

#[derive(Clone, Copy)]
struct SegmentPart<'a> {
    text: &'a str,
    units: u32,
    attrs: Option<&'a Attrs>,
}

impl ParagraphSeed {
    fn segments(&self) -> Cow<'_, [TextSegment]> {
        let Some((chunks, range)) = &self.chunks else {
            return Cow::Borrowed(&self.segments);
        };
        #[cfg(test)]
        crate::engine::TYPING_EXTRA_WORK.with(|work| {
            let mut counts = work.get();
            counts.seed_materializations += 1;
            work.set(counts);
        });
        let mut segments = Vec::new();
        for (diff, units) in chunks.diffs[range.clone()]
            .iter()
            .zip(&chunks.units[range.clone()])
        {
            if let Out::Any(Any::String(text)) = &diff.insert {
                push_segment(&mut segments, text, diff.attributes.as_deref(), units.get());
            }
        }
        Cow::Owned(segments)
    }

    fn parts(&self) -> impl Iterator<Item = SegmentPart<'_>> + Clone {
        self.segments
            .iter()
            .map(|segment| SegmentPart {
                text: &segment.text,
                units: segment.units,
                attrs: Some(&segment.attrs),
            })
            .chain(self.chunks.iter().flat_map(|(chunks, range)| {
                chunks.diffs[range.clone()]
                    .iter()
                    .zip(&chunks.units[range.clone()])
                    .filter_map(|(diff, units)| match &diff.insert {
                        Out::Any(Any::String(text)) => Some(SegmentPart {
                            text: text.as_ref(),
                            units: units.get(),
                            attrs: diff.attributes.as_deref(),
                        }),
                        _ => None,
                    })
            }))
            .filter(|part| !part.text.is_empty())
    }

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
    pub(crate) delete_bounds: Option<PatchBounds>,
}

#[derive(Clone, Copy, Debug)]
pub(crate) struct PatchBounds {
    slot: usize,
    start: usize,
    end: usize,
    left_units: u32,
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

fn push_segment(segments: &mut Vec<TextSegment>, text: &str, attrs: Option<&Attrs>, units: u32) {
    if let Some(last) = segments.last_mut()
        && same_attributes(&last.attrs, attrs)
    {
        last.text.push_str(text);
        last.units += units;
    } else {
        segments.push(TextSegment {
            text: text.to_owned(),
            units,
            attrs: attributes(attrs),
        });
    }
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

fn stateful_value(key: &str, value: &Any) -> bool {
    if matches!(value, Any::Null | Any::Undefined | Any::Bool(false)) {
        return false;
    }
    match key {
        "fieldCodeMarks" | "fieldResultBlocks" => !any_strings(Some(value)).is_empty(),
        "fieldCodeTarget" => value_string(Some(value)).is_some(),
        _ => match value {
            Any::Map(map) => map.iter().any(|(key, value)| stateful_value(key, value)),
            Any::Array(values) => values.iter().any(|value| stateful_value("", value)),
            _ => false,
        },
    }
}

pub(super) fn preview_touches_state<T: ReadTxn>(
    diff: &yrs::types::text::Diff<YChange>,
    txn: &T,
) -> bool {
    preview_touches_state_inner(diff, txn, &mut BTreeSet::new())
}

fn preview_touches_state_inner<T: ReadTxn>(
    diff: &yrs::types::text::Diff<YChange>,
    txn: &T,
    active_stories: &mut BTreeSet<String>,
) -> bool {
    if diff
        .attributes
        .as_deref()
        .is_some_and(|attrs| attrs.iter().any(|(key, value)| stateful_value(key, value)))
    {
        return true;
    }
    match &diff.insert {
        Out::Any(Any::String(_)) => false,
        Out::YMap(mark) if is_pilcrow(mark, txn) => pilcrow_values(mark, txn)
            .iter()
            .any(|(key, value)| stateful_value(key, value)),
        Out::YMap(mark) => {
            let values = pilcrow_values(mark, txn);
            if values.iter().any(|(key, value)| stateful_value(key, value)) {
                return true;
            }
            if value_string(values.get("_kind")).as_deref() == Some("table") {
                let Some(Any::Array(rows)) = values.get("rows") else {
                    return true;
                };
                for row in rows.iter() {
                    let Some(Any::Array(cells)) = any_map(row).and_then(|row| row.get("cells"))
                    else {
                        return true;
                    };
                    for cell in cells.iter() {
                        let Some(story_id) =
                            any_map(cell).and_then(|cell| map_string(cell, "story"))
                        else {
                            return true;
                        };
                        if !active_stories.insert(story_id.clone()) {
                            return true;
                        }
                        let Ok(story) = story_ref(txn, &story_id) else {
                            return true;
                        };
                        let stateful = story
                            .diff(txn, YChange::identity)
                            .iter()
                            .any(|chunk| preview_touches_state_inner(chunk, txn, active_stories));
                        active_stories.remove(&story_id);
                        if stateful {
                            return true;
                        }
                    }
                }
                return false;
            }
            let seed_only = [
                "break",
                "pageBreak",
                "columnBreak",
                "image",
                "shape",
                "chart",
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
    pub(crate) fn snapshot(
        &self,
        doc: &EditingDoc,
        map: &LoweringMap,
    ) -> impl PartialEq + std::fmt::Debug + use<> {
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
                let pm = if self.deferred {
                    map.paragraph_blocks
                        .iter()
                        .find(|(_, source)| *source == seed.source)
                        .map_or(seed.pm_start, |(pm, _)| *pm)
                } else {
                    seed.pm_start
                };
                let raw = if self.deferred {
                    map.span_at(pm + 1)
                        .filter(|span| span.paragraph == seed.source)
                        .map(|span| span.raw_start)
                        .or_else(|| doc.paragraph_index("body").ok()?.para_span(id).map(|s| s.0))
                        .unwrap_or(seed.raw_start)
                } else {
                    seed.raw_start
                };
                (
                    id.clone(),
                    seed.tainted,
                    raw,
                    seed.slot,
                    seed.source,
                    pm,
                    seed.segments()
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

    pub(crate) fn fallback(enabled: bool) -> Self {
        Self {
            legacy: true,
            ..Self::new(enabled)
        }
    }

    pub(super) fn block(&mut self, blocked: bool) {
        self.blocked |= blocked;
        self.preview_blocked |= blocked;
    }

    pub(super) fn replace_seeds(&mut self, pm: &std::ops::Range<u64>, replacement: Self) {
        self.seeds.retain(|_, seed| !pm.contains(&seed.pm_start));
        #[cfg(test)]
        for id in replacement.seeds.keys() {
            self.discarded.remove(id);
        }
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
        #[cfg(test)]
        crate::engine::TYPING_EXTRA_WORK.with(|work| {
            let mut counts = work.get();
            counts.seed_refreshes += 1;
            work.set(counts);
        });
        let mut identities = HashSet::new();
        let duplicates: HashSet<_> = map
            .paragraphs
            .iter()
            .filter_map(|(_, id)| (!identities.insert(id.as_str())).then_some(id.as_str()))
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
            if self.dependent.contains(id) || duplicates.contains(id.as_str()) {
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
        chunk_index: usize,
        units: u32,
    ) {
        if self.blocked {
            return;
        }
        let attrs = diff.attributes.as_deref();
        let unsafe_attrs = attrs
            .into_iter()
            .flatten()
            .any(|(key, value)| unsafe_value(key, value));
        if self.legacy {
            self.block(unsafe_attrs);
            match &diff.insert {
                Out::Any(Any::String(_)) if story != "body" => {}
                Out::Any(Any::String(text)) => {
                    push_segment(&mut paragraph.segments, text, attrs, units);
                }
                Out::YMap(mark) if is_pilcrow(mark, txn) => {}
                Out::YMap(mark) => {
                    let values = pilcrow_values(mark, txn);
                    self.block(
                        values.iter().any(|(key, value)| unsafe_value(key, value))
                            || value_string(values.get("_kind")).as_deref() != Some("table"),
                    );
                    *paragraph = ParagraphSeed::default();
                }
                _ => self.block(true),
            }
            return;
        }
        self.preview_blocked |= unsafe_attrs;
        paragraph.tainted |= unsafe_attrs;
        match &diff.insert {
            Out::Any(Any::String(_)) if story != "body" || paragraph.tainted => {}
            Out::Any(Any::String(text)) => {
                if let Some(chunks) = &self.chunks {
                    chunks.units[chunk_index].set(units);
                    if let Some((_, range)) = &mut paragraph.chunks {
                        range.end = chunk_index + 1;
                    } else {
                        paragraph.chunks = Some((Rc::clone(chunks), chunk_index..chunk_index + 1));
                    }
                } else {
                    #[cfg(test)]
                    {
                        self.materialized_text_units += utf16_len(text);
                    }
                    push_segment(&mut paragraph.segments, text, attrs, units);
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
        if self.legacy {
            self.block(unsafe_values || sectioned && (story != "body" || !last));
        }
        self.preview_blocked |= unsafe_values || sectioned && (story != "body" || !last);
        self.block(sectioned && story != "body");
        paragraph.tainted |= unsafe_values || (!self.legacy && !safe);
        if story == "body" && !sectioned && !paragraph.tainted && !self.blocked {
            paragraph.raw_start = start;
            paragraph.pm_start = pm_start;
            paragraph.slot = slot;
            paragraph.source = source;
            paragraph.pilcrow = Some(mark.clone());
            paragraph.mark_attrs = attrs.cloned();
            let seed = std::mem::take(paragraph);
            #[cfg(test)]
            if !self.legacy {
                crate::engine::TYPING_EXTRA_WORK.with(|work| {
                    let mut counts = work.get();
                    counts.seeds += 1;
                    work.set(counts);
                });
            }
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
        if self.legacy {
            let mut identities = BTreeSet::new();
            self.block(map.paragraphs.iter().any(|(_, id)| !identities.insert(id)));
            self.block(!blocks.iter().all(shiftable));
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
            return;
        }
        #[cfg(test)]
        crate::engine::TYPING_EXTRA_WORK.with(|work| {
            let mut counts = work.get();
            counts.seed_scans += map.paragraphs.len() + blocks.len();
            work.set(counts);
        });
        let mut identities = HashSet::new();
        let duplicates: HashSet<_> = map
            .paragraphs
            .iter()
            .filter_map(|(_, id)| (!identities.insert(id.as_str())).then_some(id.as_str()))
            .collect();
        self.preview_blocked |= !duplicates.is_empty() || !blocks.iter().all(shiftable);
        self.seeds.retain(|id, seed| {
            if duplicates.contains(id.as_str()) || self.dependent.contains(id) {
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

    pub(crate) fn can_patch(&self, edit: &TextEdit) -> bool {
        !self.blocked
            && edit.removed <= 2
            && !edit
                .attributes
                .as_ref()
                .is_some_and(|attrs| attrs.iter().any(|(key, value)| unsafe_value(key, value)))
            && self.seeds.get(&edit.paragraph).is_some_and(|seed| {
                edit.delete_bounds.is_some() || patch_part_bounds(seed.parts(), edit).is_some()
            })
    }

    pub(crate) fn can_patch_before_insertion(&self, paragraph: &str) -> bool {
        let Some(pending) = self.edit.as_ref() else {
            return false;
        };
        pending.paragraph == paragraph && self.can_patch(pending)
    }

    pub(crate) fn delete_bounds(
        &self,
        paragraph: &str,
        offset: u32,
        removed: u32,
    ) -> Option<PatchBounds> {
        if removed > 2 {
            return None;
        }
        let range = patch_part_range(
            self.seeds.get(paragraph)?.parts(),
            offset,
            removed,
            true,
            true,
        )?;
        let mut bounds = part_byte_bounds(&range, removed)?;
        bounds.slot = range.merged_slot;
        bounds.start += range.merged_bytes;
        bounds.end += range.merged_bytes;
        bounds.left_units += range.merged_units;
        Some(bounds)
    }

    pub(crate) fn resume(&mut self) {
        self.validation_epoch = self.validation_epoch.wrapping_add(1);
    }

    pub(crate) fn discard_pending(&mut self) {
        self.suspend();
    }

    #[cfg(test)]
    pub(crate) fn discarded_seeds(&self) -> &BTreeSet<String> {
        &self.discarded
    }

    #[cfg(test)]
    pub(crate) fn has_seed(&self, id: &str) -> bool {
        self.seeds.contains_key(id)
    }

    #[cfg(test)]
    pub(crate) fn exclude_discarded_seeds(&mut self, discarded: &BTreeSet<String>) {
        for id in discarded {
            self.seeds.remove(id);
        }
        self.discarded.clone_from(discarded);
    }

    pub(crate) fn refresh_coordinates(
        &mut self,
        id: &str,
        raw: u32,
        blocks: &[Rc<LayoutBlock>],
        map: &LoweringMap,
    ) -> Option<()> {
        let seed = self.seeds.get_mut(id)?;
        let LayoutBlock::Paragraph(current) = blocks.get(seed.slot)?.as_ref() else {
            return None;
        };
        if !matches!(&current.id, BlockId::Str(current_id) if current_id == id)
            || !map
                .paragraphs
                .get(seed.source as usize)
                .is_some_and(|(story, paragraph)| *story == 0 && paragraph == id)
        {
            return None;
        }
        seed.raw_start = raw;
        seed.pm_start = current.pm_start? as u64;
        Some(())
    }

    pub(crate) fn refresh_seed<T: ReadTxn>(
        &mut self,
        id: &str,
        blocks: &[Rc<LayoutBlock>],
        txn: &T,
        env: &RenderEnv,
    ) -> Option<()> {
        let seed = self.seeds.get_mut(id)?;
        let block = blocks.get(seed.slot)?;
        let identity = (self.validation_epoch, Rc::as_ptr(block) as usize);
        if seed.validated == Some(identity) {
            return Some(());
        }
        #[cfg(test)]
        crate::engine::TYPING_EXTRA_WORK.with(|work| {
            let mut counts = work.get();
            counts.seed_validations += 1;
            work.set(counts);
        });
        let LayoutBlock::Paragraph(current) = block.as_ref() else {
            return None;
        };
        let raw = seed.raw_start;
        let pm = current.pm_start? as u64;
        let mut runs = Vec::new();
        let mut units = 0;
        for segment in seed.segments().iter() {
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
        let mut output = super::preview::LoweringOutput::default();
        let mut expected = flush_paragraph(
            runs,
            seed.pilcrow.as_ref()?,
            seed.mark_attrs.as_ref(),
            txn,
            "body",
            env,
            pm,
            units,
            &mut ListState::default(),
            (&mut output, seed.source),
            Vec::new(),
        );
        if current
            .attrs
            .as_ref()
            .is_some_and(|attrs| attrs.num_pr.is_some() || attrs.list_marker.is_some())
        {
            expected.attrs = current.attrs.clone();
        }
        if expected != *current {
            return None;
        }
        seed.validated = Some(identity);
        Some(())
    }

    pub(crate) fn suspend(&mut self) {
        if let Some(edit) = self.edit.take() {
            self.discard_seed(&edit.paragraph);
        }
        self.deferred = true;
    }

    pub(crate) fn discard_seed(&mut self, paragraph: &str) {
        #[cfg(test)]
        crate::engine::TYPING_EXTRA_WORK.with(|work| {
            let mut counts = work.get();
            counts.seed_removals += 1;
            work.set(counts);
        });
        self.seeds.remove(paragraph);
        #[cfg(test)]
        self.discarded.insert(paragraph.to_owned());
    }

    pub(crate) fn patch<T: ReadTxn>(
        &mut self,
        blocks: &mut [Rc<LayoutBlock>],
        map: &mut LoweringMap,
        revealable: &mut [LayoutBlock],
        txn: &T,
        env: &RenderEnv,
        edit: &TextEdit,
    ) -> Option<super::preview::ParagraphEdit> {
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
        let old_segments = seed.segments();
        let old_units: u32 = old_segments
            .iter()
            .map(|segment| utf16_len(&segment.text))
            .sum();
        let old_end = pm_start + u64::from(old_units) + 2;
        let segments = patch_segments(&old_segments, edit)?;
        drop(old_segments);
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
        #[cfg(test)]
        if seed.chunks.is_some() {
            self.materialized_text_units += old_units;
        }
        seed.chunks = None;
        blocks[slot] = Rc::new(LayoutBlock::Paragraph(paragraph));
        seed.validated = Some((self.validation_epoch, Rc::as_ptr(&blocks[slot]) as usize));
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
        if !self.deferred {
            for seed in self.seeds.values_mut() {
                if seed.raw_start > raw {
                    seed.raw_start = (i64::from(seed.raw_start) + delta) as u32;
                    seed.pm_start = (seed.pm_start as i64 + delta) as u64;
                }
            }
        }
        Some(super::preview::ParagraphEdit {
            raw: raw..raw + old_units + 1,
            pm: pm_start..old_end,
            delta,
        })
    }
}

fn patch_segment_bounds(segments: &[TextSegment], edit: &TextEdit) -> Option<PatchBounds> {
    patch_part_bounds(
        segments.iter().map(|segment| SegmentPart {
            text: &segment.text,
            units: segment.units,
            attrs: Some(&segment.attrs),
        }),
        edit,
    )
}

fn same_part_attributes(left: Option<&Attrs>, right: Option<&Attrs>) -> bool {
    let present = |attrs: Option<&Attrs>| {
        attrs
            .into_iter()
            .flatten()
            .filter(|(_, value)| **value != Any::Null)
            .count()
    };
    present(left) == present(right)
        && left.into_iter().flatten().all(|(key, value)| {
            *value == Any::Null || right.is_some_and(|attrs| attrs.get(key) == Some(value))
        })
}

fn patch_part_range<'a>(
    parts: impl Iterator<Item = SegmentPart<'a>>,
    offset: u32,
    removed: u32,
    empty: bool,
    merge: bool,
) -> Option<PartRange<'a>> {
    let mut before = 0_u32;
    let mut previous: Option<SegmentPart<'a>> = None;
    let mut merged_slot = 0;
    let mut merged_bytes = 0;
    let mut merged_units = 0;
    let mut remaining = parts.enumerate();
    let (slot, segment) = loop {
        let (slot, segment) = remaining.next()?;
        #[cfg(test)]
        crate::engine::TYPING_EXTRA_WORK.with(|work| {
            let mut counts = work.get();
            counts.preflight_chunks += 1;
            work.set(counts);
        });
        if merge && let Some(part) = previous {
            if same_part_attributes(part.attrs, segment.attrs) {
                merged_bytes += part.text.len();
                merged_units += part.units;
            } else {
                merged_slot += 1;
                merged_bytes = 0;
                merged_units = 0;
            }
        }
        let end = before.checked_add(segment.units)?;
        if offset < end || (removed == 0 && offset == end) {
            break (slot, segment);
        }
        before = end;
        previous = Some(segment);
    };
    let start = offset.checked_sub(before)? as usize;
    let end = start.checked_add(removed as usize)?;
    if end > segment.units as usize {
        return None;
    }
    if removed == 2 && segment.units as usize == segment.text.len() {
        return None;
    }
    let next = remaining.next().map(|(_, part)| part);
    #[cfg(test)]
    if next.is_some() {
        crate::engine::TYPING_EXTRA_WORK.with(|work| {
            let mut counts = work.get();
            counts.preflight_chunks += 1;
            work.set(counts);
        });
    }
    if removed != 0
        && start == 0
        && end == segment.units as usize
        && empty
        && (previous.is_some() || next.is_some())
        && !previous.is_some_and(|part| same_part_attributes(part.attrs, segment.attrs))
        && !next.is_some_and(|part| same_part_attributes(part.attrs, segment.attrs))
    {
        return None;
    }
    Some(PartRange {
        slot,
        segment,
        start,
        end,
        merged_slot,
        merged_bytes,
        merged_units,
    })
}

struct PartRange<'a> {
    slot: usize,
    segment: SegmentPart<'a>,
    start: usize,
    end: usize,
    merged_slot: usize,
    merged_bytes: usize,
    merged_units: u32,
}

fn patch_part_bounds<'a>(
    parts: impl Iterator<Item = SegmentPart<'a>> + Clone,
    edit: &TextEdit,
) -> Option<PatchBounds> {
    let mut parts = parts.peekable();
    if parts.peek().is_none() {
        return (edit.offset == 0 && edit.removed == 0).then_some(PatchBounds {
            slot: 0,
            start: 0,
            end: 0,
            left_units: 0,
        });
    }
    let range = patch_part_range(
        parts,
        edit.offset,
        edit.removed,
        edit.text.is_empty(),
        false,
    )?;
    part_byte_bounds(&range, edit.removed)
}

fn part_byte_bounds(range: &PartRange<'_>, removed: u32) -> Option<PatchBounds> {
    let PartRange {
        slot,
        segment,
        start,
        end,
        ..
    } = *range;
    let mut units = 0;
    let mut start_byte = None;
    let mut end_byte = None;
    for (byte, ch) in segment.text.char_indices() {
        #[cfg(test)]
        crate::engine::TYPING_EXTRA_WORK.with(|work| {
            let mut counts = work.get();
            counts.preflight_text_units += ch.len_utf16();
            work.set(counts);
        });
        if units == start {
            start_byte = Some(byte);
        }
        if units == end {
            end_byte = Some(byte);
            break;
        }
        units += ch.len_utf16();
    }
    if units == start {
        start_byte.get_or_insert(segment.text.len());
    }
    if units == end {
        end_byte.get_or_insert(segment.text.len());
    }
    let (start, end) = (start_byte?, end_byte?);
    if removed != 0 && segment.text[start..end].chars().count() != 1 {
        return None;
    }
    Some(PatchBounds {
        slot,
        start,
        end,
        left_units: range.start as u32,
    })
}

fn patch_segments(segments: &[TextSegment], edit: &TextEdit) -> Option<Vec<TextSegment>> {
    let PatchBounds {
        slot,
        start,
        end,
        left_units,
    } = match edit.delete_bounds {
        Some(bounds) => bounds,
        None => patch_segment_bounds(segments, edit)?,
    };
    if segments.is_empty() {
        return Some(vec![TextSegment {
            text: edit.text.clone(),
            units: utf16_len(&edit.text),
            attrs: attributes(edit.attributes.as_ref()),
        }]);
    }
    let mut result = segments.to_vec();
    let segment = &result[slot];
    let left = segment.text[..start].to_owned();
    let right = segment.text[end..].to_owned();
    if edit.removed != 0 && left.is_empty() && right.is_empty() && edit.text.is_empty() {
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
            units: left_units,
            attrs: segment.attrs.clone(),
        },
        TextSegment {
            text: edit.text.clone(),
            units: utf16_len(&edit.text),
            attrs,
        },
        TextSegment {
            text: right,
            units: segment.units - left_units - edit.removed,
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
            last.units += segment.units;
        } else {
            merged.push(segment);
        }
    }
    Some(merged)
}

#[cfg(test)]
#[test]
fn cached_delete_bounds_match_coalesced_seed_chunks() {
    let attrs = Attrs::from([("bold".into(), Any::Bool(true))]);
    let null = Attrs::from([("italic".into(), Any::Null)]);
    let chunks = Rc::new(SeedChunks::new(vec![
        yrs::types::text::Diff::with_change(Out::Any(Any::from("ab")), None, None),
        yrs::types::text::Diff::with_change(Out::Any(Any::from("α😀")), Some(Box::new(null)), None),
        yrs::types::text::Diff::with_change(Out::Any(Any::from("cd")), Some(Box::new(attrs)), None),
        yrs::types::text::Diff::with_change(Out::Any(Any::from("ef")), None, None),
    ]));
    for (diff, units) in chunks.diffs.iter().zip(&chunks.units) {
        let Out::Any(Any::String(text)) = &diff.insert else {
            unreachable!();
        };
        units.set(utf16_len(text));
    }
    let mut local = LocalLowering::new(true);
    local.seeds.insert(
        "p".to_owned(),
        ParagraphSeed {
            chunks: Some((chunks, 0..4)),
            ..ParagraphSeed::default()
        },
    );
    for (offset, removed) in [(2, 1), (3, 2), (5, 1), (7, 1)] {
        let segments = local.seeds["p"].segments();
        let mut edit = TextEdit {
            paragraph: "p".to_owned(),
            offset,
            removed,
            text: String::new(),
            attributes: None,
            epochs: (0, 1),
            delete_bounds: None,
        };
        let expected = patch_segments(&segments, &edit).unwrap();
        edit.delete_bounds = Some(local.delete_bounds("p", offset, removed).unwrap());
        crate::engine::TYPING_EXTRA_WORK.with(|work| work.set(Default::default()));
        assert!(local.can_patch(&edit));
        let actual = patch_segments(&segments, &edit).unwrap();
        crate::engine::TYPING_EXTRA_WORK.with(|work| {
            assert_eq!(work.get().preflight_chunks, 0);
            assert_eq!(work.get().preflight_text_units, 0);
        });
        let snapshot = |segments: &[TextSegment]| {
            segments
                .iter()
                .map(|segment| (segment.text.clone(), segment.units, segment.attrs.clone()))
                .collect::<Vec<_>>()
        };
        assert_eq!(snapshot(&actual), snapshot(&expected));
    }
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

pub(crate) fn shift_block(block: &mut LayoutBlock, delta: i64) {
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
