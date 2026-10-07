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
    pub(super) chunks: Option<Rc<Vec<yrs::types::text::Diff<YChange>>>>,
    pub(crate) retained: Option<Box<LocalLowering>>,
    invalidated: BTreeSet<String>,
    validation_epoch: u64,
    rebases: Vec<SeedRebase>,
    structural: Option<SeedRebase>,
    dependent: BTreeSet<String>,
    pub(crate) edit: Option<TextEdit>,
    #[cfg(test)]
    pub(crate) materialized_text_units: u32,
}

#[derive(Debug, Default)]
pub(super) struct ParagraphSeed {
    tainted: bool,
    raw_start: u32,
    slot: usize,
    source: u32,
    segments: Vec<TextSegment>,
    chunks: Option<(
        Rc<Vec<yrs::types::text::Diff<YChange>>>,
        std::ops::Range<usize>,
    )>,
    pilcrow: Option<MapRef>,
    mark_attrs: Option<Attrs>,
    pm_start: u64,
    validated: Option<(u64, usize)>,
    rebased: usize,
}

#[derive(Debug)]
struct SeedRebase {
    after_source: Option<u32>,
    paragraphs: BTreeSet<String>,
    slot_delta: isize,
    source_delta: i64,
}

#[derive(Clone, Debug)]
struct TextSegment {
    text: String,
    attrs: Attrs,
}

impl ParagraphSeed {
    fn segments(&self) -> Cow<'_, [TextSegment]> {
        let Some((chunks, range)) = &self.chunks else {
            return Cow::Borrowed(&self.segments);
        };
        let mut segments = Vec::new();
        for diff in &chunks[range.clone()] {
            if let Out::Any(Any::String(text)) = &diff.insert {
                push_segment(&mut segments, text, diff.attributes.as_deref());
            }
        }
        Cow::Owned(segments)
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

fn push_segment(segments: &mut Vec<TextSegment>, text: &str, attrs: Option<&Attrs>) {
    if let Some(last) = segments.last_mut()
        && same_attributes(&last.attrs, attrs)
    {
        last.text.push_str(text);
    } else {
        segments.push(TextSegment {
            text: text.to_owned(),
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
    preview_touches_state_inner(diff, txn, &mut BTreeSet::new(), true)
}

pub(super) fn preview_chunk_touches_state<T: ReadTxn>(
    diff: &yrs::types::text::Diff<YChange>,
    txn: &T,
) -> bool {
    preview_touches_state_inner(diff, txn, &mut BTreeSet::new(), false)
}

fn preview_touches_state_inner<T: ReadTxn>(
    diff: &yrs::types::text::Diff<YChange>,
    txn: &T,
    active_stories: &mut BTreeSet<String>,
    descend: bool,
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
                if !descend {
                    return false;
                }
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
                            .any(|chunk| {
                                preview_touches_state_inner(chunk, txn, active_stories, true)
                            });
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

    pub(super) fn replace_seeds(&mut self, pm: &std::ops::Range<u64>, mut replacement: Self) {
        self.seeds.retain(|_, seed| !pm.contains(&seed.pm_start));
        for seed in replacement.seeds.values_mut() {
            seed.rebased = self.rebases.len();
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
            seed.rebased = self.rebases.len();
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
                    push_segment(&mut paragraph.segments, text, attrs);
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
                    push_segment(&mut paragraph.segments, text, attrs);
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
                !self.invalidated.contains(&edit.paragraph)
                    && patch_segment_bounds(&seed.segments(), edit).is_some()
            })
    }

    pub(crate) fn can_patch_after_pending(&self, edit: &TextEdit) -> bool {
        if edit.removed > 2 {
            return false;
        }
        let Some(pending) = self.edit.as_ref() else {
            return self.can_patch(edit);
        };
        if pending.paragraph != edit.paragraph {
            return false;
        }
        self.can_patch(pending)
            && self.seeds.get(&edit.paragraph).is_some_and(|seed| {
                patch_segments(&seed.segments(), pending)
                    .is_some_and(|segments| patch_segment_bounds(&segments, edit).is_some())
            })
    }

    #[allow(clippy::too_many_arguments)]
    pub(super) fn recover_seed(
        &mut self,
        id: &str,
        mark: &MapRef,
        values: &BTreeMap<String, Any>,
        attrs: Option<&Attrs>,
        positions: (u32, u64, usize, u32),
        chunks: std::ops::Range<usize>,
        safe: bool,
    ) {
        let Some(retained) = self.retained.as_mut() else {
            return;
        };
        if retained.invalidated.is_empty() || !retained.invalidated.remove(id) {
            return;
        }
        if let Some(structural) = &mut retained.structural
            && structural.paragraphs.contains(id)
        {
            structural.after_source = Some(
                structural
                    .after_source
                    .map_or(positions.3, |source| source.max(positions.3)),
            );
        }
        if !safe || retained.dependent.contains(id) {
            retained.seeds.remove(id);
            return;
        }
        let seed = if retained.structural.is_some() {
            let Some(shared) = &self.chunks else {
                retained.seeds.remove(id);
                return;
            };
            if values.iter().any(|(key, value)| unsafe_value(key, value))
                || attrs
                    .is_some_and(|attrs| attrs.iter().any(|(key, value)| unsafe_value(key, value)))
                || shared[chunks.clone()].iter().any(|diff| {
                    !matches!(diff.insert, Out::Any(Any::String(_)))
                        || diff.attributes.as_deref().is_some_and(|attrs| {
                            attrs.iter().any(|(key, value)| unsafe_value(key, value))
                        })
                })
            {
                retained.seeds.remove(id);
                return;
            }
            #[cfg(test)]
            crate::engine::TYPING_EXTRA_WORK.with(|work| {
                let mut counts = work.get();
                counts.seeds += 1;
                work.set(counts);
            });
            retained.seeds.entry(id.to_owned()).or_default()
        } else {
            let Some(seed) = retained.seeds.get_mut(id) else {
                return;
            };
            seed
        };
        let (raw, pm, slot, source) = positions;
        seed.raw_start = raw;
        seed.pm_start = pm;
        seed.slot = slot;
        seed.source = source;
        seed.pilcrow = Some(mark.clone());
        seed.mark_attrs = attrs.cloned();
        seed.segments.clear();
        seed.chunks = self
            .chunks
            .as_ref()
            .map(|shared| (Rc::clone(shared), chunks));
        seed.validated = None;
        seed.rebased = retained.rebases.len() + usize::from(retained.structural.is_some());
    }

    pub(crate) fn prepare_structural(
        &mut self,
        blocks: usize,
        map: &LoweringMap,
        paragraphs: &[String],
    ) -> bool {
        if self.blocked || self.legacy {
            return false;
        }
        self.invalidated.extend(paragraphs.iter().cloned());
        let structural = self.structural.get_or_insert_with(|| SeedRebase {
            after_source: None,
            paragraphs: BTreeSet::new(),
            slot_delta: -(blocks as isize),
            source_delta: -(map.paragraphs.len() as i64),
        });
        structural.paragraphs.extend(paragraphs.iter().cloned());
        true
    }

    pub(crate) fn finish_structural(&mut self, blocks: &[LayoutBlock], map: &LoweringMap) {
        if let Some(mut rebase) = self.structural.take() {
            for id in std::mem::take(&mut rebase.paragraphs) {
                self.invalidated.remove(&id);
                let eligible = self.seeds.get(&id).is_some_and(|seed| {
                    let Some(LayoutBlock::Paragraph(paragraph)) = blocks.get(seed.slot) else {
                        return false;
                    };
                    seed.rebased == self.rebases.len() + 1
                        && matches!(&paragraph.id, BlockId::Str(current) if current == &id)
                        && paragraph.attrs.is_some()
                        && paragraph.pm_start == Some(seed.pm_start as f64)
                        && shiftable(&blocks[seed.slot])
                        && !page_break_changes_marker(paragraph, blocks.get(seed.slot + 1))
                });
                if !eligible {
                    self.seeds.remove(&id);
                }
            }
            rebase.slot_delta += blocks.len() as isize;
            rebase.source_delta += map.paragraphs.len() as i64;
            rebase.after_source = rebase
                .after_source
                .and_then(|source| u32::try_from(i64::from(source) - rebase.source_delta).ok());
            self.rebases.push(rebase);
        }
    }

    #[cfg(test)]
    pub(crate) fn structurally_deferred(&self) -> bool {
        !self.rebases.is_empty()
    }

    pub(crate) fn resume(&mut self) {
        self.validation_epoch = self.validation_epoch.wrapping_add(1);
    }

    pub(crate) fn refresh_coordinates(
        &mut self,
        id: &str,
        raw: u32,
        blocks: &[Rc<LayoutBlock>],
        map: &LoweringMap,
    ) -> Option<()> {
        let refreshed = self.rebase_coordinates(id, raw, blocks, map);
        if refreshed.is_none() {
            self.invalidated.insert(id.to_owned());
        }
        refreshed
    }

    fn rebase_coordinates(
        &mut self,
        id: &str,
        raw: u32,
        blocks: &[Rc<LayoutBlock>],
        map: &LoweringMap,
    ) -> Option<()> {
        let rebases = self.rebases.len();
        let seed = self.seeds.get_mut(id)?;
        #[cfg(test)]
        crate::engine::TYPING_EXTRA_WORK.with(|work| {
            let mut counts = work.get();
            counts.seed_rebases += rebases - seed.rebased;
            work.set(counts);
        });
        for rebase in &self.rebases[seed.rebased..] {
            if rebase
                .after_source
                .is_some_and(|source| seed.source > source)
            {
                seed.slot = seed.slot.checked_add_signed(rebase.slot_delta)?;
                seed.source = u32::try_from(i64::from(seed.source) + rebase.source_delta).ok()?;
            }
        }
        seed.rebased = rebases;
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
            if edit
                .attributes
                .as_ref()
                .is_some_and(|attrs| attrs.iter().any(|(key, value)| unsafe_value(key, value)))
            {
                self.seeds.remove(&edit.paragraph);
            } else {
                self.invalidated.insert(edit.paragraph);
            }
        }
        self.deferred = true;
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

fn patch_segment_bounds(
    segments: &[TextSegment],
    edit: &TextEdit,
) -> Option<(usize, usize, usize)> {
    if segments.is_empty() {
        if edit.offset != 0 || edit.removed != 0 {
            return None;
        }
        return Some((0, 0, 0));
    }
    let mut before = 0;
    let slot = segments.iter().position(|segment| {
        let end = before + utf16_len(&segment.text);
        if edit.offset < end || (edit.removed == 0 && edit.offset == end) {
            true
        } else {
            before = end;
            false
        }
    })?;
    let segment = &segments[slot];
    let start = edit.offset.checked_sub(before)? as usize;
    let end = start.checked_add(edit.removed as usize)?;
    let mut units = 0;
    let mut start_byte = None;
    let mut end_byte = None;
    for (byte, ch) in segment.text.char_indices() {
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
    if edit.removed != 0 && segment.text[start..end].chars().count() != 1 {
        return None;
    }
    if edit.removed != 0
        && start == 0
        && end == segment.text.len()
        && edit.text.is_empty()
        && segments.len() > 1
    {
        return None;
    }
    Some((slot, start, end))
}

fn patch_segments(segments: &[TextSegment], edit: &TextEdit) -> Option<Vec<TextSegment>> {
    let (slot, start, end) = patch_segment_bounds(segments, edit)?;
    if segments.is_empty() {
        return Some(vec![TextSegment {
            text: edit.text.clone(),
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
