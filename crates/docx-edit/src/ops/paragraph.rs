//! Paragraph ops: `split_paragraph`, `merge_paragraphs`, `set_paragraph_attrs` (+ wrappers),
//! `apply_paragraph_style` (op-contract §1 "Paragraph"), and the R5 paraId re-uniquing pass.

use std::collections::{BTreeMap, HashMap, HashSet};
use std::sync::Arc;

use yrs::types::Attrs;
use yrs::{Any, Map, MapPrelim, MapRef, Out, ReadTxn, Text, TextRef, TransactionMut};

use crate::format::{PROTECTED_ATTRS, Patch};
use crate::op::{OpError, OpResult, ParaBounds, Receipt, SplitReceipt, para_bounds};
use crate::ops::{
    adjacent_paragraph_change_revision_id, adjacent_revision_id, adopt_pilcrow, capture_pilcrow,
    revision_id_in_range, snapshot,
};
use crate::{
    DEL, EditCtx, EditingDoc, KIND_KEY, PARA_ID, PPR_CHANGE, PPR_DEL, PPR_INS, ParagraphId,
    Position, StoryRange, check_position, insertion_attrs, next_pilcrow, revision_value, story_ref,
};

/// The paragraph attrs a style definition controls. Applying a style resets every one of these
/// to the style's value or clears it (port of `paragraphAttrsFromResolvedStyle`).
pub const STYLE_CONTROLLED_PARA_ATTRS: [&str; 15] = [
    "alignment",
    "spaceBefore",
    "spaceAfter",
    "lineSpacing",
    "lineSpacingRule",
    "indentLeft",
    "indentRight",
    "indentFirstLine",
    "hangingIndent",
    "contextualSpacing",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "outlineLevel",
    "defaultTextFormatting",
];

/// The 7 style-controlled marks swept before a style's run formats are applied (port of
/// `makeApplyStyle`'s `styleControlledMarks`).
pub const STYLE_CONTROLLED_MARKS: [&str; 7] = [
    "bold",
    "italic",
    "fontSize",
    "fontFamily",
    "textColor",
    "underline",
    "strike",
];

/// The pPr subset an empty second half inherits on split (port of `INHERITED_PARA_ATTRS`;
/// `styleId` is `pStyle` in the story vocabulary).
const INHERITED_PARA_ATTRS: [&str; 7] = [
    "defaultTextFormatting",
    "pStyle",
    "lineSpacing",
    "lineSpacingRule",
    "spaceAfter",
    "spaceBefore",
    "contextualSpacing",
];

/// The `defaultTextFormatting` keys that cross a split (port of `styleCarryDtf` — the
/// font/size/color subset; bold/italic/underline etc. deliberately do not carry).
const STYLE_CARRY_DTF_KEYS: [&str; 4] = ["fontFamily", "fontSize", "fontSizeCs", "color"];

const BORDERS: &str = "borders";
const TABS: &str = "tabs";
const INDENT_LEFT: &str = "indentLeft";
const DEFAULT_TEXT_FORMATTING: &str = "defaultTextFormatting";

/// Default indent step in twips (0.5 inch), matching the PM commands.
pub const INDENT_STEP_TWIPS: f64 = 720.0;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum MergeDirection {
    /// Delete THIS paragraph's mark, merging it with the following paragraph.
    Forward,
    /// Delete the PREVIOUS paragraph's mark, merging this paragraph into it.
    Backward,
}

/// Which paragraphs an op targets (op-contract "paras: One|Range").
#[derive(Clone, Debug, PartialEq)]
pub enum ParaSelector {
    One(ParagraphId),
    Many(Vec<ParagraphId>),
    /// Every paragraph whose content or mark intersects the range.
    Range(StoryRange),
}

/// One tab stop; `pos` is in twips, `alignment` is the `w:tab` val (`left`, `center`, `right`,
/// `decimal`, `bar`), `leader` the optional leader character name.
#[derive(Clone, Debug, PartialEq)]
pub struct TabStop {
    pub pos: f64,
    pub alignment: String,
    pub leader: Option<String>,
}

impl TabStop {
    fn to_any(&self) -> Any {
        let mut map = HashMap::from([
            ("pos".into(), Any::Number(self.pos)),
            ("val".into(), Any::from(self.alignment.as_str())),
        ]);
        if let Some(leader) = &self.leader {
            map.insert("leader".into(), Any::from(leader.as_str()));
        }
        Any::Map(Arc::new(map))
    }
}

/// Tri-state paragraph attribute delta (op-contract §1 "Paragraph"). Spacing and indent values
/// are authored OOXML units (twips / line-spacing units), never pixels.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct ParaAttrDelta {
    pub alignment: Patch<String>,
    pub line_spacing: Patch<f64>,
    pub line_spacing_rule: Patch<String>,
    pub space_before: Patch<f64>,
    pub space_after: Patch<f64>,
    pub indent_left: Patch<f64>,
    pub indent_right: Patch<f64>,
    pub indent_first_line: Patch<f64>,
    pub hanging_indent: Patch<f64>,
    pub bidi: Patch<bool>,
    pub tabs: Patch<Vec<TabStop>>,
    /// The paragraph-mark run defaults (`defaultTextFormatting`), as an opaque attr map.
    pub default_text_formatting: Patch<BTreeMap<String, Any>>,
    /// The +30 opaque paragraph properties; `None` clears the key.
    pub other: BTreeMap<String, Option<Any>>,
}

/// A host-resolved paragraph style, injected because style resolution (styles.xml cascade) stays
/// outside the CRDT until S5.
///
/// `paragraph_attrs` is the `paragraphAttrsFromResolvedStyle` projection: values for the
/// [`STYLE_CONTROLLED_PARA_ATTRS`] keys (a missing or `Any::Null` entry clears the attr), plus
/// any list attrs when the style defines numbering. `run_marks` is the style's run formatting
/// lowered to story attr values (`bold`, `fontSize`, ...), applied after sweeping the
/// [`STYLE_CONTROLLED_MARKS`].
#[derive(Clone, Debug, Default, PartialEq)]
pub struct ResolvedStyleProjection {
    pub style_id: String,
    /// Host-verified existence. `false` → [`OpError::UnknownStyle`] before any mutation.
    pub known: bool,
    pub paragraph_attrs: BTreeMap<String, Any>,
    pub run_marks: BTreeMap<String, Any>,
}

struct TargetPara {
    story_id: String,
    story: TextRef,
    bounds: ParaBounds,
    map: MapRef,
}

fn all_targets<T: ReadTxn>(txn: &T) -> Vec<TargetPara> {
    let Some(stories) = txn.get_map(crate::STORIES) else {
        return Vec::new();
    };
    let mut story_ids: Vec<String> = stories.keys(txn).map(|key| key.to_string()).collect();
    story_ids.sort();
    let mut result = Vec::new();
    for story_id in story_ids {
        let Some(Out::YText(story)) = stories.get(txn, &story_id) else {
            continue;
        };
        let pilcrow_maps: HashMap<u32, MapRef> = crate::pilcrows(&story, txn).into_iter().collect();
        for bounds in para_bounds(&story, txn) {
            if let Some(map) = pilcrow_maps.get(&bounds.pilcrow) {
                result.push(TargetPara {
                    story_id: story_id.clone(),
                    story: story.clone(),
                    bounds,
                    map: map.clone(),
                });
            }
        }
    }
    result
}

/// Resolves a selector to pilcrow targets, validating BEFORE any mutation.
fn resolve_selector<T: ReadTxn>(txn: &T, selector: &ParaSelector) -> OpResult<Vec<TargetPara>> {
    let all = all_targets(txn);
    match selector {
        ParaSelector::One(id) => {
            let target = all
                .into_iter()
                .find(|target| target.bounds.para_id == *id)
                .ok_or_else(|| OpError::UnknownPara(id.clone()))?;
            Ok(vec![target])
        }
        ParaSelector::Many(ids) => {
            let mut by_id: HashMap<String, TargetPara> = all
                .into_iter()
                .map(|target| (target.bounds.para_id.clone(), target))
                .collect();
            ids.iter()
                .map(|id| {
                    by_id
                        .remove(id)
                        .ok_or_else(|| OpError::UnknownPara(id.clone()))
                })
                .collect()
        }
        ParaSelector::Range(range) => {
            if range.end < range.start {
                return Err(OpError::InvalidRange {
                    start: range.start,
                    end: range.end,
                });
            }
            let targets: Vec<TargetPara> = all
                .into_iter()
                .filter(|target| {
                    target.story_id == range.story
                        && target.bounds.start <= range.end
                        && target.bounds.pilcrow >= range.start
                })
                .collect();
            if targets.is_empty() {
                return Err(OpError::UnknownStory(range.story.clone()));
            }
            Ok(targets)
        }
    }
}

fn set_or_remove(txn: &mut TransactionMut<'_>, map: &MapRef, key: &str, value: Option<Any>) {
    match value {
        Some(value) if value != Any::Null => {
            map.insert(txn, key.to_owned(), value);
        }
        _ => {
            map.remove(txn, key);
        }
    }
}

/// Writes a style's paragraph-attr projection: the 15 style-controlled attrs are reset to the
/// projection's value (or cleared), extra projection keys (list attrs) applied as-is.
fn apply_paragraph_attr_projection(
    txn: &mut TransactionMut<'_>,
    map: &MapRef,
    attrs: &BTreeMap<String, Any>,
) -> OpResult<()> {
    for key in STYLE_CONTROLLED_PARA_ATTRS {
        set_or_remove(txn, map, key, attrs.get(key).cloned());
    }
    for (key, value) in attrs {
        if STYLE_CONTROLLED_PARA_ATTRS.contains(&key.as_str()) {
            continue;
        }
        if matches!(key.as_str(), PARA_ID | KIND_KEY) {
            return Err(OpError::ReservedKey(key.clone()));
        }
        set_or_remove(txn, map, key, Some(value.clone()));
    }
    Ok(())
}

/// Reduces a `defaultTextFormatting` map to the font/size/color subset that crosses a split.
fn style_carry_dtf(value: &Any) -> Option<Any> {
    let Any::Map(map) = value else {
        return None;
    };
    let subset: HashMap<String, Any> = map
        .iter()
        .filter(|(key, _)| STYLE_CARRY_DTF_KEYS.contains(&key.as_str()))
        .map(|(key, value)| (key.clone(), value.clone()))
        .collect();
    if subset.is_empty() {
        None
    } else {
        Some(Any::Map(Arc::new(subset)))
    }
}

impl EditingDoc {
    /// Splits a paragraph by inserting exactly ONE pilcrow embed (op-contract §1).
    ///
    /// The new pilcrow terminates the FIRST half with the source paragraph's full pPr and its
    /// ORIGINAL paraId; the original pilcrow is re-minted with a fresh paraId and becomes the
    /// second half's mark. Second-half inheritance ports `applyPostSplitInheritance`:
    ///
    /// - mid-paragraph split: the second half keeps its pPr, borders ALWAYS cleared;
    /// - split at the paragraph end (empty second half): the second half keeps only the
    ///   `INHERITED_PARA_ATTRS` subset, with `defaultTextFormatting` reduced to font/size/color;
    /// - split at the end WITH a host-injected `w:next` style: the second half switches to that
    ///   style's projection instead (borders cleared either way).
    ///
    /// Suggesting mode stamps the inserted pilcrow with `ins` and `pPrIns`.
    pub fn split_paragraph(
        &self,
        ctx: &EditCtx,
        at: Position,
        next_style: Option<&ResolvedStyleProjection>,
    ) -> OpResult<SplitReceipt> {
        if let Some(projection) = next_style
            && !projection.known
        {
            return Err(OpError::UnknownStyle(projection.style_id.clone()));
        }
        let second_para_id = self.next_id();
        let mut txn = self.transact_for(ctx);
        let story = story_ref(&txn, &at.story)?;
        check_position(&story, &txn, at.index)?;
        let chunks = snapshot(&story, &txn);
        let revision_id = ctx.is_suggesting().then(|| {
            adjacent_revision_id(&chunks, at.index, crate::INS, &ctx.author)
                .or_else(|| {
                    adjacent_paragraph_change_revision_id(&chunks, at.index, &txn, &ctx.author)
                })
                .unwrap_or_else(|| self.next_id())
        });
        let (orig_index, orig_map) =
            next_pilcrow(&story, &txn, at.index).ok_or(OpError::ExpectedPilcrow {
                story: at.story.clone(),
                index: at.index,
            })?;
        let (first_para_id, props) = capture_pilcrow(&orig_map, &txn);
        let second_half_empty = orig_index == at.index;

        let ins = revision_id
            .as_ref()
            .map(|id| revision_value(id, &ctx.revision_author()));
        let new_pilcrow = story.insert_embed_with_attributes(
            &mut txn,
            at.index,
            MapPrelim::default(),
            insertion_attrs(ins, None),
        );
        new_pilcrow.insert(&mut txn, KIND_KEY, crate::PILCROW_KIND);
        new_pilcrow.insert(&mut txn, PARA_ID, first_para_id.as_str());
        for (key, value) in &props {
            new_pilcrow.insert(&mut txn, key.clone(), value.clone());
        }
        if let Some(id) = revision_id.as_ref() {
            new_pilcrow.insert(
                &mut txn,
                PPR_INS,
                revision_value(id, &ctx.revision_author()),
            );
        }

        // The original pilcrow now terminates the second half: re-mint its identity, then apply
        // post-split inheritance.
        orig_map.insert(&mut txn, PARA_ID, second_para_id.as_str());
        if second_half_empty {
            if let Some(next) = next_style {
                // `w:next` switch (port of applyNextParagraphStyle): fresh attrs + the next
                // style's projection; borders cleared.
                for (key, _) in &props {
                    orig_map.remove(&mut txn, key);
                }
                orig_map.insert(&mut txn, "pStyle", next.style_id.as_str());
                apply_paragraph_attr_projection(&mut txn, &orig_map, &next.paragraph_attrs)?;
                orig_map.remove(&mut txn, BORDERS);
            } else {
                // Blank-attr inheritance: keep only the inherited subset; dtf reduced to the
                // font/size/color carry. Borders fall out of the sweep.
                for (key, value) in &props {
                    if !INHERITED_PARA_ATTRS.contains(&key.as_str()) {
                        orig_map.remove(&mut txn, key);
                    } else if key == DEFAULT_TEXT_FORMATTING {
                        set_or_remove(
                            &mut txn,
                            &orig_map,
                            DEFAULT_TEXT_FORMATTING,
                            style_carry_dtf(value),
                        );
                    }
                }
            }
        } else {
            // Mid-paragraph split keeps the second half's pPr; Word never propagates w:pBdr.
            orig_map.remove(&mut txn, BORDERS);
        }
        Ok(SplitReceipt {
            first_para_id,
            second_para_id,
            revision_ids: revision_id.into_iter().collect(),
        })
    }

    /// Merges the paragraph with its neighbor by deleting the boundary pilcrow; the survivor
    /// adopts the deleted mark's pPr + paraId (the earlier paragraph's identity wins — R6).
    /// Suggesting mode retains the mark with `del` + `pPrDel` instead.
    pub fn merge_paragraphs(
        &self,
        ctx: &EditCtx,
        para: &str,
        direction: MergeDirection,
    ) -> OpResult<Receipt> {
        let mut txn = self.transact_for(ctx);
        let targets = all_targets(&txn);
        let index = targets
            .iter()
            .position(|target| target.bounds.para_id == para)
            .ok_or_else(|| OpError::UnknownPara(para.to_owned()))?;
        let boundary_index = match direction {
            MergeDirection::Forward => {
                let is_last_in_story = targets
                    .get(index + 1)
                    .is_none_or(|next| next.story_id != targets[index].story_id);
                if is_last_in_story {
                    return Err(OpError::CannotMergeFinalParagraph(para.to_owned()));
                }
                index
            }
            MergeDirection::Backward => {
                let has_previous =
                    index > 0 && targets[index - 1].story_id == targets[index].story_id;
                if !has_previous {
                    return Err(OpError::NoParagraphBefore(para.to_owned()));
                }
                index - 1
            }
        };
        let boundary = &targets[boundary_index];
        let survivor = &targets[boundary_index + 1];
        let story = boundary.story.clone();
        let pilcrow_index = boundary.bounds.pilcrow;
        let own_insert = ctx
            .is_suggesting()
            .then(|| paragraph_revision_id(&boundary.map, &txn, PPR_INS, &ctx.author))
            .flatten();
        let revision_id = (ctx.is_suggesting() && own_insert.is_none()).then(|| {
            adjacent_revision_id(&snapshot(&story, &txn), pilcrow_index, DEL, &ctx.author)
                .unwrap_or_else(|| self.next_id())
        });

        if own_insert.is_some() {
            // Backspacing over this author's still-pending split retracts the
            // suggestion itself; it must not author a second pPrDel revision.
            let (donor_id, mut donor_props) = capture_pilcrow(&boundary.map, &txn);
            donor_props.retain(|(key, _)| !matches!(key.as_str(), PPR_INS | PPR_DEL));
            story.remove_range(&mut txn, pilcrow_index, 1);
            adopt_pilcrow(&mut txn, &survivor.map, &donor_id, &donor_props);
        } else if let Some(id) = revision_id.as_ref() {
            let revision = revision_value(id, &ctx.revision_author());
            story.format(
                &mut txn,
                pilcrow_index,
                1,
                Attrs::from([(Arc::from(DEL), revision.clone())]),
            );
            boundary.map.insert(&mut txn, PPR_DEL, revision);
        } else {
            let (donor_id, donor_props) = capture_pilcrow(&boundary.map, &txn);
            story.remove_range(&mut txn, pilcrow_index, 1);
            adopt_pilcrow(&mut txn, &survivor.map, &donor_id, &donor_props);
        }
        let caret = crate::op::loc_range_in_txn(
            &boundary.story_id,
            &story,
            &txn,
            pilcrow_index,
            pilcrow_index,
        )?;
        Ok(Receipt {
            new_para_ids: Vec::new(),
            revision_ids: own_insert.into_iter().chain(revision_id).collect(),
            range: Some(caret),
        })
    }

    /// Applies a tri-state paragraph attribute delta to the selected paragraphs in one
    /// transaction.
    pub fn set_paragraph_attrs(
        &self,
        ctx: &EditCtx,
        selector: &ParaSelector,
        delta: &ParaAttrDelta,
    ) -> OpResult<Receipt> {
        for key in delta.other.keys() {
            if matches!(key.as_str(), PARA_ID | KIND_KEY) {
                return Err(OpError::ReservedKey(key.clone()));
            }
        }
        let mut txn = self.transact_for(ctx);
        let targets = resolve_selector(&txn, selector)?;
        let revision_id = ctx.is_suggesting().then(|| {
            targets
                .iter()
                .find_map(|target| {
                    revision_id_in_range(
                        &snapshot(&target.story, &txn),
                        target.bounds.start,
                        target.bounds.pilcrow + 1,
                        crate::INS,
                        &ctx.author,
                    )
                })
                .unwrap_or_else(|| self.next_id())
        });
        let mut changed = false;
        for target in &targets {
            let previous = paragraph_formatting(&target.map, &txn);
            apply_para_delta(&mut txn, &target.map, delta);
            let current = paragraph_formatting(&target.map, &txn);
            if let Some(id) = revision_id.as_ref()
                && previous != current
            {
                append_paragraph_property_change(
                    &mut txn,
                    &target.map,
                    id,
                    &ctx.revision_author(),
                    previous,
                    current,
                );
                changed = true;
            }
        }
        Ok(Receipt {
            revision_ids: revision_id.filter(|_| changed).into_iter().collect(),
            ..Receipt::default()
        })
    }

    /// Adds (or replaces, by position) a tab stop on the selected paragraphs.
    pub fn add_tab_stop(
        &self,
        ctx: &EditCtx,
        selector: &ParaSelector,
        stop: &TabStop,
    ) -> OpResult<Receipt> {
        let mut txn = self.transact_for(ctx);
        let targets = resolve_selector(&txn, selector)?;
        for target in &targets {
            let mut stops = read_tab_stops(&target.map, &txn);
            stops.retain(|existing| existing_pos(existing) != Some(stop.pos));
            stops.push(stop.to_any());
            stops.sort_by(|a, b| {
                existing_pos(a)
                    .unwrap_or(f64::MAX)
                    .total_cmp(&existing_pos(b).unwrap_or(f64::MAX))
            });
            target
                .map
                .insert(&mut txn, TABS, Any::Array(Arc::from(stops)));
        }
        Ok(Receipt::default())
    }

    /// Removes the tab stop at `pos` (twips) from the selected paragraphs.
    pub fn remove_tab_stop(
        &self,
        ctx: &EditCtx,
        selector: &ParaSelector,
        pos: f64,
    ) -> OpResult<Receipt> {
        let mut txn = self.transact_for(ctx);
        let targets = resolve_selector(&txn, selector)?;
        for target in &targets {
            let mut stops = read_tab_stops(&target.map, &txn);
            stops.retain(|existing| existing_pos(existing) != Some(pos));
            if stops.is_empty() {
                target.map.remove(&mut txn, TABS);
            } else {
                target
                    .map
                    .insert(&mut txn, TABS, Any::Array(Arc::from(stops)));
            }
        }
        Ok(Receipt::default())
    }

    /// Increases `indentLeft` by `step` twips (default 720 — half an inch, the PM command
    /// default).
    pub fn increase_indent(
        &self,
        ctx: &EditCtx,
        selector: &ParaSelector,
        step: Option<f64>,
    ) -> OpResult<Receipt> {
        let step = step.unwrap_or(INDENT_STEP_TWIPS);
        let mut txn = self.transact_for(ctx);
        let targets = resolve_selector(&txn, selector)?;
        for target in &targets {
            let current = number_prop(&target.map, &txn, INDENT_LEFT).unwrap_or(0.0);
            target
                .map
                .insert(&mut txn, INDENT_LEFT, Any::Number(current + step));
        }
        Ok(Receipt::default())
    }

    /// Decreases `indentLeft` by `step` twips (default 720), clamping at zero — a zero indent
    /// clears the attr (PM parity).
    pub fn decrease_indent(
        &self,
        ctx: &EditCtx,
        selector: &ParaSelector,
        step: Option<f64>,
    ) -> OpResult<Receipt> {
        let step = step.unwrap_or(INDENT_STEP_TWIPS);
        let mut txn = self.transact_for(ctx);
        let targets = resolve_selector(&txn, selector)?;
        for target in &targets {
            let current = number_prop(&target.map, &txn, INDENT_LEFT).unwrap_or(0.0);
            let next = (current - step).max(0.0);
            if next > 0.0 {
                target.map.insert(&mut txn, INDENT_LEFT, Any::Number(next));
            } else {
                target.map.remove(&mut txn, INDENT_LEFT);
            }
        }
        Ok(Receipt::default())
    }

    /// Sets (or clears) the paragraph-mark run defaults (`defaultTextFormatting`).
    pub fn set_paragraph_default_format(
        &self,
        ctx: &EditCtx,
        selector: &ParaSelector,
        formatting: Option<&BTreeMap<String, Any>>,
    ) -> OpResult<Receipt> {
        let mut txn = self.transact_for(ctx);
        let targets = resolve_selector(&txn, selector)?;
        for target in &targets {
            match formatting {
                Some(map) if !map.is_empty() => {
                    let value: HashMap<String, Any> =
                        map.iter().map(|(k, v)| (k.clone(), v.clone())).collect();
                    target
                        .map
                        .insert(&mut txn, DEFAULT_TEXT_FORMATTING, Any::Map(Arc::new(value)));
                }
                _ => {
                    target.map.remove(&mut txn, DEFAULT_TEXT_FORMATTING);
                }
            }
        }
        Ok(Receipt::default())
    }

    /// Applies a host-resolved paragraph style (compound; port of `makeApplyStyle`): sets the
    /// styleId, resets every style-controlled paragraph attr, sweeps the 7 style-controlled
    /// marks from the paragraph text, and applies the new style's run formats — all in ONE
    /// transaction. An unknown style errors before any mutation.
    pub fn apply_paragraph_style(
        &self,
        ctx: &EditCtx,
        selector: &ParaSelector,
        projection: &ResolvedStyleProjection,
    ) -> OpResult<Receipt> {
        if !projection.known {
            return Err(OpError::UnknownStyle(projection.style_id.clone()));
        }
        for key in projection.run_marks.keys() {
            if PROTECTED_ATTRS.contains(&key.as_str()) {
                return Err(OpError::InvalidFormatValue(format!(
                    "attribute {key:?} is not a formatting attribute"
                )));
            }
        }
        let mut txn = self.transact_for(ctx);
        let targets = resolve_selector(&txn, selector)?;
        for target in &targets {
            target
                .map
                .insert(&mut txn, "pStyle", projection.style_id.as_str());
            apply_paragraph_attr_projection(&mut txn, &target.map, &projection.paragraph_attrs)?;
            let len = target.bounds.len();
            if len > 0 {
                let mut attrs: Attrs = STYLE_CONTROLLED_MARKS
                    .iter()
                    .map(|key| (Arc::from(*key), Any::Null))
                    .collect();
                for (key, value) in &projection.run_marks {
                    attrs.insert(Arc::from(key.as_str()), value.clone());
                }
                target
                    .story
                    .format(&mut txn, target.bounds.start, len, attrs);
            }
        }
        Ok(Receipt::default())
    }

    /// R5 maintenance: re-mints duplicate paraIds (first occurrence keeps its id), under
    /// `Origin::System` so the pass never enters undo history. Returns `(old, new)` pairs.
    pub fn dedupe_para_ids(&self, now_iso: &str) -> OpResult<Vec<(ParagraphId, ParagraphId)>> {
        let ctx = EditCtx::system(now_iso);
        let mut renames = Vec::new();
        let mut txn = self.transact_for(&ctx);
        let targets = all_targets(&txn);
        let mut seen: HashSet<String> = HashSet::new();
        for target in targets {
            let id = target.bounds.para_id.clone();
            if seen.insert(id.clone()) {
                continue;
            }
            let minted = self.next_id();
            target.map.insert(&mut txn, PARA_ID, minted.as_str());
            renames.push((id, minted));
        }
        Ok(renames)
    }
}

fn paragraph_revision_id<T: ReadTxn>(
    map: &MapRef,
    txn: &T,
    key: &str,
    author: &str,
) -> Option<String> {
    let Some(Out::Any(Any::Map(revision))) = map.get(txn, key) else {
        return None;
    };
    if !matches!(revision.get("author"), Some(Any::String(value)) if value.as_ref() == author) {
        return None;
    }
    match revision.get("id").or_else(|| revision.get("revisionId")) {
        Some(Any::String(id)) => Some(id.to_string()),
        Some(Any::Number(id)) if id.is_finite() => Some(id.to_string()),
        Some(Any::BigInt(id)) => Some(id.to_string()),
        _ => None,
    }
}

fn apply_para_delta(txn: &mut TransactionMut<'_>, map: &MapRef, delta: &ParaAttrDelta) {
    fn apply<T, F: Fn(&T) -> Any>(
        txn: &mut TransactionMut<'_>,
        map: &MapRef,
        key: &str,
        patch: &Patch<T>,
        lower: F,
    ) {
        match patch {
            Patch::Keep => {}
            Patch::Clear => {
                map.remove(txn, key);
            }
            Patch::Set(value) => {
                map.insert(txn, key.to_owned(), lower(value));
            }
        }
    }
    apply(txn, map, "alignment", &delta.alignment, |v| {
        Any::from(v.as_str())
    });
    apply(txn, map, "lineSpacing", &delta.line_spacing, |v| {
        Any::Number(*v)
    });
    apply(txn, map, "lineSpacingRule", &delta.line_spacing_rule, |v| {
        Any::from(v.as_str())
    });
    apply(txn, map, "spaceBefore", &delta.space_before, |v| {
        Any::Number(*v)
    });
    apply(txn, map, "spaceAfter", &delta.space_after, |v| {
        Any::Number(*v)
    });
    apply(txn, map, INDENT_LEFT, &delta.indent_left, |v| {
        Any::Number(*v)
    });
    apply(txn, map, "indentRight", &delta.indent_right, |v| {
        Any::Number(*v)
    });
    apply(txn, map, "indentFirstLine", &delta.indent_first_line, |v| {
        Any::Number(*v)
    });
    apply(txn, map, "hangingIndent", &delta.hanging_indent, |v| {
        Any::Number(*v)
    });
    apply(txn, map, "bidi", &delta.bidi, |v| Any::Bool(*v));
    apply(txn, map, TABS, &delta.tabs, |stops| {
        Any::Array(Arc::from(
            stops.iter().map(TabStop::to_any).collect::<Vec<_>>(),
        ))
    });
    apply(
        txn,
        map,
        DEFAULT_TEXT_FORMATTING,
        &delta.default_text_formatting,
        |dtf| {
            let value: HashMap<String, Any> =
                dtf.iter().map(|(k, v)| (k.clone(), v.clone())).collect();
            Any::Map(Arc::new(value))
        },
    );
    for (key, value) in &delta.other {
        set_or_remove(txn, map, key, value.clone());
    }
}

fn paragraph_formatting<T: ReadTxn>(map: &MapRef, txn: &T) -> HashMap<String, Any> {
    map.iter(txn)
        .filter_map(|(key, value)| {
            if matches!(
                key.as_ref(),
                KIND_KEY | PARA_ID | PPR_INS | PPR_DEL | PPR_CHANGE
            ) {
                return None;
            }
            match value {
                Out::Any(value) if value != Any::Null => Some((key.to_string(), value)),
                _ => None,
            }
        })
        .collect()
}

fn append_paragraph_property_change(
    txn: &mut TransactionMut<'_>,
    map: &MapRef,
    revision_id: &str,
    author: &crate::Author,
    previous: HashMap<String, Any>,
    current: HashMap<String, Any>,
) {
    let mut changes = match map.get(txn, PPR_CHANGE) {
        Some(Out::Any(Any::Array(changes))) => changes.to_vec(),
        _ => Vec::new(),
    };
    let info = revision_value(revision_id, author);
    changes.push(Any::Map(Arc::new(HashMap::from([
        ("type".to_owned(), Any::from("paragraphPropertyChange")),
        ("info".to_owned(), info),
        (
            "previousFormatting".to_owned(),
            Any::Map(Arc::new(previous)),
        ),
        ("currentFormatting".to_owned(), Any::Map(Arc::new(current))),
    ]))));
    map.insert(txn, PPR_CHANGE, Any::Array(Arc::from(changes)));
}

fn read_tab_stops<T: ReadTxn>(map: &MapRef, txn: &T) -> Vec<Any> {
    match map.get(txn, TABS) {
        Some(Out::Any(Any::Array(stops))) => stops.to_vec(),
        _ => Vec::new(),
    }
}

fn existing_pos(stop: &Any) -> Option<f64> {
    let Any::Map(map) = stop else {
        return None;
    };
    match map.get("pos") {
        Some(Any::Number(pos)) => Some(*pos),
        Some(Any::BigInt(pos)) => Some(*pos as f64),
        _ => None,
    }
}

fn number_prop<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<f64> {
    match map.get(txn, key) {
        Some(Out::Any(Any::Number(value))) => Some(value),
        Some(Out::Any(Any::BigInt(value))) => Some(value as f64),
        _ => None,
    }
}
