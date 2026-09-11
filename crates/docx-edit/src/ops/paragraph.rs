//! Paragraph structure and properties: split, merge, attribute deltas, tab
//! stops, indent steps, style application, and paraId maintenance.
//!
//! A paragraph is not a container here. Its text is a plain run of story
//! units, and its identity and properties live on the pilcrow embed that
//! TERMINATES it — so splitting a paragraph is inserting one pilcrow, merging
//! two is removing one, and every property op writes to a pilcrow's map.
//!
//! Two consequences follow. First, splitting has to decide which half keeps
//! the original paraId: the new pilcrow takes it and terminates the FIRST
//! half, while the original pilcrow — now ending the second half — is
//! re-minted. Second, merging has to decide whose properties survive; a plain
//! merge lets the deleted mark's properties and paraId win, so the earlier
//! paragraph's identity carries over. Tracked-change resolution uses the
//! opposite rule, for the reasons its own module docs give.
//!
//! Style resolution stays outside the CRDT. Ops that apply a style take a
//! host-supplied [`ResolvedStyleProjection`] rather than reading `styles.xml`,
//! and reject an unknown style before touching the document.
//!
//! Spacing, indent and tab values are authored OOXML units — twips and
//! line-spacing units — never pixels.

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

/// The paragraph attributes a style definition owns. Applying a style resets
/// every one of them to the style's value, or clears it when the style has
/// none — an attribute here is never left over from the previous style.
pub const STYLE_CONTROLLED_PARA_ATTRS: [&str; 16] = [
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
    "widowControl",
    "pageBreakBefore",
    "outlineLevel",
    "defaultTextFormatting",
];

/// The run marks swept from a paragraph's text before the new style's own run
/// formatting is written, so direct formatting from the old style cannot
/// survive the switch.
pub const STYLE_CONTROLLED_MARKS: [&str; 7] = [
    "bold",
    "italic",
    "fontSize",
    "fontFamily",
    "textColor",
    "underline",
    "strike",
];

/// The only paragraph properties an EMPTY second half inherits on split —
/// pressing Enter at the end of a paragraph starts a clean one that keeps the
/// style and vertical rhythm but nothing else.
const INHERITED_PARA_ATTRS: [&str; 7] = [
    "defaultTextFormatting",
    "pStyle",
    "lineSpacing",
    "lineSpacingRule",
    "spaceAfter",
    "spaceBefore",
    "contextualSpacing",
];

/// The `defaultTextFormatting` keys that cross a split: font, size and color
/// only. Bold, italic, underline and the rest deliberately do not carry.
const STYLE_CARRY_DTF_KEYS: [&str; 4] = ["fontFamily", "fontSize", "fontSizeCs", "color"];

const BORDERS: &str = "borders";
const TABS: &str = "tabs";
const INDENT_LEFT: &str = "indentLeft";
const DEFAULT_TEXT_FORMATTING: &str = "defaultTextFormatting";

/// Default half-inch indent step in twips.
pub const INDENT_STEP_TWIPS: f64 = 720.0;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum MergeDirection {
    /// Delete THIS paragraph's mark, merging it with the following paragraph.
    Forward,
    /// Delete the PREVIOUS paragraph's mark, merging this paragraph into it.
    Backward,
}

/// Which paragraphs an op targets. Every variant is resolved and validated
/// before any mutation, so an unknown id never leaves a partial edit.
#[derive(Clone, Debug, PartialEq)]
pub enum ParaSelector {
    One(ParagraphId),
    /// Resolved in the order given; an unknown id in the list is an error.
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

/// A tri-state paragraph-property delta: [`Patch::Keep`] leaves the current
/// value, [`Patch::Clear`] removes the property, [`Patch::Set`] writes it.
/// Spacing and indent values are authored OOXML units, never pixels.
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
    /// Every paragraph property without a typed field above, written as given;
    /// `None` clears the key. Schema-managed identity keys are rejected.
    pub other: BTreeMap<String, Option<Any>>,
}

/// A style definition already resolved by the host, injected because the
/// `styles.xml` cascade lives outside the CRDT.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct ResolvedStyleProjection {
    pub style_id: String,
    /// Host-verified existence. `false` yields [`OpError::UnknownStyle`]
    /// before any mutation.
    pub known: bool,
    /// Values for the [`STYLE_CONTROLLED_PARA_ATTRS`] keys — a missing or null
    /// entry clears that attribute — plus any list attributes when the style
    /// defines numbering.
    pub paragraph_attrs: BTreeMap<String, Any>,
    /// The style's run formatting as story attribute values (`bold`,
    /// `fontSize`, …), applied after the [`STYLE_CONTROLLED_MARKS`] sweep.
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

/// Writes a style's paragraph-attribute projection: each
/// [`STYLE_CONTROLLED_PARA_ATTRS`] key is reset to the projection's value or
/// cleared when it has none, and any extra key (list attributes) is applied as
/// given. Errors on a schema-managed identity key.
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
    /// Splits a paragraph by inserting exactly ONE pilcrow at `at`.
    ///
    /// The new pilcrow terminates the FIRST half, carrying the source
    /// paragraph's full properties and its ORIGINAL paraId; the original
    /// pilcrow is re-minted with a fresh paraId and becomes the second half's
    /// mark. What the second half then keeps depends on where the split fell:
    ///
    /// - mid-paragraph: it keeps its own properties;
    /// - at the paragraph end, so the second half is empty: only
    ///   the `INHERITED_PARA_ATTRS` subset survives, with
    ///   `defaultTextFormatting` reduced to the font/size/color carry keys;
    /// - at the end WITH a `next_style`: it switches to that style's
    ///   projection outright instead.
    ///
    /// Paragraph borders are cleared in every case, because Word never
    /// propagates `w:pBdr` across a split. Suggesting mode stamps the inserted
    /// pilcrow `ins` and `pPrIns`, reusing an adjacent revision by the same
    /// author when there is one.
    ///
    /// Errors when `next_style` is not known, before any mutation, and when
    /// `at` does not address a position inside a story.
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
                // A `w:next` switch starts from nothing: drop the source
                // properties before writing the projection.
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

    /// Merges the paragraph with its neighbour by deleting the boundary
    /// pilcrow. The survivor adopts the deleted mark's properties and paraId,
    /// so the EARLIER paragraph's identity wins.
    ///
    /// Suggesting mode normally retains the mark, stamping it `del` and
    /// `pPrDel`. The exception is backspacing over this same author's still
    /// pending split: that retracts the suggestion, physically removing the
    /// pilcrow rather than authoring a second, contradictory revision.
    ///
    /// The receipt's range is the caret position after the merge. Errors when
    /// the paragraph is unknown, when merging forward from a story's last
    /// paragraph, and when merging backward from its first.
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

    /// Applies a tri-state paragraph attribute delta to the selected
    /// paragraphs in one transaction. In suggesting mode a `pPrChange` record
    /// capturing the before and after property maps is appended to every
    /// paragraph the delta actually changed, and the receipt carries its
    /// revision id; a delta that changes nothing stamps nothing. Errors when
    /// the delta names a schema-managed identity key, before any mutation.
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

    /// Adds a tab stop to the selected paragraphs, replacing any existing stop
    /// at the same position. Stops stay sorted by position.
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

    /// Removes the tab stop at `pos` twips from the selected paragraphs,
    /// dropping the `tabs` attribute entirely once none are left.
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

    /// Increases `indentLeft` by `step` twips, defaulting to
    /// [`INDENT_STEP_TWIPS`].
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

    /// Decreases `indentLeft` by `step` twips, defaulting to
    /// [`INDENT_STEP_TWIPS`] and clamping at zero. Reaching zero removes the
    /// attribute rather than storing it.
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

    /// Applies a host-resolved paragraph style in ONE transaction: writes the
    /// style id, resets every [`STYLE_CONTROLLED_PARA_ATTRS`] entry to the
    /// projection's value or clears it, sweeps the [`STYLE_CONTROLLED_MARKS`]
    /// from the paragraph's text, and writes the style's run formats over it.
    ///
    /// Errors before any mutation when the style is not known, and when the
    /// projection's run marks include a protected attribute such as a
    /// hyperlink or a tracked-change stamp.
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

    /// Restores paraId uniqueness after a merge of divergent replicas: every
    /// duplicate is re-minted, with the first occurrence in document order
    /// keeping its id. Runs under a system origin so the pass never enters
    /// undo history. Returns the `(old, new)` pairs.
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

/// Appends one `paragraphPropertyChange` record to the pilcrow's `pPrChange`
/// array, holding the revision info plus the complete property maps from
/// before and after — what rejection needs to rewind the paragraph.
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
