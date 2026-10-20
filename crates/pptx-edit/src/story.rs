use std::sync::Arc;

use ooxml_drawingml::{Theme, resolve_color_value_to_hex_with_theme};
use pptx_parse::{RunProperties, TextBody};
use yrs::branch::{Branch, BranchPtr};
use yrs::types::Attrs;
use yrs::types::text::YChange;
use yrs::updates::{decoder::Decode, encoder::Encode};
use yrs::{
    Any, Assoc, IndexedSequence, Map, MapPrelim, MapRef, Out, ReadTxn, StickyIndex, Text,
    TextPrelim, TextRef, Transact, TransactionMut,
};

use crate::model::validate_xml_text;
use crate::{
    CaretAnchor, DeckSession, EditError, EditResult, KIND, PARA_ID, PILCROW_KIND,
    ParagraphSnapshot, ParagraphSource, STORIES, StorySnapshot, TextCaps, TextReceipt,
    TextRunSnapshot, TextStyle, TextStylePatch,
};

const UNDERLINE_TYPES: [&str; 18] = [
    "none",
    "words",
    "sng",
    "dbl",
    "heavy",
    "dotted",
    "dottedHeavy",
    "dash",
    "dashHeavy",
    "dashLong",
    "dashLongHeavy",
    "dotDash",
    "dotDashHeavy",
    "dotDotDash",
    "dotDotDashHeavy",
    "wavy",
    "wavyHeavy",
    "wavyDbl",
];

const ALIGNMENTS: [&str; 7] = ["l", "ctr", "r", "just", "justLow", "dist", "thaiDist"];

/// The values land in schema-typed attributes, so junk must fail the edit
/// rather than the file.
pub(crate) fn validate_style_values(
    font_family: Option<&str>,
    underline: Option<&str>,
    color: Option<&str>,
    font_size_pt: Option<f64>,
    spacing_pt: Option<f64>,
    baseline_pct: Option<f64>,
    kern_pt: Option<f64>,
) -> EditResult<()> {
    if let Some(font_family) = font_family {
        validate_xml_text(font_family)?;
    }
    if let Some(underline) = underline
        && !UNDERLINE_TYPES.contains(&underline)
    {
        return Err(EditError::InvalidText(format!(
            "unrecognized underline type {underline:?}"
        )));
    }
    if let Some(color) = color {
        let rgb = color.strip_prefix('#').unwrap_or(color);
        if rgb.len() != 6 || !rgb.bytes().all(|byte| byte.is_ascii_hexdigit()) {
            return Err(EditError::InvalidText(format!(
                "color {color:?} must be a six-digit hex value"
            )));
        }
    }
    if let Some(size) = font_size_pt
        && (!size.is_finite() || !(1.0..=4_000.0).contains(&size))
    {
        return Err(EditError::InvalidText(format!(
            "font size {size}pt is outside the 1-4000pt range"
        )));
    }
    if let Some(spacing) = spacing_pt
        && (!spacing.is_finite() || !(-4_000.0..=4_000.0).contains(&spacing))
    {
        return Err(EditError::InvalidText(format!(
            "letter spacing {spacing}pt is outside the -4000-4000pt range"
        )));
    }
    if let Some(kern) = kern_pt
        && (!kern.is_finite() || !(0.0..=4_000.0).contains(&kern))
    {
        return Err(EditError::InvalidText(format!(
            "kerning threshold {kern}pt is outside the 0-4000pt range"
        )));
    }
    if let Some(baseline) = baseline_pct
        && (!baseline.is_finite()
            || !(f64::from(i32::MIN)..=f64::from(i32::MAX)).contains(&(baseline * 1000.0)))
    {
        return Err(EditError::InvalidText(
            "baseline exceeds the signed percentage range".into(),
        ));
    }
    Ok(())
}

pub(crate) fn validate_alignment(alignment: Option<&str>) -> EditResult<()> {
    if let Some(alignment) = alignment
        && !ALIGNMENTS.contains(&alignment)
    {
        return Err(EditError::InvalidText(format!(
            "unrecognized paragraph alignment {alignment:?}"
        )));
    }
    Ok(())
}

pub(crate) fn seed_story(
    stories: &MapRef,
    txn: &mut TransactionMut<'_>,
    story_id: &str,
    body: &TextBody,
    theme: Option<&Theme>,
) -> EditResult<()> {
    let story = stories.insert(txn, story_id, TextPrelim::new(""));
    if body.paragraphs.is_empty() {
        append_pilcrow(&story, txn, &format!("para:{story_id}:0"), None, 0, None);
        return Ok(());
    }
    for (paragraph_index, paragraph) in body.paragraphs.iter().enumerate() {
        for run in &paragraph.runs {
            if !run.text.is_empty() {
                let style = style_from_run_properties(&run.properties, theme);
                let index = story.len(txn);
                insert_styled_text(&story, txn, index, &run.text, &style);
            }
        }
        let bullet_json = paragraph
            .properties
            .bullet
            .as_ref()
            .map(serde_json::to_string)
            .transpose()
            .map_err(|error| EditError::Json(error.to_string()))?;
        append_pilcrow(
            &story,
            txn,
            &format!("para:{story_id}:{paragraph_index}"),
            paragraph.properties.alignment.as_deref(),
            paragraph.properties.level,
            bullet_json.as_deref(),
        );
    }
    Ok(())
}

pub(crate) fn seed_plain_story(
    stories: &MapRef,
    txn: &mut TransactionMut<'_>,
    story_id: &str,
    paragraph_id: &str,
    text: &str,
    style: &TextStyle,
) -> TextRef {
    let story = stories.insert(txn, story_id, TextPrelim::new(""));
    if !text.is_empty() {
        insert_styled_text(&story, txn, 0, text, style);
    }
    append_pilcrow(&story, txn, paragraph_id, None, 0, None);
    story
}

pub(crate) fn import_source_numbering_restarts(
    doc: &yrs::Doc,
    package: &pptx_parse::PptxPackage,
) -> EditResult<()> {
    let source = crate::doc_with_client_id(crate::BOOTSTRAP_CLIENT_ID);
    crate::deck::seed_doc(&source, package, "")?;
    let source_txn = source.transact();
    let source_stories = crate::deck::required_map(&source_txn, STORIES)?;
    let mut restarts = std::collections::HashMap::new();
    for (_, value) in source_stories.iter(&source_txn) {
        let Out::YText(story) = value else { continue };
        for diff in story.diff(&source_txn, YChange::identity) {
            let Out::YMap(map) = diff.insert else {
                continue;
            };
            let (Some(id), Some(json)) = (
                map_string(&map, &source_txn, PARA_ID),
                map_string(&map, &source_txn, "bulletJson"),
            ) else {
                continue;
            };
            let mut bullet: pptx_parse::Bullet =
                serde_json::from_str(&json).map_err(|error| EditError::Json(error.to_string()))?;
            if let pptx_parse::Bullet::AutoNumber {
                restart: restart @ true,
                ..
            } = &mut bullet
            {
                *restart = false;
                restarts.insert(id, (bullet, json));
            }
        }
    }
    if restarts.is_empty() {
        return Ok(());
    }
    let mut txn = doc.transact_mut_with(crate::MIGRATE_ORIGIN);
    let stories = crate::deck::required_map(&txn, STORIES)?;
    let mut updates = Vec::new();
    for (story_id, value) in stories.iter(&txn) {
        let Out::YText(story) = value else { continue };
        for diff in story.diff(&txn, YChange::identity) {
            let Out::YMap(map) = diff.insert else {
                continue;
            };
            let Some(id) = map_string(&map, &txn, PARA_ID) else {
                continue;
            };
            let Some((legacy, source)) = restarts.get(&id) else {
                continue;
            };
            let current = map_string(&map, &txn, "bulletJson")
                .and_then(|json| serde_json::from_str::<pptx_parse::Bullet>(&json).ok());
            if current.as_ref() == Some(legacy) {
                updates.push(((story_id.to_owned(), id), map, source.clone()));
            }
        }
    }
    updates.sort_by(|left, right| left.0.cmp(&right.0));
    for (_, map, json) in updates {
        map.insert(&mut txn, "bulletJson", json);
    }
    Ok(())
}

fn append_pilcrow(
    story: &TextRef,
    txn: &mut TransactionMut<'_>,
    paragraph_id: &str,
    alignment: Option<&str>,
    level: u32,
    bullet_json: Option<&str>,
) {
    let index = story.len(txn);
    let pilcrow =
        story.insert_embed_with_attributes(txn, index, MapPrelim::default(), Attrs::default());
    pilcrow.insert(txn, KIND, PILCROW_KIND);
    pilcrow.insert(txn, PARA_ID, paragraph_id);
    pilcrow.insert(txn, "level", level as f64);
    if let Some(alignment) = alignment {
        pilcrow.insert(txn, "alignment", alignment);
    }
    if let Some(bullet_json) = bullet_json {
        pilcrow.insert(txn, "bulletJson", bullet_json);
    }
}

fn replacement_bullet(enabled: bool) -> EditResult<String> {
    let bullet = if enabled {
        pptx_parse::Bullet::Character {
            value: "•".into()
        }
    } else {
        pptx_parse::Bullet::None
    };
    serde_json::to_string(&bullet).map_err(|error| EditError::Json(error.to_string()))
}

fn styled_characters<'a>(
    runs: impl Iterator<Item = (&'a str, &'a TextStyle)>,
) -> Vec<(char, &'a TextStyle)> {
    runs.flat_map(|(text, style)| text.chars().map(move |character| (character, style)))
        .collect()
}

fn paragraph_offsets(story: &StorySnapshot) -> Vec<u32> {
    story
        .paragraphs
        .iter()
        .scan(0, |offset, paragraph| {
            let start = *offset;
            *offset += paragraph
                .runs
                .iter()
                .map(|run| run.text.encode_utf16().count() as u32)
                .sum::<u32>()
                + 1;
            Some(start)
        })
        .collect()
}

impl DeckSession {
    pub fn story(&self, story_id: &str) -> EditResult<StorySnapshot> {
        let txn = self.doc.transact();
        let story = story_ref(&txn, story_id)?;
        snapshot_story(&story, &txn, story_id)
    }

    pub fn anchor_caret(&self, story_id: &str, index: u32) -> EditResult<CaretAnchor> {
        let txn = self.doc.transact();
        let story = story_ref(&txn, story_id)?;
        let length = final_pilcrow_index(&story, &txn)?;
        if index > length {
            return Err(EditError::OutOfBounds { index, length });
        }
        let position = if index == 0 {
            StickyIndex::from_type(&txn, &story, Assoc::Before)
        } else {
            story
                .sticky_index(&txn, index, Assoc::After)
                .ok_or(EditError::OutOfBounds { index, length })?
        };
        Ok(CaretAnchor {
            story_id: story_id.to_owned(),
            position,
        })
    }

    pub fn resolve_caret_anchor(&self, anchor: &CaretAnchor) -> Option<u32> {
        let txn = self.doc.transact();
        let story = story_ref(&txn, &anchor.story_id).ok()?;
        let offset = anchor.position.get_offset(&txn)?;
        let expected = BranchPtr::from(<TextRef as AsRef<Branch>>::as_ref(&story));
        if offset.branch != expected {
            return None;
        }
        let length = final_pilcrow_index(&story, &txn).ok()?;
        Some(offset.index.min(length))
    }

    pub fn set_story_paragraphs(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        paragraphs: &[crate::TextParagraphDraft],
    ) -> EditResult<StorySnapshot> {
        if paragraphs.is_empty() || paragraphs.len() > 100 {
            return Err(EditError::InvalidText("supply 1 to 100 paragraphs".into()));
        }
        let mut length = 0;
        for paragraph in paragraphs {
            if paragraph.runs.len() > 64 {
                return Err(EditError::InvalidText(
                    "supply up to 64 runs per paragraph".into(),
                ));
            }
            if paragraph
                .alignment
                .as_deref()
                .is_some_and(|alignment| !ALIGNMENTS.contains(&alignment))
            {
                return Err(EditError::InvalidText(
                    "unrecognized paragraph alignment".into(),
                ));
            }
            for run in &paragraph.runs {
                validate_xml_text(&run.text)?;
                if run.text.contains(['\r', '\n']) {
                    return Err(EditError::InvalidText(
                        "each run must stay in one paragraph".into(),
                    ));
                }
                let style = &run.style;
                validate_style_values(
                    style.font_family.as_deref(),
                    style.underline.as_deref(),
                    style.color.as_deref(),
                    style.font_size_pt,
                    style.spacing_pt,
                    style.baseline_pct,
                    style.kern_pt,
                )?;
                length += run.text.encode_utf16().count();
            }
        }
        if length > 16_000 {
            return Err(EditError::InvalidText(
                "text exceeds the 16000-character limit".into(),
            ));
        }
        let previous = self.story(story_id)?;
        let old_count = previous.paragraphs.len();
        let new_count = paragraphs.len();
        let old_text: Vec<_> = previous
            .paragraphs
            .iter()
            .map(|paragraph| {
                styled_characters(
                    paragraph
                        .runs
                        .iter()
                        .map(|run| (run.text.as_str(), &run.style)),
                )
            })
            .collect();
        let draft_text: Vec<_> = paragraphs
            .iter()
            .map(|paragraph| {
                styled_characters(
                    paragraph
                        .runs
                        .iter()
                        .map(|run| (run.text.as_str(), &run.style)),
                )
            })
            .collect();
        let bullets = paragraphs
            .iter()
            .map(|paragraph| paragraph.bullet.map(replacement_bullet).transpose())
            .collect::<EditResult<Vec<_>>>()?;
        let equal = |old: usize, draft: usize| {
            old_text[old] == draft_text[draft]
                && paragraphs[draft]
                    .alignment
                    .as_ref()
                    .is_none_or(|alignment| {
                        previous.paragraphs[old].alignment.as_ref() == Some(alignment)
                    })
                && bullets[draft].as_ref().is_none_or(|bullet| {
                    previous.paragraphs[old].bullet_json.as_ref() == Some(bullet)
                })
        };
        if old_count == new_count && (0..old_count).all(|index| equal(index, index)) {
            return self.story(story_id);
        }
        let shared = old_count.min(new_count);
        let prefix = (0..shared - 1)
            .take_while(|&index| equal(index, index))
            .count();
        let suffix = (0..shared - prefix)
            .take_while(|&index| equal(old_count - 1 - index, new_count - 1 - index))
            .count();
        let old_middle = old_count - prefix - suffix;
        let new_middle = new_count - prefix - suffix;
        let old_starts = paragraph_offsets(&previous);
        if old_middle > new_middle {
            let fields =
                crate::target::field_paragraphs(&self.package, &self.snapshot()?, story_id);
            let txn = self.doc.transact();
            let story = story_ref(&txn, story_id)?;
            let properties = paragraph_projections(&story, &txn, story_id);
            let first = prefix + new_middle - usize::from(suffix == 0);
            let end = old_count - suffix + usize::from(suffix > 0);
            if properties[first..end].iter().any(|(_, map, _)| {
                map_string(map, &txn, PARA_ID).is_some_and(|id| fields.contains(&id))
                    || matches!(map.get(&txn, "hasFields"), Some(Out::Any(Any::Bool(true))))
            }) {
                let start = old_starts[prefix + new_middle] - u32::from(suffix == 0);
                return Err(EditError::ParagraphBoundary {
                    start,
                    end: start + 1,
                });
            }
        }
        let mut templates = vec![None; new_count];
        if new_middle > old_middle {
            let (range, template) = if suffix > 0 {
                let template = if prefix + old_middle > 0 {
                    &previous.paragraphs[prefix + old_middle - 1]
                } else {
                    &previous.paragraphs[old_count - suffix]
                };
                (prefix + old_middle..prefix + new_middle, template)
            } else {
                (
                    old_count - 1..new_count - 1,
                    &previous.paragraphs[old_count - 1],
                )
            };
            for slot in &mut templates[range] {
                *slot = Some(template.clone());
            }
        } else if old_middle > new_middle && suffix == 0 {
            templates[new_count - 1] = Some(previous.paragraphs[new_count - 1].clone());
        }
        if templates.iter().any(Option::is_some) {
            let source_xml = self.source_paragraph_xml(story_id)?;
            let source_prefix = format!("para:{story_id}:");
            for template in templates.iter_mut().flatten() {
                if template.preserved_xml.is_none() {
                    template.preserved_xml = template
                        .id
                        .strip_prefix(&source_prefix)
                        .and_then(|index| index.parse::<usize>().ok())
                        .and_then(|index| source_xml.get(index).cloned());
                }
            }
        }
        self.with_undo_capture(|| {
            if new_middle > old_middle {
                let index = if suffix > 0 {
                    old_starts[old_count - suffix]
                } else {
                    previous.length - 1
                };
                for _ in 0..new_middle - old_middle {
                    self.insert_paragraph_break(context, story_id, index)?;
                }
            } else if old_middle > new_middle {
                for index in (prefix + new_middle..old_count - suffix).rev() {
                    let start = old_starts[index];
                    let end = old_starts
                        .get(index + 1)
                        .copied()
                        .unwrap_or(previous.length)
                        - 1;
                    if start != end {
                        self.delete_text(context, story_id, start, end)?;
                    }
                    let marker = if suffix > 0 { start } else { start - 1 };
                    self.remove_replacement_paragraph(context, story_id, marker, suffix == 0)?;
                }
            }
            let current = self.story(story_id)?;
            if current.paragraphs.len() != new_count {
                return Err(EditError::InvalidState(
                    "story replacement produced an unexpected paragraph count".into(),
                ));
            }
            let starts = paragraph_offsets(&current);
            for (index, (paragraph, draft)) in
                current.paragraphs.iter().zip(&draft_text).enumerate().rev()
            {
                let old = styled_characters(
                    paragraph
                        .runs
                        .iter()
                        .map(|run| (run.text.as_str(), &run.style)),
                );
                if old == *draft {
                    continue;
                }
                let text_prefix = old.iter().zip(draft).take_while(|(a, b)| a == b).count();
                let text_suffix = old
                    .iter()
                    .rev()
                    .zip(draft.iter().rev())
                    .take(old.len().min(draft.len()) - text_prefix)
                    .take_while(|(a, b)| a == b)
                    .count();
                let mut offset = starts[index]
                    + old[..text_prefix]
                        .iter()
                        .map(|(character, _)| character.len_utf16() as u32)
                        .sum::<u32>();
                let removed = old[text_prefix..old.len() - text_suffix]
                    .iter()
                    .map(|(character, _)| character.len_utf16() as u32)
                    .sum::<u32>();
                if removed > 0 {
                    self.delete_text(context, story_id, offset, offset + removed)?;
                }
                let middle = &draft[text_prefix..draft.len() - text_suffix];
                let mut cursor = 0;
                while cursor < middle.len() {
                    let style = middle[cursor].1;
                    let count = middle[cursor..]
                        .iter()
                        .take_while(|(_, candidate)| *candidate == style)
                        .count();
                    let text: String = middle[cursor..cursor + count]
                        .iter()
                        .map(|(character, _)| *character)
                        .collect();
                    self.insert_text(context, story_id, offset, &text, style)?;
                    offset += text.encode_utf16().count() as u32;
                    cursor += count;
                }
            }
            let current = self.story(story_id)?;
            let starts = paragraph_offsets(&current);
            for (index, template) in templates.iter().enumerate() {
                let Some(template) = template else { continue };
                self.set_replacement_properties(context, story_id, starts[index], template)?;
                if current.paragraphs[index].alignment != template.alignment {
                    self.set_paragraph_alignment(
                        context,
                        story_id,
                        starts[index],
                        starts[index],
                        template.alignment.as_deref(),
                    )?;
                }
            }
            let current = self.story(story_id)?;
            for (index, paragraph) in paragraphs.iter().enumerate() {
                if let Some(alignment) = paragraph.alignment.as_deref()
                    && current.paragraphs[index].alignment.as_deref() != Some(alignment)
                {
                    self.set_paragraph_alignment(
                        context,
                        story_id,
                        starts[index],
                        starts[index],
                        Some(alignment),
                    )?;
                }
                if let Some(bullet) = &bullets[index]
                    && current.paragraphs[index].bullet_json.as_ref() != Some(bullet)
                {
                    let mut properties = current.paragraphs[index].clone();
                    properties.bullet_json = Some(bullet.clone());
                    self.set_replacement_properties(context, story_id, starts[index], &properties)?;
                }
            }
            self.story(story_id)
        })
    }

    fn set_replacement_properties(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        paragraph_start: u32,
        properties: &ParagraphSnapshot,
    ) -> EditResult<()> {
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        let marker = selected_pilcrows(&story, &txn, story_id, paragraph_start, paragraph_start)
            .pop()
            .ok_or_else(|| EditError::InvalidState("missing replacement paragraph".into()))?;
        let level = f64::from(properties.level);
        if map_number(&marker, &txn, "level") != Some(level) {
            marker.insert(&mut txn, "level", level);
        }
        for (key, value) in [
            ("bulletJson", properties.bullet_json.as_deref()),
            ("preservedXml", properties.preserved_xml.as_deref()),
        ] {
            if map_string(&marker, &txn, key).as_deref() != value {
                if let Some(value) = value {
                    marker.insert(&mut txn, key, value);
                } else {
                    marker.remove(&mut txn, key);
                }
            }
        }
        if properties.source_paragraphs.is_empty() {
            marker.remove(&mut txn, "sourceParagraphs");
        }
        Ok(())
    }

    fn source_paragraph_xml(&self, story_id: &str) -> EditResult<Vec<String>> {
        fn find_shape<'a>(
            shapes: &'a [crate::ShapeSnapshot],
            story_id: &str,
        ) -> Option<&'a crate::ShapeSnapshot> {
            shapes.iter().find_map(|shape| {
                if shape.text_stories.iter().any(|story| story.id == story_id) {
                    Some(shape)
                } else {
                    find_shape(&shape.children, story_id)
                }
            })
        }
        for slide in self.snapshot()?.slides {
            let Some(shape) = find_shape(&slide.shapes, story_id) else {
                continue;
            };
            let Some(part) = slide.source_part_path else {
                return Ok(Vec::new());
            };
            let Some((_, path)) = shape.id.rsplit_once(":shape:") else {
                return Ok(Vec::new());
            };
            let path = path
                .split('.')
                .map(str::parse)
                .collect::<Result<Vec<usize>, _>>()
                .map_err(|error| EditError::Write(error.to_string()))?;
            let target = crate::save::text_target(story_id, &shape.id)?;
            return pptx_parse::text_paragraph_xml(&self.package, &part, &path, &target)
                .map_err(|error| EditError::Write(error.to_string()));
        }
        Ok(Vec::new())
    }

    pub fn insert_text(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        index: u32,
        text: &str,
        style: &TextStyle,
    ) -> EditResult<TextReceipt> {
        validate_xml_text(text)?;
        validate_style_values(
            style.font_family.as_deref(),
            style.underline.as_deref(),
            style.color.as_deref(),
            style.font_size_pt,
            style.spacing_pt,
            style.baseline_pct,
            style.kern_pt,
        )?;
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        let final_pilcrow = final_pilcrow_index(&story, &txn)?;
        if index > final_pilcrow {
            return Err(EditError::OutOfBounds {
                index,
                length: final_pilcrow,
            });
        }
        if !text.is_empty() {
            insert_styled_text(&story, &mut txn, index, text, style);
        }
        let length = text.encode_utf16().count() as u32;
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start: index,
            end: index + length,
            text: text.to_owned(),
        })
    }

    pub fn delete_text(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        start: u32,
        end: u32,
    ) -> EditResult<TextReceipt> {
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        check_text_range(&story, &txn, start, end)?;
        let text = text_in_range(&story, &txn, start, end);
        if end > start {
            story.remove_range(&mut txn, start, end - start);
        }
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start,
            end,
            text,
        })
    }

    pub fn delete_text_joining_paragraphs(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        start: u32,
        end: u32,
    ) -> EditResult<TextReceipt> {
        if start == end {
            return self.delete_text(context, story_id, start, end);
        }
        let fields = crate::target::field_paragraphs(&self.package, &self.snapshot()?, story_id);
        let source_xml = self.source_paragraph_xml(story_id)?;
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        check_text_bounds(&story, &txn, start, end)?;
        let final_index = final_pilcrow_index(&story, &txn)?;
        if start <= final_index && final_index < end {
            return Err(EditError::ParagraphBoundary { start, end });
        }
        let markers = paragraph_markers(&story, &txn);
        let first = markers
            .iter()
            .position(|(index, _)| *index >= start)
            .unwrap();
        let last = markers.iter().position(|(index, _)| *index >= end).unwrap();
        let properties = paragraph_projections(&story, &txn, story_id);
        if first != last {
            if properties[first..=last].iter().any(|(_, map, _)| {
                map_string(map, &txn, PARA_ID).is_some_and(|id| fields.contains(&id))
                    || matches!(map.get(&txn, "hasFields"), Some(Out::Any(Any::Bool(true))))
            }) {
                return Err(EditError::ParagraphBoundary { start, end });
            }
            let mut sources = Vec::new();
            for (_, map, joined_sources) in &properties[first..=last] {
                let paragraph_sources = if joined_sources.is_empty() {
                    let id = map_string(map, &txn, PARA_ID).unwrap_or_default();
                    let xml = map_string(map, &txn, "preservedXml").or_else(|| {
                        let index = id
                            .strip_prefix(&format!("para:{story_id}:"))?
                            .parse::<usize>()
                            .ok()?;
                        source_xml.get(index).cloned()
                    });
                    vec![ParagraphSource { id, xml }]
                } else {
                    joined_sources.clone()
                };
                for source in paragraph_sources {
                    if !sources
                        .iter()
                        .any(|previous: &ParagraphSource| previous.id == source.id)
                    {
                        sources.push(source);
                    }
                }
            }
            let source_json = serde_json::to_string(&sources)
                .map_err(|error| EditError::Json(error.to_string()))?;
            let values: Vec<_> = properties[first]
                .1
                .iter(&txn)
                .filter_map(|(key, value)| match value {
                    Out::Any(value) if key != "joinPosition" => Some((key.to_owned(), value)),
                    _ => None,
                })
                .collect();
            let position = story
                .sticky_index(&txn, markers[first].0, Assoc::After)
                .unwrap();
            let supersedes = paragraph_join_records(&story, &txn, &markers)
                .into_iter()
                .filter_map(|(offset, key, _)| {
                    let index = markers.iter().position(|(index, _)| *index >= offset)?;
                    (first <= index && index <= last).then(|| Any::String(key.into()))
                })
                .collect();
            let record = markers.last().unwrap().1.insert(
                &mut txn,
                self.next_id("join"),
                MapPrelim::default(),
            );
            for (key, value) in values {
                record.insert(&mut txn, key, value);
            }
            record.insert(
                &mut txn,
                "joinPosition",
                Any::Buffer(position.encode_v1().into()),
            );
            record.insert(&mut txn, "joinSupersedes", Any::Array(supersedes));
            record.insert(&mut txn, "sourceParagraphs", source_json);
        }
        let text = text_in_range(&story, &txn, start, end);
        if end > start {
            story.remove_range(&mut txn, start, end - start);
        }
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start,
            end,
            text,
        })
    }

    pub fn format_text(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        start: u32,
        end: u32,
        patch: &TextStylePatch,
    ) -> EditResult<TextReceipt> {
        validate_style_values(
            patch.font_family.as_deref(),
            patch.underline.as_deref(),
            patch.color.as_deref(),
            patch.font_size_pt,
            patch.spacing_pt,
            patch.baseline_pct,
            None,
        )?;
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        check_text_bounds(&story, &txn, start, end)?;
        let text = text_in_range(&story, &txn, start, end);
        for (segment_start, segment_end) in paragraph_text_segments(&story, &txn, start, end) {
            for (key, value) in patch_values(patch) {
                if let Some(value) = value {
                    story.format(
                        &mut txn,
                        segment_start,
                        segment_end - segment_start,
                        Attrs::from([(Arc::from(key), value)]),
                    );
                }
            }
        }
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start,
            end,
            text,
        })
    }

    /// Sets `a:pPr@algn` on every paragraph the range touches; `None` clears
    /// the value so the placeholder cascade applies again.
    pub fn set_paragraph_alignment(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        start: u32,
        end: u32,
        alignment: Option<&str>,
    ) -> EditResult<TextReceipt> {
        validate_alignment(alignment)?;
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        check_text_bounds(&story, &txn, start, end)?;
        let text = text_in_range(&story, &txn, start, end);
        let pilcrows = selected_pilcrows(&story, &txn, story_id, start, end);
        for pilcrow in pilcrows {
            match alignment {
                Some(alignment) => {
                    pilcrow.insert(&mut txn, "alignment", alignment);
                }
                None => {
                    pilcrow.remove(&mut txn, "alignment");
                }
            }
        }
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start,
            end,
            text,
        })
    }

    pub fn insert_paragraph_break(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        index: u32,
    ) -> EditResult<TextReceipt> {
        let fields = crate::target::field_paragraphs(&self.package, &self.snapshot()?, story_id);
        let (story, has_fields, inherited) = {
            let txn = self.doc.transact();
            let story = story_ref(&txn, story_id)?;
            let final_pilcrow = final_pilcrow_index(&story, &txn)?;
            if index > final_pilcrow {
                return Err(EditError::OutOfBounds {
                    index,
                    length: final_pilcrow,
                });
            }
            let (offset, properties, sources) = paragraph_projections(&story, &txn, story_id)
                .into_iter()
                .find(|(offset, _, _)| *offset >= index)
                .unwrap();
            let has_fields = map_string(&properties, &txn, PARA_ID)
                .is_some_and(|id| fields.contains(&id))
                || matches!(
                    properties.get(&txn, "hasFields"),
                    Some(Out::Any(Any::Bool(true)))
                );
            let marker = paragraph_markers(&story, &txn)
                .into_iter()
                .find(|(position, _)| *position == offset)
                .unwrap()
                .1;
            let inherited = if marker == properties {
                None
            } else {
                let sources = serde_json::to_string(&sources)
                    .map_err(|error| EditError::Json(error.to_string()))?;
                Some((marker, properties, sources))
            };
            (story, has_fields, inherited)
        };
        let paragraph_id = self.next_id("para");
        let mut txn = self.transact_for(context);
        let pilcrow = story.insert_embed_with_attributes(
            &mut txn,
            index,
            MapPrelim::default(),
            Attrs::default(),
        );
        pilcrow.insert(&mut txn, KIND, PILCROW_KIND);
        pilcrow.insert(&mut txn, PARA_ID, paragraph_id);
        pilcrow.insert(&mut txn, "level", 0_f64);
        if has_fields {
            pilcrow.insert(&mut txn, "hasFields", true);
        }
        if let Some((marker, properties, sources)) = inherited {
            for target in [marker, pilcrow] {
                for key in ["alignment", "level", "bulletJson", "preservedXml"] {
                    match properties.get(&txn, key) {
                        Some(Out::Any(value)) => {
                            target.insert(&mut txn, key, value);
                        }
                        _ => {
                            target.remove(&mut txn, key);
                        }
                    }
                }
                target.insert(&mut txn, "sourceParagraphs", sources.as_str());
            }
        }
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start: index,
            end: index + 1,
            text: "\n".to_owned(),
        })
    }

    pub fn delete_paragraph_break(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        index: u32,
    ) -> EditResult<TextReceipt> {
        let txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        let final_pilcrow = final_pilcrow_index(&story, &txn)?;
        if index >= final_pilcrow {
            return Err(EditError::OutOfBounds {
                index,
                length: final_pilcrow,
            });
        }
        let mut offset = 0;
        let is_pilcrow = story.diff(&txn, YChange::identity).into_iter().any(|diff| {
            let length = out_len(&diff.insert);
            let found = offset == index
                && matches!(
                    diff.insert,
                    Out::YMap(ref map)
                        if map_string(map, &txn, KIND).as_deref() == Some(PILCROW_KIND)
                );
            offset += length;
            found
        });
        if !is_pilcrow {
            return Err(EditError::ParagraphBoundary {
                start: index,
                end: index.saturating_add(1),
            });
        }
        drop(txn);
        self.delete_text_joining_paragraphs(context, story_id, index, index + 1)
    }

    fn remove_replacement_paragraph(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        index: u32,
        empty_after: bool,
    ) -> EditResult<()> {
        let fields = crate::target::field_paragraphs(&self.package, &self.snapshot()?, story_id);
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        let markers = paragraph_markers(&story, &txn);
        let selected = markers
            .iter()
            .position(|(offset, _)| *offset == index)
            .ok_or(EditError::ParagraphBoundary {
                start: index,
                end: index + 1,
            })?;
        let start = selected
            .checked_sub(1)
            .map_or(0, |previous| markers[previous].0 + 1);
        let properties = paragraph_projections(&story, &txn, story_id);
        let empty = if empty_after {
            markers
                .get(selected + 1)
                .is_some_and(|(offset, _)| *offset == index + 1)
        } else {
            start == index
        };
        if !empty
            || selected + 1 >= markers.len()
            || properties[selected..=selected + 1]
                .iter()
                .any(|(_, map, _)| {
                    map_string(map, &txn, PARA_ID).is_some_and(|id| fields.contains(&id))
                        || matches!(map.get(&txn, "hasFields"), Some(Out::Any(Any::Bool(true))))
                })
        {
            return Err(EditError::ParagraphBoundary {
                start: index,
                end: index + 1,
            });
        }
        story.remove_range(&mut txn, index, 1);
        Ok(())
    }
}

pub(crate) fn validate_story<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    story_id: &str,
) -> EditResult<()> {
    let mut offset = 0;
    let mut pilcrows = Vec::new();
    for diff in story.diff(txn, YChange::identity) {
        if let Some(kern) = diff.attributes.as_ref().and_then(|attrs| attrs.get("kern"))
            && !valid_kern(kern)
        {
            return Err(EditError::InvalidState(format!(
                "story {story_id} has a kern threshold outside 0-4000pt"
            )));
        }
        let length = out_len(&diff.insert);
        if let Out::YMap(map) = diff.insert {
            if map_string(&map, txn, KIND).as_deref() != Some(PILCROW_KIND) {
                return Err(EditError::InvalidState(format!(
                    "story {story_id} contains a non-pilcrow embed"
                )));
            }
            pilcrows.push(offset);
        }
        offset += length;
    }
    if pilcrows.last().copied() != story.len(txn).checked_sub(1) {
        return Err(EditError::InvalidState(format!(
            "story {story_id} has no final pilcrow"
        )));
    }
    Ok(())
}

/// A remote `kern` lands in a schema-typed attribute, so it must be a size.
fn valid_kern(value: &Any) -> bool {
    match value {
        Any::Null => true,
        Any::Number(kern) => (0.0..=4_000.0).contains(kern),
        Any::BigInt(kern) => (0..=4_000).contains(kern),
        _ => false,
    }
}

/// The snapshot `snapshot_story` reads back from a story `seed_story` wrote,
/// computed without materializing a scratch document.
pub(crate) fn baseline_story(
    story_id: &str,
    body: &TextBody,
    theme: Option<&Theme>,
) -> EditResult<StorySnapshot> {
    let mut paragraphs = Vec::with_capacity(body.paragraphs.len().max(1));
    let mut length = 1_u32;
    if body.paragraphs.is_empty() {
        paragraphs.push(ParagraphSnapshot {
            source_paragraphs: Vec::new(),
            id: format!("para:{story_id}:0"),
            preserved_xml: None,
            alignment: None,
            level: 0,
            bullet_json: None,
            runs: Vec::new(),
        });
    } else {
        length = 0;
        for (paragraph_index, paragraph) in body.paragraphs.iter().enumerate() {
            let mut runs: Vec<TextRunSnapshot> = Vec::new();
            let mut last_attrs: Option<TextStyle> = None;
            for run in &paragraph.runs {
                if run.text.is_empty() {
                    continue;
                }
                let style = style_from_run_properties(&run.properties, theme);
                // yrs' format-gap cleanup deletes every marker between two
                // identically styled runs, so their text items squash into a
                // single run in the snapshot.
                if last_attrs.as_ref() == Some(&style) {
                    if let Some(last) = runs.last_mut() {
                        last.text.push_str(&run.text);
                    }
                } else {
                    runs.push(TextRunSnapshot {
                        text: run.text.clone(),
                        style: baseline_style(style.clone()),
                    });
                }
                last_attrs = Some(style);
                length += run.text.encode_utf16().count() as u32;
            }
            length += 1;
            let bullet_json = paragraph
                .properties
                .bullet
                .as_ref()
                .map(serde_json::to_string)
                .transpose()
                .map_err(|error| EditError::Json(error.to_string()))?;
            paragraphs.push(ParagraphSnapshot {
                source_paragraphs: Vec::new(),
                id: format!("para:{story_id}:{paragraph_index}"),
                preserved_xml: None,
                alignment: paragraph.properties.alignment.clone(),
                level: paragraph.properties.level,
                bullet_json,
                runs,
            });
        }
    }
    Ok(StorySnapshot {
        id: story_id.to_owned(),
        length,
        paragraphs,
    })
}

/// Attribute reads drop non-finite numbers, so the snapshot applies the same
/// normalization the seeded `Any::Number` attributes go through.
fn baseline_style(style: TextStyle) -> TextStyle {
    TextStyle {
        font_size_pt: style.font_size_pt.filter(|value| value.is_finite()),
        spacing_pt: style.spacing_pt.filter(|value| value.is_finite()),
        baseline_pct: style.baseline_pct.filter(|value| value.is_finite()),
        kern_pt: style.kern_pt.filter(|value| value.is_finite()),
        ..style
    }
}

pub(crate) fn snapshot_story<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    story_id: &str,
) -> EditResult<StorySnapshot> {
    validate_story(story, txn, story_id)?;
    let mut paragraphs = Vec::new();
    let mut runs = Vec::new();
    let projections = paragraph_projections(story, txn, story_id);
    for diff in story.diff(txn, YChange::identity) {
        match diff.insert {
            Out::Any(Any::String(text)) => runs.push(TextRunSnapshot {
                text: text.to_string(),
                style: style_from_attrs(diff.attributes.as_deref()),
            }),
            Out::YMap(_) => {
                let (_, map, sources) = &projections[paragraphs.len()];
                paragraphs.push(ParagraphSnapshot {
                    source_paragraphs: sources.clone(),
                    id: map_string(map, txn, PARA_ID).unwrap_or_default(),
                    preserved_xml: map_string(map, txn, "preservedXml"),
                    alignment: map_string(map, txn, "alignment"),
                    level: map_number(map, txn, "level").unwrap_or_default() as u32,
                    bullet_json: map_string(map, txn, "bulletJson"),
                    runs: std::mem::take(&mut runs),
                });
            }
            _ => {}
        }
    }
    Ok(StorySnapshot {
        id: story_id.to_owned(),
        length: story.len(txn),
        paragraphs,
    })
}

pub(crate) fn story_ref<T: ReadTxn>(txn: &T, story_id: &str) -> EditResult<TextRef> {
    txn.get_map(STORIES)
        .and_then(|stories| stories.get(txn, story_id))
        .and_then(|value| value.cast::<TextRef>().ok())
        .ok_or_else(|| EditError::StoryNotFound(story_id.to_owned()))
}

fn final_pilcrow_index<T: ReadTxn>(story: &TextRef, txn: &T) -> EditResult<u32> {
    validate_story(story, txn, "requested")?;
    Ok(story.len(txn) - 1)
}

fn check_text_range<T: ReadTxn>(story: &TextRef, txn: &T, start: u32, end: u32) -> EditResult<()> {
    check_text_bounds(story, txn, start, end)?;
    for (index, _) in paragraph_markers(story, txn) {
        if start <= index && index < end {
            return Err(EditError::ParagraphBoundary { start, end });
        }
    }
    Ok(())
}

fn check_text_bounds<T: ReadTxn>(story: &TextRef, txn: &T, start: u32, end: u32) -> EditResult<()> {
    let length = story.len(txn);
    if start > end || end > length {
        return Err(EditError::OutOfBounds {
            index: end.max(start),
            length,
        });
    }
    Ok(())
}

/// Whether a paragraph-level edit over `start..end` touches the paragraph spanning
/// `paragraph_start..=pilcrow`. A collapsed caret picks the paragraph it sits in; a range
/// stopping at a paragraph start does not.
pub(crate) fn selects_paragraph(start: u32, end: u32, paragraph_start: u32, pilcrow: u32) -> bool {
    start <= pilcrow && (end > paragraph_start || (start == end && start >= paragraph_start))
}

/// The pilcrow of every paragraph the range touches, as [`selects_paragraph`] decides.
fn selected_pilcrows<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    story_id: &str,
    start: u32,
    end: u32,
) -> Vec<MapRef> {
    let start = start.min(story.len(txn).saturating_sub(1));
    let mut pilcrows = Vec::new();
    let mut paragraph_start = 0;
    for (offset, map, _) in paragraph_projections(story, txn, story_id) {
        if selects_paragraph(start, end, paragraph_start, offset) {
            pilcrows.push(map);
        }
        paragraph_start = offset + 1;
    }
    pilcrows
}

fn paragraph_markers<T: ReadTxn>(story: &TextRef, txn: &T) -> Vec<(u32, MapRef)> {
    let mut markers = Vec::new();
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let length = out_len(&diff.insert);
        if let Out::YMap(map) = diff.insert {
            markers.push((offset, map));
        }
        offset += length;
    }
    markers
}

fn source_paragraphs<T: ReadTxn>(map: &MapRef, txn: &T) -> Vec<ParagraphSource> {
    map_string(map, txn, "sourceParagraphs")
        .and_then(|json| serde_json::from_str(&json).ok())
        .unwrap_or_default()
}

fn paragraph_projections<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    story_id: &str,
) -> Vec<(u32, MapRef, Vec<ParagraphSource>)> {
    let markers = paragraph_markers(story, txn);
    let mut result: Vec<_> = markers
        .iter()
        .map(|(offset, map)| {
            let sources = if map.contains_key(txn, "sourceParagraphs") {
                source_paragraphs(map, txn)
            } else {
                Vec::new()
            };
            (*offset, map.clone(), sources)
        })
        .collect();
    let mut joins = paragraph_join_records(story, txn, &markers);
    let superseded: std::collections::HashSet<_> = joins
        .iter()
        .flat_map(|(_, _, map)| match map.get(txn, "joinSupersedes") {
            Some(Out::Any(Any::Array(keys))) => keys
                .iter()
                .filter_map(|key| match key {
                    Any::String(key) => Some(key.to_string()),
                    _ => None,
                })
                .collect::<Vec<_>>(),
            _ => Vec::new(),
        })
        .collect();
    joins.retain(|(_, key, _)| !superseded.contains(key));
    joins.sort_by_key(|(offset, key, map)| {
        (
            *offset,
            source_paragraphs(map, txn).first().and_then(|source| {
                source
                    .id
                    .strip_prefix(&format!("para:{story_id}:"))?
                    .parse::<usize>()
                    .ok()
            }),
            key.clone(),
        )
    });
    let mut projected = vec![false; result.len()];
    for (offset, _, map) in joins {
        let Some(index) = markers.iter().position(|(index, _)| *index >= offset) else {
            continue;
        };
        if !projected[index] {
            result[index].1 = map.clone();
            projected[index] = true;
        }
        for source in source_paragraphs(&map, txn) {
            if !result[index]
                .2
                .iter()
                .any(|previous| previous.id == source.id)
            {
                result[index].2.push(source);
            }
        }
    }
    result
}

pub(crate) fn materialize_paragraph_joins(
    target: &TextRef,
    txn: &mut TransactionMut<'_>,
    source_id: &str,
) -> EditResult<()> {
    let source = story_ref(txn, source_id)?;
    let markers = paragraph_markers(&source, txn);
    if paragraph_join_records(&source, txn, &markers).is_empty() {
        return Ok(());
    }
    let projections = paragraph_projections(&source, txn, source_id);
    for ((_, target), (_, properties, sources)) in
        paragraph_markers(target, txn).into_iter().zip(projections)
    {
        let values: Vec<_> = properties
            .iter(txn)
            .filter_map(|(key, value)| match value {
                Out::Any(value) if !matches!(key, "joinPosition" | "joinSupersedes") => {
                    Some((key.to_owned(), value))
                }
                _ => None,
            })
            .collect();
        target.clear(txn);
        for (key, value) in values {
            target.insert(txn, key, value);
        }
        if !sources.is_empty() {
            let json = serde_json::to_string(&sources)
                .map_err(|error| EditError::Json(error.to_string()))?;
            target.insert(txn, "sourceParagraphs", json);
        }
    }
    Ok(())
}

fn paragraph_join_records<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    markers: &[(u32, MapRef)],
) -> Vec<(u32, String, MapRef)> {
    let Some((_, last)) = markers.last() else {
        return Vec::new();
    };
    let branch = BranchPtr::from(<TextRef as AsRef<Branch>>::as_ref(story));
    last.iter(txn)
        .filter_map(|(key, value)| {
            if !key.starts_with("join:") {
                return None;
            }
            let Out::YMap(map) = value else { return None };
            let Some(Out::Any(Any::Buffer(bytes))) = map.get(txn, "joinPosition") else {
                return None;
            };
            let position = StickyIndex::decode_v1(&bytes).ok()?.get_offset(txn)?;
            (position.branch == branch).then(|| (position.index, key.to_owned(), map))
        })
        .collect()
}

fn paragraph_text_segments<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    start: u32,
    end: u32,
) -> Vec<(u32, u32)> {
    let mut segments = Vec::new();
    let mut paragraph_start = 0;
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let item_length = out_len(&diff.insert);
        if matches!(diff.insert, Out::YMap(_)) {
            let segment_start = start.max(paragraph_start);
            let segment_end = end.min(offset);
            if segment_start < segment_end {
                segments.push((segment_start, segment_end));
            }
            paragraph_start = offset + item_length;
        }
        offset += item_length;
    }
    segments
}

fn text_in_range<T: ReadTxn>(story: &TextRef, txn: &T, start: u32, end: u32) -> String {
    let mut output = String::new();
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let length = out_len(&diff.insert);
        if let Out::Any(Any::String(ref text)) = diff.insert {
            let overlap_start = start.saturating_sub(offset).min(length);
            let overlap_end = end.saturating_sub(offset).min(length);
            if overlap_end > overlap_start {
                let utf16 = text.encode_utf16().collect::<Vec<_>>();
                output.push_str(&String::from_utf16_lossy(
                    &utf16[overlap_start as usize..overlap_end as usize],
                ));
            }
        } else if matches!(diff.insert, Out::YMap(_))
            && start <= offset
            && offset < end
            && offset < story.len(txn) - 1
        {
            output.push('\n');
        }
        offset += length;
        if offset >= end {
            break;
        }
    }
    output
}

fn insert_styled_text(
    story: &TextRef,
    txn: &mut TransactionMut<'_>,
    index: u32,
    text: &str,
    style: &TextStyle,
) {
    story.insert(txn, index, text);
    let length = text.encode_utf16().count() as u32;
    for (key, value) in style_values(style) {
        story.format(txn, index, length, Attrs::from([(Arc::from(key), value)]));
    }
}

fn style_values(style: &TextStyle) -> [(&'static str, Any); 10] {
    [
        ("bold", style.bold.map(Any::Bool).unwrap_or(Any::Null)),
        ("italic", style.italic.map(Any::Bool).unwrap_or(Any::Null)),
        (
            "fontSize",
            style.font_size_pt.map(Any::Number).unwrap_or(Any::Null),
        ),
        (
            "color",
            style.color.as_deref().map(Any::from).unwrap_or(Any::Null),
        ),
        (
            "fontFamily",
            style
                .font_family
                .as_deref()
                .map(Any::from)
                .unwrap_or(Any::Null),
        ),
        (
            "underline",
            style
                .underline
                .as_deref()
                .map(Any::from)
                .unwrap_or(Any::Null),
        ),
        (
            "spacing",
            style.spacing_pt.map(Any::Number).unwrap_or(Any::Null),
        ),
        (
            "baseline",
            style.baseline_pct.map(Any::Number).unwrap_or(Any::Null),
        ),
        ("kern", style.kern_pt.map(Any::Number).unwrap_or(Any::Null)),
        (
            "caps",
            style
                .caps
                .map(|caps| Any::from(caps.as_attribute()))
                .unwrap_or(Any::Null),
        ),
    ]
}

fn patch_values(patch: &TextStylePatch) -> [(&'static str, Option<Any>); 8] {
    [
        ("bold", patch.bold.map(Any::Bool)),
        ("italic", patch.italic.map(Any::Bool)),
        ("fontSize", patch.font_size_pt.map(Any::Number)),
        ("color", patch.color.as_deref().map(Any::from)),
        ("fontFamily", patch.font_family.as_deref().map(Any::from)),
        ("underline", patch.underline.as_deref().map(Any::from)),
        ("spacing", patch.spacing_pt.map(Any::Number)),
        ("baseline", patch.baseline_pct.map(Any::Number)),
    ]
}

fn style_from_run_properties(properties: &RunProperties, theme: Option<&Theme>) -> TextStyle {
    TextStyle {
        bold: properties.bold,
        italic: properties.italic,
        font_size_pt: properties.font_size_pt,
        color: resolve_color_value_to_hex_with_theme(properties.color.as_ref(), theme),
        font_family: properties.font_family.clone(),
        underline: properties.underline.clone(),
        spacing_pt: properties.spacing_pt,
        baseline_pct: properties.baseline_pct,
        kern_pt: properties.kern_pt,
        caps: properties.caps,
    }
}

pub(crate) fn style_from_attrs(attrs: Option<&Attrs>) -> TextStyle {
    TextStyle {
        bold: attrs.and_then(|attrs| any_bool(attrs.get("bold"))),
        italic: attrs.and_then(|attrs| any_bool(attrs.get("italic"))),
        font_size_pt: attrs.and_then(|attrs| any_number(attrs.get("fontSize"))),
        color: attrs.and_then(|attrs| any_string(attrs.get("color"))),
        font_family: attrs.and_then(|attrs| any_string(attrs.get("fontFamily"))),
        underline: attrs.and_then(|attrs| any_string(attrs.get("underline"))),
        spacing_pt: attrs.and_then(|attrs| any_number(attrs.get("spacing"))),
        baseline_pct: attrs.and_then(|attrs| any_number(attrs.get("baseline"))),
        kern_pt: attrs.and_then(|attrs| any_number(attrs.get("kern"))),
        caps: attrs
            .and_then(|attrs| any_string(attrs.get("caps")))
            .as_deref()
            .and_then(TextCaps::from_attribute),
    }
}

fn any_bool(value: Option<&Any>) -> Option<bool> {
    match value {
        Some(Any::Bool(value)) => Some(*value),
        _ => None,
    }
}

fn any_number(value: Option<&Any>) -> Option<f64> {
    match value {
        Some(Any::Number(value)) if value.is_finite() => Some(*value),
        Some(Any::BigInt(value)) => Some(*value as f64),
        _ => None,
    }
}

fn any_string(value: Option<&Any>) -> Option<String> {
    match value {
        Some(Any::String(value)) => Some(value.to_string()),
        _ => None,
    }
}

fn map_string<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<String> {
    match map.get(txn, key) {
        Some(Out::Any(Any::String(value))) => Some(value.to_string()),
        _ => None,
    }
}

fn map_number<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<f64> {
    match map.get(txn, key) {
        Some(Out::Any(Any::Number(value))) if value.is_finite() => Some(value),
        Some(Out::Any(Any::BigInt(value))) => Some(value as f64),
        _ => None,
    }
}

fn out_len(value: &Out) -> u32 {
    match value {
        Out::Any(Any::String(value)) => value.encode_utf16().count() as u32,
        _ => 1,
    }
}
