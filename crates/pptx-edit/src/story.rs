use std::sync::Arc;

use ooxml_drawingml::{Theme, resolve_color_value_to_hex_with_theme};
use pptx_parse::{RunProperties, TextBody};
use yrs::types::Attrs;
use yrs::types::text::YChange;
use yrs::{
    Any, Map, MapPrelim, MapRef, Out, ReadTxn, Text, TextPrelim, TextRef, Transact, TransactionMut,
};

use crate::{
    DeckSession, EditError, EditResult, KIND, PARA_ID, PILCROW_KIND, ParagraphSnapshot, STORIES,
    StorySnapshot, TextReceipt, TextRunSnapshot, TextStyle, TextStylePatch,
};

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

impl DeckSession {
    pub fn story(&self, story_id: &str) -> EditResult<StorySnapshot> {
        let txn = self.doc.transact();
        let story = story_ref(&txn, story_id)?;
        snapshot_story(&story, &txn, story_id)
    }

    pub fn insert_text(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        index: u32,
        text: &str,
        style: &TextStyle,
    ) -> EditResult<TextReceipt> {
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

    pub fn format_text(
        &self,
        context: &crate::EditCtx,
        story_id: &str,
        start: u32,
        end: u32,
        patch: &TextStylePatch,
    ) -> EditResult<TextReceipt> {
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        check_text_range(&story, &txn, start, end)?;
        let text = text_in_range(&story, &txn, start, end);
        if end > start {
            story.format(&mut txn, start, end - start, attrs_from_patch(patch));
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
        let mut txn = self.transact_for(context);
        let story = story_ref(&txn, story_id)?;
        let final_pilcrow = final_pilcrow_index(&story, &txn)?;
        if index > final_pilcrow {
            return Err(EditError::OutOfBounds {
                index,
                length: final_pilcrow,
            });
        }
        let paragraph_id = self.next_id("para");
        let pilcrow = story.insert_embed_with_attributes(
            &mut txn,
            index,
            MapPrelim::default(),
            Attrs::default(),
        );
        pilcrow.insert(&mut txn, KIND, PILCROW_KIND);
        pilcrow.insert(&mut txn, PARA_ID, paragraph_id);
        pilcrow.insert(&mut txn, "level", 0_f64);
        Ok(TextReceipt {
            story_id: story_id.to_owned(),
            start: index,
            end: index + 1,
            text: "\n".to_owned(),
        })
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

pub(crate) fn snapshot_story<T: ReadTxn>(
    story: &TextRef,
    txn: &T,
    story_id: &str,
) -> EditResult<StorySnapshot> {
    validate_story(story, txn, story_id)?;
    let mut paragraphs = Vec::new();
    let mut runs = Vec::new();
    for diff in story.diff(txn, YChange::identity) {
        match diff.insert {
            Out::Any(Any::String(text)) => runs.push(TextRunSnapshot {
                text: text.to_string(),
                style: style_from_attrs(diff.attributes.as_deref()),
            }),
            Out::YMap(map) => {
                paragraphs.push(ParagraphSnapshot {
                    id: map_string(&map, txn, PARA_ID).unwrap_or_default(),
                    alignment: map_string(&map, txn, "alignment"),
                    level: map_number(&map, txn, "level").unwrap_or_default() as u32,
                    bullet_json: map_string(&map, txn, "bulletJson"),
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

fn story_ref<T: ReadTxn>(txn: &T, story_id: &str) -> EditResult<TextRef> {
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
    let length = story.len(txn);
    if start > end || end > length {
        return Err(EditError::OutOfBounds {
            index: end.max(start),
            length,
        });
    }
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let item_length = out_len(&diff.insert);
        if matches!(diff.insert, Out::YMap(_)) && start <= offset && offset < end {
            return Err(EditError::ParagraphBoundary { start, end });
        }
        offset += item_length;
    }
    Ok(())
}

fn text_in_range<T: ReadTxn>(story: &TextRef, txn: &T, start: u32, end: u32) -> String {
    let mut output = String::new();
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        let length = out_len(&diff.insert);
        if let Out::Any(Any::String(text)) = diff.insert {
            let overlap_start = start.saturating_sub(offset).min(length);
            let overlap_end = end.saturating_sub(offset).min(length);
            if overlap_end > overlap_start {
                let utf16 = text.encode_utf16().collect::<Vec<_>>();
                output.push_str(&String::from_utf16_lossy(
                    &utf16[overlap_start as usize..overlap_end as usize],
                ));
            }
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

fn style_values(style: &TextStyle) -> [(&'static str, Any); 6] {
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
    ]
}

fn attrs_from_patch(patch: &TextStylePatch) -> Attrs {
    let mut attrs = Attrs::default();
    insert_option(&mut attrs, "bold", patch.bold.map(Any::Bool));
    insert_option(&mut attrs, "italic", patch.italic.map(Any::Bool));
    insert_option(&mut attrs, "fontSize", patch.font_size_pt.map(Any::Number));
    insert_option(&mut attrs, "color", patch.color.as_deref().map(Any::from));
    insert_option(
        &mut attrs,
        "fontFamily",
        patch.font_family.as_deref().map(Any::from),
    );
    insert_option(
        &mut attrs,
        "underline",
        patch.underline.as_deref().map(Any::from),
    );
    attrs
}

fn insert_option(attrs: &mut Attrs, key: &str, value: Option<Any>) {
    if let Some(value) = value {
        attrs.insert(Arc::from(key), value);
    }
}

fn style_from_run_properties(properties: &RunProperties, theme: Option<&Theme>) -> TextStyle {
    TextStyle {
        bold: properties.bold,
        italic: properties.italic,
        font_size_pt: properties.font_size_pt,
        color: resolve_color_value_to_hex_with_theme(properties.color.as_ref(), theme),
        font_family: properties.font_family.clone(),
        underline: properties.underline.clone(),
    }
}

fn style_from_attrs(attrs: Option<&Attrs>) -> TextStyle {
    TextStyle {
        bold: attrs.and_then(|attrs| any_bool(attrs.get("bold"))),
        italic: attrs.and_then(|attrs| any_bool(attrs.get("italic"))),
        font_size_pt: attrs.and_then(|attrs| any_number(attrs.get("fontSize"))),
        color: attrs.and_then(|attrs| any_string(attrs.get("color"))),
        font_family: attrs.and_then(|attrs| any_string(attrs.get("fontFamily"))),
        underline: attrs.and_then(|attrs| any_string(attrs.get("underline"))),
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
