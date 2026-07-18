use std::collections::HashSet;
use std::sync::Arc;

use ooxml_drawingml::ShapeFill;
use pptx_parse::{PptxPackage, ShapeNode};
use serde::de::DeserializeOwned;
use yrs::{
    Any, Array, ArrayPrelim, ArrayRef, Doc, Map, MapPrelim, MapRef, Out, ReadTxn, TextRef,
    Transact, TransactionMut, WriteTxn,
};

use crate::story::{seed_plain_story, seed_story, snapshot_story, validate_story};
use crate::{
    DeckSession, DeckSnapshot, EditCtx, EditError, EditResult, META, SHAPES, SLIDE_ORDER, SLIDES,
    STORIES, ShapeDraft, ShapeKind, ShapeReceipt, ShapeRect, ShapeSnapshot, SlideReceipt,
    SlideSnapshot, TransformReceipt,
};

const SCHEMA_VERSION: f64 = 1.0;
const MAX_GEOMETRY: i64 = 1_000_000_000_000_000;

pub(crate) fn seed_doc(doc: &Doc, package: &PptxPackage, fingerprint: &str) -> EditResult<()> {
    let mut txn = doc.transact_mut_with("pptx:bootstrap");
    let meta = txn.get_or_insert_map(META);
    meta.insert(&mut txn, "schemaVersion", SCHEMA_VERSION);
    meta.insert(&mut txn, "fingerprint", fingerprint);
    meta.insert(&mut txn, "widthEmu", package.presentation.width_emu as f64);
    meta.insert(
        &mut txn,
        "heightEmu",
        package.presentation.height_emu as f64,
    );
    let order = txn.get_or_insert_array(SLIDE_ORDER);
    let slides = txn.get_or_insert_map(SLIDES);
    let shapes = txn.get_or_insert_map(SHAPES);
    let stories = txn.get_or_insert_map(STORIES);

    for (slide_index, slide) in package.slides.iter().enumerate() {
        let reference = &package.presentation.slides[slide_index];
        let slide_id = format!("slide:{slide_index}:{}", reference.id);
        order.push_back(&mut txn, slide_id.as_str());
        let slide_map = slides.insert(&mut txn, slide_id.as_str(), MapPrelim::default());
        slide_map.insert(&mut txn, "id", slide_id.as_str());
        slide_map.insert(&mut txn, "sourcePartPath", slide.part_path.as_str());
        if let Some(layout) = &slide.layout_part_path {
            slide_map.insert(&mut txn, "layoutPartPath", layout.as_str());
        }
        if let Some(name) = &slide.name {
            slide_map.insert(&mut txn, "name", name.as_str());
        }
        let shape_order = slide_map.insert(&mut txn, "shapes", ArrayPrelim::default());
        for (shape_index, shape) in slide.shapes.iter().enumerate() {
            let shape_id = seed_shape(
                &shapes,
                &stories,
                &mut txn,
                &slide_id,
                &shape_index.to_string(),
                shape,
            )?;
            shape_order.push_back(&mut txn, shape_id.as_str());
        }
    }
    Ok(())
}

fn seed_shape(
    shapes: &MapRef,
    stories: &MapRef,
    txn: &mut TransactionMut<'_>,
    slide_id: &str,
    path: &str,
    shape: &ShapeNode,
) -> EditResult<String> {
    let shape_id = format!("{slide_id}:shape:{path}");
    let shape_map = shapes.insert(txn, shape_id.as_str(), MapPrelim::default());
    let base = match shape {
        ShapeNode::Shape(shape) => &shape.base,
        ShapeNode::Picture(shape) => &shape.base,
        ShapeNode::GraphicFrame(shape) => &shape.base,
        ShapeNode::Group(shape) => &shape.base,
    };
    shape_map.insert(txn, "id", shape_id.as_str());
    shape_map.insert(txn, "sourceId", base.id as f64);
    shape_map.insert(txn, "name", base.name.as_str());
    shape_map.insert(txn, "x", base.transform.x as f64);
    shape_map.insert(txn, "y", base.transform.y as f64);
    shape_map.insert(txn, "width", base.transform.width as f64);
    shape_map.insert(txn, "height", base.transform.height as f64);
    shape_map.insert(txn, "rotationDeg", base.transform.rotation_deg);
    shape_map.insert(txn, "flipH", base.transform.flip_h);
    shape_map.insert(txn, "flipV", base.transform.flip_v);
    insert_json(
        &shape_map,
        txn,
        "placeholderJson",
        base.placeholder.as_ref(),
    )?;

    let mut text_story_ids = Vec::new();
    let mut child_ids = Vec::new();
    match shape {
        ShapeNode::Shape(shape) => {
            shape_map.insert(txn, "kind", "shape");
            shape_map.insert(txn, "geometry", shape.geometry.as_str());
            insert_json(&shape_map, txn, "fillJson", shape.fill.as_ref())?;
            insert_json(&shape_map, txn, "outlineJson", shape.outline.as_ref())?;
            if let Some(body) = &shape.text {
                let story_id = format!("story:{shape_id}:0");
                seed_story(stories, txn, &story_id, body)?;
                text_story_ids.push(story_id);
            }
        }
        ShapeNode::Picture(picture) => {
            shape_map.insert(txn, "kind", "picture");
            shape_map.insert(txn, "geometry", "rect");
            insert_json(&shape_map, txn, "fillJson", picture.fill.as_ref())?;
            insert_json(&shape_map, txn, "outlineJson", picture.outline.as_ref())?;
            if let Some(media) = &picture.media_part_path {
                shape_map.insert(txn, "mediaPartPath", media.as_str());
            }
        }
        ShapeNode::GraphicFrame(frame) => {
            shape_map.insert(txn, "kind", "graphicFrame");
            shape_map.insert(txn, "geometry", "rect");
            insert_json(&shape_map, txn, "graphicJson", Some(&frame.data))?;
            if let pptx_parse::GraphicFrameData::Table { rows } = &frame.data {
                for (row_index, row) in rows.iter().enumerate() {
                    for (cell_index, body) in row.iter().enumerate() {
                        let story_id = format!("story:{shape_id}:table:{row_index}:{cell_index}");
                        seed_story(stories, txn, &story_id, body)?;
                        text_story_ids.push(story_id);
                    }
                }
            }
        }
        ShapeNode::Group(group) => {
            shape_map.insert(txn, "kind", "group");
            shape_map.insert(txn, "geometry", "group");
            for (child_index, child) in group.children.iter().enumerate() {
                child_ids.push(seed_shape(
                    shapes,
                    stories,
                    txn,
                    slide_id,
                    &format!("{path}.{child_index}"),
                    child,
                )?);
            }
        }
    }
    shape_map.insert(txn, "textStories", string_array(&text_story_ids));
    shape_map.insert(txn, "children", string_array(&child_ids));
    Ok(shape_id)
}

fn insert_json<T: serde::Serialize>(
    map: &MapRef,
    txn: &mut TransactionMut<'_>,
    key: &str,
    value: Option<&T>,
) -> EditResult<()> {
    if let Some(value) = value {
        let json =
            serde_json::to_string(value).map_err(|error| EditError::Json(error.to_string()))?;
        map.insert(txn, key, json);
    }
    Ok(())
}

impl DeckSession {
    pub fn snapshot(&self) -> EditResult<DeckSnapshot> {
        snapshot_doc(&self.doc)
    }

    pub fn insert_slide(
        &self,
        context: &EditCtx,
        index: u32,
        layout_part_path: Option<&str>,
    ) -> EditResult<SlideReceipt> {
        let slide_id = self.next_id("slide");
        let mut txn = self.transact_for(context);
        let order = required_order(&txn)?;
        let length = order.len(&txn);
        if index > length {
            return Err(EditError::OutOfBounds { index, length });
        }
        let slides = required_map(&txn, SLIDES)?;
        let slide = slides.insert(&mut txn, slide_id.as_str(), MapPrelim::default());
        slide.insert(&mut txn, "id", slide_id.as_str());
        slide.insert(&mut txn, "name", format!("Slide {}", length + 1));
        if let Some(layout_part_path) = layout_part_path {
            slide.insert(&mut txn, "layoutPartPath", layout_part_path);
        }
        slide.insert(&mut txn, "shapes", ArrayPrelim::default());
        order.insert(&mut txn, index, slide_id.as_str());
        Ok(SlideReceipt {
            slide_id,
            from_index: None,
            to_index: Some(index),
        })
    }

    pub fn delete_slide(&self, context: &EditCtx, slide_id: &str) -> EditResult<SlideReceipt> {
        let mut txn = self.transact_for(context);
        let order = required_order(&txn)?;
        let index = array_index(&order, &txn, slide_id)
            .ok_or_else(|| EditError::SlideNotFound(slide_id.to_owned()))?;
        order.remove(&mut txn, index);
        Ok(SlideReceipt {
            slide_id: slide_id.to_owned(),
            from_index: Some(index),
            to_index: None,
        })
    }

    pub fn move_slide(
        &self,
        context: &EditCtx,
        slide_id: &str,
        to_index: u32,
    ) -> EditResult<SlideReceipt> {
        let mut txn = self.transact_for(context);
        let order = required_order(&txn)?;
        let length = order.len(&txn);
        if to_index >= length {
            return Err(EditError::OutOfBounds {
                index: to_index,
                length,
            });
        }
        let from_index = array_index(&order, &txn, slide_id)
            .ok_or_else(|| EditError::SlideNotFound(slide_id.to_owned()))?;
        if from_index != to_index {
            order.remove(&mut txn, from_index);
            order.insert(&mut txn, to_index, slide_id);
        }
        Ok(SlideReceipt {
            slide_id: slide_id.to_owned(),
            from_index: Some(from_index),
            to_index: Some(to_index),
        })
    }

    pub fn add_text_box(
        &self,
        context: &EditCtx,
        slide_id: &str,
        draft: &ShapeDraft,
    ) -> EditResult<ShapeReceipt> {
        validate_rect(draft.rect)?;
        let shape_id = self.next_id("shape");
        let story_id = format!("story:{shape_id}:0");
        let paragraph_id = self.next_id("para");
        let mut txn = self.transact_for(context);
        let slide = slide_ref(&txn, slide_id)?;
        let order = slide_shape_order(&slide, &txn)?;
        let index = order.len(&txn);
        let shapes = required_map(&txn, SHAPES)?;
        let stories = required_map(&txn, STORIES)?;
        seed_plain_story(
            &stories,
            &mut txn,
            &story_id,
            &paragraph_id,
            &draft.text,
            &draft.style,
        );
        let shape = shapes.insert(&mut txn, shape_id.as_str(), MapPrelim::default());
        shape.insert(&mut txn, "id", shape_id.as_str());
        shape.insert(&mut txn, "sourceId", 0_f64);
        shape.insert(&mut txn, "kind", "shape");
        shape.insert(&mut txn, "name", draft.name.as_str());
        shape.insert(&mut txn, "x", draft.rect.x as f64);
        shape.insert(&mut txn, "y", draft.rect.y as f64);
        shape.insert(&mut txn, "width", draft.rect.width as f64);
        shape.insert(&mut txn, "height", draft.rect.height as f64);
        shape.insert(&mut txn, "rotationDeg", 0_f64);
        shape.insert(&mut txn, "flipH", false);
        shape.insert(&mut txn, "flipV", false);
        shape.insert(&mut txn, "geometry", "rect");
        insert_json(
            &shape,
            &mut txn,
            "fillJson",
            Some(&ShapeFill::named("none")),
        )?;
        shape.insert(
            &mut txn,
            "textStories",
            string_array(std::slice::from_ref(&story_id)),
        );
        shape.insert(&mut txn, "children", string_array(&[]));
        order.push_back(&mut txn, shape_id.as_str());
        Ok(ShapeReceipt {
            slide_id: slide_id.to_owned(),
            shape_id,
            index,
        })
    }

    pub fn remove_shape(
        &self,
        context: &EditCtx,
        slide_id: &str,
        shape_id: &str,
    ) -> EditResult<ShapeReceipt> {
        let mut txn = self.transact_for(context);
        let slide = slide_ref(&txn, slide_id)?;
        let order = slide_shape_order(&slide, &txn)?;
        let index = array_index(&order, &txn, shape_id)
            .ok_or_else(|| EditError::ShapeNotFound(shape_id.to_owned()))?;
        order.remove(&mut txn, index);
        Ok(ShapeReceipt {
            slide_id: slide_id.to_owned(),
            shape_id: shape_id.to_owned(),
            index,
        })
    }

    pub fn move_shape(
        &self,
        context: &EditCtx,
        slide_id: &str,
        shape_id: &str,
        x: i64,
        y: i64,
    ) -> EditResult<TransformReceipt> {
        validate_coordinate(x)?;
        validate_coordinate(y)?;
        let mut txn = self.transact_for(context);
        require_shape_membership(&txn, slide_id, shape_id)?;
        let shape = shape_ref(&txn, shape_id)?;
        let before = shape_rect(&shape, &txn)?;
        shape.insert(&mut txn, "x", x as f64);
        shape.insert(&mut txn, "y", y as f64);
        Ok(TransformReceipt {
            slide_id: slide_id.to_owned(),
            shape_id: shape_id.to_owned(),
            before,
            after: ShapeRect { x, y, ..before },
        })
    }

    pub fn resize_shape(
        &self,
        context: &EditCtx,
        slide_id: &str,
        shape_id: &str,
        width: i64,
        height: i64,
    ) -> EditResult<TransformReceipt> {
        let rect = ShapeRect {
            width,
            height,
            ..ShapeRect::default()
        };
        validate_rect(rect)?;
        let mut txn = self.transact_for(context);
        require_shape_membership(&txn, slide_id, shape_id)?;
        let shape = shape_ref(&txn, shape_id)?;
        let before = shape_rect(&shape, &txn)?;
        shape.insert(&mut txn, "width", width as f64);
        shape.insert(&mut txn, "height", height as f64);
        Ok(TransformReceipt {
            slide_id: slide_id.to_owned(),
            shape_id: shape_id.to_owned(),
            before,
            after: ShapeRect {
                width,
                height,
                ..before
            },
        })
    }
}

pub(crate) fn validate_doc(doc: &Doc) -> EditResult<()> {
    let snapshot = snapshot_doc(doc)?;
    if snapshot.width_emu <= 0 || snapshot.height_emu <= 0 {
        return Err(EditError::InvalidState(
            "slide dimensions must be positive".to_owned(),
        ));
    }
    let txn = doc.transact();
    let meta = required_map(&txn, META)?;
    if map_number(&meta, &txn, "schemaVersion") != Some(SCHEMA_VERSION) {
        return Err(EditError::InvalidState(
            "unsupported deck schema version".to_owned(),
        ));
    }
    if map_string(&meta, &txn, "fingerprint").is_none() {
        return Err(EditError::InvalidState("missing fingerprint".to_owned()));
    }
    let stories = required_map(&txn, STORIES)?;
    for (story_id, value) in stories.iter(&txn) {
        let story = value
            .cast::<TextRef>()
            .map_err(|_| EditError::InvalidState(format!("story {story_id} is not text")))?;
        validate_story(&story, &txn, story_id)?;
    }
    Ok(())
}

fn snapshot_doc(doc: &Doc) -> EditResult<DeckSnapshot> {
    let txn = doc.transact();
    let meta = required_map(&txn, META)?;
    let order = required_order(&txn)?;
    let slides = required_map(&txn, SLIDES)?;
    let shapes = required_map(&txn, SHAPES)?;
    let stories = required_map(&txn, STORIES)?;
    let mut seen_slides = HashSet::new();
    let mut slide_snapshots = Vec::new();
    for slide_id in string_array_ref(&order, &txn) {
        if !seen_slides.insert(slide_id.clone()) {
            continue;
        }
        let slide = slides
            .get(&txn, &slide_id)
            .and_then(|value| value.cast::<MapRef>().ok())
            .ok_or_else(|| EditError::InvalidState(format!("missing slide {slide_id}")))?;
        let shape_order = slide_shape_order(&slide, &txn)?;
        let mut seen_shapes = HashSet::new();
        let mut shape_snapshots = Vec::new();
        for shape_id in string_array_ref(&shape_order, &txn) {
            if seen_shapes.insert(shape_id.clone()) {
                shape_snapshots.push(snapshot_shape(
                    &shapes,
                    &stories,
                    &txn,
                    &shape_id,
                    &mut HashSet::new(),
                )?);
            }
        }
        slide_snapshots.push(SlideSnapshot {
            id: slide_id,
            source_part_path: map_string(&slide, &txn, "sourcePartPath"),
            layout_part_path: map_string(&slide, &txn, "layoutPartPath"),
            name: map_string(&slide, &txn, "name"),
            shapes: shape_snapshots,
        });
    }
    Ok(DeckSnapshot {
        width_emu: required_i64(&meta, &txn, "widthEmu")?,
        height_emu: required_i64(&meta, &txn, "heightEmu")?,
        slides: slide_snapshots,
    })
}

fn snapshot_shape<T: ReadTxn>(
    shapes: &MapRef,
    stories: &MapRef,
    txn: &T,
    shape_id: &str,
    visiting: &mut HashSet<String>,
) -> EditResult<ShapeSnapshot> {
    if !visiting.insert(shape_id.to_owned()) {
        return Err(EditError::InvalidState(format!(
            "shape cycle at {shape_id}"
        )));
    }
    let shape = shapes
        .get(txn, shape_id)
        .and_then(|value| value.cast::<MapRef>().ok())
        .ok_or_else(|| EditError::InvalidState(format!("missing shape {shape_id}")))?;
    let mut text_snapshots = Vec::new();
    for story_id in map_string_array(&shape, txn, "textStories")? {
        let story = stories
            .get(txn, &story_id)
            .and_then(|value| value.cast::<TextRef>().ok())
            .ok_or_else(|| EditError::InvalidState(format!("missing story {story_id}")))?;
        text_snapshots.push(snapshot_story(&story, txn, &story_id)?);
    }
    let mut children = Vec::new();
    for child_id in map_string_array(&shape, txn, "children")? {
        children.push(snapshot_shape(shapes, stories, txn, &child_id, visiting)?);
    }
    visiting.remove(shape_id);
    Ok(ShapeSnapshot {
        id: shape_id.to_owned(),
        source_id: required_u32(&shape, txn, "sourceId")?,
        kind: parse_shape_kind(&required_string(&shape, txn, "kind")?)?,
        name: required_string(&shape, txn, "name")?,
        x: required_i64(&shape, txn, "x")?,
        y: required_i64(&shape, txn, "y")?,
        width: required_i64(&shape, txn, "width")?,
        height: required_i64(&shape, txn, "height")?,
        rotation_deg: map_number(&shape, txn, "rotationDeg").unwrap_or_default(),
        flip_h: map_bool(&shape, txn, "flipH").unwrap_or_default(),
        flip_v: map_bool(&shape, txn, "flipV").unwrap_or_default(),
        geometry: required_string(&shape, txn, "geometry")?,
        placeholder: optional_json(&shape, txn, "placeholderJson")?,
        fill: optional_json(&shape, txn, "fillJson")?,
        outline: optional_json(&shape, txn, "outlineJson")?,
        media_part_path: map_string(&shape, txn, "mediaPartPath"),
        graphic: optional_json(&shape, txn, "graphicJson")?,
        text_stories: text_snapshots,
        children,
    })
}

fn parse_shape_kind(value: &str) -> EditResult<ShapeKind> {
    match value {
        "shape" => Ok(ShapeKind::Shape),
        "picture" => Ok(ShapeKind::Picture),
        "graphicFrame" => Ok(ShapeKind::GraphicFrame),
        "group" => Ok(ShapeKind::Group),
        _ => Err(EditError::InvalidState(format!(
            "unknown shape kind {value}"
        ))),
    }
}

fn required_order<T: ReadTxn>(txn: &T) -> EditResult<ArrayRef> {
    txn.get_array(SLIDE_ORDER)
        .ok_or_else(|| EditError::InvalidState("missing slide order".to_owned()))
}

fn required_map<T: ReadTxn>(txn: &T, name: &str) -> EditResult<MapRef> {
    txn.get_map(name)
        .ok_or_else(|| EditError::InvalidState(format!("missing {name}")))
}

fn slide_ref<T: ReadTxn>(txn: &T, slide_id: &str) -> EditResult<MapRef> {
    required_map(txn, SLIDES)?
        .get(txn, slide_id)
        .and_then(|value| value.cast::<MapRef>().ok())
        .ok_or_else(|| EditError::SlideNotFound(slide_id.to_owned()))
}

fn shape_ref<T: ReadTxn>(txn: &T, shape_id: &str) -> EditResult<MapRef> {
    required_map(txn, SHAPES)?
        .get(txn, shape_id)
        .and_then(|value| value.cast::<MapRef>().ok())
        .ok_or_else(|| EditError::ShapeNotFound(shape_id.to_owned()))
}

fn slide_shape_order<T: ReadTxn>(slide: &MapRef, txn: &T) -> EditResult<ArrayRef> {
    slide
        .get(txn, "shapes")
        .and_then(|value| value.cast::<ArrayRef>().ok())
        .ok_or_else(|| EditError::InvalidState("slide has no shape order".to_owned()))
}

fn require_shape_membership<T: ReadTxn>(txn: &T, slide_id: &str, shape_id: &str) -> EditResult<()> {
    let slide = slide_ref(txn, slide_id)?;
    let order = slide_shape_order(&slide, txn)?;
    if array_index(&order, txn, shape_id).is_some() {
        Ok(())
    } else {
        Err(EditError::ShapeNotFound(shape_id.to_owned()))
    }
}

fn array_index<T: ReadTxn>(array: &ArrayRef, txn: &T, value: &str) -> Option<u32> {
    array
        .iter(txn)
        .enumerate()
        .find(|(_, item)| out_string(item).as_deref() == Some(value))
        .map(|(index, _)| index as u32)
}

fn string_array_ref<T: ReadTxn>(array: &ArrayRef, txn: &T) -> Vec<String> {
    array
        .iter(txn)
        .filter_map(|value| out_string(&value))
        .collect()
}

fn map_string_array<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> EditResult<Vec<String>> {
    match map.get(txn, key) {
        Some(Out::Any(Any::Array(values))) => Ok(values
            .iter()
            .filter_map(|value| match value {
                Any::String(value) => Some(value.to_string()),
                _ => None,
            })
            .collect()),
        None => Ok(Vec::new()),
        _ => Err(EditError::InvalidState(format!("{key} is not an array"))),
    }
}

fn string_array(values: &[String]) -> Any {
    Any::Array(Arc::from(
        values
            .iter()
            .map(|value| Any::from(value.as_str()))
            .collect::<Vec<_>>(),
    ))
}

fn shape_rect<T: ReadTxn>(shape: &MapRef, txn: &T) -> EditResult<ShapeRect> {
    Ok(ShapeRect {
        x: required_i64(shape, txn, "x")?,
        y: required_i64(shape, txn, "y")?,
        width: required_i64(shape, txn, "width")?,
        height: required_i64(shape, txn, "height")?,
    })
}

fn required_u32<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> EditResult<u32> {
    let value = required_i64(map, txn, key)?;
    u32::try_from(value).map_err(|_| {
        EditError::InvalidState(format!("{key} value {value} is outside the u32 range"))
    })
}

fn validate_rect(rect: ShapeRect) -> EditResult<()> {
    validate_coordinate(rect.x)?;
    validate_coordinate(rect.y)?;
    if rect.width <= 0 || rect.height <= 0 {
        return Err(EditError::InvalidGeometry(
            "shape width and height must be positive".to_owned(),
        ));
    }
    validate_coordinate(rect.width)?;
    validate_coordinate(rect.height)
}

fn validate_coordinate(value: i64) -> EditResult<()> {
    if value.unsigned_abs() > MAX_GEOMETRY as u64 {
        return Err(EditError::InvalidGeometry(format!(
            "coordinate {value} exceeds the safe range"
        )));
    }
    Ok(())
}

fn required_string<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> EditResult<String> {
    map_string(map, txn, key)
        .ok_or_else(|| EditError::InvalidState(format!("missing string {key}")))
}

fn map_string<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<String> {
    map.get(txn, key).and_then(|value| out_string(&value))
}

fn out_string(value: &Out) -> Option<String> {
    match value {
        Out::Any(Any::String(value)) => Some(value.to_string()),
        _ => None,
    }
}

fn required_i64<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> EditResult<i64> {
    let number = map_number(map, txn, key)
        .ok_or_else(|| EditError::InvalidState(format!("missing number {key}")))?;
    if !number.is_finite() || number.fract() != 0.0 || number.abs() > MAX_GEOMETRY as f64 {
        return Err(EditError::InvalidState(format!("invalid integer {key}")));
    }
    Ok(number as i64)
}

fn map_number<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<f64> {
    match map.get(txn, key) {
        Some(Out::Any(Any::Number(value))) if value.is_finite() => Some(value),
        Some(Out::Any(Any::BigInt(value))) => Some(value as f64),
        _ => None,
    }
}

fn map_bool<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<bool> {
    match map.get(txn, key) {
        Some(Out::Any(Any::Bool(value))) => Some(value),
        _ => None,
    }
}

fn optional_json<T: DeserializeOwned, R: ReadTxn>(
    map: &MapRef,
    txn: &R,
    key: &str,
) -> EditResult<Option<T>> {
    map_string(map, txn, key)
        .map(|json| {
            serde_json::from_str(&json).map_err(|error| EditError::InvalidState(error.to_string()))
        })
        .transpose()
}
