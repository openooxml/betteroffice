use ooxml_drawingml::ShapeOutline;
use serde_json::Value;
use yrs::{Any, Map, MapRef, ReadTxn, Transact};

use crate::deck::SourceImport;
use crate::{DeckSession, EditError, EditResult, META, MIGRATE_ORIGIN, ShapeSnapshot};

pub(crate) fn import_source(
    session: &DeckSession,
    import: &mut SourceImport<'_>,
) -> EditResult<()> {
    let pending = {
        let txn = session.doc.transact();
        txn.get_map(META).is_some_and(|meta| {
            meta.get(&txn, "outlineGradientsPendingSource") == Some(Any::Bool(true).into())
        })
    };
    if !pending {
        return Ok(());
    }
    let mut sources = Vec::new();
    {
        let original = import.source_snapshot()?;
        for slide in &original.slides {
            collect_outlines(&slide.shapes, &mut sources);
        }
    }
    let mut package = serde_json::to_value(&import.package)
        .map_err(|error| EditError::Json(error.to_string()))?;
    let source =
        serde_json::to_value(import.source).map_err(|error| EditError::Json(error.to_string()))?;
    let merged = merge_gradients(&mut package, &source);
    if !merged && sources.is_empty() {
        return Ok(());
    }
    if merged {
        import.package =
            serde_json::from_value(package).map_err(|error| EditError::Json(error.to_string()))?;
    }
    let mut txn = session.doc.transact_mut_with(MIGRATE_ORIGIN);
    let shapes = txn
        .get_map(crate::SHAPES)
        .ok_or_else(|| EditError::InvalidState("missing shapes".into()))?;
    for (id, outline) in sources {
        let Some(shape) = shapes
            .get(&txn, &id)
            .and_then(|value| value.cast::<MapRef>().ok())
        else {
            continue;
        };
        let Some(current) = shape.get(&txn, "outlineJson") else {
            continue;
        };
        let current: ShapeOutline = serde_json::from_str(&current.to_string(&txn))
            .map_err(|error| EditError::Json(error.to_string()))?;
        let mut legacy = outline.clone();
        legacy.gradient = None;
        if current == legacy {
            let json = serde_json::to_string(&outline)
                .map_err(|error| EditError::Json(error.to_string()))?;
            shape.insert(&mut txn, "outlineJson", json);
        }
    }
    let meta = txn
        .get_map(META)
        .ok_or_else(|| EditError::InvalidState("missing metadata".into()))?;
    meta.remove(&mut txn, "outlineGradientsPendingSource");
    Ok(())
}

fn collect_outlines(shapes: &[ShapeSnapshot], outlines: &mut Vec<(String, ShapeOutline)>) {
    for shape in shapes {
        if let Some(outline) = &shape.outline
            && outline.gradient.is_some()
        {
            outlines.push((shape.id.clone(), outline.clone()));
        }
        collect_outlines(&shape.children, outlines);
    }
}

fn merge_gradients(target: &mut Value, source: &Value) -> bool {
    let mut changed = false;
    match (target, source) {
        (Value::Object(target), Value::Object(source)) => {
            if source
                .get("outline")
                .and_then(|outline| outline.get("gradient"))
                .is_some()
                && target.get("style") != source.get("style")
            {
                if let Some(style) = source.get("style") {
                    target.insert("style".into(), style.clone());
                } else {
                    target.remove("style");
                }
                changed = true;
            }
            if !target.contains_key("gradient")
                && let Some(gradient) = source.get("gradient")
            {
                target.insert("gradient".into(), gradient.clone());
                changed = true;
            }
            for (key, target) in target {
                if let Some(source) = source.get(key) {
                    changed |= merge_gradients(target, source);
                }
            }
        }
        (Value::Array(target), Value::Array(source)) => {
            for (target, source) in target.iter_mut().zip(source) {
                changed |= merge_gradients(target, source);
            }
        }
        _ => {}
    }
    changed
}
