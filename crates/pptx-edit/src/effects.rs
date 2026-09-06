use std::{collections::HashMap, sync::Arc};

use pptx_parse::{PptxPackage, ShapeNode};
use yrs::{Any, Doc, Map, ReadTxn, Transact};

use crate::{EditError, EditResult, META, MIGRATE_ORIGIN, deck};

pub(crate) fn import_source(doc: &Doc, source: &PptxPackage) -> EditResult<()> {
    let mut package = deck::package_from_doc(doc)?;
    let sources: HashMap<_, _> = source
        .slides
        .iter()
        .map(|part| (&part.part_path, &part.shapes))
        .chain(
            source
                .layouts
                .iter()
                .map(|part| (&part.part_path, &part.shapes)),
        )
        .chain(
            source
                .masters
                .iter()
                .map(|part| (&part.part_path, &part.shapes)),
        )
        .collect();
    let mut changed = false;
    for (path, shapes) in package
        .slides
        .iter_mut()
        .map(|part| (&part.part_path, &mut part.shapes))
        .chain(
            package
                .layouts
                .iter_mut()
                .map(|part| (&part.part_path, &mut part.shapes)),
        )
        .chain(
            package
                .masters
                .iter_mut()
                .map(|part| (&part.part_path, &mut part.shapes)),
        )
    {
        if let Some(source) = sources.get(path) {
            changed |= merge_shapes(shapes, source);
        }
    }
    if changed {
        let bytes =
            serde_json::to_vec(&package).map_err(|error| EditError::Json(error.to_string()))?;
        let mut txn = doc.transact_mut_with(MIGRATE_ORIGIN);
        let meta = txn
            .get_map(META)
            .ok_or_else(|| EditError::InvalidState("missing metadata".into()))?;
        meta.insert(&mut txn, "packageJson", Any::Buffer(Arc::from(bytes)));
    }
    Ok(())
}

fn merge_shapes(targets: &mut [ShapeNode], sources: &[ShapeNode]) -> bool {
    let mut changed = false;
    for source in sources {
        let Some(target) = targets
            .iter_mut()
            .find(|target| deck::shape_base(target).id == deck::shape_base(source).id)
        else {
            continue;
        };
        let (target, source) = match (target, source) {
            (ShapeNode::Shape(target), ShapeNode::Shape(source)) => {
                (&mut target.effects, &source.effects)
            }
            (ShapeNode::Picture(target), ShapeNode::Picture(source)) => {
                (&mut target.shape_effects, &source.shape_effects)
            }
            (ShapeNode::Group(target), ShapeNode::Group(source)) => {
                changed |= merge_shapes(&mut target.children, &source.children);
                continue;
            }
            _ => continue,
        };
        if target.is_none() && source.is_some() {
            *target = source.clone();
            changed = true;
        }
    }
    changed
}
