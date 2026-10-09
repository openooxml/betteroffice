use std::collections::HashMap;

use pptx_parse::ShapeNode;

use crate::deck::{SourceImport, shape_base};

pub(crate) fn import_source(import: &mut SourceImport<'_>) {
    let sources: HashMap<_, _> = import
        .source
        .slides
        .iter()
        .map(|part| (&part.part_path, &part.shapes))
        .chain(
            import
                .source
                .layouts
                .iter()
                .map(|part| (&part.part_path, &part.shapes)),
        )
        .chain(
            import
                .source
                .masters
                .iter()
                .map(|part| (&part.part_path, &part.shapes)),
        )
        .collect();
    let package = &mut import.package;
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
            merge_shapes(shapes, source);
        }
    }
}

fn merge_shapes(targets: &mut [ShapeNode], sources: &[ShapeNode]) -> bool {
    let mut changed = false;
    for source in sources {
        let Some(target) = targets
            .iter_mut()
            .find(|target| shape_base(target).id == shape_base(source).id)
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
