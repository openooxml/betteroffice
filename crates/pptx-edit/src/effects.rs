use std::collections::HashMap;

use pptx_parse::{PptxPackage, ShapeNode};

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

pub(crate) fn preserve_theme_styles(target: &mut PptxPackage, stored: &PptxPackage) {
    for theme in &mut target.themes {
        if let Some(stored) = stored
            .themes
            .iter()
            .find(|stored| stored.part_path == theme.part_path)
            && !stored.format_scheme.effects.is_empty()
        {
            theme
                .format_scheme
                .effects
                .clone_from(&stored.format_scheme.effects);
        }
    }
    let sources: HashMap<_, _> = stored
        .slides
        .iter()
        .map(|part| (&part.part_path, &part.shapes))
        .chain(
            stored
                .layouts
                .iter()
                .map(|part| (&part.part_path, &part.shapes)),
        )
        .chain(
            stored
                .masters
                .iter()
                .map(|part| (&part.part_path, &part.shapes)),
        )
        .collect();
    for (path, shapes) in target
        .slides
        .iter_mut()
        .map(|part| (&part.part_path, &mut part.shapes))
        .chain(
            target
                .layouts
                .iter_mut()
                .map(|part| (&part.part_path, &mut part.shapes)),
        )
        .chain(
            target
                .masters
                .iter_mut()
                .map(|part| (&part.part_path, &mut part.shapes)),
        )
    {
        if let Some(stored) = sources.get(path) {
            preserve_shape_effect_references(shapes, stored);
        }
    }
}

fn preserve_shape_effect_references(targets: &mut [ShapeNode], stored: &[ShapeNode]) {
    for source in stored {
        let Some(target) = targets
            .iter_mut()
            .find(|target| shape_base(target).id == shape_base(source).id)
        else {
            continue;
        };
        let (target, source) = match (target, source) {
            (ShapeNode::Shape(target), ShapeNode::Shape(source)) => {
                (&mut target.style, &source.style)
            }
            (ShapeNode::Picture(target), ShapeNode::Picture(source)) => {
                (&mut target.style, &source.style)
            }
            (ShapeNode::Group(target), ShapeNode::Group(source)) => {
                preserve_shape_effect_references(&mut target.children, &source.children);
                continue;
            }
            _ => continue,
        };
        if let Some(effect) = source.as_ref().and_then(|style| style.effect.as_ref()) {
            target.get_or_insert_with(Default::default).effect = Some(effect.clone());
        }
    }
}
