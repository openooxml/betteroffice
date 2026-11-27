//! Placeholder geometry inheritance: the layout or master transform a slide
//! shape draws at when its own `a:xfrm` is absent.

use pptx_parse::{Placeholder, PptxPackage, ShapeNode, ShapeTransform, SlideLayout, SlideMaster};

use crate::paragraph::{SlideParents, find_placeholder};
use crate::{InheritedGeometry, ShapeSnapshot};

/// Source shapes and the parts a slide inherits from.
pub(crate) struct SlideContext<'a> {
    pub(crate) layout: Option<&'a SlideLayout>,
    pub(crate) master: Option<&'a SlideMaster>,
    pub(crate) source_shapes: &'a [ShapeNode],
}

impl<'a> SlideContext<'a> {
    pub(crate) fn new(
        package: &'a PptxPackage,
        source_part_path: Option<&str>,
        layout_part_path: Option<&str>,
    ) -> Self {
        let parents = SlideParents::resolve(package, source_part_path, layout_part_path);
        Self {
            layout: parents.layout,
            master: parents.master,
            source_shapes: parents
                .slide
                .map(|slide| slide.shapes.as_slice())
                .unwrap_or_default(),
        }
    }

    fn placeholder_nodes(&self, placeholder: &Placeholder) -> impl Iterator<Item = &'a ShapeNode> {
        [
            self.layout.map(|layout| layout.shapes.as_slice()),
            self.master.map(|master| master.shapes.as_slice()),
        ]
        .into_iter()
        .flatten()
        .filter_map(move |nodes| find_placeholder(nodes, placeholder))
    }
}

/// Records the inherited transform on every shape that draws at one.
pub(crate) fn record_inherited(shapes: &mut [ShapeSnapshot], context: &SlideContext<'_>) {
    for shape in shapes {
        shape.inherited = (shape.width <= 0 || shape.height <= 0)
            .then(|| inherited_transform(context, shape.source_id, shape.placeholder.as_ref()))
            .flatten()
            .map(|transform| InheritedGeometry {
                x: transform.x,
                y: transform.y,
                width: transform.width,
                height: transform.height,
                rotation_deg: transform.rotation_deg,
                flip_h: transform.flip_h,
                flip_v: transform.flip_v,
            });
        record_inherited(&mut shape.children, context);
    }
}

/// The transform a placeholder takes from the layout's matching placeholder,
/// then the master's, while its source node spells out none of its own.
pub(crate) fn inherited_transform<'a>(
    context: &SlideContext<'a>,
    source_id: u32,
    placeholder: Option<&Placeholder>,
) -> Option<&'a ShapeTransform> {
    let placeholder = placeholder.filter(|_| !has_own_transform(context, source_id))?;
    layout_transform(context, placeholder)
}

/// The layout's or master's transform for `placeholder`, whatever the slide
/// node spells out itself.
pub(crate) fn layout_transform<'a>(
    context: &SlideContext<'a>,
    placeholder: &Placeholder,
) -> Option<&'a ShapeTransform> {
    context
        .placeholder_nodes(placeholder)
        .map(node_transform)
        .find(|transform| transform.width > 0 && transform.height > 0)
}

pub(crate) fn layout_geometry<'a>(
    context: &SlideContext<'a>,
    placeholder: &Placeholder,
) -> Option<&'a pptx_parse::Shape> {
    context
        .placeholder_nodes(placeholder)
        .find_map(|node| match node {
            ShapeNode::Shape(shape) if shape.has_preset_geometry || shape.geometry == "custom" => {
                Some(shape)
            }
            _ => None,
        })
}

/// Whether the slide's own node spells out any part of a transform, which
/// keeps the shape from inheriting one.
fn has_own_transform(context: &SlideContext<'_>, source_id: u32) -> bool {
    source_id != 0
        && find_source_node(context.source_shapes, source_id)
            .is_some_and(|node| node_transform(node) != &ShapeTransform::default())
}

fn find_source_node(nodes: &[ShapeNode], source_id: u32) -> Option<&ShapeNode> {
    for node in nodes {
        if node.id() == source_id {
            return Some(node);
        }
        if let ShapeNode::Group(group) = node
            && let Some(found) = find_source_node(&group.children, source_id)
        {
            return Some(found);
        }
    }
    None
}

fn node_transform(node: &ShapeNode) -> &ShapeTransform {
    &node_base(node).transform
}

fn node_base(node: &ShapeNode) -> &pptx_parse::ShapeBase {
    match node {
        ShapeNode::Shape(shape) => &shape.base,
        ShapeNode::Picture(shape) => &shape.base,
        ShapeNode::GraphicFrame(shape) => &shape.base,
        ShapeNode::Group(shape) => &shape.base,
    }
}
