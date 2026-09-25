//! What a source sheet holds beyond its cells, read from the retained package within a
//! budget: drawing objects with their alt text and authored anchors, comment counts and
//! the worksheet features the model does not carry. Nothing here fails a read; a
//! malformed part is reported, and one that hits a cap or the budget stops the reading.

use std::ops::ControlFlow;

use quick_xml::events::Event;
use xlsx_model::ChartAnchor;

use crate::MAX_DEPTH;
use crate::chart::{
    NS_RELATIONSHIPS, NS_SPREADSHEET_DRAWING, anchor_geometry, directory_of, is_anchor,
    relationship_part_path, relationships_of, type_is,
};
use crate::package::PreservedPackage;
use crate::tree::{Element, exceeds_limits, parse_tree_within};
use crate::xml::{find_part, local_name, next_event, reader, resolve_part_path};

const NS_MARKUP_COMPATIBILITY: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const URI_CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const URI_CHART_EX: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const URI_DIAGRAM: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

/// Worksheet children the model drops, with the name an export reports them by.
const SHEET_FEATURES: [(&str, &str); 7] = [
    ("conditionalFormatting", "conditional formatting"),
    ("dataValidations", "data validation"),
    ("autoFilter", "autofilter"),
    ("scenarios", "scenarios"),
    ("oleObjects", "embedded objects"),
    ("controls", "form controls"),
    ("picture", "background picture"),
];

/// Sheet relationships the model drops, by type suffix.
const RELATED_FEATURES: [(&str, &str); 4] = [
    ("pivotTable", "pivot tables"),
    ("queryTable", "query tables"),
    ("slicer", "slicers"),
    ("timeline", "timelines"),
];

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DrawingObjectKind {
    Chart,
    Picture,
    Shape,
    Group,
    Connector,
    /// SmartArt.
    Diagram,
    /// A graphic frame holding neither a chart nor a diagram.
    GraphicFrame,
    /// Ink.
    ContentPart,
    Unknown,
}

/// One drawing anchor of a source sheet as its part authors it.
#[derive(Clone, Debug, PartialEq)]
pub struct DrawingObject {
    pub drawing: String,
    /// Zero-based element-child ordinal of the anchor under the drawing root.
    pub ordinal: u32,
    /// Position among the drawing's `xdr:` anchors, as `SheetChart::anchor_index`
    /// counts them; `None` for an anchor wrapped in alternate content.
    pub anchor_index: Option<usize>,
    pub kind: DrawingObjectKind,
    pub name: Option<String>,
    /// Alt text.
    pub description: Option<String>,
    pub title: Option<String>,
    pub hidden: bool,
    /// `None` when the anchor's geometry does not read.
    pub anchor: Option<ChartAnchor>,
    /// The chart or image part its relationship names, resolved.
    pub target: Option<String>,
}

/// What [`PreservedPackage::source_sheet_inventory`] found besides drawing objects.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct SheetInventory {
    /// Relationship and comment parts that are malformed.
    pub unreadable_parts: Vec<String>,
    pub comments: usize,
    pub threaded_comments: usize,
    /// Worksheet features the model does not carry, in a fixed order.
    pub features: Vec<&'static str>,
    /// The part whose reading hit a parser cap or the inspection budget; the inventory
    /// stopped there.
    pub limited: Option<String>,
}

/// One step of [`PreservedPackage::visit_source_sheet_objects`].
#[derive(Clone, Debug, PartialEq)]
pub enum SourceObject {
    Object(DrawingObject),
    /// A malformed drawing part, whose objects are unknown.
    Unreadable(String),
    /// A part whose reading hit a parser cap or the inspection budget; nothing after it is
    /// visited.
    Limited(String),
}

/// What a source inspection may still read: XML nodes (elements, attributes and stream
/// events) and part bytes. Every read spends from it; one that would overspend stops.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct InspectionBudget {
    pub nodes: u64,
    pub bytes: u64,
}

enum Read<T> {
    Done(T),
    Malformed,
    Limited,
}

impl InspectionBudget {
    fn take_bytes(&mut self, bytes: &[u8]) -> bool {
        let Some(left) = self.bytes.checked_sub(bytes.len() as u64) else {
            return false;
        };
        self.bytes = left;
        true
    }

    fn tree(&mut self, bytes: &[u8]) -> Read<Element> {
        if !self.take_bytes(bytes) {
            return Read::Limited;
        }
        let cap = usize::try_from(self.nodes).unwrap_or(usize::MAX);
        match parse_tree_within(bytes, cap) {
            Ok((root, spent)) => {
                self.nodes -= spent as u64;
                Read::Done(root)
            }
            Err(error) if exceeds_limits(&error) => {
                self.nodes = self
                    .nodes
                    .saturating_sub(cap.min(crate::MAX_TREE_NODES) as u64);
                Read::Limited
            }
            Err(_) => Read::Malformed,
        }
    }

    fn relationships(&mut self, bytes: &[u8]) -> Read<Vec<(String, String, String)>> {
        match self.tree(bytes) {
            Read::Done(root) => Read::Done(relationships_of(&root)),
            Read::Malformed => Read::Malformed,
            Read::Limited => Read::Limited,
        }
    }

    /// How many `local` elements `bytes` holds, streamed one event at a time.
    fn count(&mut self, bytes: &[u8], local: &[u8]) -> Read<usize> {
        if !self.take_bytes(bytes) {
            return Read::Limited;
        }
        let mut reader = reader(bytes);
        let mut buffer = Vec::new();
        let mut depth = 0;
        let mut count = 0;
        loop {
            let Some(left) = self.nodes.checked_sub(1) else {
                return Read::Limited;
            };
            self.nodes = left;
            match next_event(&mut reader, &mut buffer, &mut depth) {
                Ok(Event::Start(element)) if local_name(&element) == local => count += 1,
                Ok(Event::Eof) => return Read::Done(count),
                Ok(_) => {}
                Err(error) if exceeds_limits(&error) => return Read::Limited,
                Err(_) => return Read::Malformed,
            }
        }
    }
}

impl PreservedPackage {
    /// What source sheet `index` holds beyond its cells and drawing objects, read within
    /// `budget`.
    pub fn source_sheet_inventory(
        &self,
        index: usize,
        budget: &mut InspectionBudget,
    ) -> SheetInventory {
        let mut inventory = SheetInventory::default();
        let Some(sheet) = self.sheets.get(index) else {
            return inventory;
        };
        let parts = &self.parts;
        for (element, feature) in SHEET_FEATURES {
            if sheet
                .template
                .children
                .iter()
                .any(|child| child.local_name == element)
            {
                inventory.features.push(feature);
            }
        }
        let rels_path = relationship_part_path(&sheet.path);
        let relationships =
            match find_part(parts, &rels_path).map(|bytes| budget.relationships(bytes)) {
                None => Vec::new(),
                Some(Read::Done(relationships)) => relationships,
                Some(Read::Malformed) => {
                    inventory.unreadable_parts.push(rels_path);
                    Vec::new()
                }
                Some(Read::Limited) => {
                    inventory.limited = Some(rels_path);
                    return inventory;
                }
            };
        for (suffix, feature) in RELATED_FEATURES {
            if relationships
                .iter()
                .any(|(_, kind, _)| type_is(kind, suffix))
            {
                inventory.features.push(feature);
            }
        }
        let directory = directory_of(&sheet.path);
        for (_, kind, target) in &relationships {
            let local: &[u8] = if type_is(kind, "comments") {
                b"comment"
            } else if type_is(kind, "threadedComment") {
                b"threadedComment"
            } else {
                continue;
            };
            let path = resolve_part_path(directory, target);
            let Some(bytes) = find_part(parts, &path) else {
                inventory.unreadable_parts.push(path);
                continue;
            };
            match budget.count(bytes, local) {
                Read::Done(count) if local == b"comment" => inventory.comments += count,
                Read::Done(count) => inventory.threaded_comments += count,
                Read::Malformed => inventory.unreadable_parts.push(path),
                Read::Limited => {
                    inventory.limited = Some(path);
                    return inventory;
                }
            }
        }
        inventory
    }

    /// Hands `visit` the drawing objects of source sheet `index` in drawing, then anchor
    /// order, reading one part at a time within `budget`, until `visit` breaks or a part
    /// hits a limit.
    pub fn visit_source_sheet_objects(
        &self,
        index: usize,
        budget: &mut InspectionBudget,
        mut visit: impl FnMut(SourceObject) -> ControlFlow<()>,
    ) -> ControlFlow<()> {
        let Some(sheet) = self.sheets.get(index) else {
            return ControlFlow::Continue(());
        };
        let parts = &self.parts;
        let rels_path = relationship_part_path(&sheet.path);
        let relationships =
            match find_part(parts, &rels_path).map(|bytes| budget.relationships(bytes)) {
                Some(Read::Done(relationships)) => relationships,
                Some(Read::Limited) => {
                    visit(SourceObject::Limited(rels_path))?;
                    return ControlFlow::Break(());
                }
                None | Some(Read::Malformed) => return ControlFlow::Continue(()),
            };
        let directory = directory_of(&sheet.path);
        let mut drawings: Vec<String> = Vec::new();
        for (_, kind, target) in &relationships {
            let path = resolve_part_path(directory, target);
            if type_is(kind, "drawing") && !drawings.contains(&path) {
                drawings.push(path);
            }
        }
        for drawing in drawings {
            let Some(bytes) = find_part(parts, &drawing) else {
                continue;
            };
            let root = match budget.tree(bytes) {
                Read::Done(root) => root,
                Read::Malformed => {
                    visit(SourceObject::Unreadable(drawing))?;
                    continue;
                }
                Read::Limited => {
                    visit(SourceObject::Limited(drawing))?;
                    return ControlFlow::Break(());
                }
            };
            let drawing_rels = relationship_part_path(&drawing);
            let relationships =
                match find_part(parts, &drawing_rels).map(|bytes| budget.relationships(bytes)) {
                    Some(Read::Done(relationships)) => relationships,
                    Some(Read::Limited) => {
                        visit(SourceObject::Limited(drawing_rels))?;
                        return ControlFlow::Break(());
                    }
                    None | Some(Read::Malformed) => Vec::new(),
                };
            let base = directory_of(&drawing).to_owned();
            let resolve = |id: &str| {
                relationships
                    .iter()
                    .find(|(rel_id, _, _)| rel_id == id)
                    .map(|(_, _, target)| resolve_part_path(&base, target))
            };
            let mut anchor_index = 0;
            for (ordinal, child) in root.child_elements().enumerate() {
                let (anchor, index) = if is_anchor(&child) {
                    anchor_index += 1;
                    (child, Some(anchor_index - 1))
                } else if child.is(NS_MARKUP_COMPATIBILITY, "AlternateContent") {
                    match alternate(child).filter(is_anchor) {
                        Some(anchor) => (anchor, None),
                        None => continue,
                    }
                } else {
                    continue;
                };
                visit(SourceObject::Object(drawing_object(
                    &drawing,
                    ordinal as u32,
                    index,
                    anchor,
                    &resolve,
                )))?;
            }
        }
        ControlFlow::Continue(())
    }
}

fn drawing_object(
    drawing: &str,
    ordinal: u32,
    anchor_index: Option<usize>,
    anchor: &Element,
    resolve: &impl Fn(&str) -> Option<String>,
) -> DrawingObject {
    let content = anchor
        .child_elements()
        .find(|child| {
            !matches!(
                child.local_name(),
                "from" | "to" | "pos" | "ext" | "clientData"
            )
        })
        .map(|child| {
            if child.is(NS_MARKUP_COMPATIBILITY, "AlternateContent") {
                alternate(child).unwrap_or(child)
            } else {
                child
            }
        });
    let kind = content.map_or(DrawingObjectKind::Unknown, object_kind);
    let properties = content.and_then(|content| {
        content
            .child_elements()
            .find(|child| child.local_name().starts_with("nv"))
            .and_then(|non_visual| non_visual.child("cNvPr"))
    });
    let text = |name: &str| {
        properties
            .and_then(|properties| properties.attribute_local(name))
            .filter(|value| !value.is_empty())
            .map(str::to_owned)
    };
    let target = content.and_then(|content| {
        let reference = match kind {
            DrawingObjectKind::Picture => find_descendant(content, "blip", 0)
                .and_then(|blip| blip.attribute_ns(NS_RELATIONSHIPS, "embed")),
            DrawingObjectKind::Chart => find_descendant(content, "chart", 0)
                .and_then(|chart| chart.attribute_ns(NS_RELATIONSHIPS, "id")),
            _ => None,
        };
        reference.and_then(resolve)
    });
    DrawingObject {
        drawing: drawing.to_owned(),
        ordinal,
        anchor_index,
        kind,
        name: text("name"),
        description: text("descr"),
        title: text("title"),
        hidden: properties
            .and_then(|properties| properties.attribute_local("hidden"))
            .is_some_and(|hidden| matches!(hidden, "1" | "true")),
        anchor: anchor_geometry(anchor).ok(),
        target,
    }
}

fn object_kind(content: &Element) -> DrawingObjectKind {
    if content.namespace() != Some(NS_SPREADSHEET_DRAWING) {
        return DrawingObjectKind::Unknown;
    }
    match content.local_name() {
        "pic" => DrawingObjectKind::Picture,
        "sp" => DrawingObjectKind::Shape,
        "grpSp" => DrawingObjectKind::Group,
        "cxnSp" => DrawingObjectKind::Connector,
        "contentPart" => DrawingObjectKind::ContentPart,
        "graphicFrame" => match find_descendant(content, "graphicData", 0)
            .and_then(|data| data.attribute_local("uri"))
        {
            Some(URI_CHART | URI_CHART_EX) => DrawingObjectKind::Chart,
            Some(URI_DIAGRAM) => DrawingObjectKind::Diagram,
            _ => DrawingObjectKind::GraphicFrame,
        },
        _ => DrawingObjectKind::Unknown,
    }
}

/// The first element of an alternate-content block's first choice, else of its fallback.
fn alternate(block: &Element) -> Option<&Element> {
    block
        .child_elements()
        .filter(|branch| matches!(branch.local_name(), "Choice" | "Fallback"))
        .find_map(|branch| branch.child_elements().next())
}

fn find_descendant<'a>(element: &'a Element, local: &str, depth: usize) -> Option<&'a Element> {
    if depth > MAX_DEPTH {
        return None;
    }
    element.child_elements().find_map(|child| {
        if child.local_name() == local {
            Some(child)
        } else {
            find_descendant(child, local, depth + 1)
        }
    })
}
