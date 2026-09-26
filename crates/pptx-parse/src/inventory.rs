//! Source locations of a slide part's modeled content and an inventory of what the model leaves
//! out, found while parsing by walking the part's XML the way the parser does and retained on
//! the package.

use crate::model::ShapeElements;
use crate::relationships::Relationship;
use crate::xml::{XmlElement, alternate_content_branch};

/// A slide part's shapes, in the order and nesting of [`crate::Slide::shapes`], and the
/// shape-tree children the model does not represent.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SlideSource {
    pub shapes: Vec<SourceShape>,
    pub omitted: Vec<OmittedElement>,
}

/// Where one modeled shape sits in its part, and the source metadata the model does not keep.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SourceShape {
    /// Element-child ordinals from the part's root element.
    pub path: Vec<u32>,
    /// The element's qualified name, such as `p:sp`.
    pub element: String,
    /// `cNvPr/@id`.
    pub id: u32,
    /// `cNvPr/@title`.
    pub title: Option<String>,
    /// `cNvPr/@descr`.
    pub description: Option<String>,
    /// What a picture plays, when it is a media frame.
    pub media: Option<MediaKind>,
    /// `a:graphicData/@uri` of a graphic frame.
    pub graphic_uri: Option<String>,
    /// The part's relationships the element references, in document order; none for a group.
    pub relationship_ids: Vec<String>,
    /// Paragraph children the model does not represent.
    pub omitted_inlines: Vec<OmittedInline>,
    pub children: Vec<SourceShape>,
    /// Children of a group the model does not represent.
    pub omitted: Vec<OmittedElement>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MediaKind {
    Video,
    Audio,
}

/// A shape-tree child the model does not represent.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OmittedElement {
    pub path: Vec<u32>,
    /// The qualified name of the element, or of the branch content an unreadable
    /// `mc:AlternateContent` holds.
    pub element: String,
    /// Modeled siblings that precede it.
    pub position: usize,
    /// `cNvPr/@name` and `cNvPr/@descr` when the element carries non-visual properties.
    pub name: Option<String>,
    pub description: Option<String>,
    /// `cNvPr/@hidden`.
    pub hidden: bool,
    /// The part's relationships the element references, in document order.
    pub relationship_ids: Vec<String>,
}

/// A paragraph child the model does not represent, such as an equation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OmittedInline {
    pub path: Vec<u32>,
    pub element: String,
    /// The table cell holding the paragraph, as (row, cell).
    pub cell: Option<(usize, usize)>,
    /// Index of the paragraph among its text body's `a:p` elements.
    pub paragraph: usize,
    /// Modeled runs (`a:r`, `a:fld`, `a:br`) of the paragraph that precede it.
    pub position: usize,
}

/// A slide's notes page, beyond the notes text the model keeps.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct NotesSource {
    pub part_path: String,
    /// Shapes with text other than the notes body and the page's header, footer, date,
    /// slide-number and slide-image placeholders.
    pub other_text_shapes: usize,
}

/// The inventory of a parsed slide part's root element.
pub(crate) fn slide_source(
    root: &XmlElement,
    relationships: &[Relationship],
    elements: ShapeElements,
) -> SlideSource {
    let mut source = SlideSource::default();
    let Some((common_index, common)) = child_named(root, "cSld") else {
        return source;
    };
    let Some((tree_index, tree)) = child_named(common, "spTree") else {
        return source;
    };
    let walk = Walk {
        elements,
        relationships,
    };
    walk.children(
        tree,
        &[common_index, tree_index],
        &mut source.shapes,
        &mut source.omitted,
    );
    source
}

/// The inventory of a parsed notes page's root element.
pub(crate) fn notes_source(root: &XmlElement, part_path: &str) -> NotesSource {
    let other_text_shapes = root
        .child("cSld")
        .and_then(|common| common.child("spTree"))
        .map_or(0, |tree| {
            tree.descendants_named("sp")
                .into_iter()
                .filter(|shape| !is_notes_decoration(shape) && has_text(shape))
                .count()
        });
    NotesSource {
        part_path: part_path.to_owned(),
        other_text_shapes,
    }
}

fn is_notes_decoration(shape: &XmlElement) -> bool {
    let placeholder = shape
        .child("nvSpPr")
        .and_then(|properties| properties.child("nvPr"))
        .and_then(|properties| properties.child("ph"));
    placeholder.is_some_and(|placeholder| {
        matches!(
            placeholder.attribute("type"),
            Some("body" | "sldImg" | "hdr" | "ftr" | "dt" | "sldNum")
        )
    })
}

fn has_text(shape: &XmlElement) -> bool {
    shape.child("txBody").is_some_and(|body| {
        body.descendants_named("t")
            .iter()
            .any(|text| !text.text_content().is_empty())
    })
}

fn child_named<'a>(parent: &'a XmlElement, name: &str) -> Option<(u32, &'a XmlElement)> {
    parent
        .child_elements()
        .enumerate()
        .find(|(_, child)| child.local_name() == name)
        .map(|(index, child)| (index as u32, child))
}

fn extend(path: &[u32], index: usize) -> Vec<u32> {
    let mut path = path.to_vec();
    path.push(index as u32);
    path
}

/// Mirrors the parser's shape-tree walk over one part.
struct Walk<'a> {
    elements: ShapeElements,
    relationships: &'a [Relationship],
}

impl Walk<'_> {
    /// `mc:AlternateContent` contributes the shapes of the branch it reads, and elements the
    /// model counts as shapes become [`SourceShape`]s.
    fn children(
        &self,
        parent: &XmlElement,
        path: &[u32],
        shapes: &mut Vec<SourceShape>,
        omitted: &mut Vec<OmittedElement>,
    ) {
        for (index, child) in parent.child_elements().enumerate() {
            let child_path = extend(path, index);
            let local = child.local_name();
            if local == "AlternateContent" {
                match branch_with_index(child) {
                    Some((branch_index, branch)) => {
                        self.children(branch, &extend(&child_path, branch_index), shapes, omitted)
                    }
                    None => omitted.push(self.omitted(child, child_path, shapes.len())),
                }
                continue;
            }
            if self.elements.contains(local) {
                shapes.push(self.shape(child, child_path));
            } else if !matches!(local, "nvGrpSpPr" | "grpSpPr" | "extLst") {
                omitted.push(self.omitted(child, child_path, shapes.len()));
            }
        }
    }

    fn omitted(&self, element: &XmlElement, path: Vec<u32>, position: usize) -> OmittedElement {
        let properties = element.descendants_named("cNvPr").into_iter().next();
        OmittedElement {
            path,
            element: content_name(element).to_owned(),
            position,
            name: properties
                .and_then(|properties| properties.attribute("name"))
                .filter(|name| !name.is_empty())
                .map(str::to_owned),
            description: properties
                .and_then(|properties| properties.attribute("descr"))
                .map(str::to_owned),
            hidden: properties
                .and_then(|properties| properties.attribute("hidden"))
                .is_some_and(|value| matches!(value, "1" | "true" | "on")),
            relationship_ids: self.relationship_ids(element),
        }
    }

    fn shape(&self, element: &XmlElement, path: Vec<u32>) -> SourceShape {
        let non_visual = element.child_elements().find(|child| {
            matches!(
                child.local_name(),
                "nvSpPr" | "nvCxnSpPr" | "nvPicPr" | "nvGraphicFramePr" | "nvGrpSpPr"
            )
        });
        let properties = non_visual.and_then(|value| value.child("cNvPr"));
        let mut shape = SourceShape {
            element: element.name.clone(),
            id: properties
                .and_then(|value| value.attribute("id"))
                .and_then(|value| value.parse().ok())
                .unwrap_or_default(),
            title: properties
                .and_then(|value| value.attribute("title"))
                .map(str::to_owned),
            description: properties
                .and_then(|value| value.attribute("descr"))
                .map(str::to_owned),
            media: non_visual
                .and_then(|value| value.child("nvPr"))
                .and_then(media_kind),
            ..SourceShape::default()
        };
        match element.local_name() {
            "sp" | "cxnSp" => {
                if let Some((index, body)) = child_named(element, "txBody") {
                    text_body_inlines(
                        body,
                        &extend(&path, index as usize),
                        None,
                        &mut shape.omitted_inlines,
                    );
                }
            }
            "graphicFrame" => {
                let graphic = child_named(element, "graphic");
                let data = graphic.and_then(|(_, graphic)| child_named(graphic, "graphicData"));
                shape.graphic_uri = data
                    .and_then(|(_, data)| data.attribute("uri"))
                    .map(str::to_owned);
                if let (Some((graphic_index, _)), Some((data_index, data))) = (graphic, data)
                    && let Some((table_index, table)) = child_named(data, "tbl")
                {
                    let table_path =
                        [path.as_slice(), &[graphic_index, data_index, table_index]].concat();
                    table_inlines(table, &table_path, &mut shape.omitted_inlines);
                }
            }
            "grpSp" => self.children(element, &path, &mut shape.children, &mut shape.omitted),
            _ => {}
        }
        if element.local_name() != "grpSp" {
            shape.relationship_ids = self.relationship_ids(element);
        }
        shape.path = path;
        shape
    }

    /// Values of namespaced attributes, anywhere in `element`, that name one of the part's
    /// relationships.
    fn relationship_ids(&self, element: &XmlElement) -> Vec<String> {
        let mut ids = Vec::new();
        let mut pending = vec![element];
        while let Some(current) = pending.pop() {
            for (name, value) in &current.attributes {
                if name.contains(':')
                    && !name.starts_with("xmlns")
                    && !ids.contains(value)
                    && self
                        .relationships
                        .iter()
                        .any(|relationship| &relationship.id == value)
                {
                    ids.push(value.clone());
                }
            }
            let children: Vec<&XmlElement> = current.child_elements().collect();
            pending.extend(children.into_iter().rev());
        }
        ids
    }
}

fn branch_with_index(alternate: &XmlElement) -> Option<(usize, &XmlElement)> {
    let branch = alternate_content_branch(alternate)?;
    alternate
        .child_elements()
        .enumerate()
        .find(|(_, child)| std::ptr::eq(*child, branch))
}

/// An element's name, or for `mc:AlternateContent` the name of its first choice's content.
fn content_name(element: &XmlElement) -> &str {
    let content = if element.local_name() == "AlternateContent" {
        element
            .child_elements()
            .flat_map(XmlElement::child_elements)
            .next()
            .unwrap_or(element)
    } else {
        element
    };
    &content.name
}

fn media_kind(properties: &XmlElement) -> Option<MediaKind> {
    let mut media = None;
    for child in properties.child_elements() {
        match child.local_name() {
            "videoFile" | "quickTimeFile" => return Some(MediaKind::Video),
            "audioFile" | "audioCd" | "wavAudioFile" => media = Some(MediaKind::Audio),
            "extLst" if media.is_none() && !child.descendants_named("media").is_empty() => {
                media = Some(MediaKind::Video);
            }
            _ => {}
        }
    }
    media
}

fn table_inlines(table: &XmlElement, path: &[u32], output: &mut Vec<OmittedInline>) {
    let rows = table
        .child_elements()
        .enumerate()
        .filter(|(_, child)| child.local_name() == "tr");
    for (row_index, (row_ordinal, row)) in rows.enumerate() {
        let cells = row
            .child_elements()
            .enumerate()
            .filter(|(_, child)| child.local_name() == "tc");
        for (cell_index, (cell_ordinal, cell)) in cells.enumerate() {
            if let Some((body_index, body)) = child_named(cell, "txBody") {
                let body_path =
                    [path, &[row_ordinal as u32, cell_ordinal as u32, body_index]].concat();
                text_body_inlines(body, &body_path, Some((row_index, cell_index)), output);
            }
        }
    }
}

fn text_body_inlines(
    body: &XmlElement,
    path: &[u32],
    cell: Option<(usize, usize)>,
    output: &mut Vec<OmittedInline>,
) {
    let paragraphs = body
        .child_elements()
        .enumerate()
        .filter(|(_, child)| child.local_name() == "p");
    for (paragraph, (paragraph_ordinal, element)) in paragraphs.enumerate() {
        let mut runs = 0;
        for (index, child) in element.child_elements().enumerate() {
            match child.local_name() {
                "r" | "fld" | "br" => runs += 1,
                "pPr" | "endParaRPr" => {}
                _ => {
                    output.push(OmittedInline {
                        path: [path, &[paragraph_ordinal as u32, index as u32]].concat(),
                        element: content_name(child).to_owned(),
                        cell,
                        paragraph,
                        position: runs,
                    });
                }
            }
        }
    }
}
