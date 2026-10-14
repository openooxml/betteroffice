//! Input decoding and oracle of the `pptx-package-parse` fuzz target.

use std::borrow::Cow;

use pptx_parse::{
    GraphicFrameData, ParseLimits, PptxError, PptxPackage, ShapeNode, TextBody,
    parse_pptx_with_limits,
};

fn limits() -> ParseLimits {
    ParseLimits {
        max_xml_bytes: 4 << 20,
        max_xml_events: 200_000,
        max_xml_text_bytes: 4 << 20,
        max_xml_depth: 64,
        max_attributes_per_element: 128,
        max_attribute_bytes: 64 << 10,
        max_relationships: 64,
        max_shapes: 96,
        max_paragraphs: 96,
        max_runs: 128,
        max_comments: 16,
    }
}

fn package_bytes(data: &[u8]) -> Option<Cow<'_, [u8]>> {
    let (mode, payload) = data.split_first()?;
    if mode & 1 == 1 {
        return Some(Cow::Borrowed(payload));
    }
    let parts = payload
        .split(|byte| *byte == 0)
        .map(|record| {
            let split = record
                .iter()
                .position(|byte| *byte == b'\n')
                .unwrap_or(record.len());
            let (path, content) = record.split_at(split);
            (
                String::from_utf8_lossy(path).into_owned(),
                content.get(1..).unwrap_or_default().to_vec(),
            )
        })
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).ok().map(Cow::Owned)
}

#[derive(Default)]
struct Counts {
    shapes: usize,
    paragraphs: usize,
    runs: usize,
}

impl Counts {
    fn add_shapes(&mut self, shapes: &[ShapeNode]) {
        for shape in shapes {
            self.shapes += 1;
            match shape {
                ShapeNode::Shape(shape) => {
                    if let Some(text) = &shape.text {
                        self.add_text(text);
                    }
                }
                ShapeNode::Picture(_) => {}
                ShapeNode::GraphicFrame(frame) => match &frame.data {
                    GraphicFrameData::Table(table) => {
                        for cell in table.rows.iter().flat_map(|row| &row.cells) {
                            self.add_text(&cell.text);
                        }
                    }
                    GraphicFrameData::Unknown {
                        picture: Some(_), ..
                    } => self.shapes += 1,
                    _ => {}
                },
                ShapeNode::Group(group) => self.add_shapes(&group.children),
            }
        }
    }

    fn add_text(&mut self, body: &TextBody) {
        self.paragraphs += body.paragraphs.len();
        self.runs += body
            .paragraphs
            .iter()
            .map(|paragraph| paragraph.runs.len())
            .sum::<usize>();
    }
}

fn assert_within_limits(package: &PptxPackage, limits: &ParseLimits) {
    let mut counts = Counts::default();
    for slide in &package.slides {
        counts.add_shapes(&slide.shapes);
    }
    for layout in &package.layouts {
        counts.add_shapes(&layout.shapes);
    }
    for master in &package.masters {
        counts.add_shapes(&master.shapes);
    }
    let relationships = package.relationships.values().map(Vec::len).sum::<usize>();
    let comments = package.comments.len() + package.comment_authors.len();

    assert!(
        counts.shapes <= limits.max_shapes,
        "shapes {}",
        counts.shapes
    );
    assert!(
        counts.paragraphs <= limits.max_paragraphs,
        "paragraphs {}",
        counts.paragraphs
    );
    assert!(counts.runs <= limits.max_runs, "runs {}", counts.runs);
    assert!(
        relationships <= limits.max_relationships,
        "relationships {relationships}"
    );
    assert!(comments <= limits.max_comments, "comments {comments}");
}

/// Parses one fuzz input; `None` when it does not decode to a package.
pub fn run(data: &[u8]) -> Option<Result<(), PptxError>> {
    let bytes = package_bytes(data)?;
    let limits = limits();
    Some(
        parse_pptx_with_limits(&bytes, &limits)
            .map(|package| assert_within_limits(&package, &limits)),
    )
}
