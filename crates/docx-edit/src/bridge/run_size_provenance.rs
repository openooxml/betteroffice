use std::collections::BTreeMap;
use std::ops::Range;
use std::sync::{Arc, Weak};

use docx_parse::paragraph::{HexIdAllocator, parse_paragraph};
use docx_parse::xml::{ParseBudget, XmlElement, XmlNode};
use yrs::ReadTxn;
use yrs::types::text::{Diff, YChange};

use crate::{EditingDoc, SegmentContent, StorySegment};

const MARKER_START: char = '\u{fdd0}';
const MARKER_END: char = '\u{fdd1}';

#[derive(Clone, Debug, Default)]
pub(crate) struct ParagraphSizes {
    fingerprint: [u8; 32],
    runs: Vec<RunSize>,
}

#[derive(Clone, Default)]
pub(super) struct Sources(Weak<crate::identity::SourceIndex>);

impl std::fmt::Debug for Sources {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Sources").finish_non_exhaustive()
    }
}

impl Sources {
    pub(super) fn new(doc: &EditingDoc) -> Self {
        Self(
            doc.source_index()
                .as_ref()
                .map(Arc::downgrade)
                .unwrap_or_default(),
        )
    }

    pub(super) fn unchanged<T: ReadTxn>(
        &self,
        key: &str,
        story: &str,
        source: &Weak<crate::seed::SourceMetadata>,
        txn: &T,
        chunks: &[Diff<YChange>],
    ) -> Option<Arc<ParagraphSizes>> {
        let index = self.0.upgrade()?;
        let source = source.upgrade()?;
        let proof = index.run_size_source(key, story, &source)?;
        let fingerprint =
            crate::identity::run_size_fingerprint(chunks.iter().map(|diff| StorySegment {
                content: crate::segment_content(diff.insert.clone(), txn),
                attributes: crate::ordered_attrs(diff.attributes.as_deref()),
            }))?;
        (fingerprint == proof.fingerprint).then_some(proof)
    }
}

#[derive(Clone, Debug)]
struct RunSize {
    range: Range<u32>,
    direct: bool,
    direct_cs: bool,
}

impl ParagraphSizes {
    pub(super) fn inherited(&self, range: Range<u32>, rtl: bool) -> bool {
        self.runs
            .binary_search_by_key(&(range.start, range.end), |run| {
                (run.range.start, run.range.end)
            })
            .ok()
            .is_some_and(|index| {
                let run = &self.runs[index];
                !run.direct && !(rtl && run.direct_cs)
            })
    }
}

pub(crate) fn paragraph_sizes(
    paragraph: &XmlElement,
    path: &str,
    budget: &mut ParseBudget<'_>,
    source: &crate::seed::SourceMetadata,
) -> Option<ParagraphSizes> {
    let mut ids = HexIdAllocator::from_sha256(&"0".repeat(64)).ok()?;
    let parsed = parse_paragraph(
        paragraph,
        None,
        None,
        None,
        None,
        None,
        path,
        budget,
        &mut ids,
        None,
        false,
        0,
    )
    .ok()?;
    let scratch = EditingDoc::new(0);
    let blocks = [serde_json::to_value(parsed).ok()?];
    crate::seed::seed_blocks(&scratch, Some(source), &[("body".to_owned(), &blocks)]).ok()?;
    let seeded = scratch.story_segments("body").ok()?;
    let original = projection(&seeded);
    let fingerprint = crate::identity::run_size_fingerprint(seeded.into_iter())?;
    if original.chars().any(|ch| matches!(ch, MARKER_START | MARKER_END)) {
        return None;
    }
    let mut marked = paragraph.clone();
    let mut sizes = Vec::new();
    mark_runs(&mut marked, &mut sizes);
    if sizes.is_empty() {
        return None;
    }
    let mut ids = HexIdAllocator::from_sha256(&"0".repeat(64)).ok()?;
    let parsed = parse_paragraph(
        &marked,
        None,
        None,
        None,
        None,
        None,
        path,
        budget,
        &mut ids,
        None,
        false,
        0,
    )
    .ok()?;
    let scratch = EditingDoc::new(0);
    let blocks = [serde_json::to_value(parsed).ok()?];
    crate::seed::seed_blocks(&scratch, None, &[("body".to_owned(), &blocks)]).ok()?;
    let projected = projection(&scratch.story_segments("body").ok()?);
    let mut remaining = projected.as_str();
    let mut content = String::new();
    let mut starts = BTreeMap::new();
    let mut runs = Vec::new();
    let mut offset = 0;
    while let Some(at) = remaining.find(MARKER_START) {
        let text = &remaining[..at];
        content.push_str(text);
        offset += text.encode_utf16().count() as u32;
        remaining = &remaining[at + MARKER_START.len_utf8()..];
        let end = remaining.find(MARKER_END)?;
        let (id, side) = remaining[..end].split_once(':')?;
        let id = id.parse::<usize>().ok()?;
        let &(direct, direct_cs) = sizes.get(id)?;
        match side {
            "s" => {
                if starts.insert(id, offset).is_some() {
                    return None;
                }
            }
            "e" => {
                let start = starts.remove(&id)?;
                if start < offset {
                    runs.push(RunSize {
                        range: start..offset,
                        direct,
                        direct_cs,
                    });
                }
            }
            _ => return None,
        }
        remaining = &remaining[end + MARKER_END.len_utf8()..];
    }
    content.push_str(remaining);
    runs.sort_unstable_by_key(|run| (run.range.start, run.range.end));
    (starts.is_empty() && content == original).then_some(ParagraphSizes { fingerprint, runs })
}

fn projection(segments: &[StorySegment]) -> String {
    let mut content = String::new();
    for segment in segments {
        match &segment.content {
            SegmentContent::Text(text) => content.push_str(text),
            SegmentContent::OtherEmbed { .. } => content.push('\0'),
            SegmentContent::Pilcrow(_) => {}
        }
    }
    content
}

fn mark_runs(node: &mut XmlElement, sizes: &mut Vec<(bool, bool)>) {
    if node.local_name() == "r" {
        let text_only = node.child_elements().all(|child| {
            matches!(
                child.local_name(),
                "rPr" | "t" | "delText" | "tab" | "softHyphen" | "noBreakHyphen" | "sym"
            )
        });
        if text_only {
            let properties = node.child("w", "rPr");
            let direct = properties.is_some_and(|properties| properties.child("w", "sz").is_some());
            let direct_cs =
                properties.is_some_and(|properties| properties.child("w", "szCs").is_some());
            let id = sizes.len();
            sizes.push((direct, direct_cs));
            let start = marker(format!("{MARKER_START}{id}:s{MARKER_END}"));
            let end = marker(format!("{MARKER_START}{id}:e{MARKER_END}"));
            let index = node
                .children
                .iter()
                .position(|child| {
                    matches!(child, XmlNode::Element(child) if child.local_name() != "rPr")
                })
                .unwrap_or(node.children.len());
            node.children.insert(index, XmlNode::Element(start));
            node.children.push(XmlNode::Element(end));
        }
        return;
    }
    for child in &mut node.children {
        if let XmlNode::Element(child) = child {
            mark_runs(child, sizes);
        }
    }
}

fn marker(text: String) -> XmlElement {
    XmlElement {
        name: "w:t".to_owned(),
        attributes: Default::default(),
        children: vec![XmlNode::Text(text)],
        paragraph_ordinal: None,
        para_id_attribute: None,
        bookmark_xml: None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
    use crate::identity::{RUN_SIZE_FINGERPRINTS, RUN_SIZE_SOURCE_PARTS, STORY_FINGERPRINTS};
    use crate::seed::fixture::{Package, image, para, run};
    use crate::{EditCtx, EngineSession, FormatPolicy, Position, seed_from_docx};

    #[test]
    fn ordinary_body_typing_checks_only_toc_paragraph_sizes() {
        let styles = concat!(
            r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
            r#"<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr>"#,
            r#"</w:rPrDefault></w:docDefaults>"#,
            r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>"#,
            r#"<w:style w:type="paragraph" w:styleId="TOC1">"#,
            r#"<w:basedOn w:val="Normal"/></w:style>"#,
            r#"<w:style w:type="character" w:styleId="Hyperlink">"#,
            r#"<w:rPr><w:sz w:val="28"/></w:rPr></w:style></w:styles>"#,
        );
        let env = RenderEnv {
            paragraph_style_font_sizes: BTreeMap::from([("TOC1".to_owned(), 10.0)]),
            hyperlink_style_font_sizes: BTreeMap::from([("Hyperlink".to_owned(), 14.0)]),
            doc_default_font_size: Some(10.0),
            ..Default::default()
        };
        for count in [1, 128] {
            let mut body = para(
                "10000001",
                concat!(
                    r#"<w:pPr><w:pStyle w:val="TOC1"/></w:pPr>"#,
                    r#"<w:hyperlink w:anchor="entry"><w:r><w:rPr>"#,
                    r#"<w:rStyle w:val="Hyperlink"/></w:rPr>"#,
                    r#"<w:t>Entry one</w:t></w:r></w:hyperlink>"#,
                ),
            );
            for index in 0..count {
                body.push_str(&para(&format!("{:08X}", 0x20000000 + index), &run("Body")));
            }
            body.push_str(&para("30000001", &image("rIdImage", "inline")));
            let bytes = Package::new(&body)
                .styles(styles)
                .rel("rIdImage", "image", "media/image1.png")
                .bytes();
            let engine = EngineSession::new(77112);
            seed_from_docx(engine.doc(), &bytes).unwrap();
            engine
                .with_lowered_story("body", &env, |blocks| {
                    let value = serde_json::to_value(blocks).unwrap();
                    assert_eq!(value[0]["runs"][0]["fontSize"], 10.0);
                })
                .unwrap();
            RUN_SIZE_SOURCE_PARTS.with(|counter| counter.set(0));
            for text in ["x", "y"] {
                STORY_FINGERPRINTS.with(|counter| counter.set(0));
                RUN_SIZE_FINGERPRINTS.with(|counter| counter.set(0));
                engine
                    .doc()
                    .insert_text(
                        &EditCtx::local("", ""),
                        Position::new("body", 12),
                        text,
                        FormatPolicy::Inherit,
                    )
                    .unwrap();
                engine.with_lowered_story("body", &env, |_| ()).unwrap();
                STORY_FINGERPRINTS.with(|counter| assert_eq!(counter.get(), 0));
                RUN_SIZE_FINGERPRINTS.with(|counter| assert!(counter.get() <= 1));
                RUN_SIZE_FINGERPRINTS.with(|counter| counter.set(0));
                yrs_doc_to_layout_blocks(engine.doc(), "body", &env).unwrap();
                STORY_FINGERPRINTS.with(|counter| assert_eq!(counter.get(), 0));
                RUN_SIZE_FINGERPRINTS.with(|counter| assert_eq!(counter.get(), 1));
                RUN_SIZE_SOURCE_PARTS.with(|counter| assert_eq!(counter.get(), 0));
            }
        }
    }
}
