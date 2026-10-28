use std::sync::{Arc, Weak};

use docx_parse::xml::XmlElement;
use yrs::ReadTxn;
use yrs::types::text::{Diff, YChange};

use crate::{EditingDoc, StorySegment};

#[derive(Clone, Debug)]
pub(crate) struct ParagraphSizes {
    fingerprint: [u8; 32],
    direct_cs: bool,
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

    pub(super) fn unchanged<'a, T: ReadTxn>(
        &self,
        key: &str,
        story: &str,
        txn: &T,
        chunks: &[Diff<YChange>],
        markers: impl Iterator<Item = (u32, &'a str, bool)>,
    ) -> Option<Arc<ParagraphSizes>> {
        let index = self.0.upgrade()?;
        let proof = index.run_size_source(key, story)?;
        let fingerprint = crate::identity::run_size_fingerprint(
            chunks.iter().map(|diff| StorySegment {
                content: crate::segment_content(diff.insert.clone(), txn),
                attributes: crate::ordered_attrs(diff.attributes.as_deref()),
            }),
            markers,
        )?;
        (fingerprint == proof.fingerprint).then_some(proof)
    }
}

impl ParagraphSizes {
    pub(super) fn inherited(&self, rtl: bool) -> bool {
        !(rtl && self.direct_cs)
    }
}

pub(crate) fn paragraph_sizes(
    paragraph: &XmlElement,
    fingerprint: [u8; 32],
) -> Option<ParagraphSizes> {
    use docx_parse::xml::XmlNode;

    enum Run {
        Entry {
            direct: bool,
            direct_cs: bool,
        },
        Tab,
        Empty,
        Begin,
        Separate,
        End,
        Instruction(String),
    }

    #[derive(Default)]
    struct Entries {
        found: bool,
        direct_cs: bool,
    }

    fn elements(node: &XmlElement) -> Option<Vec<&XmlElement>> {
        let mut elements = Vec::new();
        for child in &node.children {
            match child {
                XmlNode::Element(element) => elements.push(element),
                XmlNode::Text(text) | XmlNode::CData(text) if text.trim().is_empty() => {}
                _ => return None,
            }
        }
        Some(elements)
    }

    fn properties(node: &XmlElement) -> bool {
        node.namespace_prefix() == Some("w")
            && node.children.iter().all(|child| match child {
                XmlNode::Element(element) => properties(element),
                XmlNode::Text(text) | XmlNode::CData(text) => text.trim().is_empty(),
            })
    }

    fn text(node: &XmlElement) -> Option<String> {
        let mut text = String::new();
        for child in &node.children {
            match child {
                XmlNode::Text(part) | XmlNode::CData(part) => text.push_str(part),
                XmlNode::Element(_) => return None,
            }
        }
        Some(text)
    }

    fn run(node: &XmlElement) -> Option<Run> {
        if node.name != "w:r" {
            return None;
        }
        let children = elements(node)?;
        let (formatting, content) = match children.split_first() {
            Some((first, rest)) if first.name == "w:rPr" && properties(first) => {
                (Some(*first), rest)
            }
            _ => (None, children.as_slice()),
        };
        if content.is_empty() {
            return Some(Run::Empty);
        }
        if content.len() == 1 && content[0].name == "w:fldChar" {
            if !content[0].children.is_empty() {
                return None;
            }
            return match content[0].attributes.get("w:fldCharType")?.as_str() {
                "begin" => Some(Run::Begin),
                "separate" => Some(Run::Separate),
                "end" => Some(Run::End),
                _ => None,
            };
        }
        if content.iter().all(|child| child.name == "w:instrText") {
            let mut instruction = String::new();
            for child in content {
                instruction.push_str(&text(child)?);
            }
            return Some(Run::Instruction(instruction));
        }
        let mut entry = false;
        let mut tab = false;
        let mut positioned_tab = false;
        for child in content {
            match child.name.as_str() {
                "w:t" => entry |= !text(child)?.is_empty(),
                "w:sym" | "w:softHyphen" | "w:noBreakHyphen" if child.children.is_empty() => {
                    entry = true;
                }
                "w:tab" if child.children.is_empty() => tab = true,
                "w:ptab" if child.children.is_empty() => positioned_tab = true,
                "w:lastRenderedPageBreak" if child.children.is_empty() => {}
                _ => return None,
            }
        }
        if entry && !positioned_tab {
            Some(Run::Entry {
                direct: formatting.is_some_and(|formatting| {
                    formatting.child_elements().any(|child| child.name == "w:sz")
                }),
                direct_cs: formatting.is_some_and(|formatting| {
                    formatting.child_elements().any(|child| child.name == "w:szCs")
                }),
            })
        } else if !entry
            && (tab || positioned_tab)
            && !content.iter().any(|child| child.name == "w:t")
        {
            Some(Run::Tab)
        } else {
            None
        }
    }

    fn marker(node: &XmlElement) -> bool {
        matches!(
            node.name.as_str(),
            "w:bookmarkStart"
                | "w:bookmarkEnd"
                | "w:proofErr"
                | "w:permStart"
                | "w:permEnd"
                | "w:commentRangeStart"
                | "w:commentRangeEnd"
        ) && node.children.is_empty()
    }

    fn kind(instruction: &str, expected: &str) -> bool {
        instruction
            .split_whitespace()
            .next()
            .is_some_and(|kind| kind.eq_ignore_ascii_case(expected))
    }

    fn opening(nodes: &[&XmlElement], start: usize) -> Option<(String, usize)> {
        if !matches!(run(nodes.get(start)?), Some(Run::Begin)) {
            return None;
        }
        let mut cursor = start + 1;
        let mut instruction = String::new();
        let mut instructed = false;
        loop {
            match nodes.get(cursor).and_then(|node| run(node)) {
                Some(Run::Instruction(part)) => {
                    instruction.push_str(&part);
                    instructed = true;
                }
                Some(Run::Empty) => {}
                _ => break,
            }
            cursor += 1;
        }
        if !instructed || !matches!(run(nodes.get(cursor)?), Some(Run::Separate)) {
            return None;
        }
        Some((instruction, cursor + 1))
    }

    fn complex_page(nodes: &[&XmlElement], start: usize) -> Option<usize> {
        let (instruction, mut cursor) = opening(nodes, start)?;
        if !kind(&instruction, "PAGEREF") {
            return None;
        }
        while matches!(
            nodes.get(cursor).and_then(|node| run(node)),
            Some(Run::Entry { .. } | Run::Empty)
        ) {
            cursor += 1;
        }
        matches!(run(nodes.get(cursor)?), Some(Run::End)).then_some(cursor + 1)
    }

    fn item(
        nodes: &[&XmlElement],
        cursor: usize,
        links: bool,
        entries: &mut Entries,
    ) -> Option<usize> {
        let node = nodes.get(cursor)?;
        if marker(node) {
            return Some(cursor + 1);
        }
        if links && node.name == "w:hyperlink" {
            let children = elements(node)?;
            let mut child = 0;
            while child < children.len() {
                child = item(&children, child, false, entries)?;
            }
            return Some(cursor + 1);
        }
        if node.name == "w:fldSimple" {
            if !kind(node.attributes.get("w:instr")?, "PAGEREF") {
                return None;
            }
            return elements(node)?
                .iter()
                .all(|child| matches!(run(child), Some(Run::Entry { .. } | Run::Empty)))
                .then_some(cursor + 1);
        }
        match run(node)? {
            Run::Entry {
                direct: false,
                direct_cs,
            } => {
                entries.found = true;
                entries.direct_cs |= direct_cs;
                Some(cursor + 1)
            }
            Run::Tab | Run::Empty => Some(cursor + 1),
            Run::Begin => complex_page(nodes, cursor),
            _ => None,
        }
    }

    if paragraph.name != "w:p" {
        return None;
    }
    let nodes = elements(paragraph)?;
    let mut cursor = 0;
    if let Some(first) = nodes.first()
        && first.name == "w:pPr"
    {
        if !properties(first) {
            return None;
        }
        cursor += 1;
    }
    while nodes.get(cursor).is_some_and(|node| marker(node)) {
        cursor += 1;
    }
    if matches!(nodes.get(cursor).and_then(|node| run(node)), Some(Run::Begin)) {
        let (instruction, after) = opening(&nodes, cursor)?;
        if kind(&instruction, "TOC") {
            cursor = after;
        }
    }
    let mut entries = Entries::default();
    while cursor < nodes.len() {
        if matches!(run(nodes[cursor]), Some(Run::End)) {
            cursor += 1;
            while nodes.get(cursor).is_some_and(|node| marker(node)) {
                cursor += 1;
            }
            if cursor != nodes.len() {
                return None;
            }
            break;
        }
        cursor = item(&nodes, cursor, true, &mut entries)?;
    }
    entries.found.then_some(ParagraphSizes {
        fingerprint,
        direct_cs: entries.direct_cs,
    })
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;

    use super::*;
    use crate::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
    use crate::identity::{
        RUN_SIZE_FINGERPRINTS, RUN_SIZE_SEEDED_SNAPSHOTS, RUN_SIZE_SOURCE_PARTS,
        STORY_FINGERPRINTS,
    };
    use crate::seed::fixture::{Package, image, para, run};
    use crate::{EditCtx, EngineSession, FormatPolicy, Position, seed_from_docx};

    fn word_toc(direct: Option<usize>, complex_page: bool, inside_sdt: bool) -> Vec<u8> {
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
        let mut body = String::new();
        for index in 0..3 {
            let open = if index == 0 {
                concat!(
                    r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
                    r#"<w:r><w:instrText> TOC \h </w:instrText></w:r>"#,
                    r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
                )
            } else {
                ""
            };
            let close = if index == 2 {
                r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#
            } else {
                ""
            };
            let size = if direct == Some(index) {
                r#"<w:sz w:val="28"/>"#
            } else {
                ""
            };
            let result = concat!(
                r#"<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:webHidden/>"#,
                r#"<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>3</w:t></w:r>"#,
            );
            let page = if complex_page {
                format!(
                    concat!(
                        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
                        r#"<w:r><w:instrText> PAGEREF _Toc1 \h </w:instrText></w:r>"#,
                        r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>{result}"#,
                        r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
                    ),
                    result = result,
                )
            } else {
                format!(r#"<w:fldSimple w:instr=" PAGEREF _Toc1 \h ">{result}</w:fldSimple>"#)
            };
            body.push_str(&para(
                &format!("{:08X}", 0x10000001 + index),
                &format!(
                    concat!(
                        r#"<w:pPr><w:pStyle w:val="TOC1"/><w:rPr>"#,
                        r#"<w:rFonts w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>"#,
                        r#"</w:rPr></w:pPr>{open}<w:hyperlink w:anchor="_Toc1">"#,
                        r#"<w:r><w:rPr><w:rStyle w:val="Hyperlink"/>{size}</w:rPr>"#,
                        r#"<w:t>Entry {number}</w:t></w:r><w:r><w:rPr><w:webHidden/>"#,
                        r#"<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>"#,
                        r#"<w:tab/></w:r>{page}</w:hyperlink>{close}"#,
                    ),
                    open = open,
                    size = size,
                    number = index + 1,
                    page = page,
                    close = close,
                ),
            ));
        }
        let body = if inside_sdt {
            format!(
                concat!(
                    r#"<w:sdt><w:sdtPr><w:id w:val="1"/><w:docPartObj>"#,
                    r#"<w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/>"#,
                    r#"</w:docPartObj></w:sdtPr><w:sdtContent>{body}</w:sdtContent></w:sdt>"#,
                ),
                body = body,
            )
        } else {
            body
        };
        let theme = concat!(
            r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#,
            r#"<a:themeElements><a:fontScheme name="Proof"><a:majorFont>"#,
            r#"<a:latin typeface="Arial"/><a:ea typeface="Proof EA"/>"#,
            r#"<a:cs typeface="Proof CS"/></a:majorFont><a:minorFont>"#,
            r#"<a:latin typeface="Arial"/><a:ea typeface="Proof EA"/>"#,
            r#"<a:cs typeface="Proof CS"/></a:minorFont></a:fontScheme>"#,
            r#"</a:themeElements></a:theme>"#,
        );
        let mut parts = Package::new(&body)
            .styles(styles)
            .part("theme/theme1.xml", "rIdTheme", "theme", "theme", theme)
            .parts();
        for (name, bytes) in &mut parts {
            if name == "[Content_Types].xml" {
                *bytes = String::from_utf8(std::mem::take(bytes))
                    .unwrap()
                    .replace(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.theme+xml",
                        "application/vnd.openxmlformats-officedocument.theme+xml",
                    )
                    .into_bytes();
            }
        }
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    fn assert_word_toc(direct: Option<usize>, inside_sdt: bool) {
        use yrs::{Out, Text, Transact};

        let env = RenderEnv {
            paragraph_style_font_sizes: BTreeMap::from([("TOC1".to_owned(), 10.0)]),
            hyperlink_style_font_sizes: BTreeMap::from([("Hyperlink".to_owned(), 14.0)]),
            doc_default_font_size: Some(10.0),
            ..Default::default()
        };
        for complex_page in [false, true] {
            let bytes = word_toc(direct, complex_page, inside_sdt);
            let story_id = if inside_sdt { "body:sdt0" } else { "body" };
            let engine = EngineSession::new(77113);
            let doc = engine.doc();
            seed_from_docx(doc, &bytes).unwrap();
            let segments = doc.story_segments(story_id).unwrap();
            assert!(segments.iter().any(|segment| {
                matches!(&segment.content, crate::SegmentContent::OtherEmbed { payload, .. }
                    if payload.get("fieldData").is_some_and(|data| {
                        matches!(data, yrs::Any::String(data)
                            if data.contains("Entry 2") && data.contains("Entry 3"))
                    }))
            }));
            let txn = doc.yrs_doc().transact();
            let story = crate::story_ref(&txn, story_id).unwrap();
            let chunks = story.diff(&txn, YChange::identity);
            let comments = crate::bridge::resolve_comment_intervals(&txn, story_id, &env).unwrap();
            let sources = Sources::new(doc);
            let mut start = 0;
            let mut raw_start = 0;
            let mut raw_end = 0;
            let mut entry = 0;
            for (index, chunk) in chunks.iter().enumerate() {
                raw_end += crate::out_len(&chunk.insert);
                if let Out::YMap(pilcrow) = &chunk.insert
                    && crate::is_pilcrow(pilcrow, &txn)
                {
                    let key = crate::map_string(pilcrow, &txn, "paraId").unwrap();
                    let proof = sources.unchanged(
                        &key,
                        story_id,
                        &txn,
                        &chunks[start..=index],
                        comments.markers(raw_start, raw_end),
                    );
                    assert_eq!(proof.is_some(), direct != Some(entry), "entry {}", entry + 1);
                    entry += 1;
                    start = index + 1;
                    raw_start = raw_end;
                }
            }
            assert_eq!(entry, 3);
            drop(txn);
            let blocks = engine
                .with_lowered_story("body", &env, |blocks| serde_json::to_value(blocks).unwrap())
                .unwrap();
            let cold = yrs_doc_to_layout_blocks(doc, "body", &env).unwrap();
            assert_eq!(blocks, serde_json::to_value(cold).unwrap());
            for index in 0..3 {
                let label = format!("Entry {}", index + 1);
                let runs = blocks
                    .as_array()
                    .unwrap()
                    .iter()
                    .filter_map(|block| block["runs"].as_array())
                    .find(|runs| runs.iter().any(|run| run["text"] == label))
                    .unwrap();
                let entry = runs.iter().find(|run| run["text"] == label).unwrap();
                let expected = if direct == Some(index) { 14.0 } else { 10.0 };
                assert_eq!(entry["fontSize"], expected, "{label}");
                if inside_sdt {
                    for run in runs {
                        assert_eq!(run["fontSize"], 10.0, "{label}: {run}");
                    }
                }
            }
        }
    }

    #[test]
    fn three_entry_toc_proofs_use_the_seeded_save_units() {
        assert_word_toc(None, false);
    }

    #[test]
    fn a_direct_size_rejects_only_its_toc_entry_paragraph() {
        for entry in 0..3 {
            assert_word_toc(Some(entry), false);
        }
    }

    const ENTRY: &str = concat!(
        r#"<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>"#,
        "<w:t>Entry</w:t></w:r>",
    );
    const BEGIN: &str = r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r>"#;
    const SEPARATE: &str = r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#;
    const END: &str = r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#;
    const PAGE: &str = concat!(
        r#"<w:r><w:rPr><w:webHidden/><w:sz w:val="28"/><w:szCs w:val="28"/>"#,
        "</w:rPr><w:t>3</w:t></w:r>",
    );

    fn grammar_open(instruction: &str) -> String {
        format!("{BEGIN}<w:r><w:instrText> {instruction} target </w:instrText></w:r>{SEPARATE}")
    }

    fn grammar_proof(content: &str) -> Option<ParagraphSizes> {
        use docx_parse::xml::{ParseBudget, ParseLimits, parse_xml};

        let xml = format!(
            concat!(
                "<w:p ",
                r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
                r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" "#,
                r#"xmlns:x="urn:foreign">{content}</w:p>"#,
            ),
            content = content,
        );
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let document = parse_xml(xml.as_bytes(), "word/document.xml", &mut budget).unwrap();
        paragraph_sizes(document.root().unwrap(), [23; 32])
    }

    #[test]
    fn strict_toc_grammar_accepts_word_entry_paragraphs() {
        let toc = grammar_open("tOc");
        let page = format!("{}{PAGE}{END}", grammar_open("pAgErEf"));
        let simple = format!(r#"<w:fldSimple w:instr=" PAGEREF target ">{PAGE}</w:fldSimple>"#);
        let tab = concat!(
            r#"<w:r><w:rPr><w:webHidden/><w:sz w:val="28"/><w:szCs w:val="28"/>"#,
            "</w:rPr><w:tab/><w:lastRenderedPageBreak/></w:r>",
        );
        let link = format!(r#"<w:hyperlink w:anchor="target">{ENTRY}{tab}{page}</w:hyperlink>"#);
        let markers = concat!(
            r#"<w:bookmarkStart w:id="0" w:name="entry"/><w:bookmarkEnd w:id="0"/>"#,
            r#"<w:proofErr w:type="spellStart"/><w:permStart w:id="1"/>"#,
            r#"<w:permEnd w:id="1"/><w:commentRangeStart w:id="2"/>"#,
            r#"<w:commentRangeEnd w:id="2"/>"#,
        );
        for (label, content) in [
            ("first", format!("{toc}{link}")),
            (
                "empty run before the page separator",
                format!(
                    concat!(
                        "<w:hyperlink>{entry}{tab}",
                        r#"<w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>"#,
                        r#"<w:r><w:rPr><w:webHidden/></w:rPr><w:instrText> PAGEREF _Toc1 \h </w:instrText></w:r>"#,
                        r#"<w:r><w:rPr><w:webHidden/></w:rPr></w:r>"#,
                        r#"<w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>"#,
                        r#"<w:r><w:rPr><w:webHidden/></w:rPr><w:t>3</w:t></w:r>"#,
                        r#"<w:r><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>"#,
                        "</w:hyperlink>",
                    ),
                    entry = ENTRY,
                    tab = tab,
                ),
            ),
            ("middle", link.clone()),
            ("last", format!("{link}{END}")),
            ("single", format!("{toc}{link}{END}")),
            ("simple page", format!("{ENTRY}{tab}{simple}")),
            ("complex page", format!("{ENTRY}{tab}{page}")),
            (
                "simple page in link",
                format!("<w:hyperlink>{ENTRY}{simple}</w:hyperlink>"),
            ),
            (
                "markers",
                format!("<w:pPr/>{markers}{toc}{markers}{link}{END}{markers}"),
            ),
            (
                "empty page results",
                format!(
                    r#"{ENTRY}<w:fldSimple w:instr=" PAGEREF target "/>{}{END}"#,
                    grammar_open("PAGEREF"),
                ),
            ),
            (
                "multiple instruction runs",
                format!(
                    concat!(
                        "{BEGIN}<w:r><w:instrText>T</w:instrText></w:r>",
                        "<w:r><w:instrText>O</w:instrText><w:instrText>C</w:instrText></w:r>",
                        "{SEPARATE}{ENTRY}{END}",
                    ),
                    BEGIN = BEGIN,
                    SEPARATE = SEPARATE,
                    ENTRY = ENTRY,
                    END = END,
                ),
            ),
            (
                "entry content",
                format!(
                    concat!(
                        "<w:r><w:t>Entry</w:t><w:tab/><w:lastRenderedPageBreak/>",
                        r#"<w:sym w:font="Arial" w:char="0041"/>"#,
                        "<w:softHyphen/><w:noBreakHyphen/></w:r>{tab}",
                        "<w:r><w:rPr/><w:ptab/><w:lastRenderedPageBreak/></w:r>",
                    ),
                    tab = tab,
                ),
            ),
            (
                "formatted field controls",
                format!(
                    "{}{}{ENTRY}{}",
                    toc.replace("<w:r>", r#"<w:r><w:rPr><w:sz w:val="28"/></w:rPr>"#),
                    "<w:hyperlink/>",
                    END.replace("<w:r>", r#"<w:r><w:rPr><w:szCs w:val="28"/></w:rPr>"#),
                ),
            ),
        ] {
            let proof = grammar_proof(&content).unwrap_or_else(|| panic!("{label}: {content}"));
            assert_eq!(proof.fingerprint, [23; 32], "{label}");
            assert!(proof.inherited(true), "{label}");
        }
    }

    #[test]
    fn strict_toc_grammar_rejects_malformed_or_unsupported_content() {
        let toc = grammar_open("TOC");
        let page = grammar_open("PAGEREF");
        let reference = format!("{}{ENTRY}{END}", grammar_open("REF"));
        let simple_open = r#"<w:fldSimple w:instr=" PAGEREF target ">"#;
        let instruction = "<w:r><w:instrText> stray </w:instrText></w:r>";
        let mut cases = vec![
            (
                "field crosses a simple field boundary",
                format!("{ENTRY}{page}{simple_open}{END}</w:fldSimple>{PAGE}{END}"),
            ),
            (
                "two ends in a simple field",
                format!("{ENTRY}{simple_open}{END}{END}</w:fldSimple>"),
            ),
            ("instruction in result", format!("{ENTRY}{page}{instruction}{PAGE}{END}")),
            ("instruction in TOC result", format!("{toc}{ENTRY}{instruction}{END}")),
            ("stray instruction", format!("{instruction}{ENTRY}")),
            (
                "text in instruction section",
                format!(
                    "{BEGIN}<w:r><w:instrText> TOC </w:instrText></w:r>{ENTRY}{SEPARATE}{ENTRY}"
                ),
            ),
            ("deleted controls", format!("<w:del>{toc}</w:del>{ENTRY}")),
            (
                "alternate controls",
                format!(
                    concat!(
                        "<mc:AlternateContent><mc:Choice Requires=\"w14\">{toc}</mc:Choice>",
                        "<mc:Fallback/></mc:AlternateContent>{ENTRY}",
                    ),
                    toc = toc,
                    ENTRY = ENTRY,
                ),
            ),
            ("REF before TOC", format!("{reference}{toc}{ENTRY}")),
            ("REF after TOC", format!("{toc}{ENTRY}{END}{reference}")),
            ("REF before entry", format!("{reference}{ENTRY}")),
            ("REF after continuation", format!("{ENTRY}{END}{reference}")),
            ("body after TOC end", format!("{toc}{ENTRY}{END}{ENTRY}")),
            ("text before TOC begin", format!("{ENTRY}{toc}{ENTRY}")),
            ("unterminated page", format!("{ENTRY}{page}{PAGE}")),
            ("nested complex page", format!("{ENTRY}{page}{page}{PAGE}{END}{END}")),
            (
                "nested simple page",
                format!("{ENTRY}{simple_open}{simple_open}{PAGE}</w:fldSimple></w:fldSimple>"),
            ),
            (
                "field crosses into link",
                format!("{ENTRY}{page}<w:hyperlink>{PAGE}{END}</w:hyperlink>"),
            ),
            (
                "field crosses out of link",
                format!("<w:hyperlink>{ENTRY}{page}{PAGE}</w:hyperlink>{END}"),
            ),
            (
                "link in page result",
                format!("{ENTRY}{page}<w:hyperlink>{PAGE}</w:hyperlink>{END}"),
            ),
            ("TOC inside link", format!("<w:hyperlink>{toc}{ENTRY}{END}</w:hyperlink>")),
            (
                "nested link",
                format!("<w:hyperlink><w:hyperlink>{ENTRY}</w:hyperlink></w:hyperlink>"),
            ),
            ("multiple ends", format!("{ENTRY}{END}{END}")),
            ("empty paragraph", String::new()),
            ("only page result", format!("{page}{PAGE}{END}")),
            ("only tab", "<w:r><w:tab/></w:r>".to_owned()),
            ("empty run", format!("{ENTRY}<w:r/>")),
            ("empty text", format!("{ENTRY}<w:r><w:t/></w:r>")),
            ("empty text with tab", format!("{ENTRY}<w:r><w:t/><w:tab/></w:r>")),
            ("only rendered break", format!("{ENTRY}<w:r><w:lastRenderedPageBreak/></w:r>")),
            ("positioned tab with text", format!("{ENTRY}<w:r><w:t>Text</w:t><w:ptab/></w:r>")),
            ("page tab only", format!("{ENTRY}{simple_open}<w:r><w:tab/></w:r></w:fldSimple>")),
            ("late paragraph properties", format!("{ENTRY}<w:pPr/>")),
            ("duplicate paragraph properties", format!("<w:pPr/><w:pPr/>{ENTRY}")),
            ("late run properties", format!("{ENTRY}<w:r><w:t>Text</w:t><w:rPr/></w:r>")),
            (
                "duplicate run properties",
                format!("{ENTRY}<w:r><w:rPr/><w:rPr/><w:t>Text</w:t></w:r>"),
            ),
            ("container text", format!("{ENTRY} stray")),
            ("nonempty marker", format!("{ENTRY}<w:bookmarkEnd><w:r/></w:bookmarkEnd>")),
            ("foreign run", format!("{ENTRY}<x:r><x:t>Text</x:t></x:r>")),
            ("foreign text", format!("{ENTRY}<w:r><x:t>Text</x:t></w:r>")),
            ("foreign properties", format!("{ENTRY}<w:r><x:rPr/><w:t>Text</w:t></w:r>")),
        ];
        for control in ["begin", "separate", "end"] {
            cases.push((
                "mixed control and text",
                format!(
                    r#"{ENTRY}<w:r><w:fldChar w:fldCharType="{control}"/><w:t>Text</w:t></w:r>"#
                ),
            ));
        }
        for name in [
            "br",
            "cr",
            "sdt",
            "smartTag",
            "customXml",
            "ins",
            "del",
            "footnoteReference",
            "delText",
        ] {
            cases.push(("unsupported element", format!("{ENTRY}<w:{name}/>")));
            cases.push(("unsupported run child", format!("{ENTRY}<w:r><w:{name}/></w:r>")));
        }
        for name in ["REF", "HYPERLINK", "SEQ", "UNKNOWN"] {
            cases.push((
                "other complex field",
                format!("{ENTRY}{}{PAGE}{END}", grammar_open(name)),
            ));
            cases.push((
                "other simple field",
                format!(r#"{ENTRY}<w:fldSimple w:instr=" {name} target ">{PAGE}</w:fldSimple>"#),
            ));
        }
        for (label, content) in cases {
            assert!(grammar_proof(&content).is_none(), "{label}: {content}");
        }
    }

    #[test]
    fn strict_toc_grammar_preserves_direct_entry_sizes() {
        for content in [
            r#"<w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t>Entry</w:t></w:r>"#,
            concat!(
                r#"<w:hyperlink><w:r><w:rPr><w:sz w:val="28"/></w:rPr>"#,
                r#"<w:sym w:font="Arial" w:char="0041"/></w:r></w:hyperlink>"#,
            ),
        ] {
            assert!(grammar_proof(content).is_none(), "{content}");
        }
        let complex_script =
            r#"<w:r><w:rPr><w:szCs w:val="28"/></w:rPr><w:t>Entry</w:t></w:r>"#;
        let proof = grammar_proof(complex_script).unwrap();
        assert!(proof.inherited(false));
        assert!(!proof.inherited(true));
    }

    #[test]
    fn an_sdt_toc_with_sized_tabs_and_pageref_results_uses_the_entry_size() {
        assert_word_toc(None, true);
    }

    #[test]
    fn split_toc_paragraphs_keep_unproven_sizes_and_match_cold() {
        let bytes = word_toc(None, false, false);
        let engine = EngineSession::new(77115);
        let env = RenderEnv {
            paragraph_style_font_sizes: BTreeMap::from([("TOC1".to_owned(), 10.0)]),
            hyperlink_style_font_sizes: BTreeMap::from([("Hyperlink".to_owned(), 14.0)]),
            doc_default_font_size: Some(10.0),
            ..Default::default()
        };
        seed_from_docx(engine.doc(), &bytes).unwrap();
        engine.with_lowered_story("body", &env, |_| ()).unwrap();
        engine
            .doc()
            .split_paragraph(&EditCtx::local("", ""), Position::new("body", 3), None)
            .unwrap();
        let resident = engine
            .with_lowered_story("body", &env, |blocks| serde_json::to_value(blocks).unwrap())
            .unwrap();
        let cold = yrs_doc_to_layout_blocks(engine.doc(), "body", &env).unwrap();
        assert_eq!(resident, serde_json::to_value(cold).unwrap());
        let mut halves = Vec::new();
        for block in resident.as_array().unwrap() {
            for run in block["runs"].as_array().into_iter().flatten() {
                if matches!(run["text"].as_str(), Some("Ent" | "ry 1")) {
                    assert_eq!(run["fontSize"], 14.0);
                    halves.push(run["text"].as_str().unwrap());
                }
            }
        }
        assert_eq!(halves, ["Ent", "ry 1"]);
    }

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
            RUN_SIZE_SEEDED_SNAPSHOTS.with(|counter| counter.set(0));
            engine
                .with_lowered_story("body", &env, |blocks| {
                    let value = serde_json::to_value(blocks).unwrap();
                    assert_eq!(value[0]["runs"][0]["fontSize"], 10.0);
                })
                .unwrap();
            RUN_SIZE_SEEDED_SNAPSHOTS.with(|counter| assert_eq!(counter.get(), 1));
            RUN_SIZE_SEEDED_SNAPSHOTS.with(|counter| counter.set(0));
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
                RUN_SIZE_SEEDED_SNAPSHOTS.with(|counter| assert_eq!(counter.get(), 0));
            }
        }
    }
}
