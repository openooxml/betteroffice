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
    #[derive(Default)]
    struct Field {
        instruction: String,
        result: bool,
    }

    impl Field {
        fn entry_result(&self) -> bool {
            self.result
                && !self
                    .instruction
                    .split_whitespace()
                    .next()
                    .is_some_and(|instruction| instruction.eq_ignore_ascii_case("PAGEREF"))
        }
    }

    fn visit(
        node: &XmlElement,
        fields: &mut Vec<Field>,
        found: &mut bool,
        direct_cs: &mut bool,
    ) -> bool {
        if node.matches_name("w", "fldSimple") {
            fields.push(Field {
                instruction: node.attribute(Some("w"), "instr").unwrap_or_default().to_owned(),
                result: true,
            });
            let valid = node
                .child_elements()
                .all(|child| visit(child, fields, found, direct_cs));
            fields.pop();
            return valid;
        }
        if node.matches_name("w", "r") {
            let mut control = false;
            let mut text = false;
            for child in node.child_elements() {
                if child.matches_name("w", "fldChar") {
                    control = true;
                    match child.attribute(Some("w"), "fldCharType") {
                        Some("begin") => fields.push(Field::default()),
                        Some("separate") => {
                            if fields.is_empty() {
                                fields.push(Field::default());
                            }
                            fields.last_mut().unwrap().result = true;
                        }
                        Some("end") => {
                            fields.pop();
                        }
                        _ => return false,
                    }
                } else if child.matches_name("w", "instrText")
                    || child.matches_name("w", "delInstrText")
                {
                    control = true;
                    if let Some(field) = fields.last_mut() {
                        field.instruction.push_str(&child.text_content());
                    }
                } else if child.matches_name("w", "t") || child.matches_name("w", "delText") {
                    text |= !child.text_content().is_empty();
                } else if ["sym", "softHyphen", "noBreakHyphen"]
                    .iter()
                    .any(|name| child.matches_name("w", name))
                {
                    text = true;
                }
            }
            if control || !text || !fields.iter().all(Field::entry_result) {
                return true;
            }
            *found = true;
            let properties = node.child("w", "rPr");
            if properties.is_some_and(|properties| properties.child("w", "sz").is_some()) {
                return false;
            }
            *direct_cs |=
                properties.is_some_and(|properties| properties.child("w", "szCs").is_some());
            return true;
        }
        node.child_elements()
            .all(|child| visit(child, fields, found, direct_cs))
    }

    let mut found = false;
    let mut direct_cs = false;
    (visit(paragraph, &mut Vec::new(), &mut found, &mut direct_cs) && found).then_some(
        ParagraphSizes {
            fingerprint,
            direct_cs,
        },
    )
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

    #[test]
    fn source_entry_sizes_exclude_tabs_and_pageref_results() {
        use docx_parse::xml::{ParseBudget, ParseLimits, parse_xml};

        let sized = r#"<w:rPr><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>"#;
        let complex = format!(
            concat!(
                r#"<w:r>{sized}<w:fldChar w:fldCharType="begin"/></w:r>"#,
                r#"<w:r>{sized}<w:instrText> PAGEREF target </w:instrText></w:r>"#,
                r#"<w:r>{sized}<w:fldChar w:fldCharType="separate"/></w:r>"#,
                r#"<w:r>{sized}<w:t>3</w:t></w:r>"#,
                r#"<w:r>{sized}<w:fldChar w:fldCharType="end"/></w:r>"#,
            ),
            sized = sized,
        );
        let mut cases = vec![
            (format!("<w:r>{sized}<w:tab/><w:ptab/></w:r>"), true),
            (
                format!(
                    concat!(
                        r#"<w:fldSimple w:instr=" PAGEREF target "><w:r>{sized}"#,
                        "<w:t>3</w:t></w:r></w:fldSimple>",
                    ),
                    sized = sized,
                ),
                true,
            ),
            (complex.clone(), true),
            (format!("<w:r>{sized}<w:t>Direct entry</w:t></w:r>"), false),
            (
                format!(r#"<w:r>{sized}<w:sym w:font="Arial" w:char="0041"/></w:r>"#),
                false,
            ),
        ];
        for instruction in ["TOC", "REF", "HYPERLINK", "SEQ", "UNKNOWN"] {
            cases.push((
                format!(
                    concat!(
                        r#"<w:fldSimple w:instr=" {instruction} target "><w:r>{sized}"#,
                        "<w:t>3</w:t></w:r></w:fldSimple>",
                    ),
                    instruction = instruction,
                    sized = sized,
                ),
                false,
            ));
            cases.push((complex.replace("PAGEREF", instruction), false));
        }
        for linked in [false, true] {
            for (content, eligible) in &cases {
                let entry = concat!(
                    r#"<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>"#,
                    r#"<w:t>Entry</w:t></w:r>"#,
                );
                let content = if linked {
                    format!("<w:hyperlink>{entry}{content}</w:hyperlink>")
                } else {
                    format!("{entry}{content}")
                };
                let xml = format!(
                    concat!(
                        "<w:p xmlns:w=\"",
                        "http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
                        "{content}</w:p>",
                    ),
                    content = content,
                );
                let limits = ParseLimits::default();
                let mut budget = ParseBudget::new(&limits);
                let document = parse_xml(xml.as_bytes(), "word/document.xml", &mut budget).unwrap();
                let proof = paragraph_sizes(document.root().unwrap(), [0; 32]);
                assert_eq!(proof.is_some(), *eligible, "{content}");
                if let Some(proof) = proof {
                    assert!(proof.inherited(true), "{content}");
                }
            }
        }
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
