use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{EditCtx, EditingDoc, MergeDirection, StoryRange, seed_from_docx};
use docx_layout::types::{LayoutBlock, Run};
use serde_json::Value;
use yrs::{Any, Map, Out, ReadTxn, Text, Transact};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn run(text: &str) -> String {
    format!(r#"<w:r><w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

/// A complex field whose cached result is "7".
fn field(instruction: &str) -> String {
    format!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> {instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
    )
}

fn hyperlink_sequence() -> String {
    format!(
        r#"<w:hyperlink w:anchor="top"><w:r><w:fldChar w:fldCharType="begin" w:fldLock="true"/></w:r><w:r><w:instrText> SEQ Figure </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:fldSimple w:instr="PAGE">{}</w:fldSimple><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink>"#,
        run("1")
    )
}

fn hyperlink_sequence_with_sdt_instruction() -> String {
    hyperlink_sequence().replacen(
        "<w:r><w:instrText> SEQ Figure </w:instrText></w:r>",
        "<w:sdt><w:sdtPr/><w:sdtContent><w:r><w:instrText> SEQ Figure </w:instrText></w:r></w:sdtContent></w:sdt>",
        1,
    )
}

fn paragraph(content: &str) -> String {
    format!("<w:p>{}{content}</w:p>", run("Caption "))
}

fn text_box(content: &str) -> String {
    format!(
        r#"<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1600000" cy="228600"/><wp:wrapNone/><wp:docPr id="1" name="Box 1"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1600000" cy="228600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p>{content}</w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"#
    )
}

fn document(body: &str) -> Vec<u8> {
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document xmlns:w="{W}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>"#)),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name.to_owned(), value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

/// The results of the body's SEQ fields in lowering order, each with whether
/// it sits in a text box or shape.
fn boxed_sequence_results_in(blocks: &[LayoutBlock]) -> Vec<(bool, String)> {
    fn collect(value: &Value, boxed: bool, out: &mut Vec<(bool, String)>) {
        match value {
            Value::Object(map) if map.get("rawType").and_then(Value::as_str) == Some("SEQ") => {
                out.push((boxed, map["fallback"].as_str().unwrap().to_owned()))
            }
            Value::Object(map) => {
                let kind = map.get("kind").and_then(Value::as_str);
                let boxed = boxed || matches!(kind, Some("shape" | "textBox"));
                map.values().for_each(|value| collect(value, boxed, out))
            }
            Value::Array(items) => items.iter().for_each(|value| collect(value, boxed, out)),
            _ => {}
        }
    }
    let mut results = Vec::new();
    collect(&serde_json::to_value(blocks).unwrap(), false, &mut results);
    results
}

fn lowered_body(body: &str, show_hidden_text: bool, hydrated: bool) -> Vec<LayoutBlock> {
    let doc = EditingDoc::new(1);
    seed_from_docx(&doc, &document(body)).unwrap();
    let doc = if hydrated {
        let peer = EditingDoc::new(2);
        peer.apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        peer
    } else {
        doc
    };
    yrs_doc_to_layout_blocks(
        &doc,
        "body",
        &RenderEnv {
            show_hidden_text,
            ..RenderEnv::default()
        },
    )
    .unwrap()
}

fn boxed_sequence_results(body: &str) -> Vec<(bool, String)> {
    boxed_sequence_results_in(&lowered_body(body, false, false))
}

fn sequence_results(body: &str) -> Vec<String> {
    boxed_sequence_results(body)
        .into_iter()
        .map(|(_, result)| result)
        .collect()
}

fn assert_sticky_caption(doc: &EditingDoc, cached: &str) {
    let txn = doc.yrs_doc().transact();
    assert_eq!(
        txn.get_map("session").unwrap().get(&txn, "opaqueSequences"),
        Some(Out::Any(Any::Array(vec![Any::from("figure")].into())))
    );
    drop(txn);
    for paragraph in doc.paragraphs("body").unwrap() {
        assert!(!paragraph.properties.contains_key("opaqueSequences"));
    }
    for show_hidden_text in [false, true] {
        let blocks = yrs_doc_to_layout_blocks(
            doc,
            "body",
            &RenderEnv {
                show_hidden_text,
                ..RenderEnv::default()
            },
        )
        .unwrap();
        assert_eq!(
            boxed_sequence_results_in(&blocks),
            [(false, cached.to_owned())]
        );
    }
}

/// Each instruction with the result Word 16 shows after updating fields
/// (checked against its PDF export), or "7", the cached result, where
/// BetterOffice keeps it.
const CASES: &[(&str, &str)] = &[
    (r"SEQ Figure \* ARABIC", "1"),
    (r"SEQ Figure", "2"),
    (r"SEQ figure", "3"),
    (r"SEQ Figure \c", "3"),
    (r"SEQ Figure \r 10", "10"),
    (r"SEQ Figure \h", ""),
    (r"SEQ Figure", "12"),
    (r"SEQ Figure \h \* ARABIC", "13"),
    (r"SEQ Figure \h \* MERGEFORMAT", ""),
    (r"SEQ Figure \c \r 20", "14"),
    (r"SEQ Figure \r 20 \c", "20"),
    (r"SEQ Figure \* Arabic \* roman", "xxi"),
    (r#"SEQ "Figure""#, "22"),
    (r"SEQ Table \* ROMAN", "I"),
    (r"SEQ Table \* roman", "ii"),
    (r"SEQ Table \* ALPHABETIC", "C"),
    (r"SEQ Table \* alphabetic", "d"),
    (r"SEQ Table \* Roman", "V"),
    (r"SEQ Other \c", "0"),
    (r"SEQ Letters \r 27 \* ALPHABETIC", "AA"),
    (r"SEQ Letters \r 53 \* alphabetic", "aaa"),
    (r"SEQ Numerals \r 4000 \* ROMAN", "MMMM"),
    (r"SEQ Zero \r 0 \* ROMAN", ""),
    (r#"SEQ Figure \# "00""#, "7"),
    (r"SEQ Figure \* CardText", "7"),
];

#[test]
fn seq_fields_number_like_word() {
    let mut body: String = CASES
        .iter()
        .map(|(instruction, _)| paragraph(&field(instruction)))
        .collect();
    let mut expected: Vec<&str> = CASES.iter().map(|(_, result)| *result).collect();
    // A locked field counts and keeps its cached result.
    body += &paragraph(&field("SEQ Figure").replacen(
        r#"w:fldCharType="begin""#,
        r#"w:fldCharType="begin" w:fldLock="true""#,
        1,
    ));
    expected.push("7");
    body += &format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>{}</w:tc></w:tr></w:tbl>"#,
        paragraph(&field("SEQ Figure"))
    );
    expected.push("26");
    body +=
        &paragraph(r#"<w:fldSimple w:instr=" SEQ Figure "><w:r><w:t>7</w:t></w:r></w:fldSimple>"#);
    expected.push("27");
    // Word numbers a chapter-reset sequence, and shows an error for a SEQ
    // without a name; both keep their cached results here.
    body += &paragraph(&field(r"SEQ Chapter \s 1"));
    body += &paragraph(&field("SEQ Chapter"));
    body += &paragraph(&field("SEQ"));
    expected.extend(["7", "7", "7"]);

    assert_eq!(sequence_results(&body), expected);
}

#[test]
fn a_sequence_with_a_nested_field_keeps_its_cached_results() {
    let nested = format!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> QUOTE "</w:instrText></w:r>{}<w:r><w:instrText xml:space="preserve">" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        field("SEQ figure")
    );
    let body = [
        paragraph(&field("SEQ Figure")),
        paragraph(&nested),
        paragraph(&field("SEQ Figure")),
        paragraph(&field("SEQ Table")),
    ]
    .concat();
    assert_eq!(sequence_results(&body), ["7", "7", "1"]);
}

#[test]
fn legacy_state_without_sequence_metadata_keeps_all_cached_results() {
    let nested = format!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> QUOTE "</w:instrText></w:r>{}<w:r><w:instrText xml:space="preserve">" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        field("SEQ Figure")
    );
    let body = [
        paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>1</w:t>", 1)),
        paragraph(&nested),
        paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>3</w:t>", 1)),
    ]
    .concat();
    let doc = EditingDoc::new(1);
    seed_from_docx(&doc, &document(&body)).unwrap();
    {
        let txn = doc.yrs_doc().transact();
        let session = txn.get_map("session").unwrap();
        assert_eq!(
            session.get(&txn, "opaqueSequences"),
            Some(Out::Any(Any::Array(vec![Any::from("figure")].into())))
        );
    }
    for (has_metadata, expected) in [(true, ["1", "2", "2"]), (false, ["1", "2", "3"])] {
        {
            let mut txn = doc.yrs_doc().transact_mut();
            let session = txn.get_map("session").unwrap();
            if has_metadata {
                session.insert(
                    &mut txn,
                    "opaqueSequences",
                    Any::Array(Vec::<Any>::new().into()),
                );
            } else {
                assert!(session.remove(&mut txn, "opaqueSequences").is_some());
            }
        }
        let peer = EditingDoc::new(2);
        peer.apply_update_v1(&doc.encode_state_as_update_v1())
            .unwrap();
        for doc in [&doc, &peer] {
            let blocks = yrs_doc_to_layout_blocks(doc, "body", &RenderEnv::default()).unwrap();
            let results: Vec<_> = blocks
                .iter()
                .filter_map(|block| match block {
                    LayoutBlock::Paragraph(paragraph) => Some(&paragraph.runs),
                    _ => None,
                })
                .flatten()
                .filter_map(|run| match run {
                    Run::Field(field) => field.fallback.as_deref(),
                    _ => None,
                })
                .collect();
            assert_eq!(results, expected);
        }
    }
}

#[test]
fn nested_sequences_keep_the_seed_operations_of_a_document_without_them() {
    let seed_clocks = |name: &str| {
        let nested = format!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> QUOTE "</w:instrText></w:r>{}<w:r><w:instrText xml:space="preserve">" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
            field(&format!("{name} Figure"))
        );
        let doc = EditingDoc::new(1);
        seed_from_docx(
            &doc,
            &document(&[paragraph(&nested), paragraph(&run("After"))].concat()),
        )
        .unwrap();
        let txn = doc.yrs_doc().transact();
        txn.state_vector()
            .iter()
            .filter(|(client, _)| **client != yrs::ClientID::new(0x1_0000_05e9))
            .map(|(client, clock)| (*client, *clock))
            .collect::<std::collections::BTreeMap<_, _>>()
    };
    assert_eq!(seed_clocks("SEQ"), seed_clocks("XEQ"));
}

#[test]
fn a_body_projected_sequence_result_is_not_duplicated() {
    for result in [
        r#"<w:hyperlink w:anchor="top"><w:r><w:t>1</w:t></w:r></w:hyperlink>"#,
        r#"<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>"#,
    ] {
        let projected = field("SEQ Figure").replacen("<w:r><w:t>7</w:t></w:r>", result, 1);
        let body = format!("<w:p>{projected}</w:p>");
        for hydrated in [false, true] {
            let blocks = lowered_body(&body, false, hydrated);
            let text: String = blocks
                .iter()
                .filter_map(|block| match block {
                    LayoutBlock::Paragraph(paragraph) => Some(&paragraph.runs),
                    _ => None,
                })
                .flatten()
                .filter_map(|run| match run {
                    Run::Text(text) => Some(text.text.as_str()),
                    Run::Field(field) => field.fallback.as_deref(),
                    _ => None,
                })
                .collect();
            assert_eq!(text, "1");
            assert_eq!(boxed_sequence_results_in(&blocks), [(false, String::new())]);
        }
    }
}

#[test]
fn legacy_projected_sequence_owners_keep_cached_results() {
    for result in [
        r#"<w:hyperlink w:anchor="top"><w:r><w:t>1</w:t></w:r></w:hyperlink>"#,
        r#"<w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>"#,
    ] {
        let projected = field("SEQ Figure").replacen("<w:r><w:t>7</w:t></w:r>", result, 1);
        let body = format!("<w:p>{projected}</w:p>");
        let seeded = EditingDoc::new(1);
        seed_from_docx(&seeded, &document(&body)).unwrap();
        let legacy = EditingDoc::new(2);
        legacy
            .apply_update_v1(&seeded.encode_state_as_update_v1())
            .unwrap();
        {
            let txn = legacy.yrs_doc().transact();
            let stories = txn.get_map("stories").unwrap();
            let Some(Out::YText(story)) = stories.get(&txn, "body") else {
                panic!("body");
            };
            let mut owners = 0;
            for diff in story.diff(&txn, yrs::types::text::YChange::identity) {
                if let Out::YMap(field) = diff.insert
                    && field.get(&txn, "resultProjection").is_some()
                {
                    assert!(field.get(&txn, "nestedSequences").is_none());
                    owners += 1;
                }
            }
            assert_eq!(owners, 1);
        }
        let hydrated = EditingDoc::new(3);
        hydrated
            .apply_update_v1(&legacy.encode_state_as_update_v1())
            .unwrap();
        for doc in [&legacy, &hydrated] {
            for show_hidden_text in [false, true] {
                let blocks = yrs_doc_to_layout_blocks(
                    doc,
                    "body",
                    &RenderEnv {
                        show_hidden_text,
                        ..RenderEnv::default()
                    },
                )
                .unwrap();
                let text: String = blocks
                    .iter()
                    .filter_map(|block| match block {
                        LayoutBlock::Paragraph(paragraph) => Some(&paragraph.runs),
                        _ => None,
                    })
                    .flatten()
                    .filter_map(|run| match run {
                        Run::Text(text) => Some(text.text.as_str()),
                        Run::Field(field) => field.fallback.as_deref(),
                        _ => None,
                    })
                    .collect();
                assert_eq!(text, "1");
                assert_eq!(boxed_sequence_results_in(&blocks), [(false, String::new())]);
            }
        }
    }
}

#[test]
fn hidden_sequence_fields_keep_visible_cached_results() {
    let hidden = r#"<w:fldSimple w:instr="SEQ Figure"><w:r><w:rPr><w:vanish/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>"#;
    let visible = r#"<w:fldSimple w:instr="SEQ Figure"><w:r><w:t>2</w:t></w:r></w:fldSimple>"#;
    for boxed in [false, true] {
        let body = if boxed {
            format!("<w:p>{}</w:p><w:p>{visible}</w:p>", text_box(hidden))
        } else {
            format!("<w:p>{hidden}</w:p><w:p>{visible}</w:p>")
        };
        for hydrated in [false, true] {
            for show_hidden_text in [false, true] {
                let blocks = lowered_body(&body, show_hidden_text, hydrated);
                let mut expected = Vec::new();
                if show_hidden_text {
                    expected.push((boxed, "1".to_owned()));
                }
                expected.push((false, "2".to_owned()));
                assert_eq!(boxed_sequence_results_in(&blocks), expected);
            }
        }
    }
}

#[test]
fn hidden_fields_keep_nested_hyperlink_sequences_opaque() {
    let hidden = format!(
        r#"<w:fldSimple w:instr="QUOTE"><w:r><w:rPr><w:vanish/></w:rPr><w:t>7</w:t></w:r>{}</w:fldSimple>"#,
        hyperlink_sequence_with_sdt_instruction()
    );
    for boxed in [false, true] {
        let hidden = if boxed {
            text_box(&hidden)
        } else {
            hidden.clone()
        };
        let body = [
            format!("<w:p>{hidden}</w:p>"),
            paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>2</w:t>", 1)),
            paragraph(&field("SEQ Table")),
        ]
        .concat();
        for hydrated in [false, true] {
            for show_hidden_text in [false, true] {
                assert_eq!(
                    boxed_sequence_results_in(&lowered_body(&body, show_hidden_text, hydrated)),
                    [(false, "2"), (false, "1")].map(|(boxed, result)| (boxed, result.to_owned()))
                );
            }
        }
    }
}

#[test]
fn hidden_text_boxes_keep_sequences_with_typed_or_raw_hyperlink_fields_opaque() {
    for hidden in [
        field("SEQ Figure"),
        hyperlink_sequence(),
        hyperlink_sequence_with_sdt_instruction(),
    ] {
        let hidden_box = text_box(&hidden).replacen(
            "<w:r><w:drawing>",
            "<w:r><w:rPr><w:vanish/></w:rPr><w:drawing>",
            1,
        );
        let body = [
            format!("<w:p>{hidden_box}</w:p>"),
            paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>2</w:t>", 1)),
            paragraph(&field("SEQ Table")),
        ]
        .concat();
        for hydrated in [false, true] {
            assert_eq!(
                boxed_sequence_results_in(&lowered_body(&body, false, hydrated)),
                [(false, "2"), (false, "1")].map(|(boxed, result)| (boxed, result.to_owned()))
            );
        }
    }
}

#[test]
fn hidden_text_boxes_preserve_stale_visible_captions() {
    let hidden_box = text_box(&field("SEQ Figure")).replacen(
        "<w:r><w:drawing>",
        "<w:r><w:rPr><w:vanish/></w:rPr><w:drawing>",
        1,
    );
    let body = [
        format!("<w:p>{hidden_box}</w:p>"),
        paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>9</w:t>", 1)),
    ]
    .concat();
    for hydrated in [false, true] {
        for show_hidden_text in [false, true] {
            let expected = if show_hidden_text {
                vec![(true, "1".to_owned()), (false, "2".to_owned())]
            } else {
                vec![(false, "9".to_owned())]
            };
            assert_eq!(
                boxed_sequence_results_in(&lowered_body(&body, show_hidden_text, hydrated)),
                expected
            );
        }
    }
}

#[test]
fn shape_fields_preserve_visible_cached_result_fragments() {
    for hidden_text in ["", "0"] {
        let hidden_result = if hidden_text.is_empty() {
            String::new()
        } else {
            format!("<w:t>{hidden_text}</w:t>")
        };
        let cached = format!(
            r#"<w:fldSimple w:instr="SEQ Figure" w:fldLock="true"><w:r><w:rPr><w:vanish/></w:rPr>{hidden_result}</w:r><w:r><w:t>1</w:t></w:r></w:fldSimple>"#
        );
        let body = [
            format!("<w:p>{}</w:p>", text_box(&cached)),
            paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>9</w:t>", 1)),
        ]
        .concat();
        for hydrated in [false, true] {
            for show_hidden_text in [false, true] {
                let boxed = if show_hidden_text {
                    format!("{hidden_text}1")
                } else {
                    "1".to_owned()
                };
                let caption = if !show_hidden_text && !hidden_text.is_empty() {
                    "9"
                } else {
                    "2"
                };
                assert_eq!(
                    boxed_sequence_results_in(&lowered_body(&body, show_hidden_text, hydrated)),
                    [(true, boxed), (false, caption.to_owned())]
                );
            }
        }
    }
}

#[test]
fn hidden_shape_runs_omit_line_breaks() {
    let content = format!(
        r#"{}<w:r><w:rPr><w:vanish/></w:rPr><w:br/></w:r>{}"#,
        run("Before"),
        run("After")
    );
    let body = format!("<w:p>{}</w:p>", text_box(&content));
    for hydrated in [false, true] {
        for show_hidden_text in [false, true] {
            let blocks = lowered_body(&body, show_hidden_text, hydrated);
            let runs = blocks
                .iter()
                .filter_map(|block| match block {
                    LayoutBlock::Shape(shape) => shape.inner_text.as_ref(),
                    _ => None,
                })
                .flatten()
                .flat_map(|paragraph| &paragraph.runs)
                .collect::<Vec<_>>();
            let text: String = runs
                .iter()
                .filter_map(|run| match run {
                    Run::Text(text) => Some(text.text.as_str()),
                    _ => None,
                })
                .collect();
            assert_eq!(text, "BeforeAfter");
            assert_eq!(
                runs.iter()
                    .filter(|run| matches!(run, Run::LineBreak(_)))
                    .count(),
                usize::from(show_hidden_text)
            );
        }
    }
}

#[test]
fn hidden_sequences_preserve_stale_cached_results_only_for_the_affected_sequence() {
    let body = [
        r#"<w:p><w:fldSimple w:instr="SEQ Figure"><w:r><w:rPr><w:vanish/></w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p>"#.to_owned(),
        paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>9</w:t>", 1)),
        paragraph(&field("SEQ Table")),
    ]
    .concat();
    for hydrated in [false, true] {
        assert_eq!(
            boxed_sequence_results_in(&lowered_body(&body, false, hydrated)),
            [(false, "9"), (false, "1")].map(|(boxed, result)| (boxed, result.to_owned()))
        );
        assert_eq!(
            boxed_sequence_results_in(&lowered_body(&body, true, hydrated)),
            [(false, "1"), (false, "2"), (false, "1")]
                .map(|(boxed, result)| (boxed, result.to_owned()))
        );
    }
}

#[test]
fn a_text_box_with_a_projected_nested_sequence_keeps_cached_results() {
    let nested = format!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> QUOTE "</w:instrText></w:r>{}<w:r><w:instrText xml:space="preserve">" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="top">{}</w:hyperlink><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>1</w:t>", 1),
        run("1")
    );
    let body = [
        format!("<w:p>{}</w:p>", text_box(&nested)),
        paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>2</w:t>", 1)),
        paragraph(&field("SEQ Table")),
    ]
    .concat();
    assert_eq!(
        boxed_sequence_results(&body),
        [(false, "2"), (false, "1")].map(|(boxed, result)| (boxed, result.to_owned()))
    );
}

#[test]
fn a_text_box_sequence_too_deep_to_lower_keeps_its_sequence_cached() {
    let deep = (0..8).fold(
        field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>2</w:t>", 1),
        |content, _| format!("<w:sdt><w:sdtPr/><w:sdtContent>{content}</w:sdtContent></w:sdt>"),
    );
    let body = [
        paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>1</w:t>", 1)),
        format!("<w:p>{}</w:p>", text_box(&deep)),
        paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>3</w:t>", 1)),
    ]
    .concat();
    for hydrated in [false, true] {
        assert_eq!(
            boxed_sequence_results_in(&lowered_body(&body, false, hydrated)),
            [(false, "1"), (false, "3")].map(|(boxed, result)| (boxed, result.to_owned()))
        );
    }
}

#[test]
fn a_text_box_with_a_projected_sequence_keeps_cached_results() {
    let boxed = format!(
        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Figure </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="top">{}</w:hyperlink><w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
        run("1")
    );
    let body = [
        format!("<w:p>{}</w:p>", text_box(&boxed)),
        paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>2</w:t>", 1)),
        paragraph(&field("SEQ Table")),
    ]
    .concat();
    assert_eq!(
        boxed_sequence_results(&body),
        [(false, "2"), (false, "1")].map(|(boxed, result)| (boxed, result.to_owned()))
    );
}

#[test]
fn a_text_box_with_a_hyperlink_sequence_keeps_cached_results() {
    let body = [
        format!("<w:p>{}</w:p>", text_box(&hyperlink_sequence())),
        paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>2</w:t>", 1)),
        paragraph(&field("SEQ Table")),
    ]
    .concat();
    assert_eq!(
        boxed_sequence_results(&body),
        [(false, "2"), (false, "1")].map(|(boxed, result)| (boxed, result.to_owned()))
    );
}

#[test]
fn a_body_hyperlink_sequence_keeps_cached_results() {
    let body = [
        paragraph(&hyperlink_sequence()),
        paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>2</w:t>", 1)),
    ]
    .concat();
    assert_eq!(sequence_results(&body), ["2"]);
}

#[test]
fn a_text_box_with_a_hyperlink_sequence_instruction_in_an_sdt_keeps_cached_results() {
    let body = [
        format!(
            "<w:p>{}</w:p>",
            text_box(&hyperlink_sequence_with_sdt_instruction())
        ),
        paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>2</w:t>", 1)),
        paragraph(&field("SEQ Table")),
    ]
    .concat();
    assert_eq!(
        boxed_sequence_results(&body),
        [(false, "2"), (false, "1")].map(|(boxed, result)| (boxed, result.to_owned()))
    );
}

#[test]
fn a_body_hyperlink_sequence_instruction_in_an_sdt_keeps_cached_results() {
    let body = [
        paragraph(&hyperlink_sequence_with_sdt_instruction()),
        paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>2</w:t>", 1)),
    ]
    .concat();
    assert_eq!(sequence_results(&body), ["2"]);
}

#[test]
fn a_body_typed_hyperlink_sequence_in_an_sdt_keeps_cached_results() {
    let hyperlink = format!(
        r#"<w:hyperlink w:anchor="top"><w:sdt><w:sdtPr/><w:sdtContent>{}</w:sdtContent></w:sdt></w:hyperlink>"#,
        field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>1</w:t>", 1)
    );
    let body = [
        paragraph(&hyperlink),
        paragraph(&field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>2</w:t>", 1)),
    ]
    .concat();
    for hydrated in [false, true] {
        for show_hidden_text in [false, true] {
            assert_eq!(
                boxed_sequence_results_in(&lowered_body(&body, show_hidden_text, hydrated)),
                [(false, "2".to_owned())]
            );
        }
    }
}

#[test]
fn deleting_a_hyperlink_sequence_keeps_document_opacity_on_state_only_peers() {
    let hyperlink = format!(
        r#"<w:hyperlink w:anchor="top"><w:fldSimple w:instr="SEQ Figure" w:fldLock="true">{}</w:fldSimple></w:hyperlink>"#,
        run("1")
    );
    let body = format!(
        "<w:p>{}</w:p><w:p>{hyperlink}</w:p><w:p>{}</w:p>",
        run("x"),
        field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>9</w:t>", 1)
    );
    let doc = EditingDoc::new(1);
    seed_from_docx(&doc, &document(&body)).unwrap();
    let paragraphs = doc.paragraphs("body").unwrap();
    let end = doc
        .paragraph_mark_position(&paragraphs[1].para_id)
        .unwrap()
        .index;
    doc.delete_range(&EditCtx::local("", ""), StoryRange::new("body", 1, end))
        .unwrap();
    let peer = EditingDoc::new(2);
    peer.apply_update_v1(&doc.encode_state_as_update_v1())
        .unwrap();
    for doc in [&doc, &peer] {
        assert_eq!(doc.paragraphs("body").unwrap()[0].text, "x");
        assert_sticky_caption(doc, "9");
    }
}

#[test]
fn merging_a_hyperlink_sdt_keeps_document_opacity_on_state_only_peers() {
    let hyperlink = format!(
        r#"<w:hyperlink w:anchor="top"><w:sdt><w:sdtPr/><w:sdtContent>{}</w:sdtContent></w:sdt></w:hyperlink>"#,
        field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>1</w:t>", 1)
    );
    let body = format!(
        "<w:p>{}</w:p><w:p>{hyperlink}</w:p><w:p>{}</w:p>",
        run("x"),
        field("SEQ Figure").replacen("<w:t>7</w:t>", "<w:t>2</w:t>", 1)
    );
    let doc = EditingDoc::new(1);
    seed_from_docx(&doc, &document(&body)).unwrap();
    let first = doc.paragraphs("body").unwrap()[0].para_id.clone();
    doc.merge_paragraphs(&EditCtx::local("", ""), &first, MergeDirection::Forward)
        .unwrap();
    let peer = EditingDoc::new(2);
    peer.apply_update_v1(&doc.encode_state_as_update_v1())
        .unwrap();
    for doc in [&doc, &peer] {
        assert_eq!(doc.paragraphs("body").unwrap().len(), 2);
        assert_sticky_caption(doc, "2");
    }
}

#[test]
fn a_text_box_anchored_at_a_paragraph_start_counts_before_the_paragraph() {
    let body = [
        paragraph(&field("SEQ Figure")),
        format!(
            "<w:p>{}{}{}</w:p>",
            text_box(&(run("Boxed ") + &field("SEQ Figure"))),
            run("Caption "),
            field("SEQ Figure")
        ),
        paragraph(&field("SEQ Figure")),
    ]
    .concat();
    assert_eq!(
        boxed_sequence_results(&body),
        [(false, "1"), (true, "2"), (false, "3"), (false, "4")]
            .map(|(boxed, result)| (boxed, result.to_owned()))
    );
}
