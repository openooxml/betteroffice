use docx_parse::ParseLimits;
use docx_parse::block::PREVIEW_MIN_BLOCKS;
use docx_parse::s9::{
    S9ParseOptions, media_table_parts, parse_docx_s9_preview_from_parts,
    parse_docx_s9_preview_from_parts_full_dom,
    parse_docx_s9_preview_from_parts_full_dom_with_budget,
    parse_docx_s9_preview_from_parts_with_budget,
    parse_docx_s9_preview_with_media_table_with_budget, parse_docx_s9_wire_with_limits,
};

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#;

fn relationships(entries: &[(&str, &str, &str)]) -> Vec<u8> {
    let mut xml =
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#
            .to_owned();
    for (id, kind, target) in entries {
        xml.push_str(&format!(
            r#"<Relationship Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{target}"/>"#
        ));
    }
    xml.push_str("</Relationships>");
    xml.into_bytes()
}

fn package(body: &str) -> Vec<(String, Vec<u8>)> {
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            relationships(&[("rDoc", "officeDocument", "word/document.xml")]),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            relationships(&[
                ("rHeader", "header", "header1.xml"),
                ("rComments", "comments", "comments.xml"),
            ]),
        ),
        (
            "word/document.xml".to_owned(),
            format!("<w:document {NS}><w:body>{body}</w:body></w:document>").into_bytes(),
        ),
        (
            "word/header1.xml".to_owned(),
            format!(r#"<w:hdr {NS}><w:p w14:paraId="00000001"><w:r><w:t>header</w:t></w:r></w:p></w:hdr>"#).into_bytes(),
        ),
        (
            "word/comments.xml".to_owned(),
            format!(r#"<w:comments {NS}><w:comment w:id="0" w:author="A"><w:p><w:r><w:t>comment</w:t></w:r></w:p></w:comment></w:comments>"#).into_bytes(),
        ),
    ];
    ooxml_opc::unzip_parts(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap()
}

fn paragraph(index: usize, properties: &str, runs: &str) -> String {
    format!(
        r#"<w:p w14:paraId="{:08X}">{properties}{runs}<w:r><w:t>paragraph {index}</w:t></w:r></w:p>"#,
        index % 100 + 1,
    )
}

fn paragraphs(mut render: impl FnMut(usize) -> String) -> String {
    (0..620).map(&mut render).collect()
}

fn compare(body: &str, refused: bool) {
    let parts = package(body);
    for source_ordinals in [false, true] {
        for blocks in [1, 5, 50, 200] {
            let options = S9ParseOptions {
                source_ordinals,
                determinism_seed: Some("7".repeat(64)),
                ..S9ParseOptions::default()
            };
            let limits = ParseLimits::default();
            let legacy =
                parse_docx_s9_preview_from_parts(&parts, blocks, options.clone(), &limits).unwrap();
            let expected =
                parse_docx_s9_preview_from_parts_full_dom(&parts, blocks, options.clone(), &limits)
                    .unwrap();
            assert_eq!(legacy.is_none(), refused, "blocks={blocks}");
            let legacy_json = serde_json::to_vec(&legacy).unwrap();
            assert_eq!(
                legacy_json,
                serde_json::to_vec(&expected).unwrap(),
                "blocks={blocks}, source_ordinals={source_ordinals}"
            );
            for paragraph_budget in [None, Some(32), Some(256)] {
                let actual = parse_docx_s9_preview_from_parts_with_budget(
                    &parts,
                    blocks,
                    options.clone(),
                    &limits,
                    paragraph_budget,
                )
                .unwrap();
                let expected = parse_docx_s9_preview_from_parts_full_dom_with_budget(
                    &parts,
                    blocks,
                    options.clone(),
                    &limits,
                    paragraph_budget,
                )
                .unwrap();
                assert_eq!(actual.is_none(), refused);
                assert_eq!(
                    actual, expected,
                    "blocks={blocks}, budget={paragraph_budget:?}"
                );
                if paragraph_budget.is_none() {
                    let envelope = actual.as_ref().map(|(envelope, budget_stopped)| {
                        assert!(!*budget_stopped);
                        envelope
                    });
                    assert_eq!(serde_json::to_vec(&envelope).unwrap(), legacy_json);
                }
            }
        }
    }
}

#[test]
fn plain_paragraphs_match_the_full_dom_preview() {
    compare(&paragraphs(|index| paragraph(index, "", "")), false);
}

#[test]
fn tables_and_wrapped_blocks_match_the_full_dom_preview() {
    let tables = paragraphs(|index| {
        let p = paragraph(index, "", "");
        if index % 4 == 0 {
            format!("<w:tbl><w:tr><w:tc>{p}</w:tc></w:tr></w:tbl>{p}")
        } else {
            p
        }
    });
    compare(&tables, false);
    let wrappers = paragraphs(|index| {
        let p = paragraph(index, "", "");
        match index % 4 {
            0 => format!("<w:sdt><w:sdtPr/><w:sdtContent>{p}</w:sdtContent></w:sdt>"),
            1 => format!("<w:customXml><w:smartTag>{p}{p}</w:smartTag></w:customXml>"),
            2 => format!("<w:smartTag>{p}</w:smartTag>"),
            _ => p,
        }
    });
    compare(&wrappers, false);
}

#[test]
fn fields_ending_inside_and_beyond_the_prefix_match_the_full_dom_preview() {
    for end in [260, 590] {
        let body = paragraphs(|index| {
            let runs = match index {
                190 => {
                    r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> IF 1 = 1 </w:instrText></w:r>"#
                }
                191 => r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
                index if index == end => r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
                _ => "",
            };
            paragraph(index, "", runs)
        });
        compare(&body, false);
    }
}

#[test]
fn later_section_properties_and_final_headers_match_the_full_dom_preview() {
    let sections = paragraphs(|index| {
        let properties = match index {
            100 => r#"<w:pPr><w:sectPr><w:pgSz w:w="10000" w:h="15000"/></w:sectPr></w:pPr>"#,
            500 => r#"<w:pPr><w:sectPr><w:pgSz w:w="20000" w:h="25000"/></w:sectPr></w:pPr>"#,
            _ => "",
        };
        paragraph(index, properties, "")
    });
    compare(&sections, false);
    let final_header = format!(
        r#"{}<w:sectPr><w:headerReference w:type="default" r:id="rHeader"/><w:pgSz w:w="18000" w:h="24000"/></w:sectPr>"#,
        paragraphs(|index| paragraph(index, "", "")),
    );
    compare(&final_header, false);
}

#[test]
fn a_section_inside_a_later_sdt_matches_the_full_dom_preview() {
    let body = paragraphs(|index| {
        if index == 500 {
            let p = paragraph(
                index,
                r#"<w:pPr><w:sectPr><w:pgSz w:w="18000" w:h="24000"/></w:sectPr></w:pPr>"#,
                "",
            );
            format!("<w:sdt><w:sdtContent>{p}</w:sdtContent></w:sdt>")
        } else {
            paragraph(index, "", "")
        }
    });
    compare(&body, false);
}

#[test]
fn a_late_page_anchor_refuses_both_preview_paths() {
    let body = paragraphs(|index| {
        let runs = if index == 550 {
            r#"<w:r><w:drawing><wp:anchor><wp:positionV relativeFrom="page"/></wp:anchor></w:drawing></w:r>"#
        } else {
            ""
        };
        paragraph(index, "", runs)
    });
    compare(&body, true);
}

#[test]
fn bookmarks_and_comment_ranges_spanning_the_cut_match_the_full_dom_preview() {
    let body = paragraphs(|index| {
        let runs = match index {
            0 => r#"<w:bookmarkStart w:id="1" w:name="whole"/><w:commentRangeStart w:id="0"/>"#,
            4 => r#"<w:bookmarkStart w:id="2" w:name="later"/>"#,
            590 => {
                r#"<w:bookmarkEnd w:id="2"/><w:bookmarkEnd w:id="1"/><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>"#
            }
            _ => "",
        };
        paragraph(index, "", runs)
    });
    compare(&body, false);
}

#[test]
fn malformed_tails_keep_the_full_dom_error() {
    for tail in ["<w:p>", "<w:p a=\"<\"/>", "<w:p>&unknown;</w:p>"] {
        let body = format!("{}{tail}", paragraphs(|index| paragraph(index, "", "")));
        let parts = package(&body);
        let options = S9ParseOptions {
            determinism_seed: Some("7".repeat(64)),
            ..S9ParseOptions::default()
        };
        let limits = ParseLimits::default();
        let actual =
            parse_docx_s9_preview_from_parts(&parts, 5, options.clone(), &limits).unwrap_err();
        let expected =
            parse_docx_s9_preview_from_parts_full_dom(&parts, 5, options, &limits).unwrap_err();
        assert_eq!(actual, expected);
    }
}

#[test]
fn xml_limits_still_count_the_removed_tail() {
    let parts = package(&paragraphs(|index| paragraph(index, "", "")));
    let options = S9ParseOptions {
        determinism_seed: Some("7".repeat(64)),
        ..S9ParseOptions::default()
    };
    let mut limits = Vec::new();
    for max_xml_events in [100, 4300, 4350, 4400, 10000] {
        limits.push(ParseLimits {
            max_xml_events,
            ..ParseLimits::default()
        });
    }
    for max_xml_text_bytes in [100, 40000, 46000, 100000] {
        limits.push(ParseLimits {
            max_xml_text_bytes,
            ..ParseLimits::default()
        });
    }
    for max_xml_bytes in [100, 50000, 100000] {
        limits.push(ParseLimits {
            max_xml_bytes,
            ..ParseLimits::default()
        });
    }
    for limits in limits {
        let actual = parse_docx_s9_preview_from_parts(&parts, 5, options.clone(), &limits)
            .map(|value| serde_json::to_value(value).unwrap());
        let expected =
            parse_docx_s9_preview_from_parts_full_dom(&parts, 5, options.clone(), &limits)
                .map(|value| serde_json::to_value(value).unwrap());
        assert_eq!(actual, expected, "{limits:?}");
    }
}

fn table(rows: usize, cells: usize, cell_paragraphs: usize) -> String {
    let content = (0..cell_paragraphs)
        .map(|index| paragraph(index, "", ""))
        .collect::<String>();
    let cells = format!("<w:tc>{content}</w:tc>").repeat(cells);
    format!(
        "<w:tbl>{}</w:tbl>",
        format!("<w:tr>{cells}</w:tr>").repeat(rows)
    )
}

fn weighted_preview(
    body: &str,
    blocks: usize,
    budget: usize,
) -> (docx_parse::S9WireEnvelope, bool) {
    preview_with_budget(&package(body), blocks, Some(budget))
}

fn preview_with_budget(
    parts: &[(String, Vec<u8>)],
    blocks: usize,
    budget: Option<usize>,
) -> (docx_parse::S9WireEnvelope, bool) {
    let options = S9ParseOptions {
        determinism_seed: Some("7".repeat(64)),
        ..S9ParseOptions::default()
    };
    let limits = ParseLimits::default();
    let streaming = parse_docx_s9_preview_from_parts_with_budget(
        parts,
        blocks,
        options.clone(),
        &limits,
        budget,
    )
    .unwrap()
    .unwrap();
    let dom = parse_docx_s9_preview_from_parts_full_dom_with_budget(
        parts,
        blocks,
        options.clone(),
        &limits,
        budget,
    )
    .unwrap()
    .unwrap();
    assert_eq!(streaming, dom);
    let (parts, media) = media_table_parts(&ooxml_opc::rezip_parts(parts).unwrap().into()).unwrap();
    let media_preview = parse_docx_s9_preview_with_media_table_with_budget(
        &parts, &media, blocks, options, &limits, budget,
    )
    .unwrap()
    .unwrap();
    assert_eq!(streaming, media_preview);
    streaming
}

fn dense_tables_with_page_count(trailing_paragraphs: usize) -> Vec<(String, Vec<u8>)> {
    let body = format!(
        r#"{}{}<w:sectPr><w:headerReference w:type="default" r:id="rHeader"/></w:sectPr>"#,
        table(1, 1, 10).repeat(40),
        paragraph(40, "", "").repeat(trailing_paragraphs),
    );
    let mut parts = package(&body);
    let header = parts
        .iter_mut()
        .find(|(path, _)| path == "word/header1.xml")
        .unwrap();
    header.1 = format!(
        r#"<w:hdr {NS}><w:p><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>99</w:t></w:r></w:fldSimple></w:p></w:hdr>"#,
    ).into_bytes();
    parts
}

fn paragraph_count(content: &[docx_parse::BlockContent]) -> usize {
    content
        .iter()
        .map(|block| match block {
            docx_parse::BlockContent::Paragraph(_) => 1,
            docx_parse::BlockContent::Table(table) => table
                .rows
                .iter()
                .flat_map(|row| &row.cells)
                .map(|cell| paragraph_count(&cell.content))
                .sum(),
            docx_parse::BlockContent::BlockSdt(sdt) => paragraph_count(&sdt.content),
            docx_parse::BlockContent::RawXml(_) => 0,
        })
        .sum()
}

#[test]
fn a_short_dense_body_with_numpages_ignores_the_budget() {
    let parts = dense_tables_with_page_count(0);
    let (preview, stopped) = preview_with_budget(&parts, 200, Some(256));
    let (legacy, _) = preview_with_budget(&parts, 200, None);
    assert!(!stopped);
    let content = &preview.document.package.document.content;
    assert_eq!(content.len(), 40);
    assert!(
        content
            .iter()
            .all(|block| matches!(block, docx_parse::BlockContent::Table(_)))
    );
    assert_eq!(paragraph_count(content), 400);
    assert_eq!(
        serde_json::to_vec(&preview).unwrap(),
        serde_json::to_vec(&legacy).unwrap()
    );
}

#[test]
fn a_dense_body_with_more_than_the_block_limit_uses_the_budget() {
    let parts = dense_tables_with_page_count(170);
    let (preview, stopped) = preview_with_budget(&parts, 200, Some(256));
    let (legacy, _) = preview_with_budget(&parts, 200, None);
    assert!(stopped);
    let content = &preview.document.package.document.content;
    let legacy_content = &legacy.document.package.document.content;
    assert_eq!(content.len(), PREVIEW_MIN_BLOCKS);
    assert_eq!(legacy_content.len(), 200);
    assert_eq!(paragraph_count(content), 320);
    assert_eq!(paragraph_count(legacy_content), 560);
    assert!(paragraph_count(content) < paragraph_count(legacy_content));
}

#[test]
fn the_budget_requires_strictly_more_than_the_block_limit() {
    for blocks in [199, 200, 201] {
        let body = format!("{}<w:sectPr/>", table(1, 1, 10).repeat(blocks));
        let parts = package(&body);
        let (preview, stopped) = preview_with_budget(&parts, 200, Some(256));
        assert_eq!(stopped, blocks > 200);
        if blocks <= 200 {
            let (legacy, _) = preview_with_budget(&parts, 200, None);
            assert_eq!(preview, legacy);
        } else {
            assert_eq!(
                preview.document.package.document.content.len(),
                PREVIEW_MIN_BLOCKS
            );
        }
    }
}

#[test]
fn fields_that_hold_the_legacy_cut_to_the_body_end_disable_the_budget() {
    for blocks in [210, 450] {
        for close in [false, true] {
            for end in ["", "<w:sectPr/>"] {
                let mut body = (0..blocks)
                    .map(|index| {
                        if index < 40 {
                            return table(1, 1, 10);
                        }
                        let runs = match index {
                            190 => r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> IF 1 = 1 </w:instrText></w:r>"#,
                            191 => r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
                            index if close && index == blocks - 1 => r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
                            _ => "",
                        };
                        paragraph(index, "", runs)
                    })
                    .collect::<String>();
                body.push_str(end);
                let parts = package(&body);
                let (preview, stopped) = preview_with_budget(&parts, 200, Some(256));
                let (legacy, _) = preview_with_budget(&parts, 200, None);
                assert!(!stopped);
                assert_eq!(preview.document.package.document.content.len(), blocks);
                assert_eq!(preview, legacy);
            }
        }
    }
}

#[test]
fn dense_tables_stop_at_the_first_complete_block_after_the_minimum() {
    let mut body = (0..60)
        .map(|index| format!("{}{}", table(4, 3, 2), paragraph(index, "", "")))
        .collect::<String>();
    body.push_str(&paragraph(0, "", "").repeat(620));
    for (budget, expected_blocks) in [(256, PREVIEW_MIN_BLOCKS), (430, 35)] {
        let (preview, stopped) = weighted_preview(&body, 200, budget);
        assert!(stopped);
        let content = &preview.document.package.document.content;
        assert_eq!(content.len(), expected_blocks);
        for block in content.iter().step_by(2) {
            let docx_parse::BlockContent::Table(table) = block else {
                panic!("table");
            };
            assert_eq!(table.rows.len(), 4);
            for row in &table.rows {
                assert_eq!(row.cells.len(), 3);
                for cell in &row.cells {
                    assert_eq!(cell.content.len(), 2);
                    assert!(
                        cell.content
                            .iter()
                            .all(|block| matches!(block, docx_parse::BlockContent::Paragraph(_)))
                    );
                }
            }
        }
    }
}

#[test]
fn an_exhausted_budget_keeps_trailing_non_block_children() {
    for markers in [
        "",
        r#"<w:bookmarkStart w:id="1" w:name="tail"/><w:bookmarkEnd w:id="1"/><w:permStart w:id="2"/><w:permEnd w:id="2"/><w:proofErr w:type="spellStart"/><w:proofErr w:type="spellEnd"/><w:customXmlInsRangeStart w:id="3"/><w:customXmlInsRangeEnd w:id="3"/><w:moveFromRangeStart w:id="4"/><w:moveFromRangeEnd w:id="4"/>"#,
    ] {
        let body = format!(
            r#"{}{markers}<w:sectPr><w:headerReference w:type="default" r:id="rHeader"/><w:pgSz w:w="18000" w:h="24000"/></w:sectPr>"#,
            table(1, 1, 10).repeat(32),
        );
        let (preview, stopped) = weighted_preview(&body, 200, 256);
        assert!(!stopped);
        let content = &preview.document.package.document.content;
        assert_eq!(content.len(), 32);
        assert!(
            content
                .iter()
                .all(|block| matches!(block, docx_parse::BlockContent::Table(_)))
        );
        let whole_body = !stopped && content.len() < 200;
        assert!(whole_body);
        let full = parse_docx_s9_wire_with_limits(
            &ooxml_opc::rezip_parts(&package(&body)).unwrap(),
            S9ParseOptions {
                determinism_seed: Some("7".repeat(64)),
                ..S9ParseOptions::default()
            },
            &ParseLimits::default(),
        )
        .unwrap();
        assert_eq!(
            serde_json::to_vec(&preview).unwrap(),
            serde_json::to_vec(&full).unwrap()
        );
    }
}

#[test]
fn an_exhausted_budget_stops_before_another_block() {
    for trailing_block in [
        paragraph(32, "", ""),
        r#"<ext:block xmlns:ext="urn:preview-test"/>"#.to_owned(),
    ] {
        let body = format!(
            r#"{}{trailing_block}{}<w:sectPr><w:headerReference w:type="default" r:id="rHeader"/><w:pgSz w:w="18000" w:h="24000"/></w:sectPr>"#,
            table(1, 1, 10).repeat(32),
            paragraph(33, "", "").repeat(170),
        );
        let (preview, stopped) = weighted_preview(&body, 200, 256);
        assert!(stopped);
        let content = &preview.document.package.document.content;
        assert_eq!(content.len(), 32);
        let whole_body = !stopped && content.len() < 200;
        assert!(!whole_body);
        let expected = parse_docx_s9_preview_from_parts(
            &package(&body),
            32,
            S9ParseOptions {
                determinism_seed: Some("7".repeat(64)),
                ..S9ParseOptions::default()
            },
            &ParseLimits::default(),
        )
        .unwrap()
        .unwrap();
        assert_eq!(preview, expected);
        let properties =
            serde_json::to_value(preview.document.package.document.final_section_properties)
                .unwrap();
        assert!(properties.to_string().contains("18000"));
    }
}

#[test]
fn nested_table_and_sdt_paragraphs_charge_the_budget() {
    let p = paragraph(0, "", "");
    let nested = format!(
        "<w:tbl><w:tr><w:tc>{p}{}{p}</w:tc></w:tr></w:tbl>",
        table(1, 2, 2)
    );
    let sdt = format!(
        "<w:sdt><w:sdtPr/><w:sdtContent>{p}{}{p}</w:sdtContent></w:sdt>",
        table(1, 2, 2)
    );
    for block in [nested, sdt] {
        let body = format!("{}{block}{}", p.repeat(32), p.repeat(620));
        let (preview, stopped) = weighted_preview(&body, 200, 38);
        assert!(stopped);
        assert_eq!(preview.document.package.document.content.len(), 33);
    }
}

#[test]
fn text_box_paragraphs_charge_the_budget() {
    let p = paragraph(0, "", "");
    let drawing = format!(
        r#"<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:txbx><w:txbxContent>{}</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#,
        p.repeat(5),
    );
    let body = format!(
        "{}{}{}",
        p.repeat(32),
        paragraph(32, "", &drawing),
        p.repeat(620)
    );
    let (preview, stopped) = weighted_preview(&body, 200, 38);
    assert!(stopped);
    assert_eq!(preview.document.package.document.content.len(), 33);
}

#[test]
fn a_field_spanning_the_budget_closes_before_the_cut() {
    for end in [40, 450] {
        let body = (0..620)
            .map(|index| {
                let runs = match index {
                    30 => {
                        r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> IF 1 = 1 </w:instrText></w:r>"#
                    }
                    31 => r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
                    index if index == end => r#"<w:r><w:fldChar w:fldCharType="end"/></w:r>"#,
                    _ => "",
                };
                paragraph(index, "", runs)
            })
            .collect::<String>();
        let (preview, stopped) = weighted_preview(&body, 200, 32);
        assert!(stopped);
        assert_eq!(preview.document.package.document.content.len(), end + 1);
        let parts = package(&body);
        let expected = parse_docx_s9_preview_from_parts(
            &parts,
            end + 1,
            S9ParseOptions {
                determinism_seed: Some("7".repeat(64)),
                ..S9ParseOptions::default()
            },
            &ParseLimits::default(),
        )
        .unwrap()
        .unwrap();
        assert_eq!(preview, expected);
    }
}

#[test]
fn a_giant_first_table_is_kept_whole() {
    let body = format!("{}{}", table(100, 3, 2), paragraph(0, "", "").repeat(620));
    for (blocks, expected_blocks, budget_stopped) in
        [(1, 1, false), (200, PREVIEW_MIN_BLOCKS, true)]
    {
        let (preview, stopped) = weighted_preview(&body, blocks, 256);
        assert_eq!(stopped, budget_stopped);
        let content = &preview.document.package.document.content;
        assert_eq!(content.len(), expected_blocks);
        let docx_parse::BlockContent::Table(table) = &content[0] else {
            panic!("table");
        };
        assert_eq!(table.rows.len(), 100);
        assert!(table.rows.iter().all(|row| {
            row.cells.len() == 3 && row.cells.iter().all(|cell| cell.content.len() == 2)
        }));
    }
}

#[test]
fn sparse_text_still_stops_at_the_block_limit() {
    let body = (0..300)
        .map(|index| paragraph(index, "", ""))
        .collect::<String>();
    let (preview, stopped) = weighted_preview(&body, 200, 256);
    assert!(!stopped);
    assert_eq!(preview.document.package.document.content.len(), 200);
    let parts = package(&body);
    let legacy = parse_docx_s9_preview_from_parts(
        &parts,
        200,
        S9ParseOptions {
            determinism_seed: Some("7".repeat(64)),
            ..S9ParseOptions::default()
        },
        &ParseLimits::default(),
    )
    .unwrap()
    .unwrap();
    assert_eq!(
        serde_json::to_vec(&preview).unwrap(),
        serde_json::to_vec(&legacy).unwrap()
    );
}

#[test]
fn weighted_cuts_keep_the_next_section_properties() {
    for section_index in [80, 500] {
        let body = (0..620)
            .map(|index| {
                let properties = if index == section_index {
                    r#"<w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rHeader"/><w:pgSz w:w="18000" w:h="24000"/></w:sectPr></w:pPr>"#
                } else {
                    ""
                };
                format!("{}{}", table(4, 3, 2), paragraph(index, properties, ""))
            })
            .collect::<String>();
        let (preview, stopped) = weighted_preview(&body, 200, 256);
        assert!(stopped);
        assert_eq!(
            preview.document.package.document.content.len(),
            PREVIEW_MIN_BLOCKS
        );
        let properties =
            serde_json::to_value(preview.document.package.document.final_section_properties)
                .unwrap();
        assert!(properties.to_string().contains("18000"));
    }
}

#[test]
fn empty_blocks_still_have_weight_and_zero_budget_obeys_the_minimum() {
    let body = "<w:sdt><w:sdtContent/></w:sdt>".repeat(620);
    for (budget, expected_blocks) in [(0, PREVIEW_MIN_BLOCKS), (40, 40)] {
        let (preview, stopped) = weighted_preview(&body, 200, budget);
        assert!(stopped);
        assert_eq!(
            preview.document.package.document.content.len(),
            expected_blocks
        );
    }
}
