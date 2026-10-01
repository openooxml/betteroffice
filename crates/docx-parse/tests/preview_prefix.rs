use docx_parse::ParseLimits;
use docx_parse::s9::{
    S9ParseOptions, parse_docx_s9_preview_from_parts, parse_docx_s9_preview_from_parts_full_dom,
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
            let actual =
                parse_docx_s9_preview_from_parts(&parts, blocks, options.clone(), &limits).unwrap();
            let expected =
                parse_docx_s9_preview_from_parts_full_dom(&parts, blocks, options, &limits)
                    .unwrap();
            assert_eq!(actual.is_none(), refused, "blocks={blocks}");
            assert_eq!(
                serde_json::to_value(actual).unwrap(),
                serde_json::to_value(expected).unwrap(),
                "blocks={blocks}, source_ordinals={source_ordinals}"
            );
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
