use ooxml_fidelity::{XmlLimits, parse_part, roundtrip_findings, structural_fingerprint};

fn fingerprint(xml: &str) -> String {
    structural_fingerprint(&parse_part(xml.as_bytes(), "test.xml", &XmlLimits::default()).unwrap())
}

#[test]
fn literal_line_endings_normalize_before_character_references() {
    for text in ["a\r\nb", "a\rb", "a\nb"] {
        assert_eq!(
            fingerprint(&format!("<t>{text}</t>")),
            fingerprint("<t>a\nb</t>")
        );
        assert_eq!(
            fingerprint(&format!("<t><![CDATA[{text}]]></t>")),
            fingerprint("<t>a\nb</t>")
        );
    }
    assert_ne!(fingerprint("<t>a&#13;b</t>"), fingerprint("<t>a\rb</t>"));
}

#[test]
fn pretty_printed_empty_property_containers_compare_equal() {
    for name in ["tcPr", "rPr", "pPr", "tblPr", "trPr", "sectPr", "tblGrid"] {
        let wrap = |body| {
            format!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>{body}</w:p></w:body></w:document>"#
            )
        };
        let before = wrap(format!("<w:{name}>\n  </w:{name}>"));
        let after = wrap(format!("<w:{name}/>"));
        assert_eq!(fingerprint(&before), fingerprint(&after));
        assert!(
            roundtrip_findings(
                &[("word/document.xml".to_owned(), before.into_bytes())],
                &[("word/document.xml".to_owned(), after.into_bytes())]
            )
            .unwrap()
            .is_empty()
        );
    }
}

#[test]
fn nonbreaking_spaces_and_preserved_whitespace_remain_significant() {
    assert_ne!(
        fingerprint("<p><a/>\u{a0}<b/></p>"),
        fingerprint("<p><a/><b/></p>")
    );
    assert_ne!(
        fingerprint(
            r#"<w:tcPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xml:space="preserve"> </w:tcPr>"#
        ),
        fingerprint(
            r#"<w:tcPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xml:space="preserve"/>"#
        )
    );
}

#[test]
fn property_child_order_is_significant() {
    let first = r#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:spacing w:after="120"/><w:ind w:left="100"/></w:pPr>"#;
    let second = r#"<w:pPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:ind w:left="100"/><w:spacing w:after="120"/></w:pPr>"#;
    assert_ne!(fingerprint(first), fingerprint(second));
}
