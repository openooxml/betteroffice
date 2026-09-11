use ooxml_fidelity::{Part, roundtrip_findings};

fn unparseable_pair(bytes: Vec<u8>) {
    let before = vec![("customXml/item1.xml".to_owned(), bytes)];
    assert!(roundtrip_findings(&before, &before).unwrap().is_empty());
    let after = vec![(
        "customXml/item1.xml".to_owned(),
        b"<data>changed</data>".to_vec(),
    )];
    assert_eq!(
        roundtrip_findings(&before, &after).unwrap(),
        [
            "unparseable part: customXml/item1.xml",
            "unmodelled part rewritten: customXml/item1.xml"
        ]
    );
    assert_eq!(
        roundtrip_findings(&after, &before).unwrap(),
        [
            "unparseable part: customXml/item1.xml",
            "unmodelled part rewritten: customXml/item1.xml"
        ]
    );
}

#[test]
fn unchanged_utf16_is_preserved_and_changed_utf16_is_reported() {
    let mut bytes = vec![0xff, 0xfe];
    for unit in "<?xml version=\"1.0\" encoding=\"UTF-16\"?><data>kept</data>".encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }
    unparseable_pair(bytes);
}

#[test]
fn unchanged_doctype_is_preserved_and_changed_doctype_is_reported() {
    unparseable_pair(b"<!DOCTYPE data><data>kept</data>".to_vec());
}

#[test]
fn unchanged_entity_is_preserved_and_changed_entity_is_reported() {
    unparseable_pair(b"<data>a&undefined;b</data>".to_vec());
}

#[test]
fn a_dropped_part_is_reported() {
    assert_eq!(
        roundtrip_findings(&[("word/media/image.png".to_owned(), vec![1, 2, 3])], &[]).unwrap(),
        ["part dropped: word/media/image.png"]
    );
}

#[test]
fn changed_non_xml_bytes_are_reported() {
    assert_eq!(
        roundtrip_findings(
            &[("word/media/image.png".to_owned(), vec![1, 2, 3])],
            &[("word/media/image.png".to_owned(), vec![1, 2, 4])]
        )
        .unwrap(),
        ["bytes differ: word/media/image.png"]
    );
}

#[test]
fn the_report_detects_census_loss_when_entry_sets_match() {
    let entry = r#"<Relationship Id="rId1" Type="link" Target="file.xml"/>"#;
    let parts = |entries: String| -> Vec<Part> {
        vec![("word/_rels/document.xml.rels".to_owned(), format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{entries}</Relationships>"#).into_bytes())]
    };
    assert_eq!(
        roundtrip_findings(&parts(entry.repeat(2)), &parts(entry.to_owned())).unwrap(),
        [
            "census loss: http://schemas.openxmlformats.org/package/2006/relationships:Relationship 2 -> 1"
        ]
    );
}

#[test]
fn text_outside_the_root_is_reported() {
    let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#;
    let before = vec![("word/document.xml".to_owned(), xml.as_bytes().to_vec())];
    for corrupted in [format!("{xml}CORRUPTED"), format!("CORRUPTED{xml}")] {
        assert_eq!(
            roundtrip_findings(
                &before,
                &[("word/document.xml".to_owned(), corrupted.into_bytes())]
            )
            .unwrap(),
            ["unparseable part: word/document.xml"]
        );
    }
}
