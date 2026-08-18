//! Round-trip fidelity gates: open → save → reopen through the package
//! oracles. The oracles parse bytes with their own reader (`ooxml-fidelity`),
//! never this crate's model, so a model gap cannot hide a save loss. Known
//! defects are pinned with exact equality: fixing one without lowering its
//! ceiling fails too, so no regression can hide behind the headroom.
//! Governed by `openspec/changes/docx-word-fidelity/specs/fidelity-oracles`.

use betteroffice_docx::Document;
use ooxml_fidelity::wml::{Difference, diff_digests, semantic_digest};
use ooxml_fidelity::{
    XmlLimits, element_census, is_xml_part, losses, parse_part, structural_fingerprint,
};

type Parts = Vec<(String, Vec<u8>)>;

fn sample_docx() -> Vec<u8> {
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/document.xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p w14:paraId="11111111"><w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">Hello </w:t></w:r><w:bookmarkStart w:id="1" w:name="mark"/><w:r><w:t>DOCX</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr><w:p w14:paraId="22222222"><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p w14:paraId="33333333"><w:r><w:tab/><w:t>Second</w:t><w:br/><w:t>section</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>"#.to_vec(),
        ),
        (
            "word/header1.xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:p w14:paraId="44444444"><w:r><w:t>Native header</w:t></w:r></w:p></w:hdr>"#.to_vec(),
        ),
        (
            "word/media/image1.png".to_owned(),
            vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x01, 0x02, 0x03],
        ),
        (
            "customXml/item1.xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><data xmlns="urn:custom"><value>kept</value></data>"#.to_vec(),
        ),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn with_document_xml(package: &[u8], edit: impl Fn(String) -> String) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(package).unwrap();
    for (name, bytes) in &mut parts {
        if name == "word/document.xml" {
            *bytes = edit(String::from_utf8(bytes.clone()).unwrap()).into_bytes();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn parts_of(package: &[u8]) -> Parts {
    ooxml_opc::unzip_parts(package).unwrap()
}

fn save_unedited(package: &[u8]) -> Vec<u8> {
    Document::open(package).unwrap().save().unwrap()
}

/// One line per violated rule; an unedited round trip must report nothing.
fn roundtrip_report(before: &Parts, after: &Parts) -> Vec<String> {
    let mut report = Vec::new();
    let limits = XmlLimits::default();
    let after_bytes: std::collections::BTreeMap<&str, &Vec<u8>> = after
        .iter()
        .map(|(name, bytes)| (name.as_str(), bytes))
        .collect();
    for (name, bytes) in before {
        match after_bytes.get(name.as_str()) {
            None => report.push(format!("part dropped: {name}")),
            Some(saved) if is_xml_part(name) => {
                let original = parse_part(bytes, name, &limits).unwrap();
                let reopened = parse_part(saved, name, &limits).unwrap();
                if structural_fingerprint(&original) != structural_fingerprint(&reopened) {
                    report.push(format!("fingerprint differs: {name}"));
                }
            }
            Some(saved) => {
                if *saved != bytes {
                    report.push(format!("bytes differ: {name}"));
                }
            }
        }
    }
    for (name, _) in after {
        if !before.iter().any(|(existing, _)| existing == name) {
            report.push(format!("part added: {name}"));
        }
    }
    for loss in losses(
        &element_census(before).unwrap(),
        &element_census(after).unwrap(),
    ) {
        report.push(format!(
            "census loss: {}:{} {} -> {}",
            loss.namespace, loss.local, loss.before, loss.after
        ));
    }
    for difference in digest_diff(before, after) {
        report.push(format!(
            "digest: {} | {} -> {}",
            difference.path, difference.before, difference.after
        ));
    }
    report
}

fn digest_diff(before: &Parts, after: &Parts) -> Vec<Difference> {
    diff_digests(
        &semantic_digest(before).unwrap(),
        &semantic_digest(after).unwrap(),
    )
}

#[test]
fn an_unedited_round_trip_passes_both_oracles_and_the_byte_rules() {
    let original = sample_docx();
    let saved = save_unedited(&original);
    assert_eq!(
        roundtrip_report(&parts_of(&original), &parts_of(&saved)),
        Vec::<String>::new()
    );
}

#[test]
fn save_reopen_save_is_byte_identical() {
    let first = save_unedited(&sample_docx());
    let second = save_unedited(&first);
    assert_eq!(first, second);
}

#[test]
fn an_edited_round_trip_loses_nothing_beyond_its_footprint() {
    let original = sample_docx();
    let mut document = Document::open(&original).unwrap();
    document
        .replace_paragraph_text("22222222", "Edited natively")
        .unwrap();
    let edited = document.save().unwrap();
    let before = parts_of(&original);
    let after = parts_of(&edited);
    assert_eq!(
        losses(
            &element_census(&before).unwrap(),
            &element_census(&after).unwrap()
        ),
        vec![]
    );
    assert_eq!(
        digest_diff(&before, &after),
        vec![Difference {
            path: "word/document.xml block[1].text".to_owned(),
            before: "Cell text".to_owned(),
            after: "Edited natively".to_owned(),
        }]
    );
}

#[test]
fn the_guard_has_teeth() {
    let saved = save_unedited(&sample_docx());
    let broken = with_document_xml(&saved, |xml| {
        xml.replace(r#"<w:bookmarkStart w:id="1" w:name="mark"/>"#, "")
    });
    let before = parts_of(&saved);
    let after = parts_of(&broken);
    let census_losses = losses(
        &element_census(&before).unwrap(),
        &element_census(&after).unwrap(),
    );
    assert_eq!(census_losses.len(), 1);
    assert_eq!(census_losses[0].local, "bookmarkStart");
    assert_ne!(digest_diff(&before, &after), vec![]);
}

/// KNOWN DEFECT (`pkg.unknown-xml`): a foreign-namespace element inside a
/// modelled part does not survive save. Ceiling: exactly three lost elements
/// on this fixture. Implementing generic preservation must lower it to zero.
#[test]
fn an_unknown_element_in_a_modelled_part_is_a_known_defect() {
    let original = with_document_xml(&sample_docx(), |xml| {
        xml.replace(
            "<w:body>",
            r#"<w:body><cx:block xmlns:cx="urn:custom-x"><cx:y cx:z="1"/></cx:block>"#,
        )
        .replace(
            "<w:r><w:t>DOCX</w:t></w:r>",
            r#"<cx:tag xmlns:cx="urn:custom-x" cx:k="v"/><w:r><w:t>DOCX</w:t></w:r>"#,
        )
    });
    let saved = save_unedited(&original);
    let lost: Vec<String> = losses(
        &element_census(&parts_of(&original)).unwrap(),
        &element_census(&parts_of(&saved)).unwrap(),
    )
    .into_iter()
    .map(|loss| format!("{}:{}", loss.namespace, loss.local))
    .collect();
    assert_eq!(
        lost,
        vec![
            "urn:custom-x:block".to_owned(),
            "urn:custom-x:tag".to_owned(),
            "urn:custom-x:y".to_owned(),
        ]
    );
}

/// KNOWN DEFECT (`pkg.unknown-xml`): a foreign-namespace attribute on a known
/// element does not survive save. The census is blind to attributes; the
/// digest names the exact loss. Ceiling: exactly one difference.
#[test]
fn an_unknown_attribute_on_a_known_element_is_a_known_defect() {
    let original = with_document_xml(&sample_docx(), |xml| {
        xml.replace(
            r#"<w:p w14:paraId="33333333">"#,
            r#"<w:p w14:paraId="33333333" cx:flag="1" xmlns:cx="urn:custom-x">"#,
        )
    });
    let saved = save_unedited(&original);
    let differences = digest_diff(&parts_of(&original), &parts_of(&saved));
    assert_eq!(differences.len(), 1);
    assert_eq!(differences[0].path, "word/document.xml block[2].attributes");
    assert!(differences[0].before.contains("{urn:custom-x}flag=1"));
    assert!(!differences[0].after.contains("{urn:custom-x}flag=1"));
}
