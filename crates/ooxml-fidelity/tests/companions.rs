use ooxml_fidelity::{Part, roundtrip_findings};

const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

fn rels(target: &str, mode: &str) -> Vec<Part> {
    vec![("word/_rels/document.xml.rels".to_owned(), format!(r#"<Relationships xmlns="{RELS}"><Relationship Id="rId5" Type="hyperlink" TargetMode="{mode}" Target="{target}"/></Relationships>"#).into_bytes())]
}

#[test]
fn retargeting_a_companion_relationship_is_a_finding() {
    let findings = roundtrip_findings(
        &rels("commentsIds.xml", "Internal"),
        &rels("commentsIds-other.xml", "Internal"),
    )
    .unwrap();
    assert_eq!(findings.len(), 2, "{findings:?}");
    assert!(
        findings
            .iter()
            .all(|line| line.starts_with("digest: word/_rels/document.xml.rels entry"))
    );
}

#[test]
fn companion_words_in_hyperlinks_do_not_allow_changes() {
    let findings = roundtrip_findings(
        &rels("https://example.org/commentsExtended", "External"),
        &rels("https://example.org/commentsIds", "External"),
    )
    .unwrap();
    assert_eq!(findings.len(), 2, "{findings:?}");
}

#[test]
fn similarly_named_custom_parts_are_not_companions() {
    let findings = roundtrip_findings(
        &[],
        &[(
            "customXml/commentsIds-payload.xml".to_owned(),
            b"<data/>".to_vec(),
        )],
    )
    .unwrap();
    assert_eq!(
        findings,
        [
            "part added: customXml/commentsIds-payload.xml",
            "digest: customXml/commentsIds-payload.xml | absent -> present"
        ]
    );
}

#[test]
fn genuine_companion_additions_are_allowed_but_removals_are_not() {
    for name in ["commentsIds", "commentsExtended", "commentsExtensible"] {
        let before = vec![
            (
                "word/_rels/document.xml.rels".to_owned(),
                format!(r#"<Relationships xmlns="{RELS}"/>"#).into_bytes(),
            ),
            (
                "[Content_Types].xml".to_owned(),
                format!(r#"<Types xmlns="{CT}"/>"#).into_bytes(),
            ),
        ];
        let after = vec![("word/_rels/document.xml.rels".to_owned(), format!(r#"<Relationships xmlns="{RELS}"><Relationship Id="added" Type="companion" Target="{name}.xml"/></Relationships>"#).into_bytes()),
            ("[Content_Types].xml".to_owned(), format!(r#"<Types xmlns="{CT}"><Override PartName="/word/{name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{name}+xml"/></Types>"#).into_bytes()),
            (format!("word/{name}.xml"), b"<companion/>".to_vec())];
        assert!(roundtrip_findings(&before, &after).unwrap().is_empty());
        let removed = roundtrip_findings(&after, &before).unwrap();
        assert!(
            removed
                .iter()
                .any(|line| line.starts_with("digest: [Content_Types].xml entry")),
            "{removed:?}"
        );
        assert!(
            removed
                .iter()
                .any(|line| line.starts_with("digest: word/_rels/document.xml.rels entry")),
            "{removed:?}"
        );
    }
}

#[test]
fn companion_words_do_not_hide_lost_paragraph_identity() {
    let before = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="11111111" label="commentsIds"/></w:body></w:document>"#;
    let after = before.replace(r#" w14:paraId="11111111""#, "");
    let findings = roundtrip_findings(
        &[("word/document.xml".to_owned(), before.as_bytes().to_vec())],
        &[("word/document.xml".to_owned(), after.into_bytes())],
    )
    .unwrap();
    assert_eq!(findings.len(), 1, "{findings:?}");
    assert!(findings[0].starts_with("digest: word/document.xml block[0].attributes"));
}
