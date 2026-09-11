use ooxml_fidelity::wml::{diff_digests, semantic_digest};
use ooxml_fidelity::{Part, XmlLimits, parse_part, roundtrip_findings, structural_fingerprint};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";

fn part(root: &str, attributes: &str, body: &str) -> Vec<Part> {
    vec![(if root == "hdr" { "word/header1.xml" } else { "word/document.xml" }.to_owned(),
        format!(r#"<w:{root} xmlns:w="{W}" xmlns:mc="{MC}" xmlns:w14="{W14}" xmlns:x="urn:custom" {attributes}>{body}</w:{root}>"#).into_bytes())]
}

fn assert_both_detect(before: &[Part], after: &[Part]) {
    let fingerprint = |parts: &[Part]| {
        structural_fingerprint(
            &parse_part(&parts[0].1, &parts[0].0, &XmlLimits::default()).unwrap(),
        )
    };
    assert_ne!(
        fingerprint(before),
        fingerprint(after),
        "fingerprint missed compatibility loss"
    );
    assert!(
        !diff_digests(
            &semantic_digest(before).unwrap(),
            &semantic_digest(after).unwrap()
        )
        .is_empty(),
        "digest missed compatibility loss"
    );
    let findings = roundtrip_findings(before, after).unwrap();
    assert!(
        findings
            .iter()
            .any(|line| line.starts_with("fingerprint differs:")),
        "{findings:?}"
    );
    assert!(
        findings.iter().any(|line| line.starts_with("digest:")),
        "{findings:?}"
    );
}

#[test]
fn custom_ignorable_cannot_disappear_while_markup_remains() {
    assert_both_detect(
        &part("document", r#"mc:Ignorable="w14 x""#, "<x:marker/>"),
        &part("document", r#"mc:Ignorable="w14""#, "<x:marker/>"),
    );
}

#[test]
fn standard_ignorable_additions_are_directional() {
    let before = part("document", r#"mc:Ignorable="x""#, "<x:marker/>");
    let after = part("document", r#"mc:Ignorable="w14 x""#, "<x:marker/>");
    assert!(roundtrip_findings(&before, &after).unwrap().is_empty());
    assert_both_detect(&after, &before);
}

#[test]
fn header_root_ignorable_cannot_disappear() {
    assert_both_detect(
        &part("hdr", r#"mc:Ignorable="w14 x""#, "<x:marker/>"),
        &part("hdr", "", "<x:marker/>"),
    );
}

#[test]
fn requires_rebinding_changes_choice_semantics() {
    let before = part(
        "document",
        r#"xmlns:w16="urn:unsupported" xmlns:spare="urn:unsupported""#,
        r#"<mc:AlternateContent><mc:Choice Requires="w16"><w:p><w:r><w:t>CHOICE</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>FALLBACK</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent>"#,
    );
    let after = vec![(
        before[0].0.clone(),
        String::from_utf8(before[0].1.clone())
            .unwrap()
            .replace(
                r#"xmlns:w16="urn:unsupported""#,
                r#"xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml""#,
            )
            .into_bytes(),
    )];
    assert_both_detect(&before, &after);
}

#[test]
fn qname_values_follow_scoped_bindings() {
    for attribute in [
        r#"mc:Ignorable="a""#,
        r#"mc:ProcessContent="a:item""#,
        r#"xsi:type="a:Kind""#,
    ] {
        let before = part(
            "document",
            "",
            &format!(
                r#"<x:item xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:a="urn:one" xmlns:b="urn:two" {attribute}/>"#
            ),
        );
        let after = vec![(
            before[0].0.clone(),
            String::from_utf8(before[0].1.clone())
                .unwrap()
                .replace(
                    r#"xmlns:a="urn:one" xmlns:b="urn:two""#,
                    r#"xmlns:a="urn:two" xmlns:b="urn:one""#,
                )
                .into_bytes(),
        )];
        assert_both_detect(&before, &after);
    }
}

#[test]
fn qname_aliases_and_prefix_set_order_are_insignificant() {
    let before = part(
        "document",
        r#"mc:Ignorable="x w14 x""#,
        r#"<mc:AlternateContent><mc:Choice Requires="x"><x:item mc:ProcessContent="x:child"/></mc:Choice></mc:AlternateContent>"#,
    );
    let after = vec![(
        before[0].0.clone(),
        String::from_utf8(before[0].1.clone())
            .unwrap()
            .replace("x:", "alias:")
            .replace("xmlns:x=", "xmlns:alias=")
            .replace("x w14 x", "w14 alias")
            .replace(r#"Requires="x""#, r#"Requires="alias""#)
            .into_bytes(),
    )];
    assert!(roundtrip_findings(&before, &after).unwrap().is_empty());
}

#[test]
fn foreign_root_attributes_are_digested() {
    assert_both_detect(
        &part("document", r#"x:flag="kept""#, "<w:body/>"),
        &part("document", "", "<w:body/>"),
    );
}

#[test]
fn unbound_qname_prefix_is_rejected() {
    let xml = part(
        "document",
        "",
        r#"<mc:AlternateContent><mc:Choice Requires="missing"/></mc:AlternateContent>"#,
    );
    assert!(parse_part(&xml[0].1, &xml[0].0, &XmlLimits::default()).is_err());
}
