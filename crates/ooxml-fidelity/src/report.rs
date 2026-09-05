//! Round-trip findings: both oracles, the census, and the byte rules over an
//! original/saved package pair, with exactly the declared normalizations
//! tolerated. Every allowance is fail-closed: a difference that is not
//! precisely the tolerated addition stays a finding.

use std::collections::{BTreeMap, BTreeSet};

use crate::census::{element_census, losses};
use crate::error::FidelityError;
use crate::fingerprint::{structural_fingerprint, structural_fingerprint_excluding};
use crate::registry::{
    COMPANION_PART_NAMES, ComparisonMode, MANAGED_IDENTITY_ATTRIBUTES, comparison_mode,
    is_modelled_xml_part, normalize_root_ignorable,
};
use crate::wml::{Difference, diff_digests, semantic_digest};
use crate::xml::{XmlLimits, parse_part};
use crate::{Part, is_xml_part};

/// One line per violated rule; an unedited round trip must report nothing.
pub fn roundtrip_findings(before: &[Part], after: &[Part]) -> Result<Vec<String>, FidelityError> {
    let mut findings = Vec::new();
    let allowed_entries = companion_entries(after)?;
    let limits = XmlLimits::default();
    require_exact("non-xml-part-bytes")?;
    require_exact("unmodelled-xml-part-bytes")?;
    let after_bytes: BTreeMap<&str, &Vec<u8>> = after
        .iter()
        .map(|(name, bytes)| (name.as_str(), bytes))
        .collect();
    for (name, bytes) in before {
        match after_bytes.get(name.as_str()) {
            None => findings.push(format!("part dropped: {name}")),
            // Relationship and content-type parts carry no authored order;
            // their oracle is the digest's sorted identity set, not the
            // ordered fingerprint ("relationship-and-content-type-order").
            Some(_) if name == "[Content_Types].xml" || name.ends_with(".rels") => {}
            // Byte rule 2: an XML part the engine does not model is never
            // re-emitted, so no lexical difference is allowed either.
            Some(saved) if is_xml_part(name) && !is_modelled_xml_part(name) => {
                if *saved != bytes {
                    findings.push(format!("unmodelled part rewritten: {name}"));
                }
            }
            Some(saved) if is_xml_part(name) => {
                let original = parse_part(bytes, name, &limits)?;
                let mut reopened = parse_part(saved, name, &limits)?;
                normalize_root_ignorable(&original.attributes, &mut reopened.attributes);
                // Identity-stripped equality classifies the difference as the
                // "paragraph-id-minting" normalization; a lost or changed
                // identity still surfaces through the digest below.
                if structural_fingerprint(&original) != structural_fingerprint(&reopened)
                    && structural_fingerprint_excluding(&original, MANAGED_IDENTITY_ATTRIBUTES)
                        != structural_fingerprint_excluding(&reopened, MANAGED_IDENTITY_ATTRIBUTES)
                {
                    findings.push(format!("fingerprint differs: {name}"));
                }
            }
            Some(saved) => {
                if *saved != bytes {
                    findings.push(format!("bytes differ: {name}"));
                }
            }
        }
    }
    for (name, _) in after {
        if !before.iter().any(|(existing, _)| existing == name) && !is_companion_part(name) {
            findings.push(format!("part added: {name}"));
        }
    }
    for loss in losses(&element_census(before)?, &element_census(after)?) {
        findings.push(format!(
            "census loss: {}:{} {} -> {}",
            loss.namespace, loss.local, loss.before, loss.after
        ));
    }
    for difference in diff_digests(&semantic_digest(before)?, &semantic_digest(after)?) {
        if allowed_difference(&difference, &allowed_entries) {
            continue;
        }
        findings.push(format!(
            "digest: {} | {} -> {}",
            difference.path, difference.before, difference.after
        ));
    }
    Ok(findings)
}

/// The byte rules compare bytes; no registry entry may loosen them.
fn require_exact(artifact: &str) -> Result<(), FidelityError> {
    match comparison_mode(artifact)? {
        ComparisonMode::Exact => Ok(()),
        mode => Err(FidelityError::LoosenedComparison {
            artifact: artifact.to_owned(),
            mode: format!("{mode:?}"),
        }),
    }
}

/// The "comment-companion-parts" normalization.
fn is_companion_part(name: &str) -> bool {
    COMPANION_PART_NAMES.contains(&name)
}

fn companion_entries(parts: &[Part]) -> Result<BTreeSet<(String, String)>, FidelityError> {
    let mut allowed = BTreeSet::new();
    for (part, bytes) in parts {
        if part != "[Content_Types].xml" && !part.ends_with(".rels") {
            continue;
        }
        let root = parse_part(bytes, part, &XmlLimits::default())?;
        for entry in root.element_children() {
            let target = if part == "[Content_Types].xml"
                && entry.is(
                    "http://schemas.openxmlformats.org/package/2006/content-types",
                    "Override",
                ) {
                entry
                    .attribute("", "PartName")
                    .and_then(|name| name.strip_prefix('/'))
                    .map(str::to_owned)
            } else if part.ends_with(".rels")
                && entry.is(
                    "http://schemas.openxmlformats.org/package/2006/relationships",
                    "Relationship",
                )
                && entry
                    .attribute("", "TargetMode")
                    .is_none_or(|mode| mode == "Internal")
            {
                entry
                    .attribute("", "Target")
                    .and_then(|target| relationship_target(part, target))
            } else {
                None
            };
            if target.as_deref().is_some_and(is_companion_part) {
                allowed.insert((format!("{part} entry"), crate::wml::element_token(entry)));
            }
        }
    }
    Ok(allowed)
}

fn relationship_target(part: &str, target: &str) -> Option<String> {
    let directory = if part == "_rels/.rels" {
        ""
    } else {
        part.rsplit_once("/_rels/")?.0
    };
    let path = if target.starts_with('/') || directory.is_empty() {
        target.to_owned()
    } else {
        format!("{directory}/{target}")
    };
    let mut segments = Vec::new();
    for segment in path.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                segments.pop()?;
            }
            _ => segments.push(segment),
        }
    }
    Some(segments.join("/"))
}

fn allowed_difference(
    difference: &Difference,
    allowed_entries: &BTreeSet<(String, String)>,
) -> bool {
    if difference.before == "absent"
        && ((difference.after == "present" && is_companion_part(&difference.path))
            || allowed_entries.contains(&(difference.path.clone(), difference.after.clone())))
    {
        return true;
    }
    if difference.path.ends_with(".attributes") {
        return minted_identity_only(&difference.before, &difference.after);
    }
    false
}

/// The "paragraph-id-minting" normalization: the after side may add managed
/// identity attributes and nothing else. Fails closed on any other change.
fn minted_identity_only(before: &str, after: &str) -> bool {
    let before_tokens: BTreeSet<&str> = before.split(',').filter(|t| !t.is_empty()).collect();
    let after_tokens: BTreeSet<&str> = after.split(',').filter(|t| !t.is_empty()).collect();
    before_tokens.is_subset(&after_tokens)
        && after_tokens.difference(&before_tokens).all(|token| {
            MANAGED_IDENTITY_ATTRIBUTES
                .iter()
                .any(|(namespace, local)| token.starts_with(&format!("{{{namespace}}}{local}=")))
        })
}

#[cfg(test)]
mod tests {
    use super::*;

    const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";

    fn package(document: &str) -> Vec<Part> {
        vec![(
            "word/document.xml".to_owned(),
            format!(r#"<w:document xmlns:w="{W}" xmlns:w14="{W14}"><w:body>{document}</w:body></w:document>"#)
                .into_bytes(),
        )]
    }

    #[test]
    fn a_clean_round_trip_reports_nothing() {
        let parts = package(r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#);
        assert_eq!(
            roundtrip_findings(&parts, &parts).unwrap(),
            Vec::<String>::new()
        );
    }

    #[test]
    fn minted_paragraph_identity_is_tolerated() {
        let before = package(r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#);
        let after = package(r#"<w:p w14:paraId="11111111"><w:r><w:t>x</w:t></w:r></w:p>"#);
        assert_eq!(
            roundtrip_findings(&before, &after).unwrap(),
            Vec::<String>::new()
        );
    }

    #[test]
    fn a_lost_paragraph_identity_is_not_tolerated() {
        let before = package(r#"<w:p w14:paraId="11111111"><w:r><w:t>x</w:t></w:r></w:p>"#);
        let after = package(r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#);
        assert_ne!(
            roundtrip_findings(&before, &after).unwrap(),
            Vec::<String>::new()
        );
    }

    #[test]
    fn a_changed_paragraph_identity_is_not_tolerated() {
        let before = package(r#"<w:p w14:paraId="11111111"><w:r><w:t>x</w:t></w:r></w:p>"#);
        let after = package(r#"<w:p w14:paraId="22222222"><w:r><w:t>x</w:t></w:r></w:p>"#);
        assert_ne!(
            roundtrip_findings(&before, &after).unwrap(),
            Vec::<String>::new()
        );
    }

    #[test]
    fn minting_alongside_another_attribute_change_is_not_tolerated() {
        let before = package(r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#);
        let after = package(
            r#"<w:p w14:paraId="11111111" w:rsidR="00AA00AA"><w:r><w:t>x</w:t></w:r></w:p>"#,
        );
        assert_ne!(
            roundtrip_findings(&before, &after).unwrap(),
            Vec::<String>::new()
        );
    }

    /// D4 rule 2: the engine does not re-emit unmodelled parts, so lexical
    /// noise there is a rewrite — while the same noise in a modelled part is
    /// exactly what the fingerprint is allowed to forgive.
    #[test]
    fn byte_rule_two_holds_unmodelled_parts_to_bytes_not_trees() {
        let styles = |attributes: &str| {
            (
                "word/styles.xml".to_owned(),
                format!(r#"<w:styles xmlns:w="{W}"><w:style {attributes}/></w:styles>"#)
                    .into_bytes(),
            )
        };
        let mut before = package(r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#);
        let mut after = before.clone();
        before.push(styles(r#"w:type="paragraph" w:styleId="Normal""#));
        after.push(styles(r#"w:styleId="Normal" w:type="paragraph""#));
        assert_eq!(
            roundtrip_findings(&before, &after).unwrap(),
            vec!["unmodelled part rewritten: word/styles.xml".to_owned()]
        );

        let modelled_before = package(r#"<w:p><w:r w:rsidR="00AA"><w:t>x</w:t></w:r></w:p>"#);
        let modelled_after = package(r#"<w:p><w:r w:rsidR="00AA" ><w:t>x</w:t></w:r  ></w:p>"#);
        assert_eq!(
            roundtrip_findings(&modelled_before, &modelled_after).unwrap(),
            Vec::<String>::new()
        );
    }

    #[test]
    fn companion_part_additions_are_tolerated_and_others_are_not() {
        let before = package(r#"<w:p><w:r><w:t>x</w:t></w:r></w:p>"#);
        let mut with_companion = before.clone();
        with_companion.push((
            "word/commentsExtended.xml".to_owned(),
            r#"<w15:commentsEx xmlns:w15="urn:w15"/>"#.as_bytes().to_vec(),
        ));
        assert_eq!(
            roundtrip_findings(&before, &with_companion).unwrap(),
            Vec::<String>::new()
        );
        let mut with_stranger = before.clone();
        with_stranger.push((
            "word/stranger.xml".to_owned(),
            "<x/>".to_owned().into_bytes(),
        ));
        assert_eq!(
            roundtrip_findings(&before, &with_stranger).unwrap(),
            vec![
                "part added: word/stranger.xml".to_owned(),
                "digest: word/stranger.xml | absent -> present".to_owned(),
            ]
        );
    }
}
