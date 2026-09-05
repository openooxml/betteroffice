//! Frozen comparison policy: how each fidelity artifact is compared, and the
//! serializer normalizations the fingerprint is allowed to forgive.
//!
//! Nothing outside this registry may loosen a comparison; an undeclared
//! serializer deviation is a bug, not a tolerance.

use crate::error::FidelityError;
use crate::xml::XmlAttribute;

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum ComparisonMode {
    /// Every field compares; nothing is dropped.
    Exact,
    /// Exact after declared canonicalization, with named ephemera dropped.
    CanonicalExact { ephemera: &'static [&'static str] },
    /// Numeric epsilon, allowed only for documented raster comparison.
    Tolerance { epsilon: f64 },
}

const COMPARISONS: &[(&str, ComparisonMode)] = &[
    ("structural-fingerprint", ComparisonMode::Exact),
    ("semantic-digest", ComparisonMode::Exact),
    ("element-census", ComparisonMode::Exact),
    ("non-xml-part-bytes", ComparisonMode::Exact),
    ("unmodelled-xml-part-bytes", ComparisonMode::Exact),
    ("saved-package-bytes", ComparisonMode::Exact),
    ("render-golden-png", ComparisonMode::Exact),
];

/// The mode an artifact compares under; an unregistered artifact refuses.
pub fn comparison_mode(artifact: &str) -> Result<ComparisonMode, FidelityError> {
    COMPARISONS
        .iter()
        .find(|(name, _)| *name == artifact)
        .map(|(_, mode)| *mode)
        .ok_or_else(|| FidelityError::UnknownArtifact(artifact.to_owned()))
}

/// The XML parts the engine re-emits from its model. Every other XML part is
/// unmodelled: byte rule 2 holds it to identical bytes through save, so a part
/// the engine starts rewriting must be declared here in the same change.
pub const MODELLED_XML_PARTS: &[&str] = &[
    "[Content_Types].xml",
    "docProps/core.xml",
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsExtensible.xml",
    "word/commentsIds.xml",
    "word/document.xml",
    "word/endnotes.xml",
    "word/footnotes.xml",
    "word/numbering.xml",
];

/// Header and footer parts carry a serial number, so they match by stem.
const MODELLED_STORY_STEMS: &[&str] = &["word/header", "word/footer"];

/// True when the engine re-emits this XML part from its model.
pub fn is_modelled_xml_part(name: &str) -> bool {
    name.ends_with(".rels")
        || MODELLED_XML_PARTS.contains(&name)
        || MODELLED_STORY_STEMS
            .iter()
            .any(|stem| name.starts_with(stem))
}

/// One intentional serializer deviation from input XML, reviewed and named.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Normalization {
    pub id: &'static str,
    pub description: &'static str,
}

/// The markup-compatibility namespace, home of `mc:Ignorable`.
pub const MC_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// Standard declaration URIs; QName attribute values are resolved before comparison.
pub const STANDARD_WML_NAMESPACE_URIS: &[&str] = &[
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://schemas.microsoft.com/office/2019/extlst",
    "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
    "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
    "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
    "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
    "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex",
    "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
    "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
    "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
    "http://schemas.microsoft.com/office/drawing/2016/ink",
    "http://schemas.microsoft.com/office/drawing/2017/model3d",
    "http://schemas.microsoft.com/office/word/2006/wordml",
    "http://schemas.microsoft.com/office/word/2010/wordml",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "http://schemas.microsoft.com/office/word/2012/wordml",
    "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "http://schemas.microsoft.com/office/word/2018/wordml",
    "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock",
    "urn:schemas-microsoft-com:office:office",
    "urn:schemas-microsoft-com:office:word",
    "urn:schemas-microsoft-com:vml",
];

/// The comment companion parts the serializer materializes on save, exactly
/// as Word does when it writes threaded comments.
pub const COMPANION_PART_NAMES: &[&str] = &[
    "word/commentsExtended.xml",
    "word/commentsIds.xml",
    "word/commentsExtensible.xml",
];

/// Managed paragraph identity the serializer may mint where absent.
pub const MANAGED_IDENTITY_ATTRIBUTES: &[(&str, &str)] = &[
    (
        "http://schemas.microsoft.com/office/word/2010/wordml",
        "paraId",
    ),
    (
        "http://schemas.microsoft.com/office/word/2010/wordml",
        "textId",
    ),
];

/// Every entry is one reviewed serializer deviation the oracles forgive.
pub const DECLARED_NORMALIZATIONS: &[Normalization] = &[
    Normalization {
        id: "root-standard-namespace-declarations",
        description: "Modelled WML parts re-emit Word's standard namespace declaration set on \
                      the root; the fingerprint excludes those URIs from the binding set.",
    },
    Normalization {
        id: "root-mc-ignorable",
        description: "Root mc:Ignorable compares resolved URI sets; only additions from \
                      STANDARD_WML_NAMESPACE_URIS are tolerated, never removals.",
    },
    Normalization {
        id: "relationship-and-content-type-order",
        description: "Relationship and content-type entries carry no authored order; the digest \
                      compares them as sorted identity sets without positions.",
    },
    Normalization {
        id: "comment-companion-parts",
        description: "Saving a document with comments materializes Word's companion parts \
                      (commentsExtended, commentsIds, commentsExtensible) and their content-type \
                      and relationship entries; the report tolerates exactly those additions.",
    },
    Normalization {
        id: "paragraph-id-minting",
        description: "The serializer mints w14:paraId/w14:textId where absent, as Word does. \
                      Additions are tolerated; a changed or lost identity is never forgiven.",
    },
];

pub(crate) fn normalize_root_ignorable(before: &[XmlAttribute], after: &mut Vec<XmlAttribute>) {
    let is_ignorable = |attribute: &&XmlAttribute| {
        attribute.namespace == MC_NAMESPACE && attribute.local == "Ignorable"
    };
    let original = before.iter().find(is_ignorable);
    let saved = after.iter().find(is_ignorable);
    let before_set: std::collections::BTreeSet<_> = original
        .map(|attribute| attribute.value.as_str())
        .unwrap_or("")
        .split_whitespace()
        .collect();
    let after_set: std::collections::BTreeSet<_> = saved
        .map(|attribute| attribute.value.as_str())
        .unwrap_or("")
        .split_whitespace()
        .collect();
    if before_set.is_subset(&after_set)
        && after_set
            .difference(&before_set)
            .all(|uri| STANDARD_WML_NAMESPACE_URIS.contains(uri))
    {
        after.retain(|attribute| !is_ignorable(&attribute));
        if let Some(attribute) = original {
            after.push(attribute.clone());
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn registered_artifacts_resolve() {
        assert_eq!(
            comparison_mode("structural-fingerprint").unwrap(),
            ComparisonMode::Exact
        );
    }

    #[test]
    fn an_unregistered_artifact_refuses() {
        assert_eq!(
            comparison_mode("layout-geometry").unwrap_err(),
            FidelityError::UnknownArtifact("layout-geometry".to_owned())
        );
    }

    #[test]
    fn every_tolerance_declares_a_positive_epsilon() {
        for (name, mode) in COMPARISONS {
            if let ComparisonMode::Tolerance { epsilon } = mode {
                assert!(*epsilon > 0.0, "{name} declares a non-positive epsilon");
            }
        }
    }

    #[test]
    fn exactly_the_reviewed_normalizations_are_declared() {
        let ids: Vec<&str> = DECLARED_NORMALIZATIONS
            .iter()
            .map(|normalization| normalization.id)
            .collect();
        assert_eq!(
            ids,
            vec![
                "root-standard-namespace-declarations",
                "root-mc-ignorable",
                "relationship-and-content-type-order",
                "comment-companion-parts",
                "paragraph-id-minting",
            ]
        );
    }
}
