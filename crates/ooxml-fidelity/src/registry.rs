//! Frozen comparison policy: how each fidelity artifact is compared, and the
//! serializer normalizations the fingerprint is allowed to forgive.
//!
//! Nothing outside this registry may loosen a comparison; an undeclared
//! serializer deviation is a bug, not a tolerance.

use crate::error::FidelityError;

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

/// One intentional serializer deviation from input XML, reviewed and named.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Normalization {
    pub id: &'static str,
    pub description: &'static str,
}

/// The markup-compatibility namespace, home of `mc:Ignorable`.
pub const MC_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// The declaration boilerplate Word writes on every WML part root. A
/// declaration binds a prefix and carries no meaning of its own: an unused
/// one can appear or vanish freely, and a used-but-undeclared prefix fails
/// the parse outright, so excluding these URIs from the fingerprint's
/// binding set forgives exactly the inert difference and nothing else.
pub const STANDARD_WML_NAMESPACE_URIS: &[&str] = &[
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://schemas.microsoft.com/office/word/2006/wordml",
    "http://schemas.microsoft.com/office/word/2010/wordml",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "http://schemas.microsoft.com/office/word/2012/wordml",
    "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "http://schemas.microsoft.com/office/word/2018/wordml",
    "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "urn:schemas-microsoft-com:office:office",
    "urn:schemas-microsoft-com:office:word",
    "urn:schemas-microsoft-com:vml",
];

/// Every entry is one reviewed serializer deviation the fingerprint forgives.
pub const DECLARED_NORMALIZATIONS: &[Normalization] = &[
    Normalization {
        id: "root-standard-namespace-declarations",
        description: "Modelled WML parts re-emit Word's standard namespace declaration set on \
                      the root; the fingerprint excludes those URIs from the binding set.",
    },
    Normalization {
        id: "root-mc-ignorable",
        description: "Modelled WML parts re-emit Word's standard mc:Ignorable list on the root; \
                      the fingerprint excludes mc:Ignorable on root elements only.",
    },
];

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
    fn exactly_the_two_reviewed_normalizations_are_declared() {
        let ids: Vec<&str> = DECLARED_NORMALIZATIONS
            .iter()
            .map(|normalization| normalization.id)
            .collect();
        assert_eq!(
            ids,
            vec!["root-standard-namespace-declarations", "root-mc-ignorable"]
        );
    }
}
