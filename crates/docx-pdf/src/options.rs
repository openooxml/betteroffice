//! Export configuration for the native PDF writer.

/// PDF conformance profile the exporter targets.
///
/// [`PdfConformance::Standard`] preserves the historical output: a plain,
/// untagged PDF 1.7 document without archival claims. [`PdfConformance::PdfA2u`]
/// opts into ISO 19005-2 (PDF/A-2u), which additionally requires embedded
/// fonts, a document information dictionary, XMP metadata, and a color output
/// intent.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PdfConformance {
    /// Plain PDF 1.7, without an archival claim.
    #[default]
    Standard,
    /// PDF/A-2u: ISO 19005-2 archival PDF with Unicode-mapped text.
    PdfA2u,
}

impl PdfConformance {
    /// Whether this profile is a PDF/A profile.
    pub(crate) const fn is_pdfa(self) -> bool {
        matches!(self, Self::PdfA2u)
    }
}

/// Document information carried into the PDF.
///
/// Every field is optional; absent fields are omitted from `/Info` and XMP
/// instead of being replaced with invented values. `creation_date` and
/// `modification_date` are PDF date strings (`D:YYYYMMDDHHmmSSOHH'mm'`); the
/// metadata writer validates or rejects them, this type stores them verbatim.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PdfMetadata {
    /// Document title (`/Title`, `dc:title`).
    pub title: Option<String>,
    /// Author (`/Author`, `dc:creator`).
    pub author: Option<String>,
    /// Subject (`/Subject`, `dc:description`).
    pub subject: Option<String>,
    /// Keywords (`/Keywords`, `pdf:Keywords`).
    pub keywords: Option<String>,
    /// Creating application (`/Creator`, `xmp:CreatorTool`).
    pub creator: Option<String>,
    /// Creation date (`/CreationDate`, `xmp:CreateDate`).
    pub creation_date: Option<String>,
    /// Last modification date (`/ModDate`, `xmp:ModifyDate`).
    pub modification_date: Option<String>,
}

impl PdfMetadata {
    /// Whether every field is absent.
    pub fn is_empty(&self) -> bool {
        self.title.is_none()
            && self.author.is_none()
            && self.subject.is_none()
            && self.keywords.is_none()
            && self.creator.is_none()
            && self.creation_date.is_none()
            && self.modification_date.is_none()
    }
}

/// Options for [`crate::export_pdf_with_options`].
///
/// [`Default`] is [`PdfConformance::Standard`] with empty metadata, matching
/// the behavior of the existing convenience exporters.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct PdfExportOptions {
    /// Conformance profile to target.
    pub conformance: PdfConformance,
    /// Document information to emit.
    ///
    /// Written for [`PdfConformance::PdfA2u`], where `/Info` and XMP are
    /// required. Standard exports keep the historical output and do not emit
    /// metadata.
    pub metadata: PdfMetadata,
}
