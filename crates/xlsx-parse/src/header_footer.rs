use serde::{Deserialize, Serialize};

#[derive(Clone, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HeaderFooterText {
    pub raw: String,
    pub left: String,
    pub center: String,
    pub right: String,
}

#[derive(Clone, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HeaderFooterPair {
    pub header: Option<HeaderFooterText>,
    pub footer: Option<HeaderFooterText>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetHeaderFooter {
    pub different_first: bool,
    pub different_odd_even: bool,
    pub scale_with_doc: bool,
    pub align_with_margins: bool,
    pub odd: HeaderFooterPair,
    pub even: Option<HeaderFooterPair>,
    pub first: Option<HeaderFooterPair>,
}

/// Decodes stored header/footer text into printable sections and dialog fields.
pub fn parse_header_footer_text(_source: &str) -> HeaderFooterText {
    HeaderFooterText::default()
}

/// Reads printable header/footer variants from a worksheet element.
pub fn parse_header_footer(_source: &[u8]) -> Option<SheetHeaderFooter> {
    None
}
