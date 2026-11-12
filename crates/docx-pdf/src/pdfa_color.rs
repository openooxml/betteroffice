//! PDF/A-2u color support: bundled sRGB output profile and output intent.
//!
//! The profile is a CC0-licensed, 736-byte sRGB IEC61966-2.1 ICC v2 display
//! profile. Provenance and licensing are recorded in `assets/ATTRIBUTION.md`.
//!
//! Integration requires registering this module in `lib.rs` with
//! `mod pdfa_color;` and, for `PdfConformance::PdfA2u` documents:
//!
//! 1. Reserve a profile object id and call
//!    [`write_output_profile`] while writing the document body.
//! 2. Open the catalog writer, write `/Pages` and `/Metadata`, then call
//!    [`attach_output_intent`] before finishing the catalog.

use pdf_writer::types::OutputIntentSubtype;
use pdf_writer::writers::Catalog;
use pdf_writer::{Finish, Pdf, Ref, TextStr};

use crate::PdfError;

/// The CC0 sRGB IEC61966-2.1 ICC v2 profile bundled in `assets/`.
///
/// This is a three-component RGB display profile (`mntr`) with an XYZ
/// connection space. See `assets/ATTRIBUTION.md` for its origin and license.
pub(crate) const SRGB_PROFILE: &[u8] = include_bytes!("../assets/sRGB-v2-magic.icc");

/// Every ICC profile starts with a 128-byte header.
const ICC_HEADER_LEN: usize = 128;

/// Component count for the bundled RGB profile and the emitted `/N` entry.
const PROFILE_COMPONENTS: i32 = 3;

/// Well-known output condition identifier for sRGB IEC61966-2.1.
const OUTPUT_CONDITION_IDENTIFIER: &str = "sRGB IEC61966-2.1";

/// Write the bundled sRGB profile as an ICC profile stream with `/N 3`.
///
/// `id` must be reserved by the caller and must not be written elsewhere. The
/// profile is checked before any PDF bytes are emitted; a malformed bundled
/// asset yields [`PdfError`] instead of a panic.
pub(crate) fn write_output_profile(pdf: &mut Pdf, id: Ref) -> Result<(), PdfError> {
    validate_profile(SRGB_PROFILE)?;

    let mut profile = pdf.icc_profile(id, SRGB_PROFILE);
    profile.n(PROFILE_COMPONENTS);
    profile.finish();

    Ok(())
}

/// Attach a single PDF/A output intent (`GTS_PDFA1`) to the catalog.
///
/// Writes one `/OutputIntents` entry with subtype `GTS_PDFA1`,
/// `OutputConditionIdentifier` and `Info` set to `sRGB IEC61966-2.1`, and
/// `/DestOutputProfile` pointing at `profile` (written by
/// [`write_output_profile`]).
pub(crate) fn attach_output_intent(
    catalog: &mut Catalog<'_>,
    profile: Ref,
) -> Result<(), PdfError> {
    validate_profile(SRGB_PROFILE)?;

    let mut intents = catalog.output_intents();
    let mut intent = intents.push();
    intent
        .subtype(OutputIntentSubtype::PDFA)
        .output_condition_identifier(TextStr(OUTPUT_CONDITION_IDENTIFIER))
        .info(TextStr(OUTPUT_CONDITION_IDENTIFIER))
        .dest_output_profile(profile);
    intent.finish();
    intents.finish();

    Ok(())
}

/// Validate the ICC fields this module relies on before writing a PDF profile.
///
/// Checks only the header fields required here: length, declared size, version,
/// device class, RGB data color space, and the resulting three-component
/// consistency with the emitted `/N 3`.
fn validate_profile(profile: &[u8]) -> Result<(), PdfError> {
    if profile.len() <= ICC_HEADER_LEN {
        return Err(PdfError::InvalidNumber(
            "ICC profile is missing its 128-byte header",
        ));
    }

    let declared_size =
        u32::from_be_bytes([profile[0], profile[1], profile[2], profile[3]]) as usize;
    if declared_size != profile.len() {
        return Err(PdfError::InvalidNumber(
            "ICC profile declared size does not match its length",
        ));
    }

    if profile[8] >= 5 {
        return Err(PdfError::Unsupported("ICC profile version is 5.0 or newer"));
    }

    let class = &profile[12..16];
    if class != &b"mntr"[..] && class != &b"prtr"[..] {
        return Err(PdfError::Unsupported(
            "ICC profile device class is not mntr or prtr",
        ));
    }

    let components = if &profile[16..20] == &b"RGB "[..] {
        3
    } else {
        return Err(PdfError::Unsupported(
            "ICC profile data color space is not RGB",
        ));
    };
    if components != PROFILE_COMPONENTS {
        return Err(PdfError::Unsupported(
            "ICC profile component count does not match /N 3",
        ));
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use pdf_writer::Content;

    fn corrupt(offset: usize, replacement: &[u8]) -> Vec<u8> {
        let mut profile = SRGB_PROFILE.to_vec();
        profile[offset..offset + replacement.len()].copy_from_slice(replacement);
        profile
    }

    #[test]
    fn bundled_profile_has_a_valid_rgb_header() {
        let declared = u32::from_be_bytes([
            SRGB_PROFILE[0],
            SRGB_PROFILE[1],
            SRGB_PROFILE[2],
            SRGB_PROFILE[3],
        ]) as usize;

        assert!(SRGB_PROFILE.len() > ICC_HEADER_LEN);
        assert_eq!(declared, SRGB_PROFILE.len(), "declared profile size");
        assert_eq!(&SRGB_PROFILE[12..16], &b"mntr"[..], "device class");
        assert_eq!(&SRGB_PROFILE[16..20], &b"RGB "[..], "data color space");
        assert!(SRGB_PROFILE[8] < 5, "major version");
        assert!(validate_profile(SRGB_PROFILE).is_ok());
    }

    #[test]
    fn malformed_profiles_are_rejected_without_panicking() {
        assert!(matches!(
            validate_profile(&SRGB_PROFILE[..ICC_HEADER_LEN]),
            Err(PdfError::InvalidNumber(_))
        ));

        let wrong_size = corrupt(0, &1u32.to_be_bytes());
        assert!(matches!(
            validate_profile(&wrong_size),
            Err(PdfError::InvalidNumber(_))
        ));

        assert!(matches!(
            validate_profile(&corrupt(12, b"scnr")),
            Err(PdfError::Unsupported(_))
        ));
        assert!(matches!(
            validate_profile(&corrupt(16, b"CMYK")),
            Err(PdfError::Unsupported(_))
        ));
        assert!(matches!(
            validate_profile(&corrupt(8, &[5, 0, 0, 0])),
            Err(PdfError::Unsupported(_))
        ));
    }

    #[test]
    fn output_profile_and_intent_emit_pdfa_bytes() {
        let mut pdf = Pdf::new();
        let profile_id = Ref::new(1);
        write_output_profile(&mut pdf, profile_id).expect("bundled profile is valid");

        let catalog_id = Ref::new(2);
        {
            let mut catalog = pdf.catalog(catalog_id);
            attach_output_intent(&mut catalog, profile_id).expect("output intent");
        }

        let content_id = Ref::new(3);
        let mut content = Content::new();
        content.rect(0.0, 0.0, 10.0, 10.0).fill_nonzero();
        let content_bytes = content.finish().into_vec();
        let stream = pdf.stream(content_id, &content_bytes);
        stream.finish();

        let bytes = pdf.finish();
        let text = String::from_utf8_lossy(&bytes);

        assert!(text.contains("/OutputIntents"));
        assert_eq!(text.matches("/Type /OutputIntent").count(), 1);
        assert!(text.contains("/S /GTS_PDFA1"));
        assert!(text.contains("/OutputConditionIdentifier (sRGB IEC61966-2.1)"));
        assert!(text.contains("/Info (sRGB IEC61966-2.1)"));
        assert_eq!(text.matches("(sRGB IEC61966-2.1)").count(), 2);
        assert!(text.contains("/DestOutputProfile 1 0 R"));
        let length_entry = format!("/Length {}", SRGB_PROFILE.len());
        assert!(text.contains(length_entry.as_str()));
        assert!(text.contains("/N 3"));
        assert!(
            bytes
                .windows(SRGB_PROFILE.len())
                .any(|window| window == SRGB_PROFILE),
            "profile bytes are embedded in the PDF"
        );
        assert!(
            bytes
                .windows(content_bytes.len())
                .any(|window| window == content_bytes.as_slice()),
            "page content stream is embedded in the PDF"
        );
    }
}
