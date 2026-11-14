//! Export a DOCX file to PDF.
//!
//! ```text
//! cargo run -p betteroffice-docx --example export_pdf -- input.docx output.pdf
//!     [--pdfa-2u] [--title TEXT] [--author TEXT] [--subject TEXT]
//!     [--keywords TEXT] [--creator TEXT] [--creation-date PDF_DATE]
//!     [--modification-date PDF_DATE]
//! ```
//!
//! Without flags the output is the historical plain PDF 1.7. `--pdfa-2u`
//! targets PDF/A-2u (ISO 19005-2); it needs embedded fonts and refuses a
//! document that would fall back to Base-14 text. The metadata flags populate
//! `/Info` and the XMP packet; dates are PDF date strings such as
//! `D:20260101120000Z` and are written verbatim.

use betteroffice_docx::{Document, PdfConformance, PdfExportOptions};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut input = None;
    let mut output = None;
    let mut options = PdfExportOptions::default();
    let mut args = std::env::args().skip(1);
    while let Some(arg) = args.next() {
        let (flag, inline) = match arg.split_once('=') {
            Some((flag, value)) if flag.starts_with("--") => {
                (flag.to_owned(), Some(value.to_owned()))
            }
            _ if arg.starts_with("--") => (arg.clone(), None),
            _ => (String::new(), None),
        };
        match flag.as_str() {
            "--pdfa-2u" => {
                if inline.is_some() {
                    return Err("--pdfa-2u takes no value".into());
                }
                options.conformance = PdfConformance::PdfA2u;
            }
            "--title" => options.metadata.title = Some(value("--title", inline, &mut args)?),
            "--author" => options.metadata.author = Some(value("--author", inline, &mut args)?),
            "--subject" => options.metadata.subject = Some(value("--subject", inline, &mut args)?),
            "--keywords" => {
                options.metadata.keywords = Some(value("--keywords", inline, &mut args)?);
            }
            "--creator" => options.metadata.creator = Some(value("--creator", inline, &mut args)?),
            "--creation-date" => {
                options.metadata.creation_date = Some(value("--creation-date", inline, &mut args)?);
            }
            "--modification-date" => {
                options.metadata.modification_date =
                    Some(value("--modification-date", inline, &mut args)?);
            }
            _ if flag.starts_with('-') => return Err(format!("unknown option {flag}").into()),
            _ => {
                if input.is_none() {
                    input = Some(arg);
                } else if output.is_none() {
                    output = Some(arg);
                } else {
                    return Err(format!("unexpected argument {arg}").into());
                }
            }
        }
    }

    let input = input.ok_or("missing input DOCX")?;
    let output = output.ok_or("missing output PDF")?;
    let bytes = std::fs::read(input)?;
    let document = Document::open(&bytes)?;
    let pdf = document.export_pdf_with_options(&options)?;
    std::fs::write(output, pdf)?;
    Ok(())
}

/// The value for a flag, from `--flag=value` or the next argument.
fn value(
    flag: &str,
    inline: Option<String>,
    args: &mut impl Iterator<Item = String>,
) -> Result<String, String> {
    inline
        .or_else(|| args.next())
        .ok_or_else(|| format!("{flag} needs a value"))
}
