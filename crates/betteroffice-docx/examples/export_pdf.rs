//! Export a DOCX file to PDF.
//!
//! `cargo run -p betteroffice-docx --example export_pdf -- input.docx output.pdf`

use betteroffice_docx::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let input = args.next().ok_or("missing input DOCX")?;
    let output = args.next().ok_or("missing output PDF")?;
    let bytes = std::fs::read(input)?;
    let document = Document::open(&bytes)?;
    let pdf = document.export_pdf()?;
    std::fs::write(output, pdf)?;
    Ok(())
}
