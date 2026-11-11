//! Smoke-export a display-list JSON fixture.
//!
//! `cargo run -p betteroffice-docx-pdf --example export_fixture -- input.json output.pdf`

use betteroffice_docx_pdf::export_pdf;
use docx_layout::display_list::DisplayList;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = std::env::args().skip(1);
    let input = args.next().ok_or("missing input display-list JSON")?;
    let output = args.next().ok_or("missing output PDF path")?;
    let json = std::fs::read_to_string(input)?;
    let display_list: DisplayList = serde_json::from_str(&json)?;
    let pdf = export_pdf(&display_list)?;
    std::fs::write(output, pdf)?;
    Ok(())
}
