//! Story seeding must stay linear in paragraph count.
//!
//! Print the timing table with
//! `cargo test --release -p betteroffice-docx-edit --test seed_scaling -- --nocapture`.

use std::time::{Duration, Instant};

use docx_edit::{EditingDoc, seed_from_docx};

const SENTENCE: &str = "The applicant confirms that the control described in this section is operated by the compliance team and reviewed quarterly.";

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#;

const DOC_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"#;

fn build_docx(paragraph_count: usize, sentences_per_paragraph: usize) -> Vec<u8> {
    let text = vec![SENTENCE; sentences_per_paragraph].join(" ");
    let paragraph = format!(r#"<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>"#);
    let mut body = paragraph.repeat(paragraph_count);
    body.push_str(r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#);
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#
    );
    ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".to_owned(), CONTENT_TYPES.into()),
        ("_rels/.rels".to_owned(), ROOT_RELS.into()),
        ("word/_rels/document.xml.rels".to_owned(), DOC_RELS.into()),
        ("word/styles.xml".to_owned(), STYLES.into()),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ])
    .expect("synthetic package zips")
}

fn seed_time(paragraph_count: usize, sentences_per_paragraph: usize) -> Duration {
    let bytes = build_docx(paragraph_count, sentences_per_paragraph);
    let doc = EditingDoc::new(7);
    let start = Instant::now();
    seed_from_docx(&doc, &bytes).expect("seed succeeds");
    start.elapsed()
}

/// A quadratic seed makes per-paragraph cost grow ~8x from 500 to 4000
/// paragraphs; a linear one keeps it flat. 3x holds a wide margin both ways.
#[test]
fn seeding_stays_linear_in_paragraph_count() {
    let mut per_paragraph = Vec::new();
    for count in [500usize, 1000, 2000, 4000] {
        let elapsed = seed_time(count, 1);
        let millis = elapsed.as_secs_f64() * 1000.0 / count as f64;
        per_paragraph.push(millis);
        println!(
            "{count:>6} paragraphs   {:>8.2} s   {millis:>7.3} ms/paragraph",
            elapsed.as_secs_f64()
        );
    }
    let ratio = per_paragraph.last().unwrap() / per_paragraph.first().unwrap();
    assert!(
        ratio < 3.0,
        "per-paragraph seed cost grew {ratio:.1}x from 500 to 4000 paragraphs; seeding is no longer linear"
    );
}

/// The same character count must not get slower by being split into more
/// paragraphs: 4000 one-sentence paragraphs vs 500 eight-sentence ones.
#[test]
fn paragraph_heavy_shape_carries_no_penalty() {
    let many_paragraphs = seed_time(4000, 1).as_secs_f64();
    let few_paragraphs = seed_time(500, 8).as_secs_f64();
    let ratio = many_paragraphs / few_paragraphs;
    assert!(
        ratio < 8.0,
        "4000x1 took {ratio:.1}x the time of 500x8 at equal character count; the pre-fix penalty was ~40x"
    );
}
