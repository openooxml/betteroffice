use betteroffice_docx::{Document, get_paragraph_text};
use serde_json::Value;

use crate::Result;

pub fn open(bytes: &[u8]) -> Result<Document> {
    Ok(Document::open(bytes)?)
}

pub fn edit(document: &mut Document, probe: &Value) -> Result<()> {
    let old = probe["old"].as_str().ok_or("Missing original text")?;
    let matches: Vec<_> = document
        .paragraphs()
        .into_iter()
        .filter(|paragraph| get_paragraph_text(paragraph) == old)
        .collect();
    if matches.len() != 1 {
        return Err("Original paragraph is missing or ambiguous".into());
    }
    let id = matches[0]
        .para_id
        .clone()
        .ok_or("Paragraph has no edit ID")?;
    document.replace_paragraph_text(
        &id,
        probe["new"].as_str().ok_or("Missing replacement text")?,
    )?;
    Ok(())
}

pub fn save(document: &Document) -> Result<Vec<u8>> {
    Ok(document.save()?)
}

pub fn verify(document: &Document, probe: &Value) -> Result<()> {
    let expected = probe["new"].as_str().ok_or("Missing replacement text")?;
    if document
        .paragraphs()
        .iter()
        .filter(|p| get_paragraph_text(p) == expected)
        .count()
        != 1
    {
        return Err("Edited paragraph did not survive reopening".into());
    }
    Ok(())
}
