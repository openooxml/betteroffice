use betteroffice_pptx::{EditCtx, Presentation, ShapeSnapshot, StorySnapshot};
use serde_json::Value;

use crate::Result;

pub fn open(bytes: &[u8]) -> Result<Presentation> {
    Ok(Presentation::open(bytes)?)
}

fn matching(shapes: &[ShapeSnapshot], text: &str, stories: &mut Vec<StorySnapshot>) {
    for shape in shapes {
        stories.extend(
            shape
                .text_stories
                .iter()
                .filter(|story| story.plain_text() == text)
                .cloned(),
        );
        matching(&shape.children, text, stories);
    }
}

fn find(document: &Presentation, text: &str) -> Result<StorySnapshot> {
    let mut stories = Vec::new();
    for slide in document.snapshot()?.slides {
        matching(&slide.shapes, text, &mut stories);
    }
    if stories.len() != 1 {
        return Err("Text is missing or ambiguous".into());
    }
    Ok(stories.remove(0))
}

pub fn edit(document: &mut Presentation, probe: &Value) -> Result<()> {
    let old = probe["old"].as_str().ok_or("Missing original text")?;
    let story = find(document, old)?;
    if story.paragraphs.len() != 1 || story.paragraphs[0].runs.len() != 1 {
        return Err("The probe requires a single plain text run".into());
    }
    let added = probe["new"]
        .as_str()
        .and_then(|new| new.strip_prefix(old))
        .ok_or("Invalid append edit")?;
    document.insert_text(
        &EditCtx::local("roundtrip-benchmark"),
        &story.id,
        u32::try_from(old.encode_utf16().count())?,
        added,
        &story.paragraphs[0].runs[0].style,
    )?;
    Ok(())
}

pub fn save(document: &Presentation) -> Result<Vec<u8>> {
    Ok(document.save()?)
}

pub fn verify(document: &Presentation, probe: &Value) -> Result<()> {
    find(
        document,
        probe["new"].as_str().ok_or("Missing replacement text")?,
    )?;
    Ok(())
}
