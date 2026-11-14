use std::time::{Duration, Instant};

use docx_edit::*;

const DELETED_PARAGRAPHS: usize = 2_000;

fn package() -> Vec<u8> {
    let paragraphs: String = (0..DELETED_PARAGRAPHS)
        .map(|index| format!("<w:p><w:r><w:t>Paragraph {index}</w:t></w:r></w:p>"))
        .collect();
    let document = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{paragraphs}<w:sectPr/></w:body></w:document>"#
    );
    ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ])
    .unwrap()
}

fn delete_body_beside(stories: usize) -> Duration {
    let doc = EditingDoc::new(71);
    seed_from_docx(&doc, &package()).unwrap();
    for story in 0..stories {
        doc.create_story(format!("s{story}"), "Cell", "Normal", "left")
            .unwrap();
    }
    let undo = UndoSession::new();
    undo.track(&doc);
    let len = doc.story_len("body").unwrap();
    let started = Instant::now();
    doc.delete_range(
        &EditCtx::local("Ada", "2026-11-05T00:00:00Z"),
        StoryRange::new("body", 0, len - 1),
    )
    .unwrap();
    started.elapsed()
}

#[test]
fn a_large_delete_under_undo_does_not_slow_with_the_other_stories() {
    let best = |stories| (0..3).map(|_| delete_body_beside(stories)).min().unwrap();
    let (few, many) = (best(500), best(8_000));
    assert!(
        many < few * 4,
        "sixteen times the stories took {many:?} against {few:?}"
    );
}
