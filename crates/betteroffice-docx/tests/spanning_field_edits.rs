use betteroffice_docx::{Document, EditCtx, Error, Receipt};

const DECLARATION: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
const NAMESPACES: &str = r#"xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 w15""#;
const SECTION: &str = r#"<w:sectPr w:rsidR="00A1B2C3"><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>"#;
const AFTER_TEXT: &str = "After the contents";

/// A paragraph of the table of contents fixture, by its ID and its index among the paragraphs.
#[derive(Clone, Copy)]
struct Target {
    id: &'static str,
    index: usize,
}

const OPENING: Target = Target {
    id: "0000001A",
    index: 0,
};
const MIDDLE: Target = Target {
    id: "0000001B",
    index: 1,
};
const SECOND_MIDDLE: Target = Target {
    id: "0000001C",
    index: 2,
};
const CLOSING: Target = Target {
    id: "0000001D",
    index: 3,
};
const AFTER: Target = Target {
    id: "0000001E",
    index: 4,
};

type Replace = fn(&mut Document, Target, &str) -> Result<Receipt, Error>;

const VARIANTS: [(&str, Replace); 4] = [
    ("replace_paragraph_text", replace),
    ("replace_paragraph_text_with", replace_with),
    ("replace_paragraph_text_at", replace_at),
    ("replace_paragraph_text_at_with", replace_at_with),
];

fn replace(document: &mut Document, target: Target, text: &str) -> Result<Receipt, Error> {
    document.replace_paragraph_text(target.id, text)
}

fn replace_with(document: &mut Document, target: Target, text: &str) -> Result<Receipt, Error> {
    document.replace_paragraph_text_with(target.id, text, 7, &context())
}

fn replace_at(document: &mut Document, target: Target, text: &str) -> Result<Receipt, Error> {
    document.replace_paragraph_text_at(target.index, text)
}

fn replace_at_with(document: &mut Document, target: Target, text: &str) -> Result<Receipt, Error> {
    document.replace_paragraph_text_at_with(target.index, text, 7, &context())
}

fn context() -> EditCtx {
    EditCtx::local("Editor", "2026-01-02T03:04:05.000Z")
}

fn main_part(body: &str) -> String {
    format!("{DECLARATION}<w:document {NAMESPACES}><w:body>{body}{SECTION}</w:body></w:document>")
}

fn docx(document: &str) -> Vec<u8> {
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        ("word/document.xml".to_owned(), document.as_bytes().to_vec()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn saved_main_part(package: &[u8]) -> String {
    let parts = ooxml_opc::unzip_parts(package).unwrap();
    let (_, bytes) = parts
        .iter()
        .find(|(name, _)| name == "word/document.xml")
        .expect("word/document.xml");
    String::from_utf8(bytes.clone()).unwrap()
}

fn paragraph(id: &str, content: &str) -> String {
    format!(r#"<w:p w14:paraId="{id}" w14:textId="77777777">{content}</w:p>"#)
}

fn run(text: &str) -> String {
    format!(r#"<w:r><w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

fn field_character(kind: &str) -> String {
    format!(r#"<w:r><w:fldChar w:fldCharType="{kind}"/></w:r>"#)
}

fn instruction(code: &str) -> String {
    format!(r#"<w:r><w:instrText xml:space="preserve">{code}</w:instrText></w:r>"#)
}

/// A table of contents whose field begins in its first entry's paragraph and ends in its last.
fn table_of_contents() -> String {
    [
        paragraph(
            OPENING.id,
            &[
                field_character("begin"),
                instruction(r#" TOC \o "1-3" \h \z \u "#),
                field_character("separate"),
                run("Introduction"),
            ]
            .concat(),
        ),
        paragraph(MIDDLE.id, &run("Scope")),
        paragraph(SECOND_MIDDLE.id, &run("Terms")),
        paragraph(
            CLOSING.id,
            &[run("Appendix"), field_character("end")].concat(),
        ),
    ]
    .concat()
}

fn spanning_source() -> Vec<u8> {
    docx(&main_part(&format!(
        "{}{}",
        table_of_contents(),
        paragraph(AFTER.id, &run(AFTER_TEXT))
    )))
}

/// What keeps the field characters of `xml` from being one begin, one separate and one end, with
/// the running depth of open fields never below zero and the end before `outside`.
fn field_faults(xml: &str, outside: &str) -> Vec<String> {
    const TYPE: &str = r#"w:fldCharType=""#;
    let mut faults = Vec::new();
    let mut counts = [0usize; 3];
    let mut depth = 0isize;
    let mut lowest = 0isize;
    let mut last_end = None;
    for (at, _) in xml.match_indices(TYPE) {
        let value = &xml[at + TYPE.len()..];
        match &value[..value.find('"').unwrap_or(value.len())] {
            "begin" => {
                counts[0] += 1;
                depth += 1;
            }
            "separate" => counts[1] += 1,
            "end" => {
                counts[2] += 1;
                depth -= 1;
                last_end = Some(at);
            }
            other => faults.push(format!("a field character of type {other:?}")),
        }
        lowest = lowest.min(depth);
    }
    if counts != [1, 1, 1] {
        faults.push(format!(
            "begin, separate, end are {counts:?}, not [1, 1, 1]"
        ));
    }
    if lowest < 0 {
        faults.push(format!("the running field depth falls to {lowest}"));
    }
    match (last_end, xml.find(outside)) {
        (_, None) => faults.push(format!("{outside:?} is missing")),
        (Some(end), Some(text)) if end > text => {
            faults.push(format!("the field ends after {outside:?}"))
        }
        _ => {}
    }
    faults
}

fn assert_balanced(saved: &str, outside: &str) {
    let faults = field_faults(saved, outside);
    assert!(faults.is_empty(), "{}: {saved}", faults.join("; "));
}

/// Why writing `text` into `target` with `replace` and saving does not keep the spanning field
/// whole: a refused edit that changed the save, a failed save, missing text or faulty field
/// characters. Empty for a refused edit that changed nothing and for a whole field.
fn edit_faults(replace: Replace, target: Target, text: &str) -> Vec<String> {
    let source = spanning_source();
    let mut document = Document::open(&source).unwrap();
    let before = match document.save() {
        Ok(bytes) => bytes,
        Err(error) => return vec![format!("the unedited save failed: {error:?}")],
    };
    if replace(&mut document, target, text).is_err() {
        return match document.save() {
            Ok(after) if after == before => Vec::new(),
            Ok(_) => vec!["the edit was refused but the save changed".to_owned()],
            Err(error) => vec![format!("the save after a refused edit failed: {error:?}")],
        };
    }
    let saved = match document.save() {
        Ok(bytes) => saved_main_part(&bytes),
        Err(error) => return vec![format!("the save failed: {error:?}")],
    };
    let mut faults = field_faults(&saved, AFTER_TEXT);
    if !saved.contains(text) {
        faults.push(format!("{text:?} is missing"));
    }
    faults
}

fn edited(target: Target, text: &str) -> String {
    let mut document = Document::open(&spanning_source()).unwrap();
    document.replace_paragraph_text(target.id, text).unwrap();
    saved_main_part(&document.save().unwrap())
}

#[test]
fn replacing_the_closing_paragraph_of_a_spanning_field_saves_the_field_balanced() {
    let saved = edited(CLOSING, "Appendix, edited");
    assert!(saved.contains("Appendix, edited"), "{saved}");
    assert_balanced(&saved, AFTER_TEXT);
}

#[test]
fn replacing_a_middle_paragraph_of_a_spanning_field_saves_the_field_balanced() {
    let saved = edited(MIDDLE, "Scope, edited");
    assert!(saved.contains("Scope, edited"), "{saved}");
    assert_balanced(&saved, AFTER_TEXT);
}

#[test]
fn replacing_the_opening_paragraph_of_a_spanning_field_is_refused_or_saves_it_balanced() {
    assert_eq!(
        edit_faults(replace, OPENING, "Introduction, edited"),
        Vec::<String>::new()
    );
}

#[test]
fn every_paragraph_text_replacement_is_refused_or_saves_a_spanning_field_balanced() {
    let positions = [
        ("opening", OPENING),
        ("middle", MIDDLE),
        ("closing", CLOSING),
    ];
    let mut failures = Vec::new();
    for (name, replace) in VARIANTS {
        for (position, target) in positions {
            for fault in edit_faults(replace, target, "Rewritten entry") {
                failures.push(format!("{name} on the {position} paragraph: {fault}"));
            }
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn replacing_the_paragraph_after_a_spanning_field_keeps_the_field_bytes() {
    let saved = edited(AFTER, "After the contents, edited");
    assert!(saved.contains("After the contents, edited"), "{saved}");
    assert!(
        saved.contains(&table_of_contents()),
        "source bytes differ in the table of contents field: {saved}"
    );
    assert_balanced(&saved, AFTER_TEXT);
}

#[test]
fn replacing_a_paragraph_beside_a_field_within_one_paragraph_keeps_the_field_bytes() {
    let field = paragraph(
        "0000002A",
        &[
            run("Page "),
            field_character("begin"),
            instruction(" PAGE "),
            field_character("separate"),
            run("1"),
            field_character("end"),
        ]
        .concat(),
    );
    let source = main_part(&format!(
        "{field}{}",
        paragraph("0000002B", &run("Body text"))
    ));
    let mut document = Document::open(&docx(&source)).unwrap();
    document
        .replace_paragraph_text("0000002B", "Body text, edited")
        .unwrap();
    let saved = saved_main_part(&document.save().unwrap());
    assert!(
        saved.contains(&field),
        "source bytes differ in the field paragraph: {saved}"
    );
    assert_balanced(&saved, "Body text, edited");
}
