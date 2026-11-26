use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Arc;

use betteroffice_docx::{
    BlockContent, Document, InlineNode, ParagraphContent, RunContent, SaveOptions,
};
use docx_parse::{
    S9ParseOptions, S13SaveOptions, S13SaveRequest, SerializerDeterminism, parse_docx_s9_wire,
    write_docx_s13_parts_with_report,
};
use ooxml_opc::{rezip_parts, unzip_parts};
use sha2::{Digest, Sha256};

const CORE: &str = "docProps/core.xml";
const NOW: &str = "2030-01-02T03:04:05.006Z";
const ORIGINAL_TIME: &str = "2024-01-02T03:04:05Z";

fn synthetic_package() -> Vec<u8> {
    let parts = [
        (
            "[Content_Types].xml",
            concat!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
                r#"<Override PartName="/word/document.xml" ContentType="application/"#,
                r#"vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>"#,
                "</Types>",
            ),
        ),
        (
            "_rels/.rels",
            concat!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/"#,
                r#"relationships"><Relationship Id="main" Type="http://"#,
                "schemas.openxmlformats.org/",
                r#"officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>"#,
                "</Relationships>",
            ),
        ),
        (
            "word/document.xml",
            concat!(
                r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/"#,
                r#"main"><w:body><w:p><w:r><w:t>Original</w:t></w:r></w:p></w:body></w:document>"#,
            ),
        ),
        (
            CORE,
            concat!(
                r#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/"#,
                r#"metadata/core-properties" xmlns:dcterms="http://purl.org/dc/terms/" "#,
                r#"xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">"#,
                r#"<dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-02T03:04:05Z"#,
                "</dcterms:modified><cp:lastModifiedBy>Before</cp:lastModifiedBy>",
                "</cp:coreProperties>",
            ),
        ),
    ];
    rezip_parts(
        &parts
            .into_iter()
            .map(|(path, xml)| (path.to_owned(), xml.as_bytes().to_vec()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    unzip_parts(bytes).unwrap().into_iter().collect()
}

fn assert_whole_model_write(source: &[u8], document: &Document, options: SaveOptions) {
    let model = document.model();
    let request = S13SaveRequest {
        determinism: SerializerDeterminism {
            seed: format!("{:x}", Sha256::digest(source)),
            now: options.now.clone(),
        },
        document: model.body.clone(),
        header_entries: model.headers.clone(),
        footer_entries: model.footers.clone(),
        footnotes: model.footnotes.clone(),
        endnotes: model.endnotes.clone(),
        footnote_separators: model.footnote_separators.clone(),
        endnote_separators: model.endnote_separators.clone(),
        relationship_entries: model.relationships.clone(),
        numbering: Some(model.numbering.clone()),
        options: S13SaveOptions {
            update_modified_date: options.update_modified_date,
            modified_by: options.modified_by.clone(),
        },
        selective: None,
        paragraph_ids: None,
    };
    let original = unzip_parts(source).unwrap();
    let (expected, _) = write_docx_s13_parts_with_report(request, &original, None).unwrap();
    let saved = document.save_report_with_options(options).unwrap();
    assert!(saved.whole_model_write);
    assert!(saved.bytes == expected, "the save differs from the plan-less writer");
}

#[test]
fn a_mutably_borrowed_unchanged_model_saves_the_original_bytes() {
    let source = synthetic_package();
    let mut document = Document::open(&source).unwrap();
    document.model_mut();
    assert_eq!(document.save().unwrap(), source);
    let saved = document.save_report().unwrap();
    assert_eq!(saved.bytes, source);
    assert!(!saved.whole_model_write);
    assert!(saved.kept_unchanged.is_none());
    assert!(saved.structure_dropped.is_none());
}

#[test]
fn stamping_an_unchanged_model_only_updates_core_properties() {
    let source = synthetic_package();
    let mut document = Document::open(&source).unwrap();
    document.model_mut();
    let options = SaveOptions {
        now: NOW.to_owned(),
        update_modified_date: true,
        modified_by: Some("Fixture author".to_owned()),
    };
    let saved = document.save_report_with_options(options.clone()).unwrap();
    assert!(!saved.whole_model_write);
    assert!(saved.kept_unchanged.is_none());
    assert!(saved.structure_dropped.is_none());
    let mut expected = parts(&source);
    let updated = std::str::from_utf8(&expected[CORE])
        .unwrap()
        .replace(ORIGINAL_TIME, NOW)
        .replace(
            "<cp:lastModifiedBy>Before</cp:lastModifiedBy>",
            "<cp:lastModifiedBy>Fixture author</cp:lastModifiedBy>",
        );
    assert_ne!(updated.as_bytes(), expected[CORE].as_slice());
    expected.insert(CORE.to_owned(), updated.into_bytes());
    assert_eq!(parts(&saved.bytes), expected);
    assert_eq!(parts(&document.save_with_options(options).unwrap()), expected);
}

#[test]
fn a_changed_model_reports_a_whole_model_write_matching_the_plain_writer() {
    let source = synthetic_package();
    let mut document = Document::open(&source).unwrap();
    let BlockContent::Paragraph(paragraph) = &mut document.model_mut().body.content[0] else {
        panic!("expected a paragraph");
    };
    let ParagraphContent::Inline(InlineNode::Run(run)) = &mut Arc::make_mut(paragraph).content[0]
    else {
        panic!("expected a run");
    };
    let RunContent::Text { text, .. } = &mut run.content[0] else {
        panic!("expected text");
    };
    *text = "Edited".to_owned();
    assert_whole_model_write(
        &source,
        &document,
        SaveOptions {
            now: NOW.to_owned(),
            update_modified_date: true,
            modified_by: None,
        },
    );
}

fn collect_docs(root: &Path, docs: &mut Vec<PathBuf>) {
    for entry in fs::read_dir(root).unwrap() {
        let entry = entry.unwrap();
        let path = entry.path();
        if entry.file_type().unwrap().is_dir() {
            collect_docs(&path, docs);
        } else if path.extension().is_some_and(|ext| ext == "docx") {
            docs.push(path);
        }
    }
}

#[test]
fn repository_documents_save_as_source_bytes_or_write_readdressed_models() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"))
        .ancestors()
        .nth(2)
        .unwrap();
    let mut docs = [
        "apps/demo/public/betteroffice-demo.docx",
        "crates/ooxml-opc/tests/fixtures/betteroffice-demo.docx",
        "crates/ooxml-redact/tests/fixtures/redaction-integrity.docx",
        "packages/docx-react/src/components/DocxEditor/hooks/__fixtures__/probe-linked-header.docx",
    ]
    .map(|path| root.join(path))
    .to_vec();
    for dir in [
        "crates/docx-edit/tests/fixtures",
        "crates/betteroffice-docx/tests/corpus/fixtures",
        "packages/docx/src/yrs/__fixtures__",
    ] {
        collect_docs(&root.join(dir), &mut docs);
    }
    docs.sort();
    let mut unchanged = 0;
    let mut readdressed = 0;
    for path in &docs {
        let source = fs::read(path).unwrap();
        let mut document = Document::open(&source).unwrap_or_else(|error| {
            panic!("{}: {error}", path.display());
        });
        let parsed = parse_docx_s9_wire(
            &source,
            S9ParseOptions {
                source_ordinals: true,
                ..S9ParseOptions::default()
            },
        )
        .unwrap();
        let raw_content = parsed.document.package.document.content;
        if document.model_mut().body.content == raw_content {
            let saved = document.save_report().unwrap();
            assert!(saved.bytes == source, "{}: bytes changed", path.display());
            assert!(!saved.whole_model_write, "{}", path.display());
            unchanged += 1;
        } else {
            assert_whole_model_write(&source, &document, SaveOptions::default());
            readdressed += 1;
        }
    }
    assert!(
        unchanged > 0 && unchanged + readdressed == docs.len(),
        "{unchanged} unchanged, {readdressed} readdressed of {} documents",
        docs.len()
    );
}
