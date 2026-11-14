use betteroffice_docx::{Document, DroppedStructure, SaveOptions, SaveReport};

const DECLARATION: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
const NAMESPACES: &str = r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14""#;
const SECTION: &str = r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>"#;
const SMART_TAG: &str = r#"<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place"><w:r><w:t>Ancient</w:t></w:r></w:smartTag>"#;
const MAIN: &str = "word/document.xml";
/// An empty paragraph with mark formatting, whose text edit the crate saves as a whole write.
const EMPTY_MARKED: &str = r#"<w:p w14:paraId="0000000E" w14:textId="77777777"><w:pPr><w:rPr><w:b/></w:rPr></w:pPr></w:p>"#;

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
        (MAIN.to_owned(), document.as_bytes().to_vec()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn saved_main_part(package: &[u8]) -> String {
    let parts = ooxml_opc::unzip_parts(package).unwrap();
    let (_, bytes) = parts
        .iter()
        .find(|(name, _)| name == MAIN)
        .expect("word/document.xml");
    String::from_utf8(bytes.clone()).unwrap()
}

fn paragraph(id: &str, content: &str) -> String {
    format!(r#"<w:p w14:paraId="{id}" w14:textId="77777777">{content}</w:p>"#)
}

fn run(text: &str) -> String {
    format!(r#"<w:r><w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

fn smart_tag_paragraph() -> String {
    paragraph("0000000B", &format!("{}{SMART_TAG}", run("Near ")))
}

fn source() -> String {
    main_part(
        &[
            paragraph("0000000A", &run("Alpha.")),
            smart_tag_paragraph(),
            EMPTY_MARKED.to_owned(),
            paragraph("0000000D", &run("Delta.")),
        ]
        .concat(),
    )
}

fn open() -> Document {
    Document::open(&docx(&source())).unwrap()
}

fn report_of(document: &Document) -> SaveReport {
    let report = document.save_report().unwrap();
    assert_eq!(report.bytes, document.save().unwrap(), "report bytes differ from save");
    report
}

fn assert_nothing_reported(report: &SaveReport) {
    assert!(report.kept_unchanged.is_none(), "{:?}", report.kept_unchanged);
    assert!(report.structure_dropped.is_none(), "{:?}", report.structure_dropped);
}

#[test]
fn an_unedited_save_reports_nothing_kept_or_dropped() {
    let source = source();
    let package = docx(&source);
    let document = Document::open(&package).unwrap();
    let report = report_of(&document);
    assert_nothing_reported(&report);
    assert_eq!(report.bytes, package);
    assert_eq!(saved_main_part(&report.bytes), source);
}

#[test]
fn an_unedited_save_with_options_reports_the_bytes_save_with_options_writes() {
    let document = open();
    let options = SaveOptions {
        update_modified_date: true,
        modified_by: Some("Reviewer".to_owned()),
        ..SaveOptions::default()
    };
    let report = document.save_report_with_options(options.clone()).unwrap();
    assert_eq!(report.bytes, document.save_with_options(options).unwrap());
}

#[test]
fn one_edited_paragraph_reports_nothing_and_keeps_the_smart_tag() {
    let mut document = open();
    document
        .replace_paragraph_text("0000000D", "Delta edited.")
        .unwrap();
    let report = report_of(&document);
    assert_nothing_reported(&report);
    let saved = saved_main_part(&report.bytes);
    assert!(saved.contains("Delta edited."), "{saved}");
    assert!(saved.contains(&smart_tag_paragraph()), "{saved}");
}

#[test]
fn one_edited_paragraph_with_options_reports_the_bytes_save_with_options_writes() {
    let mut document = open();
    document
        .replace_paragraph_text("0000000D", "Delta edited.")
        .unwrap();
    let options = SaveOptions {
        now: "2026-01-02T03:04:05.000Z".to_owned(),
        ..SaveOptions::default()
    };
    let report = document.save_report_with_options(options.clone()).unwrap();
    assert_eq!(report.bytes, document.save_with_options(options).unwrap());
    assert_nothing_reported(&report);
}

#[test]
fn a_whole_write_without_a_save_plan_reports_nothing() {
    let mut document = open();
    document.replace_paragraph_text("0000000E", "X").unwrap();
    let report = report_of(&document);
    assert_nothing_reported(&report);
    assert!(saved_main_part(&report.bytes).contains(">X<"));
}

#[test]
fn a_whole_write_with_options_reports_the_bytes_save_with_options_writes() {
    let mut document = open();
    document.replace_paragraph_text("0000000E", "X").unwrap();
    let options = SaveOptions {
        now: "2026-01-02T03:04:05.000Z".to_owned(),
        update_modified_date: true,
        modified_by: None,
    };
    let report = document.save_report_with_options(options.clone()).unwrap();
    assert_eq!(report.bytes, document.save_with_options(options).unwrap());
    assert_nothing_reported(&report);
}

/// A body whose text edit to 0000000D leaves it out of the save plan, written whole.
const STRUCTURE_DROPPED_DOCUMENT: &str = concat!(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
    r#"<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"><w:body>"#,
    r#"<w:p w14:paraId="0000000A" w14:textId="77777777"><w:r><w:t xml:space="preserve">Near </w:t></w:r><w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place"><w:r><w:t>Ancient</w:t></w:r></w:smartTag></w:p>"#,
    r#"<w:p w14:paraId="0000000B" w14:textId="77777777"><w:customXml w:uri="urn:pr1" w:element="tagged"><w:r><w:t>Tagged</w:t></w:r></w:customXml></w:p>"#,
    r#"<w:p w14:paraId="0000000C" w14:textId="77777777"><w:bookmarkStart w:id="1" w:name="pr1mark"/><w:r><w:t>Marked</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>"#,
    r#"<w:p w14:paraId="0000000D" w14:textId="77777777"><w:pPr><mc:AlternateContent><mc:Choice Requires="w14"><w:jc w:val="center"/></mc:Choice><mc:Fallback><w:jc w:val="left"/></mc:Fallback></mc:AlternateContent></w:pPr><w:r><w:t>aaaa</w:t></w:r></w:p>"#,
    r#"<w:p><w:r><w:br w:type="page"/></w:r></w:p>"#,
    r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#,
);

#[test]
fn a_whole_write_in_the_save_plan_reports_the_structure_it_drops() {
    let mut document = Document::open(&docx(STRUCTURE_DROPPED_DOCUMENT)).unwrap();
    document.replace_paragraph_text("0000000D", "aaa").unwrap();
    let report = report_of(&document);
    assert!(report.kept_unchanged.is_none(), "{:?}", report.kept_unchanged);
    let dropped = report
        .structure_dropped
        .expect("the whole write's dropped structure is reported");
    assert_eq!(dropped.len(), 1, "{dropped:?}");
    let DroppedStructure {
        part,
        dropped: names,
        paragraphs,
    } = &dropped[0];
    assert_eq!(part, MAIN);
    assert_eq!(names, &["w:smartTag", "w:customXml"]);
    let paragraphs: Vec<(u32, &str, &str, &[String])> = paragraphs
        .iter()
        .map(|paragraph| {
            (
                paragraph.ordinal,
                paragraph.para_id.as_str(),
                paragraph.text.as_str(),
                paragraph.dropped.as_slice(),
            )
        })
        .collect();
    let smart_tag = ["w:smartTag".to_owned()];
    let custom_xml = ["w:customXml".to_owned()];
    assert_eq!(
        paragraphs,
        [
            (0, "0000000A", "Near Ancient", &smart_tag[..]),
            (1, "0000000B", "Tagged", &custom_xml[..]),
        ]
    );
}
