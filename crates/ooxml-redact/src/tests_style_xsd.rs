use std::collections::HashSet;

use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use quick_xml::events::{BytesStart, Event};
use quick_xml::{NsReader, XmlVersion};

use crate::{
    Format, RedactError, RedactionOptions, redact_with_options, redact_with_report_and_options,
};

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DOCX_CONTENT_TYPES: &str = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;

fn package(parts: Vec<(&str, Vec<u8>)>) -> Vec<u8> {
    let owned: Vec<(String, Vec<u8>)> = parts
        .into_iter()
        .map(|(path, bytes)| (path.to_owned(), bytes))
        .collect();
    ooxml_opc::rezip_parts(&owned).unwrap()
}

fn xml(value: &str) -> Vec<u8> {
    value.as_bytes().to_vec()
}

fn part<'a>(parts: &'a [(String, Vec<u8>)], path: &str) -> &'a [u8] {
    parts
        .iter()
        .find(|(candidate, _)| candidate == path)
        .map(|(_, bytes)| bytes.as_slice())
        .unwrap_or_else(|| panic!("missing part {path}"))
}

fn redact_docx(styles: &str, instruction: &str) -> Result<Vec<(String, Vec<u8>)>, RedactError> {
    let document = format!(
        r#"<w:document xmlns:w="{W_NS}"><w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">{}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Result</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>"#,
        quick_xml::escape::escape(instruction)
    );
    let source = package(vec![
        ("[Content_Types].xml", xml(DOCX_CONTENT_TYPES)),
        ("word/document.xml", xml(&document)),
        ("word/styles.xml", xml(styles)),
    ]);
    let output = redact_with_options(
        &source,
        Format::Docx,
        &RedactionOptions {
            random_characters: true,
            ..Default::default()
        },
    )?;
    Ok(ooxml_opc::unzip_parts(&output).unwrap())
}

#[test]
fn docx_custom_styles_renamed_builtins_kept_and_all_references_resolve() {
    let styles_xml = format!(
        r#"<w:styles xmlns:w="{W_NS}">
            <w:latentStyles w:defLockedState="0" w:defUIPriority="99">
                <w:lsdException w:name="Normal" w:uiPriority="0" w:qFormat="1"/>
                <w:lsdException w:name="heading 1" w:uiPriority="9" w:qFormat="1"/>
                <w:lsdException w:name="Table Grid" w:uiPriority="39"/>
                <w:lsdException w:name="Grid Table 5 Dark Accent 2" w:uiPriority="50"/>
                <w:lsdException w:name="Normal (Web)" w:uiPriority="99"/>
                <w:lsdException w:name="Bespoke Body Style" w:uiPriority="10" w:qFormat="1"/>
                <w:lsdException w:name="OrphanLatentCustomStyle" w:uiPriority="11"/>
            </w:latentStyles>
            <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
                <w:name w:val="Normal"/>
            </w:style>
            <w:style w:type="paragraph" w:styleId="Heading1">
                <w:name w:val="heading 1"/>
                <w:basedOn w:val="Normal"/>
                <w:next w:val="Normal"/>
                <w:link w:val="Heading1Char"/>
            </w:style>
            <w:style w:type="character" w:styleId="Heading1Char">
                <w:name w:val="Heading 1 Char"/>
                <w:link w:val="Heading1"/>
            </w:style>
            <w:style w:type="paragraph" w:styleId="Title">
                <w:name w:val="Title"/>
            </w:style>
            <w:style w:type="paragraph" w:styleId="TOC1">
                <w:name w:val="toc 1"/>
            </w:style>
            <w:style w:type="paragraph" w:styleId="Caption">
                <w:name w:val="caption"/>
            </w:style>
            <w:style w:type="paragraph" w:styleId="ListParagraph">
                <w:name w:val="List Paragraph"/>
            </w:style>
            <w:style w:type="table" w:styleId="TableGrid">
                <w:name w:val="Table Grid"/>
            </w:style>
            <w:style w:type="character" w:styleId="Hyperlink">
                <w:name w:val="Hyperlink"/>
            </w:style>
            <w:style w:type="paragraph" w:styleId="FootnoteText">
                <w:name w:val="footnote text"/>
            </w:style>
            <w:style w:type="paragraph" w:styleId="Header">
                <w:name w:val="header"/>
            </w:style>
            <w:style w:type="paragraph" w:styleId="Footer">
                <w:name w:val="footer"/>
            </w:style>
            <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
                <w:name w:val="Default Paragraph Font"/>
            </w:style>
            <w:style w:type="table" w:default="1" w:styleId="TableNormal">
                <w:name w:val="Normal Table"/>
            </w:style>
            <w:style w:type="numbering" w:default="1" w:styleId="NoList">
                <w:name w:val="No List"/>
            </w:style>
            <w:style w:type="paragraph" w:customStyle="1" w:styleId="BespokeBodyId">
                <w:name w:val="Bespoke Body Style"/>
                <w:aliases w:val="BespokePrimaryAlias,BespokeSecondaryAlias"/>
                <w:basedOn w:val="Heading1"/>
                <w:next w:val="ImplicitCustomId"/>
                <w:link w:val="BespokeCharId"/>
            </w:style>
            <w:style w:type="character" w:customStyle="1" w:styleId="BespokeCharId">
                <w:name w:val="Bespoke Character Style"/>
                <w:basedOn w:val="DefaultParagraphFont"/>
                <w:link w:val="BespokeBodyId"/>
            </w:style>
            <w:style w:type="paragraph" w:styleId="ImplicitCustomId">
                <w:name w:val="Implicit Custom Style"/>
                <w:basedOn w:val="BespokeBodyId"/>
                <w:next w:val="Normal"/>
            </w:style>
            <w:style w:type="paragraph" w:customStyle="1" w:default="1" w:styleId="CustomDefaultOverride">
                <w:name w:val="Normal"/>
            </w:style>
            <w:style w:type="table" w:customStyle="1" w:styleId="BespokeTableId">
                <w:name w:val="Bespoke Table Style"/>
                <w:basedOn w:val="TableNormal"/>
            </w:style>
            <w:style w:type="numbering" w:customStyle="1" w:styleId="BespokeNumDef">
                <w:name w:val="Bespoke Numbering Def"/>
                <w:numStyleLink w:val="BespokeListTarget"/>
            </w:style>
            <w:style w:type="numbering" w:customStyle="1" w:styleId="BespokeListTarget">
                <w:name w:val="Bespoke List Target"/>
            </w:style>
        </w:styles>"#
    );

    let styles_with_effects_xml = format!(
        r#"<w:styles xmlns:w="{W_NS}">
            <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
                <w:name w:val="Normal"/>
            </w:style>
            <w:style w:type="paragraph" w:customStyle="1" w:styleId="BespokeBodyId">
                <w:name w:val="Bespoke Body Style"/>
                <w:basedOn w:val="Heading1"/>
                <w:next w:val="ImplicitCustomId"/>
                <w:link w:val="BespokeCharId"/>
            </w:style>
            <w:style w:type="character" w:customStyle="1" w:styleId="BespokeCharId">
                <w:name w:val="Bespoke Character Style"/>
                <w:link w:val="BespokeBodyId"/>
            </w:style>
            <w:style w:type="paragraph" w:styleId="ImplicitCustomId">
                <w:name w:val="Implicit Custom Style"/>
                <w:basedOn w:val="BespokeBodyId"/>
            </w:style>
        </w:styles>"#
    );

    let numbering_xml = format!(
        r#"<w:numbering xmlns:w="{W_NS}">
            <w:abstractNum w:abstractNumId="0">
                <w:styleLink w:val="BespokeListTarget"/>
                <w:lvl w:ilvl="0">
                    <w:pStyle w:val="BespokeBodyId"/>
                </w:lvl>
            </w:abstractNum>
            <w:num w:numId="1">
                <w:abstractNumId w:val="0"/>
            </w:num>
        </w:numbering>"#
    );

    let document_xml = format!(
        r#"<w:document xmlns:w="{W_NS}">
            <w:body>
                <w:p>
                    <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
                    <w:r><w:t>Heading text</w:t></w:r>
                </w:p>
                <w:p>
                    <w:pPr><w:pStyle w:val="BespokeBodyId"/></w:pPr>
                    <w:r>
                        <w:rPr><w:rStyle w:val="BespokeCharId"/></w:rPr>
                        <w:t>Body text</w:t>
                    </w:r>
                </w:p>
                <w:p>
                    <w:pPr><w:pStyle w:val="ImplicitCustomId"/></w:pPr>
                    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                    <w:r><w:instrText xml:space="preserve"> STYLEREF "Bespoke Body Style" \* MERGEFORMAT </w:instrText></w:r>
                    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                    <w:r><w:t>Cached heading</w:t></w:r>
                    <w:r><w:fldChar w:fldCharType="end"/></w:r>
                </w:p>
                <w:p>
                    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                    <w:r><w:instrText xml:space="preserve"> STYLEREF </w:instrText></w:r>
                    <w:r><w:instrText>&quot;Implicit Custom Style&quot; \l</w:instrText></w:r>
                    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                    <w:r><w:t>Split run field</w:t></w:r>
                    <w:r><w:fldChar w:fldCharType="end"/></w:r>
                </w:p>
                <w:p>
                    <w:fldSimple w:instr="STYLEREF &quot;heading 1&quot; \* MERGEFORMAT">
                        <w:r><w:t>Builtin field</w:t></w:r>
                    </w:fldSimple>
                </w:p>
                <w:p>
                    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                    <w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \t "Bespoke Body Style,1,Implicit Custom Style,2,heading 1,3" </w:instrText></w:r>
                    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                    <w:r><w:t>TOC entry</w:t></w:r>
                    <w:r><w:fldChar w:fldCharType="end"/></w:r>
                </w:p>
                <w:tbl>
                    <w:tblPr><w:tblStyle w:val="BespokeTableId"/></w:tblPr>
                    <w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>
                </w:tbl>
                <w:sectPr/>
            </w:body>
        </w:document>"#
    );

    let glossary_styles_xml = format!(
        r#"<w:styles xmlns:w="{W_NS}">
            <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
                <w:name w:val="Normal"/>
            </w:style>
            <w:style w:type="paragraph" w:customStyle="1" w:styleId="GlossaryCustomStyleId">
                <w:name w:val="Glossary Custom Style Name"/>
                <w:basedOn w:val="Normal"/>
            </w:style>
        </w:styles>"#
    );

    let glossary_doc_xml = format!(
        r#"<w:glossaryDocument xmlns:w="{W_NS}">
            <w:docParts>
                <w:docPart>
                    <w:docPartBody>
                        <w:p>
                            <w:pPr><w:pStyle w:val="GlossaryCustomStyleId"/></w:pPr>
                            <w:r><w:t>Glossary text</w:t></w:r>
                        </w:p>
                    </w:docPartBody>
                </w:docPart>
            </w:docParts>
        </w:glossaryDocument>"#
    );

    let source = package(vec![
        (
            "[Content_Types].xml",
            xml(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/stylesWithEffects.xml" ContentType="application/vnd.ms-word.stylesWithEffects+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/><Override PartName="/word/glossary/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#,
            ),
        ),
        (
            "_rels/.rels",
            xml(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
            ),
        ),
        ("word/document.xml", xml(&document_xml)),
        ("word/styles.xml", xml(&styles_xml)),
        ("word/stylesWithEffects.xml", xml(&styles_with_effects_xml)),
        ("word/numbering.xml", xml(&numbering_xml)),
        ("word/glossary/document.xml", xml(&glossary_doc_xml)),
        ("word/glossary/styles.xml", xml(&glossary_styles_xml)),
    ]);

    let options = RedactionOptions {
        random_characters: true,
        ..Default::default()
    };
    let (output, _) = redact_with_report_and_options(&source, Format::Docx, &options).unwrap();
    let parts = ooxml_opc::unzip_parts(&output).unwrap();

    for custom_ident in [
        "BespokeBodyId",
        "Bespoke Body Style",
        "BespokePrimaryAlias",
        "BespokeSecondaryAlias",
        "BespokeCharId",
        "Bespoke Character Style",
        "ImplicitCustomId",
        "Implicit Custom Style",
        "CustomDefaultOverride",
        "BespokeTableId",
        "Bespoke Table Style",
        "BespokeNumDef",
        "Bespoke Numbering Def",
        "BespokeListTarget",
        "Bespoke List Target",
        "OrphanLatentCustomStyle",
        "GlossaryCustomStyleId",
        "Glossary Custom Style Name",
    ] {
        for (path, bytes) in &parts {
            let text = String::from_utf8_lossy(bytes);
            assert!(
                !text.contains(custom_ident),
                "custom style identifier '{custom_ident}' survived in {path}: {text}"
            );
        }
    }

    let redacted_styles = String::from_utf8_lossy(part(&parts, "word/styles.xml"));
    for builtin in [
        "Normal",
        "Heading1",
        "heading 1",
        "Heading1Char",
        "Heading 1 Char",
        "Title",
        "TOC1",
        "toc 1",
        "Caption",
        "caption",
        "ListParagraph",
        "List Paragraph",
        "TableGrid",
        "Table Grid",
        "Hyperlink",
        "FootnoteText",
        "footnote text",
        "Header",
        "header",
        "Footer",
        "footer",
        "DefaultParagraphFont",
        "Default Paragraph Font",
        "TableNormal",
        "Normal Table",
        "NoList",
        "No List",
        "Grid Table 5 Dark Accent 2",
        "Normal (Web)",
    ] {
        assert!(
            redacted_styles.contains(builtin),
            "built-in style '{builtin}' was altered in word/styles.xml: {redacted_styles}"
        );
    }

    let (defined_ids, defined_names) = collect_defined_docx_styles(part(&parts, "word/styles.xml"));
    for check_part in [
        "word/styles.xml",
        "word/stylesWithEffects.xml",
        "word/numbering.xml",
        "word/document.xml",
    ] {
        assert_docx_style_refs_resolve(part(&parts, check_part), &defined_ids, check_part);
    }
    let (glossary_ids, _) = collect_defined_docx_styles(part(&parts, "word/glossary/styles.xml"));
    for check_part in ["word/glossary/styles.xml", "word/glossary/document.xml"] {
        assert_docx_style_refs_resolve(part(&parts, check_part), &glossary_ids, check_part);
    }

    assert_eq!(
        field_instructions(part(&parts, "word/document.xml")),
        [
            r#" STYLEREF "RedactedStyle1" \* MERGEFORMAT "#,
            " STYLEREF ",
            r#""RedactedStyle3" \l"#,
            r#"STYLEREF "heading 1" \* MERGEFORMAT"#,
            r#" TOC \o "1-3" \h \z \t "RedactedStyle1,1,RedactedStyle3,2,heading 1,3" "#,
        ]
    );
    for field_style in ["RedactedStyle1", "RedactedStyle3", "heading 1"] {
        assert!(
            defined_names.contains(&field_style.to_ascii_lowercase()),
            "field style reference '{field_style}' does not resolve to a style name in styles.xml"
        );
    }

    parse_docx_s9_wire(&output, S9ParseOptions::default()).unwrap();
}

#[test]
fn docx_field_style_references_keep_builtins_and_refuse_unknown_or_ambiguous_names() {
    let styles = format!(
        r#"<w:styles xmlns:w="{W_NS}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:customStyle="1" w:styleId="MemoId"><w:name w:val="Memo Body"/><w:aliases w:val="Shared Alias"/></w:style><w:style w:type="paragraph" w:customStyle="1" w:styleId="OtherId"><w:name w:val="Shared Alias"/></w:style></w:styles>"#
    );
    for (instruction, expected) in [
        (
            r#" STYLEREF "heading 1" \l "#,
            r#" STYLEREF "heading 1" \l "#,
        ),
        (
            r#" STYLEREF Memo Body \n"#,
            r#" STYLEREF "RedactedStyle1" \n"#,
        ),
        (
            r#" TOC \o "1-3" \h \z \t "heading 1;1;memo body;2" "#,
            r#" TOC \o "1-3" \h \z \t "heading 1;1;RedactedStyle1;2" "#,
        ),
    ] {
        let parts = redact_docx(&styles, instruction).unwrap();
        assert_eq!(
            field_instructions(part(&parts, "word/document.xml")),
            [expected],
            "{instruction}"
        );
    }
    for unreadable in [r#" TOC \c "Figure" \h "#, r#" STYLEREF "Unclosed "#] {
        let parts = redact_docx(&styles, unreadable).unwrap();
        let [instruction] = &field_instructions(part(&parts, "word/document.xml"))[..] else {
            panic!("{unreadable}");
        };
        assert!(
            instruction.trim_start_matches(['0', 'x']).is_empty(),
            "{instruction}"
        );
    }
    for name in ["Überschrift 1", "Undefined Name", "Shared Alias"] {
        let error = redact_docx(&styles, &format!(r#"STYLEREF "{name}""#)).unwrap_err();
        assert!(matches!(error, RedactError::StyleReference { .. }));
    }
    let error = redact_docx(&styles, r#" STYLEREF "Shared Alias" "#).unwrap_err();
    assert!(
        matches!(error, RedactError::StyleReference { .. }),
        "{error:?}"
    );
}

#[test]
fn xlsx_custom_cell_table_and_pivot_styles_renamed_builtins_kept() {
    let styles_xml = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
        <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
        <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
        <cellStyleXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
        <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
        <cellStyles count="3">
            <cellStyle name="Normal" xfId="0" builtinId="0"/>
            <cellStyle name="CustomCurrencyStyle" xfId="1"/>
            <cellStyle name="CustomHighlightStyle" xfId="2"/>
        </cellStyles>
        <tableStyles count="2" defaultTableStyle="BespokeTableStyle" defaultPivotStyle="BespokePivotStyle">
            <tableStyle name="BespokeTableStyle" pivot="0" count="0"/>
            <tableStyle name="BespokePivotStyle" pivot="1" count="0"/>
        </tableStyles>
    </styleSheet>"#;

    let table1_xml = r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="A1:B2">
        <tableColumns count="2"><tableColumn id="1" name="C1"/><tableColumn id="2" name="C2"/></tableColumns>
        <tableStyleInfo name="BespokeTableStyle" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
    </table>"#;

    let table2_xml = r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="2" name="T2" displayName="T2" ref="D1:E2">
        <tableColumns count="2"><tableColumn id="1" name="C1"/><tableColumn id="2" name="C2"/></tableColumns>
        <tableStyleInfo name="TableStyleMedium9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
    </table>"#;

    let pivot_xml = r#"<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P1" cacheId="1" dataCaption="Values">
        <location ref="A5:B8" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>
        <pivotTableStyleInfo name="BespokePivotStyle" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>
    </pivotTableDefinition>"#;

    let source = package(vec![
        (
            "[Content_Types].xml",
            xml(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/></Types>"#,
            ),
        ),
        (
            "_rels/.rels",
            xml(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
            ),
        ),
        (
            "xl/workbook.xml",
            xml(
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            xml(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#,
            ),
        ),
        (
            "xl/worksheets/sheet1.xml",
            xml(
                r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#,
            ),
        ),
        ("xl/styles.xml", xml(styles_xml)),
        ("xl/tables/table1.xml", xml(table1_xml)),
        ("xl/tables/table2.xml", xml(table2_xml)),
        ("xl/pivotTables/pivotTable1.xml", xml(pivot_xml)),
    ]);

    let output = redact_with_options(
        &source,
        Format::Xlsx,
        &RedactionOptions {
            random_characters: true,
            ..Default::default()
        },
    )
    .unwrap();
    let parts = ooxml_opc::unzip_parts(&output).unwrap();

    for custom_ident in [
        "CustomCurrencyStyle",
        "CustomHighlightStyle",
        "BespokeTableStyle",
        "BespokePivotStyle",
    ] {
        for (path, bytes) in &parts {
            let text = String::from_utf8_lossy(bytes);
            assert!(
                !text.contains(custom_ident),
                "custom XLSX style '{custom_ident}' survived in {path}: {text}"
            );
        }
    }

    let styles_out = String::from_utf8_lossy(part(&parts, "xl/styles.xml"));
    let table1_out = String::from_utf8_lossy(part(&parts, "xl/tables/table1.xml"));
    let table2_out = String::from_utf8_lossy(part(&parts, "xl/tables/table2.xml"));
    let pivot_out = String::from_utf8_lossy(part(&parts, "xl/pivotTables/pivotTable1.xml"));

    assert!(styles_out.contains(r#"cellStyle name="Normal" xfId="0" builtinId="0""#));
    assert!(styles_out.contains(r#"cellStyle name="RedactedCellStyle1" xfId="1""#));
    assert!(styles_out.contains(r#"cellStyle name="RedactedCellStyle2" xfId="2""#));
    assert!(styles_out.contains(r#"defaultTableStyle="RedactedTableStyle1""#));
    assert!(styles_out.contains(r#"defaultPivotStyle="RedactedTableStyle2""#));
    assert!(styles_out.contains(r#"tableStyle name="RedactedTableStyle1""#));
    assert!(styles_out.contains(r#"tableStyle name="RedactedTableStyle2""#));
    assert!(table1_out.contains(r#"tableStyleInfo name="RedactedTableStyle1""#));
    assert!(table2_out.contains(r#"tableStyleInfo name="TableStyleMedium9""#));
    assert!(pivot_out.contains(r#"pivotTableStyleInfo name="RedactedTableStyle2""#));

    xlsx_parse::parse_workbook(&parts).unwrap();
}

fn attribute(reader: &NsReader<&[u8]>, start: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    start
        .attributes()
        .flatten()
        .find(|attr| attr.key.local_name().as_ref() == name)
        .map(|attr| {
            attr.decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .unwrap()
                .into_owned()
        })
}

fn collect_defined_docx_styles(bytes: &[u8]) -> (HashSet<String>, HashSet<String>) {
    let mut reader = NsReader::from_reader(bytes);
    let mut ids = HashSet::new();
    let mut names = HashSet::new();
    let mut in_style = false;
    loop {
        match reader.read_event().unwrap() {
            Event::Start(start) | Event::Empty(start) => match start.local_name().as_ref() {
                b"style" => {
                    in_style = true;
                    ids.extend(attribute(&reader, &start, b"styleId"));
                }
                b"name" if in_style => {
                    names.extend(
                        attribute(&reader, &start, b"val").map(|name| name.to_ascii_lowercase()),
                    );
                }
                _ => {}
            },
            Event::End(end) if end.local_name().as_ref() == b"style" => in_style = false,
            Event::Eof => return (ids, names),
            _ => {}
        }
    }
}

fn assert_docx_style_refs_resolve(bytes: &[u8], defined_ids: &HashSet<String>, part_name: &str) {
    let mut reader = NsReader::from_reader(bytes);
    loop {
        match reader.read_event().unwrap() {
            Event::Start(start) | Event::Empty(start) => {
                let local = String::from_utf8_lossy(start.local_name().as_ref()).into_owned();
                if !matches!(
                    local.as_str(),
                    "pStyle"
                        | "rStyle"
                        | "tblStyle"
                        | "basedOn"
                        | "next"
                        | "link"
                        | "numStyleLink"
                        | "styleLink"
                ) {
                    continue;
                }
                let target = attribute(&reader, &start, b"val").unwrap();
                assert!(
                    defined_ids.contains(&target),
                    "{local}/@val='{target}' in {part_name} resolves to no styleId in {defined_ids:?}"
                );
            }
            Event::Eof => return,
            _ => {}
        }
    }
}

fn field_instructions(document: &[u8]) -> Vec<String> {
    let mut reader = NsReader::from_reader(document);
    let mut instructions = Vec::new();
    let mut inside = false;
    loop {
        match reader.read_event().unwrap() {
            Event::Start(start) if start.local_name().as_ref() == b"instrText" => {
                inside = true;
                instructions.push(String::new());
            }
            Event::Start(start) | Event::Empty(start)
                if start.local_name().as_ref() == b"fldSimple" =>
            {
                instructions.extend(attribute(&reader, &start, b"instr"));
            }
            Event::Text(text) if inside => {
                let decoded = text.decode().unwrap();
                let text = quick_xml::escape::unescape(&decoded).unwrap().into_owned();
                instructions.last_mut().unwrap().push_str(&text);
            }
            Event::GeneralRef(reference) if inside => {
                let entity = format!("&{};", reference.decode().unwrap());
                let text = quick_xml::escape::unescape(&entity).unwrap().into_owned();
                instructions.last_mut().unwrap().push_str(&text);
            }
            Event::End(_) => inside = false,
            Event::Eof => return instructions,
            _ => {}
        }
    }
}

#[test]
fn p1_field_literal_and_unresolved_style_are_not_preserved() {
    let styles = format!(r#"<w:styles xmlns:w="{W_NS}"/>"#);
    let parts = redact_docx(&styles, r#"TOC \o "1-3" \p "SyntheticPrivate""#).unwrap();
    assert!(
        !field_instructions(part(&parts, "word/document.xml"))
            .join("")
            .contains("SyntheticPrivate"),
        "field literal survived"
    );
    assert!(
        matches!(
            redact_docx(&styles, r#"STYLEREF "SecretClient""#),
            Err(RedactError::StyleReference { .. })
        ),
        "unresolved style must return StyleReference"
    );
}

#[test]
fn p3_customized_builtin_cell_style_is_renamed() {
    let parts = vec![(
        "xl/styles.xml".to_owned(),
        xml(
            r#"<styleSheet><cellStyles><cellStyle name="SyntheticConfidential" xfId="1" builtinId="4" customBuiltin="1"/></cellStyles></styleSheet>"#,
        ),
    )];
    let map = crate::styles::StyleMap::collect_xlsx(&parts).unwrap();
    assert!(
        map.xlsx_replacement("cellStyle", "name", "SyntheticConfidential", true)
            .is_some(),
        "customized builtin name survived"
    );
}

#[test]
fn p10_repeated_toc_switches_are_linear() {
    let styles = format!(
        r#"<w:styles xmlns:w="{W_NS}"><w:style w:type="paragraph" w:customStyle="1" w:styleId="S"><w:name w:val="S"/></w:style></w:styles>"#
    );
    let map = crate::styles::StyleMap::collect("word/styles.xml", styles.as_bytes()).unwrap();
    let instruction = format!("TOC {}", "\\t \"S,1\" ".repeat(500_000));
    let start = std::time::Instant::now();
    let output = map
        .rewrite_field_instruction("word/document.xml", &instruction)
        .unwrap()
        .unwrap();
    assert_eq!(output.matches("RedactedStyle1,1").count(), 500_000);
    assert!(
        start.elapsed() < std::time::Duration::from_secs(20),
        "repeated TOC rewrites exceeded 20 seconds: {:?}",
        start.elapsed()
    );
}

#[test]
fn case_insensitive_style_references_match_renamed_definitions() {
    let references = [
        "basedOn",
        "next",
        "link",
        "pStyle",
        "rStyle",
        "tblStyle",
        "numStyleLink",
        "styleLink",
    ];
    let styles = format!(
        r#"<w:styles xmlns:w="{W_NS}"><w:style w:customStyle="1" w:styleId="ClientBody"><w:name w:val="Client Body"/></w:style><w:style w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:styleId="ClientHeading"><w:name w:val="heading 1"/></w:style></w:styles>"#
    );
    let body: String = references
        .iter()
        .flat_map(|element| {
            [
                "ClientBody",
                "clientbody",
                "CLIENTBODY",
                "cLiEnTbOdY",
                "clientheading",
            ]
            .map(|id| format!(r#"<w:{element} w:val="{id}"/>"#))
        })
        .collect();
    let document = format!(
        r#"<w:document xmlns:w="{W_NS}"><w:body><w:p>{body}<w:pStyle w:val="nOrMaL"/><w:rStyle w:val="MissingClient"/></w:p></w:body></w:document>"#
    );
    let source = package(vec![
        ("[Content_Types].xml", xml(DOCX_CONTENT_TYPES)),
        ("word/document.xml", xml(&document)),
        ("word/styles.xml", xml(&styles)),
    ]);
    for random_characters in [false, true] {
        let output = redact_with_options(
            &source,
            Format::Docx,
            &RedactionOptions {
                random_characters,
                ..Default::default()
            },
        )
        .unwrap();
        let parts = ooxml_opc::unzip_parts(&output).unwrap();
        for (path, bytes) in &parts {
            let text = String::from_utf8_lossy(bytes);
            assert!(
                !text.to_ascii_lowercase().contains("client"),
                "{path}: {text}"
            );
        }
        let document = String::from_utf8_lossy(part(&parts, "word/document.xml"));
        for element in references {
            assert_eq!(
                document
                    .matches(&format!(r#"<w:{element} w:val="RedactedStyle1"/>"#))
                    .count(),
                4,
                "{document}"
            );
            assert!(
                document.contains(&format!(r#"<w:{element} w:val="RedactedStyle2"/>"#)),
                "{document}"
            );
        }
        assert!(
            document.contains(r#"<w:pStyle w:val="nOrMaL"/>"#),
            "{document}"
        );
        assert!(
            document.contains(r#"<w:rStyle w:val="UnresolvedStyle"/>"#),
            "{document}"
        );
    }
}

#[test]
fn undefined_style_references_and_list_names_are_neutral() {
    let styles = format!(
        r#"<w:styles xmlns:w="{W_NS}"><w:style w:type="paragraph" w:customStyle="1" w:styleId="ClientBody"><w:name w:val="Client Body"/></w:style><w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/></w:style></w:styles>"#
    );
    let document = format!(
        r#"<w:document xmlns:w="{W_NS}"><w:body><w:p><w:pPr><w:pStyle w:val="ClientBody"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>a</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="ClientLegacyHeading"/></w:pPr><w:r><w:rPr><w:rStyle w:val="ClientLegacyEmphasis"/></w:rPr><w:t>b</w:t></w:r></w:p></w:body></w:document>"#
    );
    let numbering = format!(
        r#"<w:numbering xmlns:w="{W_NS}"><w:abstractNum w:abstractNumId="0"><w:name w:val="ClientClauseList"/><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
    );
    let source = package(vec![
        ("[Content_Types].xml", xml(DOCX_CONTENT_TYPES)),
        ("word/document.xml", xml(&document)),
        ("word/styles.xml", xml(&styles)),
        ("word/numbering.xml", xml(&numbering)),
    ]);
    let output = redact_with_options(
        &source,
        Format::Docx,
        &RedactionOptions {
            random_characters: true,
            ..Default::default()
        },
    )
    .unwrap();
    let parts = ooxml_opc::unzip_parts(&output).unwrap();
    let joined: String = parts
        .iter()
        .map(|(_, bytes)| String::from_utf8_lossy(bytes).into_owned())
        .collect();
    assert!(
        !joined.contains("Client"),
        "a source name survived: {joined}"
    );
    let document = String::from_utf8_lossy(part(&parts, "word/document.xml")).into_owned();
    assert!(
        document.contains(r#"<w:rStyle w:val="Strong"/>"#),
        "{document}"
    );
    let (defined, _) = collect_defined_docx_styles(part(&parts, "word/styles.xml"));
    let first = document.split("</w:p>").next().unwrap();
    let id = first
        .split(r#"<w:pStyle w:val=""#)
        .nth(1)
        .and_then(|rest| rest.split('"').next())
        .unwrap();
    assert!(defined.contains(id), "{id} not in {defined:?}");
}

const REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn custom_xml_fixture(format: Format) -> Vec<u8> {
    let (main, content_type, document) = match format {
        Format::Docx => (
            "word/document.xml",
            "wordprocessingml.document",
            format!(
                r#"<w:document xmlns:w="{W_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w:body><w:sdt><w:sdtPr><w:dataBinding w:xpath="/c:Record" w:prefixMappings="xmlns:c='urn:synthetic'"/><w14:dataBinding w:xpath="/Record"/><w15:dataBinding w:xpath="/Record"></w15:dataBinding><w:docPartPr><w:docPartUnique/></w:docPartPr></w:sdtPr><w:sdtContent><w:p><w:r><w:t>CachedSyntheticText</w:t></w:r></w:p></w:sdtContent></w:sdt><w:sectPr/></w:body></w:document>"#
            ),
        ),
        Format::Xlsx => (
            "xl/workbook.xml",
            "spreadsheetml.sheet",
            format!(
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{OFFICE_REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="sheet"/></sheets></workbook>"#
            ),
        ),
        Format::Pptx => (
            "ppt/presentation.xml",
            "presentationml.presentation",
            format!(
                r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="{OFFICE_REL}"><p:sldIdLst><p:sldId id="256" r:id="slide"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#
            ),
        ),
        _ => unreachable!(),
    };
    let (directory, file) = main.rsplit_once('/').unwrap();
    let owner_rels = format!("{directory}/_rels/{file}.rels");
    let extra_rel = match format {
        Format::Xlsx => format!(
            r#"<Relationship Id="sheet" Type="{OFFICE_REL}/worksheet" Target="worksheets/sheet1.xml"/>"#
        ),
        Format::Pptx => format!(
            r#"<Relationship Id="slide" Type="{OFFICE_REL}/slide" Target="slides/slide1.xml"/>"#
        ),
        _ => String::new(),
    };
    let mut parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            xml(&format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/{main}" ContentType="application/vnd.openxmlformats-officedocument.{content_type}.main+xml"/><Override PartName="/customXml/item1.xml" ContentType="application/xml"/><Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/><Override PartName="/customXml/item2.xml" ContentType="application/xml"/></Types>"#
            )),
        ),
        (
            "_rels/.rels".to_owned(),
            xml(&format!(
                r#"<Relationships xmlns="{REL_NS}"><Relationship Id="main" Type="{OFFICE_REL}/officeDocument" Target="{main}"/><Relationship Id="customRoot" Type="{OFFICE_REL}/customXml" Target="customXml/item2.xml"/></Relationships>"#
            )),
        ),
        (main.to_owned(), xml(&document)),
        (
            owner_rels,
            xml(&format!(
                r#"<Relationships xmlns="{REL_NS}">{extra_rel}<Relationship Id="custom" Type="{OFFICE_REL}/customXml" Target="../customXml/item1.xml"/><Relationship Id="props" Type="{OFFICE_REL}/customXmlProps" Target="../customXml/itemProps1.xml"/></Relationships>"#
            )),
        ),
        (
            "customXml/item1.xml".to_owned(),
            xml(
                r#"<SyntheticPrefix:Record xmlns:SyntheticPrefix="urn:synthetic" SyntheticAttribute="SyntheticPayload"/>"#,
            ),
        ),
        (
            "customXml/itemProps1.xml".to_owned(),
            xml(
                r#"<ds:datastoreItem xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml" ds:itemID="{00000000-0000-0000-0000-000000000001}"><ds:schemaRefs><ds:schemaRef ds:uri="urn:SyntheticSchema"/></ds:schemaRefs></ds:datastoreItem>"#,
            ),
        ),
        (
            "customXml/_rels/item1.xml.rels".to_owned(),
            xml(&format!(
                r#"<Relationships xmlns="{REL_NS}"><Relationship Id="props" Type="{OFFICE_REL}/customXmlProps" Target="itemProps1.xml"/><Relationship Id="schema" Type="{OFFICE_REL}/customXml" Target="item2.xml"/></Relationships>"#
            )),
        ),
        (
            "customXml/_rels/itemProps1.xml.rels".to_owned(),
            xml(&format!(
                r#"<Relationships xmlns="{REL_NS}"><Relationship Id="item" Type="{OFFICE_REL}/customXml" Target="item1.xml"/></Relationships>"#
            )),
        ),
        (
            "customXml/item2.xml".to_owned(),
            xml(
                r#"<xs:schema xmlns:xs="http://www.w3.org/2001/XMLSche&#109;a" targetNamespace="urn:SyntheticSchema"><xs:element name="SyntheticDeclaration" type="xs:string"/></xs:schema>"#,
            ),
        ),
    ];
    if format == Format::Xlsx {
        parts.push(("xl/worksheets/sheet1.xml".to_owned(), xml(r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"#)));
    }
    if format == Format::Pptx {
        parts.push(("ppt/slides/slide1.xml".to_owned(), xml(r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>"#)));
    }
    if format == Format::Docx {
        parts.push(("word/header1.xml".to_owned(), xml(&format!(r#"<w:hdr xmlns:w="{W_NS}"><w:sdt><w:sdtPr><w:dataBinding w:xpath="/Record"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>CachedHeader</w:t></w:r></w:p></w:sdtContent></w:sdt></w:hdr>"#))));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn assert_custom_xml_removed(format: Format) {
    for random_characters in [false, true] {
        let (output, _) = redact_with_report_and_options(
            &custom_xml_fixture(format),
            format,
            &RedactionOptions {
                random_characters,
                ..Default::default()
            },
        )
        .unwrap();
        let parts = ooxml_opc::unzip_parts(&output).unwrap();
        assert!(
            !parts
                .iter()
                .any(|(path, _)| path.to_ascii_lowercase().starts_with("customxml/")),
            "customXml data, properties or relationships survived"
        );
        for (path, bytes) in &parts {
            let text = String::from_utf8_lossy(bytes);
            assert!(
                !text.to_ascii_lowercase().contains("customxml/"),
                "customXml target or override survived in {path}"
            );
            for sentinel in [
                "SyntheticPrefix",
                "SyntheticAttribute",
                "SyntheticSchema",
                "SyntheticDeclaration",
                "SyntheticPayload",
            ] {
                assert!(
                    !text.contains(sentinel),
                    "customXml content survived in {path}: {sentinel}"
                );
            }
            if path.ends_with(".rels") {
                let mut reader = NsReader::from_reader(bytes.as_slice());
                loop {
                    let event = reader.read_event().unwrap();
                    match event {
                        Event::Start(start) | Event::Empty(start)
                            if start.local_name().as_ref() == b"Relationship" =>
                        {
                            let target = attribute(&reader, &start, b"Target").unwrap();
                            let resolved =
                                crate::scrub::resolve_relationship_target(path, &target).unwrap();
                            assert!(
                                parts
                                    .iter()
                                    .any(|(name, _)| crate::scrub::normalize_part_name(name)
                                        == resolved),
                                "dangling relationship in {path}: {target}"
                            );
                        }
                        Event::Eof => break,
                        _ => {}
                    }
                }
            }
        }
        match format {
            Format::Docx => {
                parse_docx_s9_wire(&output, S9ParseOptions::default()).unwrap();
            }
            Format::Xlsx => {
                xlsx_parse::parse_workbook(&parts).unwrap();
            }
            Format::Pptx => {
                pptx_parse::parse_pptx(&output).unwrap();
            }
            _ => unreachable!(),
        }
    }
}

#[test]
fn custom_xml_removed_docx_reparses_without_dangling_relationships() {
    assert_custom_xml_removed(Format::Docx);
}

#[test]
fn custom_xml_removed_xlsx_reparses_without_dangling_relationships() {
    assert_custom_xml_removed(Format::Xlsx);
}

#[test]
fn custom_xml_removed_pptx_reparses_without_dangling_relationships() {
    assert_custom_xml_removed(Format::Pptx);
}

#[test]
fn bound_sdt_keeps_redacted_cached_text_without_data_bindings() {
    let output = redact_with_options(
        &custom_xml_fixture(Format::Docx),
        Format::Docx,
        &RedactionOptions::default(),
    )
    .unwrap();
    let parts = ooxml_opc::unzip_parts(&output).unwrap();
    for path in ["word/document.xml", "word/header1.xml"] {
        let text = String::from_utf8_lossy(part(&parts, path));
        assert!(
            !text.contains("dataBinding"),
            "data binding survived in {path}"
        );
        assert!(
            text.contains("w:sdtContent") && text.contains("<w:t>"),
            "cached content removed in {path}"
        );
        let expected = if path == "word/document.xml" {
            "CachedSyntheticText"
        } else {
            "CachedHeader"
        };
        assert!(
            text.contains(&format!("<w:t>{}</w:t>", "x".repeat(expected.len()))),
            "cached text changed length in {path}"
        );
        assert!(
            !text.contains("Cached"),
            "cached content not redacted in {path}"
        );
    }
    assert!(String::from_utf8_lossy(part(&parts, "word/document.xml")).contains("w:docPartPr"));
}

fn assert_undefined_refs_neutral(styles: Option<&str>) {
    let refs: String = [
        "pStyle",
        "rStyle",
        "tblStyle",
        "numStyleLink",
        "styleLink",
        "basedOn",
        "next",
        "link",
    ]
    .iter()
    .map(|name| format!(r#"<w:{name} w:val="SyntheticUndefined{name}"/>"#))
    .collect();
    let document = format!(
        r#"<w:document xmlns:w="{W_NS}"><w:body><w:p><w:pPr>{refs}</w:pPr><w:r><w:t>Text</w:t></w:r></w:p></w:body></w:document>"#
    );
    let mut parts = vec![
        ("[Content_Types].xml", xml(DOCX_CONTENT_TYPES)),
        ("word/document.xml", xml(&document)),
    ];
    if let Some(styles) = styles {
        parts.push(("word/styles.xml", xml(styles)));
    }
    let output =
        redact_with_options(&package(parts), Format::Docx, &RedactionOptions::default()).unwrap();
    let parts = ooxml_opc::unzip_parts(&output).unwrap();
    let document = String::from_utf8_lossy(part(&parts, "word/document.xml"));
    assert!(
        !document.contains("SyntheticUndefined"),
        "undefined style reference survived: {document}"
    );
    assert_eq!(document.matches("UnresolvedStyle").count(), 8);
}

#[test]
fn undefined_style_refs_with_empty_styles_are_neutral() {
    assert_undefined_refs_neutral(Some(&format!(r#"<w:styles xmlns:w="{W_NS}"/>"#)));
}

#[test]
fn undefined_style_refs_with_missing_styles_are_neutral() {
    assert_undefined_refs_neutral(None);
}

#[test]
fn large_style_alias_merging_completes_in_bounded_time() {
    let (sender, receiver) = std::sync::mpsc::channel();
    std::thread::spawn(move || {
        let parts: Vec<_> = ["A", "B"].iter().zip(["word/styles.xml", "word/styleswitheffects.xml"]).map(|(prefix, path)| {
            let aliases = (0..100_000).map(|n| format!("{prefix}{n:06}")).collect::<Vec<_>>().join(",");
            (path.to_owned(), xml(&format!(r#"<w:styles xmlns:w="{W_NS}"><w:style w:styleId="S" w:customStyle="1"><w:name w:val="SyntheticStyle"/><w:aliases w:val="{aliases}"/></w:style></w:styles>"#)))
        }).collect();
        let map = crate::styles::StyleMap::collect_docx(
            &parts,
            &["word/styles.xml", "word/styleswitheffects.xml"],
        )
        .unwrap();
        for alias in ["A000000", "A099999", "B000000", "B099999"] {
            assert!(
                map.rewrite_field_instruction("word/document.xml", &format!("STYLEREF {alias}"))
                    .unwrap()
                    .unwrap()
                    .contains("RedactedStyle1")
            );
        }
        sender.send(()).unwrap();
    });
    assert!(
        receiver
            .recv_timeout(std::time::Duration::from_secs(20))
            .is_ok(),
        "merging 200000 style aliases exceeded 20 seconds"
    );
}

#[test]
fn custom_xml_parts_outside_conventional_directory_are_removed() {
    for format in [Format::Docx, Format::Xlsx, Format::Pptx] {
        let parts = ooxml_opc::unzip_parts(&custom_xml_fixture(format)).unwrap();
        let relocated: Vec<_> = parts
            .into_iter()
            .map(|(path, bytes)| {
                (
                    path.replace("customXml/", "metadata/"),
                    String::from_utf8(bytes)
                        .unwrap()
                        .replace("customXml/", "metadata/")
                        .into_bytes(),
                )
            })
            .collect();
        let source = ooxml_opc::rezip_parts(&relocated).unwrap();
        let output = redact_with_options(&source, format, &RedactionOptions::default()).unwrap();
        let parts = ooxml_opc::unzip_parts(&output).unwrap();
        assert!(
            !parts.iter().any(|(path, _)| path.starts_with("metadata/")),
            "relocated customXml part survived"
        );
        for (path, bytes) in &parts {
            assert!(
                !String::from_utf8_lossy(bytes).contains("metadata/"),
                "relocated target or override survived in {path}"
            );
        }
    }
}
