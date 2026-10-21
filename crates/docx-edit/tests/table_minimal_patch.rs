use docx_edit::{CellLoc, EditCtx, EditingDoc, StoryRange, TableRange, seed_from_docx};
use docx_parse::serializer::{S13SaveRequest, write_docx_s13_with_warnings};
use serde_json::{Value, json};

const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#
);

fn package(xml: &str) -> Vec<u8> {
    ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        ("word/document.xml".to_owned(), xml.as_bytes().to_vec()),
    ])
    .unwrap()
}

fn paragraph(id: u32, text: &str) -> String {
    format!(
        r#"<w:p w14:paraId="{id:08X}"><w:r w:rsidR="AABBCCDD"><w:rPr><w:noProof/></w:rPr><w:t>{text}</w:t></w:r></w:p>"#
    )
}

fn source(first: &str) -> String {
    let second = paragraph(2, "def");
    let unrelated = paragraph(3, "keep");
    let tail = paragraph(4, "tail");
    format!(
        r#"<w:document {NS}><w:body><w:tbl><w:tblPr><w:tblBorders><w:top w:val="nil" w:sz="18" w:space="3" w:color="123456"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr w14:paraId="ABCDEF01" w14:textId="ABCDEF02" w:rsidR="12345678" w:rsidTr="87654321"><w:trPr><w:gridAfter w:val="1"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:fill="AABBCC"/></w:tcPr>{first}</w:tc><w:tc><w:tcPr><w:shd w:val="clear" w:fill="DDEEFF"/></w:tcPr>{second}</w:tc></w:tr></w:tbl><w:tbl><w:tblPr><w:tblCaption w:val="caption"/><w:tblDescription w:val="description"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>{unrelated}</w:tc></w:tr></w:tbl>{tail}</w:body></w:document>"#
    )
}

fn open(xml: &str) -> (EditingDoc, Value) {
    let bytes = package(xml);
    let doc = EditingDoc::new(71);
    seed_from_docx(&doc, &bytes).unwrap();
    let parsed = docx_parse::parse_docx_s9_wire(&bytes, Default::default()).unwrap();
    let mut model = serde_json::to_value(parsed.document.package.document).unwrap();
    fn ordinals(value: &mut Value, next: &mut u32) {
        match value {
            Value::Object(object) => {
                if object.get("type").and_then(Value::as_str) == Some("paragraph") {
                    object.insert("sourceOrdinal".to_owned(), json!(*next));
                    *next += 1;
                }
                for child in object.values_mut() {
                    ordinals(child, next);
                }
            }
            Value::Array(values) => {
                for child in values {
                    ordinals(child, next);
                }
            }
            _ => {}
        }
    }
    ordinals(&mut model, &mut 0);
    (doc, model)
}

fn ctx() -> EditCtx {
    EditCtx::local("", "2026-10-21T00:00:00Z")
}

fn save(xml: &str, doc: &EditingDoc, model: Value) -> (String, Vec<String>) {
    let plan = doc.paragraph_save_plan();
    let part = plan
        .spliced_parts
        .iter()
        .find(|part| part.part == "word/document.xml")
        .expect("edited tables must retain a splice plan");
    let entry = json!({
        "part": part.part, "sha256": part.sha256,
        "paragraphs": part.paragraphs.iter().map(|(ordinal, _)| *ordinal).collect::<Vec<_>>(),
        "changed": part.changed, "removed": part.removed,
        "inserted": part.inserted.iter().map(|anchor| match anchor {
            docx_edit::SpliceAnchor::Before(ordinal) => json!({"before": ordinal}),
            docx_edit::SpliceAnchor::After(ordinal) => json!({"after": ordinal}),
        }).collect::<Vec<_>>(),
        "paragraphPatches": part.paragraph_patches, "legacyParagraphs": part.legacy_paragraphs,
        "tablePatches": part.table_patches
    });
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2026-10-21T00:00:00.000Z"},
        "document": model, "options": {"updateModifiedDate": false},
        "paragraphIds": {"splicedParts": [entry]}
    }))
    .unwrap();
    let (bytes, warnings) = write_docx_s13_with_warnings(request, &package(xml)).unwrap();
    let output = ooxml_opc::unzip_parts(&bytes)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap()
        .1;
    (String::from_utf8(output).unwrap(), warnings)
}

fn model_cell(model: &mut Value, column: usize) -> &mut Value {
    &mut model["content"][0]["rows"][0]["cells"][column]
}

fn model_text(cell: &mut Value, text: &str) {
    cell["content"][0]["content"][0]["content"][0]["text"] = json!(text);
}

fn replace(doc: &EditingDoc, cell: u32, text: &str) {
    doc.replace_range(&ctx(), StoryRange::new(format!("body:t0:r0c{cell}"), 0, 3), text)
        .unwrap();
}

#[test]
fn cell_text_preserves_both_tables_and_all_nonparagraph_bytes() {
    let original = paragraph(1, "abc");
    let xml = source(&original);
    let (doc, mut model) = open(&xml);
    replace(&doc, 0, "xyz");
    model_text(model_cell(&mut model, 0), "xyz");
    let (output, warnings) = save(&xml, &doc, model);
    assert_eq!(
        output,
        xml.replace(&original, &original.replace("abc", "xyz"))
    );
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn merge_cells_changes_only_the_merged_cells() {
    let first = paragraph(1, "abc");
    let second = paragraph(2, "def");
    let xml = source(&first);
    let (doc, mut model) = open(&xml);
    doc.merge_cells(
        &ctx(),
        &TableRange::new(CellLoc::new("body", 0, 0, 0), CellLoc::new("body", 0, 0, 1)),
    )
    .unwrap();
    let cells = model["content"][0]["rows"][0]["cells"]
        .as_array_mut()
        .unwrap();
    let removed = cells.remove(1);
    cells[0]["formatting"]["gridSpan"] = json!(2);
    cells[0]["content"]
        .as_array_mut()
        .unwrap()
        .extend(removed["content"].as_array().unwrap().iter().cloned());
    let (output, warnings) = save(&xml, &doc, model);
    let old = format!(
        r#"<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:fill="AABBCC"/></w:tcPr>{first}</w:tc><w:tc><w:tcPr><w:shd w:val="clear" w:fill="DDEEFF"/></w:tcPr>{second}</w:tc>"#
    );
    let new = format!(
        r#"<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:gridSpan w:val="2"/><w:shd w:val="clear" w:fill="AABBCC"/></w:tcPr>{first}{second}</w:tc>"#
    );
    assert_eq!(output, xml.replace(&old, &new));
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn cell_format_preserves_row_ids_grid_after_shading_and_borders() {
    let original = paragraph(1, "abc");
    let xml = source(&original);
    let (doc, mut model) = open(&xml);
    doc.format_range(
        &ctx(),
        StoryRange::new("body:t0:r0c0", 0, 3),
        &docx_edit::InlineFormatDelta {
            bold: docx_edit::Patch::Set(true),
            ..Default::default()
        },
    )
    .unwrap();
    model_cell(&mut model, 0)["content"][0]["content"][0]["formatting"]["bold"] = json!(true);
    let (output, warnings) = save(&xml, &doc, model);
    assert_eq!(
        output,
        xml.replace(
            &original,
            &original.replace("<w:noProof/>", "<w:b/><w:noProof/>"),
        )
    );
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn field_replace_keeps_the_patched_cell_neighbour_and_reports_its_reason() {
    let field = r#"<w:p w14:paraId="00000001"><w:r><w:fldChar w:fldCharType="begin"/><w:t>abc</w:t><w:fldChar w:fldCharType="end"/></w:r></w:p>"#;
    let xml = source(field);
    let (doc, mut model) = open(&xml);
    replace(&doc, 0, "xyz");
    replace(&doc, 1, "uvw");
    model_cell(&mut model, 0)["content"][0] = json!({
        "type": "paragraph", "paraId": "00000001", "sourceOrdinal": 0,
        "content": [{"type": "run", "content": [{"type": "text", "text": "xyz"}]}]
    });
    model_text(model_cell(&mut model, 1), "uvw");
    let (output, warnings) = save(&xml, &doc, model);
    assert_eq!(
        output,
        xml.replace(
            field,
            r#"<w:p w14:paraId="00000001"><w:r><w:t>xyz</w:t></w:r></w:p>"#,
        )
        .replace(&paragraph(2, "def"), &paragraph(2, "uvw"))
    );
    assert_eq!(warnings, [r#"DOCX_MINIMAL_PATCH replace counts {"field":1}"#]);
}

#[test]
fn row_insertion_keeps_cell_edits_and_clones_source_row_properties() {
    let first = paragraph(1, "abc");
    let xml = source(&first);
    let (doc, mut model) = open(&xml);
    replace(&doc, 0, "xyz");
    let receipt = doc
        .insert_row(&ctx(), &CellLoc::new("body", 0, 0, 0), true)
        .unwrap();
    model_text(model_cell(&mut model, 0), "xyz");
    let mut row = model["content"][0]["rows"][0].clone();
    for (cell, key) in row["cells"]
        .as_array_mut()
        .unwrap()
        .iter_mut()
        .zip(&receipt.new_para_ids)
    {
        let id = doc
            .paragraph_identities()
            .paragraphs
            .into_iter()
            .find_map(|identity| match identity.paragraph {
                docx_edit::ParagraphRef::Session { para_id, .. } if para_id == *key => {
                    identity.ooxml_para_id
                }
                _ => None,
            })
            .unwrap();
        cell["content"] = json!([{
            "type": "paragraph", "paraId": id, "content": []
        }]);
    }
    model["content"][0]["rows"].as_array_mut().unwrap().push(row);
    let (output, warnings) = save(&xml, &doc, model);
    let start = xml.find("<w:tr ").unwrap();
    let end = xml[start..].find("</w:tr>").unwrap() + start + "</w:tr>".len();
    assert!(output.starts_with(&xml[..start]));
    assert!(output.ends_with(&xml[end..]));
    let rows = &output[start..output.len() - xml[end..].len()];
    assert!(rows.starts_with(&xml[start..end].replace(&first, &first.replace("abc", "xyz"))));
    assert_eq!(rows.matches(r#"w:rsidTr="87654321""#).count(), 2);
    assert_eq!(
        rows.matches(r#"<w:trPr><w:gridAfter w:val="1"/></w:trPr>"#).count(),
        2
    );
    assert_eq!(
        rows.matches(r#"<w:shd w:val="clear" w:fill="AABBCC"/>"#).count(),
        2
    );
    assert_eq!(
        rows.matches(r#"<w:shd w:val="clear" w:fill="DDEEFF"/>"#).count(),
        2
    );
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn unsupported_cell_property_replaces_only_its_row_and_counts_the_key() {
    let xml = source(&paragraph(1, "abc"));
    let (doc, mut model) = open(&xml);
    doc.set_cell_shading(
        &ctx(),
        &TableRange::cell(CellLoc::new("body", 0, 0, 0)),
        Some("112233"),
    )
    .unwrap();
    model_cell(&mut model, 0)["formatting"]["shading"]["fill"] = json!({"rgb": "112233"});
    let (output, warnings) = save(&xml, &doc, model);
    let start = xml.find("<w:tr ").unwrap();
    let end = xml[start..].find("</w:tr>").unwrap() + start + "</w:tr>".len();
    assert!(output.starts_with(&xml[..start]));
    assert!(output.ends_with(&xml[end..]));
    assert!(output.contains(r#"w:fill="112233""#));
    assert_eq!(
        warnings,
        [r#"DOCX_MINIMAL_PATCH replace counts {"table:backgroundColor":1}"#]
    );
}

#[test]
fn vertical_merge_patches_restart_and_continuation_without_rewriting_rows() {
    let first = paragraph(1, "abc");
    let second = paragraph(2, "def");
    let xml = source(&first);
    let start = xml.find("<w:tr ").unwrap();
    let end = xml[start..].find("</w:tr>").unwrap() + start + "</w:tr>".len();
    let lower = xml[start..end]
        .replace(&first, &paragraph(5, "ghi"))
        .replace(&second, &paragraph(6, "jkl"));
    let xml = format!("{}{}{}", &xml[..end], lower, &xml[end..]);
    let (doc, mut model) = open(&xml);
    doc.merge_cells(
        &ctx(),
        &TableRange::new(CellLoc::new("body", 0, 0, 0), CellLoc::new("body", 0, 1, 1)),
    )
    .unwrap();
    let rows = model["content"][0]["rows"].as_array_mut().unwrap();
    let mut content = Vec::new();
    for row in rows.iter() {
        for cell in row["cells"].as_array().unwrap() {
            content.extend(cell["content"].as_array().unwrap().iter().cloned());
        }
    }
    let mut top = rows[0]["cells"][0].clone();
    top["formatting"]["gridSpan"] = json!(2);
    top["formatting"]["vMerge"] = json!("restart");
    top["content"] = json!(content);
    rows[0]["cells"] = json!([top]);
    let mut bottom = rows[1]["cells"][0].clone();
    bottom["formatting"]["gridSpan"] = json!(2);
    bottom["formatting"]["vMerge"] = json!("continue");
    bottom["content"] = json!([]);
    rows[1]["cells"] = json!([bottom]);
    let (output, warnings) = save(&xml, &doc, model);
    let row_prefix = &xml[start..xml[start..].find("<w:tc>").unwrap() + start];
    let tc_pr = r#"<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:gridSpan w:val="2"/>"#;
    let shd = r#"<w:shd w:val="clear" w:fill="AABBCC"/></w:tcPr>"#;
    let merged = format!(
        "{row_prefix}{tc_pr}<w:vMerge w:val=\"restart\"/>{shd}{first}{second}{}{}</w:tc></w:tr>{row_prefix}{tc_pr}<w:vMerge/>{shd}<w:p/></w:tc></w:tr>",
        paragraph(5, "ghi"),
        paragraph(6, "jkl"),
    );
    let tail = end + lower.len();
    assert_eq!(output, format!("{}{merged}{}", &xml[..start], &xml[tail..]));
    assert!(warnings.is_empty(), "{warnings:?}");
}

fn table_document(rows: &str, columns: usize) -> String {
    let grid = r#"<w:gridCol w:w="2000"/>"#.repeat(columns);
    let other = paragraph(90, "keep");
    let tail = paragraph(91, "tail");
    format!(
        r#"<w:document {NS}><w:body><w:tbl><w:tblPr/><w:tblGrid>{grid}</w:tblGrid>{rows}</w:tbl><w:tbl><w:tblPr><w:tblCaption w:val="caption"/><w:tblDescription w:val="description"/></w:tblPr><w:tr><w:tc>{other}</w:tc></w:tr></w:tbl>{tail}</w:body></w:document>"#
    )
}

fn table_cell(properties: &str, content: &str) -> String {
    format!("<w:tc><w:tcPr>{properties}</w:tcPr>{content}</w:tc>")
}

fn clear_continuations(model: &mut Value) {
    for row in model["content"][0]["rows"].as_array_mut().unwrap() {
        for cell in row["cells"].as_array_mut().unwrap() {
            if cell["formatting"]["vMerge"] == "continue" {
                cell["content"] = json!([]);
            }
        }
    }
}

#[test]
fn merging_vertical_columns_keeps_the_untouched_continuation_cell() {
    let top = (0..3)
        .map(|column| {
            table_cell(
                r#"<w:vMerge w:val="restart"/>"#,
                &paragraph(column + 1, ["abc", "def", "ghi"][column as usize]),
            )
        })
        .collect::<String>();
    let bottom = (0..3)
        .map(|column| {
            table_cell(
                &format!(r#"<w:vMerge/><w:shd w:fill="{}"/>"#, ["112233", "445566", "778899"][column]),
                &format!(r#"<w:p w14:paraId="{:08X}"><w:pPr><w:keepNext/></w:pPr></w:p>"#, column + 4),
            )
        })
        .collect::<Vec<_>>();
    let lower = format!("<w:tr>{}</w:tr>", bottom.concat());
    let xml = table_document(&format!("<w:tr>{top}</w:tr>{lower}"), 3);
    let (doc, mut model) = open(&xml);
    clear_continuations(&mut model);
    doc.merge_cells(
        &ctx(),
        &TableRange::new(CellLoc::new("body", 0, 0, 0), CellLoc::new("body", 0, 1, 1)),
    )
    .unwrap();
    let rows = model["content"][0]["rows"].as_array_mut().unwrap();
    for row in rows {
        let cells = row["cells"].as_array_mut().unwrap();
        let removed = cells.remove(1);
        cells[0]["formatting"]["gridSpan"] = json!(2);
        cells[0]["content"]
            .as_array_mut()
            .unwrap()
            .extend(removed["content"].as_array().unwrap().iter().cloned());
    }
    let (output, warnings) = save(&xml, &doc, model);
    let merged_lower = format!(
        "<w:tr>{}{}</w:tr>",
        bottom[0].replace("<w:vMerge/>", r#"<w:gridSpan w:val="2"/><w:vMerge/>"#),
        bottom[2],
    );
    assert!(output.contains(&merged_lower), "{output}");
    let unchanged = &xml[xml.find("</w:tbl>").unwrap()..];
    assert!(output.ends_with(unchanged));
    assert!(warnings.is_empty(), "{warnings:?}");
}
