use docx_edit::{
    CellLoc, EditCtx, EditingDoc, Position, RawOp, StoryRange, TableLocator, TableRange,
    seed_from_docx,
};
use docx_parse::serializer::{S13SaveRequest, write_docx_s13_with_warnings};
use serde_json::{Value, json};
use yrs::Any;

const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#
);

/// The package of `xml`, with a comments part holding every comment it references.
fn package(xml: &str) -> Vec<u8> {
    let comments: String = xml
        .split(r#"<w:commentReference w:id=""#)
        .skip(1)
        .filter_map(|rest| rest.split('"').next())
        .map(|id| {
            format!(
                r#"<w:comment w:id="{id}" w:author="Ada"><w:p><w:r><w:t>Note</w:t></w:r></w:p></w:comment>"#
            )
        })
        .collect();
    let override_comments = if comments.is_empty() {
        ""
    } else {
        r#"<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>"#
    };
    let mut parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{override_comments}</Types>"#
            )
            .into_bytes(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        ("word/document.xml".to_owned(), xml.as_bytes().to_vec()),
    ];
    if !comments.is_empty() {
        parts.push((
            "word/_rels/document.xml.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="comments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>"#.to_vec(),
        ));
        parts.push((
            "word/comments.xml".to_owned(),
            format!(
                r#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{comments}</w:comments>"#
            )
            .into_bytes(),
        ));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
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
        r#"<w:document {NS}><w:body><w:tbl><w:tblPr><w:tblBorders><w:top w:val="nil" w:sz="18" w:space="3" w:color="123456"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr w14:paraId="ABCDEF01" w14:textId="ABCDEF02" w:rsidR="12345678" w:rsidTr="87654321"><w:trPr/><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:shd w:val="clear" w:fill="AABBCC"/></w:tcPr>{first}</w:tc><w:tc><w:tcPr><w:shd w:val="clear" w:fill="DDEEFF"/></w:tcPr>{second}</w:tc></w:tr></w:tbl><w:tbl><w:tblPr><w:tblCaption w:val="caption"/><w:tblDescription w:val="description"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>{unrelated}</w:tc></w:tr></w:tbl>{tail}</w:body></w:document>"#
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
    try_save(xml, doc, model).unwrap()
}

fn try_save(xml: &str, doc: &EditingDoc, model: Value) -> Result<(String, Vec<String>), String> {
    save_model(xml, doc, model, save_entry(doc))
}

fn save_entry(doc: &EditingDoc) -> Option<Value> {
    let plan = doc.paragraph_save_plan();
    plan.spliced_parts
        .iter()
        .find(|part| part.part == "word/document.xml")
        .map(|part| {
            json!({
                "part": part.part, "sha256": part.sha256,
                "paragraphs": part.paragraphs.iter().map(|(ordinal, _)| *ordinal).collect::<Vec<_>>(),
                "changed": part.changed, "removed": part.removed,
                "inserted": part.inserted.iter().map(|anchor| match anchor {
                    docx_edit::SpliceAnchor::Before(ordinal) => json!({"before": ordinal}),
                    docx_edit::SpliceAnchor::After(ordinal) => json!({"after": ordinal}),
                }).collect::<Vec<_>>(),
                "paragraphPatches": part.paragraph_patches, "legacyParagraphs": part.legacy_paragraphs,
                "tablePatches": part.table_patches
            })
        })
}

fn save_model(
    xml: &str,
    doc: &EditingDoc,
    mut model: Value,
    entry: Option<Value>,
) -> Result<(String, Vec<String>), String> {
    if let Some(sections) = model["sections"].as_array() {
        let sections: Vec<_> = sections
            .iter()
            .map(|section| {
                let content = model["content"].as_array().unwrap();
                let start = section["contentStart"].as_u64().unwrap() as usize;
                let end = section["contentEnd"].as_u64().unwrap() as usize;
                json!({
                    "id": section["id"],
                    "properties": section["properties"],
                    "content": &content[start..end]
                })
            })
            .collect();
        model["sections"] = json!(sections);
    }
    let patched_parts: Vec<_> = doc
        .paragraph_save_plan()
        .patched_parts
        .iter()
        .map(|(part, para_ids)| json!({"part": part, "paraIds": para_ids}))
        .collect();
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2026-10-21T00:00:00.000Z"},
        "document": model, "options": {"updateModifiedDate": false},
        "paragraphIds": {
            "patchedParts": patched_parts,
            "splicedParts": entry.into_iter().collect::<Vec<_>>()
        }
    }))
    .unwrap();
    let (bytes, warnings) =
        write_docx_s13_with_warnings(request, &package(xml)).map_err(|error| error.to_string())?;
    let output = ooxml_opc::unzip_parts(&bytes)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap()
        .1;
    Ok((String::from_utf8(output).unwrap(), warnings))
}

fn save_with_base_fallback(xml: &str, doc: &EditingDoc, model: Value) -> (String, Vec<String>) {
    let expected = save_model(xml, doc, model.clone(), None).unwrap();
    let actual = save(xml, doc, model);
    assert_eq!(actual, expected);
    actual
}

fn save_with_base_row_splice(xml: &str, doc: &EditingDoc, model: Value) -> (String, Vec<String>) {
    let mut entry = save_entry(doc).unwrap();
    entry["tablePatches"] = json!([]);
    let expected = save_model(xml, doc, model.clone(), Some(entry)).unwrap();
    let actual = save(xml, doc, model);
    assert_eq!(actual, expected);
    assert!(
        doc.paragraph_save_plan()
            .spliced_parts
            .iter()
            .all(|part| part.table_patches.is_empty())
    );
    actual
}

#[test]
fn complex_source_tables_never_produce_table_patches() {
    let simple = source(&paragraph(1, "abc"));
    for xml in [
        simple.replacen("<w:tcPr>", "<w:tcPr><w:vMerge w:val=\"restart\"/>", 1),
        simple.replacen("<w:tcPr>", "<w:tcPr><w:gridSpan w:val=\"2\"/>", 1),
        simple.replacen(
            "<w:trPr/>",
            "<w:trPr><w:gridBefore w:val=\"1\"/></w:trPr>",
            1,
        ),
        simple.replacen(
            "<w:trPr/>",
            "<w:trPr><w:gridAfter w:val=\"1\"/></w:trPr>",
            1,
        ),
        simple.replacen(
            &paragraph(1, "abc"),
            &format!(
                "<w:tbl><w:tr><w:tc>{}</w:tc></w:tr></w:tbl>{}",
                paragraph(10, "nested"),
                paragraph(1, "abc"),
            ),
            1,
        ),
    ] {
        let (doc, _) = open(&xml);
        doc.set_cell_shading(
            &ctx(),
            &TableRange::cell(CellLoc::new("body", 0, 0, 1)),
            Some("112233"),
        )
        .unwrap();
        assert!(
            doc.paragraph_save_plan()
                .spliced_parts
                .iter()
                .all(|part| part.table_patches.is_empty())
        );
    }
}

fn model_cell(model: &mut Value, column: usize) -> &mut Value {
    &mut model["content"][0]["rows"][0]["cells"][column]
}

fn model_text(cell: &mut Value, text: &str) {
    cell["content"][0]["content"][0]["content"][0]["text"] = json!(text);
}

fn replace(doc: &EditingDoc, cell: u32, text: &str) {
    doc.replace_range(
        &ctx(),
        StoryRange::new(format!("body:t0:r0c{cell}"), 0, 3),
        text,
    )
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
fn merge_cells_uses_the_base_story_fallback() {
    let first = paragraph(1, "abc");
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
    let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert!(warnings.is_empty(), "{warnings:?}");
    assert!(output.contains(r#"<w:gridSpan w:val="2"/>"#));
    assert!(output.contains("abc"));
    assert!(output.contains("def"));
}

#[test]
fn cell_format_preserves_row_ids_shading_and_borders() {
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
fn field_replace_with_a_cell_neighbour_edit_uses_the_base_story_fallback() {
    let field = concat!(
        r#"<w:p w14:paraId="00000001"><w:r><w:fldChar w:fldCharType="begin"/></w:r>"#,
        r#"<w:r><w:instrText> DATE </w:instrText></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="separate"/></w:r>"#,
        r#"<w:r><w:t>abc</w:t></w:r>"#,
        r#"<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#,
    );
    let xml = source(field);
    let (doc, mut model) = open(&xml);
    doc.replace_range(&ctx(), StoryRange::new("body:t0:r0c0", 0, 1), "xyz")
        .unwrap();
    replace(&doc, 1, "uvw");
    model_cell(&mut model, 0)["content"][0] = json!({
        "type": "paragraph", "paraId": "00000001", "sourceOrdinal": 0,
        "content": [{"type": "run", "content": [{"type": "text", "text": "xyz"}]}]
    });
    model_text(model_cell(&mut model, 1), "uvw");
    assert!(doc.paragraph_save_plan().spliced_parts.is_empty());
    let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert!(warnings.is_empty(), "{warnings:?}");
    assert!(output.contains(r#"<w:p w14:paraId="00000001"><w:r><w:t>xyz</w:t></w:r></w:p>"#));
    assert!(output.contains(r#"<w:p w14:paraId="00000002"><w:r><w:t>uvw</w:t></w:r></w:p>"#));
    assert!(!output.contains("fldChar"), "{output}");
    assert!(output.contains(r#"<w:tblCaption w:val="caption"/>"#));
}

#[test]
fn row_insertion_keeps_cell_edits_through_the_base_row_splice() {
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
    model["content"][0]["rows"]
        .as_array_mut()
        .unwrap()
        .push(row);
    let (output, warnings) = save_with_base_row_splice(&xml, &doc, model);
    let start = xml.find("<w:tr ").unwrap();
    let end = xml[start..].find("</w:tr>").unwrap() + start + "</w:tr>".len();
    assert!(output.starts_with(&xml[..start]));
    assert!(output.ends_with(&xml[end..]));
    let rows = &output[start..output.len() - xml[end..].len()];
    assert!(rows.starts_with(&xml[start..end].replace(&first, &first.replace("abc", "xyz"))));
    assert_eq!(rows.matches(r#"w:rsidTr="87654321""#).count(), 1);
    assert_eq!(rows.matches("xyz").count(), 1);
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
fn vertical_merge_uses_the_base_story_fallback() {
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
    let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert!(warnings.is_empty(), "{warnings:?}");
    assert!(output.contains(r#"<w:vMerge w:val="restart"/>"#));
    assert!(output.contains("<w:vMerge/>"));
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
fn merging_vertical_columns_uses_the_base_story_fallback() {
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
                &format!(
                    r#"<w:vMerge/><w:shd w:fill="{}"/>"#,
                    ["112233", "445566", "778899"][column],
                ),
                &format!(
                    r#"<w:p w14:paraId="{:08X}"><w:pPr><w:keepNext/></w:pPr></w:p>"#,
                    column + 4,
                ),
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
    let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert!(warnings.is_empty(), "{warnings:?}");
    assert!(output.contains(r#"<w:gridSpan w:val="2"/>"#));
}

fn inserted_model_row(
    doc: &EditingDoc,
    mut row: Value,
    receipt: &docx_edit::TableReceipt,
) -> Value {
    let identities = doc.paragraph_identities();
    for (cell, key) in row["cells"]
        .as_array_mut()
        .unwrap()
        .iter_mut()
        .zip(&receipt.new_para_ids)
    {
        let id = identities
            .paragraphs
            .iter()
            .find_map(|identity| match &identity.paragraph {
                docx_edit::ParagraphRef::Session { para_id, .. } if para_id == key => {
                    identity.ooxml_para_id.as_ref()
                }
                _ => None,
            })
            .unwrap();
        cell["content"] = json!([{
            "type": "paragraph", "paraId": id, "content": []
        }]);
    }
    row
}

fn formatted_row(id: u32, text: &str, height: u32, fill: &str) -> String {
    let cell = table_cell(
        &format!(
            r#"<w:tcBorders><w:top w:val="single" w:sz="12" w:color="{fill}"/></w:tcBorders><w:shd w:fill="{fill}"/>"#
        ),
        &paragraph(id, text),
    );
    format!(
        r#"<w:tr w:rsidR="{id:08X}"><w:trPr><w:trHeight w:val="{height}" w:hRule="exact"/></w:trPr>{cell}</w:tr>"#
    )
}

#[test]
fn column_insertion_uses_the_base_story_fallback() {
    let row = format!("<w:tr>{}</w:tr>", table_cell("", &paragraph(1, "A")));
    let xml = table_document(&row, 1);
    let (doc, mut model) = open(&xml);
    let receipt = doc
        .insert_column(&ctx(), &CellLoc::new("body", 0, 0, 0), true)
        .unwrap();
    let inserted = inserted_model_row(&doc, model["content"][0]["rows"][0].clone(), &receipt);
    model["content"][0]["rows"][0]["cells"]
        .as_array_mut()
        .unwrap()
        .push(inserted["cells"][0].clone());
    model["content"][0]["columnWidths"] = json!([2000, 2000]);
    let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert_eq!(output.matches("<w:t>A</w:t>").count(), 1);
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn table_formatting_with_a_crossing_body_comment_uses_the_base_story_fallback() {
    let row = format!("<w:tr>{}</w:tr>", table_cell("", &paragraph(1, "A")));
    let xml = table_document(&row, 1)
        .replace("</w:body>", &format!("{}</w:body>", paragraph(92, "last")));
    let (doc, mut model) = open(&xml);
    doc.set_cell_shading(
        &ctx(),
        &TableRange::cell(CellLoc::new("body", 0, 0, 0)),
        Some("112233"),
    )
    .unwrap();
    model_cell(&mut model, 0)["formatting"]["shading"] = json!({"fill": {"rgb": "112233"}});
    model["content"][2]["content"]
        .as_array_mut()
        .unwrap()
        .insert(0, json!({"type": "commentRangeStart", "id": 7}));
    model["content"][3]["content"]
        .as_array_mut()
        .unwrap()
        .push(json!({"type": "commentRangeEnd", "id": 7}));
    let expected = save_model(&xml, &doc, model.clone(), None).unwrap();
    let mut entry = save_entry(&doc).unwrap();
    entry["changed"] = json!([2, 3]);
    let actual = save_model(&xml, &doc, model, Some(entry)).unwrap();
    assert_eq!(actual, expected);
    assert!(actual.0.contains(r#"<w:commentRangeStart w:id="7"/>"#));
    assert!(actual.0.contains(r#"<w:commentRangeEnd w:id="7"/>"#));
    assert!(actual.0.contains(r#"w:fill="112233""#));
    assert_eq!(actual.0.matches("<w:t>A</w:t>").count(), 1);
}

#[test]
fn row_insertion_with_table_scoped_attributes_uses_the_base_row_splice() {
    let row = format!(
        r#"<w:tr x:keep="1">{}</w:tr>"#,
        table_cell("", &paragraph(1, "A")),
    );
    let xml = table_document(&row, 1).replacen("<w:tbl>", r#"<w:tbl xmlns:x="urn:custom">"#, 1);
    let (doc, mut model) = open(&xml);
    let receipt = doc
        .insert_row(&ctx(), &CellLoc::new("body", 0, 0, 0), true)
        .unwrap();
    let inserted = inserted_model_row(&doc, model["content"][0]["rows"][0].clone(), &receipt);
    model["content"][0]["rows"]
        .as_array_mut()
        .unwrap()
        .push(inserted);
    let (output, warnings) = save_with_base_row_splice(&xml, &doc, model);
    assert!(output.contains(&row), "{output}");
    assert_eq!(output.matches(r#"x:keep="1""#).count(), 1);
    assert_eq!(output.matches("<w:t>A</w:t>").count(), 1);
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn empty_inserted_cell_with_table_width_uses_the_base_story_fallback() {
    let row = format!("<w:tr>{}</w:tr>", table_cell("", &paragraph(1, "A")));
    let xml = table_document(&row, 1);
    let (doc, mut model) = open(&xml);
    let receipt = doc
        .insert_row(&ctx(), &CellLoc::new("body", 0, 0, 0), true)
        .unwrap();
    doc.apply_raw_ops(
        &receipt.created_story_ids[0],
        vec![RawOp::Delete { index: 0, len: 1 }],
        &ctx(),
    )
    .unwrap();
    doc.set_table_width(&ctx(), &TableLocator::new("body", 0), 5000.0)
        .unwrap();
    let mut inserted = model["content"][0]["rows"][0].clone();
    inserted["cells"][0]["content"] = json!([]);
    model["content"][0]["rows"]
        .as_array_mut()
        .unwrap()
        .push(inserted);
    model["content"][0]["formatting"] = json!({"width": {"value": 5000, "type": "dxa"}});
    let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert_eq!(output.matches("<w:t>A</w:t>").count(), 1);
    assert!(output.contains(r#"<w:tblW w:w="5000" w:type="dxa"/>"#));
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn table_patch_rejection_uses_the_base_story_fallback_after_a_paragraph_patch() {
    let row = format!("<w:tr>{}</w:tr>", table_cell("", &paragraph(1, "abc")));
    let xml = table_document(&row, 1)
        .replacen("<w:tbl>", r#"<w:tbl xmlns:x="urn:custom">"#, 1)
        .replacen("<w:tblPr/>", r#"<w:tblPr x:keep="1"/>"#, 1);
    let (doc, mut model) = open(&xml);
    replace(&doc, 0, "xyz");
    model_text(model_cell(&mut model, 0), "xyz");
    doc.set_table_width(&ctx(), &TableLocator::new("body", 0), 5000.0)
        .unwrap();
    model["content"][0]["formatting"] = json!({"width": {"value": 5000, "type": "dxa"}});
    let entry = save_entry(&doc).unwrap();
    assert!(!entry["tablePatches"].as_array().unwrap().is_empty());
    assert!(!entry["paragraphPatches"].as_array().unwrap().is_empty());
    let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert_eq!(output.matches("<w:t>xyz</w:t>").count(), 1);
    assert!(!output.contains("abc"), "{output}");
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn insert_before_uses_the_base_row_splice_and_preserves_suggested_insertion() {
    let first = formatted_row(1, "abc", 300, "112233");
    let second = formatted_row(2, "def", 600, "445566");
    let xml = table_document(&format!("{first}{second}"), 1);
    for suggesting in [false, true] {
        let (doc, mut model) = open(&xml);
        let context = if suggesting {
            ctx().suggesting()
        } else {
            ctx()
        };
        let receipt = doc
            .insert_row(&context, &CellLoc::new("body", 0, 1, 0), false)
            .unwrap();
        let mut inserted =
            inserted_model_row(&doc, model["content"][0]["rows"][1].clone(), &receipt);
        let revision_id = suggesting.then(|| {
            let id = &receipt.revision_ids[0];
            let hash = id.bytes().fold(0xcbf29ce484222325_u64, |hash, byte| {
                (hash ^ u64::from(byte)).wrapping_mul(0x100000001b3)
            });
            hash & ((1_u64 << 53) - 1)
        });
        if let Some(id) = revision_id {
            inserted["structuralChange"] = json!({
                "type": "tableRowInsertion",
                "info": {"id": id, "author": "", "date": context.now_iso}
            });
        }
        model["content"][0]["rows"]
            .as_array_mut()
            .unwrap()
            .insert(1, inserted);
        let (output, warnings) = save_with_base_row_splice(&xml, &doc, model);
        let prefix = &xml[..xml.find(&first).unwrap() + first.len()];
        let suffix = &xml[xml.find(&second).unwrap()..];
        assert!(output.starts_with(prefix));
        assert!(output.ends_with(suffix));
        let inserted = &output[prefix.len()..output.len() - suffix.len()];
        assert!(inserted.contains(r#"<w:trHeight w:val="600" w:hRule="exact"/>"#));
        assert!(inserted.contains(r#"w:fill="445566""#));
        assert!(inserted.contains(r#"<w:top w:val="single" w:sz="12" w:color="445566"/>"#));
        if let Some(id) = revision_id {
            assert!(
                inserted.contains(&format!(r#"<w:ins w:id="{id}""#)),
                "{inserted}"
            );
        }
        assert!(warnings.is_empty(), "{warnings:?}");
    }
}

#[test]
fn appending_beside_continuations_uses_the_base_story_fallback() {
    let restart = table_cell(r#"<w:vMerge w:val="restart"/>"#, &paragraph(1, "abc"));
    let continuation = table_cell(
        r#"<w:vMerge/><w:shd w:fill="ABCDEF"/>"#,
        r#"<w:p w14:paraId="00000002"><w:pPr><w:keepNext/></w:pPr></w:p>"#,
    );
    let lower = format!(
        r#"<w:tr w:rsidR="11223344"><w:trPr><w:trHeight w:val="720" w:hRule="exact"/></w:trPr>{continuation}</w:tr>"#
    );
    let ordinary = formatted_row(3, "def", 300, "445566");
    let xml = table_document(&format!("<w:tr>{restart}</w:tr>{lower}{ordinary}"), 1);
    let (doc, mut model) = open(&xml);
    clear_continuations(&mut model);
    let receipt = doc
        .insert_row(&ctx(), &CellLoc::new("body", 0, 2, 0), true)
        .unwrap();
    let inserted = inserted_model_row(&doc, model["content"][0]["rows"][2].clone(), &receipt);
    model["content"][0]["rows"]
        .as_array_mut()
        .unwrap()
        .push(inserted);
    let (_, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn balanced_row_insertion_and_deletion_keeps_retained_row_identities() {
    let first = formatted_row(1, "abc", 300, "112233");
    let second = formatted_row(2, "def", 600, "445566");
    let third = formatted_row(3, "ghi", 900, "778899");
    let xml = table_document(&format!("{first}{second}{third}"), 1);
    let (doc, mut model) = open(&xml);
    let receipt = doc
        .insert_row(&ctx(), &CellLoc::new("body", 0, 0, 0), false)
        .unwrap();
    let inserted = inserted_model_row(&doc, model["content"][0]["rows"][0].clone(), &receipt);
    model["content"][0]["rows"]
        .as_array_mut()
        .unwrap()
        .insert(0, inserted);
    doc.delete_row(&ctx(), &TableRange::cell(CellLoc::new("body", 0, 2, 0)))
        .unwrap();
    model["content"][0]["rows"]
        .as_array_mut()
        .unwrap()
        .remove(2);
    let (output, warnings) = save_with_base_row_splice(&xml, &doc, model);
    assert!(output.contains(&format!("{first}{third}")), "{output}");
    assert!(!output.contains(&second));
    assert!(output.ends_with(&xml[xml.find("</w:tbl>").unwrap()..]));
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn splitting_into_new_rows_does_not_duplicate_source_row_identity() {
    let first = formatted_row(1, "abc", 300, "112233");
    let second = formatted_row(2, "def", 600, "445566");
    let xml = table_document(&format!("{first}{second}"), 1);
    let (doc, mut model) = open(&xml);
    let receipt = doc
        .split_cell_grid(&ctx(), &CellLoc::new("body", 0, 0, 0), Some(2), Some(1))
        .unwrap();
    let inserted = inserted_model_row(&doc, model["content"][0]["rows"][0].clone(), &receipt);
    model["content"][0]["rows"]
        .as_array_mut()
        .unwrap()
        .insert(1, inserted);
    let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert_eq!(output.matches("<w:t>def</w:t>").count(), 1);
    assert_eq!(output.matches("</w:tr>").count(), 4);
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn nested_row_insertion_is_emitted_once_inside_its_own_table() {
    let inner_row = formatted_row(1, "abc", 300, "112233");
    let inner = format!(
        r#"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>{inner_row}</w:tbl>"#
    );
    let outer_cell = table_cell("", &format!("{inner}{}", paragraph(2, "outer")));
    let xml = table_document(&format!("<w:tr>{outer_cell}</w:tr>"), 1);
    for after in [false, true] {
        let (doc, mut model) = open(&xml);
        let receipt = doc
            .insert_row(&ctx(), &CellLoc::new("body:t0:r0c0", 0, 0, 0), after)
            .unwrap();
        let rows = model["content"][0]["rows"][0]["cells"][0]["content"][0]["rows"]
            .as_array_mut()
            .unwrap();
        let inserted = inserted_model_row(&doc, rows[0].clone(), &receipt);
        let id = inserted["cells"][0]["content"][0]["paraId"]
            .as_str()
            .unwrap()
            .to_owned();
        rows.insert(usize::from(after), inserted);
        let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
        assert_eq!(output.matches(&format!(r#"w14:paraId="{id}""#)).count(), 1);
        assert_eq!(output.matches("</w:tr>").count(), 4);
        assert!(warnings.is_empty(), "{warnings:?}");
    }
}

#[test]
fn grid_offset_merge_keeps_the_continuation_in_its_source_columns() {
    let top = (1..=4)
        .map(|id| table_cell("", &paragraph(id, "abc")))
        .collect::<String>();
    let lower_properties = r#"<w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>"#;
    let lower_cells = [
        table_cell(r#"<w:shd w:fill="112233"/>"#, &paragraph(5, "def")),
        table_cell("", &paragraph(6, "ghi")),
    ];
    let lower = format!("<w:tr>{lower_properties}{}</w:tr>", lower_cells.concat());
    let xml = table_document(&format!("<w:tr>{top}</w:tr>{lower}"), 4);
    let (doc, mut model) = open(&xml);
    doc.merge_cells(
        &ctx(),
        &TableRange::new(CellLoc::new("body", 0, 0, 1), CellLoc::new("body", 0, 1, 2)),
    )
    .unwrap();
    let rows = model["content"][0]["rows"].as_array_mut().unwrap();
    let bottom = rows[1]["cells"].as_array().unwrap().clone();
    let cells = rows[0]["cells"].as_array_mut().unwrap();
    let removed = cells.remove(2);
    cells[1]["formatting"]["gridSpan"] = json!(2);
    cells[1]["formatting"]["vMerge"] = json!("restart");
    cells[1]["content"]
        .as_array_mut()
        .unwrap()
        .extend(removed["content"].as_array().unwrap().iter().cloned());
    for cell in &bottom {
        cells[1]["content"]
            .as_array_mut()
            .unwrap()
            .extend(cell["content"].as_array().unwrap().iter().cloned());
    }
    let mut continuation = bottom[0].clone();
    continuation["formatting"]["gridSpan"] = json!(2);
    continuation["formatting"]["vMerge"] = json!("continue");
    continuation["content"] = json!([]);
    rows[1]["cells"] = json!([continuation]);
    let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert!(warnings.is_empty(), "{warnings:?}");
    assert!(output.contains(r#"<w:gridSpan w:val="2"/>"#));
}

#[test]
fn grid_offset_vertical_merges_seed_on_the_native_grid() {
    let restart = r#"<w:vMerge w:val="restart"/>"#;
    let offsets = r#"<w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>"#;
    let top = format!(
        "<w:tr>{offsets}{}{}</w:tr>",
        table_cell(restart, &paragraph(1, "A")),
        table_cell(restart, &paragraph(2, "B")),
    );
    let bottom = format!(
        "<w:tr>{}{}{}</w:tr>",
        table_cell("", &paragraph(3, "X")),
        table_cell("<w:vMerge/>", "<w:p/>"),
        table_cell(r#"<w:gridSpan w:val="2"/>"#, &paragraph(4, "D")),
    );
    let xml = table_document(&format!("{top}{bottom}"), 4);
    let (doc, _) = open(&xml);
    for (row, column) in [(1, 2), (1, 1), (0, 2), (1, 0)] {
        doc.set_cell_shading(
            &ctx(),
            &TableRange::cell(CellLoc::new("body", 0, row, column)),
            Some("112233"),
        )
        .unwrap();
    }
}

#[test]
fn grid_offset_properties_use_the_base_story_fallback() {
    let properties = r#"<w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>"#;
    let cells = format!(
        "{}{}",
        table_cell("", &paragraph(1, "abc")),
        table_cell("", &paragraph(2, "def")),
    );
    let xml = table_document(&format!("<w:tr>{properties}{cells}</w:tr>"), 4);
    let (doc, mut model) = open(&xml);
    doc.set_cell_shading(
        &ctx(),
        &TableRange::cell(CellLoc::new("body", 0, 0, 1)),
        Some("112233"),
    )
    .unwrap();
    model_cell(&mut model, 0)["formatting"]["shading"] = json!({"fill": {"rgb": "112233"}});
    let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert!(output.contains(r#"w:fill="112233""#));
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn merges_with_continuation_paragraphs_use_the_base_story_fallback() {
    for replace_row in [false, true] {
        let owners = (1..=3)
            .map(|id| table_cell(r#"<w:vMerge w:val="restart"/>"#, &paragraph(id, "abc")))
            .collect::<String>();
        let hidden = format!(
            "{}{}",
            paragraph(6, "hidden one"),
            paragraph(7, "hidden two")
        );
        let untouched = table_cell(r#"<w:vMerge/><w:shd w:fill="778899"/>"#, &hidden);
        let lower = format!(
            "{}{}{}",
            table_cell("<w:vMerge/>", &paragraph(4, "")),
            table_cell("<w:vMerge/>", &paragraph(5, "")),
            untouched,
        );
        let top_side = if replace_row {
            table_cell("", &paragraph(10, "side"))
        } else {
            String::new()
        };
        let lower_side = if replace_row {
            table_cell("", &paragraph(11, "side"))
        } else {
            String::new()
        };
        let xml = table_document(
            &format!("<w:tr>{owners}{top_side}</w:tr><w:tr>{lower}{lower_side}</w:tr>"),
            if replace_row { 4 } else { 3 },
        );
        let (doc, mut model) = open(&xml);
        clear_continuations(&mut model);
        if replace_row {
            doc.set_cell_shading(
                &ctx(),
                &TableRange::cell(CellLoc::new("body", 0, 1, 3)),
                Some("112233"),
            )
            .unwrap();
            model["content"][0]["rows"][1]["cells"][3]["formatting"]["shading"] =
                json!({"fill": {"rgb": "112233"}});
        }
        doc.merge_cells(
            &ctx(),
            &TableRange::new(CellLoc::new("body", 0, 0, 0), CellLoc::new("body", 0, 1, 1)),
        )
        .unwrap();
        for row in model["content"][0]["rows"].as_array_mut().unwrap() {
            let cells = row["cells"].as_array_mut().unwrap();
            let removed = cells.remove(1);
            cells[0]["formatting"]["gridSpan"] = json!(2);
            cells[0]["content"]
                .as_array_mut()
                .unwrap()
                .extend(removed["content"].as_array().unwrap().iter().cloned());
        }
        let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
        assert!(warnings.is_empty(), "{warnings:?}");
        if replace_row {
            assert!(output.contains(r#"w:fill="112233""#));
        }
    }
}

#[test]
fn continuation_markers_follow_the_base_story_fallback() {
    let owner = table_cell(r#"<w:vMerge w:val="restart"/>"#, &paragraph(1, "abc"));
    let hidden = concat!(
        r#"<w:p w14:paraId="00000002"><w:bookmarkStart w:id="7" w:name="Kept"/>"#,
        r#"<w:r><w:t>hidden one</w:t></w:r></w:p><w:p w14:paraId="00000003">"#,
        r#"<w:r><w:t>hidden two</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p>"#,
    );
    let continuation = table_cell("<w:vMerge/>", hidden);
    let ordinary = formatted_row(4, "def", 300, "445566");
    for outside_markers in [false, true] {
        let outside = if outside_markers {
            r#"<w:bookmarkStart w:id="8" w:name="Outside"/><w:bookmarkEnd w:id="8"/>"#
        } else {
            ""
        };
        let xml = table_document(
            &format!("<w:tr>{owner}</w:tr><w:tr>{continuation}</w:tr>{ordinary}{outside}"),
            1,
        );
        let (doc, mut model) = open(&xml);
        clear_continuations(&mut model);
        let receipt = doc
            .insert_row(&ctx(), &CellLoc::new("body", 0, 2, 0), true)
            .unwrap();
        let inserted = inserted_model_row(&doc, model["content"][0]["rows"][2].clone(), &receipt);
        model["content"][0]["rows"]
            .as_array_mut()
            .unwrap()
            .push(inserted);
        let (_, warnings) = save_with_base_fallback(&xml, &doc, model);
        assert!(warnings.is_empty(), "{warnings:?}");
    }
}

fn merge_model_row(model: &mut Value) {
    let cells = model["content"][0]["rows"][0]["cells"]
        .as_array_mut()
        .unwrap();
    let removed = cells.remove(1);
    cells[0]["formatting"]["gridSpan"] = json!(2);
    cells[0]["content"]
        .as_array_mut()
        .unwrap()
        .extend(removed["content"].as_array().unwrap().iter().cloned());
}

#[test]
fn nested_continuations_seed_and_save_the_base_paragraphs() {
    let inner = format!(
        "<w:tbl><w:tr>{}</w:tr><w:tr>{}</w:tr></w:tbl>",
        table_cell(r#"<w:vMerge w:val="restart"/>"#, &paragraph(3, "owner")),
        table_cell(
            "<w:vMerge/>",
            &format!("{}{}", paragraph(4, "one"), paragraph(5, "two")),
        ),
    );
    let row = format!(
        "<w:tr>{}{}{}</w:tr>",
        table_cell("", &paragraph(1, "abc")),
        table_cell("", &paragraph(2, "def")),
        table_cell("", &format!("{inner}{}", paragraph(6, "outer"))),
    );
    let xml = table_document(&row, 3);
    let (doc, mut model) = open(&xml);
    assert_eq!(
        doc.paragraphs("body:t0:r0c2:t0:r1c0")
            .unwrap()
            .iter()
            .map(|paragraph| paragraph.text.as_str())
            .collect::<Vec<_>>(),
        ["one", "two"],
    );
    doc.merge_cells(
        &ctx(),
        &TableRange::new(CellLoc::new("body", 0, 0, 0), CellLoc::new("body", 0, 0, 1)),
    )
    .unwrap();
    merge_model_row(&mut model);
    let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert!(output.contains("<w:t>one</w:t>"), "{output}");
    assert!(output.contains("<w:t>two</w:t>"), "{output}");
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn merge_then_delete_the_first_paragraph_keeps_live_cell_properties() {
    let xml = source(&paragraph(1, "abc"));
    let (doc, mut model) = open(&xml);
    doc.merge_cells(
        &ctx(),
        &TableRange::new(CellLoc::new("body", 0, 0, 0), CellLoc::new("body", 0, 0, 1)),
    )
    .unwrap();
    merge_model_row(&mut model);
    doc.apply_raw_ops(
        "body:t0:r0c0",
        vec![RawOp::Delete { index: 0, len: 4 }],
        &ctx(),
    )
    .unwrap();
    model_cell(&mut model, 0)["content"]
        .as_array_mut()
        .unwrap()
        .remove(0);
    let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert!(warnings.is_empty(), "{warnings:?}");
    let first_cell = &output[..output.find("</w:tc>").unwrap()];
    assert!(first_cell.contains(r#"w:fill="AABBCC""#), "{output}");
    assert!(!first_cell.contains(r#"w:fill="DDEEFF""#), "{output}");
}

#[test]
fn inserted_rows_with_offsets_use_the_base_row_splice() {
    let offsets = r#"<w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/></w:trPr>"#;
    let row = format!(
        "<w:tr>{offsets}{}{}</w:tr>",
        table_cell("", &paragraph(1, "abc")),
        table_cell("", &paragraph(2, "def")),
    );
    let xml = table_document(&row, 4);
    let (doc, mut model) = open(&xml);
    let receipt = doc
        .insert_row(&ctx(), &CellLoc::new("body", 0, 0, 1), true)
        .unwrap();
    let mut template = model["content"][0]["rows"][0].clone();
    let cell = template["cells"][0].clone();
    template["cells"] = json!([cell.clone(), cell.clone(), cell.clone(), cell]);
    let inserted = inserted_model_row(&doc, template, &receipt);
    model["content"][0]["rows"]
        .as_array_mut()
        .unwrap()
        .push(inserted);
    let (output, warnings) = save(&xml, &doc, model);
    assert!(warnings.is_empty(), "{warnings:?}");
    assert!(output.contains("abc"), "{output}");
    assert!(
        doc.paragraph_save_plan()
            .spliced_parts
            .iter()
            .all(|part| part.table_patches.is_empty())
    );
}

#[test]
fn retained_cell_reordering_uses_the_base_story_fallback() {
    let first = format!(
        "<w:tr>{}{}</w:tr>",
        table_cell("", &paragraph(1, "abc")),
        table_cell("", &paragraph(2, "def")),
    );
    let second = formatted_row(3, "ghi", 300, "112233").replace(
        "</w:tr>",
        &format!("{}</w:tr>", table_cell("", &paragraph(4, "jkl"))),
    );
    let xml = table_document(&format!("{first}{second}"), 2);
    let (doc, mut model) = open(&xml);
    let payload = doc.table_payload(&TableLocator::new("body", 0)).unwrap();
    let mut rows = serde_json::to_value(&payload["rows"]).unwrap();
    rows[0]["cells"].as_array_mut().unwrap().swap(0, 1);
    doc.set_embed_attrs(
        &ctx(),
        Position::new("body", 0),
        vec![(
            "rows".to_owned(),
            Any::from_json(&rows.to_string()).unwrap(),
        )],
    )
    .unwrap();
    model["content"][0]["rows"][0]["cells"]
        .as_array_mut()
        .unwrap()
        .swap(0, 1);
    let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert!(output.find("def").unwrap() < output.find("abc").unwrap());
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn row_replace_keeps_a_verified_patch_when_projection_is_stale() {
    let row = formatted_row(1, "abc", 300, "AABBCC");
    let second = formatted_row(2, "def", 300, "112233");
    let xml = table_document(&format!("{row}{second}"), 1);
    let (doc, mut model) = open(&xml);
    replace(&doc, 0, "xyz");
    doc.set_cell_shading(
        &ctx(),
        &TableRange::cell(CellLoc::new("body", 0, 0, 0)),
        Some("112233"),
    )
    .unwrap();
    model_cell(&mut model, 0)["formatting"]["shading"]["fill"] = json!({"rgb": "112233"});
    let (output, warnings) = save(&xml, &doc, model);
    let start = xml.find(&row).unwrap();
    assert!(output.starts_with(&xml[..start]));
    assert!(output.ends_with(&xml[start + row.len()..]));
    assert!(output.contains("xyz"), "{output}");
    assert!(!output.contains("abc"), "{output}");
    assert_eq!(
        warnings,
        [r#"DOCX_MINIMAL_PATCH replace counts {"table:backgroundColor":1}"#]
    );
}

#[test]
fn merge_save_retains_comment_anchors_and_unrelated_table_metadata() {
    let marked = concat!(
        r#"<w:p w14:paraId="00000002"><w:commentRangeStart w:id="7"/>"#,
        r#"<w:r><w:t>def</w:t></w:r><w:commentRangeEnd w:id="7"/>"#,
        r#"<w:r><w:commentReference w:id="7"/></w:r></w:p>"#,
    );
    let xml = source(&paragraph(1, "abc")).replace(&paragraph(2, "def"), marked);
    let (doc, mut model) = open(&xml);
    let comment = doc
        .add_comment(
            &[StoryRange::new("body:t0:r0c1", 0, 3)],
            "",
            "2026-10-21",
            Any::from("note"),
        )
        .unwrap();
    doc.merge_cells(
        &ctx(),
        &TableRange::new(CellLoc::new("body", 0, 0, 0), CellLoc::new("body", 0, 0, 1)),
    )
    .unwrap();
    merge_model_row(&mut model);
    let resolved = doc.resolve_comment(&comment).unwrap();
    assert_eq!(resolved.len(), 1);
    assert_eq!(resolved[0].story, "body:t0:r0c0");
    assert_eq!((resolved[0].start, resolved[0].end), (4, 7));
    let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert!(warnings.is_empty(), "{warnings:?}");
    assert!(output.contains(r#"<w:commentRangeStart w:id="7"/>"#));
    assert!(output.contains(r#"<w:commentRangeEnd w:id="7"/>"#));
    assert!(output.contains(r#"<w:commentReference w:id="7"/>"#));
    assert!(output.contains(r#"<w:tblCaption w:val="caption"/>"#));
    assert!(output.contains(r#"<w:tblDescription w:val="description"/>"#));
}

#[test]
fn rectangular_text_and_row_insertion_use_the_base_row_splice() {
    let first = formatted_row(1, "abc", 300, "112233");
    let second = formatted_row(2, "def", 600, "445566");
    let third = formatted_row(3, "ghi", 900, "778899");
    let xml = table_document(&format!("{first}{second}{third}"), 1);
    let (doc, mut model) = open(&xml);
    replace(&doc, 0, "xyz");
    model_text(model_cell(&mut model, 0), "xyz");
    let receipt = doc
        .insert_row(&ctx(), &CellLoc::new("body", 0, 0, 0), true)
        .unwrap();
    let inserted = inserted_model_row(&doc, model["content"][0]["rows"][0].clone(), &receipt);
    model["content"][0]["rows"]
        .as_array_mut()
        .unwrap()
        .insert(1, inserted);
    let (output, warnings) = save_with_base_row_splice(&xml, &doc, model);
    assert!(output.contains(&first.replace("abc", "xyz")), "{output}");
    assert!(
        output.ends_with(&xml[xml.find(&second).unwrap()..]),
        "{output}"
    );
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn complex_table_text_edit_keeps_the_legacy_paragraph_splice() {
    let original = paragraph(1, "abc");
    let xml =
        source(&original).replace("<w:trPr/>", r#"<w:trPr><w:gridAfter w:val="1"/></w:trPr>"#);
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
fn a_simple_row_patch_keeps_an_untouched_table_with_omitted_nested_content() {
    let hidden = format!(
        "<w:tbl><w:tr>{}</w:tr></w:tbl>",
        table_cell("", &paragraph(6, "nested")),
    );
    let complex = format!(
        "<w:tbl><w:tr>{}{}</w:tr><w:tr>{}{}</w:tr></w:tbl>",
        table_cell(r#"<w:vMerge w:val="restart"/>"#, &paragraph(5, "owner")),
        table_cell("", &paragraph(7, "side")),
        table_cell("<w:vMerge/>", &hidden),
        table_cell("", &paragraph(8, "lower")),
    );
    let xml = source(&paragraph(1, "abc"));
    let at = xml.find("<w:tbl>").unwrap();
    let xml = format!("{}{complex}{}", &xml[..at], &xml[at..]);
    let (doc, mut model) = open(&xml);
    model["content"][0]["rows"][1]["cells"][0]["content"] = json!([]);
    doc.set_cell_shading(
        &ctx(),
        &TableRange::cell(CellLoc::new("body", 1, 0, 0)),
        Some("112233"),
    )
    .unwrap();
    model["content"][1]["rows"][0]["cells"][0]["formatting"]["shading"]["fill"] =
        json!({"rgb": "112233"});
    let (output, warnings) = save(&xml, &doc, model);
    assert!(output.contains(&complex), "{output}");
    assert!(output.contains(r#"w:fill="112233""#), "{output}");
    assert_eq!(
        warnings,
        [r#"DOCX_MINIMAL_PATCH replace counts {"table:backgroundColor":1}"#],
    );
}

#[test]
fn complex_table_changes_use_the_base_projected_paragraphs() {
    let xml = source(&paragraph(1, "abc"));
    let (doc, mut model) = open(&xml);
    replace(&doc, 0, "xyz");
    doc.merge_cells(
        &ctx(),
        &TableRange::new(CellLoc::new("body", 0, 0, 0), CellLoc::new("body", 0, 0, 1)),
    )
    .unwrap();
    merge_model_row(&mut model);
    model_text(model_cell(&mut model, 0), "xyz");
    let (output, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert!(warnings.is_empty(), "{warnings:?}");
    assert!(output.contains("xyz"), "{output}");
    assert!(!output.contains("abc"), "{output}");
}

#[test]
fn complex_table_changes_keep_the_base_unrelated_table_output() {
    let xml = source(&paragraph(1, "abc")).replace(
        "<w:tr><w:tc>",
        r#"<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr><w:tc>"#,
    );
    let (doc, mut model) = open(&xml);
    doc.merge_cells(
        &ctx(),
        &TableRange::new(CellLoc::new("body", 0, 0, 0), CellLoc::new("body", 0, 0, 1)),
    )
    .unwrap();
    merge_model_row(&mut model);
    let (_, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert!(warnings.is_empty(), "{warnings:?}");
}

fn set_source_table_formatting(doc: &EditingDoc, entries: &[(&str, Value)]) {
    let payload = doc.table_payload(&TableLocator::new("body", 0)).unwrap();
    let mut tbl_pr = serde_json::to_value(&payload["tblPr"]).unwrap();
    for (key, value) in entries {
        tbl_pr["_originalFormatting"][*key] = value.clone();
    }
    doc.set_embed_attrs(
        &ctx(),
        Position::new("body", 0),
        vec![(
            "tblPr".to_owned(),
            Any::from_json(&tbl_pr.to_string()).unwrap(),
        )],
    )
    .unwrap();
    let entry = save_entry(doc).unwrap();
    assert!(!entry["tablePatches"].as_array().unwrap().is_empty());
}

#[test]
fn table_caption_description_and_band_sizes_are_patched_into_tbl_pr() {
    let xml = source(&paragraph(1, "abc"));
    let (doc, mut model) = open(&xml);
    let entries = [
        ("caption", json!("Totals")),
        ("description", json!("Quarterly totals")),
        ("styleRowBandSize", json!(2)),
        ("styleColBandSize", json!(3)),
    ];
    set_source_table_formatting(&doc, &entries);
    for (key, value) in entries {
        model["content"][0]["formatting"][key] = value;
    }
    let (output, warnings) = save(&xml, &doc, model);
    let expected = xml
        .replacen(
            "<w:tblPr><w:tblBorders>",
            r#"<w:tblPr><w:tblStyleRowBandSize w:val="2"/><w:tblStyleColBandSize w:val="3"/><w:tblBorders>"#,
            1,
        )
        .replacen(
            "</w:tblBorders></w:tblPr>",
            r#"</w:tblBorders><w:tblCaption w:val="Totals"/><w:tblDescription w:val="Quarterly totals"/></w:tblPr>"#,
            1,
        );
    assert_eq!(output, expected);
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn unmapped_table_formatting_changes_use_the_base_story_fallback() {
    let xml = source(&paragraph(1, "abc"));
    let (doc, model) = open(&xml);
    set_source_table_formatting(&doc, &[("unmappedProperty", json!("x"))]);
    let (_, warnings) = save_with_base_fallback(&xml, &doc, model);
    assert!(warnings.is_empty(), "{warnings:?}");
}

#[test]
fn table_width_patches_tbl_pr_and_keeps_rows_byte_identical() {
    let original = paragraph(1, "abc");
    let xml = source(&original);
    let (doc, mut model) = open(&xml);
    doc.set_table_width(&ctx(), &TableLocator::new("body", 0), 5000.0)
        .unwrap();
    model["content"][0]["formatting"]["width"] = json!({"value": 5000, "type": "dxa"});
    let (output, warnings) = save(&xml, &doc, model);
    assert_eq!(
        output,
        xml.replacen(
            "<w:tblPr><w:tblBorders>",
            r#"<w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:tblBorders>"#,
            1,
        )
    );
    assert!(warnings.is_empty(), "{warnings:?}");
}
