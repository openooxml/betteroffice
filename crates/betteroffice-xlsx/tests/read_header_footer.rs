use betteroffice_xlsx::{
    CalculationOptions, CellRef, CellState, CellValue, Op, ReadRequest, Sheet, SheetId, Workbook,
    WorkbookModel,
};
use serde_json::{Value, json};

const ELEMENTS: [&str; 6] = [
    r#"<headerFooter><oddHeader>&amp;L&amp;"Arial,Bold"&amp;14Left &amp;A&amp;C&amp;KFF0000Center &amp;P of &amp;N&amp;RRight &amp;F</oddHeader><oddFooter>Bare footer &amp;P</oddFooter></headerFooter>"#,
    r#"<headerFooter differentOddEven="1"><oddHeader>Odd</oddHeader><evenHeader>Even</evenHeader><evenFooter>&amp;N</evenFooter></headerFooter>"#,
    r#"<headerFooter differentFirst="true"><oddHeader>Odd</oddHeader><firstHeader>First</firstHeader><firstFooter>&amp;D</firstFooter></headerFooter>"#,
    r#"<headerFooter differentOddEven="0" differentFirst="false"><oddHeader>Odd</oddHeader><evenHeader>Ignored</evenHeader><firstHeader>Ignored</firstHeader></headerFooter>"#,
    "",
    "<headerFooter><oddHeader></oddHeader></headerFooter>",
];

fn package() -> Vec<u8> {
    let mut model = WorkbookModel::default();
    for index in 0..ELEMENTS.len() {
        model.sheets.push(Sheet::new(format!("Sheet{}", index + 1)));
    }
    let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
    for (index, element) in ELEMENTS.iter().enumerate() {
        let path = format!("xl/worksheets/sheet{}.xml", index + 1);
        let sheet = parts.iter_mut().find(|(name, _)| name == &path).unwrap();
        sheet.1 = String::from_utf8(std::mem::take(&mut sheet.1))
            .unwrap()
            .replace("</worksheet>", &format!("{element}</worksheet>"))
            .into_bytes();
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn text(raw: &str, left: &str, center: &str, right: &str) -> Value {
    json!({ "raw": raw, "left": left, "center": center, "right": right })
}

fn centered(raw: &str, center: &str) -> Value {
    text(raw, "", center, "")
}

fn pair(header: Value, footer: Value) -> Value {
    json!({ "header": header, "footer": footer })
}

fn expected() -> Vec<Value> {
    let odd = pair(centered("Odd", "Odd"), Value::Null);
    let variants = [
        (pair(text(r#"&L&"Arial,Bold"&14Left &A&C&KFF0000Center &P of &N&RRight &F"#, "Left &[Tab]", "Center &[Page] of &[Pages]", "Right &[File]"), centered("Bare footer &P", "Bare footer &[Page]")), Value::Null, Value::Null),
        (odd.clone(), pair(centered("Even", "Even"), centered("&N", "&[Pages]")), Value::Null),
        (odd.clone(), Value::Null, pair(centered("First", "First"), centered("&D", "&[Date]"))),
        (odd, Value::Null, Value::Null),
    ];
    let mut expected = variants.into_iter().enumerate().map(|(index, (odd, even, first))| json!({
        "differentFirst": index == 2, "differentOddEven": index == 1,
        "scaleWithDoc": true, "alignWithMargins": true,
        "odd": odd, "even": even, "first": first,
    })).collect::<Vec<_>>();
    expected.extend([Value::Null, Value::Null]);
    expected
}

fn catalog(workbook: &Workbook) -> Vec<Value> {
    workbook.read_cells(&ReadRequest::default()).unwrap().unwrap().sheets
        .into_iter().map(|sheet| serde_json::to_value(sheet).unwrap()).collect()
}

fn assert_facts(workbook: &Workbook, expected: &[Value]) {
    let sheets = catalog(workbook);
    assert_eq!(sheets.len(), expected.len());
    for (sheet, expected) in sheets.iter().zip(expected) {
        assert_eq!(sheet.get("headerFooter"), Some(expected), "{}", sheet["name"]);
    }
}

#[test]
fn read_cells_reports_source_header_footer_and_session_sheet_absence() {
    let mut workbook = Workbook::open(&package()).unwrap();
    let mut expected = expected();
    assert_facts(&workbook, &expected);
    workbook.apply_ops(vec![Op::AddSheet { index: 6, name: "Fresh".to_owned() }], CalculationOptions::default()).unwrap();
    expected.push(Value::Null);
    assert_facts(&workbook, &expected);
}

#[test]
fn read_cells_omits_header_footer_for_model_origins() {
    for collaborative in [false, true] {
        let mut model = WorkbookModel::default();
        model.sheets.push(Sheet::new("Plain"));
        let workbook = if collaborative {
            Workbook::from_model_collaborative(model, 7).unwrap()
        } else {
            Workbook::from_model(model).unwrap()
        };
        let sheets = catalog(&workbook);
        assert_eq!(sheets[0]["visibility"], "unknown");
        assert!(!sheets[0].as_object().unwrap().contains_key("headerFooter"));
    }
}

fn elements(bytes: &[u8]) -> Vec<Vec<u8>> {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    (1..=ELEMENTS.len()).map(|index| {
        let path = format!("xl/worksheets/sheet{index}.xml");
        let bytes = &parts.iter().find(|(name, _)| name == &path).unwrap().1;
        let xml = std::str::from_utf8(bytes).unwrap();
        match xml.find("<headerFooter") {
            None => Vec::new(),
            Some(start) => {
                let end = start + xml[start..].find("</headerFooter>").unwrap() + "</headerFooter>".len();
                bytes[start..end].to_vec()
            }
        }
    }).collect()
}

#[test]
fn header_footer_survives_save_and_unrelated_cell_edit_byte_identically() {
    let source = package();
    for edit in [false, true] {
        let mut workbook = Workbook::open(&source).unwrap();
        let before = catalog(&workbook);
        if edit {
            workbook.apply_ops(vec![Op::SetCell {
                sheet: SheetId(0), at: CellRef::parse_a1("A1").unwrap(),
                cell: CellState { value: CellValue::Text { value: "Unrelated".into() }, ..Default::default() },
            }], CalculationOptions::default()).unwrap();
        }
        let saved = workbook.save().unwrap();
        assert_eq!(elements(&saved), elements(&source));
        let reopened = Workbook::open(&saved).unwrap();
        let after = catalog(&reopened);
        for (before, after) in before.iter().zip(&after) {
            assert_eq!(before.get("headerFooter"), after.get("headerFooter"));
        }
        assert_facts(&reopened, &expected());
    }
}
