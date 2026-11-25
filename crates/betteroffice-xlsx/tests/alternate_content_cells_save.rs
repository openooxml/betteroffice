use betteroffice_xlsx::{
    CalculationOptions, CellRange, CellRef, CellValue, Error, NumberFormatMutation, SheetId, Workbook,
};
use xlsx_model::CellProvider;

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const EXT: &str = "urn:test:ext";

const WRAPPED_A1: &str = concat!(
    "<mc:AlternateContent>",
    r#"<mc:Choice Requires="x"><c r="A1" t="inlineStr"><is><t>choice text</t></is></c></mc:Choice>"#,
    r#"<mc:Fallback><c r="A1" t="inlineStr"><is><t>fallback text</t></is></c></mc:Fallback>"#,
    "</mc:AlternateContent>",
);

fn cell_ref(address: &str) -> CellRef {
    CellRef::parse_a1(address).unwrap()
}

fn sheet2_data() -> String {
    format!(
        concat!(
            r#"<sheetData><row r="1">{wrapped}"#,
            r#"<c r="B1"><f>1+1</f><v>2</v></c>"#,
            r#"<c r="C1" t="str"><f>_xlfn.FORMULATEXT(Sheet1!H1)</f><v>=1+1</v></c>"#,
            "</row></sheetData>",
        ),
        wrapped = WRAPPED_A1
    )
}

fn package() -> Vec<u8> {
    package_with_wrapped_a1(WRAPPED_A1)
}

fn package_with_wrapped_a1(wrapped_a1: &str) -> Vec<u8> {
    let sheet1 = format!(
        r#"<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="A1"><v>10</v></c><c r="H1"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>"#
    );
    let sheet2 = format!(
        r#"<worksheet xmlns="{MAIN}" xmlns:mc="{MC}" xmlns:x="{EXT}" mc:Ignorable="x">{}</worksheet>"#,
        sheet2_data().replace(WRAPPED_A1, wrapped_a1)
    );
    let parts = [
        (
            "[Content_Types].xml",
            concat!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
                r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
                r#"<Default Extension="xml" ContentType="application/xml"/>"#,
                r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#,
                r#"<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                r#"<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                "</Types>",
            )
            .to_owned(),
        ),
        (
            "_rels/.rels",
            format!(
                r#"<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
            ),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<workbook xmlns="{MAIN}" xmlns:r="{RELS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Sheet2" sheetId="2" r:id="rId2"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!(
                r#"<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{RELS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{RELS}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#
            ),
        ),
        ("xl/worksheets/sheet1.xml", sheet1),
        ("xl/worksheets/sheet2.xml", sheet2),
    ];
    let parts: Vec<_> = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[derive(Clone, Copy)]
enum XmlToken<'a> {
    Start {
        name: &'a str,
        raw: &'a str,
        start: usize,
    },
    Empty {
        name: &'a str,
        raw: &'a str,
        start: usize,
        end: usize,
    },
    End {
        name: &'a str,
        end: usize,
    },
    Text(&'a str),
}

fn is_xml_space(c: char) -> bool {
    matches!(c, ' ' | '\t' | '\r' | '\n')
}

fn tokenize_xml(xml: &str) -> Vec<XmlToken<'_>> {
    let bytes = xml.as_bytes();
    let mut tokens = Vec::new();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] != b'<' {
            let start = i;
            while i < bytes.len() && bytes[i] != b'<' {
                i += 1;
            }
            tokens.push(XmlToken::Text(&xml[start..i]));
            continue;
        }
        let start = i;
        if xml[i..].starts_with("<!--") {
            let Some(rel) = xml[i + 4..].find("-->") else {
                break;
            };
            i += 4 + rel + 3;
            continue;
        }
        if xml[i..].starts_with("<?") {
            let Some(rel) = xml[i + 2..].find("?>") else {
                break;
            };
            i += 2 + rel + 2;
            continue;
        }
        if xml[i..].starts_with("<![CDATA[") {
            let Some(rel) = xml[i + 9..].find("]]>") else {
                break;
            };
            tokens.push(XmlToken::Text(&xml[i + 9..i + 9 + rel]));
            i += 9 + rel + 3;
            continue;
        }
        if xml[i..].starts_with("</") {
            let mut j = i + 2;
            while j < bytes.len() && bytes[j] != b'>' {
                j += 1;
            }
            if j >= bytes.len() {
                break;
            }
            let name = xml[i + 2..j].trim();
            let end = j + 1;
            tokens.push(XmlToken::End { name, end });
            i = end;
            continue;
        }
        let mut j = i + 1;
        let mut quote: Option<u8> = None;
        while j < bytes.len() {
            let b = bytes[j];
            if let Some(q) = quote {
                if b == q {
                    quote = None;
                }
            } else if b == b'"' || b == b'\'' {
                quote = Some(b);
            } else if b == b'>' {
                break;
            }
            j += 1;
        }
        if j >= bytes.len() {
            break;
        }
        let end = j + 1;
        let inner = &xml[i + 1..j];
        let (empty, body) = match inner.strip_suffix('/') {
            Some(stripped) => (true, stripped),
            None => (false, inner),
        };
        let name_end = body.find(is_xml_space).unwrap_or(body.len());
        let name = &body[..name_end];
        let raw = &xml[start..end];
        if empty {
            tokens.push(XmlToken::Empty {
                name,
                raw,
                start,
                end,
            });
        } else {
            tokens.push(XmlToken::Start { name, raw, start });
        }
        i = end;
    }
    tokens
}

fn local_name(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, local)| local)
}

fn tag_attr<'a>(raw_tag: &'a str, want: &str) -> Option<&'a str> {
    let inner = raw_tag.strip_prefix('<')?.strip_suffix('>')?;
    let inner = inner.strip_suffix('/').unwrap_or(inner);
    let name_end = inner.find(is_xml_space).unwrap_or(inner.len());
    let mut rest = &inner[name_end..];
    loop {
        rest = rest.trim_start_matches(is_xml_space);
        if rest.is_empty() {
            return None;
        }
        let eq = rest.find('=')?;
        let key = rest[..eq].trim_end_matches(is_xml_space);
        let after_eq = rest[eq + 1..].trim_start_matches(is_xml_space);
        let quote = after_eq.as_bytes().first().copied()?;
        if quote != b'"' && quote != b'\'' {
            return None;
        }
        let after_quote = &after_eq[1..];
        let close = after_quote.find(quote as char)?;
        let value = &after_quote[..close];
        if key == want {
            return Some(value);
        }
        rest = &after_quote[close + 1..];
    }
}

fn unescape_xml(text: &str) -> String {
    text.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&amp;", "&")
}

fn parse_shared_strings_xml(xml: &str) -> Vec<String> {
    let tokens = tokenize_xml(xml);
    let mut stack: Vec<&str> = Vec::new();
    let mut strings = Vec::new();
    let mut current: Option<String> = None;
    for token in &tokens {
        match *token {
            XmlToken::Start { name, .. } => {
                let local = local_name(name);
                if local == "si" && stack.last() == Some(&"sst") {
                    current = Some(String::new());
                }
                stack.push(local);
            }
            XmlToken::Empty { name, .. } => {
                if local_name(name) == "si" && stack.last() == Some(&"sst") {
                    strings.push(String::new());
                }
            }
            XmlToken::End { name, .. } => {
                stack.pop();
                if local_name(name) == "si" && stack.last() == Some(&"sst") {
                    strings.push(current.take().unwrap_or_default());
                }
            }
            XmlToken::Text(text) => {
                if stack.last() == Some(&"t")
                    && !stack.contains(&"rPh")
                    && let Some(item) = &mut current
                {
                    item.push_str(&unescape_xml(text));
                }
            }
        }
    }
    strings
}

fn extract_cell_text(cell_xml: &str, shared_strings: &[String]) -> Option<String> {
    let tokens = tokenize_xml(cell_xml);
    let ty = match tokens.first()? {
        XmlToken::Start { raw, .. } | XmlToken::Empty { raw, .. } => tag_attr(raw, "t"),
        _ => None,
    };
    let mut stack: Vec<&str> = Vec::new();
    let mut v_text = String::new();
    let mut has_v = false;
    let mut is_text = String::new();
    let mut has_is = false;
    for token in &tokens {
        match *token {
            XmlToken::Start { name, .. } => {
                let local = local_name(name);
                if local == "v" {
                    has_v = true;
                }
                if local == "is" {
                    has_is = true;
                }
                stack.push(local);
            }
            XmlToken::Empty { name, .. } => {
                let local = local_name(name);
                if local == "v" {
                    has_v = true;
                }
                if local == "is" {
                    has_is = true;
                }
            }
            XmlToken::End { .. } => {
                stack.pop();
            }
            XmlToken::Text(text) => {
                let decoded = unescape_xml(text);
                if stack.last() == Some(&"v") {
                    v_text.push_str(&decoded);
                } else if stack.last() == Some(&"t")
                    && stack.contains(&"is")
                    && !stack.contains(&"rPh")
                {
                    is_text.push_str(&decoded);
                }
            }
        }
    }
    match ty {
        Some("inlineStr") if has_is => Some(is_text),
        Some("str") if has_v => Some(v_text),
        Some("s") if has_v => v_text
            .trim()
            .parse::<usize>()
            .ok()
            .and_then(|idx| shared_strings.get(idx).cloned()),
        _ => None,
    }
}

struct BranchObservation {
    kind: String,
    a1_cells: Vec<Option<String>>,
}

struct SheetInspection {
    standalone_a1: usize,
    wrappers: Vec<String>,
    choice_a1: usize,
    fallback_a1: usize,
}

fn inspect_sheet(sheet_xml: &str, shared_strings: &[String]) -> SheetInspection {
    let tokens = tokenize_xml(sheet_xml);
    let mut stack: Vec<&str> = Vec::new();
    let mut standalone_a1 = 0usize;
    let mut wrappers = Vec::new();
    let mut branches: Vec<BranchObservation> = Vec::new();
    let mut active_ac: Option<(usize, usize)> = None;
    let mut active_branch: Option<(usize, usize)> = None;
    let mut active_a1: Option<(usize, usize, Option<usize>, bool)> = None;

    for token in &tokens {
        match *token {
            XmlToken::Start { name, raw, start } => {
                let local = local_name(name);
                let in_sheet_data = stack.contains(&"sheetData");
                if in_sheet_data
                    && active_ac.is_none()
                    && active_a1.is_none()
                    && local == "AlternateContent"
                {
                    active_ac = Some((stack.len(), start));
                } else if active_ac.is_some()
                    && active_branch.is_none()
                    && matches!(local, "Choice" | "Fallback")
                {
                    let idx = branches.len();
                    branches.push(BranchObservation {
                        kind: local.to_owned(),
                        a1_cells: Vec::new(),
                    });
                    active_branch = Some((stack.len(), idx));
                } else if in_sheet_data
                    && active_a1.is_none()
                    && local == "c"
                    && tag_attr(raw, "r") == Some("A1")
                {
                    active_a1 = Some((
                        stack.len(),
                        start,
                        active_branch.map(|(_, idx)| idx),
                        active_ac.is_some(),
                    ));
                }
                stack.push(local);
            }
            XmlToken::Empty {
                name,
                raw,
                start,
                end,
            } => {
                let local = local_name(name);
                let in_sheet_data = stack.contains(&"sheetData");
                if in_sheet_data
                    && active_ac.is_none()
                    && active_a1.is_none()
                    && local == "AlternateContent"
                {
                    wrappers.push(sheet_xml[start..end].to_owned());
                } else if active_ac.is_some()
                    && active_branch.is_none()
                    && matches!(local, "Choice" | "Fallback")
                {
                    branches.push(BranchObservation {
                        kind: local.to_owned(),
                        a1_cells: Vec::new(),
                    });
                } else if in_sheet_data
                    && active_a1.is_none()
                    && local == "c"
                    && tag_attr(raw, "r") == Some("A1")
                {
                    let value = extract_cell_text(&sheet_xml[start..end], shared_strings);
                    match (active_branch.map(|(_, idx)| idx), active_ac.is_some()) {
                        (Some(idx), _) => branches[idx].a1_cells.push(value),
                        (None, false) => standalone_a1 += 1,
                        (None, true) => {}
                    }
                }
            }
            XmlToken::End { end, .. } => {
                stack.pop();
                let depth = stack.len();
                if let Some((a1_depth, a1_start, branch_idx, in_ac)) = active_a1
                    && depth == a1_depth
                {
                    active_a1 = None;
                    let value = extract_cell_text(&sheet_xml[a1_start..end], shared_strings);
                    match (branch_idx, in_ac) {
                        (Some(idx), _) => branches[idx].a1_cells.push(value),
                        (None, false) => standalone_a1 += 1,
                        (None, true) => {}
                    }
                }
                if let Some((branch_depth, _)) = active_branch
                    && depth == branch_depth
                {
                    active_branch = None;
                }
                if let Some((ac_depth, ac_start)) = active_ac
                    && depth == ac_depth
                {
                    active_ac = None;
                    wrappers.push(sheet_xml[ac_start..end].to_owned());
                }
            }
            XmlToken::Text(_) => {}
        }
    }

    let choice_a1 = branches
        .iter()
        .filter(|branch| branch.kind == "Choice")
        .map(|branch| branch.a1_cells.len())
        .sum();
    let fallback_a1 = branches
        .iter()
        .filter(|branch| branch.kind == "Fallback")
        .map(|branch| branch.a1_cells.len())
        .sum();

    SheetInspection {
        standalone_a1,
        wrappers,
        choice_a1,
        fallback_a1,
    }
}

fn sheet_data_xml(sheet_xml: &str) -> &str {
    let tokens = tokenize_xml(sheet_xml);
    let mut depth = 0usize;
    let mut start = None;
    for token in &tokens {
        match *token {
            XmlToken::Start {
                name,
                start: tag_start,
                ..
            } => {
                if depth == 1 && local_name(name) == "sheetData" {
                    start = Some(tag_start);
                }
                depth += 1;
            }
            XmlToken::Empty {
                name,
                start: tag_start,
                end,
                ..
            } => {
                if depth == 1 && local_name(name) == "sheetData" {
                    return &sheet_xml[tag_start..end];
                }
            }
            XmlToken::End { name, end } => {
                depth = depth.saturating_sub(1);
                if depth == 1 && local_name(name) == "sheetData" {
                    let tag_start = start.expect("sheetData start");
                    return &sheet_xml[tag_start..end];
                }
            }
            XmlToken::Text(_) => {}
        }
    }
    panic!("missing <sheetData> in worksheet XML: {sheet_xml}");
}

fn read_saved_sheet2(saved: &[u8]) -> (String, Vec<String>) {
    let parts = ooxml_opc::unzip_parts(saved).unwrap();
    let sheet_bytes = parts
        .iter()
        .find(|(path, _)| path.trim_start_matches('/') == "xl/worksheets/sheet2.xml")
        .map(|(_, bytes)| bytes.as_slice())
        .expect("saved archive should contain xl/worksheets/sheet2.xml");
    let sheet_xml = std::str::from_utf8(sheet_bytes)
        .expect("sheet2 XML should be UTF-8")
        .to_owned();
    let shared_strings = parts
        .iter()
        .find(|(path, _)| path.trim_start_matches('/') == "xl/sharedStrings.xml")
        .map(|(_, bytes)| {
            let xml = std::str::from_utf8(bytes).expect("sharedStrings XML should be UTF-8");
            parse_shared_strings_xml(xml)
        })
        .unwrap_or_default();
    (sheet_xml, shared_strings)
}

fn assert_wrapped_a1_preserved(sheet_xml: &str, shared_strings: &[String]) {
    let info = inspect_sheet(sheet_xml, shared_strings);
    assert_eq!(
        info.standalone_a1, 0,
        "expected no standalone <c r=\"A1\"> in saved sheet2 XML: {sheet_xml}"
    );
    assert_eq!(
        info.choice_a1, 1,
        "expected <c r=\"A1\"> once in mc:Choice: {sheet_xml}"
    );
    assert_eq!(
        info.fallback_a1, 1,
        "expected <c r=\"A1\"> once in mc:Fallback: {sheet_xml}"
    );
    assert_eq!(
        info.wrappers.len(),
        1,
        "expected one mc:AlternateContent wrapper: {sheet_xml}"
    );
    assert_eq!(
        info.wrappers[0], WRAPPED_A1,
        "expected mc:AlternateContent wrapper bytes to stay unchanged: {sheet_xml}"
    );
}

#[test]
fn observe_wrapped_cells_after_save() {
    let cases = [
        ("edit_sheet_keeps_wrapped_cell_once", SheetId(1), "B1", "=2+2"),
        (
            "recalc_dirtied_sheet_keeps_wrapped_cell_once",
            SheetId(0),
            "H1",
            "=2+3",
        ),
        ("untouched_sheet_keeps_wrapper_bytes", SheetId(0), "A1", "20"),
        (
            "edit_of_wrapped_cell_is_consistent_or_refused",
            SheetId(1),
            "A1",
            "updated",
        ),
    ];

    for (case, sheet, address, input) in cases {
        let mut workbook = Workbook::open(&package()).unwrap();
        let edit_result = workbook.edit_cell(
            sheet,
            cell_ref(address),
            input,
            CalculationOptions::default(),
        );
        let edit_status = if edit_result.is_ok() { "Ok" } else { "Err" };
        match workbook.save() {
            Ok(saved) => {
                let (sheet_xml, shared_strings) = read_saved_sheet2(&saved);
                let info = inspect_sheet(&sheet_xml, &shared_strings);
                eprintln!(
                    "OBSERVE case={case} edit={edit_status} standalone_a1={} choice_a1={} fallback_a1={} wrapper_present={}",
                    info.standalone_a1,
                    info.choice_a1,
                    info.fallback_a1,
                    !info.wrappers.is_empty()
                );
            }
            Err(error) => {
                eprintln!(
                    "OBSERVE case={case} edit={edit_status} save=Err({error}) standalone_a1=0 choice_a1=0 fallback_a1=0 wrapper_present=false"
                );
            }
        }
    }
}

#[test]
fn edit_sheet_keeps_wrapped_cell_once() {
    let mut workbook = Workbook::open(&package()).unwrap();
    workbook
        .edit_cell(
            SheetId(1),
            cell_ref("B1"),
            "=2+2",
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let (sheet_xml, shared_strings) = read_saved_sheet2(&saved);
    assert_wrapped_a1_preserved(&sheet_xml, &shared_strings);
}

#[test]
fn recalc_dirtied_sheet_keeps_wrapped_cell_once() {
    let mut workbook = Workbook::open(&package()).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            cell_ref("H1"),
            "=2+3",
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        workbook.model().value(SheetId(1), cell_ref("C1")),
        CellValue::Text {
            value: "=2+3".to_owned(),
        }
    );
    let saved = workbook.save().unwrap();
    let (sheet_xml, shared_strings) = read_saved_sheet2(&saved);
    assert_wrapped_a1_preserved(&sheet_xml, &shared_strings);
}

#[test]
fn untouched_sheet_keeps_wrapper_bytes() {
    let mut workbook = Workbook::open(&package()).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            cell_ref("A1"),
            "20",
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    let (sheet_xml, _) = read_saved_sheet2(&saved);
    assert_eq!(sheet_data_xml(&sheet_xml), sheet2_data());
}

#[test]
fn edit_of_wrapped_cell_is_consistent_or_refused() {
    let bytes = package();
    let baseline = Workbook::open(&bytes).unwrap().save().unwrap();
    let mut workbook = Workbook::open(&bytes).unwrap();
    let result = workbook.edit_cell(
        SheetId(1),
        cell_ref("A1"),
        "updated",
        CalculationOptions::default(),
    );
    assert_alternate_cell_error(result.unwrap_err());
    assert_eq!(workbook.save().unwrap(), baseline);
}

#[test]
fn style_edit_of_wrapped_cell_is_refused() {
    let bytes = package();
    let mut workbook = Workbook::open(&bytes).unwrap();
    let baseline = workbook.save().unwrap();
    let at = cell_ref("A1");
    let result = workbook.set_range_number_format(
        SheetId(1),
        CellRange::new(at, at),
        NumberFormatMutation::Percent,
        CalculationOptions::default(),
    );
    assert_alternate_cell_error(result.unwrap_err());
    assert_eq!(workbook.save().unwrap(), baseline);
}

#[test]
fn changed_wrapped_formula_cache_refuses_save() {
    let bytes = package_with_wrapped_a1(concat!(
        "<mc:AlternateContent>",
        r#"<mc:Choice Requires="x"><c r="A1"><f>Sheet1!A1*2</f><v>20</v></c></mc:Choice>"#,
        r#"<mc:Fallback><c r="A1"><f>Sheet1!A1*2</f><v>20</v></c></mc:Fallback>"#,
        "</mc:AlternateContent>",
    ));
    let mut workbook = Workbook::open(&bytes).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            cell_ref("A1"),
            "15",
            CalculationOptions::default(),
        )
        .unwrap();
    assert_eq!(
        workbook.model().value(SheetId(1), cell_ref("A1")),
        CellValue::Number { value: 30.0 }
    );
    assert_alternate_cell_error(workbook.save().unwrap_err());
}

fn assert_alternate_cell_error(error: Error) {
    assert!(matches!(
        error,
        Error::Spreadsheet(xlsx_parse::ParseError::UnsupportedEdit(message))
            if message.contains("A1") && message.contains("mc:AlternateContent")
    ));
}
