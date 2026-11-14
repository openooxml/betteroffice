use betteroffice_xlsx::{CalculationOptions, CellRange, DifferentialFormat, Op, SheetId, Workbook};

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const X14_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const SHEET_PART: &str = "xl/worksheets/sheet1.xml";
const OTHER_SQREF: &str = "C1:C5";

#[derive(Clone, Copy, Debug)]
enum Placement {
    Direct,
    Ext,
    ExtWrapper,
}

impl Placement {
    fn label(self) -> &'static str {
        match self {
            Placement::Direct => "direct",
            Placement::Ext => "extLst/ext",
            Placement::ExtWrapper => "extLst/ext/wrapper",
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum Edit {
    None,
    AddNewGroup,
    AddSameGroup,
    RemoveOtherRule,
    RenameReferencedSheet,
}

impl Edit {
    fn label(self) -> &'static str {
        match self {
            Edit::None => "none",
            Edit::AddNewGroup => "AddConditionalFormat new group",
            Edit::AddSameGroup => "AddConditionalFormat same group",
            Edit::RemoveOtherRule => "RemoveConditionalFormat other rule",
            Edit::RenameReferencedSheet => "RenameSheet referenced by other rule",
        }
    }
}

fn nested_element(prefixed: bool, children: bool) -> String {
    let p = if prefixed { "x14:" } else { "" };
    if children {
        format!(
            r#"<{p}cfRule type="expression" priority="9"><{p}dxf><{p}font><{p}b/></{p}font></{p}dxf></{p}cfRule>"#
        )
    } else {
        format!(r#"<{p}cfRule type="expression" priority="9"/>"#)
    }
}

fn nested_rule(placement: Placement, prefixed: bool, children: bool) -> String {
    let nested = nested_element(prefixed, children);
    let body = match placement {
        Placement::Direct => {
            format!(r#"{nested}<extLst><ext uri="urn:tail"><x14:tail v="KEEPME"/></ext></extLst>"#)
        }
        Placement::Ext => format!(
            r#"<extLst><ext uri="urn:nested">{nested}<x14:tail v="KEEPME"/></ext></extLst>"#
        ),
        Placement::ExtWrapper => format!(
            r#"<extLst><ext uri="urn:nested"><x14:wrapper>{nested}<x14:tail v="KEEPME"/></x14:wrapper></ext></extLst>"#
        ),
    };
    format!(
        r#"<cfRule type="expression" dxfId="0" priority="1"><formula>$B$1&gt;0</formula>{body}</cfRule>"#
    )
}

fn sheet(rule: &str) -> String {
    let mut rows = String::new();
    for row in 1..=5u32 {
        rows.push_str(&format!(
            r#"<row r="{row}"><c r="A{row}"><v>{row}</v></c><c r="B{row}"><v>{row}</v></c></row>"#
        ));
    }
    format!(
        r#"<worksheet xmlns="{MAIN_NS}" xmlns:x14="{X14_NS}"><sheetData>{rows}</sheetData><conditionalFormatting sqref="A1:A5">{rule}</conditionalFormatting><conditionalFormatting sqref="{OTHER_SQREF}"><cfRule type="expression" dxfId="0" priority="2"><formula>Data!$B$2&gt;0</formula></cfRule></conditionalFormatting></worksheet>"#
    )
}

fn package(sheet: &str) -> Vec<u8> {
    let parts: Vec<(String, Vec<u8>)> = vec![
        (
            "xl/workbook.xml".into(),
            format!(
                r#"<workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId3"/></sheets></workbook>"#
            )
            .into_bytes(),
        ),
        (SHEET_PART.into(), sheet.as_bytes().to_vec()),
        (
            "xl/worksheets/sheet2.xml".into(),
            format!(
                r#"<worksheet xmlns="{MAIN_NS}"><sheetData><row r="2"><c r="B2"><v>1</v></c></row></sheetData></worksheet>"#
            )
            .into_bytes(),
        ),
        (
            "xl/styles.xml".into(),
            format!(
                r#"<styleSheet xmlns="{MAIN_NS}"><cellXfs count="1"><xf/></cellXfs><dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>"#
            )
            .into_bytes(),
        ),
        (
            "xl/_rels/workbook.xml.rels".into(),
            format!(
                r#"<Relationships><Relationship Id="rId1" Type="{REL_NS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{REL_NS}/styles" Target="styles.xml"/><Relationship Id="rId3" Type="{REL_NS}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#
            )
            .into_bytes(),
        ),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

struct Tag<'a> {
    start: usize,
    end: usize,
    name: &'a str,
    closing: bool,
    empty: bool,
}

fn local(name: &str) -> &str {
    name.rsplit(':').next().unwrap_or(name)
}

fn skip_to(xml: &str, from: usize, marker: &str) -> Result<usize, String> {
    xml[from..]
        .find(marker)
        .map(|at| from + at + marker.len())
        .ok_or_else(|| format!("unterminated markup at byte {from}"))
}

fn tags(xml: &str) -> Result<Vec<Tag<'_>>, String> {
    let bytes = xml.as_bytes();
    let mut found = Vec::new();
    let mut at = 0;
    while let Some(open) = xml[at..].find('<') {
        let start = at + open;
        let rest = &xml[start..];
        if rest.starts_with("<?") {
            at = skip_to(xml, start, "?>")?;
            continue;
        }
        if rest.starts_with("<!--") {
            at = skip_to(xml, start, "-->")?;
            continue;
        }
        if rest.starts_with("<![CDATA[") {
            at = skip_to(xml, start, "]]>")?;
            continue;
        }
        if rest.starts_with("<!") {
            at = skip_to(xml, start, ">")?;
            continue;
        }
        let mut quote = None;
        let mut close = None;
        for (index, byte) in bytes.iter().enumerate().skip(start + 1) {
            match (quote, *byte) {
                (Some(q), b) if b == q => quote = None,
                (Some(_), _) => {}
                (None, b'"' | b'\'') => quote = Some(*byte),
                (None, b'<') => return Err(format!("'<' inside a tag at byte {index}")),
                (None, b'>') => {
                    close = Some(index);
                    break;
                }
                _ => {}
            }
        }
        let close = close.ok_or_else(|| format!("unterminated tag at byte {start}"))?;
        let end = close + 1;
        let tag = &xml[start..end];
        let closing = tag.starts_with("</");
        let empty = !closing && tag.ends_with("/>");
        let name = tag[1..tag.len() - 1]
            .trim_start_matches('/')
            .split(|c: char| c.is_whitespace() || c == '/')
            .next()
            .unwrap_or("");
        if name.is_empty() {
            return Err(format!("tag without a name at byte {start}"));
        }
        found.push(Tag {
            start,
            end,
            name,
            closing,
            empty,
        });
        at = end;
    }
    Ok(found)
}

fn well_formed(xml: &str) -> Result<(), String> {
    let mut stack: Vec<&str> = Vec::new();
    let mut roots = 0;
    for tag in tags(xml)? {
        let (name, at) = (tag.name, tag.start);
        if tag.closing {
            match stack.pop() {
                Some(open) if open == name => {}
                Some(open) => return Err(format!("</{name}> closes <{open}> at byte {at}")),
                None => return Err(format!("</{name}> with nothing open at byte {at}")),
            }
            continue;
        }
        if stack.is_empty() {
            roots += 1;
        }
        if !tag.empty {
            stack.push(tag.name);
        }
    }
    if let Some(open) = stack.last() {
        return Err(format!("<{open}> is never closed ({} open)", stack.len()));
    }
    if roots != 1 {
        return Err(format!("{roots} root elements"));
    }
    Ok(())
}

// cfRule children of conditionalFormatting: start tag count and the closed ones' text
fn outer_rules(xml: &str) -> (usize, Vec<&str>) {
    let Ok(all) = tags(xml) else {
        return (0, Vec::new());
    };
    let mut stack: Vec<(&str, Option<usize>)> = Vec::new();
    let mut count = 0;
    let mut texts = Vec::new();
    for tag in &all {
        if tag.closing {
            if let Some((_, Some(start))) = stack.pop() {
                texts.push(&xml[start..tag.end]);
            }
            continue;
        }
        let outer = local(tag.name) == "cfRule"
            && stack
                .last()
                .is_some_and(|(parent, _)| local(parent) == "conditionalFormatting");
        if outer {
            count += 1;
        }
        if tag.empty {
            if outer {
                texts.push(&xml[tag.start..tag.end]);
            }
        } else {
            stack.push((tag.name, outer.then_some(tag.start)));
        }
    }
    (count, texts)
}

fn trim(text: &str) -> String {
    if text.chars().count() > 200 {
        format!("{}...", text.chars().take(200).collect::<String>())
    } else {
        text.to_owned()
    }
}

fn model_rule_count(wb: &Workbook) -> usize {
    wb.model().sheets[0]
        .conditional_formats
        .iter()
        .map(|group| group.rules.len())
        .sum()
}

fn add(range: &str) -> Op {
    Op::AddConditionalFormat {
        sheet: SheetId(0),
        range: CellRange::parse_a1(range).unwrap(),
        kind: "expression".into(),
        operator: None,
        formulas: vec!["$A$1>2".into()],
        format: DifferentialFormat {
            font_color: Some("#FF0000".into()),
            ..Default::default()
        },
    }
}

fn op_for(edit: Edit, wb: &Workbook) -> Result<Option<Op>, String> {
    match edit {
        Edit::None => Ok(None),
        Edit::AddNewGroup => Ok(Some(add("E1:E5"))),
        Edit::AddSameGroup => Ok(Some(add("A1:A5"))),
        Edit::RemoveOtherRule => {
            let id = wb.model().sheets[0]
                .conditional_formats
                .iter()
                .find(|group| group.sqref == OTHER_SQREF)
                .and_then(|group| group.rules.first())
                .map(|rule| rule.id.clone())
                .ok_or_else(|| format!("no rule read in {OTHER_SQREF}"))?;
            Ok(Some(Op::RemoveConditionalFormat {
                sheet: SheetId(0),
                id,
            }))
        }
        Edit::RenameReferencedSheet => Ok(Some(Op::RenameSheet {
            sheet: SheetId(1),
            name: "Inputs".into(),
        })),
    }
}

fn run_case(source_rule: &str, edit: Edit) -> Vec<String> {
    let mut problems = Vec::new();
    let mut wb = match Workbook::open(&package(&sheet(source_rule))) {
        Ok(wb) => wb,
        Err(error) => return vec![format!("open failed {}", trim(&format!("{error:?}")))],
    };
    match op_for(edit, &wb) {
        Ok(None) => {}
        Ok(Some(op)) => {
            if let Err(error) = wb.apply_ops(vec![op], CalculationOptions::default()) {
                return vec![format!("op refused {}", trim(&format!("{error:?}")))];
            }
        }
        Err(error) => return vec![error],
    }
    let model = model_rule_count(&wb);
    let saved = match wb.save() {
        Ok(saved) => saved,
        Err(error) => return vec![format!("save failed {}", trim(&format!("{error:?}")))],
    };
    if let Err(error) = Workbook::open(&saved) {
        problems.push(format!("reopen failed {}", trim(&format!("{error:?}"))));
    }
    let part = ooxml_opc::unzip_parts(&saved).ok().and_then(|parts| {
        parts
            .into_iter()
            .find(|(name, _)| name == SHEET_PART)
            .map(|(_, data)| String::from_utf8_lossy(&data).into_owned())
    });
    let Some(part) = part else {
        problems.push(format!("saved package has no {SHEET_PART}"));
        return problems;
    };
    if let Err(error) = well_formed(&part) {
        problems.push(format!("saved sheet not well formed: {error}"));
    }
    let (count, texts) = outer_rules(&part);
    if count != model {
        problems.push(format!("model rules={model} saved cfRule={count}"));
    }
    if !texts.contains(&source_rule) {
        let first = texts.first().map_or("none".to_owned(), |text| trim(text));
        problems.push(format!(
            "nested rule not saved as its source; first saved rule {first}"
        ));
    }
    problems
}

#[test]
fn conditional_format_rule_with_nested_elements_is_kept_on_save() {
    let edits = [
        Edit::None,
        Edit::AddNewGroup,
        Edit::AddSameGroup,
        Edit::RemoveOtherRule,
        Edit::RenameReferencedSheet,
    ];
    let mut total = 0;
    let mut failures = Vec::new();
    for placement in [Placement::Direct, Placement::Ext, Placement::ExtWrapper] {
        for prefixed in [false, true] {
            for children in [false, true] {
                let rule = nested_rule(placement, prefixed, children);
                for edit in edits {
                    total += 1;
                    let label = format!(
                        "at={} prefix={} nested={} op={}",
                        placement.label(),
                        if prefixed { "x14" } else { "none" },
                        if children { "children" } else { "empty" },
                        edit.label()
                    );
                    let problems = std::panic::catch_unwind(|| run_case(&rule, edit))
                        .unwrap_or_else(|_| vec!["panicked".to_owned()]);
                    if !problems.is_empty() {
                        failures.push(format!("{label}: {}", problems.join("; ")));
                    }
                }
            }
        }
    }
    failures.sort_by_key(String::len);
    let failed = failures.len();
    let shortest = failures.iter().take(10).cloned().collect::<Vec<_>>();
    println!("CASES total={total} failed={failed}");
    for failure in &shortest {
        println!("FAIL {failure}");
    }
    assert!(
        failures.is_empty(),
        "CASES total={total} failed={failed}\n{}",
        shortest.join("\n")
    );
}
