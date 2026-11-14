use betteroffice_xlsx::{CalculationOptions, CellRange, DifferentialFormat, Op, SheetId, Workbook};

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const X14_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const SHEET_PART: &str = "xl/worksheets/sheet1.xml";
const OTHER_SQREF: &str = "C1:C5";
const OTHER_GROUP: &str = r#"<conditionalFormatting sqref="C1:C5"><cfRule type="expression" dxfId="0" priority="2"><formula>Data!$B$2&gt;0</formula></cfRule></conditionalFormatting>"#;
const EXT_FORMULA_RULE: &str = r#"<cfRule type="expression" dxfId="0" priority="1"><formula>Data!$A$1&gt;0</formula><extLst><ext uri="{X}"><y:formula>Data!$B$1</y:formula></ext></extLst></cfRule>"#;
const SECOND_RULE: &str = r#"<cfRule type="expression" dxfId="0" priority="3"><formula>$A$1&gt;1</formula></cfRule>"#;

#[derive(Clone, Copy, Debug)]
enum Placement {
    Direct,
    Ext,
    ExtWrapper,
    ExtraFirst,
}

impl Placement {
    fn label(self) -> &'static str {
        match self {
            Placement::Direct => "direct",
            Placement::Ext => "extLst/ext",
            Placement::ExtWrapper => "extLst/ext/wrapper",
            Placement::ExtraFirst => "extLst/ext/extra-first",
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum Nested {
    None,
    Bare,
    Dxf,
}

impl Nested {
    fn label(self) -> &'static str {
        match self {
            Nested::None => "none",
            Nested::Bare => "bare",
            Nested::Dxf => "dxf",
        }
    }
}

struct Piece {
    xml: String,
    values: Vec<String>,
    nested: bool,
}

fn markup(xml: &str) -> Piece {
    Piece {
        xml: xml.to_owned(),
        values: Vec::new(),
        nested: false,
    }
}

fn values(parts: &[Piece]) -> Vec<String> {
    parts.iter().flat_map(|part| part.values.clone()).collect()
}

// main_prefix: values up to the first nested closing tag; suffix: the values after it
struct Fixture {
    xml: String,
    main_prefix: Vec<String>,
    suffix: Vec<String>,
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
    fn id(self) -> &'static str {
        match self {
            Edit::None => "none",
            Edit::AddNewGroup => "add-new-group",
            Edit::AddSameGroup => "add-same-group",
            Edit::RemoveOtherRule => "remove-other-rule",
            Edit::RenameReferencedSheet => "rename-referenced-sheet",
        }
    }
}

fn nested_element(prefixed: bool, nested: Nested, payload: bool) -> Piece {
    let p = if prefixed { "x14:" } else { "" };
    if let Nested::None = nested {
        return markup("");
    }
    let mut inner = String::new();
    let mut values = Vec::new();
    if payload {
        inner.push_str(&format!(
            r#"<{p}formula>$B$3&gt;0</{p}formula><{p}colorScale><{p}cfvo type="formula" val="$B$4"/></{p}colorScale>"#
        ));
        values = vec!["$B$3>0".to_owned(), "$B$4".to_owned()];
    }
    if let Nested::Dxf = nested {
        inner.push_str(&format!(r#"<{p}dxf><{p}font><{p}b/></{p}font></{p}dxf>"#));
    }
    let xml = if inner.is_empty() {
        format!(r#"<{p}cfRule type="expression" priority="9"/>"#)
    } else {
        format!(r#"<{p}cfRule type="expression" priority="9">{inner}</{p}cfRule>"#)
    };
    Piece {
        xml,
        values,
        nested: true,
    }
}

fn nested_rule(placement: Placement, prefixed: bool, nested: Nested, payload: bool) -> Fixture {
    let nested = nested_element(prefixed, nested, payload);
    // a formula and a formula cfvo below the rule's own children, inside its extLst
    let extra = if payload {
        r#"<x14:extra><x14:formula>$B$5</x14:formula><x14:cfvo type="formula" val="$B$6"/></x14:extra>"#
    } else {
        ""
    };
    let extra = Piece {
        xml: extra.to_owned(),
        values: if payload {
            vec!["$B$5".to_owned(), "$B$6".to_owned()]
        } else {
            Vec::new()
        },
        nested: false,
    };
    let own = Piece {
        xml: "<formula>$B$1&gt;0</formula>".to_owned(),
        values: vec!["$B$1>0".to_owned()],
        nested: false,
    };
    let (open, tail, close) = (
        r#"<extLst><ext uri="urn:nested">"#,
        r#"<x14:tail v="KEEPME"/>"#,
        "</ext></extLst>",
    );
    let parts = match placement {
        Placement::Direct => vec![
            own,
            nested,
            markup(r#"<extLst><ext uri="urn:tail">"#),
            markup(tail),
            extra,
            markup(close),
        ],
        Placement::Ext => vec![
            own,
            markup(open),
            nested,
            markup(tail),
            extra,
            markup(close),
        ],
        Placement::ExtWrapper => vec![
            own,
            markup(open),
            markup("<x14:wrapper>"),
            nested,
            markup(tail),
            markup("</x14:wrapper>"),
            extra,
            markup(close),
        ],
        Placement::ExtraFirst => vec![
            own,
            markup(open),
            extra,
            nested,
            markup(tail),
            markup(close),
        ],
    };
    let split = parts
        .iter()
        .position(|part| part.nested)
        .map_or(parts.len(), |at| at + 1);
    let body: String = parts.iter().map(|part| part.xml.as_str()).collect();
    Fixture {
        xml: format!(
            r#"<cfRule type="expression" dxfId="0" priority="1">{body}</cfRule>"#
        ),
        main_prefix: values(&parts[..split]),
        suffix: values(&parts[split..]),
    }
}

fn worksheet(groups: &str) -> String {
    let mut rows = String::new();
    for row in 1..=5u32 {
        rows.push_str(&format!(
            r#"<row r="{row}"><c r="A{row}"><v>{row}</v></c><c r="B{row}"><v>{row}</v></c></row>"#
        ));
    }
    format!(
        r#"<worksheet xmlns="{MAIN_NS}" xmlns:x14="{X14_NS}" xmlns:y="urn:y"><sheetData>{rows}</sheetData>{groups}</worksheet>"#
    )
}

fn sheet(rule: &str) -> String {
    worksheet(&format!(
        r#"<conditionalFormatting sqref="A1:A5">{rule}</conditionalFormatting>{OTHER_GROUP}"#
    ))
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

fn only_markup(mut gap: &str) -> bool {
    loop {
        gap = gap.trim_start();
        if gap.is_empty() {
            return true;
        }
        let end = if gap.starts_with("<?") {
            "?>"
        } else if gap.starts_with("<!--") {
            "-->"
        } else {
            return false;
        };
        let Some(at) = gap.find(end) else {
            return false;
        };
        gap = &gap[at + end.len()..];
    }
}

// one root, matched tags, nothing but whitespace outside the root, and the part reopens on its own
fn well_formed(xml: &str) -> Result<(), String> {
    let xml = xml.strip_prefix('\u{feff}').unwrap_or(xml);
    let mut stack: Vec<&str> = Vec::new();
    let mut roots = 0;
    let mut last = 0;
    for tag in tags(xml)? {
        let (name, at) = (tag.name, tag.start);
        if stack.is_empty() && !only_markup(&xml[last..at]) {
            return Err(format!("text outside the root before byte {at}"));
        }
        last = tag.end;
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
    if !only_markup(&xml[last..]) {
        return Err(format!("text outside the root after byte {last}"));
    }
    if roots != 1 {
        return Err(format!("{roots} root elements"));
    }
    Workbook::open(&package(xml))
        .map(|_| ())
        .map_err(|error| format!("part alone does not reopen {}", trim(&format!("{error:?}"))))
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

fn rule_formulas(wb: &Workbook) -> Vec<String> {
    wb.model().sheets[0]
        .conditional_formats
        .iter()
        .flat_map(|group| group.rules.iter())
        .flat_map(|rule| rule.formulas.iter().cloned())
        .collect()
}

// one line per rule the model read, compared across trees outside the test
fn print_reading(case: &str, wb: &Workbook) {
    let groups = &wb.model().sheets[0].conditional_formats;
    println!("READ case={case} groups={} rules={}", groups.len(), model_rule_count(wb));
    for (g, group) in groups.iter().enumerate() {
        for (r, rule) in group.rules.iter().enumerate() {
            println!(
                "READ case={case} rule={g}.{r} sqref={} type={} operator={} priority={} dxf={} formulas={:?}",
                group.sqref,
                rule.rule_type,
                rule.operator.as_deref().unwrap_or("-"),
                rule.priority,
                rule.dxf_id.map_or("-".to_owned(), |dxf| dxf.to_string()),
                rule.formulas
            );
        }
    }
}

fn sheet_part(saved: &[u8]) -> Option<String> {
    ooxml_opc::unzip_parts(saved).ok().and_then(|parts| {
        parts
            .into_iter()
            .find(|(name, _)| name == SHEET_PART)
            .map(|(_, data)| String::from_utf8_lossy(&data).into_owned())
    })
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

fn panic_text(panic: Box<dyn std::any::Any + Send>) -> String {
    let text = panic
        .downcast_ref::<String>()
        .cloned()
        .or_else(|| panic.downcast_ref::<&str>().map(|text| (*text).to_owned()))
        .unwrap_or_default();
    format!("panicked {text}")
}

fn run_case(case: &str, fixture: &Fixture, edit: Edit) -> Vec<String> {
    let source_rule = fixture.xml.as_str();
    let mut problems = Vec::new();
    let mut wb = match Workbook::open(&package(&sheet(source_rule))) {
        Ok(wb) => wb,
        Err(error) => {
            println!("READ case={case} open failed");
            return vec![format!("open failed {}", trim(&format!("{error:?}")))];
        }
    };
    print_reading(case, &wb);
    let expected: Vec<String> = fixture
        .main_prefix
        .iter()
        .chain(&fixture.suffix)
        .cloned()
        .collect();
    let read = wb.model().sheets[0]
        .conditional_formats
        .first()
        .and_then(|group| group.rules.first())
        .map(|rule| rule.formulas.clone());
    assert_eq!(read, Some(expected), "case={case} rule=0.0 formulas");
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
    let Some(part) = sheet_part(&saved) else {
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

fn grid_case(
    placement: Placement,
    prefixed: bool,
    nested: Nested,
    payload: bool,
    edit: Edit,
) -> Option<String> {
    let fixture = nested_rule(placement, prefixed, nested, payload);
    let case = format!(
        "at={},prefix={},nested={},payload={},op={}",
        placement.label(),
        if prefixed { "x14" } else { "none" },
        nested.label(),
        if payload { "formula+cfvo" } else { "none" },
        edit.id()
    );
    if !matches!(nested, Nested::None) {
        println!(
            "EXPECT case={case} main_prefix={:?} suffix={:?}",
            fixture.main_prefix, fixture.suffix
        );
    }
    let problems = std::panic::catch_unwind(|| run_case(&case, &fixture, edit))
        .unwrap_or_else(|panic| vec![panic_text(panic)]);
    (!problems.is_empty()).then(|| format!("{case}: {}", problems.join("; ")))
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
            for nested in [Nested::None, Nested::Bare, Nested::Dxf] {
                for payload in [false, true] {
                    for edit in edits {
                        total += 1;
                        failures.extend(grid_case(placement, prefixed, nested, payload, edit));
                    }
                }
            }
        }
    }
    // extra values ahead of the nested element: main already reads every value
    let first = nested_rule(Placement::ExtraFirst, true, Nested::Dxf, true);
    let order = ["$B$1>0", "$B$5", "$B$6", "$B$3>0", "$B$4"];
    assert_eq!(first.main_prefix, order.map(str::to_owned).to_vec());
    assert!(first.suffix.is_empty(), "suffix {:?}", first.suffix);
    total += 1;
    failures.extend(grid_case(Placement::ExtraFirst, true, Nested::Dxf, true, Edit::None));
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

#[test]
fn rule_formula_in_ext_list_is_rewritten_by_sheet_rename() {
    let source = worksheet(&format!(
        r#"<conditionalFormatting sqref="A1:A5">{EXT_FORMULA_RULE}</conditionalFormatting>"#
    ));
    let mut wb = Workbook::open(&package(&source)).unwrap();
    print_reading("ext-formula-rename", &wb);
    let rename = Op::RenameSheet {
        sheet: SheetId(1),
        name: "Info".into(),
    };
    let result = wb.apply_ops(vec![rename], CalculationOptions::default());
    assert!(
        result.is_ok(),
        "RenameSheet Data -> Info refused: {:?}",
        result.as_ref().err()
    );
    result.unwrap();
    let renamed = vec!["Info!$A$1>0".to_owned(), "Info!$B$1".to_owned()];
    assert_eq!(rule_formulas(&wb), renamed);
    let saved = wb.save().unwrap();
    let part = sheet_part(&saved).unwrap();
    well_formed(&part).unwrap();
    assert!(part.contains("<formula>Info!$A$1&gt;0</formula>"), "{part}");
    assert!(part.contains("<y:formula>Info!$B$1</y:formula>"), "{part}");
    assert!(!part.contains("Data!"), "{part}");
    assert_eq!(rule_formulas(&Workbook::open(&saved).unwrap()), renamed);
}

fn nested_group_case(case: &str, first: &str, second: &str) -> Vec<String> {
    let mut problems = Vec::new();
    let source = worksheet(&format!(
        r#"<conditionalFormatting sqref="A1:A5">{first}{second}</conditionalFormatting>{OTHER_GROUP}"#
    ));
    let mut wb = match Workbook::open(&package(&source)) {
        Ok(wb) => wb,
        Err(error) => return vec![format!("open failed {}", trim(&format!("{error:?}")))],
    };
    print_reading(case, &wb);
    let groups = &wb.model().sheets[0].conditional_formats;
    let read: Vec<(String, usize)> = groups
        .iter()
        .map(|group| (group.sqref.clone(), group.rules.len()))
        .collect();
    let expected = vec![("A1:A5".to_owned(), 2), (OTHER_SQREF.to_owned(), 1)];
    if read != expected {
        problems.push(format!("read groups {read:?}, expected {expected:?}"));
    }
    let raw = groups
        .first()
        .and_then(|group| group.rules.first())
        .and_then(|rule| rule.raw_xml.clone());
    if raw.as_deref() != Some(first) {
        problems.push(format!("rule 1 raw xml read as {:?}", raw.map(|text| trim(&text))));
    }
    if let Err(error) = wb.apply_ops(vec![add("E1:E5")], CalculationOptions::default()) {
        problems.push(format!("op refused {}", trim(&format!("{error:?}"))));
        return problems;
    }
    let model = model_rule_count(&wb);
    let saved = match wb.save() {
        Ok(saved) => saved,
        Err(error) => {
            problems.push(format!("save failed {}", trim(&format!("{error:?}"))));
            return problems;
        }
    };
    let Some(part) = sheet_part(&saved) else {
        problems.push(format!("saved package has no {SHEET_PART}"));
        return problems;
    };
    if let Err(error) = well_formed(&part) {
        problems.push(format!("saved sheet not well formed: {error}"));
    }
    let (count, texts) = outer_rules(&part);
    if model != 4 || count != model {
        problems.push(format!("model rules={model} saved cfRule={count}, expected 4"));
    }
    if !texts.contains(&first) {
        problems.push("rule 1 not saved byte-equal".to_owned());
    }
    if !texts.contains(&second) {
        problems.push("rule 2 not saved byte-equal".to_owned());
    }
    match Workbook::open(&saved) {
        Ok(reopened) if model_rule_count(&reopened) == model => {}
        Ok(reopened) => problems.push(format!(
            "reopened rules={} saved model rules={model}",
            model_rule_count(&reopened)
        )),
        Err(error) => problems.push(format!("reopen failed {}", trim(&format!("{error:?}")))),
    }
    problems
}

#[test]
fn nested_conditional_formatting_tag_in_a_rule_keeps_the_group_open() {
    let mut failures = Vec::new();
    for prefixed in [false, true] {
        let p = if prefixed { "x14:" } else { "" };
        let first = format!(
            r#"<cfRule type="expression" dxfId="0" priority="1"><formula>$B$1&gt;0</formula><extLst><ext uri="{{X}}"><{p}conditionalFormatting/></ext></extLst></cfRule>"#
        );
        let case = format!("nested-group,prefix={}", if prefixed { "x14" } else { "none" });
        let problems = std::panic::catch_unwind(|| nested_group_case(&case, &first, SECOND_RULE))
            .unwrap_or_else(|_| vec!["panicked".to_owned()]);
        if !problems.is_empty() {
            failures.push(format!("{case}: {}", problems.join("; ")));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}
