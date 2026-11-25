use std::fmt::Write;
use std::panic::{AssertUnwindSafe, catch_unwind};

use betteroffice_xlsx::{
    CalculationOptions, CalculationRequest, CellRange, CellRef, CellValue, Color, EditHistory,
    EditOperation, EditRequest, EditSource, EditStep, Hyperlink, RangeAddress, RangeTarget,
    Workbook, WorkbookModel,
};
use sha2::{Digest, Sha256};
use xlsx_ops::PreservedCellFormat;

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const CHART: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const HYPERLINK_LIMIT: usize = 16 * 1024 * 1024;
const CHART_LIMIT: usize = 8 * 1024 * 1024;
const STYLE_LIMIT: usize = 64 * 1024;
const HYPERLINK_ERROR: &str =
    "invalid collaborative workbook state: sheet hyperlink state exceeds its size limit";
const CHART_ERROR: &str =
    "invalid collaborative workbook state: sheet chart state exceeds its size limit";

fn package(
    sheet_data: &str,
    sheet_tail: &str,
    sheet_rels: Option<String>,
    styles: &str,
) -> Vec<u8> {
    zip_parts(package_parts(sheet_data, sheet_tail, sheet_rels, styles))
}

fn package_parts(
    sheet_data: &str,
    sheet_tail: &str,
    sheet_rels: Option<String>,
    styles: &str,
) -> Vec<(String, String)> {
    let mut parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels".to_owned(),
            format!(r#"<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#),
        ),
        (
            "xl/workbook.xml".to_owned(),
            format!(r#"<workbook xmlns="{MAIN}" xmlns:r="{RELS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#),
        ),
        (
            "xl/_rels/workbook.xml.rels".to_owned(),
            format!(r#"<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{RELS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{RELS}/styles" Target="styles.xml"/></Relationships>"#),
        ),
        (
            "xl/worksheets/sheet1.xml".to_owned(),
            format!(r#"<worksheet xmlns="{MAIN}" xmlns:r="{RELS}"><sheetData>{sheet_data}</sheetData>{sheet_tail}</worksheet>"#),
        ),
        ("xl/styles.xml".to_owned(), styles.to_owned()),
    ];
    if let Some(rels) = sheet_rels {
        parts.push(("xl/worksheets/_rels/sheet1.xml.rels".to_owned(), rels));
    }
    parts
}

fn zip_parts(parts: Vec<(String, String)>) -> Vec<u8> {
    let parts: Vec<_> = parts
        .into_iter()
        .map(|(name, xml)| (name, xml.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn styles(size: &str) -> String {
    format!(
        r#"<styleSheet xmlns="{MAIN}"><fonts count="1"><font><sz val="{size}"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>"#
    )
}

fn outcome<T: std::fmt::Debug>(run: impl FnOnce() -> T) -> (bool, String) {
    match catch_unwind(AssertUnwindSafe(run)) {
        Ok(value) => (true, format!("{value:?}")),
        Err(_) => (false, "PANIC".to_owned()),
    }
}

fn probe(name: &str, bytes: &[u8]) {
    let workbook = match Workbook::open(bytes) {
        Ok(workbook) => workbook,
        Err(error) => {
            println!("F7 {name} open=error {error}");
            return;
        }
    };
    let results = [
        ("can_undo", outcome(|| workbook.can_undo())),
        ("can_redo", outcome(|| workbook.can_redo())),
        (
            "history_undo_depth",
            outcome(|| workbook.history_state().undo_depth),
        ),
        ("client_id", outcome(|| workbook.client_id())),
        (
            "state_update_bytes",
            outcome(|| workbook.encode_state_as_update_v1().len()),
        ),
    ];
    let line: Vec<String> = results
        .iter()
        .map(|(label, (_, value))| format!("{label}={value}"))
        .collect();
    println!("F7 {name} open=ok {}", line.join(" "));
    assert!(
        results.iter().all(|(_, (ok, _))| *ok),
        "{name}: an accessor panicked"
    );
}

#[test]
fn a_non_finite_font_size_starts_collaboration_without_a_panic() {
    probe(
        "non-finite-font-size",
        &package(
            r#"<row r="1"><c r="A1"><v>1</v></c></row>"#,
            "",
            None,
            &styles("INF"),
        ),
    );
}

#[test]
fn many_long_hyperlinks_start_collaboration_without_a_panic() {
    let links: u32 = std::env::var("F7_LINKS")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(60_000);
    let url_len: usize = std::env::var("F7_URL")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(300);
    println!("F7 hyperlinks links={links} url_len={url_len}");
    probe(
        "many-long-hyperlinks",
        &long_hyperlink_package(links, url_len),
    );
}

fn external_hyperlink(row: u32, url_len: usize) -> Hyperlink {
    let at = CellRef::new(row - 1, 0);
    let mut url = format!("https://example.com/catalog/item/{row}?ref=");
    url.extend(std::iter::repeat_n('x', url_len.saturating_sub(url.len())));
    Hyperlink {
        range: CellRange::new(at, at),
        external_target: Some(url),
        location: None,
        tooltip: None,
        display: None,
    }
}

fn long_hyperlink_package(count: u32, url_len: usize) -> Vec<u8> {
    let links: Vec<_> = (1..=count)
        .map(|row| external_hyperlink(row, url_len))
        .collect();
    hyperlink_package(&links)
}

fn hyperlink_package(links: &[Hyperlink]) -> Vec<u8> {
    let mut rows = String::new();
    let mut hyperlinks = String::from("<hyperlinks>");
    let mut rels = format!(r#"<Relationships xmlns="{PKG}">"#);
    for link in links {
        let row = link.range.start.row + 1;
        let url = link.external_target.as_deref().unwrap();
        write!(
            rows,
            r#"<row r="{row}"><c r="A{row}" t="inlineStr"><is><t>item {row}</t></is></c></row>"#
        )
        .unwrap();
        write!(hyperlinks, r#"<hyperlink ref="A{row}" r:id="rIdL{row}"/>"#).unwrap();
        write!(rels, r#"<Relationship Id="rIdL{row}" Type="{HYPERLINK}" Target="{url}" TargetMode="External"/>"#).unwrap();
    }
    hyperlinks.push_str("</hyperlinks>");
    rels.push_str("</Relationships>");
    package(&rows, &hyperlinks, Some(rels), &styles("11"))
}

fn parse(bytes: &[u8]) -> WorkbookModel {
    xlsx_parse::parse_workbook(&ooxml_opc::unzip_parts(bytes).unwrap()).unwrap()
}

fn keep_for_comparison(bytes: &[u8], expected_error: Option<&str>) {
    static KEPT: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);
    if let Ok(dir) = std::env::var("F7_OUT") {
        let index = KEPT.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
        let side = if expected_error.is_some() {
            "refuse"
        } else {
            "accept"
        };
        std::fs::write(format!("{dir}/{index:03}-{side}.xlsx"), bytes).unwrap();
    }
}

#[test]
fn keep_size_files_for_comparison() {
    if std::env::var("F7_OUT").is_err() {
        return;
    }
    let dir = std::env::var("F7_OUT").unwrap();
    std::fs::write(
        format!("{dir}/size-30k-accept.xlsx"),
        long_hyperlink_package(30_000, 300),
    )
    .unwrap();
    std::fs::write(
        format!("{dir}/size-50k-refuse.xlsx"),
        long_hyperlink_package(50_000, 300),
    )
    .unwrap();
}

fn assert_open_matches_build(bytes: &[u8], expected_error: Option<&str>) {
    keep_for_comparison(bytes, expected_error);
    let standalone = Workbook::open(bytes).err().map(|error| error.to_string());
    let collaborative = Workbook::open_collaborative(bytes, 1)
        .err()
        .map(|error| error.to_string());
    assert_eq!(standalone.as_deref(), expected_error);
    assert_eq!(standalone, collaborative);
}

#[test]
fn probe_0_5_1_authority_build_failure() {
    let bytes = long_hyperlink_package(50_000, 300);
    match Workbook::open(&bytes) {
        Ok(workbook) => {
            let (_, client_id) = outcome(|| workbook.client_id());
            let (_, update_bytes) = outcome(|| workbook.encode_state_as_update_v1().len());
            println!("standalone open=ok client_id={client_id} state_update_bytes={update_bytes}");
        }
        Err(error) => println!("standalone open=error {error}"),
    }
    match Workbook::open_collaborative(&bytes, 1) {
        Ok(_) => println!("collaborative open=ok"),
        Err(error) => println!("collaborative open=error {error}"),
    }
}

#[test]
fn standalone_open_refuses_fifty_thousand_long_hyperlinks() {
    let bytes = long_hyperlink_package(50_000, 300);
    let error = Workbook::open(&bytes).err().unwrap();
    assert_eq!(error.to_string(), HYPERLINK_ERROR);
}

#[test]
fn hyperlink_state_boundary_matches_authority_build() {
    let mut links = Vec::new();
    let mut encoded_len = serde_json::to_string(&links).unwrap().len();
    for row in 1.. {
        let link = external_hyperlink(row, 16_000);
        let entry_len = serde_json::to_string(&link).unwrap().len();
        let next_len = encoded_len + entry_len + usize::from(!links.is_empty());
        if next_len >= HYPERLINK_LIMIT {
            break;
        }
        links.push(link);
        encoded_len = next_len;
    }
    for (size, error) in [
        (HYPERLINK_LIMIT - 1, None),
        (HYPERLINK_LIMIT + 1, Some(HYPERLINK_ERROR)),
    ] {
        links
            .last_mut()
            .unwrap()
            .external_target
            .as_mut()
            .unwrap()
            .push_str(&"x".repeat(size - encoded_len));
        encoded_len = size;
        let bytes = hyperlink_package(&links);
        let parsed_len = serde_json::to_string(&parse(&bytes).sheets[0].hyperlinks)
            .unwrap()
            .len();
        assert_eq!(parsed_len, size);
        assert_open_matches_build(&bytes, error);
    }
}

fn chart_package(refs: usize, last_padding: usize) -> Vec<u8> {
    let sheet_rels = format!(
        r#"<Relationships xmlns="{PKG}"><Relationship Id="rIdDrawing" Type="{RELS}/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#
    );
    let mut parts = package_parts(
        "",
        r#"<drawing r:id="rIdDrawing"/>"#,
        Some(sheet_rels),
        &styles("11"),
    );
    parts[0].1 = parts[0].1.replace(
        "</Types>",
        concat!(
            r#"<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>"#,
            r#"<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>"#,
            "</Types>",
        ),
    );
    let drawing = format!(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:c="{CHART}" xmlns:r="{RELS}"><xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="100" cy="100"/><xdr:graphicFrame><c:chart r:id="rIdChart"/></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>"#
    );
    let drawing_rels = format!(
        r#"<Relationships xmlns="{PKG}"><Relationship Id="rIdChart" Type="{RELS}/chart" Target="../charts/chart1.xml"/></Relationships>"#
    );
    let mut chart = format!(r#"<c:chartSpace xmlns:c="{CHART}"><c:chart><c:plotArea><c:barChart>"#);
    let formula_padding = " ".repeat(16_000);
    let last_padding = " ".repeat(last_padding);
    for index in 0..refs {
        let extra = if index + 1 == refs {
            last_padding.as_str()
        } else {
            ""
        };
        write!(
            chart,
            r#"<c:ser><c:idx val="{index}"/><c:order val="{index}"/><c:val><c:numRef><c:f>Sheet1!$A$1{formula_padding}{extra}+0</c:f></c:numRef></c:val></c:ser>"#
        )
        .unwrap();
    }
    chart.push_str("</c:barChart></c:plotArea></c:chart></c:chartSpace>");
    parts.extend([
        ("xl/drawings/drawing1.xml".to_owned(), drawing),
        (
            "xl/drawings/_rels/drawing1.xml.rels".to_owned(),
            drawing_rels,
        ),
        ("xl/charts/chart1.xml".to_owned(), chart),
    ]);
    zip_parts(parts)
}

#[test]
fn chart_state_boundary_matches_authority_build() {
    let sample = parse(&chart_package(1, 0));
    let charts = &sample.sheets[0].charts;
    let one_ref = serde_json::to_string(charts).unwrap().len();
    let extra_ref = serde_json::to_string(&charts[0].refs[0]).unwrap().len() + 1;
    let more_refs = (CHART_LIMIT - 1 - one_ref) / extra_ref;
    let encoded_len = one_ref + more_refs * extra_ref;
    for (size, error) in [
        (CHART_LIMIT - 1, None),
        (CHART_LIMIT + 1, Some(CHART_ERROR)),
    ] {
        let bytes = chart_package(1 + more_refs, size - encoded_len);
        let parsed_len = serde_json::to_string(&parse(&bytes).sheets[0].charts)
            .unwrap()
            .len();
        assert_eq!(parsed_len, size);
        assert_open_matches_build(&bytes, error);
    }
}

fn style_package(styles: &str) -> Vec<u8> {
    package(
        r#"<row r="1"><c r="A1" s="0"><v>1</v></c></row>"#,
        "",
        None,
        styles,
    )
}

fn number_format_package(code_len: usize) -> Vec<u8> {
    let code = "0".repeat(code_len);
    let stylesheet = styles("11")
        .replace(
            "<fonts",
            &format!(
                r#"<numFmts count="1"><numFmt numFmtId="164" formatCode="{code}"/></numFmts><fonts"#
            ),
        )
        .replace(
            r#"<cellXfs count="1"><xf numFmtId="0""#,
            r#"<cellXfs count="1"><xf applyNumberFormat="1" numFmtId="164""#,
        );
    style_package(&stylesheet)
}

fn style_entry(bytes: &[u8]) -> PreservedCellFormat {
    PreservedCellFormat::from_style(&parse(bytes).styles, Some(0))
}

#[test]
fn cell_format_size_boundary_matches_authority_build() {
    let sample_len = serde_json::to_string(&style_entry(&number_format_package(1)))
        .unwrap()
        .len();
    for size in [STYLE_LIMIT - 1, STYLE_LIMIT + 1] {
        let bytes = number_format_package(size - sample_len + 1);
        let payload = serde_json::to_string(&style_entry(&bytes)).unwrap();
        assert_eq!(payload.len(), size);
        let key = Sha256::digest(payload.as_bytes());
        let oversized = format!(
            "invalid collaborative workbook state: cell format {key:x} exceeds its size limit"
        );
        assert_open_matches_build(&bytes, (size > STYLE_LIMIT).then_some(oversized.as_str()));
    }
}

#[test]
fn non_finite_theme_tint_matches_authority_build() {
    let tinted = styles("11").replace("</font>", r#"<color theme="4" tint="NaN"/></font>"#);
    let bytes = style_package(&tinted);
    let entry = style_entry(&bytes);
    assert!(
        matches!(entry.format.font.color, Some(Color::Theme { tint, .. }) if !tint.is_finite())
    );
    let payload = serde_json::to_string(&entry).unwrap();
    let key = Sha256::digest(payload.as_bytes());
    let error = serde_json::from_str::<PreservedCellFormat>(&payload).unwrap_err();
    let expected =
        format!("invalid collaborative workbook state: invalid cell format {key:x}: {error}");
    assert_open_matches_build(&bytes, Some(&expected));
}

#[test]
fn thirty_thousand_long_hyperlinks_support_edit_undo_and_reopen() {
    let bytes = long_hyperlink_package(30_000, 300);
    let mut workbook = Workbook::open(&bytes).unwrap();
    workbook.client_id();
    let applied = workbook
        .apply_edits(&EditRequest {
            expect_version: workbook.version(),
            source: EditSource::Host,
            history: EditHistory::Separate,
            calculation: CalculationRequest::default(),
            steps: vec![EditStep::new(EditOperation::SetCellInputs {
                target: RangeTarget {
                    sheet_id: "sheet:0".into(),
                    range: RangeAddress::A1 { a1: "B1".into() },
                },
                inputs: vec![vec!["edited".into()]],
            })],
        })
        .unwrap()
        .unwrap();
    assert!(applied.applied);
    let at = CellRef::new(0, 1);
    assert_eq!(
        workbook.model().sheets[0].cell(at).unwrap().value,
        CellValue::Text {
            value: "edited".into()
        }
    );
    workbook.undo(CalculationOptions::default()).unwrap();
    assert!(workbook.model().sheets[0].cell(at).is_none());
    let saved = workbook.save().unwrap();
    let reopened = Workbook::open(&saved).unwrap();
    assert!(reopened.model().sheets[0].cell(at).is_none());
    assert_eq!(reopened.model().sheets[0].hyperlinks.len(), 30_000);
}
