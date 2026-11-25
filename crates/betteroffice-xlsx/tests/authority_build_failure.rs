use std::fmt::Write;
use std::panic::{AssertUnwindSafe, catch_unwind};

use betteroffice_xlsx::Workbook;

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const HYPERLINK: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

fn package(
    sheet_data: &str,
    sheet_tail: &str,
    sheet_rels: Option<String>,
    styles: &str,
) -> Vec<u8> {
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
    let links: usize = std::env::var("F7_LINKS")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(60_000);
    let url_len: usize = std::env::var("F7_URL")
        .ok()
        .and_then(|v| v.parse().ok())
        .unwrap_or(300);
    let mut rows = String::new();
    let mut hyperlinks = String::from("<hyperlinks>");
    let mut rels = format!(r#"<Relationships xmlns="{PKG}">"#);
    for index in 0..links {
        let row = index + 1;
        let mut url = format!("https://example.com/catalog/item/{row}?ref=");
        while url.len() < url_len {
            url.push('x');
        }
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
    println!("F7 hyperlinks links={links} url_len={url_len}");
    probe(
        "many-long-hyperlinks",
        &package(&rows, &hyperlinks, Some(rels), &styles("11")),
    );
}
