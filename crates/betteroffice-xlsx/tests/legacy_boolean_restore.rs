use betteroffice_xlsx::{CalculationOptions, CellRef, CellValue, Error, SheetId, Workbook};
use xlsx_model::CellProvider;
use yrs::updates::decoder::Decode;
use yrs::{Doc, Map, ReadTxn, Transact, Update};

const EMPTY_BOOLEAN_UPDATE_HEX: &str = "0211bc9a859ce889d00c00280111786c73783a63656c6c2d666f726d61747340353830326436323864633734373638353161653461616163623135663365356238633539646239393431366537353738643832643864376365376562656665360177ab027b22666f6e74223a7b226e616d65223a6e756c6c2c2273697a655f7074223a6e756c6c2c22626f6c64223a66616c73652c226974616c6963223a66616c73652c22756e6465726c696e65223a66616c73652c22737472696b65223a66616c73652c22636f6c6f72223a6e756c6c7d2c2266696c6c223a224e6f6e65222c22626f72646572223a7b226c656674223a6e756c6c2c227269676874223a6e756c6c2c22746f70223a6e756c6c2c22626f74746f6d223a6e756c6c7d2c226e756d626572466f726d6174223a7b226b696e64223a226275696c74696e222c226964223a307d2c22616c69676e6d656e74223a7b2268223a6e756c6c2c2276223a6e756c6c2c22777261705f74657874223a66616c73652c22736872696e6b5f746f5f666974223a66616c73657d7d280104786c73780f6261736546696e6765727072696e7401774a78662d6661636574733a38313739343034653833383134643363396336373335326666303638653235363438323431333435616630376463353733323262313439396636323039613266280104786c73780d736368656d6156657273696f6e017d06280104786c73781373747275637475726547656e65726174696f6e017d00080110786c73783a73686565742d6f7264657201770773686565743a3027010b786c73783a7368656574730773686565743a30012700bc9a859ce889d00c0509636f6c576964746873012700bc9a859ce889d00c0508636f6e74656e7473012800bc9a859ce889d00c050a667265657a6550616e65017e2800bc9a859ce889d00c050a68797065726c696e6b730177025b5d2800bc9a859ce889d00c05066368617274730177025b5d2800bc9a859ce889d00c05066d65726765730175002800bc9a859ce889d00c05046e616d650177065368656574312700bc9a859ce889d00c050a726f7748656967687473012700bc9a859ce889d00c05067374796c6573012800bc9a859ce889d00c0703303a300175027a000000000000000075027a0000000000000003792800bc9a859ce889d00c0703303a310175027a000000000000000075027a00000000000000017d01010700a8bc9a859ce889d00c100175027a000000000000000075027a00000000000000017d0201bc9a859ce889d00c011001";
const EMPTY_BOOLEAN_FINGERPRINT: &str =
    "xf-facets:8179404e83814d3c9c67352ff068e25648241345af07dc57322b1499f6209a2f";
const FORMULA_EMPTY_BOOLEAN_UPDATE_HEX: &str = "0211fabf9af8fbf1e40e00280111786c73783a63656c6c2d666f726d61747340353830326436323864633734373638353161653461616163623135663365356238633539646239393431366537353738643832643864376365376562656665360177ab027b22666f6e74223a7b226e616d65223a6e756c6c2c2273697a655f7074223a6e756c6c2c22626f6c64223a66616c73652c226974616c6963223a66616c73652c22756e6465726c696e65223a66616c73652c22737472696b65223a66616c73652c22636f6c6f72223a6e756c6c7d2c2266696c6c223a224e6f6e65222c22626f72646572223a7b226c656674223a6e756c6c2c227269676874223a6e756c6c2c22746f70223a6e756c6c2c22626f74746f6d223a6e756c6c7d2c226e756d626572466f726d6174223a7b226b696e64223a226275696c74696e222c226964223a307d2c22616c69676e6d656e74223a7b2268223a6e756c6c2c2276223a6e756c6c2c22777261705f74657874223a66616c73652c22736872696e6b5f746f5f666974223a66616c73657d7d280104786c73780f6261736546696e6765727072696e7401774a78662d6661636574733a30353564393338666266303639666661353637613261323738353661646265356434353062346334306230313730396337316461336133326631616265306132280104786c73780d736368656d6156657273696f6e017d06280104786c73781373747275637475726547656e65726174696f6e017d00080110786c73783a73686565742d6f7264657201770773686565743a3027010b786c73783a7368656574730773686565743a30012700fabf9af8fbf1e40e0509636f6c576964746873012700fabf9af8fbf1e40e0508636f6e74656e7473012800fabf9af8fbf1e40e050a667265657a6550616e65017e2800fabf9af8fbf1e40e050a68797065726c696e6b730177025b5d2800fabf9af8fbf1e40e05066368617274730177025b5d2800fabf9af8fbf1e40e05066d65726765730175002800fabf9af8fbf1e40e05046e616d650177065368656574312700fabf9af8fbf1e40e050a726f7748656967687473012700fabf9af8fbf1e40e05067374796c6573012800fabf9af8fbf1e40e0703303a300175037a0000000000000001770654525545282975027a0000000000000003792800fabf9af8fbf1e40e0703303a310175027a000000000000000075027a00000000000000017d01010700a8fabf9af8fbf1e40e100175027a000000000000000075027a00000000000000017d0201fabf9af8fbf1e40e011001";
const FORMULA_EMPTY_BOOLEAN_FINGERPRINT: &str =
    "xf-facets:055d938fbf069ffa567a2a27856adbe5d450b4c40b01709c71da3a32f1abe0a2";

const EMPTY_BOOLEAN: &str = r#"<c r="A1" t="b"><v></v></c>"#;
const FORMULA_EMPTY_BOOLEAN: &str = r#"<c r="A1" t="b"><f>TRUE()</f><v/></c>"#;
const TRUE_BOOLEAN: &str = r#"<c r="A1" t="b"><v>1</v></c>"#;
const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
fn package(a1: &str, extra_cells: &str, calc_pr: &str) -> Vec<u8> {
    let parts = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            format!(r#"<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#),
        ),
        (
            "xl/workbook.xml",
            format!(r#"<workbook xmlns="{MAIN}" xmlns:r="{RELS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>{calc_pr}</workbook>"#),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!(r#"<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{RELS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!(r#"<worksheet xmlns="{MAIN}"><sheetData><row r="1">{a1}<c r="B1"><v>1</v></c>{extra_cells}</row></sheetData></worksheet>"#),
        ),
    ];
    let parts: Vec<_> = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn decode_hex(hex: &str) -> Vec<u8> {
    assert!(hex.len().is_multiple_of(2), "expected a pinned hex update");
    (0..hex.len())
        .step_by(2)
        .map(|at| u8::from_str_radix(&hex[at..at + 2], 16).expect("expected a pinned hex update"))
        .collect()
}

fn base_fingerprint(state: &[u8]) -> String {
    let doc = Doc::new();
    doc.transact_mut()
        .apply_update(Update::decode_v1(state).unwrap())
        .unwrap();
    let txn = doc.transact();
    txn.get_map("xlsx")
        .and_then(|meta| meta.get(&txn, "baseFingerprint"))
        .and_then(|value| value.cast::<String>().ok())
        .unwrap()
}

fn apply_and_check_b1(workbook: &mut Workbook, update: &[u8]) {
    assert!(
        workbook
            .apply_update_v1(update, CalculationOptions::default())
            .unwrap()
            .applied
    );
    assert_eq!(
        workbook.model().value(SheetId(0), CellRef::new(0, 1)),
        CellValue::Number { value: 2.0 }
    );
}

fn assert_base_mismatch(workbook: &mut Workbook, update: &[u8]) {
    let before = workbook.encode_state_as_update_v1();
    let error = workbook
        .apply_update_v1(update, CalculationOptions::default())
        .unwrap_err();
    assert!(
        matches!(&error, Error::CollaborativeState(message)
            if message == "workbook base fingerprint does not match shared state"),
        "{error}"
    );
    assert_eq!(workbook.encode_state_as_update_v1(), before);
}

fn restore_legacy_update(a1: &str, hex: &str, fingerprint: &str) {
    let update = decode_hex(hex);
    assert_eq!(base_fingerprint(&update), fingerprint);
    let mut reopened = Workbook::open_collaborative(&package(a1, "", ""), 8).unwrap();
    apply_and_check_b1(&mut reopened, &update);
}

#[test]
fn restores_0_5_1_update_with_an_empty_boolean() {
    restore_legacy_update(
        EMPTY_BOOLEAN,
        EMPTY_BOOLEAN_UPDATE_HEX,
        EMPTY_BOOLEAN_FINGERPRINT,
    );
}

#[test]
fn restores_0_5_1_update_with_an_empty_boolean_formula_cache() {
    restore_legacy_update(
        FORMULA_EMPTY_BOOLEAN,
        FORMULA_EMPTY_BOOLEAN_UPDATE_HEX,
        FORMULA_EMPTY_BOOLEAN_FINGERPRINT,
    );
}

#[test]
fn rejects_0_5_1_update_from_a_different_base() {
    let update = decode_hex(EMPTY_BOOLEAN_UPDATE_HEX);
    for source in [
        package(TRUE_BOOLEAN, "", ""),
        package(EMPTY_BOOLEAN, r#"<c r="C1"><v>5</v></c>"#, ""),
    ] {
        let mut reopened = Workbook::open_collaborative(&source, 8).unwrap();
        assert_base_mismatch(&mut reopened, &update);
    }
}

#[test]
fn rejects_0_5_1_update_for_an_empty_boolean_in_another_cell() {
    let update = decode_hex(EMPTY_BOOLEAN_UPDATE_HEX);
    let source = package("", r#"<c r="D1" t="b"><v></v></c>"#, "");
    let mut reopened = Workbook::open_collaborative(&source, 8).unwrap();
    assert_base_mismatch(&mut reopened, &update);
}
