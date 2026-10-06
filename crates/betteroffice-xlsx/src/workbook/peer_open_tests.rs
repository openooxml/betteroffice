use std::fmt::Write;
use std::future::Future;
use std::task::{Context, Poll, Waker};

use super::edit_tests::{options, r};
use super::*;

fn sliced_authority(model: &WorkbookModel, units: usize) -> (WorkbookAuthority, usize) {
    let work = ooxml_opc::WorkBudget::default();
    let mut opening = Box::pin(WorkbookAuthority::from_source_with_projection_sliced(
        model,
        Some(73),
        &[],
        None,
        &work,
    ));
    let mut context = Context::from_waker(Waker::noop());
    let mut charged = 0;
    loop {
        work.reset(units);
        let state = opening.as_mut().poll(&mut context);
        assert!(work.touched() <= units);
        charged += work.touched();
        if let Poll::Ready(result) = state {
            return (result.unwrap().0, charged);
        }
    }
}

#[test]
fn indexed_palette_fingerprints_stay_sliced() {
    let mut parts = ooxml_opc::unzip_parts(&source_bytes(24)).unwrap();
    let styles = &mut parts
        .iter_mut()
        .find(|(path, _)| path == "xl/styles.xml")
        .unwrap()
        .1;
    *styles = String::from_utf8(std::mem::take(styles))
        .unwrap()
        .replace(
            "</styleSheet>",
            &format!(
                "<colors><indexedColors>{}</indexedColors></colors></styleSheet>",
                "<rgbColor rgb=\"FF123456\"/>".repeat(66),
            ),
        )
        .into_bytes();
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    let model = xlsx_parse::parse_workbook(&parts).unwrap();
    assert!(!model.styles.indexed_colors.is_empty());
    let (expected, _, _) =
        WorkbookAuthority::from_source_with_projection(&model, Some(73), &[], None).unwrap();
    for units in [1, 256, usize::MAX] {
        crate::authority::SYNC_FINGERPRINTS.set(0);
        let (actual, _) = sliced_authority(&model, units);
        assert_eq!(crate::authority::SYNC_FINGERPRINTS.get(), 0);
        assert_eq!(
            actual.encode_state_as_update_v1(),
            expected.encode_state_as_update_v1()
        );
        assert_eq!(
            actual.encode_state_vector_v1(),
            expected.encode_state_vector_v1()
        );
        let worker = worker(&bytes, false, Some(73));
        let mut actual = complete(&bytes, worker.peer_hydration().unwrap(), units, true);
        let mut expected =
            Workbook::open_with_peer_hydration_oracle(&bytes, worker.peer_hydration().unwrap())
                .unwrap();
        assert_equal(&mut actual, &mut expected);
    }
}

#[test]
fn seed_batches_charge_long_values_and_sheet_metadata() {
    let mut model = WorkbookModel::default();
    let mut sheet = Sheet::new("Large");
    for row in 0..40 {
        sheet.set_cell(
            CellRef::new(row, 0),
            xlsx_model::Cell {
                value: CellValue::Text {
                    value: "x".repeat(8192),
                },
                formula: None,
                style: None,
            },
        );
    }
    for row in 0..8 {
        let at = CellRef::new(row, 0);
        sheet.hyperlinks.push(xlsx_model::Hyperlink {
            range: CellRange { start: at, end: at },
            external_target: Some("https://example.com".into()),
            location: None,
            tooltip: Some("m".repeat(16 * 1024)),
            display: None,
        });
    }
    model.sheets.push(sheet);
    let (expected, _, _) =
        WorkbookAuthority::from_source_with_projection(&model, Some(73), &[], None).unwrap();
    for units in [1, 256, usize::MAX] {
        crate::authority::take_seed_batches();
        let (actual, charged) = sliced_authority(&model, units);
        let batches = crate::authority::take_seed_batches();
        assert!(batches.len() > 4);
        assert!(batches.iter().any(|(bytes, _)| *bytes > 64 * 1024));
        for (bytes, count) in &batches {
            assert!(*bytes <= 64 * 1024 || *count == 1);
        }
        assert!(
            charged
                >= batches
                    .iter()
                    .map(|(bytes, _)| bytes.div_ceil(64))
                    .sum::<usize>()
        );
        assert_eq!(
            actual.encode_state_as_update_v1(),
            expected.encode_state_as_update_v1()
        );
        assert_eq!(
            actual.encode_state_vector_v1(),
            expected.encode_state_vector_v1()
        );
    }
}

#[test]
fn seed_sheet_order_spans_byte_bounded_batches_and_matches_oracle() {
    let mut model = WorkbookModel::default();
    for index in 0..8192 {
        let name = format!("Sheet-{index:04}-{}", "n".repeat(20));
        model.sheets.push(Sheet::new(name));
    }
    let order_bytes = (0..model.sheets.len())
        .map(|index| format!("sheet:{index}").len() + 2)
        .sum::<usize>();
    assert!(order_bytes > 64 * 1024);
    let (expected, _, _) =
        WorkbookAuthority::from_source_with_projection(&model, Some(73), &[], None).unwrap();
    for units in [1, 256, usize::MAX] {
        crate::authority::take_seed_batches();
        let (actual, _) = sliced_authority(&model, units);
        let batches = crate::authority::take_seed_batches();
        assert!(batches.iter().all(|(bytes, _)| *bytes <= 64 * 1024));
        let entries = batches.iter().map(|(_, count)| count).sum::<usize>();
        let order_entries = entries - (4 + 10 * model.sheets.len());
        assert!(order_entries > 1);
        assert_eq!(
            actual.encode_state_vector_v1(),
            expected.encode_state_vector_v1()
        );
        assert_eq!(
            actual.encode_state_as_update_v1(),
            expected.encode_state_as_update_v1()
        );
    }
}

#[test]
fn hydration_chunks_bound_bytes_and_split_oversized_cells() {
    let mut parts = ooxml_opc::unzip_parts(&source_bytes(0)).unwrap();
    let sheet = &mut parts
        .iter_mut()
        .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
        .unwrap()
        .1;
    let text: String = "l🙂🙂\""
        .chars()
        .cycle()
        .take(xlsx_calc::eval::MAX_CELL_TEXT_CHARS)
        .collect();
    *sheet = String::from_utf8(std::mem::take(sheet)).unwrap().replace(
        "</sheetData>",
        &format!("<row r=\"30\"><c r=\"A30\" t=\"inlineStr\"><is><t>{text}</t></is></c></row></sheetData>"),
    ).into_bytes();
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    let worker = worker(&bytes, true, Some(73));
    let chunks = worker.peer_hydration().unwrap().into_chunks();
    let mut parts = 0;
    for chunk in chunks {
        let record = serde_json::json!({
            "workbook": chunk, "calculation_context": null,
        });
        if record["workbook"]["kind"] == "cell-part" {
            parts += 1;
        }
        let json = serde_json::to_vec(&record).unwrap();
        assert!(json.len() <= 64 * 1024);
    }
    assert!(parts > 1);
    for units in [1, 256, usize::MAX] {
        let mut actual = complete(&bytes, worker.peer_hydration().unwrap(), units, true);
        let mut expected =
            Workbook::open_with_peer_hydration_oracle(&bytes, worker.peer_hydration().unwrap())
                .unwrap();
        assert_equal(&mut actual, &mut expected);
    }
}

fn source_bytes(rows: usize) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&edit_tests::workbook_bytes()).unwrap();
    fn replace(parts: &mut [(String, Vec<u8>)], path: &str, from: &str, to: &str) {
        let bytes = &mut parts.iter_mut().find(|(name, _)| name == path).unwrap().1;
        *bytes = String::from_utf8(std::mem::take(bytes))
            .unwrap()
            .replace(from, to)
            .into_bytes();
    }
    replace(
        &mut parts,
        "xl/workbook.xml",
        "<sheets>",
        "<workbookPr date1904=\"1\"/><sheets>",
    );
    let styles = &mut parts
        .iter_mut()
        .find(|(name, _)| name == "xl/styles.xml")
        .unwrap()
        .1;
    let styles_xml = br#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF336699"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FF112233"/></left><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="1"><dxf><font><b/><color rgb="FFCC0000"/></font><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></dxf></dxfs>
</styleSheet>"#;
    *styles = styles_xml.to_vec();
    replace(
        &mut parts,
        "xl/worksheets/sheet1.xml",
        "<sheetData>",
        "<sheetViews><sheetView workbookViewId=\"0\"><pane xSplit=\"1\" ySplit=\"2\" topLeftCell=\"B3\" activePane=\"bottomRight\" state=\"frozen\"/></sheetView></sheetViews><cols><col min=\"1\" max=\"2\" width=\"18.5\" customWidth=\"1\"/></cols><sheetData>",
    );
    replace(
        &mut parts,
        "xl/worksheets/sheet1.xml",
        "<row r=\"1\">",
        "<row r=\"1\" ht=\"30\" customHeight=\"1\">",
    );
    replace(
        &mut parts,
        "xl/worksheets/sheet1.xml",
        "</sheetData>",
        "<row r=\"6\"><c r=\"E6\" t=\"inlineStr\"><is><t>Inline plain</t></is></c><c r=\"F6\" t=\"inlineStr\"><is><r><rPr><i/></rPr><t>Inline rich</t></r></is></c><c r=\"I6\" s=\"2\"><v>123</v></c></row></sheetData>",
    );
    replace(
        &mut parts,
        "_rels/.rels",
        "</Relationships>",
        "<Relationship Id=\"custom1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml\" Target=\"customXml/item1.xml\"/></Relationships>",
    );
    replace(
        &mut parts,
        "xl/workbook.xml",
        "</workbook>",
        "<definedNames><definedName name=\"Inputs\">Sheet1!$A$1:$B$1</definedName></definedNames></workbook>",
    );
    replace(
        &mut parts,
        "xl/worksheets/sheet1.xml",
        "<worksheet xmlns=",
        "<worksheet xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns=",
    );
    replace(
        &mut parts,
        "xl/worksheets/sheet1.xml",
        "</worksheet>",
        "<mergeCells><mergeCell ref=\"G1:H1\"/></mergeCells><hyperlinks><hyperlink ref=\"A2\" r:id=\"h1\"/></hyperlinks><drawing r:id=\"d1\"/><tableParts count=\"1\"><tablePart r:id=\"t1\"/></tableParts></worksheet>",
    );
    replace(
        &mut parts,
        "[Content_Types].xml",
        "</Types>",
        "<Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/xl/charts/chart1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/><Override PartName=\"/xl/tables/table1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/></Types>",
    );
    for (path, xml) in [
        (
            "customXml/item1.xml",
            r#"<coverage xmlns="urn:peer-coverage">preserved source part</coverage>"#,
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="h1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/><Relationship Id="d1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>"#,
        ),
        (
            "xl/drawings/drawing1.xml",
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="c1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#,
        ),
        (
            "xl/drawings/_rels/drawing1.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="c1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#,
        ),
        (
            "xl/charts/chart1.xml",
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Sheet1!$A$1:$B$1</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#,
        ),
        (
            "xl/tables/table1.xml",
            r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B2" headerRowCount="1" totalsRowCount="0"><tableColumns count="2"><tableColumn id="1" name="Quantity"/><tableColumn id="2" name="Cost"/></tableColumns></table>"#,
        ),
    ] {
        parts.push((path.into(), xml.as_bytes().to_vec()));
    }
    if rows > 0 {
        for sheet in 1..=3 {
            let mut data = String::from("<sheetData>");
            let other = sheet % 3 + 1;
            for row in 1..=rows {
                write!(
                    data,
                    "<row r=\"{row}\" ht=\"27\" customHeight=\"1\"><c r=\"A{row}\" s=\"2\"><v>{row}</v></c><c r=\"B{row}\" t=\"s\" s=\"1\"><v>0</v></c><c r=\"C{row}\"><f t=\"array\" ref=\"C{row}:C{row}\">A{row}*2</f><v>0</v></c><c r=\"D{row}\"><f>C{row}+Sheet{other}!A{row}</f><v>0</v></c></row>"
                )
                .unwrap();
            }
            data.push_str("</sheetData>");
            let path = format!("xl/worksheets/sheet{sheet}.xml");
            let bytes = &mut parts.iter_mut().find(|(name, _)| name == &path).unwrap().1;
            let mut xml = String::from_utf8(std::mem::take(bytes)).unwrap();
            let start = xml.find("<sheetData>").unwrap();
            let end = xml.find("</sheetData>").unwrap() + "</sheetData>".len();
            xml.replace_range(start..end, &data);
            *bytes = xml.into_bytes();
        }
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn worker(bytes: &[u8], full: bool, client_id: Option<u64>) -> Workbook {
    let mut worker = Workbook::open_internal(bytes, true, client_id).unwrap();
    worker.set_rand_seed(Some(123));
    worker.set_active_sheet(SheetId(1)).unwrap();
    worker.proposals = ProposalSet::with_id_counter(7);
    worker.rebuild_and_recalculate(options());
    if full {
        worker.rebuild_and_recalculate(options());
    }
    worker
}

fn complete(bytes: &[u8], hydration: PeerHydration, units: usize, early: bool) -> Workbook {
    let mut opener = WorkbookPeerOpener::new(bytes.to_vec(), hydration.client_id);
    if !early {
        while opener.advance(units).unwrap() != OpenerState::NeedsHydration {
            assert!(opener.touched_units() <= units);
        }
    }
    for chunk in hydration.into_chunks() {
        opener.push_hydration(chunk).unwrap();
    }
    loop {
        let state = opener.advance(units).unwrap();
        assert!(opener.touched_units() <= units);
        if state == OpenerState::Ready {
            break;
        }
    }
    opener.finish().unwrap()
}

fn assert_equal(actual: &mut Workbook, expected: &mut Workbook) {
    assert_eq!(actual.is_collaborative(), expected.is_collaborative());
    if actual.is_collaborative() {
        assert_eq!(actual.client_id(), expected.client_id());
    }
    assert_eq!(actual.model, expected.model);
    assert_eq!(
        actual.authority.encode_state_vector_v1(),
        expected.authority.encode_state_vector_v1()
    );
    assert_eq!(
        actual.authority.encode_state_as_update_v1(),
        expected.authority.encode_state_as_update_v1()
    );
    assert_eq!(
        actual.authority.materialize().unwrap(),
        expected.authority.materialize().unwrap()
    );
    assert_eq!(
        actual.authority.structure().unwrap(),
        expected.authority.structure().unwrap()
    );
    actual.ensure_graph();
    expected.ensure_graph();
    actual
        .graph
        .as_ref()
        .unwrap()
        .assert_matches(expected.graph.as_ref().unwrap());
    assert_eq!(actual.version(), expected.version());
    assert_eq!(actual.last_calculation, expected.last_calculation);
    assert_eq!(
        actual.calculations_since_open,
        expected.calculations_since_open
    );
    assert_eq!(
        actual.recalculated_since_open,
        expected.recalculated_since_open
    );
    assert_eq!(actual.rand_seed, expected.rand_seed);
    assert_eq!(
        actual.proposals.id_counter(),
        expected.proposals.id_counter()
    );
    assert_eq!(actual.active_sheet, expected.active_sheet);
    assert_eq!(actual.history_state(), expected.history_state());
    assert_eq!(actual.preserved.origins, expected.preserved.origins);
    assert_eq!(
        actual.preserved.shared_string_cells,
        expected.preserved.shared_string_cells
    );
    assert_eq!(actual.preserved.axes, expected.preserved.axes);
    assert_eq!(actual.opened_anchors, expected.opened_anchors);
    let saved = actual.save().unwrap();
    assert_eq!(saved, expected.save().unwrap());
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    assert_eq!(
        parts
            .iter()
            .find(|(name, _)| name == "customXml/item1.xml")
            .unwrap()
            .1
            .as_slice(),
        br#"<coverage xmlns="urn:peer-coverage">preserved source part</coverage>"#.as_slice()
    );
}

#[test]
fn sliced_authority_chunks_match_oracle_with_repeated_cell_formats() {
    let mut model = WorkbookModel::default();
    model.styles.fonts.push(xlsx_model::Font {
        bold: true,
        ..Default::default()
    });
    let format = xlsx_model::Xf {
        font: Some(0),
        ..Default::default()
    };
    model.styles.cell_xfs = vec![xlsx_model::Xf::default(), format.clone(), format];
    let mut sheet = Sheet::new("Chunks");
    for row in 0..4097 {
        sheet.set_cell(
            CellRef::new(row, 0),
            xlsx_model::Cell {
                value: CellValue::Text {
                    value: format!("value {row}"),
                },
                formula: (row % 3 == 0).then(|| format!("\"value {row}\"")),
                style: Some(1 + row % 2),
            },
        );
    }
    sheet.col_widths = BTreeMap::from([(2, 12.5), (10, 18.0)]);
    sheet.row_heights = BTreeMap::from([(2, 24.0), (10, 30.5)]);
    model.sheets = vec![sheet, Sheet::new("Empty")];
    let (expected, _, _) =
        WorkbookAuthority::from_source_with_projection(&model, Some(73), &[], None).unwrap();
    for units in [1, 256] {
        let work = ooxml_opc::WorkBudget::default();
        let mut opening = Box::pin(WorkbookAuthority::from_source_with_projection_sliced(
            &model,
            Some(73),
            &[],
            None,
            &work,
        ));
        let mut context = Context::from_waker(Waker::noop());
        let actual = loop {
            work.reset(units);
            let state = opening.as_mut().poll(&mut context);
            assert!(work.touched() <= units);
            if let Poll::Ready(result) = state {
                break result.unwrap().0;
            }
        };
        assert_eq!(
            actual.encode_state_as_update_v1(),
            expected.encode_state_as_update_v1()
        );
        assert_eq!(
            actual.encode_state_vector_v1(),
            expected.encode_state_vector_v1()
        );
    }
}

#[test]
fn sliced_peer_matches_oracle_at_every_budget_and_hydration_arrival() {
    let bytes = source_bytes(0);
    for full in [false, true] {
        let worker = worker(&bytes, full, Some(73));
        assert_eq!(worker.model.sheets.len(), 3);
        assert_eq!(worker.model.sheets[0].charts.len(), 1);
        assert_eq!(worker.model.tables.len(), 1);
        assert_eq!(worker.model.defined_names.len(), 1);
        assert_eq!(worker.model.sheets[0].hyperlinks.len(), 1);
        assert!(!worker.model.sheets[0].merges.is_empty());
        assert!(worker.model.sheets[0].array_formulas().next().is_some());
        assert_eq!(worker.peer_hydration().unwrap().delta, !full);
        for units in [1, 256, usize::MAX] {
            for early in [false, true] {
                let mut actual = complete(&bytes, worker.peer_hydration().unwrap(), units, early);
                let mut expected = Workbook::open_with_peer_hydration_oracle(
                    &bytes,
                    worker.peer_hydration().unwrap(),
                )
                .unwrap();
                assert_equal(&mut actual, &mut expected);
            }
        }
    }
}

#[test]
fn standalone_peer_matches_oracle_at_every_budget_and_hydration_arrival() {
    let bytes = source_bytes(0);
    for full in [false, true] {
        let worker = worker(&bytes, full, None);
        assert!(!worker.is_collaborative());
        assert_eq!(worker.peer_hydration().unwrap().client_id, None);
        for units in [1, 256, 4096] {
            for early in [false, true] {
                let mut actual = complete(&bytes, worker.peer_hydration().unwrap(), units, early);
                let mut expected = Workbook::open_with_peer_hydration_oracle(
                    &bytes,
                    worker.peer_hydration().unwrap(),
                )
                .unwrap();
                assert_equal(&mut actual, &mut expected);
            }
        }
    }
}

#[test]
fn multi_chunk_hydration_matches_delta_and_full_oracle() {
    let bytes = source_bytes(1200);
    for full in [false, true] {
        let worker = worker(&bytes, full, Some(73));
        let mut hydration = worker.peer_hydration().unwrap();
        assert_eq!(hydration.delta, !full);
        for sheet in 0..3 {
            hydration.cells[sheet].retain(|(at, _, _, _)| at.col != 1);
            hydration.deleted_cells[sheet] = (0..1200).map(|row| CellRef::new(row, 1)).collect();
        }
        let json = serde_json::to_vec(&hydration).unwrap();
        let chunks = hydration.into_chunks();
        let mut counts = [0; 3];
        for chunk in &chunks {
            match chunk {
                PeerHydrationChunk::Cells { sheet: 0, cells } => {
                    assert!(cells.len() <= 512);
                    counts[0] += 1;
                }
                PeerHydrationChunk::DeletedCells { sheet: 0, cells } => {
                    assert!(cells.len() <= 512);
                    counts[1] += 1;
                }
                PeerHydrationChunk::Arrays { sheet: 0, arrays } => {
                    assert!(arrays.len() <= 512);
                    counts[2] += 1;
                }
                _ => {}
            }
        }
        assert!(counts.into_iter().all(|count| count >= 3));
        for units in [256, 4096] {
            for early in [false, true] {
                let hydration = serde_json::from_slice(&json).unwrap();
                let mut actual = complete(&bytes, hydration, units, early);
                let mut expected = Workbook::open_with_peer_hydration_oracle(
                    &bytes,
                    serde_json::from_slice(&json).unwrap(),
                )
                .unwrap();
                for sheet in &actual.model.sheets {
                    assert!(sheet.cell(r("B1200")).is_none());
                    assert_eq!(
                        sheet.cell(r("D1200")).unwrap().value,
                        CellValue::Number { value: 3600.0 }
                    );
                    assert!(sheet.array_formula(r("C1200")).is_some());
                }
                assert_equal(&mut actual, &mut expected);
            }
        }
    }
}

#[test]
fn deleted_cells_match_delta_and_full_oracle() {
    let bytes = source_bytes(0);
    for full in [false, true] {
        let worker = worker(&bytes, full, Some(73));
        let mut hydration = worker.peer_hydration().unwrap();
        hydration.cells[0].retain(|(at, _, _, _)| *at != r("A2"));
        hydration.deleted_cells[0].push(r("A2"));
        let json = serde_json::to_vec(&hydration).unwrap();
        let mut actual = complete(&bytes, hydration, 1, true);
        let mut expected = Workbook::open_with_peer_hydration_oracle(
            &bytes,
            serde_json::from_slice(&json).unwrap(),
        )
        .unwrap();
        assert!(actual.model.sheets[0].cell(r("A2")).is_none());
        assert_equal(&mut actual, &mut expected);
    }
}

#[test]
fn every_edit_kind_and_history_match_oracle() {
    let bytes = source_bytes(0);
    let worker = worker(&bytes, false, Some(73));
    let sheet = SheetId(0);
    let range = CellRange::parse_a1("G1:H1").unwrap();
    let chart = &worker.model.sheets[0].charts[0];
    let to = match chart.anchor {
        ChartAnchor::TwoCell {
            mut from,
            mut to,
            edit_as,
        } => {
            from.row += 1;
            to.row += 1;
            ChartAnchor::TwoCell { from, to, edit_as }
        }
        _ => panic!("expected a two-cell chart"),
    };
    let edits = vec![
        Op::SetCell {
            sheet,
            at: r("A1"),
            cell: CellState {
                value: CellValue::Number { value: 11.0 },
                ..Default::default()
            },
        },
        Op::InsertRows {
            sheet,
            at: 1,
            count: 1,
        },
        Op::DeleteRows {
            sheet,
            at: 1,
            count: 1,
        },
        Op::InsertCols {
            sheet,
            at: 1,
            count: 1,
        },
        Op::DeleteCols {
            sheet,
            at: 1,
            count: 1,
        },
        Op::SetColWidth {
            sheet,
            col: 0,
            width: Some(18.0),
        },
        Op::SetRowHeight {
            sheet,
            row: 0,
            height: Some(24.0),
        },
        Op::SetFreezePane {
            sheet,
            pane: Some(FreezePane::new(1, 1, r("B2"))),
        },
        Op::SetHyperlinks {
            sheet,
            hyperlinks: Vec::new(),
        },
        Op::RestoreColStyles {
            sheet,
            styles: vec![xlsx_model::ColStyle {
                first: 0,
                last: 1,
                xf: 1,
            }],
        },
        Op::SetCharts {
            sheet,
            charts: Vec::new(),
        },
        Op::SetChartAnchor {
            sheet,
            frame: chart.frame_id(),
            part: chart.part.clone(),
            from: chart.anchor,
            to,
        },
        Op::MergeCells {
            sheet,
            range: CellRange::parse_a1("G2:H2").unwrap(),
        },
        Op::UnmergeCells { sheet, range },
        Op::PatchRangeStyle {
            sheet,
            range,
            patch: StylePatch {
                bold: Some(true),
                ..Default::default()
            },
        },
        Op::SetRangeNumberFormat {
            sheet,
            range,
            format: NumberFormatMutation::Percent,
        },
        Op::ApplyRangeFormat {
            sheet,
            range,
            format: CapturedFormat {
                rows: 1,
                columns: 1,
                formats: vec![CellFormat::default()],
            },
        },
        Op::AddSheet {
            index: 3,
            name: "Added".into(),
        },
        Op::RemoveSheet { index: 2 },
        Op::RenameSheet {
            sheet: SheetId(2),
            name: "Renamed".into(),
        },
        Op::SetDefinedNames {
            defined_names: Vec::new(),
        },
        Op::RestoreSheet {
            sheet: SheetId(2),
            name: "Restored".into(),
            formulas: Vec::new(),
        },
    ];
    for op in edits {
        let mut actual = complete(&bytes, worker.peer_hydration().unwrap(), 1, true);
        let mut expected =
            Workbook::open_with_peer_hydration_oracle(&bytes, worker.peer_hydration().unwrap())
                .unwrap();
        actual.mode = WorkbookMode::Standalone;
        expected.mode = WorkbookMode::Standalone;
        assert_eq!(
            format!("{:?}", actual.apply_ops(vec![op.clone()], options())),
            format!("{:?}", expected.apply_ops(vec![op], options()))
        );
        assert_equal(&mut actual, &mut expected);
        assert_eq!(
            format!("{:?}", actual.undo(options())),
            format!("{:?}", expected.undo(options()))
        );
        assert_equal(&mut actual, &mut expected);
        assert_eq!(
            format!("{:?}", actual.redo(options())),
            format!("{:?}", expected.redo(options()))
        );
        assert_equal(&mut actual, &mut expected);
    }
}

#[test]
fn opener_rejects_corruption_mismatch_and_premature_finish() {
    assert!(WorkbookPeerOpener::new(Vec::new(), None).finish().is_err());
    assert!(
        WorkbookPeerOpener::new(b"invalid zip".to_vec(), None)
            .advance(usize::MAX)
            .is_err()
    );
    let bytes = source_bytes(0);
    let worker = worker(&bytes, false, Some(73));
    let mut hydration = worker.peer_hydration().unwrap();
    hydration.cells.pop();
    let mut opener = WorkbookPeerOpener::new(bytes.clone(), Some(73));
    for chunk in hydration.into_chunks() {
        opener.push_hydration(chunk).unwrap();
    }
    assert!(opener.advance(usize::MAX).is_err());
    let mut duplicate = bytes;
    let needle = b"xl/worksheets/sheet2.xml";
    for offset in 0..duplicate.len().saturating_sub(needle.len()) {
        if &duplicate[offset..offset + needle.len()] == needle {
            duplicate[offset + needle.len() - 5] = b'1';
        }
    }
    assert!(
        WorkbookPeerOpener::new(duplicate, Some(73))
            .advance(usize::MAX)
            .is_err()
    );
}

#[test]
fn hydration_chunks_bound_cells_arrays_and_calculation_lists() {
    let bytes = source_bytes(0);
    let worker = worker(&bytes, false, Some(73));
    let mut hydration = worker.peer_hydration().unwrap();
    let limit = peer_open::PEER_HYDRATION_CHUNK_CELLS;
    hydration.cells[0] = vec![(r("A1"), CellValue::Empty, None, None); limit * 3 + 1];
    hydration.deleted_cells[0] = vec![r("A2"); limit * 3 + 1];
    hydration.arrays[0] = vec![(r("A1"), CellRange::parse_a1("A1:B1").unwrap()); limit * 3 + 1];
    hydration.last_calculation.changed = vec![
        crate::CellAddress {
            sheet: SheetId(0),
            cell: r("A1")
        };
        limit * 3 + 1
    ];
    for chunk in hydration.into_chunks() {
        let size = match chunk {
            PeerHydrationChunk::Cells { cells, .. } => cells.len(),
            PeerHydrationChunk::DeletedCells { cells, .. } => cells.len(),
            PeerHydrationChunk::Arrays { arrays, .. } => arrays.len(),
            PeerHydrationChunk::Changed { cells }
            | PeerHydrationChunk::CycleCells { cells }
            | PeerHydrationChunk::LimitedCells { cells } => cells.len(),
            _ => 0,
        };
        assert!(size <= limit);
    }
}
