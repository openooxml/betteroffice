use super::*;
use super::edit_tests::{options, r};

fn source_bytes() -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&edit_tests::workbook_bytes()).unwrap();
    fn replace(parts: &mut [(String, Vec<u8>)], path: &str, from: &str, to: &str) {
        let bytes = &mut parts.iter_mut().find(|(name, _)| name == path).unwrap().1;
        *bytes = String::from_utf8(std::mem::take(bytes)).unwrap().replace(from, to).into_bytes();
    }
    replace(&mut parts, "xl/workbook.xml", "</workbook>",
        "<definedNames><definedName name=\"Inputs\">Sheet1!$A$1:$B$1</definedName></definedNames></workbook>");
    replace(&mut parts, "xl/worksheets/sheet1.xml", "<worksheet xmlns=", "<worksheet xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns=");
    replace(&mut parts, "xl/worksheets/sheet1.xml", "</worksheet>",
        "<mergeCells><mergeCell ref=\"G1:H1\"/></mergeCells><hyperlinks><hyperlink ref=\"A2\" r:id=\"h1\"/></hyperlinks><drawing r:id=\"d1\"/><tableParts count=\"1\"><tablePart r:id=\"t1\"/></tableParts></worksheet>");
    replace(&mut parts, "[Content_Types].xml", "</Types>",
        "<Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/xl/charts/chart1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/><Override PartName=\"/xl/tables/table1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/></Types>");
    for (path, xml) in [
        ("xl/worksheets/_rels/sheet1.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="h1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/><Relationship Id="d1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/><Relationship Id="t1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>"#),
        ("xl/drawings/drawing1.xml", r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="c1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#),
        ("xl/drawings/_rels/drawing1.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="c1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>"#),
        ("xl/charts/chart1.xml", r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>Sheet1!$A$1:$B$1</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#),
        ("xl/tables/table1.xml", r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="Sales" ref="A1:B2" headerRowCount="1" totalsRowCount="0"><tableColumns count="2"><tableColumn id="1" name="Quantity"/><tableColumn id="2" name="Cost"/></tableColumns></table>"#),
    ] {
        parts.push((path.into(), xml.as_bytes().to_vec()));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn worker(bytes: &[u8], full: bool) -> Workbook {
    let mut worker = Workbook::open_internal(bytes, true, Some(73)).unwrap();
    worker.set_rand_seed(Some(123));
    worker.set_active_sheet(SheetId(1)).unwrap();
    worker.proposals = ProposalSet::with_id_counter(7);
    worker.rebuild_and_recalculate(options());
    if full { worker.rebuild_and_recalculate(options()); }
    worker
}

fn complete(bytes: &[u8], hydration: PeerHydration, units: usize, early: bool) -> Workbook {
    let mut opener = WorkbookPeerOpener::new(bytes.to_vec(), hydration.client_id);
    if !early {
        while opener.advance(units).unwrap() != OpenerState::NeedsHydration {
            assert!(opener.touched_units() <= units);
        }
    }
    for chunk in hydration.into_chunks() { opener.push_hydration(chunk).unwrap(); }
    loop {
        let state = opener.advance(units).unwrap();
        assert!(opener.touched_units() <= units);
        if state == OpenerState::Ready { break; }
    }
    opener.finish().unwrap()
}

fn assert_equal(actual: &mut Workbook, expected: &mut Workbook) {
    assert_eq!(actual.model, expected.model);
    assert_eq!(actual.authority.encode_state_vector_v1(), expected.authority.encode_state_vector_v1());
    assert_eq!(actual.authority.encode_state_as_update_v1(), expected.authority.encode_state_as_update_v1());
    assert_eq!(actual.authority.materialize().unwrap(), expected.authority.materialize().unwrap());
    assert_eq!(actual.authority.structure().unwrap(), expected.authority.structure().unwrap());
    actual.ensure_graph();
    expected.ensure_graph();
    actual.graph.as_ref().unwrap().assert_matches(expected.graph.as_ref().unwrap());
    assert_eq!(actual.version(), expected.version());
    assert_eq!(actual.last_calculation, expected.last_calculation);
    assert_eq!(actual.calculations_since_open, expected.calculations_since_open);
    assert_eq!(actual.recalculated_since_open, expected.recalculated_since_open);
    assert_eq!(actual.rand_seed, expected.rand_seed);
    assert_eq!(actual.proposals.id_counter(), expected.proposals.id_counter());
    assert_eq!(actual.active_sheet, expected.active_sheet);
    assert_eq!(actual.history_state(), expected.history_state());
    assert_eq!(actual.preserved.origins, expected.preserved.origins);
    assert_eq!(actual.preserved.shared_string_cells, expected.preserved.shared_string_cells);
    assert_eq!(actual.preserved.axes, expected.preserved.axes);
    assert_eq!(actual.opened_anchors, expected.opened_anchors);
    assert_eq!(actual.save().unwrap(), expected.save().unwrap());
}

#[test]
fn sliced_peer_matches_oracle_at_every_budget_and_hydration_arrival() {
    let bytes = source_bytes();
    for full in [false, true] {
        let worker = worker(&bytes, full);
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
                let mut expected = Workbook::open_with_peer_hydration_oracle(&bytes, worker.peer_hydration().unwrap()).unwrap();
                assert_equal(&mut actual, &mut expected);
            }
        }
    }
}

#[test]
fn deleted_cells_match_delta_and_full_oracle() {
    let bytes = source_bytes();
    for full in [false, true] {
        let worker = worker(&bytes, full);
        let mut hydration = worker.peer_hydration().unwrap();
        hydration.cells[0].retain(|(at, _, _, _)| *at != r("A2"));
        hydration.deleted_cells[0].push(r("A2"));
        let json = serde_json::to_vec(&hydration).unwrap();
        let mut actual = complete(&bytes, hydration, 1, true);
        let mut expected = Workbook::open_with_peer_hydration_oracle(&bytes, serde_json::from_slice(&json).unwrap()).unwrap();
        assert!(actual.model.sheets[0].cell(r("A2")).is_none());
        assert_equal(&mut actual, &mut expected);
    }
}

#[test]
fn every_edit_kind_and_history_match_oracle() {
    let bytes = source_bytes();
    let worker = worker(&bytes, false);
    let sheet = SheetId(0);
    let range = CellRange::parse_a1("G1:H1").unwrap();
    let chart = &worker.model.sheets[0].charts[0];
    let to = match chart.anchor {
        ChartAnchor::TwoCell { mut from, mut to, edit_as } => {
            from.row += 1;
            to.row += 1;
            ChartAnchor::TwoCell { from, to, edit_as }
        }
        _ => panic!("expected a two-cell chart"),
    };
    let edits = vec![
        Op::SetCell { sheet, at: r("A1"), cell: CellState { value: CellValue::Number { value: 11.0 }, ..Default::default() } },
        Op::InsertRows { sheet, at: 1, count: 1 },
        Op::DeleteRows { sheet, at: 1, count: 1 },
        Op::InsertCols { sheet, at: 1, count: 1 },
        Op::DeleteCols { sheet, at: 1, count: 1 },
        Op::SetColWidth { sheet, col: 0, width: Some(18.0) },
        Op::SetRowHeight { sheet, row: 0, height: Some(24.0) },
        Op::SetFreezePane { sheet, pane: Some(FreezePane::new(1, 1, r("B2"))) },
        Op::SetHyperlinks { sheet, hyperlinks: Vec::new() },
        Op::RestoreColStyles { sheet, styles: vec![xlsx_model::ColStyle { first: 0, last: 1, xf: 1 }] },
        Op::SetCharts { sheet, charts: Vec::new() },
        Op::SetChartAnchor { sheet, frame: chart.frame_id(), part: chart.part.clone(), from: chart.anchor, to },
        Op::MergeCells { sheet, range: CellRange::parse_a1("G2:H2").unwrap() },
        Op::UnmergeCells { sheet, range },
        Op::PatchRangeStyle { sheet, range, patch: StylePatch { bold: Some(true), ..Default::default() } },
        Op::SetRangeNumberFormat { sheet, range, format: NumberFormatMutation::Percent },
        Op::ApplyRangeFormat { sheet, range, format: CapturedFormat { rows: 1, columns: 1, formats: vec![CellFormat::default()] } },
        Op::AddSheet { index: 3, name: "Added".into() },
        Op::RemoveSheet { index: 2 },
        Op::RenameSheet { sheet: SheetId(2), name: "Renamed".into() },
        Op::SetDefinedNames { defined_names: Vec::new() },
        Op::RestoreSheet { sheet: SheetId(2), name: "Restored".into(), formulas: Vec::new() },
    ];
    for op in edits {
        let mut actual = complete(&bytes, worker.peer_hydration().unwrap(), 1, true);
        let mut expected = Workbook::open_with_peer_hydration_oracle(&bytes, worker.peer_hydration().unwrap()).unwrap();
        actual.mode = WorkbookMode::Standalone;
        expected.mode = WorkbookMode::Standalone;
        assert_eq!(format!("{:?}", actual.apply_ops(vec![op.clone()], options())), format!("{:?}", expected.apply_ops(vec![op], options())));
        assert_equal(&mut actual, &mut expected);
        assert_eq!(format!("{:?}", actual.undo(options())), format!("{:?}", expected.undo(options())));
        assert_equal(&mut actual, &mut expected);
        assert_eq!(format!("{:?}", actual.redo(options())), format!("{:?}", expected.redo(options())));
        assert_equal(&mut actual, &mut expected);
    }
}

#[test]
fn opener_rejects_corruption_mismatch_and_premature_finish() {
    assert!(WorkbookPeerOpener::new(Vec::new(), None).finish().is_err());
    assert!(WorkbookPeerOpener::new(b"invalid zip".to_vec(), None).advance(usize::MAX).is_err());
    let bytes = source_bytes();
    let worker = worker(&bytes, false);
    let mut hydration = worker.peer_hydration().unwrap();
    hydration.cells.pop();
    let mut opener = WorkbookPeerOpener::new(bytes.clone(), Some(73));
    for chunk in hydration.into_chunks() { opener.push_hydration(chunk).unwrap(); }
    assert!(opener.advance(usize::MAX).is_err());
    let mut duplicate = bytes;
    let needle = b"xl/worksheets/sheet2.xml";
    for offset in 0..duplicate.len().saturating_sub(needle.len()) {
        if &duplicate[offset..offset + needle.len()] == needle {
            duplicate[offset + needle.len() - 5] = b'1';
        }
    }
    assert!(WorkbookPeerOpener::new(duplicate, Some(73)).advance(usize::MAX).is_err());
}

#[test]
fn hydration_chunks_bound_cells_arrays_and_calculation_lists() {
    let bytes = source_bytes();
    let worker = worker(&bytes, false);
    let mut hydration = worker.peer_hydration().unwrap();
    let limit = peer_open::PEER_HYDRATION_CHUNK_CELLS;
    hydration.cells[0] = vec![(r("A1"), CellValue::Empty, None, None); limit * 3 + 1];
    hydration.deleted_cells[0] = vec![r("A2"); limit * 3 + 1];
    hydration.arrays[0] = vec![(r("A1"), CellRange::parse_a1("A1:B1").unwrap()); limit * 3 + 1];
    hydration.last_calculation.changed = vec![crate::CellAddress { sheet: SheetId(0), cell: r("A1") }; limit * 3 + 1];
    for chunk in hydration.into_chunks() {
        let size = match chunk {
            PeerHydrationChunk::Cells { cells, .. } => cells.len(),
            PeerHydrationChunk::DeletedCells { cells, .. } => cells.len(),
            PeerHydrationChunk::Arrays { arrays, .. } => arrays.len(),
            PeerHydrationChunk::Changed { cells } | PeerHydrationChunk::CycleCells { cells } | PeerHydrationChunk::LimitedCells { cells } => cells.len(),
            _ => 0,
        };
        assert!(size <= limit);
    }
}
