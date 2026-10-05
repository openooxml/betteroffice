use crate::{
    CalculationRequest, Cell, CellAddress, CellRef, CellValue, EditOperation, EditRequest,
    EditStep, ProposalEditInput, ProposalRequest, RangeAddress, RangeTarget, UpdateEvent, Viewport,
    WorkbookModel, XlsxExportOptions,
};
use xlsx_ops::{CellState, Op};

use super::*;

fn budgets() -> [SnapshotBudget; 3] {
    [
        SnapshotBudget::new(1, 16 * 1024).unwrap(),
        SnapshotBudget::new(7, 16 * 1024).unwrap(),
        SnapshotBudget::new(usize::MAX, usize::MAX).unwrap(),
    ]
}

fn context() -> CalculationOptions {
    CalculationOptions {
        now_serial: Some(45_001.25),
    }
}

fn source(protected: bool, unpatchable: bool) -> Vec<u8> {
    let workbook = br#"<workbook
        xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <bookViews><workbookView activeTab="1"/></bookViews>
        <sheets><sheet name="Data" sheetId="1" r:id="s1"/>
        <sheet name="Other" sheetId="2" r:id="s2"/>
        <sheet name="Empty" sheetId="3" r:id="s3"/></sheets>
        <definedNames><definedName name="Input">Data!$A$1</definedName>
        <definedName name="Local" localSheetId="0" hidden="1">Data!$A$5</definedName>
        </definedNames></workbook>"#;
    let worksheet = br#"<worksheet
        xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheetViews><sheetView workbookViewId="0"><pane ySplit="1"
        topLeftCell="A2" state="frozen"/></sheetView></sheetViews>
        <sheetFormatPr defaultRowHeight="15"/>
        <cols><col min="1" max="1" width="12" customWidth="1" style="0"/></cols>
        <sheetData><row r="1"><c r="A1" s="0"><v>2</v></c>
        <c r="B1" s="1"><f>A1+1</f><v>3</v></c>
        <c r="C1" t="s"><v>0</v></c><c r="D1" t="s"><v>1</v></c>
        <c r="E1"><f>NOW()</f><v>45000</v></c>
        <c r="F1"><f>RAND()</f><v>0.5</v></c>
        <c r="G1"><f t="array" ref="G1:H2">SEQUENCE(2,2)</f><v>1</v></c>
        <c r="H1"><v>2</v></c></row>
        <row r="2"><c r="A2"><v>-0</v></c><c r="B2" s="0"><v>42</v></c>
        <c r="C2" t="inlineStr"><is><t></t></is></c>
        <c r="G2"><v>3</v></c><c r="H2"><v>4</v></c></row>
        <row r="3"><c r="A3"><f>SUM(Items[Qty])+Input</f><v>8</v></c></row>
        <row r="4"><c r="A4" t="inlineStr"><is><t>Qty</t></is></c>
        <c r="B4" t="inlineStr"><is><t>Cost</t></is></c></row>
        <row r="5"><c r="A5"><v>2</v></c><c r="B5"><v>3</v></c></row>
        <row r="6"><c r="A6"><v>4</v></c><c r="B6"><v>5</v></c></row>
        </sheetData><mergeCells count="1"><mergeCell ref="I1:J1"/></mergeCells>
        <drawing r:id="drawing"/><tableParts count="1"><tablePart r:id="table"/>
        </tableParts></worksheet>"#;
    let other = format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData><row r="1"><c r="A1"><f>Data!A1+1</f><v>3</v></c></row></sheetData>
        {}</worksheet>"#,
        if protected {
            r#"<sheetProtection sheet="1"/>"#
        } else {
            ""
        }
    );
    let mut parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels"
            ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Override PartName="/xl/workbook.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
            <Override PartName="/xl/worksheets/sheet1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
            <Override PartName="/xl/worksheets/sheet2.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
            <Override PartName="/xl/worksheets/sheet3.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
            <Override PartName="/xl/styles.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
            <Override PartName="/xl/sharedStrings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
            <Override PartName="/xl/tables/table1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
            <Override PartName="/xl/drawings/drawing1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
            <Override PartName="/xl/charts/chart1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
            </Types>"#
                .to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships
            xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="workbook" Target="xl/workbook.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>
            </Relationships>"#
                .to_vec(),
        ),
        ("xl/workbook.xml".to_owned(), workbook.to_vec()),
        (
            "xl/_rels/workbook.xml.rels".to_owned(),
            br#"<Relationships
            xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="s1" Target="worksheets/sheet1.xml"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
            <Relationship Id="s2" Target="worksheets/sheet2.xml"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
            <Relationship Id="s3" Target="worksheets/sheet3.xml"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
            <Relationship Id="styles" Target="styles.xml"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"/>
            <Relationship Id="strings" Target="sharedStrings.xml"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"/>
            </Relationships>"#
                .to_vec(),
        ),
        ("xl/worksheets/sheet1.xml".to_owned(), worksheet.to_vec()),
        ("xl/worksheets/sheet2.xml".to_owned(), other.into_bytes()),
        (
            "xl/worksheets/sheet3.xml".to_owned(),
            br#"<worksheet
            xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData/></worksheet>"#
                .to_vec(),
        ),
        (
            "xl/styles.xml".to_owned(),
            br#"<styleSheet
            xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>
            <fonts count="1"><font><name val="Arial"/><sz val="11"/></font></fonts>
            <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
            <borders count="1"><border><left/><right/><top/><bottom/></border></borders>
            <cellXfs count="2"><xf fontId="0" fillId="0" borderId="0" numFmtId="0"/>
            <xf fontId="0" fillId="0" borderId="0" numFmtId="164"/></cellXfs>
            <colors><indexedColors><rgbColor rgb="FF123456"/></indexedColors></colors>
            </styleSheet>"#
                .to_vec(),
        ),
        (
            "xl/sharedStrings.xml".to_owned(),
            br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            count="2" uniqueCount="2"><si><t>duplicate</t></si><si><t>duplicate</t></si></sst>"#
                .to_vec(),
        ),
        (
            "xl/tables/table1.xml".to_owned(),
            br#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            id="1" name="Items" displayName="Items" ref="A4:B6" headerRowCount="1"
            totalsRowCount="0"><autoFilter ref="A4:B6"/><tableColumns count="2">
            <tableColumn id="1" name="Qty"/><tableColumn id="2" name="Cost"/>
            </tableColumns></table>"#
                .to_vec(),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
            br#"<Relationships
            xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="drawing" Target="../drawings/drawing1.xml"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"/>
            <Relationship Id="table" Target="../tables/table1.xml"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"/>
            </Relationships>"#
                .to_vec(),
        ),
        (
            "xl/drawings/drawing1.xml".to_owned(),
            br#"<xdr:wsDr
            xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>
            <xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
            <xdr:ext cx="914400" cy="914400"/><xdr:graphicFrame><a:graphic><a:graphicData>
            <c:chart r:id="chart"/></a:graphicData></a:graphic></xdr:graphicFrame>
            <xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>"#
                .to_vec(),
        ),
        (
            "xl/drawings/_rels/drawing1.xml.rels".to_owned(),
            br#"<Relationships
            xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="chart" Target="../charts/chart1.xml"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"/>
            </Relationships>"#
                .to_vec(),
        ),
        (
            "xl/charts/chart1.xml".to_owned(),
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/>
            <c:cat><c:numRef><c:f>Data!$A$5:$A$6</c:f><c:numCache><c:ptCount val="2"/>
            <c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt>
            </c:numCache></c:numRef></c:cat>
            <c:val><c:numRef><c:f>Data!$B$5:$B$6</c:f><c:numCache><c:ptCount val="2"/>
            <c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt>
            </c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea>
            </c:chart></c:chartSpace>"#
                .to_vec(),
        ),
    ];
    if unpatchable {
        parts.push((
            "xl/pivotcache/pivotCacheDefinition1.xml".to_owned(),
            br#"<pivotCacheDefinition><cacheSource>
            <worksheetSource sheet="Data" ref="A1:B6"/>
            </cacheSource></pivotCacheDefinition>"#
                .to_vec(),
        ));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn worker(bytes: &[u8]) -> Workbook {
    Workbook::open_recalculated_with_seed(bytes, context(), Some(123)).unwrap()
}

fn encode(workbook: &Workbook, budget: SnapshotBudget) -> Vec<Vec<u8>> {
    let mut encoder = WorkbookSnapshotEncoder::new(workbook, Some(context()), budget).unwrap();
    let mut chunks = Vec::new();
    while let Some(chunk) = encoder.next(workbook).unwrap() {
        assert!(chunk.len() <= budget.max_bytes());
        chunks.push(chunk);
    }
    assert!(encoder.next(workbook).unwrap().is_none());
    chunks
}

fn completed_builder(chunks: &[Vec<u8>], budget: SnapshotBudget) -> WorkbookSnapshotBuilder {
    let mut builder = WorkbookSnapshotBuilder::new();
    for chunk in chunks {
        assert!(!builder.push(chunk).unwrap().is_ready());
        builder.advance(budget).unwrap();
    }
    for _ in 0..100_000 {
        if builder.advance(budget).unwrap().is_ready() {
            return builder;
        }
    }
    panic!("snapshot hydration did not complete");
}

fn hydrate(workbook: &Workbook, budget: SnapshotBudget) -> Workbook {
    let chunks = encode(workbook, budget);
    let (peer, received_context) = completed_builder(&chunks, budget)
        .finish()
        .unwrap()
        .into_parts();
    assert_eq!(received_context, Some(context()));
    assert_eq!(peer.proposals().len(), 0);
    assert!(!peer.can_undo());
    assert!(!peer.can_redo());
    peer
}

fn canonical_model(model: &WorkbookModel) -> Vec<Vec<u8>> {
    let budget = SnapshotBudget::new(1, usize::MAX).unwrap();
    let mut encoder = ModelSnapshotEncoder::new();
    let mut chunks = Vec::new();
    while let Some(chunk) = encoder.next(model, budget).unwrap() {
        chunks.push(chunk);
    }
    chunks
}

fn canonical_base(workbook: &Workbook) -> Vec<Vec<u8>> {
    let budget = SnapshotBudget::new(1, usize::MAX).unwrap();
    let mut encoder = AuthoritySnapshotEncoder::new(&workbook.authority, budget).unwrap();
    let mut chunks = Vec::new();
    while let Some(chunk) = encoder.next(budget).unwrap() {
        if unframe(&chunk).unwrap().0 == ChunkKind::AuthorityBase {
            chunks.push(chunk);
        }
    }
    chunks
}

fn canonical_preserved(workbook: &Workbook) -> Vec<Vec<u8>> {
    let budget = SnapshotBudget::new(1, usize::MAX).unwrap();
    let mut encoder = PreservedSnapshotEncoder::new();
    let mut chunks = Vec::new();
    while let Some(chunk) = encoder.next(&workbook.preserved, budget).unwrap() {
        chunks.push(chunk);
    }
    chunks
}

fn assert_current_identity(worker: &Workbook, peer: &Workbook) {
    assert_eq!(
        canonical_model(worker.model()),
        canonical_model(peer.model())
    );
    assert_eq!(canonical_base(worker), canonical_base(peer));
    assert_eq!(canonical_preserved(worker), canonical_preserved(peer));
    assert_eq!(worker.preserved.origins, peer.preserved.origins);
    assert_eq!(
        worker.preserved.shared_string_cells,
        peer.preserved.shared_string_cells
    );
    assert_eq!(worker.preserved.axes, peer.preserved.axes);
    assert_eq!(worker.preserved.created, peer.preserved.created);
    assert_eq!(
        worker
            .source_container
            .as_ref()
            .map(SourceContainer::as_bytes),
        peer.source_container.as_ref().map(SourceContainer::as_bytes),
    );
    let facts = |slot: &Option<PackageSlot>| {
        slot.as_ref().map(|slot| match slot {
            PackageSlot::Present(package) => xlsx_parse::PackageFacts::from_package(package),
            PackageSlot::Deferred { facts, .. } => facts.clone(),
        })
    };
    assert_eq!(facts(&worker.source_package), facts(&peer.source_package));
    assert_eq!(worker.active_sheet, peer.active_sheet);
    assert_eq!(worker.rand_seed, peer.rand_seed);
    assert_eq!(worker.model_epoch, peer.model_epoch);
    assert_eq!(worker.version(), peer.version());
    assert_eq!(worker.last_calculation(), peer.last_calculation());
    assert_eq!(worker.edited_since_open, peer.edited_since_open);
    assert_eq!(worker.recalculated_since_open, peer.recalculated_since_open);
    assert_eq!(
        worker.moved_references_since_open,
        peer.moved_references_since_open
    );
    assert_eq!(worker.opened_anchors, peer.opened_anchors);
    assert_eq!(worker.graph.is_some(), peer.graph.is_some());
    assert_eq!(
        worker.encode_state_vector_v1(),
        peer.encode_state_vector_v1()
    );
    assert!(peer.pending_remote_updates.is_empty());
    assert!(!peer.authority.has_pending_updates());
}

fn assert_step_budget(budget: SnapshotBudget) -> crate::snapshot::step::StepWork {
    let work = crate::snapshot::step::current();
    assert!(work.records <= budget.max_records(), "{work:?}");
    assert!(work.bytes <= budget.max_bytes(), "{work:?}");
    work
}

fn decoded_snapshot(chunks: &[Vec<u8>]) -> Vec<(ChunkKind, u64, Vec<u8>)> {
    let mut records = Vec::new();
    let mut logical = Vec::new();
    for chunk in chunks {
        let (_, _, payload) = unframe(chunk).unwrap();
        let mut reader = Reader::new(payload);
        reader.var_u64().unwrap();
        reader.var_u64().unwrap();
        let length = reader.var_usize().unwrap();
        let offset = reader.var_usize().unwrap();
        assert_eq!(offset, logical.len());
        logical.extend_from_slice(reader.rest());
        if logical.len() != length {
            continue;
        }
        let (kind, ordinal, payload) = unframe(&logical).unwrap();
        if kind == ChunkKind::Header && ordinal == 0 {
            let mut reader = Reader::new(payload);
            let mut header = SnapshotHeader::decode(reader.bytes().unwrap()).unwrap();
            header.snapshot_id = 0;
            let mut writer = Writer::new();
            writer.bytes(&header.encode());
            writer.raw(reader.rest());
            records.push((kind, ordinal, writer.into_bytes()));
        } else if kind != ChunkKind::End {
            records.push((kind, ordinal, payload.to_vec()));
        }
        logical.clear();
    }
    assert!(logical.is_empty());
    records
}

#[test]
fn snapshot_unchanged_hydrated_peer_recaptures_without_package_parse() {
    for budget in budgets() {
        for bytes in [source(false, false), source(true, true)] {
            let worker = worker(&bytes);
            let parses = crate::snapshot::package::rebuild_count();
            let chunks = encode(&worker, budget);
            let (peer, _) = completed_builder(&chunks, budget)
                .finish()
                .unwrap()
                .into_parts();
            assert_current_identity(&worker, &peer);
            assert!(peer.source_package_is_unmaterialized_for_test());
            let recaptured = encode(&peer, budget);
            assert_eq!(decoded_snapshot(&chunks), decoded_snapshot(&recaptured));
            assert_eq!(crate::snapshot::package::rebuild_count(), parses);
            assert!(peer.source_package_is_unmaterialized_for_test());
        }
        let worker = Workbook::from_model(WorkbookModel {
            sheets: vec![crate::Sheet::new("Empty")],
            ..WorkbookModel::default()
        })
        .unwrap();
        let chunks = encode(&worker, budget);
        let (peer, _) = completed_builder(&chunks, budget)
            .finish()
            .unwrap()
            .into_parts();
        assert_eq!(
            decoded_snapshot(&chunks),
            decoded_snapshot(&encode(&peer, budget))
        );
    }
}

#[test]
fn snapshot_structured_reference_insert_rows_refusal_matches_source() {
    let mut worker = worker(&source(false, false));
    let mut peer = hydrate(&worker, budgets()[0]);
    assert_edited_identity(&worker, &peer);
    for workbook in [&mut worker, &mut peer] {
        workbook
            .edit_cell(SheetId(0), CellRef::new(0, 0), "7", context())
            .unwrap();
        workbook.recalculate_all(context());
        workbook.set_active_sheet(SheetId(0)).unwrap();
    }
    let operation = Op::InsertRows {
        sheet: SheetId(0),
        at: 0,
        count: 1,
    };
    for workbook in [&mut worker, &mut peer] {
        let before = canonical_model(workbook.model());
        let failure = workbook
            .apply_ops(vec![operation.clone()], context())
            .unwrap_err();
        assert!(matches!(
            failure,
            crate::Error::Operation(xlsx_ops::OpError::FormulaNotRewritable {
                sheet: SheetId(0),
                cell,
            }) if cell == CellRef::new(2, 0)
        ));
        assert_eq!(canonical_model(workbook.model()), before);
    }
    assert_edited_identity(&worker, &peer);
}

#[test]
fn snapshot_push_rejects_every_skipped_chunk_without_advance() {
    let worker = worker(&source(false, false));
    for budget in budgets() {
        let chunks = encode(&worker, budget);
        for skipped in 1..chunks.len() - 1 {
            let mut builder = WorkbookSnapshotBuilder::new();
            for chunk in &chunks[..skipped] {
                builder.push(chunk).unwrap();
                assert_step_budget(budget);
            }
            assert!(
                builder.push(&chunks[skipped + 1]).is_err(),
                "skipped {skipped}"
            );
            assert_step_budget(budget);
        }
    }
}

fn corrupt_header_counts(
    chunk: &[u8],
    source_length: Option<usize>,
    calculation_counts: Option<[usize; 3]>,
) -> Vec<u8> {
    let (kind, ordinal, payload) = unframe(chunk).unwrap();
    let mut reader = Reader::new(payload);
    let snapshot_id = reader.var_u64().unwrap();
    let logical_ordinal = reader.var_u64().unwrap();
    let length = reader.var_usize().unwrap();
    assert_eq!(reader.var_usize().unwrap(), 0);
    assert_eq!(reader.rest().len(), length);
    let (_, _, payload) = unframe(reader.rest()).unwrap();
    let mut reader = Reader::new(payload);
    let mut header = SnapshotHeader::decode(reader.bytes().unwrap()).unwrap();
    let flags = [
        reader.bool().unwrap(),
        reader.bool().unwrap(),
        reader.bool().unwrap(),
    ];
    let source_length = source_length.or(reader.option(Reader::var_usize).unwrap());
    let mut styles = [0; 7];
    for count in &mut styles {
        *count = reader.var_usize().unwrap();
    }
    let mut counts = [0; 3];
    for count in &mut counts {
        *count = reader.var_usize().unwrap();
    }
    reader.finish().unwrap();
    if let Some(declared) = calculation_counts {
        counts = declared;
        header.chunk_counts[index(ChunkKind::Header)] =
            1 + counts.iter().map(|n| *n as u64).sum::<u64>();
    }
    let mut writer = Writer::new();
    writer.bytes(&header.encode());
    for flag in flags {
        writer.bool(flag);
    }
    writer.option(source_length, Writer::var_usize);
    for count in styles.into_iter().chain(counts) {
        writer.var_usize(count);
    }
    let logical = frame(kind, logical_ordinal, &writer.into_bytes());
    let mut writer = Writer::new();
    writer.var_u64(snapshot_id);
    writer.var_u64(logical_ordinal);
    writer.var_usize(logical.len());
    writer.var_usize(0);
    writer.raw(&logical);
    frame(kind, ordinal, &writer.into_bytes())
}

#[test]
fn snapshot_huge_source_length_is_rejected_before_allocation() {
    let budget = budgets()[0];
    let worker = worker(&source(false, false));
    let mut chunks = encode(&worker, budget);
    chunks[0] = corrupt_header_counts(&chunks[0], Some(usize::MAX / 4), None);
    let mut builder = WorkbookSnapshotBuilder::new();
    for chunk in &chunks[..chunks.len() - 1] {
        builder.push(chunk).unwrap();
        assert_step_budget(budget);
        builder.advance(budget).unwrap();
        assert_eq!(assert_step_budget(budget).initialized_bytes, 0);
        assert!(builder.source.is_none());
    }
    let failure = builder.push(chunks.last().unwrap()).unwrap_err();
    assert_eq!(failure.to_string(), "snapshot source length differs");
    assert_step_budget(budget);
    assert!(builder.source.is_none());
}

#[test]
fn snapshot_huge_calculation_counts_do_not_allocate_declared_storage() {
    let budget = budgets()[0];
    let worker = worker(&source(false, false));
    let chunks = encode(&worker, budget);
    let header = corrupt_header_counts(&chunks[0], None, Some([usize::MAX / 4, 0, 0]));
    let mut builder = WorkbookSnapshotBuilder::new();
    builder.push(&header).unwrap();
    assert_step_budget(budget);
    builder.advance(budget).unwrap();
    assert_step_budget(budget);
    let state = builder.header.as_ref().unwrap();
    assert_eq!(state.header.last_calculation.changed.capacity(), 0);
    assert_eq!(state.header.last_calculation.cycle_cells.capacity(), 0);
    assert_eq!(state.header.last_calculation.limited_cells.capacity(), 0);
    let mut refused = false;
    for chunk in &chunks[1..] {
        match builder.push(chunk) {
            Ok(_) => {
                assert_step_budget(budget);
            }
            Err(failure) => {
                assert_eq!(
                    failure.to_string(),
                    "snapshot chunks are missing or reordered"
                );
                assert_step_budget(budget);
                refused = true;
                break;
            }
        };
    }
    assert!(refused);
}

#[test]
fn snapshot_calculation_storage_growth_respects_step_budget() {
    let budget = budgets()[0];
    let mut worker = worker(&source(false, false));
    worker.last_calculation.changed = (0..257)
        .map(|row| CellAddress {
            sheet: SheetId(0),
            cell: CellRef::new(row, 0),
        })
        .collect();
    let chunks = encode(&worker, budget);
    let mut builder = WorkbookSnapshotBuilder::new();
    for chunk in chunks {
        builder.push(&chunk).unwrap();
        assert_step_budget(budget);
        builder.advance(budget).unwrap();
        assert_step_budget(budget);
    }
    let mut ready = false;
    for _ in 0..100_000 {
        ready = builder.advance(budget).unwrap().is_ready();
        assert_step_budget(budget);
        if ready {
            break;
        }
    }
    assert!(ready);
    let (peer, _) = builder.finish().unwrap().into_parts();
    assert_current_identity(&worker, &peer);
}

#[test]
fn snapshot_long_defined_name_chain_respects_single_record_budget() {
    let mut model = WorkbookModel::default();
    let mut sheet = crate::Sheet::new("Data");
    sheet.set_cell(
        CellRef::new(0, 0),
        Cell {
            value: CellValue::Number { value: 2.0 },
            formula: None,
            style: None,
        },
    );
    sheet.set_cell(
        CellRef::new(0, 1),
        Cell {
            value: CellValue::Number { value: 3.0 },
            formula: Some("N_0".to_owned()),
            style: None,
        },
    );
    model.sheets = vec![sheet];
    let names = 2_048;
    model.defined_names = (0..names)
        .map(|index| xlsx_model::DefinedName {
            name: format!("N_{index}"),
            formula: if index + 1 == names {
                "Data!A1+NOW()".to_owned()
            } else {
                format!("N_{}", index + 1)
            },
            local_sheet: None,
            hidden: false,
        })
        .collect();
    let mut worker = Workbook::from_model(model).unwrap();
    worker.graph = Some(xlsx_calc::graph::DepGraph::build(worker.model()));
    assert!(worker.graph.is_some());
    let budget = SnapshotBudget::new(1, 1_024).unwrap();
    let chunks = encode(&worker, budget);
    let mut builder = WorkbookSnapshotBuilder::new();
    for chunk in chunks {
        builder.push(&chunk).unwrap();
        assert_step_budget(budget);
        builder.advance(budget).unwrap();
        assert_step_budget(budget);
    }
    let mut graph_steps = 0;
    let mut ready = false;
    for _ in 0..100_000 {
        let graph = builder.graph.is_some();
        ready = builder.advance(budget).unwrap().is_ready();
        let work = assert_step_budget(budget);
        graph_steps += usize::from(graph && work.records != 0);
        if ready {
            break;
        }
    }
    assert!(ready);
    assert!(graph_steps >= names * 2);
    let (peer, _) = builder.finish().unwrap().into_parts();
    assert_current_identity(&worker, &peer);
    let graph = peer.graph.as_ref().unwrap();
    assert_eq!(
        graph
            .dependents_of(SheetId(0), CellRef::new(0, 0))
            .collect::<Vec<_>>(),
        vec![(SheetId(0), CellRef::new(0, 1))]
    );
    assert_eq!(
        graph.volatile_cells().collect::<Vec<_>>(),
        vec![(SheetId(0), CellRef::new(0, 1))]
    );
}

fn single_large_sheet_snapshot(rows: u32) {
    let mut model = WorkbookModel::default();
    let mut sheet = xlsx_model::Sheet::new("Large");
    for row in 0..rows {
        for col in 0..20 {
            sheet.set_cell(
                CellRef::new(row, col),
                Cell {
                    value: CellValue::Number {
                        value: f64::from(row * 20 + col),
                    },
                    formula: None,
                    style: None,
                },
            );
        }
    }
    let mut small = xlsx_model::Sheet::new("Small");
    small.set_cell(
        CellRef::new(0, 0),
        Cell {
            value: CellValue::Number { value: 7.0 },
            formula: Some("Large!A1+7".to_owned()),
            style: None,
        },
    );
    model.sheets = vec![sheet, small];
    let bytes = ooxml_opc::rezip_parts(&xlsx_parse::serialize_workbook(&model).unwrap()).unwrap();
    drop(model);
    let worker = worker(&bytes);
    let budget = SnapshotBudget::new(256, 16 * 1024).unwrap();
    let mut encoder = WorkbookSnapshotEncoder::new(&worker, Some(context()), budget).unwrap();
    assert_eq!(encoder.split_fallback_reason(), None);
    let mut builder = WorkbookSnapshotBuilder::new();
    let mut updates = 0;
    let mut initialized = 0;
    let mut initialization_steps = 0;
    let mut max_push = std::time::Duration::ZERO;
    let mut max_advance = std::time::Duration::ZERO;
    let mut max_records = 0;
    let mut max_bytes = 0;
    while let Some(chunk) = encoder.next(&worker).unwrap() {
        assert!(chunk.len() <= budget.max_bytes());
        let (kind, _, payload) = unframe(&chunk).unwrap();
        let mut reader = Reader::new(payload);
        reader.var_u64().unwrap();
        reader.var_u64().unwrap();
        reader.var_usize().unwrap();
        let offset = reader.var_usize().unwrap();
        if kind == ChunkKind::Yrs && offset == 0 {
            updates += 1;
        }
        let start = std::time::Instant::now();
        builder.push(&chunk).unwrap();
        max_push = max_push.max(start.elapsed());
        let work = assert_step_budget(budget);
        max_records = max_records.max(work.records);
        max_bytes = max_bytes.max(work.bytes);
        let start = std::time::Instant::now();
        builder.advance(budget).unwrap();
        max_advance = max_advance.max(start.elapsed());
        let work = assert_step_budget(budget);
        max_records = max_records.max(work.records);
        max_bytes = max_bytes.max(work.bytes);
        initialized += work.initialized_bytes;
        initialization_steps += usize::from(work.initialized_bytes != 0);
    }
    assert!(updates > 2, "the large sheet authority was not subdivided");
    let mut steps = 0;
    loop {
        let start = std::time::Instant::now();
        let ready = builder.advance(budget).unwrap().is_ready();
        max_advance = max_advance.max(start.elapsed());
        let work = assert_step_budget(budget);
        max_records = max_records.max(work.records);
        max_bytes = max_bytes.max(work.bytes);
        initialized += work.initialized_bytes;
        initialization_steps += usize::from(work.initialized_bytes != 0);
        if ready {
            break;
        }
        steps += 1;
        assert!(steps < rows as usize * 40 + bytes.len().div_ceil(budget.max_bytes()) + 10_000);
    }
    eprintln!(
        "{rows} x 20: max push={max_push:?}, max advance={max_advance:?}, \
         max records={max_records}/{}, max bytes={max_bytes}/{}",
        budget.max_records(),
        budget.max_bytes(),
    );
    assert_eq!(initialized, bytes.len());
    assert_eq!(
        initialization_steps,
        bytes.len().div_ceil(budget.max_bytes())
    );
    assert!(initialization_steps > 1);
    let (peer, received_context) = builder.finish().unwrap().into_parts();
    assert_eq!(received_context, Some(context()));
    assert_eq!(peer.source_container.as_ref().unwrap().as_bytes(), bytes);
    assert_current_identity(&worker, &peer);
    assert_eq!(
        worker.authority.encode_state_as_update_v1(),
        peer.authority.encode_state_as_update_v1()
    );
    assert_eq!(worker.save().unwrap(), peer.save().unwrap());
}

#[test]
fn snapshot_single_large_sheet_and_source_steps_respect_budget() {
    single_large_sheet_snapshot(5_000);
}

#[test]
#[ignore]
fn snapshot_single_million_cell_sheet_and_source_steps_respect_budget() {
    single_large_sheet_snapshot(50_000);
}

#[test]
fn snapshot_authority_steps_accept_a_smaller_record_budget() {
    let worker = worker(&source(false, false));
    let encode_budget = SnapshotBudget::new(256, 16 * 1024).unwrap();
    let budget = SnapshotBudget::new(1, 16 * 1024).unwrap();
    let chunks = encode(&worker, encode_budget);
    let mut builder = WorkbookSnapshotBuilder::new();
    for chunk in chunks {
        builder.push(&chunk).unwrap();
        assert_step_budget(budget);
        builder.advance(budget).unwrap();
        assert_step_budget(budget);
    }
    loop {
        let ready = builder.advance(budget).unwrap().is_ready();
        assert_step_budget(budget);
        if ready {
            break;
        }
    }
    let (peer, received_context) = builder.finish().unwrap().into_parts();
    assert_eq!(received_context, Some(context()));
    assert_current_identity(&worker, &peer);
    assert_eq!(
        worker.authority.encode_state_as_update_v1(),
        peer.authority.encode_state_as_update_v1()
    );
    assert_eq!(worker.save().unwrap(), peer.save().unwrap());
}

fn assert_edited_identity(worker: &Workbook, peer: &Workbook) {
    assert_current_identity(worker, peer);
    assert_eq!(worker.history_state(), peer.history_state());
    assert_eq!(worker.save().unwrap(), peer.save().unwrap());
}

fn observe(workbook: &Workbook) -> (Arc<Mutex<Vec<UpdateEvent>>>, crate::UpdateSubscription) {
    let events = Arc::new(Mutex::new(Vec::new()));
    let captured = events.clone();
    let subscription = workbook
        .observe_update_v1(move |event| captured.lock().unwrap().push(event))
        .unwrap();
    (events, subscription)
}

fn assert_events(worker: &Arc<Mutex<Vec<UpdateEvent>>>, peer: &Arc<Mutex<Vec<UpdateEvent>>>) {
    assert_eq!(*worker.lock().unwrap(), *peer.lock().unwrap());
}

#[test]
fn snapshot_finish_only_moves_completed_state() {
    fn send_sync<T: Send + Sync>() {}
    send_sync::<Workbook>();
    let bytes = source(false, false);
    for budget in budgets() {
        let worker = worker(&bytes);
        let chunks = encode(&worker, budget);
        let builder = completed_builder(&chunks, budget);
        let ready = builder.ready.as_ref().unwrap();
        let sheets = ready.workbook.model.sheets.as_ptr();
        let source = ready
            .workbook
            .source_container
            .as_ref()
            .unwrap()
            .as_bytes()
            .as_ptr();
        let changed = ready.workbook.last_calculation.changed.as_ptr();
        let nonce = ready.workbook.version_nonce.as_ptr();
        assert!(ready.workbook.source_package_is_unmaterialized_for_test());
        let (peer, received_context) = builder.finish().unwrap().into_parts();
        assert_eq!(peer.model.sheets.as_ptr(), sheets);
        assert_eq!(
            peer.source_container.as_ref().unwrap().as_bytes().as_ptr(),
            source
        );
        assert_eq!(peer.last_calculation.changed.as_ptr(), changed);
        assert_eq!(peer.version_nonce.as_ptr(), nonce);
        assert_eq!(received_context, Some(context()));
        assert!(peer.source_package_is_unmaterialized_for_test());
        assert_current_identity(&worker, &peer);
    }
}

#[test]
fn snapshot_rejects_noninitial_or_incomplete_state() {
    let bytes = source(false, false);
    let budget = SnapshotBudget::new(1, usize::MAX).unwrap();
    let mut edited = worker(&bytes);
    edited
        .edit_cell(SheetId(0), CellRef::new(0, 0), "9", context())
        .unwrap();
    assert!(WorkbookSnapshotEncoder::new(&edited, None, budget).is_err());
    edited.undo(context()).unwrap();
    assert!(edited.can_redo());
    assert!(WorkbookSnapshotEncoder::new(&edited, None, budget).is_err());
    let mut proposed = worker(&bytes);
    let proposal = proposed.propose(proposal(), context()).unwrap();
    assert!(WorkbookSnapshotEncoder::new(&proposed, None, budget).is_err());
    proposed.reject_proposal(&proposal.id);
    assert!(WorkbookSnapshotEncoder::new(&proposed, None, budget).is_err());
    let collaborative = Workbook::open_collaborative(&bytes, 37).unwrap();
    assert!(WorkbookSnapshotEncoder::new(&collaborative, None, budget).is_err());
    let observed = worker(&bytes);
    let subscription = observed.observe_update_v1(|_| {}).unwrap();
    assert!(WorkbookSnapshotEncoder::new(&observed, None, budget).is_err());
    drop(subscription);
    assert!(WorkbookSnapshotEncoder::new(&observed, None, budget).is_err());
    let mut pending = worker(&bytes);
    pending.pending_remote_updates.push(vec![0]);
    assert!(WorkbookSnapshotEncoder::new(&pending, None, budget).is_err());
    let in_flight = worker(&bytes);
    let transaction = in_flight.authority.snapshot_transaction_for_test();
    assert!(WorkbookSnapshotEncoder::new(&in_flight, None, budget).is_err());
    drop(transaction);
    let mut changed = worker(&bytes);
    let mut encoder = WorkbookSnapshotEncoder::new(&changed, None, budget).unwrap();
    changed.set_active_sheet(SheetId(0)).unwrap();
    assert!(encoder.next(&changed).is_err());

    let worker = worker(&bytes);
    let chunks = encode(&worker, budget);
    let mut wrong = WorkbookSnapshotBuilder::new();
    let (kind, ordinal, payload) = unframe(&chunks[0]).unwrap();
    assert!(wrong.push(&frame(kind, ordinal + 1, payload)).is_err());
    assert!(wrong.push(&chunks[0]).is_err());
    let mut missing = WorkbookSnapshotBuilder::new();
    missing.push(&chunks[0]).unwrap();
    assert!(missing.push(&chunks[2]).is_err());
    let other = encode(&worker, budget);
    let mut stale = WorkbookSnapshotBuilder::new();
    stale.push(&chunks[0]).unwrap();
    assert!(stale.push(&other[1]).is_err());
    assert!(WorkbookSnapshotBuilder::new().finish().is_err());
    let mut premature = WorkbookSnapshotBuilder::new();
    for chunk in &chunks {
        premature.push(chunk).unwrap();
    }
    assert!(premature.finish().is_err());
    let mut missing_end = WorkbookSnapshotBuilder::new();
    for chunk in &chunks[..chunks.len() - 1] {
        missing_end.push(chunk).unwrap();
        missing_end.advance(budget).unwrap();
    }
    assert!(missing_end.finish().is_err());
    let mut extra = completed_builder(&chunks, budget);
    assert!(extra.push(chunks.last().unwrap()).is_err());
    assert!(extra.finish().is_err());

    let mut corrupted = chunks.clone();
    let source_index = corrupted
        .iter()
        .position(|chunk| unframe(chunk).unwrap().0 == ChunkKind::Source)
        .unwrap();
    *corrupted[source_index].last_mut().unwrap() ^= 1;
    let mut builder = WorkbookSnapshotBuilder::new();
    for chunk in &corrupted[..corrupted.len() - 1] {
        builder.push(chunk).unwrap();
    }
    assert!(builder.push(corrupted.last().unwrap()).is_err());
}

#[test]
fn snapshot_authority_identity() {
    let bytes = source(false, false);
    for budget in budgets() {
        let worker = worker(&bytes);
        let peer = hydrate(&worker, budget);
        assert_eq!(worker.client_id(), peer.client_id());
        assert_eq!(
            worker.authority.snapshot_identity(),
            peer.authority.snapshot_identity()
        );
        assert_eq!(
            worker.encode_state_vector_v1(),
            peer.encode_state_vector_v1()
        );
        assert_eq!(
            canonical_model(&worker.authority.materialize().unwrap()),
            canonical_model(&peer.authority.materialize().unwrap()),
        );
        assert!(!peer.authority.can_undo());
        assert!(!peer.authority.can_redo());
        assert!(!peer.authority.has_pending_updates());
        let mut simple = WorkbookModel::default();
        let mut sheet = crate::Sheet::new("Simple");
        sheet.set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Number { value: 2.0 },
                formula: None,
                style: None,
            },
        );
        simple.sheets.push(sheet);
        let simple = Workbook::from_model(simple).unwrap();
        let twin = hydrate(&simple, budget);
        assert_eq!(
            simple.encode_state_as_update_v1(),
            twin.encode_state_as_update_v1()
        );
    }
}

#[test]
fn snapshot_current_state_identity() {
    let bytes = source(false, false);
    for budget in budgets() {
        for valid in [false, true] {
            let mut worker = worker(&bytes);
            worker.authority.set_snapshot_projection_valid(valid);
            worker.model.sheets[0].set_cell(
                CellRef::new(8, 0),
                Cell {
                    value: CellValue::Number { value: -0.0 },
                    formula: None,
                    style: Some(0),
                },
            );
            worker.model.sheets[0].set_cell(
                CellRef::new(8, 1),
                Cell {
                    value: CellValue::Number {
                        value: f64::from_bits(0x7ff8_0000_0000_0042),
                    },
                    formula: Some(String::new()),
                    style: None,
                },
            );
            worker.model.styles.fonts[0].size_pt = Some(-0.0);
            worker.model_epoch = u64::MAX;
            worker.version_nonce = String::new();
            worker.committed_changes = 42;
            worker.rand_seed = Some(0);
            worker.last_calculation.cycle_cells = vec![CellAddress {
                sheet: SheetId(1),
                cell: CellRef::parse_a1("$B$2").unwrap(),
            }];
            worker.last_calculation.limited_cells = vec![CellAddress {
                sheet: SheetId(0),
                cell: CellRef::new(8, 0),
            }];
            let peer = hydrate(&worker, budget);
            assert_current_identity(&worker, &peer);
            assert_eq!(peer.authority.snapshot_projection_valid().unwrap(), valid);
            assert_eq!(worker.model.sheets[0].charts.len(), 1);
            assert_eq!(worker.model.tables.len(), 1);
            assert!(worker.model.sheets[2].iter_cells().next().is_none());
            assert_eq!(worker.preserved.shared_string_cells[0][&(0, 2)], 0);
            assert_eq!(worker.preserved.shared_string_cells[0][&(0, 3)], 1);
        }
        for context in [
            None,
            Some(CalculationOptions { now_serial: None }),
            Some(CalculationOptions {
                now_serial: Some(-0.0),
            }),
            Some(CalculationOptions {
                now_serial: Some(f64::from_bits(0x7ff8_0000_0000_0017)),
            }),
        ] {
            let worker = worker(&bytes);
            let mut encoder = WorkbookSnapshotEncoder::new(&worker, context, budget).unwrap();
            let mut chunks = Vec::new();
            while let Some(chunk) = encoder.next(&worker).unwrap() {
                chunks.push(chunk);
            }
            let (_, decoded) = completed_builder(&chunks, budget)
                .finish()
                .unwrap()
                .into_parts();
            assert_eq!(
                context.map(|options| options.now_serial.map(f64::to_bits)),
                decoded.map(|options| options.now_serial.map(f64::to_bits)),
            );
        }
    }
}

#[test]
fn snapshot_initial_save_identity() {
    let bytes = source(false, false);
    for budget in budgets() {
        for worker in [
            Workbook::open(&bytes).unwrap(),
            Workbook::open_for_read(&bytes).unwrap(),
        ] {
            let peer = hydrate(&worker, budget);
            assert!(peer.source_package_is_unmaterialized_for_test());
            assert_eq!(worker.save().unwrap(), bytes);
            assert_eq!(peer.save().unwrap(), bytes);
            assert_edited_identity(&worker, &peer);
        }
        let worker = worker(&bytes);
        let peer = hydrate(&worker, budget);
        assert_eq!(worker.save().unwrap(), peer.save().unwrap());
    }
}

#[test]
fn snapshot_rebuilt_package_identity() {
    let bytes = source(false, false);
    let mut parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let (_, data) = parts
        .iter_mut()
        .find(|(name, _)| name == "xl/worksheets/sheet1.xml")
        .unwrap();
    *data = std::str::from_utf8(data)
        .unwrap()
        .replace("SUM(Items[Qty])+Input", "SUM(A5:A6)+Input")
        .into_bytes();
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    for budget in budgets() {
        let mut worker = worker(&bytes);
        let mut peer = hydrate(&worker, budget);
        assert!(peer.source_package_is_unmaterialized_for_test());
        assert_edited_identity(&worker, &peer);
        assert!(!peer.source_package_is_unmaterialized_for_test());
        for workbook in [&mut worker, &mut peer] {
            workbook
                .edit_cell(SheetId(0), CellRef::new(0, 0), "7", context())
                .unwrap();
        }
        assert_edited_identity(&worker, &peer);
        assert_eq!(
            worker.recalculate_all(context()),
            peer.recalculate_all(context())
        );
        assert_edited_identity(&worker, &peer);
        worker.set_active_sheet(SheetId(0)).unwrap();
        peer.set_active_sheet(SheetId(0)).unwrap();
        assert_edited_identity(&worker, &peer);
        for workbook in [&mut worker, &mut peer] {
            workbook
                .apply_ops(
                    vec![Op::InsertRows {
                        sheet: SheetId(1),
                        at: 0,
                        count: 1,
                    }],
                    context(),
                )
                .unwrap();
        }
        assert_edited_identity(&worker, &peer);
    }
}

#[test]
fn snapshot_next_edit_identity() {
    let bytes = source(false, false);
    for budget in budgets() {
        let mut worker = worker(&bytes);
        let mut peer = hydrate(&worker, budget);
        assert!(worker.authority.snapshot_projection_valid().unwrap());
        assert!(peer.authority.snapshot_projection_valid().unwrap());
        let (worker_events, _worker_subscription) = observe(&worker);
        let (peer_events, _peer_subscription) = observe(&peer);
        let before = crate::authority::fast_set_cell_count();
        assert_eq!(
            worker
                .edit_cell(SheetId(0), CellRef::new(0, 0), "8", context())
                .unwrap(),
            peer.edit_cell(SheetId(0), CellRef::new(0, 0), "8", context())
                .unwrap(),
        );
        assert_eq!(crate::authority::fast_set_cell_count() - before, 2);
        assert_events(&worker_events, &peer_events);
        assert!(peer.source_package_is_unmaterialized_for_test());
        assert_edited_identity(&worker, &peer);
        let ops = vec![
            Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(4, 1),
                cell: CellState {
                    value: CellValue::Number { value: 12.0 },
                    formula: None,
                    style: Some(0),
                },
            },
            Op::SetCell {
                sheet: SheetId(1),
                at: CellRef::new(2, 2),
                cell: CellState {
                    value: CellValue::Number { value: 4.0 },
                    formula: Some("Data!A1+2".to_owned()),
                    style: None,
                },
            },
            Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(5, 0),
                cell: CellState {
                    value: CellValue::Number { value: 6.0 },
                    formula: None,
                    style: None,
                },
            },
        ];
        assert_eq!(
            worker.apply_ops(ops.clone(), context()).unwrap(),
            peer.apply_ops(ops, context()).unwrap(),
        );
        assert_events(&worker_events, &peer_events);
        assert_edited_identity(&worker, &peer);
    }
}

#[test]
fn snapshot_replay_history_identity() {
    let bytes = source(false, false);
    for budget in budgets() {
        let mut worker = worker(&bytes);
        let mut peer = hydrate(&worker, budget);
        assert!(peer.preserved_undo.is_empty());
        assert!(peer.preserved_redo.is_empty());
        assert!(peer.update_observers.lock().unwrap().listeners.is_empty());
        assert_eq!(peer.update_observers.lock().unwrap().next_id, 0);
        let (worker_events, _worker_subscription) = observe(&worker);
        let (peer_events, _peer_subscription) = observe(&peer);
        assert_events(&worker_events, &peer_events);
        assert_eq!(
            worker
                .edit_cell(SheetId(0), CellRef::new(0, 0), "2", context())
                .unwrap(),
            peer.edit_cell(SheetId(0), CellRef::new(0, 0), "2", context())
                .unwrap(),
        );
        assert!(worker_events.lock().unwrap().is_empty());
        assert_eq!(
            worker
                .edit_cell(SheetId(99), CellRef::new(0, 0), "3", context())
                .unwrap_err()
                .to_string(),
            peer.edit_cell(SheetId(99), CellRef::new(0, 0), "3", context())
                .unwrap_err()
                .to_string(),
        );
        for workbook in [&mut worker, &mut peer] {
            workbook
                .edit_cell(SheetId(0), CellRef::new(0, 0), "6", context())
                .unwrap();
        }
        assert_edited_identity(&worker, &peer);
        assert_eq!(
            worker.undo(context()).unwrap(),
            peer.undo(context()).unwrap()
        );
        assert_events(&worker_events, &peer_events);
        assert_edited_identity(&worker, &peer);
        assert_eq!(
            worker.redo(context()).unwrap(),
            peer.redo(context()).unwrap()
        );
        assert_edited_identity(&worker, &peer);
        let ops = vec![Op::AddSheet {
            index: 1,
            name: "Added".to_owned(),
        }];
        assert_eq!(
            worker.apply_ops(ops.clone(), context()).unwrap(),
            peer.apply_ops(ops, context()).unwrap(),
        );
        assert_edited_identity(&worker, &peer);
        assert_eq!(
            worker.undo(context()).unwrap(),
            peer.undo(context()).unwrap()
        );
        assert_edited_identity(&worker, &peer);
        assert_eq!(
            worker.redo(context()).unwrap(),
            peer.redo(context()).unwrap()
        );
        assert_edited_identity(&worker, &peer);
        let frame = worker.model.sheets[0].charts[0].frame_id();
        assert_eq!(
            worker
                .move_chart(SheetId(0), &frame, 20.0, 10.0, context())
                .unwrap(),
            peer.move_chart(SheetId(0), &frame, 20.0, 10.0, context())
                .unwrap(),
        );
        assert_edited_identity(&worker, &peer);
        assert_eq!(
            worker.undo(context()).unwrap(),
            peer.undo(context()).unwrap()
        );
        assert_edited_identity(&worker, &peer);
        assert_eq!(
            worker.redo(context()).unwrap(),
            peer.redo(context()).unwrap()
        );
        assert_edited_identity(&worker, &peer);
        assert_events(&worker_events, &peer_events);
    }
}

fn proposal() -> ProposalRequest {
    ProposalRequest {
        agent_id: "test".to_owned(),
        note: Some(String::new()),
        edits: vec![ProposalEditInput {
            sheet: SheetId(0),
            cell: CellRef::new(0, 0),
            input: "6".to_owned(),
            number_format: None,
        }],
    }
}

fn request(workbook: &Workbook, sheet: usize) -> EditRequest {
    EditRequest {
        expect_version: workbook.version(),
        source: Default::default(),
        history: Default::default(),
        calculation: CalculationRequest::default(),
        steps: vec![EditStep::new(EditOperation::SetCellInputs {
            target: RangeTarget {
                sheet_id: format!("sheet:{sheet}"),
                range: RangeAddress::A1 {
                    a1: "A1".to_owned(),
                },
            },
            inputs: vec![vec!["5".to_owned()]],
        })],
    }
}

#[test]
fn snapshot_facts_behavior_identity() {
    let bytes = source(true, true);
    for budget in budgets() {
        let mut worker = worker(&bytes);
        let mut peer = hydrate(&worker, budget);
        let viewport = Viewport {
            x: 0.0,
            y: 0.0,
            width: 600.0,
            height: 400.0,
        };
        for sheet in [SheetId(0), SheetId(1), SheetId(2)] {
            assert_eq!(
                worker.sheet_info_for(sheet).unwrap(),
                peer.sheet_info_for(sheet).unwrap()
            );
            assert_eq!(
                worker.display_list_for(sheet, &viewport).unwrap(),
                peer.display_list_for(sheet, &viewport).unwrap(),
            );
        }
        assert!(peer.source_package_is_unmaterialized_for_test());
        for sheet in [0, 1] {
            assert_eq!(
                format!(
                    "{:?}",
                    worker.validate_edits(&request(&worker, sheet)).unwrap()
                ),
                format!("{:?}", peer.validate_edits(&request(&peer, sheet)).unwrap()),
            );
        }
        let ops = vec![Op::RenameSheet {
            sheet: SheetId(0),
            name: "Renamed".to_owned(),
        }];
        assert_eq!(
            worker
                .apply_ops(ops.clone(), context())
                .unwrap_err()
                .to_string(),
            peer.apply_ops(ops, context()).unwrap_err().to_string(),
        );
        assert_eq!(
            worker.propose(proposal(), context()).unwrap(),
            peer.propose(proposal(), context()).unwrap(),
        );
        assert_eq!(
            worker.display_list_for(SheetId(0), &viewport).unwrap(),
            peer.display_list_for(SheetId(0), &viewport).unwrap(),
        );
        assert!(peer.source_package_is_unmaterialized_for_test());
        let options = XlsxExportOptions::default();
        assert_eq!(
            serde_json::to_value(worker.export_structured(&options).unwrap().unwrap().content,)
                .unwrap(),
            serde_json::to_value(peer.export_structured(&options).unwrap().unwrap().content)
                .unwrap(),
        );
        assert!(!peer.source_package_is_unmaterialized_for_test());
    }
}

#[test]
fn snapshot_calculation_lists_advance_one_address_at_a_time() {
    let bytes = source(false, false);
    let mut worker = worker(&bytes);
    let address = |sheet, cell: &str| CellAddress {
        sheet: SheetId(sheet),
        cell: CellRef::parse_a1(cell).unwrap(),
    };
    worker.last_calculation = crate::CalculationResult {
        changed: vec![address(1, "$B$2"), address(0, "C$4"), address(1, "$B$2")],
        cycle_cells: vec![address(2, "A1"), address(0, "$D5")],
        limited_cells: vec![address(0, "E6")],
    };
    let budget = SnapshotBudget::new(1, usize::MAX).unwrap();
    let chunks = encode(&worker, budget);
    let mut builder = WorkbookSnapshotBuilder::new();
    builder.push(&chunks[0]).unwrap();
    builder.advance(budget).unwrap();
    let mut addresses = 0;
    for chunk in &chunks[1..] {
        let kind = unframe(chunk).unwrap().0;
        builder.push(chunk).unwrap();
        builder.advance(budget).unwrap();
        if kind == ChunkKind::Header {
            addresses += 1;
            let calculation = &builder.header.as_ref().unwrap().header.last_calculation;
            assert_eq!(
                calculation.changed.len()
                    + calculation.cycle_cells.len()
                    + calculation.limited_cells.len(),
                addresses,
            );
        }
    }
    assert_eq!(addresses, 6);
    for _ in 0..100_000 {
        if builder.advance(budget).unwrap().is_ready() {
            break;
        }
    }
    let (peer, _) = builder.finish().unwrap().into_parts();
    assert_eq!(peer.last_calculation, worker.last_calculation);
}

#[test]
fn snapshot_transport_fragments_and_hydration_cap() {
    let bytes = source(false, false);
    let worker = worker(&bytes);
    let small = SnapshotBudget::new(1, 64).unwrap();
    let encoder = WorkbookSnapshotEncoder::new(&worker, None, small).unwrap();
    assert!(encoder.split_fallback_reason().is_some());
    let chunks = encode(&worker, small);
    assert!(unframe(&chunks[1]).unwrap().0 == ChunkKind::Header);
    let mut builder = WorkbookSnapshotBuilder::new();
    for chunk in &chunks {
        builder.push(chunk).unwrap();
    }
    let mut refused = false;
    for _ in 0..100_000 {
        match builder.advance(small) {
            Ok(progress) => assert!(!progress.is_ready()),
            Err(error) => {
                assert!(error.to_string().contains("advance byte budget"));
                refused = true;
                break;
            }
        }
    }
    assert!(refused);
    let large = SnapshotBudget::new(1, usize::MAX).unwrap();
    while !builder.advance(large).unwrap().is_ready() {}
    let (peer, _) = builder.finish().unwrap().into_parts();
    assert_current_identity(&worker, &peer);
}
