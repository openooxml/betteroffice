use std::sync::OnceLock;

use ooxml_opc::SourceContainer;
use xlsx_parse::{PackageFacts, PackageFactsEncoder, PackageFactsView, PreservedPackage};

use super::{SnapshotError, SnapshotResult};
use crate::{Error, Result};

const PACKAGE_DEFERRED: bool = true;

pub(crate) enum PackageSlot {
    Present(PreservedPackage),
    Deferred {
        source: SourceContainer,
        facts: PackageFacts,
        rebuilt: OnceLock<std::result::Result<PreservedPackage, String>>,
    },
}

impl PackageSlot {
    pub(crate) fn deferred(source: SourceContainer, facts: PackageFacts) -> Self {
        Self::Deferred {
            source,
            facts,
            rebuilt: OnceLock::new(),
        }
    }

    pub(crate) fn from_snapshot(
        source: SourceContainer,
        facts: PackageFacts,
    ) -> SnapshotResult<Self> {
        if PACKAGE_DEFERRED {
            Ok(Self::deferred(source, facts))
        } else {
            rebuild(&source)
                .map(Self::Present)
                .map_err(SnapshotError::new)
        }
    }

    pub(crate) fn facts(&self) -> PackageFactsView<'_> {
        match self {
            Self::Present(package) => PackageFactsView::from_package(package),
            Self::Deferred { facts, .. } => PackageFactsView::from_facts(facts),
        }
    }

    pub(crate) fn next_facts(
        &self,
        encoder: &mut PackageFactsEncoder,
        max_bytes: usize,
        retained: bool,
    ) -> SnapshotResult<Option<Vec<u8>>> {
        match self {
            Self::Present(package) => encoder.next(package, max_bytes),
            Self::Deferred { facts, .. } if retained => encoder.next_from_facts(facts, max_bytes),
            Self::Deferred { .. } => encoder.next(
                self.materialize()
                    .map_err(|failure| SnapshotError::new(failure.to_string()))?,
                max_bytes,
            ),
        }
        .map_err(|failure| SnapshotError::new(failure.to_string()))
    }

    pub(crate) fn materialize(&self) -> Result<&PreservedPackage> {
        match self {
            Self::Present(package) => Ok(package),
            Self::Deferred {
                source, rebuilt, ..
            } => rebuilt
                .get_or_init(|| rebuild(source))
                .as_ref()
                .map_err(|message| Error::Package(message.clone())),
        }
    }
}

fn rebuild(source: &SourceContainer) -> std::result::Result<PreservedPackage, String> {
    #[cfg(test)]
    REBUILDS.set(REBUILDS.get() + 1);
    let parts = ooxml_opc::unzip_parts(source.as_bytes())?;
    xlsx_parse::parse_workbook_with_owned_package(parts)
        .map(|parsed| parsed.package)
        .map_err(|error| error.to_string())
}

#[cfg(test)]
thread_local! {
    static REBUILDS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

#[cfg(test)]
pub(crate) fn rebuild_count() -> usize {
    REBUILDS.get()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        CalculationOptions, CalculationRequest, Cell, CellRef, CellValue, EditFailureCode,
        EditOperation, EditRequest, EditStep, ProposalEditInput, ProposalRequest, RangeAddress,
        RangeTarget, Sheet, SheetId, Viewport, Workbook, XlsxExportOptions,
    };
    use xlsx_ops::Op;

    fn source(protected: bool, unpatchable: bool) -> Vec<u8> {
        let mut model = xlsx_model::Workbook::default();
        let mut sheet = Sheet::new("Data");
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
                formula: Some("A1+1".to_owned()),
                style: None,
            },
        );
        model.sheets = vec![sheet, Sheet::new("Other")];
        let mut parts = xlsx_parse::serialize_workbook(&model).unwrap();
        let worksheet = parts
            .iter_mut()
            .find(|(path, _)| path == "xl/worksheets/sheet1.xml")
            .unwrap();
        worksheet.1 = String::from_utf8(std::mem::take(&mut worksheet.1))
            .unwrap()
            .replace(
                "</worksheet>",
                r#"<drawing
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                r:id="draw"/></worksheet>"#,
            )
            .into_bytes();
        if protected {
            let worksheet = parts
                .iter_mut()
                .find(|(path, _)| path == "xl/worksheets/sheet2.xml")
                .unwrap();
            worksheet.1 = String::from_utf8(std::mem::take(&mut worksheet.1))
                .unwrap()
                .replace(
                    "</worksheet>",
                    r#"<sheetProtection sheet="1"/></worksheet>"#,
                )
                .into_bytes();
        }
        parts.extend([
            (
                "xl/worksheets/_rels/sheet1.xml.rels".to_owned(),
                br#"<Relationships
                    xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship Id="draw" Target="../drawings/drawing1.xml"
Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"/>
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
                    <xdr:ext cx="914400" cy="914400"/>
                    <xdr:graphicFrame><a:graphic><a:graphicData><c:chart r:id="chart"/>
                    </a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>
                    </xdr:oneCellAnchor></xdr:wsDr>"#
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
                    <c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/>
                    <c:cat><c:numRef><c:f>Data!$A$1</c:f><c:numCache>
                    <c:pt idx="0"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:cat>
                    <c:val><c:numRef><c:f>Data!$A$1</c:f><c:numCache>
                    <c:pt idx="0"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>
                    </c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
                    .to_vec(),
            ),
        ]);
        if unpatchable {
            parts.push((
                "xl/pivotcache/pivotCacheDefinition1.xml".to_owned(),
                br#"<pivotCacheDefinition><cacheSource>
                    <worksheetSource sheet="Data" ref="A1:B4"/>
                    </cacheSource></pivotCacheDefinition>"#
                    .to_vec(),
            ));
        }
        ooxml_opc::rezip_parts(&parts).unwrap()
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
    fn deferred_slot_does_not_parse_until_save() {
        let bytes = source(true, true);
        let mut workbook = Workbook::open(&bytes).unwrap();
        let mut present = Workbook::open(&bytes).unwrap();
        assert_eq!(workbook.model().sheets[0].charts.len(), 1);
        workbook.defer_source_package_for_test().unwrap();
        assert!(workbook.source_package_is_unmaterialized_for_test());
        let viewport = Viewport {
            x: 0.0,
            y: 0.0,
            width: 600.0,
            height: 400.0,
        };
        let initial = workbook.display_list(&viewport).unwrap();
        assert_eq!(initial.charts.len(), 1);
        assert_eq!(initial, present.display_list(&viewport).unwrap());
        for workbook in [&mut workbook, &mut present] {
            workbook
                .edit_cell(
                    SheetId(0),
                    CellRef::new(0, 0),
                    "4",
                    CalculationOptions::default(),
                )
                .unwrap();
        }
        let repainted = workbook.display_list(&viewport).unwrap();
        assert_eq!(repainted.charts.len(), 1);
        assert_eq!(repainted, present.display_list(&viewport).unwrap());
        assert_ne!(initial, repainted);
        workbook.sheet_info().unwrap();
        workbook.sheet_info_for(SheetId(1)).unwrap();
        let _ = workbook.version();
        assert!(
            workbook
                .validate_edits(&request(&workbook, 0))
                .unwrap()
                .is_ok()
        );
        let refused = workbook
            .validate_edits(&request(&workbook, 1))
            .unwrap()
            .unwrap_err();
        assert_eq!(refused.failure.code, EditFailureCode::LockedTarget);
        let error = workbook
            .apply_ops(
                vec![Op::RenameSheet {
                    sheet: SheetId(0),
                    name: "Renamed".to_owned(),
                }],
                CalculationOptions::default(),
            )
            .unwrap_err();
        assert!(error.to_string().contains("cannot be rewritten"));
        workbook
            .apply_ops(
                vec![Op::InsertRows {
                    sheet: SheetId(0),
                    at: 10,
                    count: 1,
                }],
                CalculationOptions::default(),
            )
            .unwrap();
        workbook
            .propose(
                ProposalRequest {
                    agent_id: "test".to_owned(),
                    note: Some(String::new()),
                    edits: vec![ProposalEditInput {
                        sheet: SheetId(0),
                        cell: CellRef::new(0, 0),
                        input: "6".to_owned(),
                        number_format: None,
                    }],
                },
                CalculationOptions::default(),
            )
            .unwrap();
        workbook.display_list(&viewport).unwrap();
        assert!(workbook.source_package_is_unmaterialized_for_test());
        workbook.save().unwrap();
        assert!(!workbook.source_package_is_unmaterialized_for_test());
    }

    #[test]
    fn rebuilt_package_save_matches_present() {
        let bytes = source(false, false);
        let mut present = Workbook::open(&bytes).unwrap();
        let mut deferred = Workbook::open(&bytes).unwrap();
        deferred.defer_source_package_for_test().unwrap();
        assert_eq!(present.save().unwrap(), deferred.save().unwrap());
        assert!(!deferred.source_package_is_unmaterialized_for_test());
        for workbook in [&mut present, &mut deferred] {
            workbook
                .edit_cell(
                    SheetId(0),
                    CellRef::new(0, 0),
                    "7",
                    CalculationOptions::default(),
                )
                .unwrap();
        }
        assert_eq!(present.save().unwrap(), deferred.save().unwrap());
        for workbook in [&mut present, &mut deferred] {
            workbook.recalculate_all(CalculationOptions::default());
        }
        assert_eq!(present.save().unwrap(), deferred.save().unwrap());
        for workbook in [&mut present, &mut deferred] {
            workbook.set_active_sheet(SheetId(1)).unwrap();
        }
        assert_eq!(present.save().unwrap(), deferred.save().unwrap());
        for workbook in [&mut present, &mut deferred] {
            workbook
                .apply_ops(
                    vec![Op::InsertRows {
                        sheet: SheetId(0),
                        at: 0,
                        count: 1,
                    }],
                    CalculationOptions::default(),
                )
                .unwrap();
        }
        assert_eq!(present.save().unwrap(), deferred.save().unwrap());
    }

    #[test]
    fn deferred_structured_export_materializes_and_matches_present() {
        let bytes = source(true, true);
        let present = Workbook::open_for_read(&bytes).unwrap();
        let mut deferred = Workbook::open_for_read(&bytes).unwrap();
        deferred.defer_source_package_for_test().unwrap();
        assert!(deferred.source_package_is_unmaterialized_for_test());
        let options = XlsxExportOptions::default();
        let present = present
            .export_structured(&options)
            .unwrap()
            .unwrap()
            .content;
        let deferred_content = deferred
            .export_structured(&options)
            .unwrap()
            .unwrap()
            .content;
        assert_eq!(
            serde_json::to_value(present).unwrap(),
            serde_json::to_value(deferred_content).unwrap()
        );
        assert!(!deferred.source_package_is_unmaterialized_for_test());
    }

    #[test]
    fn deferred_rebuild_failure_is_cached() {
        let parts = ooxml_opc::unzip_parts(&source(false, false)).unwrap();
        let package = xlsx_parse::parse_workbook_with_owned_package(parts)
            .unwrap()
            .package;
        let slot = PackageSlot::deferred(
            SourceContainer::new(b"invalid zip".to_vec()),
            PackageFacts::from_package(&package),
        );
        let first = slot.materialize().unwrap_err().to_string();
        let PackageSlot::Deferred { rebuilt, .. } = &slot else {
            panic!("expected deferred package");
        };
        let cached = rebuilt.get().unwrap();
        assert!(cached.is_err());
        assert_eq!(slot.materialize().unwrap_err().to_string(), first);
        assert!(std::ptr::eq(cached, rebuilt.get().unwrap()));
    }

    #[test]
    fn snapshot_package_policy_uses_the_selected_mode() {
        let source = SourceContainer::new(source(false, false));
        let package = rebuild(&source).unwrap();
        let slot =
            PackageSlot::from_snapshot(source, PackageFacts::from_package(&package)).unwrap();
        if PACKAGE_DEFERRED {
            let PackageSlot::Deferred { rebuilt, .. } = slot else {
                panic!("expected deferred package");
            };
            assert!(rebuilt.get().is_none());
        } else {
            assert!(matches!(slot, PackageSlot::Present(_)));
        }
    }

    #[test]
    fn deferred_slot_remains_send_and_sync() {
        fn assert_send_sync<T: Send + Sync>() {}
        assert_send_sync::<PackageSlot>();
        assert_send_sync::<Workbook>();
    }
}
