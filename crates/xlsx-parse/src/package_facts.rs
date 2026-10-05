use std::collections::HashSet;

use xlsx_model::CellRef;

use crate::{PreservedPackage, SheetVisibility, SourceSheetKind};

#[derive(Clone, Debug, PartialEq, Eq)]
#[doc(hidden)]
pub struct PackageFacts {
    pub(crate) sheets: Vec<SheetFacts>,
    pub(crate) references: Vec<ReferenceFacts>,
    pub(crate) charts: Vec<(String, Vec<u8>)>,
    pub(crate) any_uncached_formula: bool,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SheetFacts {
    pub(crate) path: String,
    pub(crate) kind: SourceSheetKind,
    pub(crate) visibility: SheetVisibility,
    pub(crate) worksheet: bool,
    pub(crate) protected: bool,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct ReferenceFacts {
    pub(crate) part: String,
    pub(crate) areas: Option<Vec<ReferenceAreaFacts>>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct ReferenceAreaFacts {
    pub(crate) sheet: String,
    pub(crate) end: CellRef,
}

impl PackageFacts {
    #[doc(hidden)]
    pub fn from_package(package: &PreservedPackage) -> Self {
        Self {
            sheets: (0..package.source_sheet_count())
                .map(|index| sheet_facts(package, index))
                .collect(),
            references: package
                .unpatchable_references()
                .iter()
                .map(|reference| reference.package_facts())
                .collect(),
            charts: chart_part_indices(package)
                .into_iter()
                .map(|index| package.parts[index].clone())
                .collect(),
            any_uncached_formula: any_uncached_formula(package),
        }
    }
}

pub(crate) fn sheet_facts(package: &PreservedPackage, index: usize) -> SheetFacts {
    SheetFacts {
        path: package
            .source_sheet_part(index)
            .unwrap_or_default()
            .to_owned(),
        kind: package
            .source_sheet_kind(index)
            .unwrap_or(SourceSheetKind::Worksheet),
        visibility: package
            .source_sheet_visibility(index)
            .unwrap_or(SheetVisibility::Visible),
        worksheet: package.source_sheet_is_worksheet(index),
        protected: package.source_sheet_is_protected(index),
    }
}

pub(crate) fn any_uncached_formula(package: &PreservedPackage) -> bool {
    (0..package.source_sheet_count()).any(|index| {
        package
            .source_cell_facts(index)
            .is_some_and(|facts| !facts.uncached_formulas.is_empty())
    })
}

pub(crate) fn chart_part_indices(package: &PreservedPackage) -> Vec<usize> {
    let mut paths: HashSet<&str> = package
        .original_workbook
        .sheets
        .iter()
        .flat_map(|sheet| &sheet.charts)
        .map(|chart| chart.part.trim_start_matches('/'))
        .collect();
    paths.extend(package.content_types.iter().filter_map(|entry| {
        entry
            .attribute("ContentType")
            .filter(|kind| {
                crate::chart::CHART_CONTENT_TYPES
                    .iter()
                    .any(|known| kind.eq_ignore_ascii_case(known))
            })
            .and_then(|_| entry.attribute("PartName"))
            .map(|path| path.trim_start_matches('/'))
    }));
    package
        .parts
        .iter()
        .enumerate()
        .filter_map(|(index, (path, _))| {
            let path = path.trim_start_matches('/');
            let conventional = path.to_ascii_lowercase();
            (paths.contains(path)
                || (conventional.starts_with("xl/charts/")
                    && conventional.ends_with(".xml")
                    && !conventional.contains("/_rels/")))
            .then_some(index)
        })
        .collect()
}

#[derive(Clone, Copy)]
#[doc(hidden)]
pub struct PackageFactsView<'a>(FactsSource<'a>);

#[derive(Clone, Copy)]
enum FactsSource<'a> {
    Package(&'a PreservedPackage),
    Facts(&'a PackageFacts),
}

impl<'a> PackageFactsView<'a> {
    pub fn from_package(package: &'a PreservedPackage) -> Self {
        Self(FactsSource::Package(package))
    }

    pub fn from_facts(facts: &'a PackageFacts) -> Self {
        Self(FactsSource::Facts(facts))
    }

    pub fn source_present(&self) -> bool {
        true
    }

    pub fn source_sheet_count(&self) -> usize {
        match self.0 {
            FactsSource::Package(package) => package.source_sheet_count(),
            FactsSource::Facts(facts) => facts.sheets.len(),
        }
    }

    pub fn has_uncached_source_formulas(&self) -> bool {
        match self.0 {
            FactsSource::Package(package) => any_uncached_formula(package),
            FactsSource::Facts(facts) => facts.any_uncached_formula,
        }
    }

    pub fn source_sheet_part(&self, index: usize) -> Option<&'a str> {
        match self.0 {
            FactsSource::Package(package) => package.source_sheet_part(index),
            FactsSource::Facts(facts) => facts.sheets.get(index).map(|sheet| sheet.path.as_str()),
        }
    }

    pub fn source_sheet_kind(&self, index: usize) -> Option<SourceSheetKind> {
        match self.0 {
            FactsSource::Package(package) => package.source_sheet_kind(index),
            FactsSource::Facts(facts) => facts.sheets.get(index).map(|sheet| sheet.kind),
        }
    }

    pub fn source_sheet_visibility(&self, index: usize) -> Option<SheetVisibility> {
        match self.0 {
            FactsSource::Package(package) => package.source_sheet_visibility(index),
            FactsSource::Facts(facts) => facts.sheets.get(index).map(|sheet| sheet.visibility),
        }
    }

    pub fn source_sheet_is_worksheet(&self, index: usize) -> bool {
        match self.0 {
            FactsSource::Package(package) => package.source_sheet_is_worksheet(index),
            FactsSource::Facts(facts) => {
                facts.sheets.get(index).is_none_or(|sheet| sheet.worksheet)
            }
        }
    }

    pub fn source_sheet_is_protected(&self, index: usize) -> bool {
        match self.0 {
            FactsSource::Package(package) => package.source_sheet_is_protected(index),
            FactsSource::Facts(facts) => {
                facts.sheets.get(index).is_some_and(|sheet| sheet.protected)
            }
        }
    }

    pub fn unpatchable_reference_part(&self) -> Option<&'a str> {
        match self.0 {
            FactsSource::Package(package) => package.unpatchable_reference_part(),
            FactsSource::Facts(facts) => facts
                .references
                .first()
                .map(|reference| reference.part.as_str()),
        }
    }

    pub fn reference_naming_sheet(&self, sheet: &str) -> Option<&'a str> {
        match self.0 {
            FactsSource::Package(package) => package.reference_naming_sheet(sheet),
            FactsSource::Facts(facts) => {
                stranded(facts, |area| area.sheet.eq_ignore_ascii_case(sheet))
            }
        }
    }

    pub fn reference_moved_by_rows(&self, sheet: &str, at: u32) -> Option<&'a str> {
        match self.0 {
            FactsSource::Package(package) => package.reference_moved_by_rows(sheet, at),
            FactsSource::Facts(facts) => stranded(facts, |area| {
                area.sheet.eq_ignore_ascii_case(sheet) && area.end.row >= at
            }),
        }
    }

    pub fn reference_moved_by_cols(&self, sheet: &str, at: u32) -> Option<&'a str> {
        match self.0 {
            FactsSource::Package(package) => package.reference_moved_by_cols(sheet, at),
            FactsSource::Facts(facts) => stranded(facts, |area| {
                area.sheet.eq_ignore_ascii_case(sheet) && area.end.col >= at
            }),
        }
    }

    pub fn chart_part_bytes(&self, path: &str) -> Option<&'a [u8]> {
        match self.0 {
            FactsSource::Package(package) => package.part_bytes(path),
            FactsSource::Facts(facts) => facts
                .charts
                .iter()
                .find(|(part, _)| part.trim_start_matches('/') == path.trim_start_matches('/'))
                .map(|(_, bytes)| bytes.as_slice()),
        }
    }
}

fn stranded(facts: &PackageFacts, disturbed: impl Fn(&ReferenceAreaFacts) -> bool) -> Option<&str> {
    facts
        .references
        .iter()
        .find(|reference| match &reference.areas {
            None => true,
            Some(areas) => areas.iter().any(&disturbed),
        })
        .map(|reference| reference.part.as_str())
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;

    pub(crate) fn package(
        protected: bool,
        uncached: bool,
        chart: bool,
        references: bool,
    ) -> PreservedPackage {
        let mut parts =
            vec![
            (
                "xl/workbook.xml".to_owned(),
                br#"<workbook
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                    <sheets><sheet name="Data" sheetId="1" r:id="rId1"/>
                    <sheet name="Report" sheetId="2" state="hidden" r:id="rId2"/></sheets>
                    </workbook>"#
                    .to_vec(),
            ),
            (
                "xl/_rels/workbook.xml.rels".to_owned(),
                br#"<Relationships
                    xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship Id="rId1" Target="worksheets/sheet1.xml"
Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/>
                    <Relationship Id="rId2" Target="chartsheets/sheet2.xml"
Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"/>
                    </Relationships>"#
                    .to_vec(),
            ),
            ("xl/chartsheets/sheet2.xml".to_owned(), b"<chartsheet/>".to_vec()),
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                    <Default Extension="xml" ContentType="application/xml"/>
                    <Default Extension="rels"
                    ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                    </Types>"#
                    .to_vec(),
            ),
        ];
        let formula = if uncached {
            "<c r=\"A1\"><f>1+1</f></c>"
        } else {
            ""
        };
        let protection = if protected {
            "<sheetProtection sheet=\"1\"/>"
        } else {
            ""
        };
        let drawing = if chart {
            "<drawing r:id=\"draw\"/>"
        } else {
            ""
        };
        parts.push((
            "xl/worksheets/sheet1.xml".to_owned(),
            format!(
                r#"<worksheet
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheetData><row r="1">{formula}</row></sheetData>
                {protection}{drawing}</worksheet>"#,
            )
            .into_bytes(),
        ));
        if chart {
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
                    br#"<c:chartSpace
                        xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                        <c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/>
                        <c:val><c:numRef><c:f>Data!$A$1</c:f><c:numCache>
                        <c:pt idx="0"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val>
                        </c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#
                        .to_vec(),
                ),
            ]);
        }
        if references {
            parts.extend([
                (
                    "xl/pivotcache/pivotCacheDefinition1.xml".to_owned(),
                    br#"<pivotCacheDefinition><cacheSource>
                        <worksheetSource sheet="Data" ref="A1:B4"/>
                        </cacheSource></pivotCacheDefinition>"#
                        .to_vec(),
                ),
                (
                    "xl/pivotcache/pivotCacheDefinition2.xml".to_owned(),
                    br#"<pivotCacheDefinition><cacheSource type="external"/>
                        </pivotCacheDefinition>"#
                        .to_vec(),
                ),
            ]);
        }
        crate::parse_workbook_with_owned_package(parts)
            .unwrap()
            .package
    }

    #[test]
    fn facts_match_package_queries_and_missing_defaults() {
        for flags in [
            (false, false, false, false),
            (true, false, false, false),
            (false, true, false, false),
            (false, false, true, false),
            (false, false, false, true),
            (true, true, true, true),
        ] {
            let package = package(flags.0, flags.1, flags.2, flags.3);
            let facts = PackageFacts::from_package(&package);
            let present = PackageFactsView::from_package(&package);
            let deferred = PackageFactsView::from_facts(&facts);
            assert_eq!(present.source_present(), deferred.source_present());
            assert_eq!(present.source_sheet_count(), deferred.source_sheet_count());
            assert_eq!(present.has_uncached_source_formulas(), flags.1);
            assert_eq!(
                present.has_uncached_source_formulas(),
                deferred.has_uncached_source_formulas()
            );
            for index in [0, 1, 2, usize::MAX] {
                assert_eq!(
                    present.source_sheet_part(index),
                    deferred.source_sheet_part(index)
                );
                assert_eq!(
                    present.source_sheet_kind(index),
                    deferred.source_sheet_kind(index)
                );
                assert_eq!(
                    present.source_sheet_visibility(index),
                    deferred.source_sheet_visibility(index)
                );
                assert_eq!(
                    present.source_sheet_is_worksheet(index),
                    deferred.source_sheet_is_worksheet(index)
                );
                assert_eq!(
                    present.source_sheet_is_protected(index),
                    deferred.source_sheet_is_protected(index)
                );
            }
            assert_eq!(present.source_sheet_is_protected(0), flags.0);
            assert_eq!(
                present.unpatchable_reference_part(),
                deferred.unpatchable_reference_part()
            );
            for name in ["Data", "dAtA", "Report", "Missing", ""] {
                assert_eq!(
                    present.reference_naming_sheet(name),
                    deferred.reference_naming_sheet(name)
                );
                for at in [0, 1, 3, 4, u32::MAX] {
                    assert_eq!(
                        present.reference_moved_by_rows(name, at),
                        deferred.reference_moved_by_rows(name, at)
                    );
                    assert_eq!(
                        present.reference_moved_by_cols(name, at),
                        deferred.reference_moved_by_cols(name, at)
                    );
                }
            }
            for path in ["missing", "xl/charts/chart1.xml", "/xl/charts/chart1.xml"] {
                assert_eq!(
                    present.chart_part_bytes(path),
                    deferred.chart_part_bytes(path)
                );
            }
            if flags.2 {
                assert!(!facts.charts.is_empty());
            }
            if flags.3 {
                assert!(
                    facts
                        .references
                        .iter()
                        .any(|reference| reference.areas.is_some())
                );
                assert!(
                    facts
                        .references
                        .iter()
                        .any(|reference| reference.areas.is_none())
                );
            }
        }
    }
}
