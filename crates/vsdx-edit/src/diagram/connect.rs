use std::collections::{BTreeMap, HashSet};

use vsdx_parse::{Connect, ConnectsChild, ShapesChild, SheetChild};
use vsdx_resolve::Resolver;
use yrs::{Doc, Map, MapPrelim, Out, ReadTxn, Transact, WriteTxn};

use super::{
    ShapeOrigin, insert_shape, largest_shape_id, map_ref, map_string, materialize_shape,
    required_map, required_string, shape_from_snapshot, shape_origin, validate_shape_draft,
};
use crate::{
    CONNECTS, ConnectorGlue, DiagramSession, EditCtx, EditError, EditResult, PAGES, PageSnapshot,
    SHEETS, ShapeDraft, ShapeReceipt, ShapeSnapshot,
};

#[derive(PartialEq, Eq)]
pub(super) struct GlueRecord {
    id: String,
    page_id: String,
    connector_id: String,
    endpoint: GlueEndpoint,
    target_id: String,
    to_cell: String,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum GlueEndpoint {
    Begin,
    End,
}

impl GlueEndpoint {
    fn name(self) -> &'static str {
        match self {
            Self::Begin => "begin",
            Self::End => "end",
        }
    }

    fn endpoint_cell(self) -> &'static str {
        match self {
            Self::Begin => "BeginX",
            Self::End => "EndX",
        }
    }

    fn parse(value: &str) -> Option<Self> {
        match value {
            "begin" => Some(Self::Begin),
            "end" => Some(Self::End),
            _ => None,
        }
    }
}

/// One glue record addressed by the source IDs a projection assigns.
pub(super) struct PageGlue {
    pub from_sheet: u32,
    pub from_cell: &'static str,
    pub to_sheet: u32,
    pub to_cell: String,
}

impl GlueRecord {
    fn project(&self, sources: &BTreeMap<&str, u32>) -> Option<PageGlue> {
        Some(PageGlue {
            from_sheet: *sources.get(self.connector_id.as_str())?,
            from_cell: self.endpoint.endpoint_cell(),
            to_sheet: *sources.get(self.target_id.as_str())?,
            to_cell: self.to_cell.clone(),
        })
    }
}

impl DiagramSession {
    /// Adds a connector and its glue in one transaction.
    pub fn add_connector(
        &self,
        context: &EditCtx,
        page_id: &str,
        draft: &ShapeDraft,
        from: &ConnectorGlue,
        to: &ConnectorGlue,
    ) -> EditResult<ShapeReceipt> {
        validate_shape_draft(draft)?;
        let glue = [(GlueEndpoint::Begin, from), (GlueEndpoint::End, to)];
        self.validate_connector(page_id, draft, &glue)?;
        let mut txn = self.transact_for(context);
        let receipt = insert_shape(&mut txn, self.client_id, page_id, draft)?;
        let connects = txn.get_or_insert_map(CONNECTS);
        for (endpoint, target) in glue {
            let key = format!("{}:{}", receipt.shape_id, endpoint.name());
            let entry = connects.insert(&mut txn, key.as_str(), MapPrelim::default());
            entry.insert(&mut txn, "id", key.as_str());
            entry.insert(&mut txn, "pageId", page_id);
            entry.insert(&mut txn, "connectorId", receipt.shape_id.as_str());
            entry.insert(&mut txn, "endpoint", endpoint.name());
            entry.insert(&mut txn, "targetId", target.shape_id.as_str());
            entry.insert(
                &mut txn,
                "toCell",
                target.to_cell.as_deref().unwrap_or("PinX"),
            );
        }
        Ok(receipt)
    }

    fn validate_connector(
        &self,
        page_id: &str,
        draft: &ShapeDraft,
        glue: &[(GlueEndpoint, &ConnectorGlue); 2],
    ) -> EditResult<()> {
        let mut package = self.package()?;
        let snapshot = self.snapshot()?;
        let page = snapshot
            .pages
            .iter()
            .find(|page| page.id == page_id)
            .ok_or_else(|| EditError::InvalidState("connector page does not exist".to_owned()))?;
        let sources = snapshot_shape_sources(page);
        let sheet = package
            .page_contents
            .get_mut(&page.source_part_path)
            .ok_or_else(|| {
                EditError::InvalidState("connector page part does not exist".to_owned())
            })?;
        let source_id = largest_shape_id(sheet).checked_add(1).ok_or_else(|| {
            EditError::InvalidState("cannot allocate a connector source ID".to_owned())
        })?;
        let candidate = ShapeSnapshot {
            id: String::new(),
            source_id,
            name: draft.name.clone(),
            cells: draft.cells.clone(),
            children: Vec::new(),
        };
        let mut shape = shape_from_snapshot(&candidate);
        materialize_shape(&mut shape, &candidate, &HashSet::new(), 1)?;
        let Some(SheetChild::Shapes(shapes)) = sheet
            .children
            .iter_mut()
            .find(|child| matches!(child, SheetChild::Shapes(_)))
        else {
            return Err(EditError::InvalidState(
                "connector page has no shapes".to_owned(),
            ));
        };
        shapes.push(ShapesChild::Shape(shape));
        let mut connects = Vec::new();
        for (endpoint, target) in glue {
            let to_sheet = sources
                .get(target.shape_id.as_str())
                .ok_or_else(|| EditError::ShapeNotFound(target.shape_id.clone()))?;
            connects.push(ConnectsChild::Connect(Connect {
                from_sheet: source_id,
                from_cell: Some(endpoint.endpoint_cell().to_owned()),
                from_part: None,
                to_sheet: *to_sheet,
                to_cell: Some(target.to_cell.as_deref().unwrap_or("PinX").to_owned()),
                to_part: None,
                other_attrs: Vec::new(),
            }));
        }
        sheet.children.push(SheetChild::Connects(connects));
        let connectivity = Resolver::new(&package)
            .resolve_page_connectivity(&page.source_part_path)
            .map_err(|error| EditError::InvalidState(error.to_string()))?;
        let connector = connectivity
            .connectors
            .get(&source_id)
            .ok_or_else(|| EditError::InvalidState("connector draft did not resolve".to_owned()))?;
        if !connector.is_1d {
            return Err(EditError::InvalidState(
                "connector draft must describe a 1D shape".to_owned(),
            ));
        }
        if connector.glue.iter().any(|glue| {
            glue.to
                .as_ref()
                .and_then(|target| target.connection_point.as_ref())
                .is_none()
        }) {
            return Err(EditError::InvalidState(
                "connector glue must resolve to a connection point".to_owned(),
            ));
        }
        Ok(())
    }
}

pub(super) fn glue_records<T: ReadTxn>(txn: &T) -> EditResult<Vec<GlueRecord>> {
    let Some(connects) = txn.get_map(CONNECTS) else {
        return Ok(Vec::new());
    };
    let pages = required_map(txn, PAGES)?;
    let sheets = required_map(txn, SHEETS)?;
    let mut records = Vec::with_capacity(connects.len(txn) as usize);
    for (key, value) in connects.iter(txn) {
        let Out::YMap(entry) = value else {
            return Err(EditError::InvalidState(
                "connector glue is not a map".to_owned(),
            ));
        };
        let endpoint = GlueEndpoint::parse(&required_string(&entry, txn, "endpoint")?)
            .ok_or_else(|| EditError::InvalidState("connector glue has no endpoint".to_owned()))?;
        let record = GlueRecord {
            id: required_string(&entry, txn, "id")?,
            page_id: required_string(&entry, txn, "pageId")?,
            connector_id: required_string(&entry, txn, "connectorId")?,
            endpoint,
            target_id: required_string(&entry, txn, "targetId")?,
            to_cell: required_string(&entry, txn, "toCell")?,
        };
        if record.id != key || key != format!("{}:{}", record.connector_id, endpoint.name()) {
            return Err(EditError::InvalidState(
                "connector glue ID does not match its endpoint".to_owned(),
            ));
        }
        map_ref(&pages, txn, &record.page_id)?;
        for id in [&record.connector_id, &record.target_id] {
            if let Some(Out::YMap(shape)) = sheets.get(txn, id) {
                if map_string(&shape, txn, "pageId").as_deref() != Some(&record.page_id) {
                    return Err(EditError::InvalidState(
                        "connector glue crosses pages".to_owned(),
                    ));
                }
                if id == &record.connector_id && shape_origin(&shape, txn)? != ShapeOrigin::Added {
                    return Err(EditError::InvalidState(
                        "connector glue needs a session connector".to_owned(),
                    ));
                }
            }
        }
        records.push(record);
    }
    records.sort_by(|left, right| left.id.cmp(&right.id));
    Ok(records)
}

fn snapshot_shape_sources(page: &PageSnapshot) -> BTreeMap<&str, u32> {
    let mut sources = BTreeMap::new();
    let mut pending = page.shapes.iter().collect::<Vec<_>>();
    while let Some(shape) = pending.pop() {
        sources.insert(shape.id.as_str(), shape.source_id);
        pending.extend(shape.children.iter());
    }
    sources
}

/// Omits glue whose endpoints were deleted, including by concurrent edits.
pub(super) fn page_glue(page: &PageSnapshot, glue: &[GlueRecord]) -> Vec<PageGlue> {
    let sources = snapshot_shape_sources(page);
    glue.iter()
        .filter(|record| record.page_id == page.id)
        .filter_map(|record| record.project(&sources))
        .collect()
}

pub(super) fn materialize_page_glue(
    sheet: &mut vsdx_parse::Sheet,
    page: &PageSnapshot,
    glue: &[GlueRecord],
) {
    let pending = page_glue(page, glue);
    if pending.is_empty() {
        return;
    }
    let pending = pending.into_iter().map(|glue| {
        ConnectsChild::Connect(Connect {
            from_sheet: glue.from_sheet,
            from_cell: Some(glue.from_cell.to_owned()),
            from_part: None,
            to_sheet: glue.to_sheet,
            to_cell: Some(glue.to_cell),
            to_part: None,
            other_attrs: Vec::new(),
        })
    });
    match sheet.children.iter_mut().find_map(|child| match child {
        SheetChild::Connects(connects) => Some(connects),
        _ => None,
    }) {
        Some(connects) => connects.extend(pending),
        None => sheet.children.push(SheetChild::Connects(pending.collect())),
    }
}

pub(super) fn validate_remote_glue(before: &Doc, staged: &Doc) -> EditResult<()> {
    let before_glue = glue_records(&before.transact())?;
    let txn = staged.transact();
    let after_glue = glue_records(&txn)?
        .into_iter()
        .map(|record| (record.id.clone(), record))
        .collect::<BTreeMap<_, _>>();
    let sheets = required_map(&txn, SHEETS)?;
    for record in before_glue {
        match after_glue.get(&record.id) {
            Some(after) if after == &record => {}
            Some(_) => {
                return Err(EditError::InvalidState(format!(
                    "remote update changes connector glue {}",
                    record.id
                )));
            }
            None if sheets.contains_key(&txn, &record.connector_id)
                && sheets.contains_key(&txn, &record.target_id) =>
            {
                return Err(EditError::InvalidState(format!(
                    "remote update removes connector glue {} while its shapes survive",
                    record.id
                )));
            }
            None => {}
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::cell::RefCell;

    use vsdx_parse::{CellLocator, CellRow, CellSheet};

    use super::*;
    use crate::CellSnapshot;

    fn singleton_cell(name: &str, formula: &str) -> CellSnapshot {
        CellSnapshot {
            row_type: None,
            locator: CellLocator {
                sheet: CellSheet::Page(1),
                shape_id: None,
                section: None,
                section_index: None,
                row: None,
                cell_name: name.to_owned(),
            },
            name: name.to_owned(),
            formula: Some(formula.to_owned()),
            value: None,
        }
    }

    fn connection_cell(name: &str, formula: &str) -> CellSnapshot {
        CellSnapshot {
            row_type: None,
            locator: CellLocator {
                sheet: CellSheet::Page(1),
                shape_id: None,
                section: Some("Connection".to_owned()),
                section_index: None,
                row: Some(CellRow::Index(0)),
                cell_name: name.to_owned(),
            },
            name: name.to_owned(),
            formula: Some(formula.to_owned()),
            value: None,
        }
    }

    fn rect_draft(pin_x: &str, pin_y: &str) -> ShapeDraft {
        ShapeDraft {
            name: None,
            cells: [
                ("Width", "1"),
                ("Height", "1"),
                ("PinX", pin_x),
                ("PinY", pin_y),
                ("LocPinX", "0"),
                ("LocPinY", "0"),
            ]
            .into_iter()
            .map(|(name, formula)| singleton_cell(name, formula))
            .collect(),
        }
    }

    fn connector_draft() -> ShapeDraft {
        ShapeDraft {
            name: Some("Connector".to_owned()),
            cells: [
                ("OneD", "1"),
                ("BeginX", "1"),
                ("BeginY", "2"),
                ("EndX", "4"),
                ("EndY", "2"),
            ]
            .into_iter()
            .map(|(name, formula)| singleton_cell(name, formula))
            .collect(),
        }
    }

    fn glued_fixture() -> (DiagramSession, String, String, ShapeReceipt) {
        let session = DiagramSession::open(
            include_bytes!("../../../vsdx-parse/tests/fixtures/foundation.vsdx"),
            701,
        )
        .unwrap();
        let context = EditCtx::local("connector");
        let from = session
            .add_shape(&context, "page:1", &rect_draft("1", "1"))
            .unwrap();
        let to = session
            .add_shape(&context, "page:1", &rect_draft("5", "1"))
            .unwrap();
        let connector = session
            .add_connector(
                &context,
                "page:1",
                &connector_draft(),
                &ConnectorGlue {
                    shape_id: from.shape_id.clone(),
                    to_cell: None,
                },
                &ConnectorGlue {
                    shape_id: to.shape_id.clone(),
                    to_cell: None,
                },
            )
            .unwrap();
        (session, from.shape_id, to.shape_id, connector)
    }

    fn page_part(session: &DiagramSession) -> String {
        session.package().unwrap().page_part_paths[0].clone()
    }

    fn connector_route_debug(session: &DiagramSession, source_id: u32) -> String {
        let package = session.package().unwrap();
        let part = package.page_part_paths[0].clone();
        let list = vsdx_render::Renderer::default()
            .layout_page(&package, &part)
            .unwrap();
        let id = format!("{part}:{source_id}");
        list.primitives
            .iter()
            .find_map(|primitive| match primitive {
                vsdx_render::Primitive::Shape {
                    id: actual, path, ..
                } if actual == &id => Some(format!("{path:?}")),
                _ => None,
            })
            .unwrap_or_else(|| panic!("connector {source_id} did not paint"))
    }

    #[test]
    fn connector_creation_binds_both_endpoints() {
        let (session, _, _, _) = glued_fixture();
        let package = session.package().unwrap();
        let part = package.page_part_paths[0].clone();
        let connectivity = vsdx_resolve::Resolver::new(&package)
            .resolve_page_connectivity(&part)
            .unwrap();
        let connector = connectivity.connectors.get(&4).unwrap();
        assert!(connector.is_1d);
        assert_eq!(connector.glue.len(), 2);
        let begin = connector
            .glue
            .iter()
            .find(|glue| glue.endpoint == vsdx_resolve::ConnectorEndpoint::Begin)
            .unwrap();
        let end = connector
            .glue
            .iter()
            .find(|glue| glue.endpoint == vsdx_resolve::ConnectorEndpoint::End)
            .unwrap();
        assert_eq!(begin.to.as_ref().unwrap().shape_id, 2);
        assert_eq!(end.to.as_ref().unwrap().shape_id, 3);
        assert!(begin.to.as_ref().unwrap().connection_point.is_some());
        assert!(end.to.as_ref().unwrap().connection_point.is_some());
        assert!(begin.diagnostics.is_empty());
        assert!(end.diagnostics.is_empty());
    }

    #[test]
    fn connector_glue_to_a_connection_row_resolves() {
        let session = DiagramSession::open(
            include_bytes!("../../../vsdx-parse/tests/fixtures/foundation.vsdx"),
            702,
        )
        .unwrap();
        let context = EditCtx::local("connector");
        let mut target_cells = rect_draft("2", "3").cells;
        target_cells.push(connection_cell("X", "0.5"));
        target_cells.push(connection_cell("Y", "0.5"));
        let target = session
            .add_shape(
                &context,
                "page:1",
                &ShapeDraft {
                    name: None,
                    cells: target_cells,
                },
            )
            .unwrap();
        let other = session
            .add_shape(&context, "page:1", &rect_draft("5", "1"))
            .unwrap();
        session
            .add_connector(
                &context,
                "page:1",
                &connector_draft(),
                &ConnectorGlue {
                    shape_id: target.shape_id.clone(),
                    to_cell: Some("Connections.X1".to_owned()),
                },
                &ConnectorGlue {
                    shape_id: other.shape_id.clone(),
                    to_cell: None,
                },
            )
            .unwrap();
        let package = session.package().unwrap();
        let part = package.page_part_paths[0].clone();
        let connectivity = vsdx_resolve::Resolver::new(&package)
            .resolve_page_connectivity(&part)
            .unwrap();
        let point = connectivity.connectors[&4].glue[0]
            .to
            .as_ref()
            .unwrap()
            .connection_point
            .as_ref()
            .unwrap();
        assert_eq!(point.row, 0);
        assert_eq!(point.position, vsdx_resolve::ScenePoint { x: 2.5, y: 3.5 });
    }

    #[test]
    fn moved_glued_shape_drags_the_connector_route() {
        let (session, from_id, _, _) = glued_fixture();
        assert_eq!(
            connector_route_debug(&session, 4),
            "[Move { x: 1.0, y: 1.0 }, Line { x: 5.0, y: 1.0 }]"
        );
        session
            .move_shape(&EditCtx::local("move"), "page:1", &from_id, "8", "3")
            .unwrap();
        assert_eq!(
            connector_route_debug(&session, 4),
            "[Move { x: 8.0, y: 3.0 }, Line { x: 5.0, y: 1.0 }]"
        );
    }

    #[test]
    fn deleting_the_connector_removes_its_glue() {
        let (session, _, _, connector) = glued_fixture();
        session
            .delete_shape(&EditCtx::local("delete"), "page:1", &connector.shape_id)
            .unwrap();
        let package = session.package().unwrap();
        let part = package.page_part_paths[0].clone();
        assert!(
            package.page_contents[&part]
                .connects()
                .all(|connect| connect.from_sheet != 4 && connect.to_sheet != 4)
        );
        let connectivity = vsdx_resolve::Resolver::new(&package)
            .resolve_page_connectivity(&part)
            .unwrap();
        assert!(!connectivity.connectors.contains_key(&4));
        let saved = vsdx_parse::parse_vsdx(&session.save().unwrap()).unwrap();
        assert_eq!(
            package.page_contents[&part].connects().collect::<Vec<_>>(),
            saved.page_contents[&part].connects().collect::<Vec<_>>()
        );
    }

    #[test]
    fn deleting_a_glued_shape_leaves_no_dangling_glue() {
        let (session, from_id, _, _) = glued_fixture();
        session
            .delete_shape(&EditCtx::local("delete"), "page:1", &from_id)
            .unwrap();
        let package = session.package().unwrap();
        let part = package.page_part_paths[0].clone();
        let mut live_ids = std::collections::HashSet::new();
        let mut pending = package.page_contents[&part].shapes().collect::<Vec<_>>();
        while let Some(shape) = pending.pop() {
            live_ids.insert(shape.id);
            pending.extend(shape.shapes());
        }
        assert!(package.page_contents[&part].connects().all(|connect| {
            live_ids.contains(&connect.from_sheet) && live_ids.contains(&connect.to_sheet)
        }));
        let connectivity = vsdx_resolve::Resolver::new(&package)
            .resolve_page_connectivity(&part)
            .unwrap();
        assert!(connectivity.diagnostics.iter().all(|diagnostic| !matches!(
            diagnostic,
            vsdx_resolve::ConnectivityDiagnostic::MissingToShape { .. }
                | vsdx_resolve::ConnectivityDiagnostic::MissingFromShape { .. }
        )));
        let glue = &connectivity.connectors[&3].glue;
        assert_eq!(glue.len(), 1);
        assert_eq!(glue[0].endpoint, vsdx_resolve::ConnectorEndpoint::End);
        assert_eq!(glue[0].to.as_ref().unwrap().shape_id, 2);
        assert_eq!(
            connector_route_debug(&session, 3),
            "[Move { x: 1.0, y: 2.0 }, Line { x: 5.0, y: 1.0 }]"
        );
    }

    #[test]
    fn saved_connector_reloads_identically() {
        let (session, from, _, _) = glued_fixture();
        assert_saved_connector_projection(&session, 4);
        session
            .delete_shape(&EditCtx::local("delete"), "page:1", &from)
            .unwrap();
        assert_saved_connector_projection(&session, 3);
    }

    fn assert_saved_connector_projection(session: &DiagramSession, source_id: u32) {
        let part = page_part(session);
        let live_package = session.package().unwrap();
        let live_connectivity = vsdx_resolve::Resolver::new(&live_package)
            .resolve_page_connectivity(&part)
            .unwrap();
        let saved = session.save().unwrap();
        assert_eq!(session.save().unwrap(), saved);
        let reparsed = vsdx_parse::parse_vsdx(&saved).unwrap();
        let mut untouched = vec![live_package.document_part_path.clone()];
        untouched.extend(live_package.pages_part_path.clone());
        untouched.extend(live_package.masters_part_path.clone());
        untouched.extend(live_package.windows_part_path.clone());
        untouched.extend(live_package.theme_part_paths.clone());
        untouched.extend(
            live_package
                .page_part_paths
                .iter()
                .filter(|path| *path != &part)
                .cloned(),
        );
        untouched.extend(live_package.master_part_paths.clone());
        for path in untouched {
            assert_eq!(
                reparsed.part_bytes(&path),
                live_package.part_bytes(&path),
                "untouched part changed: {path}"
            );
        }
        let page_xml = String::from_utf8(reparsed.part_bytes(&part).unwrap().to_vec()).unwrap();
        assert!(page_xml.contains("Mystery='yes'"));
        assert!(page_xml.contains("<UnknownConnect Flag='yes'>"));
        let reopened = DiagramSession::open(&saved, 703).unwrap();
        let reopened_package = reopened.package().unwrap();
        let live_connects = live_package.page_contents[&part]
            .connects()
            .map(|connect| {
                (
                    connect.from_sheet,
                    connect.from_cell.clone(),
                    connect.to_sheet,
                    connect.to_cell.clone(),
                )
            })
            .collect::<Vec<_>>();
        let reopened_connects = reopened_package.page_contents[&part]
            .connects()
            .map(|connect| {
                (
                    connect.from_sheet,
                    connect.from_cell.clone(),
                    connect.to_sheet,
                    connect.to_cell.clone(),
                )
            })
            .collect::<Vec<_>>();
        assert_eq!(live_connects, reopened_connects);
        let connector_glue = live_connects
            .iter()
            .filter(|connect| connect.0 == source_id)
            .cloned()
            .collect::<Vec<_>>();
        let expected = if source_id == 4 {
            vec![
                (4, Some("BeginX".to_owned()), 2, Some("PinX".to_owned())),
                (4, Some("EndX".to_owned()), 3, Some("PinX".to_owned())),
            ]
        } else {
            vec![(3, Some("EndX".to_owned()), 2, Some("PinX".to_owned()))]
        };
        assert_eq!(connector_glue, expected);
        assert_eq!(
            vsdx_resolve::Resolver::new(&reopened_package)
                .resolve_page_connectivity(&part)
                .unwrap(),
            live_connectivity
        );
        let renderer = vsdx_render::Renderer::default();
        assert_eq!(
            renderer.layout_page(&reopened_package, &part).unwrap(),
            renderer.layout_page(&live_package, &part).unwrap()
        );
        let reopened_connector = reopened.snapshot().unwrap().pages[0]
            .shapes
            .iter()
            .find(|shape| shape.source_id == source_id)
            .cloned()
            .expect("saved connector must reload");
        assert_eq!(reopened_connector.name.as_deref(), Some("Connector"));
        for (name, formula) in [
            ("OneD", "1"),
            ("BeginX", "1"),
            ("BeginY", "2"),
            ("EndX", "4"),
            ("EndY", "2"),
        ] {
            assert!(
                reopened_connector
                    .cells
                    .iter()
                    .any(|cell| { cell.name == name && cell.formula.as_deref() == Some(formula) })
            );
        }
    }

    #[test]
    fn refused_connector_creation_writes_neither_shape_nor_glue() {
        let (session, from_id, _, _) = glued_fixture();
        let part = page_part(&session);
        let shapes = session.snapshot().unwrap().pages[0].shapes.len();
        let glue = session.package().unwrap().page_contents[&part]
            .connects()
            .count();
        let before = session.save().unwrap();
        let before_update = session.encode_state_as_update_v1();
        let context = EditCtx::local("refused");
        assert!(matches!(
            session.add_connector(
                &context,
                "page:1",
                &connector_draft(),
                &ConnectorGlue {
                    shape_id: "page:1:shape:missing".to_owned(),
                    to_cell: None,
                },
                &ConnectorGlue {
                    shape_id: from_id.clone(),
                    to_cell: None,
                },
            ),
            Err(EditError::ShapeNotFound(_))
        ));
        assert!(
            session
                .add_connector(
                    &context,
                    "page:1",
                    &ShapeDraft {
                        name: None,
                        cells: Vec::new(),
                    },
                    &ConnectorGlue {
                        shape_id: from_id.clone(),
                        to_cell: None,
                    },
                    &ConnectorGlue {
                        shape_id: from_id.clone(),
                        to_cell: None,
                    },
                )
                .is_err()
        );
        assert!(
            session
                .add_connector(
                    &context,
                    "page:1",
                    &connector_draft(),
                    &ConnectorGlue {
                        shape_id: from_id.clone(),
                        to_cell: Some("Width".to_owned()),
                    },
                    &ConnectorGlue {
                        shape_id: from_id.clone(),
                        to_cell: None,
                    },
                )
                .is_err()
        );
        assert_eq!(session.snapshot().unwrap().pages[0].shapes.len(), shapes);
        assert_eq!(
            session.package().unwrap().page_contents[&part]
                .connects()
                .count(),
            glue
        );
        assert_eq!(session.save().unwrap(), before);
        assert_eq!(session.encode_state_as_update_v1(), before_update);
    }

    #[test]
    fn remote_forged_glue_for_an_original_shape_is_rejected() {
        let session = DiagramSession::open(
            include_bytes!("../../../vsdx-parse/tests/fixtures/foundation.vsdx"),
            704,
        )
        .unwrap();
        let before = session.encode_state_as_update_v1();
        let peer = DiagramSession::open_from_update(&before, 705).unwrap();
        {
            let mut txn = peer.yrs_doc().transact_mut();
            let connects = txn.get_map(crate::CONNECTS).unwrap();
            let entry = connects.insert(&mut txn, "page:1:shape:1:begin", MapPrelim::default());
            entry.insert(&mut txn, "id", "page:1:shape:1:begin");
            entry.insert(&mut txn, "pageId", "page:1");
            entry.insert(&mut txn, "connectorId", "page:1:shape:1");
            entry.insert(&mut txn, "endpoint", "begin");
            entry.insert(&mut txn, "targetId", "page:1:shape:1");
            entry.insert(&mut txn, "toCell", "PinX");
        }
        let update = peer
            .encode_diff_v1(&session.encode_state_vector_v1())
            .unwrap();
        assert!(session.apply_update_v1(&update).is_err());
        assert_eq!(before, session.encode_state_as_update_v1());
    }

    #[test]
    fn remote_glue_strip_while_shapes_survive_is_rejected() {
        let (session, _, _, connector) = glued_fixture();
        let before = session.encode_state_as_update_v1();
        let peer = DiagramSession::open_from_update(&before, 706).unwrap();
        {
            let mut txn = peer.yrs_doc().transact_mut();
            let connects = txn.get_map(crate::CONNECTS).unwrap();
            connects.remove(&mut txn, &format!("{}:begin", connector.shape_id));
        }
        let update = peer
            .encode_diff_v1(&session.encode_state_vector_v1())
            .unwrap();
        assert!(session.apply_update_v1(&update).is_err());
        assert_eq!(before, session.encode_state_as_update_v1());
    }

    #[test]
    fn concurrent_connector_creation_and_target_deletion_converge() {
        for original_target in [false, true] {
            for reverse_delivery in [false, true] {
                let (left, from, to, old_connector) = glued_fixture();
                let (left, from, to, old_id) = if original_target {
                    (
                        DiagramSession::open(&left.save().unwrap(), 701).unwrap(),
                        "page:1:shape:2".to_owned(),
                        "page:1:shape:3".to_owned(),
                        "page:1:shape:4".to_owned(),
                    )
                } else {
                    (left, from, to, old_connector.shape_id)
                };
                left.delete_shape(&EditCtx::local("delete"), "page:1", &old_id)
                    .unwrap();
                let target = &from;
                let right =
                    DiagramSession::open_from_update(&left.encode_state_as_update_v1(), 707)
                        .unwrap();
                let connector = left
                    .add_connector(
                        &EditCtx::local("connector"),
                        "page:1",
                        &connector_draft(),
                        &ConnectorGlue {
                            shape_id: target.to_owned(),
                            to_cell: None,
                        },
                        &ConnectorGlue {
                            shape_id: to.clone(),
                            to_cell: None,
                        },
                    )
                    .unwrap();
                right
                    .delete_shape(&EditCtx::local("delete"), "page:1", target)
                    .unwrap();
                let add = left
                    .encode_diff_v1(&right.encode_state_vector_v1())
                    .unwrap();
                let delete = right
                    .encode_diff_v1(&left.encode_state_vector_v1())
                    .unwrap();
                if reverse_delivery {
                    right.apply_update_v1(&add).unwrap();
                    left.apply_update_v1(&delete).unwrap();
                } else {
                    left.apply_update_v1(&delete).unwrap();
                    right.apply_update_v1(&add).unwrap();
                }
                assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
                assert_eq!(left.save().unwrap(), right.save().unwrap());
                for session in [&left, &right] {
                    let live = session.package().unwrap();
                    let saved = session.save().unwrap();
                    assert_eq!(saved, session.save().unwrap());
                    let reopened = vsdx_parse::parse_vsdx(&saved).unwrap();
                    let path = &live.page_part_paths[0];
                    assert_eq!(
                        live.page_contents[path].connects().collect::<Vec<_>>(),
                        reopened.page_contents[path].connects().collect::<Vec<_>>()
                    );
                    let connectivity = vsdx_resolve::Resolver::new(&live)
                        .resolve_page_connectivity(path)
                        .unwrap();
                    assert!(connectivity.diagnostics.iter().all(|diagnostic| !matches!(
                        diagnostic,
                        vsdx_resolve::ConnectivityDiagnostic::MissingToShape { .. }
                            | vsdx_resolve::ConnectivityDiagnostic::MissingFromShape { .. }
                    )));
                    let snapshot = session.snapshot().unwrap();
                    let source = |id: &str| {
                        snapshot.pages[0]
                            .shapes
                            .iter()
                            .find(|shape| shape.id == id)
                            .unwrap()
                            .source_id
                    };
                    let (connector_id, target_id) = (source(&connector.shape_id), source(&to));
                    let glue = &connectivity.connectors[&connector_id].glue;
                    assert_eq!(glue.len(), 1);
                    assert_eq!(glue[0].endpoint, vsdx_resolve::ConnectorEndpoint::End);
                    assert_eq!(glue[0].to.as_ref().unwrap().shape_id, target_id);
                    assert_eq!(
                        connectivity,
                        vsdx_resolve::Resolver::new(&reopened)
                            .resolve_page_connectivity(path)
                            .unwrap()
                    );
                }
            }
        }
    }

    #[test]
    fn connector_creation_uses_resolved_one_d_values() {
        for (formula, endpoints, accepted) in [
            (Some("0"), true, false),
            (Some("1-1"), true, false),
            (None, true, true),
            (None, false, false),
            (Some("User.Unknown"), false, false),
            (Some("1"), false, true),
        ] {
            let (session, from, to, _) = glued_fixture();
            let mut draft = connector_draft();
            draft.cells[0].formula = formula.map(str::to_owned);
            if !endpoints {
                draft.cells.truncate(1);
            }
            let before = session.encode_state_as_update_v1();
            let result = session.add_connector(
                &EditCtx::local("connector"),
                "page:1",
                &draft,
                &ConnectorGlue {
                    shape_id: from,
                    to_cell: None,
                },
                &ConnectorGlue {
                    shape_id: to,
                    to_cell: None,
                },
            );
            assert_eq!(
                result.is_ok(),
                accepted,
                "formula={formula:?}, endpoints={endpoints}"
            );
            if !accepted {
                assert_eq!(before, session.encode_state_as_update_v1());
            }
        }
    }

    #[test]
    fn missing_or_unusable_connection_points_refuse_the_whole_edit() {
        for (cell, x, y) in [
            ("Connections.X9", Some("0.5"), Some("0.5")),
            ("Connections.X1", None, Some("0.5")),
            ("Connections.X1", Some("0.5"), None),
            ("Connections.X1", Some("User.Unknown"), Some("0.5")),
        ] {
            let (session, _, to, _) = glued_fixture();
            let mut draft = rect_draft("2", "3");
            for (name, formula) in [("X", x), ("Y", y)] {
                if let Some(formula) = formula {
                    draft.cells.push(connection_cell(name, formula));
                }
            }
            let from = session
                .add_shape(&EditCtx::local("shape"), "page:1", &draft)
                .unwrap();
            let before = session.encode_state_as_update_v1();
            assert!(
                session
                    .add_connector(
                        &EditCtx::local("connector"),
                        "page:1",
                        &connector_draft(),
                        &ConnectorGlue {
                            shape_id: from.shape_id,
                            to_cell: Some(cell.to_owned())
                        },
                        &ConnectorGlue {
                            shape_id: to,
                            to_cell: None
                        },
                    )
                    .is_err(),
                "{cell}: X={x:?}, Y={y:?}"
            );
            assert_eq!(before, session.encode_state_as_update_v1());
        }
    }

    #[test]
    fn deleted_connection_points_refuse_the_whole_edit() {
        for (section, row, cell) in [
            (" Del='1'", "", ""),
            ("", " Del='1'", ""),
            ("", "", " Del='1'"),
        ] {
            let package = vsdx_parse::parse_vsdx(include_bytes!(
                "../../../vsdx-parse/tests/fixtures/foundation.vsdx"
            ))
            .unwrap();
            let bytes = vsdx_parse::save_structural_edits(&package, &[vsdx_parse::StructuralEdit::AddShape {
                page_id: 1,
                shape_xml: format!("<Shape><Cell N='PinX' V='1'/><Cell N='PinY' V='1'/><Cell N='Width' V='1'/><Cell N='Height' V='1'/><Section N='Connection'{section}><Row IX='0'{row}><Cell N='X' V='0.5'{cell}/><Cell N='Y' V='0.5'/></Row></Section></Shape>").into_bytes(),
            }]).unwrap();
            let session = DiagramSession::open(&bytes, 708).unwrap();
            let before = session.encode_state_as_update_v1();
            assert!(
                session
                    .add_connector(
                        &EditCtx::local("connector"),
                        "page:1",
                        &connector_draft(),
                        &ConnectorGlue {
                            shape_id: "page:1:shape:2".to_owned(),
                            to_cell: Some("Connections.X1".to_owned())
                        },
                        &ConnectorGlue {
                            shape_id: "page:1:shape:2".to_owned(),
                            to_cell: None
                        },
                    )
                    .is_err()
            );
            assert_eq!(before, session.encode_state_as_update_v1());
        }
    }

    #[test]
    fn connector_creation_and_deletion_are_single_undoable_updates() {
        let (session, from, to, old_connector) = glued_fixture();
        session
            .delete_shape(&EditCtx::local("delete"), "page:1", &old_connector.shape_id)
            .unwrap();
        session.add_undo_barrier();
        let before = session.save().unwrap();
        let glue_before = glue_records(&session.yrs_doc().transact()).unwrap().len();
        let updates = std::rc::Rc::new(RefCell::new(Vec::new()));
        let observed = updates.clone();
        let _subscription = session
            .observe_update_v1(move |event| observed.borrow_mut().push(event))
            .unwrap();
        let receipt = session
            .add_connector(
                &EditCtx::local("connector"),
                "page:1",
                &connector_draft(),
                &ConnectorGlue {
                    shape_id: from.clone(),
                    to_cell: None,
                },
                &ConnectorGlue {
                    shape_id: to,
                    to_cell: None,
                },
            )
            .unwrap();
        assert_eq!(updates.borrow().len(), 1);
        let added = session.save().unwrap();
        assert!(session.undo());
        assert_eq!(session.save().unwrap(), before);
        assert_eq!(
            glue_records(&session.yrs_doc().transact()).unwrap().len(),
            glue_before
        );
        assert!(session.redo());
        assert_eq!(session.save().unwrap(), added);
        for target in [&receipt.shape_id, &from] {
            session.add_undo_barrier();
            updates.borrow_mut().clear();
            session
                .delete_shape(&EditCtx::local("delete"), "page:1", target)
                .unwrap();
            assert_eq!(updates.borrow().len(), 1);
            assert_ne!(session.save().unwrap(), added);
            assert!(session.undo());
            assert_eq!(session.save().unwrap(), added);
        }
    }
}
