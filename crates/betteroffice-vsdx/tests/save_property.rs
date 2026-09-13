use std::collections::BTreeMap;
use std::fmt::Write;
use std::panic::{AssertUnwindSafe, catch_unwind};

use betteroffice_vsdx::{
    CellLocator, CellRow, CellSheet, Diagram, MutationGesture, SemanticCellEdit,
};
use ooxml_opc::{rezip_parts, unzip_parts};
use vsdx_parse::{Cell, Shape, Sheet};
use vsdx_resolve::{Lookup, ResolvedShape, Resolver};

const CASES: usize = 256;
const SEED: u64 = 0x5A9E_5AFE_5A0E_0001;

struct Rng(u64);

impl Rng {
    fn new(seed: u64) -> Self {
        Self(seed)
    }

    fn next(&mut self) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0
    }

    fn choose(&mut self, upper: usize) -> usize {
        (self.next() % upper as u64) as usize
    }

    fn chance(&mut self, numerator: usize, denominator: usize) -> bool {
        self.choose(denominator) < numerator
    }
}

#[test]
fn save_cell_edits_preserve_spans_and_projections_property() {
    for case in 0..CASES {
        let result = catch_unwind(AssertUnwindSafe(|| run_save_case(case)));
        assert!(
            result.is_ok(),
            "save path panicked or violated an invariant; seed={SEED:#018X}; case={case}"
        );
    }
}

#[test]
fn refused_save_cell_edits_do_not_mutate_property() {
    for case in 0..CASES {
        let result = catch_unwind(AssertUnwindSafe(|| run_refusal_case(case)));
        assert!(
            result.is_ok(),
            "refusal path panicked or mutated its source; seed={SEED:#018X}; case={case}"
        );
    }
}

fn run_save_case(case: usize) {
    let generated = generated_case(case);
    let original = parts(&generated.source);
    let diagram = Diagram::open(&generated.source).unwrap_or_else(|error| {
        panic!("generated package did not open; seed={SEED:#018X}; case={case}; {error:?}")
    });
    let before_resolved = resolved_cells(&diagram, &generated.page_path, generated.shape_id);
    let target_span = generated
        .existing_target
        .then(|| {
            cell_spans(&generated.page_xml)
                .into_iter()
                .find(|span| has_name(span, &generated.edit.locator.cell_name))
        })
        .flatten()
        .unwrap_or_else(|| {
            if generated.existing_target {
                panic!("target cell span missing; seed={SEED:#018X}; case={case}");
            }
            String::new()
        });

    let saved = diagram
        .save_cell_edits(std::slice::from_ref(&generated.edit))
        .unwrap_or_else(|error| {
            panic!("valid generated edit was refused; seed={SEED:#018X}; case={case}; {error:?}")
        });
    let saved_parts = parts(&saved);
    assert_eq!(
        saved_parts.len(),
        original.len(),
        "save added or dropped a package part; seed={SEED:#018X}; case={case}"
    );
    for (path, bytes) in &original {
        if path == &generated.page_path {
            continue;
        }
        assert_eq!(
            saved_parts.get(path),
            Some(bytes),
            "untouched part changed: {path}; seed={SEED:#018X}; case={case}"
        );
    }

    let saved_page = std::str::from_utf8(&saved_parts[&generated.page_path]).unwrap();
    for span in cell_spans(&generated.page_xml) {
        if generated.existing_target && span == target_span {
            continue;
        }
        assert!(
            saved_page.contains(&span),
            "untouched Cell span changed: {span:?}; seed={SEED:#018X}; case={case}"
        );
    }

    let reopened = Diagram::open(&saved).unwrap_or_else(|error| {
        panic!("saved package did not reopen; seed={SEED:#018X}; case={case}; {error:?}")
    });
    let changed = find_cell(
        reopened
            .package()
            .page_contents
            .get(&generated.page_path)
            .unwrap(),
        generated.shape_id,
        &generated.edit.locator,
    )
    .unwrap_or_else(|| panic!("saved target cell missing; seed={SEED:#018X}; case={case}"));
    assert_eq!(
        changed.formula, generated.edit.formula,
        "formula projection diverged; seed={SEED:#018X}; case={case}"
    );
    assert_eq!(
        changed.value.as_deref(),
        generated.expected_cache.as_deref(),
        "cache projection diverged; seed={SEED:#018X}; case={case}"
    );
    let mut after_resolved = resolved_cells(&reopened, &generated.page_path, generated.shape_id);
    after_resolved.retain(|key, _| !key.ends_with(&generated.edit.locator.cell_name));
    let mut before_resolved = before_resolved;
    before_resolved.retain(|key, _| !key.ends_with(&generated.edit.locator.cell_name));
    assert_eq!(
        after_resolved, before_resolved,
        "unrelated resolved cells changed; seed={SEED:#018X}; case={case}"
    );
}

fn run_refusal_case(case: usize) {
    let kind = case % 3;
    let (xml, name, gesture, expected_reason) = match kind {
        0 => (
            "<PageContents><Shapes><Shape ID='1'><Cell Extra='α' V='1' N='Width' F='GUARD(1)'/></Shape></Shapes></PageContents>",
            "Width",
            MutationGesture::CellEdit,
            "GUARD",
        ),
        1 => (
            "<PageContents><Shapes><Shape ID='1'><Cell N='LockWidth' V='1' Unknown='🙂'/><Cell V='1' N='Width'/></Shape></Shapes></PageContents>",
            "Width",
            MutationGesture::ResizeWidth,
            "LockWidth",
        ),
        _ => (
            "<PageContents><Shapes><Shape ID='1'><Cell N='Width' F='SETATREF(Target)' V='1'/><Cell V='1' F='GUARD(1)' N='Target'/></Shape></Shapes></PageContents>",
            "Width",
            MutationGesture::CellEdit,
            "GUARD",
        ),
    };
    let source = package_with_page(xml);
    let original = parts(&source.bytes);
    let diagram = Diagram::open(&source.bytes).unwrap();
    let error = diagram
        .save_cell_edits(&[SemanticCellEdit {
            locator: CellLocator {
                sheet: CellSheet::Page(source.page_id),
                shape_id: Some(1),
                section: None,
                section_index: None,
                row: None,
                cell_name: name.to_owned(),
            },
            gesture,
            formula: Some("42".to_owned()),
            value: None,
        }])
        .expect_err("generated guarded, locked, or redirected edit must refuse");
    assert!(
        format!("{error:?}").contains(expected_reason),
        "refusal reason was not reported; seed={SEED:#018X}; case={case}; {error:?}"
    );
    for (path, bytes) in original {
        assert_eq!(
            diagram.package().part_bytes(&path),
            Some(bytes.as_slice()),
            "refusal changed source part {path}; seed={SEED:#018X}; case={case}"
        );
    }
}

struct GeneratedCase {
    source: Vec<u8>,
    page_path: String,
    page_xml: String,
    shape_id: u32,
    edit: SemanticCellEdit,
    existing_target: bool,
    expected_cache: Option<String>,
}

struct PagePackage {
    bytes: Vec<u8>,
    page_path: String,
    page_id: u32,
}

fn generated_case(case: usize) -> GeneratedCase {
    let mut rng = Rng::new(SEED ^ (case as u64).wrapping_mul(0x9E37_79B9_7F4A_7C15));
    let mut xml = String::from(
        "<PageContents Root='α🙂'><UnknownPage Flag=\"kept\"> β </UnknownPage><Shapes>",
    );
    let mut next_id = 1;
    let mut ids = Vec::new();
    for _ in 0..(1 + rng.choose(5)) {
        let depth = rng.choose(3);
        shape(&mut xml, &mut rng, &mut next_id, depth, &mut ids);
    }
    xml.push_str("</Shapes><UnknownTail Attr='γ'> tail 🙂 </UnknownTail></PageContents>");
    let package = package_with_page(&xml);
    let shape_id = ids[rng.choose(ids.len())];
    let kind = case % 4;
    let (section, row, cell_name, existing_target) = match kind {
        0 => (None, None, format!("Plain{shape_id}"), true),
        1 => (
            Some("User".to_owned()),
            Some(CellRow::Index(0)),
            format!("Row{shape_id}"),
            true,
        ),
        2 => (None, None, format!("Inserted{shape_id}"), false),
        _ => (
            Some("User".to_owned()),
            Some(CellRow::Index(99)),
            format!("InsertedRow{shape_id}"),
            false,
        ),
    };
    let formula = if rng.chance(1, 2) {
        "42".to_owned()
    } else {
        "Unknown(1)".to_owned()
    };
    let expected_cache = (formula == "42").then(|| "42".to_owned());
    GeneratedCase {
        source: package.bytes,
        page_path: package.page_path,
        page_xml: xml,
        shape_id,
        edit: SemanticCellEdit {
            locator: CellLocator {
                sheet: CellSheet::Page(package.page_id),
                shape_id: Some(shape_id),
                section,
                section_index: None,
                row,
                cell_name,
            },
            gesture: MutationGesture::CellEdit,
            formula: Some(formula),
            value: None,
        },
        existing_target,
        expected_cache,
    }
}

fn package_with_page(page_xml: &str) -> PagePackage {
    let fixture = include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx");
    let package = vsdx_parse::parse_vsdx(fixture).unwrap();
    let page_path = package.page_part_paths[0].clone();
    let page_id = package.page_part_ids[&page_path];
    let mut archive = unzip_parts(fixture).unwrap();
    archive
        .iter_mut()
        .find(|(path, _)| path == &page_path)
        .unwrap()
        .1 = page_xml.as_bytes().to_vec();
    PagePackage {
        bytes: rezip_parts(&archive).unwrap(),
        page_path,
        page_id,
    }
}

fn shape(output: &mut String, rng: &mut Rng, next_id: &mut u32, depth: usize, ids: &mut Vec<u32>) {
    let id = *next_id;
    *next_id += 1;
    ids.push(id);
    let q = quote(rng);
    let group = depth > 0 && rng.chance(3, 4);
    let kind = if group { "Group" } else { "Shape" };
    if rng.chance(1, 2) {
        let _ = write!(
            output,
            "<Shape Name='S α {id}' Weird={q}🙂-{id}{q} Type='{kind}' ID='{id}'>"
        );
    } else {
        let _ = write!(
            output,
            "<Shape ID='{id}' Type='{kind}' Weird={q}🙂-{id}{q} Name='S α {id}'>"
        );
    }
    output.push_str("\n <UnknownShape A='kept'> α 🙂 </UnknownShape>\n");
    local_cells(output, rng, id);
    section(output, rng, id);
    output.push_str("<Text> α <cp IX='0'/> 🙂 </Text>");
    if group {
        output.push_str("<Shapes>");
        for _ in 0..(1 + rng.choose(3)) {
            shape(output, rng, next_id, depth - 1, ids);
        }
        output.push_str("</Shapes>");
    }
    output.push_str("</Shape>");
}

fn local_cells(output: &mut String, rng: &mut Rng, id: u32) {
    let q = quote(rng);
    let deleted = if rng.chance(1, 3) { " Del='1'" } else { "" };
    if rng.chance(1, 2) {
        let _ = write!(
            output,
            " <Cell U='α' V={q}1{q} N='Plain{id}' Extra='plain'/><Cell N='FOnly{id}' F='Width*2' Weird='🙂'/><Cell N='VOnly{id}' V='5'{deleted}/><Cell F='2' V='2' N='Both{id}' Extra='both'/>"
        );
    } else {
        let _ = write!(
            output,
            " <Cell Extra='plain' N='Plain{id}' V={q}1{q} U='α'/><Cell Weird='🙂' F='Width*2' N='FOnly{id}'/><Cell{deleted} V='5' N='VOnly{id}'/><Cell Extra='both' N='Both{id}' V='2' F='2'/>"
        );
    }
}

fn section(output: &mut String, rng: &mut Rng, id: u32) {
    let q = quote(rng);
    let deleted = if rng.chance(1, 3) { " Del='1'" } else { "" };
    let _ = write!(
        output,
        "<Section Extra={q}α🙂{q} N='User'{deleted}><UnknownSection K='yes'> β </UnknownSection>"
    );
    for row in 0..(1 + rng.choose(3)) {
        let row_deleted = if rng.chance(1, 3) { " Del='1'" } else { "" };
        if rng.chance(1, 2) {
            let _ = write!(
                output,
                "<Row IX='{row}' LocalName='R α {row}'{row_deleted}><Cell Extra='row' V='1' N='Row{id}'/><Cell N='RowFOnly{id}' F='2'/></Row>"
            );
        } else {
            let _ = write!(
                output,
                "<Row LocalName='R α {row}'{row_deleted} IX='{row}'><Cell N='Row{id}' V='1' Extra='row'/><Cell F='2' N='RowFOnly{id}'/></Row>"
            );
        }
    }
    output.push_str("</Section>");
}

fn quote(rng: &mut Rng) -> char {
    if rng.chance(1, 2) { '\'' } else { '"' }
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    unzip_parts(bytes).unwrap().into_iter().collect()
}

fn cell_spans(xml: &str) -> Vec<String> {
    let mut spans = Vec::new();
    let mut offset = 0;
    while let Some(start) = xml[offset..].find("<Cell") {
        let start = offset + start;
        let end = start + xml[start..].find("/>").unwrap() + 2;
        spans.push(xml[start..end].to_owned());
        offset = end;
    }
    spans
}

fn has_name(span: &str, name: &str) -> bool {
    span.contains(&format!("N='{name}'")) || span.contains(&format!("N=\"{name}\""))
}

fn find_shape(sheet: &Sheet, id: u32) -> Option<&Shape> {
    sheet.shapes().find_map(|shape| find_shape_in(shape, id))
}

fn find_shape_in(shape: &Shape, id: u32) -> Option<&Shape> {
    if shape.id == id {
        return Some(shape);
    }
    shape.shapes().find_map(|child| find_shape_in(child, id))
}

fn find_cell<'a>(sheet: &'a Sheet, shape_id: u32, locator: &CellLocator) -> Option<&'a Cell> {
    let shape = find_shape(sheet, shape_id)?;
    match (&locator.section, &locator.row) {
        (Some(section), Some(row)) => shape
            .sections()
            .find(|candidate| candidate.name == *section)?
            .rows()
            .find(|candidate| match row {
                CellRow::Index(index) => candidate.index == Some(*index),
                CellRow::Name(name) => candidate.name.as_deref() == Some(name),
            })?
            .cells()
            .find(|cell| cell.name == locator.cell_name),
        _ => shape.cells().find(|cell| cell.name == locator.cell_name),
    }
}

fn resolved_cells(diagram: &Diagram, page_path: &str, shape_id: u32) -> BTreeMap<String, Lookup> {
    let resolved = Resolver::new(diagram.package())
        .resolve_shape(page_path, shape_id)
        .unwrap();
    flatten_resolved(&resolved)
}

fn flatten_resolved(shape: &ResolvedShape) -> BTreeMap<String, Lookup> {
    let mut result = shape.cells.clone();
    for section in shape.sections.values() {
        for row in section.rows.values() {
            for (name, cell) in &row.cells {
                result.insert(
                    format!("{}.{}.{}", section.name, row.key, name),
                    cell.clone(),
                );
            }
        }
    }
    result
}
