//! Shared formula groups keep every formula across structural edits, a save and a reopen.

use std::collections::{BTreeMap, BTreeSet};
use std::panic::{AssertUnwindSafe, catch_unwind};

use betteroffice_xlsx::{CalculationOptions, CellValue, Op, SheetId, Workbook};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const START: u32 = 10;
const LAG: u32 = 8;
const DATA: u32 = 8;

#[derive(Clone, Copy, PartialEq)]
enum Shape {
    Contiguous,
    OneGap,
    TwoGaps,
}

#[derive(Clone, Copy, PartialEq)]
enum Orient {
    Column,
    Row,
}

#[derive(Clone, Copy, PartialEq)]
enum MasterAt {
    First,
    Middle,
}

#[derive(Clone, Copy, PartialEq)]
enum Edit {
    InsertBefore,
    InsertInside,
    DeleteGap,
    DeleteFollower,
    DeleteMaster,
}

const SHAPES: [Shape; 3] = [Shape::Contiguous, Shape::OneGap, Shape::TwoGaps];
const ORIENTS: [Orient; 2] = [Orient::Column, Orient::Row];
const MASTERS: [MasterAt; 2] = [MasterAt::First, MasterAt::Middle];
const EDITS: [Edit; 5] = [
    Edit::InsertBefore,
    Edit::InsertInside,
    Edit::DeleteGap,
    Edit::DeleteFollower,
    Edit::DeleteMaster,
];

impl Shape {
    fn members(self) -> &'static [u32] {
        match self {
            Shape::Contiguous => &[10, 11, 12, 13],
            Shape::OneGap => &[10, 12, 13],
            Shape::TwoGaps => &[10, 12, 14],
        }
    }

    fn name(self) -> &'static str {
        match self {
            Shape::Contiguous => "contiguous",
            Shape::OneGap => "one-gap",
            Shape::TwoGaps => "two-gaps",
        }
    }

    fn master(self, at: MasterAt) -> u32 {
        match (at, self) {
            (MasterAt::First, _) => START,
            (MasterAt::Middle, Shape::Contiguous) => 11,
            (MasterAt::Middle, _) => 12,
        }
    }
}

impl Orient {
    fn name(self) -> &'static str {
        match self {
            Orient::Column => "column",
            Orient::Row => "row",
        }
    }

    fn member(self, pos: u32) -> String {
        match self {
            Orient::Column => format!("B{pos}"),
            Orient::Row => format!("{}2", letters(pos)),
        }
    }

    fn data(self, k: u32) -> String {
        match self {
            Orient::Column => format!("A{k}"),
            Orient::Row => format!("{}1", letters(k)),
        }
    }

    fn line(self, pos: u32) -> String {
        match self {
            Orient::Column => format!("row {pos}"),
            Orient::Row => format!("col {}", letters(pos)),
        }
    }

    fn op(self, insert: bool, pos: u32) -> Op {
        let sheet = SheetId(0);
        let at = pos - 1;
        let count = 1;
        match (self, insert) {
            (Orient::Column, true) => Op::InsertRows { sheet, at, count },
            (Orient::Column, false) => Op::DeleteRows { sheet, at, count },
            (Orient::Row, true) => Op::InsertCols { sheet, at, count },
            (Orient::Row, false) => Op::DeleteCols { sheet, at, count },
        }
    }
}

impl MasterAt {
    fn name(self) -> &'static str {
        match self {
            MasterAt::First => "first",
            MasterAt::Middle => "middle",
        }
    }
}

impl Edit {
    fn name(self) -> &'static str {
        match self {
            Edit::InsertBefore => "insert-before",
            Edit::InsertInside => "insert-inside",
            Edit::DeleteGap => "delete-gap",
            Edit::DeleteFollower => "delete-follower",
            Edit::DeleteMaster => "delete-master",
        }
    }
}

fn letters(mut n: u32) -> String {
    let mut out = Vec::new();
    while n > 0 {
        let d = (n - 1) % 26;
        out.push(b'A' + d as u8);
        n = (n - 1) / 26;
    }
    out.reverse();
    String::from_utf8(out).unwrap_or_default()
}

fn formula(orient: Orient, pos: u32) -> String {
    format!("{}*2", orient.data(pos - LAG))
}

fn sheet_xml(shape: Shape, orient: Orient, master_at: MasterAt) -> String {
    let members = shape.members();
    let master = shape.master(master_at);
    let end = *members.last().unwrap_or(&START);
    let reference = format!("{}:{}", orient.member(START), orient.member(end));
    let member_cell = |pos: u32| {
        let f = if pos == master {
            format!(
                r#"<f t="shared" ref="{reference}" si="0">{}</f>"#,
                formula(orient, pos)
            )
        } else {
            r#"<f t="shared" si="0"/>"#.to_owned()
        };
        format!(
            r#"<c r="{}">{f}<v>{}</v></c>"#,
            orient.member(pos),
            20 * (pos - LAG)
        )
    };
    let data_cell = |k: u32| format!(r#"<c r="{}"><v>{}</v></c>"#, orient.data(k), 10 * k);
    let mut rows = String::new();
    match orient {
        Orient::Column => {
            for k in 1..=DATA {
                rows.push_str(&format!(r#"<row r="{k}">{}</row>"#, data_cell(k)));
            }
            for &pos in members {
                rows.push_str(&format!(r#"<row r="{pos}">{}</row>"#, member_cell(pos)));
            }
        }
        Orient::Row => {
            let data: String = (1..=DATA).map(data_cell).collect();
            let cells: String = members.iter().map(|&pos| member_cell(pos)).collect();
            let row = format!(r#"<row r="1">{data}</row><row r="2">{cells}</row>"#);
            rows.push_str(&row);
        }
    }
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{MAIN}"><sheetData>{rows}</sheetData></worksheet>"#
    )
}

fn package(sheet: String) -> Result<Vec<u8>, String> {
    let parts = [
        (
            "[Content_Types].xml",
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
                r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
                r#"<Default Extension="xml" ContentType="application/xml"/>"#,
                r#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#,
                r#"<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#,
                r#"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
                r#"</Types>"#,
            )
            .to_owned(),
        ),
        (
            "_rels/.rels",
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
                r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>"#,
                r#"</Relationships>"#,
            )
            .to_owned(),
        ),
        (
            "xl/workbook.xml",
            format!(
                concat!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                    r#"<workbook xmlns="{main}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
                    r#"<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
                ),
                main = MAIN,
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            concat!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
                r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>"#,
                r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#,
                r#"</Relationships>"#,
            )
            .to_owned(),
        ),
        (
            "xl/styles.xml",
            format!(
                concat!(
                    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
                    r#"<styleSheet xmlns="{main}">"#,
                    r#"<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>"#,
                    r#"<fills count="2"><fill><patternFill patternType="none"/></fill>"#,
                    r#"<fill><patternFill patternType="gray125"/></fill></fills>"#,
                    r#"<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"#,
                    r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
                    r#"<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>"#,
                    r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#,
                    r#"</styleSheet>"#,
                ),
                main = MAIN,
            ),
        ),
        ("xl/worksheets/sheet1.xml", sheet),
    ]
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()));
    ooxml_opc::rezip_parts(&parts)
}

/// Positions along the edited axis as the edits move them.
struct Track {
    shape: Shape,
    members: Vec<Option<u32>>,
    master: usize,
    start: u32,
    end: u32,
}

impl Track {
    fn new(shape: Shape, master_at: MasterAt) -> Self {
        let members = shape.members();
        let master_pos = shape.master(master_at);
        Track {
            shape,
            members: members.iter().map(|&p| Some(p)).collect(),
            master: members.iter().position(|&p| p == master_pos).unwrap_or(0),
            start: START,
            end: *members.last().unwrap_or(&START),
        }
    }

    fn current(&self) -> Vec<u32> {
        self.members.iter().flatten().copied().collect()
    }

    /// The edit's line as `(insert, position)`, or why it does not apply now.
    fn resolve(&self, edit: Edit) -> Result<(bool, u32), &'static str> {
        let current = self.current();
        let master = self.members[self.master];
        let followers: Vec<u32> = self
            .members
            .iter()
            .enumerate()
            .filter(|(i, _)| *i != self.master)
            .filter_map(|(_, p)| *p)
            .collect();
        match edit {
            Edit::InsertBefore => Ok((true, self.start)),
            Edit::InsertInside => current
                .last()
                .copied()
                .filter(|&last| last > self.start)
                .map(|last| (true, last))
                .ok_or("the range has a single line left"),
            Edit::DeleteGap => {
                let reason = if self.shape == Shape::Contiguous {
                    "contiguous group and no earlier insert made a gap"
                } else {
                    "an earlier edit removed the only gap"
                };
                (self.start..=self.end)
                    .find(|p| !current.contains(p))
                    .map(|p| (false, p))
                    .ok_or(reason)
            }
            Edit::DeleteFollower => {
                let pick = match master {
                    Some(m) => followers
                        .iter()
                        .copied()
                        .find(|&p| p > m)
                        .or_else(|| followers.iter().copied().rev().find(|&p| p < m)),
                    None => followers.first().copied(),
                };
                pick.map(|p| (false, p)).ok_or("no follower left")
            }
            Edit::DeleteMaster => master
                .map(|p| (false, p))
                .ok_or("an earlier edit deleted the master"),
        }
    }

    fn apply(&mut self, insert: bool, q: u32) {
        for slot in &mut self.members {
            *slot = match *slot {
                Some(p) if insert && p >= q => Some(p + 1),
                Some(p) if !insert && p == q => None,
                Some(p) if !insert && p > q => Some(p - 1),
                other => other,
            };
        }
        if insert {
            if self.start >= q {
                self.start += 1;
            }
            if self.end >= q {
                self.end += 1;
            }
        } else {
            if self.start > q {
                self.start -= 1;
            }
            if self.end >= q && self.end > 0 {
                self.end -= 1;
            }
        }
    }
}

type Snapshot = BTreeMap<(u32, u32), (String, Option<String>)>;

fn snapshot(workbook: &Workbook) -> Snapshot {
    let mut out = BTreeMap::new();
    if let Some(sheet) = workbook.model().sheets.first() {
        for (at, cell) in sheet.iter_cells() {
            if cell.formula.is_none() && cell.value == CellValue::Empty {
                continue;
            }
            out.insert((at.row, at.col), (at.to_a1(), cell.formula.clone()));
        }
    }
    out
}

fn show(f: Option<&Option<String>>) -> String {
    match f {
        None => "absent".to_owned(),
        Some(None) => "-".to_owned(),
        Some(Some(f)) => format!("={f}"),
    }
}

enum Outcome {
    Pass,
    NotApplicable(&'static str),
    Fail {
        cell: String,
        model: String,
        reopened: String,
    },
}

fn broken(stage: &str, error: String) -> Outcome {
    Outcome::Fail {
        cell: "-".to_owned(),
        model: stage.to_owned(),
        reopened: error,
    }
}

fn run_case(
    shape: Shape,
    orient: Orient,
    master_at: MasterAt,
    edits: &[Edit],
) -> (String, Outcome) {
    let mut track = Track::new(shape, master_at);
    let mut steps = Vec::new();
    let mut words = Vec::new();
    let mut missing = None;
    for &edit in edits {
        match track.resolve(edit) {
            Ok((insert, q)) => {
                words.push(format!("{} {}", edit.name(), orient.line(q)));
                track.apply(insert, q);
                steps.push((insert, q));
            }
            Err(reason) => {
                words.push(format!("{} n/a", edit.name()));
                if missing.is_none() {
                    missing = Some(reason);
                }
            }
        }
    }
    let case = format!(
        "{} {} master-{}: {}",
        shape.name(),
        orient.name(),
        master_at.name(),
        words.join(" then ")
    );
    if let Some(reason) = missing {
        return (case, Outcome::NotApplicable(reason));
    }
    let run = AssertUnwindSafe(|| execute(shape, orient, master_at, &steps));
    let outcome = match catch_unwind(run) {
        Ok(outcome) => outcome,
        Err(_) => broken("run", "panic".to_owned()),
    };
    (case, outcome)
}

fn execute(shape: Shape, orient: Orient, master_at: MasterAt, steps: &[(bool, u32)]) -> Outcome {
    let source = match package(sheet_xml(shape, orient, master_at)) {
        Ok(source) => source,
        Err(e) => return broken("build", e),
    };
    let mut workbook = match Workbook::open(&source) {
        Ok(workbook) => workbook,
        Err(e) => return broken("open", format!("{e:?}")),
    };
    let base = snapshot(&workbook);
    for &pos in shape.members() {
        let got = base
            .values()
            .find(|(a1, _)| *a1 == orient.member(pos))
            .map(|(_, f)| f.clone());
        if got != Some(Some(formula(orient, pos))) {
            return Outcome::Fail {
                cell: orient.member(pos),
                model: format!("={}", formula(orient, pos)),
                reopened: format!("baseline:{}", show(got.as_ref())),
            };
        }
    }
    for &(insert, q) in steps {
        let op = orient.op(insert, q);
        if let Err(e) = workbook.apply_ops(vec![op], CalculationOptions::default()) {
            return broken("apply", format!("{e:?}"));
        }
    }
    let model = snapshot(&workbook);
    let saved = match workbook.save() {
        Ok(saved) => saved,
        Err(e) => return broken("save", format!("{e:?}")),
    };
    let reopened = match Workbook::open(&saved) {
        Ok(reopened) => snapshot(&reopened),
        Err(e) => return broken("reopen", format!("{e:?}")),
    };
    let keys: BTreeSet<(u32, u32)> = model.keys().chain(reopened.keys()).copied().collect();
    for key in keys {
        let m = model.get(&key);
        let r = reopened.get(&key);
        if m.map(|(_, f)| f) != r.map(|(_, f)| f) {
            return Outcome::Fail {
                cell: m.or(r).map(|(a1, _)| a1.clone()).unwrap_or_default(),
                model: show(m.map(|(_, f)| f)),
                reopened: show(r.map(|(_, f)| f)),
            };
        }
    }
    Outcome::Pass
}

#[derive(Default)]
struct Tally {
    cases: usize,
    failed: usize,
    skipped: usize,
}

impl Tally {
    fn record(&mut self, case: &str, outcome: Outcome) {
        self.cases += 1;
        match outcome {
            Outcome::Pass => {}
            Outcome::NotApplicable(reason) => {
                self.skipped += 1;
                println!("SKIP {case}: {reason}");
            }
            Outcome::Fail {
                cell,
                model,
                reopened,
            } => {
                self.failed += 1;
                println!("FAIL {case}: cell={cell} model={model} reopened={reopened}");
            }
        }
    }
}

fn configs() -> Vec<(Shape, Orient, MasterAt)> {
    let mut out = Vec::new();
    for shape in SHAPES {
        for orient in ORIENTS {
            for master_at in MASTERS {
                out.push((shape, orient, master_at));
            }
        }
    }
    out
}

#[test]
fn shared_formulas_keep_their_formula_after_structural_edits() {
    let mut grid = Tally::default();
    for (shape, orient, master_at) in configs() {
        for first in EDITS {
            for second in EDITS {
                let (case, outcome) = run_case(shape, orient, master_at, &[first, second]);
                grid.record(&case, outcome);
            }
        }
    }
    println!(
        "GRID cases={} failed={} skipped={}",
        grid.cases, grid.failed, grid.skipped
    );

    let mut triples = Tally::default();
    let pairs = [
        [Edit::InsertInside, Edit::DeleteGap],
        [Edit::DeleteGap, Edit::InsertInside],
    ];
    for (shape, orient, master_at) in configs() {
        if shape == Shape::Contiguous {
            continue;
        }
        for pair in pairs {
            for extra in EDITS {
                for slot in 0..=2 {
                    let mut edits = pair.to_vec();
                    edits.insert(slot, extra);
                    let (case, outcome) = run_case(shape, orient, master_at, &edits);
                    triples.record(&case, outcome);
                }
            }
        }
    }
    println!(
        "TRIPLES cases={} failed={} skipped={}",
        triples.cases, triples.failed, triples.skipped
    );

    assert_eq!(grid.failed, 0, "GRID failures");
    assert_eq!(triples.failed, 0, "TRIPLES failures");
}
