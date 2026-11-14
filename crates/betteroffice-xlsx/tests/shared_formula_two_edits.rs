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

/// Where the group's `ref` lies relative to its members.
#[derive(Clone, Copy, PartialEq)]
enum Span {
    Exact,
    Tail1,
    Tail2,
    Lead1,
}

#[derive(Clone, Copy, PartialEq)]
enum Edit {
    InsertBefore,
    InsertInside,
    DeleteGap,
    DeleteFollower,
    DeleteMaster,
    #[cfg(feature = "unstable-track-moves")]
    MoveHeadPastEnd,
    #[cfg(feature = "unstable-track-moves")]
    MoveGroupPastEnd,
    /// The member block, this far past the ref end, or before its start when negative.
    #[cfg(feature = "unstable-track-moves")]
    MoveMembers(i32),
}

/// One structural edit in 1-based lines along the edited axis.
#[derive(Clone, Copy)]
enum Step {
    Insert(u32),
    Delete(u32),
    #[cfg(feature = "unstable-track-moves")]
    Move {
        at: u32,
        count: u32,
        before: u32,
    },
}

const FEATURES: &str = if cfg!(feature = "unstable-track-moves") {
    "moves"
} else {
    "default"
};

const SHAPES: [Shape; 3] = [Shape::Contiguous, Shape::OneGap, Shape::TwoGaps];
const ORIENTS: [Orient; 2] = [Orient::Column, Orient::Row];
const MASTERS: [MasterAt; 2] = [MasterAt::First, MasterAt::Middle];
const SPANS: [Span; 4] = [Span::Exact, Span::Tail1, Span::Tail2, Span::Lead1];
const EDITS: [Edit; 5] = [
    Edit::InsertBefore,
    Edit::InsertInside,
    Edit::DeleteGap,
    Edit::DeleteFollower,
    Edit::DeleteMaster,
];

fn grid_edits() -> Vec<Edit> {
    #[allow(unused_mut)]
    let mut edits = EDITS.to_vec();
    #[cfg(feature = "unstable-track-moves")]
    edits.extend([Edit::MoveHeadPastEnd, Edit::MoveGroupPastEnd]);
    edits
}

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

impl Span {
    fn name(self) -> &'static str {
        match self {
            Span::Exact => "0",
            Span::Tail1 => "1",
            Span::Tail2 => "2",
            Span::Lead1 => "lead",
        }
    }

    fn lead(self) -> u32 {
        u32::from(self == Span::Lead1)
    }

    fn tail(self) -> u32 {
        match self {
            Span::Tail1 => 1,
            Span::Tail2 => 2,
            Span::Exact | Span::Lead1 => 0,
        }
    }

    fn suffix(self) -> String {
        match self {
            Span::Exact => String::new(),
            Span::Lead1 => " lead=1".to_owned(),
            _ => format!(" tail={}", self.tail()),
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

    fn op(self, step: Step) -> Op {
        let sheet = SheetId(0);
        let count = 1;
        match (self, step) {
            (Orient::Column, Step::Insert(pos)) => Op::InsertRows {
                sheet,
                at: pos - 1,
                count,
            },
            (Orient::Column, Step::Delete(pos)) => Op::DeleteRows {
                sheet,
                at: pos - 1,
                count,
            },
            (Orient::Row, Step::Insert(pos)) => Op::InsertCols {
                sheet,
                at: pos - 1,
                count,
            },
            (Orient::Row, Step::Delete(pos)) => Op::DeleteCols {
                sheet,
                at: pos - 1,
                count,
            },
            #[cfg(feature = "unstable-track-moves")]
            (Orient::Column, Step::Move { at, count, before }) => Op::MoveRows {
                sheet,
                row: at - 1,
                count,
                before: before - 1,
            },
            #[cfg(feature = "unstable-track-moves")]
            (Orient::Row, Step::Move { at, count, before }) => Op::MoveCols {
                sheet,
                col: at - 1,
                count,
                before: before - 1,
            },
        }
    }

    fn describe(self, step: Step) -> String {
        match step {
            Step::Insert(pos) | Step::Delete(pos) => self.line(pos),
            #[cfg(feature = "unstable-track-moves")]
            Step::Move { at, count, before } => {
                format!("{} x{count} before {}", self.line(at), self.line(before))
            }
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
            #[cfg(feature = "unstable-track-moves")]
            Edit::MoveHeadPastEnd => "move-head-past-end",
            #[cfg(feature = "unstable-track-moves")]
            Edit::MoveGroupPastEnd => "move-group-past-end",
            #[cfg(feature = "unstable-track-moves")]
            Edit::MoveMembers(_) => "move-members",
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

fn sheet_xml(shape: Shape, orient: Orient, master_at: MasterAt, span: Span) -> String {
    let members = shape.members();
    let master = shape.master(master_at);
    let end = *members.last().unwrap_or(&START) + span.tail();
    let reference = format!(
        "{}:{}",
        orient.member(START - span.lead()),
        orient.member(end)
    );
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
    fn new(shape: Shape, master_at: MasterAt, span: Span) -> Self {
        let members = shape.members();
        let master_pos = shape.master(master_at);
        Track {
            shape,
            members: members.iter().map(|&p| Some(p)).collect(),
            master: members.iter().position(|&p| p == master_pos).unwrap_or(0),
            start: START - span.lead(),
            end: *members.last().unwrap_or(&START) + span.tail(),
        }
    }

    fn current(&self) -> Vec<u32> {
        self.members.iter().flatten().copied().collect()
    }

    /// The edit as a step on the current lines, or why it does not apply now.
    fn resolve(&self, edit: Edit) -> Result<Step, &'static str> {
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
            Edit::InsertBefore => Ok(Step::Insert(self.start)),
            Edit::InsertInside => current
                .last()
                .copied()
                .filter(|&last| last > self.start)
                .map(Step::Insert)
                .ok_or("the range has a single line left"),
            Edit::DeleteGap => {
                let reason = if self.shape == Shape::Contiguous {
                    "contiguous group and no earlier insert made a gap"
                } else {
                    "an earlier edit removed the only gap"
                };
                (self.start..=self.end)
                    .find(|p| !current.contains(p))
                    .map(Step::Delete)
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
                pick.map(Step::Delete).ok_or("no follower left")
            }
            Edit::DeleteMaster => master
                .map(Step::Delete)
                .ok_or("an earlier edit deleted the master"),
            #[cfg(feature = "unstable-track-moves")]
            Edit::MoveHeadPastEnd => Ok(Step::Move {
                at: self.start,
                count: 2,
                before: self.end + 3,
            }),
            #[cfg(feature = "unstable-track-moves")]
            Edit::MoveGroupPastEnd => Ok(Step::Move {
                at: self.start,
                count: self.end - self.start + 1,
                before: self.end + 3,
            }),
            #[cfg(feature = "unstable-track-moves")]
            Edit::MoveMembers(offset) => {
                let first = current.iter().copied().min().ok_or("no member left")?;
                let last = current.iter().copied().max().ok_or("no member left")?;
                let shift = offset.unsigned_abs();
                let before = if offset > 0 {
                    Some(self.end + shift)
                } else {
                    self.start.checked_sub(shift).filter(|&b| b >= 1)
                };
                match before {
                    Some(before) => Ok(Step::Move {
                        at: first,
                        count: last - first + 1,
                        before,
                    }),
                    None => Err("the destination would lie before line 1"),
                }
            }
        }
    }

    fn apply(&mut self, step: Step) {
        let (insert, q) = match step {
            Step::Insert(q) => (true, q),
            Step::Delete(q) => (false, q),
            #[cfg(feature = "unstable-track-moves")]
            Step::Move { at, count, before } => {
                for p in self.members.iter_mut().flatten() {
                    if *p >= at && *p < at + count {
                        *p = before - count + (*p - at);
                    } else if *p >= at + count && *p < before {
                        *p -= count;
                    } else if *p >= before && *p < at {
                        *p += count;
                    }
                }
                let current = self.current();
                self.start = current.iter().copied().min().unwrap_or(self.start);
                self.end = current.iter().copied().max().unwrap_or(self.end);
                return;
            }
        };
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
    (shape, orient, master_at, span): Config,
    edits: &[Edit],
    audit: &mut Audit,
) -> (String, Outcome) {
    let mut track = Track::new(shape, master_at, span);
    let mut steps = Vec::new();
    let mut words = Vec::new();
    let mut missing = None;
    for &edit in edits {
        match track.resolve(edit) {
            Ok(step) => {
                words.push(format!("{} {}", edit.name(), orient.describe(step)));
                track.apply(step);
                steps.push(step);
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
        "{} {} master-{}{}: {}",
        shape.name(),
        orient.name(),
        master_at.name(),
        span.suffix(),
        words.join(" then ")
    );
    if let Some(reason) = missing {
        return (case, Outcome::NotApplicable(reason));
    }
    let config = (shape, orient, master_at, span);
    let run = AssertUnwindSafe(|| execute(config, &steps, &case, audit));
    let outcome = match catch_unwind(run) {
        Ok(outcome) => outcome,
        Err(_) => broken("run", "panic".to_owned()),
    };
    (case, outcome)
}

fn execute(
    (shape, orient, master_at, span): Config,
    steps: &[Step],
    case: &str,
    audit: &mut Audit,
) -> Outcome {
    let source = match package(sheet_xml(shape, orient, master_at, span)) {
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
    for &step in steps {
        let op = orient.op(step);
        if let Err(e) = workbook.apply_ops(vec![op], CalculationOptions::default()) {
            return broken("apply", format!("{e:?}"));
        }
    }
    finish(&workbook, case, audit)
}

/// Saves, audits the saved worksheets, reopens and compares every formula with the model.
fn finish(workbook: &Workbook, case: &str, audit: &mut Audit) -> Outcome {
    let model = snapshot(workbook);
    let saved = match workbook.save() {
        Ok(saved) => saved,
        Err(e) => return broken("save", format!("{e:?}")),
    };
    audit.check(case, &saved);
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

impl Tally {
    fn add(&mut self, other: &Tally) {
        self.cases += other.cases;
        self.failed += other.failed;
        self.skipped += other.skipped;
    }
}

type Config = (Shape, Orient, MasterAt, Span);

fn configs(span: Span) -> Vec<Config> {
    let mut out = Vec::new();
    for shape in SHAPES {
        for orient in ORIENTS {
            for master_at in MASTERS {
                out.push((shape, orient, master_at, span));
            }
        }
    }
    out
}

/// Every saved worksheet part checked so far and the shared-group violations found in them.
#[derive(Default)]
struct Audit {
    parts: usize,
    violations: Vec<String>,
}

impl Audit {
    fn check(&mut self, case: &str, saved: &[u8]) {
        let parts = match ooxml_opc::unzip_parts(saved) {
            Ok(parts) => parts,
            Err(e) => {
                self.violations.push(format!("{case}: unzip {e}"));
                return;
            }
        };
        for (path, bytes) in parts {
            if !path.starts_with("xl/worksheets/") || !path.ends_with(".xml") {
                continue;
            }
            self.parts += 1;
            let xml = String::from_utf8_lossy(&bytes);
            for violation in shared_formula_violations(&xml) {
                self.violations.push(format!("{case} {path}: {violation}"));
            }
        }
    }

    fn report(&self) {
        let (parts, violations) = (self.parts, self.violations.len());
        println!("\nINVARIANT parts={parts} violations={violations}");
        for violation in self.violations.iter().take(3) {
            println!("VIOLATION {violation}");
        }
    }
}

/// Every `<c>` whose `<f t="shared">` names an `si` must find exactly one master `<f>` with
/// that `si` and a `ref`, and its own `r` must lie inside that `ref`.
pub fn shared_formula_violations(xml: &str) -> Vec<String> {
    let mut violations = Vec::new();
    let mut masters: BTreeMap<String, Vec<String>> = BTreeMap::new();
    let mut members: Vec<(Option<String>, String)> = Vec::new();
    let mut rest = xml;
    while let Some((cell, body, after)) = next_element(rest, "c") {
        rest = after;
        let Some((formula, _, _)) = body.and_then(|body| next_element(body, "f")) else {
            continue;
        };
        if attribute(formula, "t").as_deref() != Some("shared") {
            continue;
        }
        let r = attribute(cell, "r");
        let Some(si) = attribute(formula, "si") else {
            violations.push(format!("{}: shared formula without si", show_r(&r)));
            continue;
        };
        if let Some(reference) = attribute(formula, "ref") {
            masters.entry(si.clone()).or_default().push(reference);
        }
        members.push((r, si));
    }
    for (r, si) in members {
        let shown = show_r(&r);
        let refs = masters.get(&si).map(Vec::as_slice).unwrap_or_default();
        let [reference] = refs else {
            violations.push(format!(
                "{shown} si={si}: {} masters with a ref",
                refs.len()
            ));
            continue;
        };
        let Some(at) = r.as_deref().and_then(parse_a1) else {
            violations.push(format!("{shown} si={si}: unreadable r"));
            continue;
        };
        let Some((from, to)) = parse_range(reference) else {
            violations.push(format!("{shown} si={si}: unreadable ref {reference}"));
            continue;
        };
        let cols = from.0.min(to.0)..=from.0.max(to.0);
        let rows = from.1.min(to.1)..=from.1.max(to.1);
        if !cols.contains(&at.0) || !rows.contains(&at.1) {
            violations.push(format!("{shown} si={si}: outside ref {reference}"));
        }
    }
    violations
}

fn show_r(r: &Option<String>) -> String {
    r.clone().unwrap_or_else(|| "<no r>".to_owned())
}

/// The next `<name ...>` element as (start-tag attributes, body if not self-closing, rest).
fn next_element<'a>(text: &'a str, name: &str) -> Option<(&'a str, Option<&'a str>, &'a str)> {
    let open = format!("<{name}");
    let mut from = 0;
    loop {
        let start = from + text[from..].find(&open)?;
        let after_name = start + open.len();
        let next = text[after_name..].chars().next()?;
        if !(next.is_ascii_whitespace() || next == '>' || next == '/') {
            from = after_name;
            continue;
        }
        let mut quote = None;
        let mut tag_end = None;
        for (i, ch) in text[after_name..].char_indices() {
            match (quote, ch) {
                (None, '"' | '\'') => quote = Some(ch),
                (Some(q), _) if ch == q => quote = None,
                (None, '>') => {
                    tag_end = Some(after_name + i);
                    break;
                }
                _ => {}
            }
        }
        let tag_end = tag_end?;
        let tag = &text[after_name..tag_end];
        if let Some(attrs) = tag.strip_suffix('/') {
            return Some((attrs, None, &text[tag_end + 1..]));
        }
        let close = format!("</{name}>");
        let body_start = tag_end + 1;
        let body_end = body_start + text[body_start..].find(&close)?;
        let body = &text[body_start..body_end];
        return Some((tag, Some(body), &text[body_end + close.len()..]));
    }
}

fn attribute(tag: &str, name: &str) -> Option<String> {
    let mut rest = tag;
    loop {
        rest = rest.trim_start();
        let eq = rest.find('=')?;
        let key = rest[..eq].trim();
        let value = rest[eq + 1..].trim_start();
        let quote = value.chars().next()?;
        if quote != '"' && quote != '\'' {
            return None;
        }
        let close = value[1..].find(quote)?;
        if key == name {
            return Some(value[1..1 + close].to_owned());
        }
        rest = &value[close + 2..];
    }
}

/// `A1`-style address as 1-based (column, row), ignoring `$`.
fn parse_a1(text: &str) -> Option<(u32, u32)> {
    let text: String = text.chars().filter(|c| *c != '$').collect();
    let split = text.find(|c: char| c.is_ascii_digit())?;
    let (letters, digits) = text.split_at(split);
    if letters.is_empty() || !letters.chars().all(|c| c.is_ascii_uppercase()) {
        return None;
    }
    let col = letters
        .bytes()
        .fold(0u32, |acc, b| acc * 26 + u32::from(b - b'A' + 1));
    Some((col, digits.parse().ok()?))
}

fn parse_range(text: &str) -> Option<((u32, u32), (u32, u32))> {
    match text.split_once(':') {
        Some((from, to)) => Some((parse_a1(from)?, parse_a1(to)?)),
        None => parse_a1(text).map(|at| (at, at)),
    }
}

#[test]
fn shared_formula_violations_flag_a_member_outside_its_ref() {
    let inside = r#"<sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:A4" si="0">1</f></c></row><row r="2"><c r="A2"><f t="shared" si="0"/></c></row><cols/></sheetData>"#;
    assert!(shared_formula_violations(inside).is_empty());
    let outside = r#"<c r="A2"><f t="shared" ref="A2:A5" si="0">1</f></c><c r="A6"><f t="shared" si="0"/></c><c r="B1"><f t="shared" si="1"/></c>"#;
    assert_eq!(shared_formula_violations(outside).len(), 2);
}

#[cfg(feature = "unstable-track-moves")]
#[test]
fn shared_group_moved_past_its_ref_reopens_with_model_formulas() {
    let sheet = format!(
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{main}"><sheetData>"#,
            r#"<row r="1"><c r="A1"><f t="shared" ref="A1:A4" si="0">1</f><v>1</v></c></row>"#,
            r#"<row r="2"><c r="A2"><f t="shared" si="0"/><v>1</v></c></row>"#,
            r#"</sheetData></worksheet>"#,
        ),
        main = MAIN,
    );
    let mut workbook = Workbook::open(&package(sheet).unwrap()).unwrap();
    let op = Op::MoveRows {
        sheet: SheetId(0),
        row: 0,
        count: 2,
        before: 6,
    };
    workbook
        .apply_ops(vec![op], CalculationOptions::default())
        .unwrap();
    let mut audit = Audit::default();
    let outcome = finish(&workbook, "move rows 1..2 before 7", &mut audit);
    audit.report();
    match outcome {
        Outcome::Pass => {}
        Outcome::NotApplicable(reason) => panic!("not applicable: {reason}"),
        Outcome::Fail {
            cell,
            model,
            reopened,
        } => panic!("cell={cell} model={model} reopened={reopened}"),
    }
    assert!(audit.violations.is_empty(), "INVARIANT violations");
}

#[cfg(feature = "unstable-track-moves")]
#[test]
fn shared_formula_members_moved_around_their_ref_reopen_with_model_formulas() {
    let mut audit = Audit::default();
    let mut total = Tally::default();
    for span in [Span::Tail1, Span::Tail2] {
        let mut tally = Tally::default();
        let mut red = BTreeSet::new();
        for config in configs(span) {
            let track = Track::new(config.0, config.2, span);
            let reach = i32::try_from(track.end - track.start + 3).unwrap_or(0);
            for offset in (1..=reach).flat_map(|k| [k, -k]) {
                let sweep = Edit::MoveMembers(offset);
                let mut runs = vec![vec![sweep]];
                for extra in EDITS {
                    runs.push(vec![sweep, extra]);
                    runs.push(vec![extra, sweep]);
                }
                for edits in runs {
                    let (case, outcome) = run_case(config, &edits, &mut audit);
                    if matches!(outcome, Outcome::Fail { .. }) {
                        red.insert(offset);
                    }
                    tally.record(&case, outcome);
                }
            }
        }
        let reds: Vec<String> = red.iter().map(i32::to_string).collect();
        let reds = reds.join(",");
        println!(
            "\nSWEEP tail={} cases={} failed={} skipped={}",
            span.name(),
            tally.cases,
            tally.failed,
            tally.skipped
        );
        println!("\nSWEEP tail={} red_offsets={reds}", span.name());
        total.add(&tally);
    }
    audit.report();
    assert_eq!(total.failed, 0, "SWEEP failures");
    assert_eq!(audit.violations.len(), 0, "INVARIANT violations");
}

#[test]
fn shared_formulas_keep_their_formula_after_structural_edits() {
    let mut audit = Audit::default();
    let mut grid = Tally::default();
    let mut triples = Tally::default();
    let edits = grid_edits();
    let pairs = [
        [Edit::InsertInside, Edit::DeleteGap],
        [Edit::DeleteGap, Edit::InsertInside],
    ];
    for span in SPANS {
        let mut tally = Tally::default();
        for config in configs(span) {
            for &only in &edits {
                let (case, outcome) = run_case(config, &[only], &mut audit);
                tally.record(&case, outcome);
            }
            for &first in &edits {
                for &second in &edits {
                    let (case, outcome) = run_case(config, &[first, second], &mut audit);
                    tally.record(&case, outcome);
                }
            }
        }
        println!(
            "\nGRID tail={} cases={} failed={} skipped={}",
            span.name(),
            tally.cases,
            tally.failed,
            tally.skipped
        );
        grid.add(&tally);

        let mut tally = Tally::default();
        for config in configs(span) {
            if config.0 == Shape::Contiguous && span == Span::Exact {
                continue;
            }
            for pair in pairs {
                for extra in EDITS {
                    for slot in 0..=2 {
                        let mut edits = pair.to_vec();
                        edits.insert(slot, extra);
                        let (case, outcome) = run_case(config, &edits, &mut audit);
                        tally.record(&case, outcome);
                    }
                }
            }
        }
        println!(
            "\nTRIPLES tail={} cases={} failed={} skipped={}",
            span.name(),
            tally.cases,
            tally.failed,
            tally.skipped
        );
        triples.add(&tally);
    }
    println!(
        "\nGRID features={FEATURES} cases={} failed={} skipped={}",
        grid.cases, grid.failed, grid.skipped
    );
    println!(
        "\nTRIPLES cases={} failed={} skipped={}",
        triples.cases, triples.failed, triples.skipped
    );
    audit.report();

    assert_eq!(grid.failed, 0, "GRID failures");
    assert_eq!(triples.failed, 0, "TRIPLES failures");
    assert_eq!(audit.violations.len(), 0, "INVARIANT violations");
}
