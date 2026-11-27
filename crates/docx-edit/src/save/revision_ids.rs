use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::fmt::{self, Write};
use std::sync::Arc;
#[cfg(not(target_arch = "wasm32"))]
use std::time::{Duration, Instant};

use docx_parse::S13SaveRequest;
use docx_parse::block::BlockContent;
use docx_parse::inline::{
    InlineNode, Run, RunContent, RunPropertyChange, StructuredFieldContent, StructuredFieldTree,
};
use docx_parse::paragraph::{Paragraph, ParagraphContent, TrackedInline};
use docx_parse::shape::{DrawingSceneNode, Shape};
use docx_parse::table::Table;
use serde_json::{Map, Value};

mod collect;

use collect::collect;

const MAX_ID: u32 = 1 << 31;
const PLACEHOLDER_LIMIT: f64 = (1u64 << 53) as f64;
const NO_SITE: usize = usize::MAX;
const HYPERLINK_FIELDS: [&str; 2] = ["children", "structuredChildren"];
const STRUCTURED_FIELDS: [&str; 2] = ["inline", "blocks"];
const FIELD_TREE_FIELDS: [&str; 3] = ["code", "result", "children"];
const SCENE_NODE_FIELDS: [&str; 2] = ["shape", "children"];
static NO_VIEWS: BTreeSet<String> = BTreeSet::new();

type FastMap<K, V> = HashMap<K, V, foldhash::fast::RandomState>;
type FastSet<T> = HashSet<T, foldhash::fast::RandomState>;

#[cfg(test)]
thread_local! {
    static WORK: std::cell::RefCell<Work> = std::cell::RefCell::new(Work::default());
}

#[cfg(test)]
#[derive(Clone, Copy, Debug, Default)]
struct Work {
    fields: usize,
    strings: usize,
    string_bytes: usize,
    reservation_inserts: usize,
    site_inserts: usize,
    collected_blocks: usize,
    body_block_selections: usize,
    block_inserts: usize,
    block_probes: usize,
    make_mut: usize,
    id_probes: usize,
    verified: usize,
}

#[cfg(test)]
fn count_work(count: impl FnOnce(&mut Work)) {
    WORK.with_borrow_mut(count);
}

#[derive(Debug, PartialEq, Eq)]
pub enum RevisionIdsError {
    Exhausted,
    Unassigned(u64),
    Serialize(String),
}

impl fmt::Display for RevisionIdsError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Exhausted => formatter.write_str("no Word revision ids left in range"),
            Self::Unassigned(id) => write!(formatter, "unassigned session revision id {id}"),
            Self::Serialize(message) => formatter.write_str(message),
        }
    }
}

impl std::error::Error for RevisionIdsError {}

impl serde::ser::Error for RevisionIdsError {
    fn custom<T: fmt::Display>(message: T) -> Self {
        Self::Serialize(message.to_string())
    }
}

/// Assigns revision IDs in a typed save request in each container's projection order.
/// `session` holds the placeholders minted by the projection for session revisions.
/// `session` must be disjoint from every non-session revision ID in the request.
pub fn assign_session_revision_ids(
    request: &mut S13SaveRequest,
    session: &BTreeSet<u64>,
) -> Result<BTreeMap<u64, u32>, RevisionIdsError> {
    assign(request, session, None, &NO_VIEWS, || {})
}

#[doc(hidden)]
pub fn assign_session_revision_ids_with_paths(
    request: &mut S13SaveRequest,
    session: &BTreeSet<u64>,
    paths: &mut Vec<(String, u64)>,
) -> Result<BTreeMap<u64, u32>, RevisionIdsError> {
    assign_session_revision_ids_with_paths_skipping_views(request, session, paths, &NO_VIEWS)
}

#[doc(hidden)]
pub fn assign_session_revision_ids_with_paths_skipping_views(
    request: &mut S13SaveRequest,
    session: &BTreeSet<u64>,
    paths: &mut Vec<(String, u64)>,
    skipped_views: &BTreeSet<String>,
) -> Result<BTreeMap<u64, u32>, RevisionIdsError> {
    paths.clear();
    assign(request, session, Some(paths), skipped_views, || {})
}

#[doc(hidden)]
#[cfg(not(target_arch = "wasm32"))]
pub fn assign_session_revision_ids_with_timings(
    request: &mut S13SaveRequest,
    session: &BTreeSet<u64>,
) -> Result<(BTreeMap<u64, u32>, [Duration; 3]), RevisionIdsError> {
    let mut passes = [Duration::ZERO; 3];
    let mut timings = passes.iter_mut();
    let mut started = Instant::now();
    let ids = assign(request, session, None, &NO_VIEWS, || {
        let finished = Instant::now();
        *timings.next().unwrap() = finished.duration_since(started);
        started = finished;
    })?;
    Ok((ids, passes))
}

fn assign(
    request: &mut S13SaveRequest,
    session: &BTreeSet<u64>,
    paths: Option<&mut Vec<(String, u64)>>,
    skipped_views: &BTreeSet<String>,
    mut pass_complete: impl FnMut(),
) -> Result<BTreeMap<u64, u32>, RevisionIdsError> {
    if session.is_empty() {
        return Ok(BTreeMap::new());
    }
    let session_ids = SessionIds::new(session);
    let mut found = collect::<false>(request, &session_ids, skipped_views);
    pass_complete();
    if found.sites.is_empty() {
        return Ok(BTreeMap::new());
    }
    let mut walker = SiteWalker::new(
        &session_ids,
        paths,
        skipped_views,
        &found.blocks_with_sites,
        &found.body_blocks_with_sites,
        found.site_count,
    );
    walker.request(request);
    walker.order_paths();

    let mut next = found.max;
    let mut index = walker.order.first;
    while let Some(site) = walker.sites.get_mut(index) {
        let Some(count) = found.sites.get_mut(&site.placeholder) else {
            return Err(RevisionIdsError::Unassigned(site.placeholder));
        };
        count.remaining = count
            .remaining
            .checked_sub(1)
            .ok_or(RevisionIdsError::Unassigned(site.placeholder))?;
        if count.assigned == 0 {
            count.assigned = next_id(&mut next, found.max, &found.used)?;
        }
        #[cfg(test)]
        count_work(|work| work.verified += 1);
        site.assigned = count.assigned;
        index = site.next;
    }
    if let Some(id) = found
        .sites
        .iter()
        .filter_map(|(&id, count)| (count.remaining != 0).then_some(id))
        .min()
    {
        return Err(RevisionIdsError::Unassigned(id));
    }
    let mut assigned = Vec::with_capacity(found.sites.len());
    assigned.extend(
        session
            .iter()
            .filter_map(|id| found.sites.get(id).map(|count| (*id, count.assigned))),
    );
    let ids = assigned.into_iter().collect();
    pass_complete();
    for site in &mut walker.sites {
        site.write();
    }
    pass_complete();
    Ok(ids)
}

struct SessionIds {
    first: f64,
    ids: FastSet<u64>,
}

impl SessionIds {
    fn new(session: &BTreeSet<u64>) -> Self {
        Self {
            first: session.first().map_or(f64::INFINITY, |&id| id as f64),
            ids: session.iter().copied().collect(),
        }
    }

    fn placeholder(&self, number: f64) -> Option<u64> {
        if number < self.first {
            return None;
        }
        placeholder(number).filter(|id| self.ids.contains(id))
    }

    fn model_placeholder(&self, record: &Map<String, Value>) -> Option<u64> {
        record.get("author").and_then(Value::as_str)?;
        self.placeholder(record.get("id").and_then(Value::as_f64)?)
    }
}

#[doc(hidden)]
pub struct RevisionIdReservations {
    pub used: BTreeSet<u32>,
    pub max: u32,
    pub sites: BTreeMap<(String, u32), usize>,
    pub string_paths: BTreeSet<String>,
}

#[doc(hidden)]
pub fn session_revision_id_reservations(
    request: &S13SaveRequest,
    session: &BTreeSet<u64>,
) -> Result<RevisionIdReservations, RevisionIdsError> {
    session_revision_id_reservations_skipping_views(request, session, &NO_VIEWS)
}

#[doc(hidden)]
pub fn session_revision_id_reservations_skipping_views(
    request: &S13SaveRequest,
    session: &BTreeSet<u64>,
    skipped_views: &BTreeSet<String>,
) -> Result<RevisionIdReservations, RevisionIdsError> {
    let found = collect::<true>(request, &SessionIds::new(session), skipped_views);
    Ok(RevisionIdReservations {
        used: found.used.into_iter().collect(),
        max: found.max,
        sites: found.reservations,
        string_paths: found.string_paths,
    })
}

fn next_id(next: &mut u32, max: u32, used: &FastSet<u32>) -> Result<u32, RevisionIdsError> {
    loop {
        *next = if *next + 1 < MAX_ID { *next + 1 } else { 1 };
        if *next == max {
            return Err(RevisionIdsError::Exhausted);
        }
        #[cfg(test)]
        count_work(|work| work.id_probes += 1);
        if !used.contains(next) {
            return Ok(*next);
        }
    }
}

fn placeholder(id: f64) -> Option<u64> {
    ((0.0..PLACEHOLDER_LIMIT).contains(&id) && id.fract() == 0.0).then_some(id as u64)
}

#[derive(Default)]
struct SiteCount {
    remaining: usize,
    assigned: u32,
}

#[derive(Default)]
struct Collected {
    max: u32,
    used: FastSet<u32>,
    sites: FastMap<u64, SiteCount>,
    reservations: BTreeMap<(String, u32), usize>,
    string_paths: BTreeSet<String>,
    site_count: usize,
    blocks_with_sites: FastSet<*const ()>,
    body_blocks_with_sites: Vec<usize>,
}

fn block_key(block: &BlockContent) -> Option<*const ()> {
    match block {
        BlockContent::Paragraph(block) => Some(Arc::as_ptr(block).cast()),
        BlockContent::Table(block) => Some(Arc::as_ptr(block).cast()),
        BlockContent::BlockSdt(block) => Some(Arc::as_ptr(block).cast()),
        BlockContent::RawXml(_) => None,
    }
}

fn push_pointer(pointer: &mut String, key: &str) {
    pointer.push('/');
    let mut start = 0;
    for (index, character) in key.match_indices(['~', '/']) {
        pointer.push_str(&key[start..index]);
        pointer.push_str(if character == "~" { "~0" } else { "~1" });
        start = index + 1;
    }
    pointer.push_str(&key[start..]);
}

enum Site<'a> {
    Typed(&'a mut f64),
    Model(&'a mut Value),
}

struct PendingSite<'a> {
    site: Site<'a>,
    placeholder: u64,
    assigned: u32,
    next: usize,
}

impl PendingSite<'_> {
    fn write(&mut self) {
        match &mut self.site {
            Site::Typed(id) => **id = f64::from(self.assigned),
            Site::Model(id) => **id = Value::from(self.assigned),
        }
    }
}

struct SiteOrder {
    first: usize,
    last: usize,
}

impl Default for SiteOrder {
    fn default() -> Self {
        Self {
            first: NO_SITE,
            last: NO_SITE,
        }
    }
}

struct SiteWalker<'a, 's, 'p> {
    session: &'s SessionIds,
    skipped_views: &'s BTreeSet<String>,
    blocks_with_sites: &'s FastSet<*const ()>,
    body_blocks_with_sites: &'s [usize],
    sites: Vec<PendingSite<'a>>,
    order: SiteOrder,
    paths: Option<&'p mut Vec<(String, u64)>>,
    path: String,
}

impl<'a, 's, 'p> SiteWalker<'a, 's, 'p> {
    fn new(
        session: &'s SessionIds,
        mut paths: Option<&'p mut Vec<(String, u64)>>,
        skipped_views: &'s BTreeSet<String>,
        blocks_with_sites: &'s FastSet<*const ()>,
        body_blocks_with_sites: &'s [usize],
        site_count: usize,
    ) -> Self {
        if let Some(paths) = &mut paths {
            paths.reserve(site_count);
        }
        Self {
            session,
            skipped_views,
            blocks_with_sites,
            body_blocks_with_sites,
            sites: Vec::with_capacity(site_count),
            order: SiteOrder::default(),
            paths,
            path: String::new(),
        }
    }

    fn at(&mut self, key: &str, visit: impl FnOnce(&mut Self)) {
        if self.paths.is_none() {
            return visit(self);
        }
        let length = self.path.len();
        push_pointer(&mut self.path, key);
        if !self.skipped_views.contains(&self.path) {
            visit(self);
        }
        self.path.truncate(length);
    }

    fn index(&mut self, index: usize, visit: impl FnOnce(&mut Self)) {
        let length = self.path.len();
        write!(self.path, "/{index}").unwrap();
        if !self.skipped_views.contains(&self.path) {
            visit(self);
        }
        self.path.truncate(length);
    }

    fn entries<T>(
        &mut self,
        values: impl IntoIterator<Item = T>,
        mut visit: impl FnMut(&mut Self, T),
    ) {
        if self.paths.is_some() {
            for (index, value) in values.into_iter().enumerate() {
                self.index(index, |walker| visit(walker, value));
            }
        } else {
            for value in values {
                visit(self, value);
            }
        }
    }

    fn field_entries<T>(
        &mut self,
        key: &str,
        values: impl IntoIterator<Item = T>,
        visit: impl FnMut(&mut Self, T),
    ) {
        self.at(key, |walker| walker.entries(values, visit));
    }

    fn request(&mut self, request: &'a mut S13SaveRequest) {
        self.at("document", |walker| {
            walker.at("content", |walker| {
                walker.body_blocks(&mut request.document.content)
            });
            walker.field_entries(
                "sections",
                request.document.sections.iter_mut().flatten(),
                |walker, section| {
                    walker.at("content", |walker| walker.blocks(&mut section.content));
                },
            );
            walker.field_entries(
                "comments",
                request.document.comments.iter_mut().flatten(),
                |walker, comment| {
                    walker.field_entries("content", &mut comment.content, Self::paragraph);
                    walker.at("blockContent", |walker| {
                        walker.blocks(&mut comment.block_content)
                    });
                },
            );
        });
        for (key, entries) in [
            ("headerEntries", &mut request.header_entries),
            ("footerEntries", &mut request.footer_entries),
        ] {
            self.field_entries(key, entries, |walker, (_, header_footer)| {
                walker.at("1", |walker| {
                    walker.at("content", |walker| {
                        walker.blocks(&mut header_footer.content)
                    });
                });
            });
        }
        for (key, notes) in [
            ("footnotes", &mut request.footnotes),
            ("endnotes", &mut request.endnotes),
            ("footnoteSeparators", &mut request.footnote_separators),
            ("endnoteSeparators", &mut request.endnote_separators),
        ] {
            self.field_entries(key, notes, |walker, note| {
                walker.at("content", |walker| walker.blocks(&mut note.content));
            });
        }
        if let Some(plan) = &mut request.paragraph_ids {
            self.at("paragraphIds", |walker| {
                walker.field_entries("splicedParts", &mut plan.spliced_parts, |walker, part| {
                    walker.field_entries(
                        "tablePatches",
                        &mut part.table_patches,
                        |walker, table| {
                            walker.at("before", |walker| {
                                walker.model(&mut table.before, ModelKind::Any);
                            });
                            walker.at("after", |walker| {
                                walker.model(&mut table.after, ModelKind::Any);
                            });
                        },
                    );
                });
            });
        }
    }

    fn push(&mut self, placeholder: u64, site: Site<'a>) {
        if let Some(paths) = &mut self.paths {
            paths.push((format!("{}/id", self.path), placeholder));
        }
        let index = self.sites.len();
        self.sites.push(PendingSite {
            site,
            placeholder,
            assigned: 0,
            next: NO_SITE,
        });
        self.append(SiteOrder {
            first: index,
            last: index,
        });
    }

    fn append(&mut self, order: SiteOrder) {
        if order.first == NO_SITE {
            return;
        }
        if let Some(last) = self.sites.get_mut(self.order.last) {
            last.next = order.first;
        } else {
            self.order.first = order.first;
        }
        self.order.last = order.last;
    }

    fn order_paths(&mut self) {
        let Some(paths) = &mut self.paths else {
            return;
        };
        let mut index = self.order.first;
        let in_visit_order = self.sites.iter().enumerate().all(|(position, site)| {
            let matches = index == position;
            index = site.next;
            matches
        });
        if !in_visit_order {
            let mut unordered = std::mem::take(*paths);
            paths.reserve(unordered.len());
            let mut index = self.order.first;
            while let Some(site) = self.sites.get(index) {
                paths.push(std::mem::take(&mut unordered[index]));
                index = site.next;
            }
        }
    }

    fn id(&mut self, id: &'a mut f64) {
        if let Some(placeholder) = self.session.placeholder(*id) {
            self.push(placeholder, Site::Typed(id));
        }
    }

    fn info(&mut self, id: &'a mut f64) {
        self.at("info", |walker| walker.id(id));
    }

    fn body_blocks(&mut self, blocks: &'a mut [BlockContent]) {
        let mut blocks = blocks.iter_mut();
        let mut next = 0;
        for &index in self.body_blocks_with_sites {
            let block = blocks.nth(index - next).unwrap();
            next = index + 1;
            #[cfg(test)]
            count_work(|work| work.body_block_selections += 1);
            if self.paths.is_some() {
                self.index(index, |walker| walker.block(block));
            } else {
                self.block(block);
            }
        }
    }

    fn blocks(&mut self, blocks: &'a mut [BlockContent]) {
        self.entries(blocks, |walker, block| {
            let Some(key) = block_key(block) else {
                return;
            };
            #[cfg(test)]
            count_work(|work| work.block_probes += 1);
            if walker.blocks_with_sites.contains(&key) {
                walker.block(block);
            }
        });
    }

    fn block(&mut self, block: &'a mut BlockContent) {
        #[cfg(test)]
        count_work(|work| work.make_mut += 1);
        match block {
            BlockContent::Paragraph(paragraph) => self.paragraph(Arc::make_mut(paragraph)),
            BlockContent::Table(table) => self.table(Arc::make_mut(table)),
            BlockContent::BlockSdt(control) => {
                let control = Arc::make_mut(control);
                self.at("content", |walker| walker.blocks(&mut control.content));
            }
            BlockContent::RawXml(_) => unreachable!(),
        }
    }

    fn table(&mut self, table: &'a mut Table) {
        self.field_entries("rows", &mut table.rows, |walker, row| {
            walker.field_entries("cells", &mut row.cells, |walker, cell| {
                walker.at("content", |walker| walker.blocks(&mut cell.content));
                if let Some(change) = &mut cell.structural_change {
                    walker.at("structuralChange", |walker| {
                        walker.info(&mut change.info.id)
                    });
                }
                walker.field_entries(
                    "propertyChanges",
                    cell.property_changes.iter_mut().flatten(),
                    |walker, change| walker.info(&mut change.info.id),
                );
            });
            if let Some(change) = &mut row.structural_change {
                walker.at("structuralChange", |walker| {
                    walker.info(&mut change.info.id)
                });
            }
            walker.field_entries(
                "propertyChanges",
                row.property_changes.iter_mut().flatten(),
                |walker, change| walker.info(&mut change.info.id),
            );
        });
        self.field_entries(
            "propertyChanges",
            table.property_changes.iter_mut().flatten(),
            |walker, change| walker.info(&mut change.info.id),
        );
    }

    fn paragraph(&mut self, paragraph: &'a mut Paragraph) {
        self.field_entries(
            "content",
            &mut paragraph.content,
            |walker, child| match child {
                ParagraphContent::Inline(node) => walker.inline(node),
                ParagraphContent::Tracked(change) => walker.tracked(change),
                ParagraphContent::RangeStart(_)
                | ParagraphContent::RangeEnd(_)
                | ParagraphContent::CommentRange(_) => {}
            },
        );
        if let Some(info) = &mut paragraph.p_pr_ins {
            self.at("pPrIns", |walker| walker.id(&mut info.id));
        }
        if let Some(info) = &mut paragraph.p_pr_del {
            self.at("pPrDel", |walker| walker.id(&mut info.id));
        }
        self.field_entries(
            "propertyChanges",
            paragraph.property_changes.iter_mut().flatten(),
            |walker, change| walker.info(&mut change.info.id),
        );
    }

    fn tracked(&mut self, change: &'a mut TrackedInline) {
        self.info(&mut change.info.id);
        self.inlines("content", &mut change.content);
    }

    fn inlines(&mut self, key: &str, nodes: &'a mut [InlineNode]) {
        self.field_entries(key, nodes, Self::inline);
    }

    fn inline(&mut self, node: &'a mut InlineNode) {
        match node {
            InlineNode::Run(run) => self.run(run),
            InlineNode::Tracked(change) => self.tracked(change),
            InlineNode::Hyperlink(link) => {
                let [children_key, structured_key] = HYPERLINK_FIELDS;
                self.inlines(children_key, &mut link.children);
                if let Some(children) = &mut link.structured_children {
                    self.inlines(structured_key, children);
                }
            }
            InlineNode::InlineSdt(control) => self.inlines("content", &mut control.content),
            InlineNode::SimpleField(field) => {
                self.field_entries("content", &mut field.content, Self::run);
                self.structured("structuredResult", field.structured_result.as_mut());
                if let Some(tree) = &mut field.field_tree {
                    self.at("fieldTree", |walker| walker.field_tree(tree));
                }
            }
            InlineNode::ComplexField(field) => {
                self.field_entries("fieldCode", &mut field.field_code, Self::run);
                self.field_entries("fieldResult", &mut field.field_result, Self::run);
                self.structured("structuredCode", field.structured_code.as_mut());
                self.structured("structuredResult", field.structured_result.as_mut());
                if let Some(tree) = &mut field.field_tree {
                    self.at("fieldTree", |walker| walker.field_tree(tree));
                }
            }
            InlineNode::BookmarkStart(_)
            | InlineNode::BookmarkEnd(_)
            | InlineNode::Math(_)
            | InlineNode::RawXml(_) => {}
        }
    }

    fn structured(&mut self, key: &str, field: Option<&'a mut StructuredFieldContent>) {
        if let Some(field) = field {
            let [inline_key, blocks_key] = STRUCTURED_FIELDS;
            self.at(key, |walker| {
                if let Some(nodes) = &mut field.inline {
                    walker.inlines(inline_key, nodes);
                }
                if let Some(blocks) = &mut field.blocks {
                    walker.at(blocks_key, |walker| walker.blocks(blocks));
                }
            });
        }
    }

    fn field_tree(&mut self, tree: &'a mut StructuredFieldTree) {
        let [code_key, result_key, children_key] = FIELD_TREE_FIELDS;
        self.structured(code_key, tree.code.as_mut());
        self.structured(result_key, tree.result.as_mut());
        self.field_entries(
            children_key,
            tree.children.iter_mut().flatten(),
            Self::field_tree,
        );
    }

    fn run(&mut self, run: &'a mut Run) {
        if matches!(
            run.content.as_slice(),
            [RunContent::Shape { .. } | RunContent::OpaqueDrawing { .. }]
        ) {
            self.run_content(&mut run.content);
            self.run_changes(&mut run.property_changes);
        } else {
            self.run_changes(&mut run.property_changes);
            self.run_content(&mut run.content);
        }
    }

    fn run_changes(&mut self, changes: &'a mut Option<Vec<RunPropertyChange>>) {
        self.field_entries(
            "propertyChanges",
            changes.iter_mut().flatten(),
            |walker, change| {
                walker.info(&mut change.info.id);
            },
        );
    }

    fn run_content(&mut self, content: &'a mut [RunContent]) {
        self.field_entries("content", content, |walker, child| match child {
            RunContent::Shape { shape, .. } => walker.at("shape", |walker| walker.shape(shape)),
            RunContent::OpaqueDrawing {
                text_box: Some(shape),
                ..
            } => walker.at("textBox", |walker| walker.shape(shape)),
            _ => {}
        });
    }

    fn shape(&mut self, shape: &'a mut Shape) {
        if let Some(model) = &mut shape.drawing_model {
            self.at("drawingModel", |walker| walker.model(model, ModelKind::Any));
        }
        if let Some(body) = &mut shape.text_body {
            self.at("textBody", |walker| {
                walker.field_entries("content", &mut body.content, |walker, block| {
                    walker.model(block, ModelKind::Any);
                });
            });
        }
        self.field_entries("children", shape.children.iter_mut().flatten(), Self::shape);
        if let Some(root) = shape.scene.as_mut().and_then(|scene| scene.root.as_mut()) {
            self.at("scene", |walker| {
                walker.at("root", |walker| walker.scene(root))
            });
        }
    }

    fn scene(&mut self, node: &'a mut DrawingSceneNode) {
        let [shape_key, children_key] = SCENE_NODE_FIELDS;
        if let Some(shape) = &mut node.shape {
            self.at(shape_key, |walker| walker.shape(shape));
        }
        self.field_entries(
            children_key,
            node.children.iter_mut().flatten(),
            Self::scene,
        );
    }

    fn model(&mut self, model: &'a mut Value, kind: ModelKind) {
        match model {
            Value::Object(record) => {
                if let Some(placeholder) = self.session.model_placeholder(record) {
                    self.push(placeholder, Site::Model(record.get_mut("id").unwrap()));
                    return;
                }
                let kind = kind.for_record(record);
                let order = kind.fields(record);
                if order.is_empty() {
                    for (key, value) in record {
                        self.model_field(kind, key, value);
                    }
                } else {
                    let prefix = std::mem::take(&mut self.order);
                    let mut fields = [const { None }; 5];
                    for (key, value) in record {
                        if let Some(index) = order.iter().position(|name| *name == key.as_str()) {
                            fields[index] = Some((key, value));
                        } else {
                            self.model_field(kind, key, value);
                        }
                    }
                    let extra = std::mem::replace(&mut self.order, prefix);
                    for (key, value) in fields.into_iter().flatten() {
                        self.model_field(kind, key, value);
                    }
                    self.append(extra);
                }
            }
            Value::Array(items) => {
                self.entries(items, |walker, item| walker.model(item, kind));
            }
            Value::Null | Value::Bool(_) | Value::Number(_) | Value::String(_) => {}
        }
    }

    fn model_field(&mut self, parent: ModelKind, key: &str, value: &'a mut Value) {
        self.at(key, |walker| walker.model(value, parent.child(key)));
    }
}

#[derive(Clone, Copy)]
enum ModelKind {
    Any,
    Paragraph,
    Table,
    Row,
    Cell,
    Tracked,
    Run,
    Hyperlink,
    SimpleField,
    ComplexField,
    Structured,
    FieldTree,
    Shape,
    Scene,
    SceneNode,
}

impl ModelKind {
    fn for_record(self, record: &serde_json::Map<String, Value>) -> Self {
        if !matches!(self, Self::Any) {
            return self;
        }
        match record.get("type").and_then(Value::as_str) {
            Some("paragraph") => Self::Paragraph,
            Some("table") => Self::Table,
            Some("tableRow") => Self::Row,
            Some("tableCell") => Self::Cell,
            Some("insertion" | "deletion" | "moveFrom" | "moveTo") => Self::Tracked,
            Some("run") => Self::Run,
            Some("hyperlink") => Self::Hyperlink,
            Some("simpleField") => Self::SimpleField,
            Some("complexField") => Self::ComplexField,
            Some("shape") => Self::Shape,
            _ => Self::Any,
        }
    }

    fn fields(self, record: &serde_json::Map<String, Value>) -> &'static [&'static str] {
        match self {
            Self::Paragraph => &["content", "pPrIns", "pPrDel", "propertyChanges"],
            Self::Table => &["rows", "propertyChanges"],
            Self::Row => &["cells", "structuralChange", "propertyChanges"],
            Self::Cell => &["content", "structuralChange", "propertyChanges"],
            Self::Tracked => &["info", "content"],
            Self::Run => {
                let drawing = record
                    .get("content")
                    .and_then(Value::as_array)
                    .and_then(|content| match content.as_slice() {
                        [child] => child.get("type").and_then(Value::as_str),
                        _ => None,
                    });
                if matches!(drawing, Some("shape" | "opaqueDrawing")) {
                    &["content", "propertyChanges"]
                } else {
                    &["propertyChanges", "content"]
                }
            }
            Self::Hyperlink => &HYPERLINK_FIELDS,
            Self::SimpleField => &["content", "structuredResult", "fieldTree"],
            Self::ComplexField => &[
                "fieldCode",
                "fieldResult",
                "structuredCode",
                "structuredResult",
                "fieldTree",
            ],
            Self::Structured => &STRUCTURED_FIELDS,
            Self::FieldTree => &FIELD_TREE_FIELDS,
            Self::Shape => &["drawingModel", "textBody", "children", "scene"],
            Self::SceneNode => &SCENE_NODE_FIELDS,
            Self::Any | Self::Scene => &[],
        }
    }

    fn child(self, key: &str) -> Self {
        match (self, key) {
            (Self::SimpleField | Self::ComplexField, "structuredCode" | "structuredResult")
            | (Self::FieldTree, "code" | "result") => Self::Structured,
            (Self::SimpleField | Self::ComplexField, "fieldTree")
            | (Self::FieldTree, "children") => Self::FieldTree,
            (Self::Shape, "scene") => Self::Scene,
            (Self::Scene, "root") | (Self::SceneNode, "children") => Self::SceneNode,
            _ => Self::Any,
        }
    }
}

#[cfg(test)]
mod tests;
