use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::fmt::{self, Write};
use std::ptr;
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
use serde::Serialize;
use serde::ser::{
    self, SerializeMap, SerializeSeq, SerializeStruct, SerializeStructVariant, SerializeTuple,
    SerializeTupleStruct, SerializeTupleVariant, Serializer,
};
use serde_json::{Map, Value};

const MAX_ID: u32 = 1 << 31;
const PLACEHOLDER_LIMIT: f64 = (1u64 << 53) as f64;
const NO_SITE: usize = usize::MAX;
const HYPERLINK_FIELDS: [&str; 2] = ["children", "structuredChildren"];
const STRUCTURED_FIELDS: [&str; 2] = ["inline", "blocks"];
const FIELD_TREE_FIELDS: [&str; 3] = ["code", "result", "children"];
const SCENE_NODE_FIELDS: [&str; 2] = ["shape", "children"];

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

impl ser::Error for RevisionIdsError {
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
    assign(request, session, None, || {})
}

#[doc(hidden)]
pub fn assign_session_revision_ids_with_paths(
    request: &mut S13SaveRequest,
    session: &BTreeSet<u64>,
    paths: &mut Vec<(String, u64)>,
) -> Result<BTreeMap<u64, u32>, RevisionIdsError> {
    paths.clear();
    assign(request, session, Some(paths), || {})
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
    let ids = assign(request, session, None, || {
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
    mut pass_complete: impl FnMut(),
) -> Result<BTreeMap<u64, u32>, RevisionIdsError> {
    if session.is_empty() {
        return Ok(BTreeMap::new());
    }
    let session_ids = SessionIds::new(session);
    let mut found = collect_with::<false, _>(request, &session_ids)?;
    pass_complete();
    if found.sites.is_empty() {
        return Ok(BTreeMap::new());
    }
    let mut walker = SiteWalker::new(
        &session_ids,
        paths,
        &found.blocks_with_sites,
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
    for site in &mut walker.sites {
        site.write();
    }
    let mut assigned = Vec::with_capacity(found.sites.len());
    assigned.extend(
        session
            .iter()
            .filter_map(|id| found.sites.get(id).map(|count| (*id, count.assigned))),
    );
    let ids = assigned.into_iter().collect();
    pass_complete();

    if let Some(id) = walker
        .sites
        .iter()
        .filter_map(PendingSite::unassigned)
        .chain(found.sites.values().filter_map(|count| {
            count
                .hidden
                .filter(|_| !session_ids.contains(u64::from(count.assigned)))
        }))
        .min()
    {
        return Err(RevisionIdsError::Unassigned(id));
    }
    pass_complete();
    Ok(ids)
}

struct SessionIds {
    first: u64,
    ids: FastSet<u64>,
}

impl SessionIds {
    fn new(session: &BTreeSet<u64>) -> Self {
        Self {
            first: session.first().copied().unwrap_or(u64::MAX),
            ids: session.iter().copied().collect(),
        }
    }

    fn placeholder(&self, number: f64) -> Option<u64> {
        placeholder(number).filter(|id| self.contains(*id))
    }

    fn contains(&self, id: u64) -> bool {
        id >= self.first && self.ids.contains(&id)
    }
}

#[doc(hidden)]
pub struct RevisionIdReservations {
    pub used: BTreeSet<u32>,
    pub max: u32,
    pub sites: BTreeMap<(String, u32), usize>,
}

#[doc(hidden)]
pub fn session_revision_id_reservations(
    request: &S13SaveRequest,
    session: &BTreeSet<u64>,
) -> Result<RevisionIdReservations, RevisionIdsError> {
    let found = collect_with::<true, _>(request, &SessionIds::new(session))?;
    Ok(RevisionIdReservations {
        used: found.used.into_iter().collect(),
        max: found.max,
        sites: found.reservations,
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
    hidden: Option<u64>,
}

#[derive(Default)]
struct Collected {
    max: u32,
    used: FastSet<u32>,
    sites: FastMap<u64, SiteCount>,
    reservations: BTreeMap<(String, u32), usize>,
    pending: Vec<Observation>,
    pending_depth: usize,
    pointer: Option<String>,
    site_count: usize,
    blocks_with_sites: FastSet<*const ()>,
}

enum Observation {
    Reservation(u32, bool, Option<String>),
    Site(u64, Option<u64>),
}

impl Collected {
    fn begin_pointer(&mut self, key: &str) -> Option<usize> {
        self.pointer.as_mut().map(|pointer| {
            let length = pointer.len();
            push_pointer(pointer, key);
            length
        })
    }

    fn end_pointer(&mut self, start: Option<usize>) {
        if let Some(start) = start {
            self.pointer.as_mut().unwrap().truncate(start);
        }
    }

    fn note(&mut self, id: Kind, raises_base: bool, key: Option<&str>) {
        if let Kind::Number(id) = id
            && (0.0..f64::from(MAX_ID)).contains(&id)
            && id.fract() == 0.0
        {
            let pointer = self.pointer.as_mut().map(|pointer| {
                let length = pointer.len();
                if let Some(key) = key {
                    push_pointer(pointer, key);
                }
                let reservation = pointer.clone();
                pointer.truncate(length);
                reservation
            });
            self.observe(Observation::Reservation(id as u32, raises_base, pointer));
        }
    }

    fn observe(&mut self, observation: Observation) {
        if self.pending_depth == 0 {
            self.record(observation);
        } else {
            self.pending.push(observation);
        }
    }

    fn record(&mut self, observation: Observation) {
        match observation {
            Observation::Reservation(id, raises_base, pointer) => {
                #[cfg(test)]
                count_work(|work| work.reservation_inserts += 1);
                self.used.insert(id);
                if let Some(pointer) = pointer {
                    *self.reservations.entry((pointer, id)).or_default() += 1;
                }
                if raises_base {
                    self.max = self.max.max(id);
                }
            }
            Observation::Site(id, hidden) => {
                #[cfg(test)]
                count_work(|work| work.site_inserts += 1);
                self.site_count += 1;
                let count = self.sites.entry(id).or_default();
                count.remaining += 1;
                if let Some(hidden) = hidden {
                    count.hidden = Some(count.hidden.map_or(hidden, |prior| prior.min(hidden)));
                }
            }
        }
    }

    fn finish_info(&mut self, start: usize, session_id: Option<u64>) {
        self.pending_depth -= 1;
        if let Some(id) = session_id {
            let hidden = self
                .pending
                .drain(start..)
                .filter_map(|observation| match observation {
                    Observation::Site(id, _) => Some(id),
                    Observation::Reservation(..) => None,
                })
                .min();
            self.observe(Observation::Site(id, hidden));
        }
        if self.pending_depth == 0 {
            while let Some(observation) = self.pending.pop() {
                self.record(observation);
            }
        }
    }

    fn scan_xml(&mut self, text: &str) {
        #[cfg(test)]
        count_work(|work| {
            work.strings += 1;
            work.string_bytes += text.len();
        });
        let bytes = text.as_bytes();
        if bytes.len() < 8 {
            return;
        }
        for index in memchr::memchr_iter(b'w', bytes) {
            if !bytes[index..].starts_with(b"w:id")
                || (index > 0
                    && (bytes[index - 1].is_ascii_alphanumeric() || bytes[index - 1] == b'_'))
            {
                continue;
            }
            let tail = text[index + 4..].trim_start_matches(js_space);
            let Some(tail) = tail.strip_prefix('=') else {
                continue;
            };
            let Some(tail) = tail.trim_start_matches(js_space).strip_prefix(['\'', '"']) else {
                continue;
            };
            let digits = tail.bytes().take_while(u8::is_ascii_digit).count();
            if matches!(tail.as_bytes().get(digits), Some(b'\'' | b'"'))
                && let Ok(id) = tail[..digits].parse::<u32>()
            {
                self.note(Kind::Number(f64::from(id)), true, None);
            }
        }
    }
}

fn js_space(character: char) -> bool {
    matches!(
        character,
        '\u{0009}'..='\u{000D}'
            | ' '
            | '\u{00A0}'
            | '\u{1680}'
            | '\u{2000}'..='\u{200A}'
            | '\u{2028}'..='\u{2029}'
            | '\u{202F}'
            | '\u{205F}'
            | '\u{3000}'
            | '\u{FEFF}'
    )
}

#[derive(Clone, Copy, Default)]
enum Symbol {
    #[default]
    Other,
    Author,
    Id,
    RevisionId,
    Content,
    Type,
    Info,
    CommentMarker,
    Annotation,
}

impl From<&str> for Symbol {
    fn from(text: &str) -> Self {
        match text {
            "author" => Self::Author,
            "id" => Self::Id,
            "revisionId" => Self::RevisionId,
            "content" => Self::Content,
            "type" => Self::Type,
            "info" | "pPrIns" | "pPrDel" => Self::Info,
            "commentRangeStart" | "commentRangeEnd" | "commentReference" => Self::CommentMarker,
            "bookmarkStart" | "bookmarkEnd" | "moveFromRangeStart" | "moveFromRangeEnd"
            | "moveToRangeStart" | "moveToRangeEnd" => Self::Annotation,
            _ => Self::Other,
        }
    }
}

#[derive(Default)]
enum Kind {
    #[default]
    Other,
    String(Symbol),
    Number(f64),
    Array,
}

#[cfg(test)]
fn collect<T: Serialize>(
    value: &T,
    session: &BTreeSet<u64>,
) -> Result<Collected, RevisionIdsError> {
    collect_with::<false, _>(value, &SessionIds::new(session))
}

fn collect_with<const PATHS: bool, T: Serialize>(
    value: &T,
    session: &SessionIds,
) -> Result<Collected, RevisionIdsError> {
    let mut found = Collected {
        pointer: PATHS.then(String::new),
        ..Collected::default()
    };
    value.serialize(Visitor::<PATHS> {
        session,
        found: &mut found,
        model: false,
        info: false,
        key: false,
        node_type: false,
    })?;
    Ok(found)
}

struct Visitor<'a, const PATHS: bool> {
    session: &'a SessionIds,
    found: &'a mut Collected,
    model: bool,
    info: bool,
    key: bool,
    node_type: bool,
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

impl<const PATHS: bool> Visitor<'_, PATHS> {
    fn at<T>(&mut self, key: &str, visit: impl FnOnce(&mut Self) -> T) -> T {
        if !PATHS {
            return visit(self);
        }
        let start = self.found.begin_pointer(key);
        let value = visit(self);
        self.found.end_pointer(start);
        value
    }

    fn child(&mut self, key: Symbol, model: bool) -> Visitor<'_, PATHS> {
        Visitor {
            session: self.session,
            found: self.found,
            model,
            info: model && matches!(key, Symbol::Info),
            key: false,
            node_type: matches!(key, Symbol::Type),
        }
    }

    fn number(self, value: impl fmt::Display, number: f64) -> Kind {
        if self.key {
            if PATHS && let Some(pointer) = &mut self.found.pointer {
                write!(pointer, "/{value}").unwrap();
            }
            Kind::Other
        } else {
            Kind::Number(number)
        }
    }
}

struct Record<'a, const PATHS: bool> {
    visitor: Visitor<'a, PATHS>,
    kind: RecordKind,
    key: Symbol,
    key_start: Option<usize>,
    variant_start: Option<usize>,
    pending_start: Option<usize>,
    block: Option<(*const (), usize)>,
    author: bool,
    id: Kind,
    revision_id: Kind,
    content_array: bool,
    node_type: Symbol,
}

#[derive(Clone, Copy)]
enum RecordKind {
    Other,
    Block,
    ShapeTextBody,
    Shape,
    TablePatch,
}

impl<'a, const PATHS: bool> Record<'a, PATHS> {
    fn new(visitor: Visitor<'a, PATHS>, name: &'static str) -> Self {
        let kind = match name {
            "Paragraph" | "Table" | "BlockSdt" => RecordKind::Block,
            "ShapeTextBody" => RecordKind::ShapeTextBody,
            "Shape" => RecordKind::Shape,
            "S13TablePatch" => RecordKind::TablePatch,
            _ => RecordKind::Other,
        };
        let info = visitor.info || matches!(name, "TrackedChangeInfo" | "PropertyChangeInfo");
        let pending_start = info.then(|| {
            visitor.found.pending_depth += 1;
            visitor.found.pending.len()
        });
        let block = (!PATHS && matches!(kind, RecordKind::Block))
            .then(|| (ptr::null(), visitor.found.site_count));
        Self {
            visitor,
            kind,
            key: Symbol::Other,
            key_start: None,
            variant_start: None,
            pending_start,
            block,
            author: false,
            id: Kind::Other,
            revision_id: Kind::Other,
            content_array: false,
            node_type: Symbol::Other,
        }
    }

    fn field(&mut self, key: Symbol, kind: Kind) {
        match key {
            Symbol::Author => self.author = matches!(kind, Kind::String(_)),
            Symbol::Id => self.id = kind,
            Symbol::RevisionId => self.revision_id = kind,
            Symbol::Content => self.content_array = matches!(kind, Kind::Array),
            Symbol::Type => {
                self.node_type = match kind {
                    Kind::String(symbol) => symbol,
                    _ => Symbol::Other,
                };
            }
            _ => {}
        }
    }

    fn named_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), RevisionIdsError> {
        #[cfg(test)]
        count_work(|work| work.fields += 1);
        let symbol = Symbol::from(key);
        if matches!(symbol, Symbol::Type) {
            #[cfg(test)]
            if matches!(self.kind, RecordKind::Block) {
                count_work(|work| work.collected_blocks += 1);
            }
            if let Some((key, _)) = &mut self.block {
                *key = ptr::from_ref(value).cast::<()>();
            }
        }
        let model = self.visitor.model
            || matches!(
                (self.kind, key),
                (RecordKind::ShapeTextBody, "content")
                    | (RecordKind::Shape, "drawingModel")
                    | (RecordKind::TablePatch, "before" | "after")
            );
        let kind = self
            .visitor
            .at(key, |visitor| value.serialize(visitor.child(symbol, model)))?;
        self.field(symbol, kind);
        Ok(())
    }

    fn finish(self) -> Kind {
        let found = self.visitor.found;
        if let Some((key, site_count)) = self.block
            && !key.is_null()
            && found.site_count != site_count
        {
            #[cfg(test)]
            count_work(|work| work.block_inserts += 1);
            found.blocks_with_sites.insert(key);
        }
        if let Some(start) = self.pending_start {
            let session_id = match (self.author, &self.id) {
                (true, Kind::Number(id)) => self.visitor.session.placeholder(*id),
                _ => None,
            };
            found.finish_info(start, session_id);
            if session_id.is_some() {
                if PATHS {
                    found.end_pointer(self.variant_start);
                }
                return Kind::Other;
            }
        }
        if self.author {
            let raises_base = !self.content_array;
            found.note(self.id, raises_base, Some("id"));
            found.note(self.revision_id, raises_base, Some("revisionId"));
        } else if let symbol @ (Symbol::CommentMarker | Symbol::Annotation) = self.node_type {
            found.note(self.id, matches!(symbol, Symbol::Annotation), Some("id"));
        }
        if PATHS {
            found.end_pointer(self.variant_start);
        }
        Kind::Other
    }
}

impl<const PATHS: bool> SerializeMap for Record<'_, PATHS> {
    type Ok = Kind;
    type Error = RevisionIdsError;

    fn serialize_key<T: Serialize + ?Sized>(&mut self, key: &T) -> Result<(), Self::Error> {
        if PATHS {
            self.key_start = self.visitor.found.pointer.as_ref().map(String::len);
        }
        let kind = key.serialize(Visitor::<PATHS> {
            session: self.visitor.session,
            found: self.visitor.found,
            model: false,
            info: false,
            key: true,
            node_type: false,
        })?;
        self.key = match kind {
            Kind::String(symbol) => symbol,
            _ => Symbol::Other,
        };
        Ok(())
    }

    fn serialize_value<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Self::Error> {
        #[cfg(test)]
        count_work(|work| work.fields += 1);
        let model = self.visitor.model;
        let kind = value.serialize(self.visitor.child(self.key, model))?;
        if PATHS {
            self.visitor.found.end_pointer(self.key_start);
        }
        self.field(self.key, kind);
        Ok(())
    }

    fn end(self) -> Result<Kind, Self::Error> {
        Ok(self.finish())
    }
}

impl<const PATHS: bool> SerializeStruct for Record<'_, PATHS> {
    type Ok = Kind;
    type Error = RevisionIdsError;

    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), Self::Error> {
        self.named_field(key, value)
    }

    fn end(self) -> Result<Kind, Self::Error> {
        Ok(self.finish())
    }
}

impl<const PATHS: bool> SerializeStructVariant for Record<'_, PATHS> {
    type Ok = Kind;
    type Error = RevisionIdsError;

    fn serialize_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), Self::Error> {
        self.named_field(key, value)
    }

    fn end(self) -> Result<Kind, Self::Error> {
        Ok(self.finish())
    }
}

struct Sequence<'a, const PATHS: bool> {
    visitor: Visitor<'a, PATHS>,
    index: usize,
    variant_start: Option<usize>,
}

macro_rules! sequence {
    ($trait:ident, $field:ident, $kind:expr) => {
        impl<const PATHS: bool> $trait for Sequence<'_, PATHS> {
            type Ok = Kind;
            type Error = RevisionIdsError;

            fn $field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Self::Error> {
                #[cfg(test)]
                count_work(|work| work.fields += 1);
                let model = self.visitor.model;
                if !PATHS {
                    value.serialize(self.visitor.child(Symbol::Other, model))?;
                    return Ok(());
                }
                let start = self.visitor.found.pointer.as_mut().map(|pointer| {
                    let start = pointer.len();
                    write!(pointer, "/{}", self.index).unwrap();
                    self.index += 1;
                    start
                });
                value.serialize(self.visitor.child(Symbol::Other, model))?;
                self.visitor.found.end_pointer(start);
                Ok(())
            }

            fn end(self) -> Result<Kind, Self::Error> {
                if PATHS {
                    self.visitor.found.end_pointer(self.variant_start);
                }
                Ok($kind)
            }
        }
    };
}

sequence!(SerializeSeq, serialize_element, Kind::Array);
sequence!(SerializeTuple, serialize_element, Kind::Array);
sequence!(SerializeTupleStruct, serialize_field, Kind::Array);
sequence!(SerializeTupleVariant, serialize_field, Kind::Other);

macro_rules! number {
    ($($method:ident: $type:ty),+ $(,)?) => {
        $(
            fn $method(self, value: $type) -> Result<Kind, Self::Error> {
                Ok(self.number(value, value as f64))
            }
        )+
    };
}

impl<'a, const PATHS: bool> Serializer for Visitor<'a, PATHS> {
    type Ok = Kind;
    type Error = RevisionIdsError;
    type SerializeSeq = Sequence<'a, PATHS>;
    type SerializeTuple = Sequence<'a, PATHS>;
    type SerializeTupleStruct = Sequence<'a, PATHS>;
    type SerializeTupleVariant = Sequence<'a, PATHS>;
    type SerializeMap = Record<'a, PATHS>;
    type SerializeStruct = Record<'a, PATHS>;
    type SerializeStructVariant = Record<'a, PATHS>;

    number!(
        serialize_i8: i8, serialize_i16: i16, serialize_i32: i32, serialize_i64: i64,
        serialize_i128: i128, serialize_u8: u8, serialize_u16: u16, serialize_u32: u32,
        serialize_u64: u64, serialize_u128: u128, serialize_f32: f32, serialize_f64: f64,
    );

    fn serialize_bool(self, value: bool) -> Result<Kind, Self::Error> {
        if PATHS
            && self.key
            && let Some(pointer) = &mut self.found.pointer
        {
            write!(pointer, "/{value}").unwrap();
        }
        Ok(Kind::Other)
    }

    fn serialize_char(self, value: char) -> Result<Kind, Self::Error> {
        self.serialize_str(value.encode_utf8(&mut [0; 4]))
    }

    fn serialize_str(self, value: &str) -> Result<Kind, Self::Error> {
        if self.key {
            if PATHS {
                self.found.begin_pointer(value);
            }
            return Ok(Kind::String(Symbol::from(value)));
        }
        self.found.scan_xml(value);
        Ok(Kind::String(if self.node_type {
            Symbol::from(value)
        } else {
            Symbol::Other
        }))
    }

    fn serialize_bytes(self, _: &[u8]) -> Result<Kind, Self::Error> {
        Ok(Kind::Other)
    }

    fn serialize_none(self) -> Result<Kind, Self::Error> {
        Ok(Kind::Other)
    }

    fn serialize_some<T: Serialize + ?Sized>(self, value: &T) -> Result<Kind, Self::Error> {
        value.serialize(self)
    }

    fn serialize_unit(self) -> Result<Kind, Self::Error> {
        Ok(Kind::Other)
    }

    fn serialize_unit_struct(self, _: &'static str) -> Result<Kind, Self::Error> {
        self.serialize_unit()
    }

    fn serialize_unit_variant(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
    ) -> Result<Kind, Self::Error> {
        self.serialize_str(variant)
    }

    fn serialize_newtype_struct<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        value: &T,
    ) -> Result<Kind, Self::Error> {
        value.serialize(self)
    }

    fn serialize_newtype_variant<T: Serialize + ?Sized>(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
        value: &T,
    ) -> Result<Kind, Self::Error> {
        if !PATHS {
            value.serialize(self)?;
            return Ok(Kind::Other);
        }
        let start = self.found.begin_pointer(variant);
        value.serialize(Visitor {
            found: &mut *self.found,
            ..self
        })?;
        self.found.end_pointer(start);
        Ok(Kind::Other)
    }

    fn serialize_seq(self, _: Option<usize>) -> Result<Self::SerializeSeq, Self::Error> {
        Ok(Sequence {
            visitor: self,
            index: 0,
            variant_start: None,
        })
    }

    fn serialize_tuple(self, _: usize) -> Result<Self::SerializeTuple, Self::Error> {
        self.serialize_seq(None)
    }

    fn serialize_tuple_struct(
        self,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleStruct, Self::Error> {
        self.serialize_seq(None)
    }

    fn serialize_tuple_variant(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleVariant, Self::Error> {
        let variant_start = if PATHS {
            self.found.begin_pointer(variant)
        } else {
            None
        };
        Ok(Sequence {
            visitor: self,
            index: 0,
            variant_start,
        })
    }

    fn serialize_map(self, _: Option<usize>) -> Result<Self::SerializeMap, Self::Error> {
        Ok(Record::new(self, ""))
    }

    fn serialize_struct(
        self,
        name: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStruct, Self::Error> {
        Ok(Record::new(self, name))
    }

    fn serialize_struct_variant(
        self,
        _: &'static str,
        _: u32,
        variant: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStructVariant, Self::Error> {
        let variant_start = if PATHS {
            self.found.begin_pointer(variant)
        } else {
            None
        };
        let mut record = Record::new(
            Visitor {
                info: false,
                ..self
            },
            "",
        );
        record.variant_start = variant_start;
        Ok(record)
    }
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

    fn unassigned(&self) -> Option<u64> {
        #[cfg(test)]
        count_work(|work| work.verified += 1);
        let assigned = match &self.site {
            Site::Typed(id) => **id == f64::from(self.assigned),
            Site::Model(id) => id.as_u64() == Some(u64::from(self.assigned)),
        };
        (!assigned).then_some(self.placeholder)
    }
}

fn model_info_id(record: &Map<String, Value>) -> Option<f64> {
    if record.get("author").is_some_and(Value::is_string) {
        record.get("id").and_then(Value::as_f64)
    } else {
        None
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
    blocks_with_sites: &'s FastSet<*const ()>,
    sites: Vec<PendingSite<'a>>,
    order: SiteOrder,
    paths: Option<&'p mut Vec<(String, u64)>>,
    path: String,
}

impl<'a, 's, 'p> SiteWalker<'a, 's, 'p> {
    fn new(
        session: &'s SessionIds,
        mut paths: Option<&'p mut Vec<(String, u64)>>,
        blocks_with_sites: &'s FastSet<*const ()>,
        site_count: usize,
    ) -> Self {
        if let Some(paths) = &mut paths {
            paths.reserve(site_count);
        }
        Self {
            session,
            blocks_with_sites,
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
        visit(self);
        self.path.truncate(length);
    }

    fn entries<T>(
        &mut self,
        values: impl IntoIterator<Item = T>,
        mut visit: impl FnMut(&mut Self, T),
    ) {
        if self.paths.is_some() {
            for (index, value) in values.into_iter().enumerate() {
                let length = self.path.len();
                write!(self.path, "/{index}").unwrap();
                visit(self, value);
                self.path.truncate(length);
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
                walker.blocks(&mut request.document.content)
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
                                walker.model(&mut table.before, ModelKind::Any, false);
                            });
                            walker.at("after", |walker| {
                                walker.model(&mut table.after, ModelKind::Any, false);
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

    fn block_mut<T: Clone>(
        &mut self,
        block: &'a mut Arc<T>,
        node_type: impl FnOnce(&T) -> &String,
    ) -> Option<&'a mut T> {
        #[cfg(test)]
        count_work(|work| work.block_probes += 1);
        let key = ptr::from_ref(node_type(block.as_ref())).cast::<()>();
        if !self.blocks_with_sites.contains(&key) {
            return None;
        }
        #[cfg(test)]
        count_work(|work| work.make_mut += 1);
        Some(Arc::make_mut(block))
    }

    fn blocks(&mut self, blocks: &'a mut [BlockContent]) {
        self.entries(blocks, |walker, block| match block {
            BlockContent::Paragraph(paragraph) => {
                if let Some(paragraph) = walker.block_mut(paragraph, |block| &block.node_type) {
                    walker.paragraph(paragraph);
                }
            }
            BlockContent::Table(table) => {
                if let Some(table) = walker.block_mut(table, |block| &block.node_type) {
                    walker.table(table);
                }
            }
            BlockContent::BlockSdt(control) => {
                if let Some(control) = walker.block_mut(control, |block| &block.node_type) {
                    walker.at("content", |walker| walker.blocks(&mut control.content));
                }
            }
            BlockContent::RawXml(_) => {}
        });
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
            self.at("drawingModel", |walker| {
                walker.model(model, ModelKind::Any, false)
            });
        }
        if let Some(body) = &mut shape.text_body {
            self.at("textBody", |walker| {
                walker.field_entries("content", &mut body.content, |walker, block| {
                    walker.model(block, ModelKind::Any, false);
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

    fn model(&mut self, model: &'a mut Value, kind: ModelKind, info: bool) {
        match model {
            Value::Object(record) => {
                if info
                    && let Some(placeholder) =
                        model_info_id(record).and_then(|id| self.session.placeholder(id))
                {
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
                self.entries(items, |walker, item| walker.model(item, kind, false));
            }
            _ => {}
        }
    }

    fn model_field(&mut self, parent: ModelKind, key: &str, value: &'a mut Value) {
        let info = matches!(Symbol::from(key), Symbol::Info);
        self.at(key, |walker| walker.model(value, parent.child(key), info));
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
