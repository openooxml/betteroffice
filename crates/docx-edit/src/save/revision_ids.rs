use std::collections::btree_map::Entry;
use std::collections::{BTreeMap, BTreeSet, HashSet};
use std::fmt::{self, Write};
use std::ptr;
use std::sync::Arc;

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
use serde_json::Value;

const MAX_ID: u32 = 1 << 31;
const PLACEHOLDER_LIMIT: f64 = (1u64 << 53) as f64;
const HYPERLINK_FIELDS: [&str; 2] = ["children", "structuredChildren"];
const STRUCTURED_FIELDS: [&str; 2] = ["inline", "blocks"];
const FIELD_TREE_FIELDS: [&str; 3] = ["code", "result", "children"];
const SCENE_NODE_FIELDS: [&str; 2] = ["shape", "children"];

#[cfg(test)]
thread_local! {
    static BLOCK_VISITS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
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
    assign(request, session, None)
}

#[doc(hidden)]
pub fn assign_session_revision_ids_with_paths(
    request: &mut S13SaveRequest,
    session: &BTreeSet<u64>,
    paths: &mut Vec<(String, u64)>,
) -> Result<BTreeMap<u64, u32>, RevisionIdsError> {
    paths.clear();
    assign(request, session, Some(paths))
}

fn assign(
    request: &mut S13SaveRequest,
    session: &BTreeSet<u64>,
    paths: Option<&mut Vec<(String, u64)>>,
) -> Result<BTreeMap<u64, u32>, RevisionIdsError> {
    if session.is_empty() {
        return Ok(BTreeMap::new());
    }
    let mut found = collect(request, session)?;
    if found.sites.is_empty() {
        return Ok(BTreeMap::new());
    }
    let mut walker = SiteWalker::new(session, paths, &found.blocks_with_sites);
    walker.request(request);

    let mut next = found.max;
    let mut ids = BTreeMap::new();
    let mut expected = BTreeMap::new();
    for (placeholder, _) in &walker.sites {
        let Some(remaining) = found.sites.get_mut(placeholder) else {
            return Err(RevisionIdsError::Unassigned(*placeholder));
        };
        *remaining = remaining
            .checked_sub(1)
            .ok_or(RevisionIdsError::Unassigned(*placeholder))?;
        let id = match ids.entry(*placeholder) {
            Entry::Occupied(entry) => *entry.get(),
            Entry::Vacant(entry) => *entry.insert(next_id(&mut next, found.max, &found.used)?),
        };
        if session.contains(&u64::from(id)) {
            *expected.entry(u64::from(id)).or_default() += 1;
        }
    }
    if let Some((&id, _)) = found.sites.iter().find(|(_, count)| **count != 0) {
        return Err(RevisionIdsError::Unassigned(id));
    }
    for (placeholder, site) in walker.sites {
        match site {
            Site::Typed(id) => *id = f64::from(ids[&placeholder]),
            Site::Model(id) => *id = Value::from(ids[&placeholder]),
        }
    }

    let rewritten = collect_with(request, session, Collection::SitesOnly)?;
    if let Some((&id, _)) = rewritten
        .sites
        .iter()
        .find(|&(id, count)| expected.get(id) != Some(count))
        .or_else(|| {
            expected
                .iter()
                .find(|&(id, count)| rewritten.sites.get(id) != Some(count))
        })
    {
        return Err(RevisionIdsError::Unassigned(id));
    }
    Ok(ids)
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
    let found = collect_with(request, session, Collection::Reservations)?;
    Ok(RevisionIdReservations {
        used: found.used,
        max: found.max,
        sites: found.reservations,
    })
}

fn next_id(next: &mut u32, max: u32, used: &BTreeSet<u32>) -> Result<u32, RevisionIdsError> {
    loop {
        *next = if *next + 1 < MAX_ID { *next + 1 } else { 1 };
        if *next == max {
            return Err(RevisionIdsError::Exhausted);
        }
        if !used.contains(next) {
            return Ok(*next);
        }
    }
}

fn placeholder(id: f64) -> Option<u64> {
    ((0.0..PLACEHOLDER_LIMIT).contains(&id) && id.fract() == 0.0).then_some(id as u64)
}

#[derive(Default)]
enum Collection {
    #[default]
    Assignment,
    Reservations,
    SitesOnly,
}

#[derive(Default)]
struct Collected {
    max: u32,
    used: BTreeSet<u32>,
    sites: BTreeMap<u64, usize>,
    reservations: BTreeMap<(String, u32), usize>,
    pending: Vec<Observation>,
    pending_depth: usize,
    collection: Collection,
    site_count: usize,
    blocks_with_sites: HashSet<*const ()>,
    #[cfg(test)]
    reservation_inserts: usize,
}

enum Observation {
    Reservation(u32, bool, Option<String>),
    Site(u64),
}

impl Collected {
    fn note(&mut self, id: Kind, raises_base: bool, pointer: Option<&str>) {
        if !matches!(self.collection, Collection::SitesOnly)
            && let Kind::Number(id) = id
            && (0.0..f64::from(MAX_ID)).contains(&id)
            && id.fract() == 0.0
        {
            self.observe(Observation::Reservation(
                id as u32,
                raises_base,
                pointer.map(str::to_owned),
            ));
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
                {
                    self.reservation_inserts += 1;
                }
                self.used.insert(id);
                if let Some(pointer) = pointer {
                    *self.reservations.entry((pointer, id)).or_default() += 1;
                }
                if raises_base {
                    self.max = self.max.max(id);
                }
            }
            Observation::Site(id) => {
                self.site_count += 1;
                *self.sites.entry(id).or_default() += 1;
            }
        }
    }

    fn finish_info(&mut self, start: usize, session_id: Option<u64>) {
        self.pending_depth -= 1;
        if let Some(id) = session_id {
            self.pending.truncate(start);
            self.observe(Observation::Site(id));
        }
        if self.pending_depth == 0 {
            while let Some(observation) = self.pending.pop() {
                self.record(observation);
            }
        }
    }

    fn scan_xml(&mut self, text: &str, pointer: Option<&str>) {
        for (index, _) in text.match_indices("w:id") {
            if index > 0
                && (text.as_bytes()[index - 1].is_ascii_alphanumeric()
                    || text.as_bytes()[index - 1] == b'_')
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
                self.note(Kind::Number(f64::from(id)), true, pointer);
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
    Key(String),
    Number(f64),
    Array,
}

fn collect<T: Serialize>(value: &T, session: &BTreeSet<u64>) -> Result<Collected, RevisionIdsError> {
    collect_with(value, session, Collection::Assignment)
}

fn collect_with<T: Serialize>(
    value: &T,
    session: &BTreeSet<u64>,
    collection: Collection,
) -> Result<Collected, RevisionIdsError> {
    let pointer = matches!(collection, Collection::Reservations).then(String::new);
    let mut found = Collected {
        collection,
        ..Collected::default()
    };
    value.serialize(Visitor {
        session,
        found: &mut found,
        model: false,
        info: false,
        scan_strings: true,
        pointer,
    })?;
    Ok(found)
}

struct Visitor<'a> {
    session: &'a BTreeSet<u64>,
    found: &'a mut Collected,
    model: bool,
    info: bool,
    scan_strings: bool,
    pointer: Option<String>,
}

fn child_pointer(pointer: Option<&str>, key: &str) -> Option<String> {
    pointer.map(|pointer| format!("{pointer}/{}", key.replace('~', "~0").replace('/', "~1")))
}

struct Record<'a> {
    visitor: Visitor<'a>,
    name: &'static str,
    key: Kind,
    pending_start: Option<usize>,
    site_count_at_start: usize,
    block_key: Option<*const ()>,
    author: Kind,
    id: Kind,
    revision_id: Kind,
    content: Kind,
    node_type: Kind,
}

impl<'a> Record<'a> {
    fn new(visitor: Visitor<'a>, name: &'static str) -> Self {
        let info = visitor.info || matches!(name, "TrackedChangeInfo" | "PropertyChangeInfo");
        let pending_start = info.then(|| {
            visitor.found.pending_depth += 1;
            visitor.found.pending.len()
        });
        Self {
            site_count_at_start: visitor.found.site_count,
            visitor,
            name,
            key: Kind::Other,
            pending_start,
            block_key: None,
            author: Kind::Other,
            id: Kind::Other,
            revision_id: Kind::Other,
            content: Kind::Other,
            node_type: Kind::Other,
        }
    }

    fn field<T: Serialize + ?Sized>(
        &mut self,
        key: Symbol,
        model: bool,
        pointer: Option<String>,
        value: &T,
    ) -> Result<(), RevisionIdsError> {
        let kind = value.serialize(Visitor {
            session: self.visitor.session,
            found: self.visitor.found,
            model,
            info: model && matches!(key, Symbol::Info),
            scan_strings: self.visitor.scan_strings,
            pointer,
        })?;
        match key {
            Symbol::Author => self.author = kind,
            Symbol::Id => self.id = kind,
            Symbol::RevisionId => self.revision_id = kind,
            Symbol::Content => self.content = kind,
            Symbol::Type => self.node_type = kind,
            _ => {}
        }
        Ok(())
    }

    fn named_field<T: Serialize + ?Sized>(
        &mut self,
        key: &'static str,
        value: &T,
    ) -> Result<(), RevisionIdsError> {
        if key == "type" && matches!(self.name, "Paragraph" | "Table" | "BlockSdt") {
            #[cfg(test)]
            BLOCK_VISITS.set(BLOCK_VISITS.get() + 1);
            if matches!(self.visitor.found.collection, Collection::Assignment) {
                self.block_key = Some(ptr::from_ref(value).cast::<()>());
            }
        }
        let model = self.visitor.model
            || matches!(
                (self.name, key),
                ("ShapeTextBody", "content")
                    | ("Shape", "drawingModel")
                    | ("S13TablePatch", "before" | "after")
            );
        let pointer = child_pointer(self.visitor.pointer.as_deref(), key);
        self.field(Symbol::from(key), model, pointer, value)
    }

    fn finish(self) -> Kind {
        let found = self.visitor.found;
        if let Some(key) = self.block_key
            && found.site_count != self.site_count_at_start
        {
            found.blocks_with_sites.insert(key);
        }
        if let Some(start) = self.pending_start {
            let session_id = match (&self.author, &self.id) {
                (Kind::String(_), Kind::Number(id)) => {
                    placeholder(*id).filter(|id| self.visitor.session.contains(id))
                }
                _ => None,
            };
            found.finish_info(start, session_id);
            if session_id.is_some() {
                return Kind::Other;
            }
        }
        let pointer = self.visitor.pointer.as_deref();
        if matches!(self.author, Kind::String(_)) {
            let raises_base = !matches!(self.content, Kind::Array);
            found.note(
                self.id,
                raises_base,
                child_pointer(pointer, "id").as_deref(),
            );
            found.note(
                self.revision_id,
                raises_base,
                child_pointer(pointer, "revisionId").as_deref(),
            );
        } else if let Kind::String(symbol @ (Symbol::CommentMarker | Symbol::Annotation)) =
            self.node_type
        {
            found.note(
                self.id,
                matches!(symbol, Symbol::Annotation),
                child_pointer(pointer, "id").as_deref(),
            );
        }
        Kind::Other
    }
}

impl SerializeMap for Record<'_> {
    type Ok = Kind;
    type Error = RevisionIdsError;

    fn serialize_key<T: Serialize + ?Sized>(&mut self, key: &T) -> Result<(), Self::Error> {
        self.key = key.serialize(Visitor {
            session: self.visitor.session,
            found: self.visitor.found,
            model: false,
            info: false,
            scan_strings: false,
            pointer: self.visitor.pointer.as_ref().map(|_| String::new()),
        })?;
        Ok(())
    }

    fn serialize_value<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Self::Error> {
        let (symbol, pointer) = match &self.key {
            Kind::Key(key) => (
                Symbol::from(key.as_str()),
                child_pointer(self.visitor.pointer.as_deref(), key),
            ),
            Kind::String(symbol) => (*symbol, None),
            _ => (Symbol::Other, None),
        };
        self.field(symbol, self.visitor.model, pointer, value)
    }

    fn end(self) -> Result<Kind, Self::Error> {
        Ok(self.finish())
    }
}

impl SerializeStruct for Record<'_> {
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

impl SerializeStructVariant for Record<'_> {
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

struct Sequence<'a> {
    visitor: Visitor<'a>,
    index: usize,
}

macro_rules! sequence {
    ($trait:ident, $field:ident, $kind:expr) => {
        impl $trait for Sequence<'_> {
            type Ok = Kind;
            type Error = RevisionIdsError;

            fn $field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Self::Error> {
                let pointer = self.visitor.pointer.as_ref().map(|pointer| {
                    let index = self.index;
                    self.index += 1;
                    format!("{pointer}/{index}")
                });
                value.serialize(Visitor {
                    session: self.visitor.session,
                    found: self.visitor.found,
                    model: self.visitor.model,
                    info: false,
                    scan_strings: self.visitor.scan_strings,
                    pointer,
                })?;
                Ok(())
            }

            fn end(self) -> Result<Kind, Self::Error> {
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
                if !self.scan_strings && self.pointer.is_some() {
                    Ok(Kind::Key(value.to_string()))
                } else {
                    Ok(Kind::Number(value as f64))
                }
            }
        )+
    };
}

impl<'a> Serializer for Visitor<'a> {
    type Ok = Kind;
    type Error = RevisionIdsError;
    type SerializeSeq = Sequence<'a>;
    type SerializeTuple = Sequence<'a>;
    type SerializeTupleStruct = Sequence<'a>;
    type SerializeTupleVariant = Sequence<'a>;
    type SerializeMap = Record<'a>;
    type SerializeStruct = Record<'a>;
    type SerializeStructVariant = Record<'a>;

    number!(
        serialize_i8: i8, serialize_i16: i16, serialize_i32: i32, serialize_i64: i64,
        serialize_i128: i128, serialize_u8: u8, serialize_u16: u16, serialize_u32: u32,
        serialize_u64: u64, serialize_u128: u128, serialize_f32: f32, serialize_f64: f64,
    );

    fn serialize_bool(self, value: bool) -> Result<Kind, Self::Error> {
        if !self.scan_strings && self.pointer.is_some() {
            Ok(Kind::Key(value.to_string()))
        } else {
            Ok(Kind::Other)
        }
    }

    fn serialize_char(self, value: char) -> Result<Kind, Self::Error> {
        self.serialize_str(value.encode_utf8(&mut [0; 4]))
    }

    fn serialize_str(self, value: &str) -> Result<Kind, Self::Error> {
        if self.scan_strings && !matches!(self.found.collection, Collection::SitesOnly) {
            self.found.scan_xml(value, self.pointer.as_deref());
        } else if !self.scan_strings && self.pointer.is_some() {
            return Ok(Kind::Key(value.to_owned()));
        }
        Ok(Kind::String(Symbol::from(value)))
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
        mut self,
        _: &'static str,
        _: u32,
        variant: &'static str,
        value: &T,
    ) -> Result<Kind, Self::Error> {
        self.pointer = child_pointer(self.pointer.as_deref(), variant);
        value.serialize(self)?;
        Ok(Kind::Other)
    }

    fn serialize_seq(self, _: Option<usize>) -> Result<Self::SerializeSeq, Self::Error> {
        Ok(Sequence {
            visitor: self,
            index: 0,
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
        mut self,
        _: &'static str,
        _: u32,
        variant: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleVariant, Self::Error> {
        self.pointer = child_pointer(self.pointer.as_deref(), variant);
        self.serialize_seq(None)
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
        mut self,
        _: &'static str,
        _: u32,
        variant: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStructVariant, Self::Error> {
        self.pointer = child_pointer(self.pointer.as_deref(), variant);
        Ok(Record::new(
            Visitor {
                info: false,
                ..self
            },
            "",
        ))
    }
}

enum Site<'a> {
    Typed(&'a mut f64),
    Model(&'a mut Value),
}

struct SiteWalker<'a, 's, 'p> {
    session: &'s BTreeSet<u64>,
    blocks_with_sites: &'s HashSet<*const ()>,
    sites: Vec<(u64, Site<'a>)>,
    paths: Option<&'p mut Vec<(String, u64)>>,
    path: String,
    #[cfg(test)]
    block_make_mut: usize,
}

impl<'a, 's, 'p> SiteWalker<'a, 's, 'p> {
    fn new(
        session: &'s BTreeSet<u64>,
        paths: Option<&'p mut Vec<(String, u64)>>,
        blocks_with_sites: &'s HashSet<*const ()>,
    ) -> Self {
        Self {
            session,
            blocks_with_sites,
            sites: Vec::new(),
            paths,
            path: String::new(),
            #[cfg(test)]
            block_make_mut: 0,
        }
    }

    fn at(&mut self, key: &str, visit: impl FnOnce(&mut Self)) {
        let length = self.path.len();
        if self.paths.is_some() {
            self.path.push('/');
            if key.contains(['~', '/']) {
                self.path.push_str(&key.replace('~', "~0").replace('/', "~1"));
            } else {
                self.path.push_str(key);
            }
        }
        visit(self);
        self.path.truncate(length);
    }

    fn entries<T>(
        &mut self,
        values: impl IntoIterator<Item = T>,
        mut visit: impl FnMut(&mut Self, T),
    ) {
        for (index, value) in values.into_iter().enumerate() {
            let length = self.path.len();
            if self.paths.is_some() {
                write!(self.path, "/{index}").unwrap();
            }
            visit(self, value);
            self.path.truncate(length);
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
            walker.at("content", |walker| walker.blocks(&mut request.document.content));
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
                    walker.at("blockContent", |walker| walker.blocks(&mut comment.block_content));
                },
            );
        });
        for (key, entries) in [
            ("headerEntries", &mut request.header_entries),
            ("footerEntries", &mut request.footer_entries),
        ] {
            self.field_entries(key, entries, |walker, (_, header_footer)| {
                walker.at("1", |walker| {
                    walker.at("content", |walker| walker.blocks(&mut header_footer.content));
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
                    walker.field_entries("tablePatches", &mut part.table_patches, |walker, table| {
                        walker.at("before", |walker| {
                            walker.model(&mut table.before, ModelKind::Any, false);
                        });
                        walker.at("after", |walker| {
                            walker.model(&mut table.after, ModelKind::Any, false);
                        });
                    });
                });
            });
        }
    }

    fn push(&mut self, placeholder: u64, site: Site<'a>) {
        if let Some(paths) = &mut self.paths {
            paths.push((format!("{}/id", self.path), placeholder));
        }
        self.sites.push((placeholder, site));
    }

    fn id(&mut self, id: &'a mut f64) {
        if let Some(placeholder) = self.session_placeholder(*id) {
            self.push(placeholder, Site::Typed(id));
        }
    }

    fn info(&mut self, id: &'a mut f64) {
        self.at("info", |walker| walker.id(id));
    }

    fn session_placeholder(&self, id: f64) -> Option<u64> {
        placeholder(id).filter(|id| self.session.contains(id))
    }

    fn block_mut<T: Clone>(
        &mut self,
        block: &'a mut Arc<T>,
        node_type: impl FnOnce(&T) -> &String,
    ) -> Option<&'a mut T> {
        #[cfg(test)]
        BLOCK_VISITS.set(BLOCK_VISITS.get() + 1);
        let key = ptr::from_ref(node_type(block.as_ref())).cast::<()>();
        if !self.blocks_with_sites.contains(&key) {
            return None;
        }
        #[cfg(test)]
        {
            self.block_make_mut += 1;
        }
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
                    walker.at("structuralChange", |walker| walker.info(&mut change.info.id));
                }
                walker.field_entries(
                    "propertyChanges",
                    cell.property_changes.iter_mut().flatten(),
                    |walker, change| walker.info(&mut change.info.id),
                );
            });
            if let Some(change) = &mut row.structural_change {
                walker.at("structuralChange", |walker| walker.info(&mut change.info.id));
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
        self.field_entries("content", &mut paragraph.content, |walker, child| {
            match child {
                ParagraphContent::Inline(node) => walker.inline(node),
                ParagraphContent::Tracked(change) => walker.tracked(change),
                ParagraphContent::RangeStart(_)
                | ParagraphContent::RangeEnd(_)
                | ParagraphContent::CommentRange(_) => {}
            }
        });
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
        self.field_entries(children_key, tree.children.iter_mut().flatten(), Self::field_tree);
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
        self.field_entries("propertyChanges", changes.iter_mut().flatten(), |walker, change| {
            walker.info(&mut change.info.id);
        });
    }

    fn run_content(&mut self, content: &'a mut [RunContent]) {
        self.field_entries("content", content, |walker, child| {
            match child {
                RunContent::Shape { shape, .. } => walker.at("shape", |walker| walker.shape(shape)),
                RunContent::OpaqueDrawing {
                    text_box: Some(shape),
                    ..
                } => walker.at("textBox", |walker| walker.shape(shape)),
                _ => {}
            }
        });
    }

    fn shape(&mut self, shape: &'a mut Shape) {
        if let Some(model) = &mut shape.drawing_model {
            self.at("drawingModel", |walker| walker.model(model, ModelKind::Any, false));
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
            self.at("scene", |walker| walker.at("root", |walker| walker.scene(root)));
        }
    }

    fn scene(&mut self, node: &'a mut DrawingSceneNode) {
        let [shape_key, children_key] = SCENE_NODE_FIELDS;
        if let Some(shape) = &mut node.shape {
            self.at(shape_key, |walker| walker.shape(shape));
        }
        self.field_entries(children_key, node.children.iter_mut().flatten(), Self::scene);
    }

    fn model(&mut self, model: &'a mut Value, kind: ModelKind, info: bool) {
        match model {
            Value::Object(record) => {
                if info
                    && record.get("author").is_some_and(Value::is_string)
                    && let Some(number) = record.get("id").and_then(Value::as_f64)
                    && let Some(placeholder) = self.session_placeholder(number)
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
                    let mut fields: Vec<_> = record.iter_mut().collect();
                    fields.sort_by_key(|(key, _)| {
                        order
                            .iter()
                            .position(|name| *name == key.as_str())
                            .unwrap_or(order.len())
                    });
                    for (key, value) in fields {
                        self.model_field(kind, key, value);
                    }
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
                let drawing = record.get("content").and_then(Value::as_array).and_then(|content| {
                    match content.as_slice() {
                        [child] => child.get("type").and_then(Value::as_str),
                        _ => None,
                    }
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
