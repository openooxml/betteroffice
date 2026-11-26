use std::collections::btree_map::Entry;
use std::collections::{BTreeMap, BTreeSet};
use std::fmt;
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
    if session.is_empty() {
        return Ok(BTreeMap::new());
    }
    let mut found = collect(request, session)?;
    if found.sites.is_empty() {
        return Ok(BTreeMap::new());
    }
    let mut walker = SiteWalker {
        session,
        sites: Vec::new(),
    };
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

    let rewritten = collect(request, session)?;
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
pub fn session_revision_id_reservations(
    request: &S13SaveRequest,
    session: &BTreeSet<u64>,
) -> Result<(BTreeSet<u32>, u32), RevisionIdsError> {
    let found = collect(request, session)?;
    Ok((found.used, found.max))
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
struct Collected {
    max: u32,
    used: BTreeSet<u32>,
    sites: BTreeMap<u64, usize>,
}

impl Collected {
    fn note(&mut self, id: Kind, raises_base: bool) {
        if let Kind::Number(id) = id
            && (0.0..f64::from(MAX_ID)).contains(&id)
            && id.fract() == 0.0
        {
            let id = id as u32;
            self.used.insert(id);
            if raises_base {
                self.max = self.max.max(id);
            }
        }
    }

    fn site(&mut self, id: u64, count: usize) {
        *self.sites.entry(id).or_default() += count;
    }

    fn merge(&mut self, other: Self) {
        self.max = self.max.max(other.max);
        self.used.extend(other.used);
        for (id, count) in other.sites {
            self.site(id, count);
        }
    }

    fn scan_xml(&mut self, text: &str) {
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
                self.note(Kind::Number(f64::from(id)), true);
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

#[derive(Clone, Copy, Default)]
enum Kind {
    #[default]
    Other,
    String(Symbol),
    Number(f64),
    Array,
}

fn collect<T: Serialize>(
    value: &T,
    session: &BTreeSet<u64>,
) -> Result<Collected, RevisionIdsError> {
    let mut found = Collected::default();
    value.serialize(Visitor {
        session,
        found: &mut found,
        model: false,
        info: false,
        scan_strings: true,
    })?;
    Ok(found)
}

struct Visitor<'a> {
    session: &'a BTreeSet<u64>,
    found: &'a mut Collected,
    model: bool,
    info: bool,
    scan_strings: bool,
}

struct Record<'a> {
    visitor: Visitor<'a>,
    name: &'static str,
    key: Symbol,
    pending: Option<Collected>,
    author: Kind,
    id: Kind,
    revision_id: Kind,
    content: Kind,
    node_type: Kind,
}

impl<'a> Record<'a> {
    fn new(visitor: Visitor<'a>, name: &'static str) -> Self {
        let info = visitor.info || matches!(name, "TrackedChangeInfo" | "PropertyChangeInfo");
        Self {
            visitor,
            name,
            key: Symbol::Other,
            pending: info.then(Collected::default),
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
        value: &T,
    ) -> Result<(), RevisionIdsError> {
        let kind = value.serialize(Visitor {
            session: self.visitor.session,
            found: self.pending.as_mut().unwrap_or(&mut *self.visitor.found),
            model,
            info: model && matches!(key, Symbol::Info),
            scan_strings: self.visitor.scan_strings,
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
        let model = self.visitor.model
            || matches!(
                (self.name, key),
                ("ShapeTextBody", "content")
                    | ("Shape", "drawingModel")
                    | ("S13TablePatch", "before" | "after")
            );
        self.field(Symbol::from(key), model, value)
    }

    fn finish(self) -> Kind {
        let found = self.visitor.found;
        if let Some(pending) = self.pending {
            if matches!(self.author, Kind::String(_))
                && let Kind::Number(id) = self.id
                && let Some(id) = placeholder(id)
                && self.visitor.session.contains(&id)
            {
                found.site(id, 1);
                return Kind::Other;
            }
            found.merge(pending);
        }
        if matches!(self.author, Kind::String(_)) {
            let raises_base = !matches!(self.content, Kind::Array);
            found.note(self.id, raises_base);
            found.note(self.revision_id, raises_base);
        } else if let Kind::String(symbol @ (Symbol::CommentMarker | Symbol::Annotation)) =
            self.node_type
        {
            found.note(self.id, matches!(symbol, Symbol::Annotation));
        }
        Kind::Other
    }
}

impl SerializeMap for Record<'_> {
    type Ok = Kind;
    type Error = RevisionIdsError;

    fn serialize_key<T: Serialize + ?Sized>(&mut self, key: &T) -> Result<(), Self::Error> {
        self.key = match key.serialize(Visitor {
            session: self.visitor.session,
            found: self.pending.as_mut().unwrap_or(&mut *self.visitor.found),
            model: false,
            info: false,
            scan_strings: false,
        })? {
            Kind::String(symbol) => symbol,
            _ => Symbol::Other,
        };
        Ok(())
    }

    fn serialize_value<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Self::Error> {
        self.field(self.key, self.visitor.model, value)
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

struct Sequence<'a>(Visitor<'a>);

macro_rules! sequence {
    ($trait:ident, $field:ident, $kind:expr) => {
        impl $trait for Sequence<'_> {
            type Ok = Kind;
            type Error = RevisionIdsError;

            fn $field<T: Serialize + ?Sized>(&mut self, value: &T) -> Result<(), Self::Error> {
                value.serialize(Visitor {
                    session: self.0.session,
                    found: self.0.found,
                    model: self.0.model,
                    info: false,
                    scan_strings: self.0.scan_strings,
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
                Ok(Kind::Number(value as f64))
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

    fn serialize_bool(self, _: bool) -> Result<Kind, Self::Error> {
        Ok(Kind::Other)
    }

    fn serialize_char(self, value: char) -> Result<Kind, Self::Error> {
        self.serialize_str(value.encode_utf8(&mut [0; 4]))
    }

    fn serialize_str(self, value: &str) -> Result<Kind, Self::Error> {
        if self.scan_strings {
            self.found.scan_xml(value);
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
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        value: &T,
    ) -> Result<Kind, Self::Error> {
        value.serialize(self)?;
        Ok(Kind::Other)
    }

    fn serialize_seq(self, _: Option<usize>) -> Result<Self::SerializeSeq, Self::Error> {
        Ok(Sequence(self))
    }

    fn serialize_tuple(self, _: usize) -> Result<Self::SerializeTuple, Self::Error> {
        Ok(Sequence(self))
    }

    fn serialize_tuple_struct(
        self,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleStruct, Self::Error> {
        Ok(Sequence(self))
    }

    fn serialize_tuple_variant(
        self,
        _: &'static str,
        _: u32,
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeTupleVariant, Self::Error> {
        Ok(Sequence(self))
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
        _: &'static str,
        _: usize,
    ) -> Result<Self::SerializeStructVariant, Self::Error> {
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

struct SiteWalker<'a, 's> {
    session: &'s BTreeSet<u64>,
    sites: Vec<(u64, Site<'a>)>,
}

impl<'a> SiteWalker<'a, '_> {
    fn request(&mut self, request: &'a mut S13SaveRequest) {
        self.blocks(&mut request.document.content);
        for section in request.document.sections.iter_mut().flatten() {
            self.blocks(&mut section.content);
        }
        for comment in request.document.comments.iter_mut().flatten() {
            for paragraph in &mut comment.content {
                self.paragraph(paragraph);
            }
            self.blocks(&mut comment.block_content);
        }
        for (_, header_footer) in request
            .header_entries
            .iter_mut()
            .chain(&mut request.footer_entries)
        {
            self.blocks(&mut header_footer.content);
        }
        for note in request
            .footnotes
            .iter_mut()
            .chain(&mut request.endnotes)
            .chain(&mut request.footnote_separators)
            .chain(&mut request.endnote_separators)
        {
            self.blocks(&mut note.content);
        }
        if let Some(plan) = &mut request.paragraph_ids {
            for part in &mut plan.spliced_parts {
                for table in &mut part.table_patches {
                    self.model(&mut table.before, false);
                    self.model(&mut table.after, false);
                }
            }
        }
    }

    fn id(&mut self, id: &'a mut f64) {
        if let Some(placeholder) = self.session_placeholder(*id) {
            self.sites.push((placeholder, Site::Typed(id)));
        }
    }

    fn session_placeholder(&self, id: f64) -> Option<u64> {
        placeholder(id).filter(|id| self.session.contains(id))
    }

    fn blocks(&mut self, blocks: &'a mut [BlockContent]) {
        for block in blocks {
            match block {
                BlockContent::Paragraph(paragraph) => self.paragraph(Arc::make_mut(paragraph)),
                BlockContent::Table(table) => self.table(Arc::make_mut(table)),
                BlockContent::BlockSdt(control) => {
                    self.blocks(&mut Arc::make_mut(control).content);
                }
                BlockContent::RawXml(_) => {}
            }
        }
    }

    fn table(&mut self, table: &'a mut Table) {
        for row in &mut table.rows {
            for cell in &mut row.cells {
                self.blocks(&mut cell.content);
                if let Some(change) = &mut cell.structural_change {
                    self.id(&mut change.info.id);
                }
                for change in cell.property_changes.iter_mut().flatten() {
                    self.id(&mut change.info.id);
                }
            }
            if let Some(change) = &mut row.structural_change {
                self.id(&mut change.info.id);
            }
            for change in row.property_changes.iter_mut().flatten() {
                self.id(&mut change.info.id);
            }
        }
        for change in table.property_changes.iter_mut().flatten() {
            self.id(&mut change.info.id);
        }
    }

    fn paragraph(&mut self, paragraph: &'a mut Paragraph) {
        for child in &mut paragraph.content {
            match child {
                ParagraphContent::Inline(node) => self.inline(node),
                ParagraphContent::Tracked(change) => self.tracked(change),
                ParagraphContent::RangeStart(_)
                | ParagraphContent::RangeEnd(_)
                | ParagraphContent::CommentRange(_) => {}
            }
        }
        for info in paragraph.p_pr_ins.iter_mut().chain(&mut paragraph.p_pr_del) {
            self.id(&mut info.id);
        }
        for change in paragraph.property_changes.iter_mut().flatten() {
            self.id(&mut change.info.id);
        }
    }

    fn tracked(&mut self, change: &'a mut TrackedInline) {
        self.id(&mut change.info.id);
        self.inlines(&mut change.content);
    }

    fn inlines(&mut self, nodes: &'a mut [InlineNode]) {
        for node in nodes {
            self.inline(node);
        }
    }

    fn inline(&mut self, node: &'a mut InlineNode) {
        match node {
            InlineNode::Run(run) => self.run(run),
            InlineNode::Tracked(change) => self.tracked(change),
            InlineNode::Hyperlink(link) => {
                self.inlines(&mut link.children);
                if let Some(children) = &mut link.structured_children {
                    self.inlines(children);
                }
            }
            InlineNode::InlineSdt(control) => self.inlines(&mut control.content),
            InlineNode::SimpleField(field) => {
                for run in &mut field.content {
                    self.run(run);
                }
                self.structured(field.structured_result.as_mut());
                if let Some(tree) = &mut field.field_tree {
                    self.field_tree(tree);
                }
            }
            InlineNode::ComplexField(field) => {
                for run in field.field_code.iter_mut().chain(&mut field.field_result) {
                    self.run(run);
                }
                self.structured(field.structured_code.as_mut());
                self.structured(field.structured_result.as_mut());
                if let Some(tree) = &mut field.field_tree {
                    self.field_tree(tree);
                }
            }
            InlineNode::BookmarkStart(_)
            | InlineNode::BookmarkEnd(_)
            | InlineNode::Math(_)
            | InlineNode::RawXml(_) => {}
        }
    }

    fn structured(&mut self, field: Option<&'a mut StructuredFieldContent>) {
        if let Some(field) = field {
            if let Some(nodes) = &mut field.inline {
                self.inlines(nodes);
            }
            if let Some(blocks) = &mut field.blocks {
                self.blocks(blocks);
            }
        }
    }

    fn field_tree(&mut self, tree: &'a mut StructuredFieldTree) {
        self.structured(tree.code.as_mut());
        self.structured(tree.result.as_mut());
        for child in tree.children.iter_mut().flatten() {
            self.field_tree(child);
        }
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
        for change in changes.iter_mut().flatten() {
            self.id(&mut change.info.id);
        }
    }

    fn run_content(&mut self, content: &'a mut [RunContent]) {
        for child in content {
            match child {
                RunContent::Shape { shape, .. } => self.shape(shape),
                RunContent::OpaqueDrawing {
                    text_box: Some(shape),
                    ..
                } => self.shape(shape),
                _ => {}
            }
        }
    }

    fn shape(&mut self, shape: &'a mut Shape) {
        if let Some(model) = &mut shape.drawing_model {
            self.model(model, false);
        }
        if let Some(body) = &mut shape.text_body {
            for block in &mut body.content {
                self.model(block, false);
            }
        }
        for child in shape.children.iter_mut().flatten() {
            self.shape(child);
        }
        if let Some(root) = shape.scene.as_mut().and_then(|scene| scene.root.as_mut()) {
            self.scene(root);
        }
    }

    fn scene(&mut self, node: &'a mut DrawingSceneNode) {
        if let Some(shape) = &mut node.shape {
            self.shape(shape);
        }
        for child in node.children.iter_mut().flatten() {
            self.scene(child);
        }
    }

    fn model(&mut self, model: &'a mut Value, info: bool) {
        match model {
            Value::Object(record) => {
                if info
                    && record.get("author").is_some_and(Value::is_string)
                    && let Some(number) = record.get("id").and_then(Value::as_f64)
                    && let Some(placeholder) = self.session_placeholder(number)
                {
                    self.sites
                        .push((placeholder, Site::Model(record.get_mut("id").unwrap())));
                    return;
                }
                let order: &[&str] = match record.get("type").and_then(Value::as_str) {
                    Some("paragraph") => &["content", "pPrIns", "pPrDel", "propertyChanges"],
                    Some("table") => &["rows", "propertyChanges"],
                    Some("tableRow") => &["cells", "structuralChange", "propertyChanges"],
                    Some("tableCell") => &["content", "structuralChange", "propertyChanges"],
                    Some("insertion" | "deletion" | "moveFrom" | "moveTo") => &["info", "content"],
                    Some("run") => {
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
                    Some("simpleField") => &["content", "structuredResult", "fieldTree"],
                    Some("complexField") => &[
                        "fieldCode",
                        "fieldResult",
                        "structuredCode",
                        "structuredResult",
                        "fieldTree",
                    ],
                    Some("shape") => &["drawingModel", "textBody", "children", "scene"],
                    _ => &[],
                };
                if order.is_empty() {
                    for (key, value) in record {
                        self.model(value, matches!(Symbol::from(key.as_str()), Symbol::Info));
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
                        self.model(value, matches!(Symbol::from(key.as_str()), Symbol::Info));
                    }
                }
            }
            Value::Array(items) => {
                for value in items {
                    self.model(value, false);
                }
            }
            _ => {}
        }
    }
}

#[cfg(test)]
mod tests;
