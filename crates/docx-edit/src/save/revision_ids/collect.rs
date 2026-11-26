use docx_parse::inline::PropertyChangeInfo;
use docx_parse::paragraph::{CommentRange, RangeEnd, RangeStart, TrackedChangeInfo};

use super::{
    BlockContent, Collected, MAX_ID, Map, ModelKind, ParagraphContent, RunContent, SessionIds,
    Value, Write, block_key, is_model_info, model_info_id, push_pointer,
};

#[cfg(test)]
use super::count_work;

pub(super) fn collect<const PATHS: bool>(value: &impl Visit, session: &SessionIds) -> Collected {
    let mut collector = Collector::<PATHS> {
        session,
        found: Collected::default(),
        path: String::new(),
    };
    value.visit(&mut collector);
    collector.found
}

pub(super) trait Visit {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>);

    fn visit_field<const PATHS: bool>(&self, key: &str, collector: &mut Collector<PATHS>) {
        collector.at(key, |collector| self.visit(collector));
    }
}

pub(super) struct Collector<'a, const PATHS: bool> {
    session: &'a SessionIds,
    found: Collected,
    path: String,
}

impl<const PATHS: bool> Collector<'_, PATHS> {
    fn at(&mut self, key: &str, visit: impl FnOnce(&mut Self)) {
        #[cfg(test)]
        count_work(|work| work.fields += 1);
        if PATHS {
            let length = self.path.len();
            push_pointer(&mut self.path, key);
            visit(self);
            self.path.truncate(length);
        } else {
            visit(self);
        }
    }

    fn index(&mut self, index: usize, visit: impl FnOnce(&mut Self)) {
        #[cfg(test)]
        count_work(|work| work.fields += 1);
        if PATHS {
            let length = self.path.len();
            write!(self.path, "/{index}").unwrap();
            visit(self);
            self.path.truncate(length);
        } else {
            visit(self);
        }
    }

    fn field(&mut self, key: &str, value: &impl Visit) {
        value.visit_field(key, self);
    }

    fn reserve(&mut self, id: u32, raises_base: bool) {
        if id >= MAX_ID {
            return;
        }
        #[cfg(test)]
        count_work(|work| work.reservation_inserts += 1);
        self.found.used.insert(id);
        if PATHS {
            *self
                .found
                .reservations
                .entry((self.path.clone(), id))
                .or_default() += 1;
        }
        if raises_base {
            self.found.max = self.found.max.max(id);
        }
    }

    fn number(&mut self, key: &str, id: f64, raises_base: bool) {
        let word_id = id as u32;
        if id == f64::from(word_id) {
            self.at(key, |collector| collector.reserve(word_id, raises_base));
        }
    }

    fn annotation(&mut self, node_type: &str, id: f64) {
        let raises_base = match node_type {
            "bookmarkStart" | "bookmarkEnd" | "moveFromRangeStart" | "moveFromRangeEnd"
            | "moveToRangeStart" | "moveToRangeEnd" => true,
            "commentRangeStart" | "commentRangeEnd" | "commentReference" => false,
            _ => return,
        };
        self.number("id", id, raises_base);
    }

    fn site(&mut self, id: u64, hidden: Option<u64>) {
        #[cfg(test)]
        count_work(|work| work.site_inserts += 1);
        self.found.site_count += 1;
        let count = self.found.sites.entry(id).or_default();
        count.remaining += 1;
        if let Some(hidden) = hidden {
            count.hidden = Some(count.hidden.map_or(hidden, |prior| prior.min(hidden)));
        }
    }

    fn info(&mut self, id: f64, author: &str, date: Option<&str>, rsid: Option<&str>) {
        if let Some(id) = self.session.placeholder(id) {
            self.site(id, None);
            return;
        }
        self.number("id", id, true);
        self.at("author", |collector| collector.scan_xml(author));
        if let Some(date) = date {
            self.at("date", |collector| collector.scan_xml(date));
        }
        if let Some(rsid) = rsid {
            self.at("rsid", |collector| collector.scan_xml(rsid));
        }
    }

    fn block(&mut self, block: &BlockContent) {
        #[cfg(test)]
        if !matches!(block, BlockContent::RawXml(_)) {
            count_work(|work| work.collected_blocks += 1);
        }
        match block {
            BlockContent::Paragraph(paragraph) => paragraph.visit(self),
            BlockContent::Table(table) => table.visit(self),
            BlockContent::BlockSdt(control) => control.visit(self),
            BlockContent::RawXml(raw) => raw.visit(self),
        }
    }

    fn body_blocks(&mut self, blocks: &[BlockContent]) {
        for (index, block) in blocks.iter().enumerate() {
            let before = self.found.site_count;
            self.index(index, |collector| collector.block(block));
            if !PATHS && self.found.site_count != before {
                #[cfg(test)]
                count_work(|work| work.block_inserts += 1);
                self.found.body_blocks_with_sites.push(index);
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
                self.reserve(id, true);
            }
        }
    }

    fn json_record<'v>(&mut self, get: impl Fn(&str) -> Option<&'v Value>) {
        let id = get("id").and_then(Value::as_f64);
        if get("author").is_some_and(Value::is_string) {
            let raises_base = !get("content").is_some_and(Value::is_array);
            if let Some(id) = id {
                self.number("id", id, raises_base);
            }
            if let Some(id) = get("revisionId").and_then(Value::as_f64) {
                self.number("revisionId", id, raises_base);
            }
        } else if let (Some(id), Some(node_type)) = (id, get("type").and_then(Value::as_str)) {
            self.annotation(node_type, id);
        }
    }

    fn json(&mut self, value: &Value, model: Option<ModelKind>, info: bool) {
        match value {
            Value::Object(record) => {
                if info
                    && let Some(id) =
                        model_info_id(record).and_then(|id| self.session.placeholder(id))
                {
                    let hidden = hidden_model_children(record, self.session);
                    self.site(id, hidden);
                    return;
                }
                self.json_record(|key| record.get(key));
                let kind = model.map(|kind| kind.for_record(record));
                let fields = kind.map_or(&[][..], |kind| kind.fields(record));
                for &key in fields {
                    if let Some(value) = record.get(key) {
                        self.json_field(kind, key, value);
                    }
                }
                for (key, value) in record {
                    if !fields.contains(&key.as_str()) {
                        self.json_field(kind, key, value);
                    }
                }
            }
            Value::Array(items) => {
                for (index, item) in items.iter().enumerate() {
                    self.index(index, |collector| collector.json(item, model, false));
                }
            }
            Value::String(text) => self.scan_xml(text),
            _ => {}
        }
    }

    fn json_field(&mut self, model: Option<ModelKind>, key: &str, value: &Value) {
        self.at(key, |collector| {
            collector.json(
                value,
                model.map(|kind| kind.child(key)),
                model.is_some() && is_model_info(key),
            );
        });
    }

    fn model(&mut self, value: &Value) {
        self.json(value, Some(ModelKind::Any), false);
    }
}

fn hidden_model_children(record: &Map<String, Value>, session: &SessionIds) -> Option<u64> {
    record
        .iter()
        .filter_map(|(key, value)| hidden_model_site(value, is_model_info(key), session))
        .min()
}

fn hidden_model_site(value: &Value, info: bool, session: &SessionIds) -> Option<u64> {
    match value {
        Value::Object(record) => {
            if info
                && let Some(id) = model_info_id(record).and_then(|id| session.placeholder(id))
            {
                return Some(id);
            }
            hidden_model_children(record, session)
        }
        Value::Array(items) => items
            .iter()
            .filter_map(|item| hidden_model_site(item, false, session))
            .min(),
        _ => None,
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

impl Visit for String {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.scan_xml(self);
    }
}

impl<T: Visit> Visit for Option<T> {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        if let Some(value) = self {
            value.visit(collector);
        }
    }

    fn visit_field<const PATHS: bool>(&self, key: &str, collector: &mut Collector<PATHS>) {
        if let Some(value) = self {
            value.visit_field(key, collector);
        }
    }
}

impl<T: Visit> Visit for Vec<T> {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        for (index, value) in self.iter().enumerate() {
            collector.index(index, |collector| value.visit(collector));
        }
    }
}

impl<T: Visit> Visit for Box<T> {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        self.as_ref().visit(collector);
    }
}

impl<T: Visit> Visit for (String, T) {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.index(0, |collector| self.0.visit(collector));
        collector.index(1, |collector| self.1.visit(collector));
    }
}

impl Visit for (u32, String) {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.index(1, |collector| self.1.visit(collector));
    }
}

impl Visit for Value {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.json(self, None, false);
    }
}

impl Visit for docx_parse::paragraph_patch::Attrs {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.json_record(|key| self.get(key));
        for (key, child) in self {
            collector.field(key, child);
        }
    }
}

impl Visit for BlockContent {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        let before = collector.found.site_count;
        collector.block(self);
        if !PATHS && collector.found.site_count != before {
            #[cfg(test)]
            count_work(|work| work.block_inserts += 1);
            collector.found.blocks_with_sites.insert(block_key(self).unwrap());
        }
    }
}

impl Visit for ParagraphContent {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        match self {
            Self::Inline(node) => node.visit(collector),
            Self::Tracked(change) => change.visit(collector),
            Self::RangeStart(range) => range.visit(collector),
            Self::RangeEnd(range) => range.visit(collector),
            Self::CommentRange(range) => range.visit(collector),
        }
    }
}

impl Visit for docx_parse::inline::InlineNode {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        match self {
            Self::Run(run) => run.visit(collector),
            Self::Hyperlink(link) => link.visit(collector),
            Self::BookmarkStart(bookmark) => bookmark.visit(collector),
            Self::BookmarkEnd(bookmark) => bookmark.visit(collector),
            Self::SimpleField(field) => field.visit(collector),
            Self::ComplexField(field) => field.visit(collector),
            Self::InlineSdt(control) => control.visit(collector),
            Self::Math(math) => math.visit(collector),
            Self::RawXml(raw) => raw.visit(collector),
            Self::Tracked(change) => change.visit(collector),
        }
    }
}

impl Visit for RunContent {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        match self {
            Self::Text { text, .. } | Self::InstrText { text } => collector.field("text", text),
            Self::Break { break_type, clear } => {
                collector.field("breakType", break_type);
                collector.field("clear", clear);
            }
            Self::Symbol { font, char } => {
                collector.field("font", font);
                collector.field("char", char);
            }
            Self::CommentReference { id } => {
                if let Some(id) = id {
                    collector.number("id", *id, false);
                }
            }
            Self::FieldChar {
                char_type, form_data, ..
            } => {
                collector.field("charType", char_type);
                collector.field("formData", form_data);
            }
            Self::Drawing {
                image,
                drawing_properties,
            } => {
                collector.field("image", image);
                collector.field("drawing", drawing_properties);
            }
            Self::Shape {
                shape,
                drawing_properties,
            } => {
                collector.field("shape", shape);
                collector.field("drawing", drawing_properties);
            }
            Self::Chart {
                chart,
                drawing_properties,
            } => {
                collector.field("chart", chart);
                collector.field("drawing", drawing_properties);
            }
            Self::HorizontalRule { rule } => collector.field("rule", rule),
            Self::OpaqueDrawing {
                kind,
                xml,
                preview,
                previews,
                text_box,
                drawing_properties,
            } => {
                collector.field("kind", kind);
                collector.field("xml", xml);
                collector.field("preview", preview);
                collector.field("previews", previews);
                collector.field("textBox", text_box);
                collector.field("drawing", drawing_properties);
            }
            Self::Tab
            | Self::FootnoteRef { .. }
            | Self::EndnoteRef { .. }
            | Self::FootnoteRefMark
            | Self::EndnoteRefMark
            | Self::Separator
            | Self::ContinuationSeparator
            | Self::SoftHyphen
            | Self::NoBreakHyphen => {}
        }
    }
}

impl Visit for TrackedChangeInfo {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.info(self.id, &self.author, self.date.as_deref(), None);
    }
}

impl Visit for PropertyChangeInfo {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.info(
            self.id,
            &self.author,
            self.date.as_deref(),
            self.rsid.as_deref(),
        );
    }
}

impl Visit for docx_parse::inline::BookmarkStart {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.number("id", self.id, true);
        collector.field("name", &self.name);
    }
}

impl Visit for docx_parse::inline::BookmarkEnd {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.number("id", self.id, true);
    }
}

impl Visit for RangeStart {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.annotation(&self.node_type, self.id);
        collector.field("type", &self.node_type);
        collector.field("name", &self.name);
    }
}

impl Visit for RangeEnd {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.annotation(&self.node_type, self.id);
        collector.field("type", &self.node_type);
    }
}

impl Visit for CommentRange {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.annotation(&self.node_type, self.id);
        collector.field("type", &self.node_type);
    }
}

impl Visit for docx_parse::paragraph_patch::UnitV1 {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        match self {
            Self::Text { text, attrs } => {
                collector.field("text", text);
                collector.field("attrs", attrs);
            }
            Self::Atom { atom_kind, payload, attrs } => {
                collector.field("atomKind", atom_kind);
                collector.field("payload", payload);
                collector.field("attrs", attrs);
            }
        }
    }
}

impl Visit for docx_parse::paragraph_patch::SlotV1 {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        if let Self::Present { value } = self {
            collector.field("value", value);
        }
    }
}

impl Visit for docx_parse::vml::Watermark {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        match self {
            Self::Text {
                text,
                font,
                color,
                layout,
                ..
            } => {
                collector.field("text", text);
                collector.field("font", font);
                collector.field("color", color);
                collector.field("layout", layout);
            }
            Self::Picture {
                relationship_id,
                media_path,
                content_type,
                data_url,
                ..
            } => {
                collector.field("relId", relationship_id);
                collector.field("mediaPath", media_path);
                collector.field("contentType", content_type);
                collector.field("dataUrl", data_url);
            }
        }
    }
}

macro_rules! wire_name {
    ($field:ident) => { stringify!($field) };
    ($field:ident => $name:literal) => { $name };
}

macro_rules! record {
    ($ty:path { $($field:ident $(=> $name:literal)?),* $(,)? }) => {
        impl Visit for $ty {
            fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
                $(collector.field(wire_name!($field $(=> $name)?), &self.$field);)*
            }
        }
    };
}

impl Visit for docx_parse::document::DocumentBody {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.at("content", |collector| collector.body_blocks(&self.content));
        collector.field("sections", &self.sections);
        collector.field("finalSectionProperties", &self.final_section_properties);
        collector.field("customRootBindings", &self.custom_root_bindings);
        collector.field("comments", &self.comments);
    }
}

impl Visit for docx_parse::notes::Note {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.annotation(&self.story_type, self.id);
        collector.field("type", &self.story_type);
        collector.field("noteType", &self.note_type);
        collector.field("content", &self.content);
        collector.field("customRootBindings", &self.custom_root_bindings);
        collector.field("verbatimXml", &self.verbatim_xml);
    }
}

impl Visit for docx_parse::comments::Comment {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.number("id", self.id, false);
        collector.field("author", &self.author);
        collector.field("initials", &self.initials);
        collector.field("date", &self.date);
        collector.field("content", &self.content);
        collector.field("status", &self.status);
        collector.field("authorId", &self.author_id);
        collector.field("durableId", &self.durable_id);
        collector.field("paraId", &self.para_id);
        collector.field("dateUtc", &self.date_utc);
        collector.field("blockContent", &self.block_content);
    }
}

impl Visit for docx_parse::shape::Shape {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        if let Some(model) = &self.drawing_model {
            collector.at("drawingModel", |collector| collector.model(model));
        }
        collector.field("type", &self.shape_kind);
        collector.field("shapeType", &self.shape_type);
        collector.field("drawingXml", &self.drawing_xml);
        collector.field("drawingRunXml", &self.drawing_run_xml);
        collector.field("id", &self.id);
        collector.field("name", &self.name);
        collector.field("position", &self.position);
        collector.field("wrap", &self.wrap);
        collector.field("fill", &self.fill);
        collector.field("outline", &self.outline);
        collector.field("textBody", &self.text_body);
        collector.field("children", &self.children);
        collector.field("customGeometry", &self.custom_geometry);
        collector.field("scene", &self.scene);
        collector.field("fillPaint", &self.fill_paint);
        collector.field("strokePaint", &self.stroke_paint);
        collector.field("effects", &self.effects);
        collector.field("textBodyProperties", &self.text_body_properties);
        collector.field("title", &self.title);
        collector.field("description", &self.description);
    }
}

impl Visit for docx_parse::shape::ShapeTextBody {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.field("anchor", &self.anchor);
        collector.field("autoFit", &self.auto_fit);
        collector.at("content", |collector| {
            for (index, block) in self.content.iter().enumerate() {
                collector.index(index, |collector| collector.model(block));
            }
        });
    }
}

impl Visit for docx_parse::serializer::S13TablePatch {
    fn visit<const PATHS: bool>(&self, collector: &mut Collector<PATHS>) {
        collector.at("before", |collector| collector.model(&self.before));
        collector.at("after", |collector| collector.model(&self.after));
    }
}

record!(docx_parse::serializer::S13SaveRequest {
    determinism,
    document,
    header_entries => "headerEntries",
    footer_entries => "footerEntries",
    footnotes,
    endnotes,
    footnote_separators => "footnoteSeparators",
    endnote_separators => "endnoteSeparators",
    relationship_entries => "relationshipEntries",
    numbering,
    options,
    selective,
    paragraph_ids => "paragraphIds",
});

record!(docx_parse::document::Section { id, properties, content });

record!(docx_parse::header_footer::HeaderFooter {
    story_type => "type",
    hdr_ftr_type => "hdrFtrType",
    content,
    custom_root_bindings => "customRootBindings",
    watermark,
    watermark_box => "watermarkBox",
});

record!(docx_parse::block::BlockSdt { node_type => "type", properties, content });

record!(docx_parse::paragraph::Paragraph {
    node_type => "type",
    para_id => "paraId",
    para_id_attribute => "paraIdAttribute",
    text_id => "textId",
    extra_attributes => "extraAttributes",
    formatting,
    property_changes => "propertyChanges",
    p_pr_ins => "pPrIns",
    p_pr_del => "pPrDel",
    content,
    list_rendering => "listRendering",
    section_properties => "sectionProperties",
});

record!(docx_parse::paragraph::TrackedInline { node_type => "type", info, content });

record!(docx_parse::paragraph::ParagraphPropertyChange {
    node_type => "type",
    info,
    previous_formatting => "previousFormatting",
    current_formatting => "currentFormatting",
});

record!(docx_parse::table::Table {
    node_type => "type",
    formatting,
    property_changes => "propertyChanges",
    rows,
    trailing_bookmarks => "trailingBookmarks",
});

record!(docx_parse::table::TableRow {
    node_type => "type",
    extra_attributes => "extraAttributes",
    formatting,
    property_changes => "propertyChanges",
    structural_change => "structuralChange",
    cells,
    leading_bookmarks => "leadingBookmarks",
    trailing_bookmarks => "trailingBookmarks",
});

record!(docx_parse::table::TableCell {
    node_type => "type",
    formatting,
    property_changes => "propertyChanges",
    structural_change => "structuralChange",
    content,
    leading_bookmarks => "leadingBookmarks",
});

record!(docx_parse::table::TablePropertyChange {
    node_type => "type",
    info,
    previous_formatting => "previousFormatting",
    current_formatting => "currentFormatting",
});

record!(docx_parse::table::TableRowPropertyChange {
    node_type => "type",
    info,
    previous_formatting => "previousFormatting",
    current_formatting => "currentFormatting",
});

record!(docx_parse::table::TableCellPropertyChange {
    node_type => "type",
    info,
    previous_formatting => "previousFormatting",
    current_formatting => "currentFormatting",
});

record!(docx_parse::table::TableStructuralChangeInfo {
    node_type => "type",
    info,
    v_merge => "vMerge",
    v_merge_orig => "vMergeOrig",
});

record!(docx_parse::inline::Run {
    formatting,
    property_changes => "propertyChanges",
    content,
    direction_scopes => "directionScopes",
});

record!(docx_parse::inline::RunPropertyChange {
    info,
    previous_formatting => "previousFormatting",
    current_formatting => "currentFormatting",
});

record!(docx_parse::inline::RawInlineXml { xml });

record!(docx_parse::inline::MathEquation {
    display,
    omml_xml => "ommlXml",
    plain_text => "plainText",
});

record!(docx_parse::inline::Hyperlink {
    relationship_id => "rId",
    href,
    anchor,
    tooltip,
    target,
    doc_location => "docLocation",
    children,
    structured_children => "structuredChildren",
});

record!(docx_parse::inline::InlineSdt { properties, content });

record!(docx_parse::inline::SimpleField {
    instruction,
    field_type => "fieldType",
    content,
    structured_result => "structuredResult",
    field_tree => "fieldTree",
});

record!(docx_parse::inline::ComplexField {
    instruction,
    field_type => "fieldType",
    field_code => "fieldCode",
    field_result => "fieldResult",
    formatting,
    raw_form_data_xml => "rawFormDataXml",
    source_run_xml => "sourceRunXml",
    source_run_digest => "sourceRunDigest",
    structured_code => "structuredCode",
    structured_result => "structuredResult",
    field_tree => "fieldTree",
    direction_scopes => "directionScopes",
});

record!(docx_parse::inline::StructuredFieldContent { inline, blocks });

record!(docx_parse::inline::StructuredFieldTree {
    code,
    result,
    children,
    display_mode => "displayMode",
});

record!(docx_parse::numbering::AbstractNumbering {
    multi_level_type => "multiLevelType",
    num_style_link => "numStyleLink",
    style_link => "styleLink",
    levels,
    name,
});

record!(docx_parse::paragraph_patch::AttrDeltaV1 { path, before, after });

record!(docx_parse::borders::BorderSpec { style, color });

record!(docx_parse::borders::Borders {
    top,
    bottom,
    left,
    right,
    inside_h => "insideH",
    inside_v => "insideV",
    between,
    bar,
    start,
    end,
    tl2br,
    tr2bl,
});

record!(docx_parse::formatting::CellMargins { top, bottom, left, right, start, end });

record!(docx_parse::chart::Chart {
    content_type => "type",
    chart_type => "chartType",
    relationship_id => "rId",
    path,
    title,
    legend,
    series,
    axes,
    wrap,
    position,
    plot_groups => "plotGroups",
    axis_list => "axisList",
    description,
    drawing_xml => "drawingXml",
});

record!(ooxml_drawingml::chart::ChartAxes { category, value });

record!(ooxml_drawingml::chart::ChartAxis {
    id,
    title,
    labels,
    axis_type => "axisType",
    position,
    cross_axis_id => "crossAxisId",
    crosses,
    cross_between => "crossBetween",
    number_format => "numberFormat",
    major_tick_mark => "majorTickMark",
    minor_tick_mark => "minorTickMark",
    tick_label_position => "tickLabelPosition",
    major_gridline_line => "majorGridlineLine",
    minor_gridline_line => "minorGridlineLine",
    text,
    line,
});

record!(ooxml_drawingml::chart::ChartDataLabels {
    separator,
    position,
    number_format => "numberFormat",
    text,
    points,
});

record!(ooxml_drawingml::chart::ChartLegend { position, text });

record!(ooxml_drawingml::chart::ChartLine { color });

record!(ooxml_drawingml::chart::ChartMarker { symbol, color });

record!(ooxml_drawingml::chart::ChartPlotGroup {
    chart_type => "chartType",
    grouping,
    axis_ids => "axisIds",
    series,
    scatter_style => "scatterStyle",
    radar_style => "radarStyle",
    size_represents => "sizeRepresents",
    data_labels => "dataLabels",
});

record!(ooxml_drawingml::chart::ChartPoint { color });

record!(ooxml_drawingml::chart::ChartPointLabel { text, labels });

record!(ooxml_drawingml::chart::ChartSeries {
    name,
    categories,
    color,
    category_formula => "categoryFormula",
    value_formula => "valueFormula",
    value_format => "valueFormat",
    axis_ids => "axisIds",
    points,
    grouping,
    marker,
    data_labels => "dataLabels",
    line,
});

record!(ooxml_drawingml::chart::ChartTextProperties { font, color });

record!(ooxml_drawingml::ColorValue {
    rgb,
    theme_color => "themeColor",
    theme_tint => "themeTint",
    theme_shade => "themeShade",
});

record!(docx_parse::inline::DirectionScope { val, extra });

record!(docx_parse::section::DocumentGrid { grid_type => "type" });

record!(docx_parse::image::DrawingProperties {
    doc_pr_id => "docPrId",
    name,
    title,
    descr,
});

record!(docx_parse::shape::DrawingScene { root, title, description });

record!(docx_parse::shape::DrawingSceneNode {
    kind,
    id,
    name,
    shape,
    children,
    fill,
    effects,
    relationship_id => "relationshipId",
});

record!(docx_parse::inline::FieldFormData {
    name,
    help_text => "helpText",
    status_text => "statusText",
    entry_macro => "entryMacro",
    exit_macro => "exitMacro",
    control_type => "controlType",
    default_value => "defaultValue",
    text_type => "textType",
    text_format => "textFormat",
    list_entries => "listEntries",
});

record!(docx_parse::formatting::FloatingTableProperties {
    horz_anchor => "horzAnchor",
    vert_anchor => "vertAnchor",
    tblp_x_spec => "tblpXSpec",
    tblp_y_spec => "tblpYSpec",
});

record!(docx_parse::formatting::FontFamily {
    ascii,
    h_ansi => "hAnsi",
    east_asia => "eastAsia",
    cs,
    ascii_theme => "asciiTheme",
    h_ansi_theme => "hAnsiTheme",
    east_asia_theme => "eastAsiaTheme",
    cs_theme => "csTheme",
    hint,
});

record!(ooxml_drawingml::GradientFill { gradient_type => "type", stops });

record!(ooxml_drawingml::GradientStop { color });

record!(docx_parse::vml::HorizontalRule { alignment, color, xml });

record!(docx_parse::image::Image {
    image_type => "type",
    id,
    name,
    relationship_id => "rId",
    src,
    mime_type => "mimeType",
    filename,
    alt,
    title,
    wrap,
    position,
    shape_type => "shapeType",
    hlink_href => "hlinkHref",
    outline,
    effects,
    hyperlink,
});

record!(ooxml_drawingml::ImageEffect { kind, colors, raw_name => "rawName" });

record!(ooxml_drawingml::ImageEffects { ordered });

record!(ooxml_drawingml::ImageHyperlink { href, tooltip, target, doc_location });

record!(docx_parse::image::ImagePosition { horizontal, vertical });

record!(docx_parse::image::ImageWrap { wrap_type => "type", wrap_text => "wrapText" });

record!(docx_parse::paragraph_patch::InheritanceV1 {
    paragraph,
    runs,
    paragraph_mark => "paragraphMark",
});

record!(docx_parse::paragraph_patch::InsertedTextV1 { text, attrs, inherit });

record!(docx_parse::paragraph_patch::LegacyParagraphV1 { detail });

record!(docx_parse::numbering::LevelOverride { lvl });

record!(ooxml_drawingml::LineEnd { end_type => "type", width, length });

record!(docx_parse::section::LineNumbering { restart });

record!(docx_parse::numbering::ListLevel {
    num_fmt => "numFmt",
    lvl_text => "lvlText",
    lvl_jc => "lvlJc",
    suffix,
    p_pr => "pPr",
    r_pr => "rPr",
});

record!(docx_parse::numbering::ListRendering {
    marker,
    num_fmt => "numFmt",
    marker_font_family => "markerFontFamily",
    marker_color => "markerColor",
    marker_suffix => "markerSuffix",
    level_num_fmts => "levelNumFmts",
});

record!(docx_parse::notes::NoteProperties {
    position,
    num_fmt => "numFmt",
    num_restart => "numRestart",
    custom_number_format => "customNumberFormat",
    separators,
});

record!(docx_parse::notes::NoteSeparatorReference { kind });

record!(docx_parse::numbering::NumberingDefinitions {
    abstract_nums => "abstractNums",
    nums,
});

record!(docx_parse::numbering::NumberingInstance {
    level_overrides => "levelOverrides",
});

record!(docx_parse::section::PageBorders {
    top,
    bottom,
    left,
    right,
    display,
    offset_from => "offsetFrom",
    z_order => "zOrder",
});

record!(docx_parse::section::PageNumberingProperties {
    format,
    chapter_separator => "chapterSeparator",
});

record!(docx_parse::shape::PaintStop { color });

record!(docx_parse::formatting::ParagraphCnfStyle {
    val,
    first_row => "firstRow",
    last_row => "lastRow",
    first_column => "firstColumn",
    last_column => "lastColumn",
    odd_v_band => "oddVBand",
    even_v_band => "evenVBand",
    odd_h_band => "oddHBand",
    even_h_band => "evenHBand",
    first_row_first_column => "firstRowFirstColumn",
    first_row_last_column => "firstRowLastColumn",
    last_row_first_column => "lastRowFirstColumn",
    last_row_last_column => "lastRowLastColumn",
});

record!(docx_parse::formatting::ParagraphFormatting {
    alignment,
    line_spacing_rule => "lineSpacingRule",
    borders,
    shading,
    tabs,
    style_id => "styleId",
    frame,
    cnf_style => "cnfStyle",
    run_properties => "runProperties",
});

record!(docx_parse::formatting::ParagraphFrame {
    h_anchor => "hAnchor",
    v_anchor => "vAnchor",
    x_align => "xAlign",
    y_align => "yAlign",
    wrap,
});

record!(docx_parse::paragraph_patch::ParagraphPatchV1 {
    session_key => "sessionKey",
    seed,
    text_edits => "textEdits",
    run_edits => "runEdits",
    para_deltas => "paraDeltas",
    para_inherit => "paraInherit",
    mark_inherit => "markInherit",
    result_sha256 => "resultSha256",
});

record!(docx_parse::shape::PictureTile { alignment, flip });

record!(docx_parse::image::PositionAxis { relative_to => "relativeTo", alignment });

record!(docx_parse::paragraph::RawAttribute { name, value });

record!(docx_parse::relationships::Relationship {
    id,
    relationship_type => "type",
    target,
});

record!(docx_parse::paragraph_patch::RunEditV1 { deltas, target, inherit });

record!(docx_parse::formatting::RunLanguage { latin, east_asia => "eastAsia", bidi });

record!(docx_parse::serializer::S13ParagraphId { part, para_id => "paraId" });

record!(docx_parse::serializer::S13ParagraphIds {
    removed_comment_ids => "removedCommentIds",
    section_edits => "sectionEdits",
    assignments,
    patched_parts => "patchedParts",
    spliced_parts => "splicedParts",
    whole_parts => "wholeParts",
    touched_parts => "touchedParts",
});

record!(docx_parse::serializer::S13PatchedPart { part, para_ids => "paraIds" });

record!(docx_parse::serializer::S13SaveOptions { modified_by => "modifiedBy" });

record!(docx_parse::serializer::S13SelectiveSave {
    changed_para_ids => "changedParaIds",
    source_paragraphs => "sourceParagraphs",
});

record!(docx_parse::serializer::S13SourceParagraphs { part_sha256 => "partSha256" });

record!(docx_parse::serializer::S13SplicedPart {
    part,
    sha256,
    paragraph_patches => "paragraphPatches",
    legacy_paragraphs => "legacyParagraphs",
    table_patches => "tablePatches",
});

record!(docx_parse::serializer::paragraph_ids::S13TouchedPart { part });

record!(docx_parse::serializer::paragraph_ids::S13WholePart { part, sha256 });

record!(docx_parse::inline::SdtCheckboxGlyph { value, font });

record!(docx_parse::inline::SdtControlState {
    value,
    selected_value => "selectedValue",
});

record!(docx_parse::inline::SdtDataBinding {
    xpath,
    store_item_id => "storeItemID",
    prefix_mappings => "prefixMappings",
});

record!(docx_parse::inline::SdtDateState {
    full_date => "fullDate",
    format,
    language,
    calendar,
    storage_format => "storageFormat",
});

record!(docx_parse::inline::SdtGalleryState { gallery, category });

record!(docx_parse::inline::SdtListItem { display_text => "displayText", value });

record!(docx_parse::inline::SdtProperties {
    sdt_type => "sdtType",
    alias,
    tag,
    lock,
    placeholder,
    date_format => "dateFormat",
    list_items => "listItems",
    run_properties => "runProperties",
    date_state => "dateState",
    list_last_value => "listLastValue",
    checked_state => "checkedState",
    unchecked_state => "uncheckedState",
    gallery,
    appearance,
    color,
    control_state => "controlState",
    data_binding => "dataBinding",
    raw_properties_xml => "rawPropertiesXml",
    raw_end_properties_xml => "rawEndPropertiesXml",
});

record!(docx_parse::section::SectionBackground {
    color,
    theme_color => "themeColor",
    theme_tint => "themeTint",
    theme_shade => "themeShade",
});

record!(docx_parse::serializer::SectionEdit { properties });

record!(docx_parse::serializer::SectionEdits { sha256, edits });

record!(docx_parse::section::SectionProperties {
    section_id => "sectionId",
    orientation,
    section_start => "sectionStart",
    vertical_align => "verticalAlign",
    header_references => "headerReferences",
    footer_references => "footerReferences",
    line_numbers => "lineNumbers",
    page_numbering => "pageNumbering",
    page_borders => "pageBorders",
    background,
    footnote_pr => "footnotePr",
    endnote_pr => "endnotePr",
    doc_grid => "docGrid",
});

record!(docx_parse::paragraph_patch::SeedV1 {
    units,
    properties,
    unit_sha256 => "unitSha256",
    inherit,
});

record!(docx_parse::serializer::SerializerDeterminism { seed, now });

record!(docx_parse::scalars::ShadingProperties { color, fill, pattern });

record!(docx_parse::shape::ShapeEffect { kind, color, raw_name => "rawName" });

record!(ooxml_drawingml::ShapeFill { fill_type => "type", color, gradient });

record!(docx_parse::shape::ShapeFillPaint {
    kind,
    color,
    theme_color => "themeColor",
    gradient_type => "gradientType",
    stops,
    path_shape => "pathShape",
    pattern_preset => "patternPreset",
    foreground_color => "foregroundColor",
    background_color => "backgroundColor",
    picture,
    fill_mode => "fillMode",
    tile,
});

record!(ooxml_drawingml::ShapeOutline {
    color,
    gradient,
    style,
    cap,
    join,
    head_end => "headEnd",
    tail_end => "tailEnd",
});

record!(docx_parse::shape::ShapeStrokePaint {
    fill,
    dash,
    compound,
    alignment,
    cap,
    join,
    head_end => "headEnd",
    tail_end => "tailEnd",
});

record!(docx_parse::shape::ShapeTextBodyProperties {
    vertical,
    anchor,
    wrap,
    horizontal_overflow => "horizontalOverflow",
    vertical_overflow => "verticalOverflow",
    auto_fit => "autoFit",
    preset_text_warp => "presetTextWarp",
});

record!(docx_parse::section::StoryReference {
    reference_type => "type",
    relationship_id => "rId",
});

record!(docx_parse::shape::StrokeEnd { end_type => "type", width, length });

record!(docx_parse::tabs::TabStop { alignment, leader });

record!(docx_parse::formatting::TableCellFormatting {
    width,
    borders,
    margins,
    shading,
    vertical_align => "verticalAlign",
    text_direction => "textDirection",
    v_merge => "vMerge",
});

record!(docx_parse::formatting::TableFormatting {
    caption,
    description,
    width,
    justification,
    cell_spacing => "cellSpacing",
    indent,
    borders,
    cell_margins => "cellMargins",
    layout,
    style_id => "styleId",
    look,
    shading,
    overlap,
    floating,
});

record!(docx_parse::formatting::TableLook { value });

record!(docx_parse::formatting::TableMeasurement { kind => "type" });

record!(docx_parse::formatting::TableRowFormatting {
    height,
    height_rule => "heightRule",
    justification,
    width_before => "widthBefore",
    width_after => "widthAfter",
});

record!(docx_parse::paragraph_patch::TextEditV1 { insert });

record!(docx_parse::formatting::TextFill { kind, color, stops });

record!(docx_parse::formatting::TextFormatting {
    underline,
    vert_align => "vertAlign",
    color,
    highlight,
    shading,
    font_family => "fontFamily",
    language,
    ligatures,
    effect,
    emphasis_mark => "emphasisMark",
    modern_effects => "modernEffects",
    style_id => "styleId",
});

record!(docx_parse::formatting::TextGlow { color });

record!(docx_parse::formatting::TextGradientStop { color });

record!(docx_parse::formatting::TextModernEffects {
    glow,
    shadow,
    text_fill => "textFill",
    text_outline => "textOutline",
});

record!(docx_parse::formatting::TextOutline { color, dash });

record!(docx_parse::formatting::TextShadow { color });

record!(docx_parse::scalars::UnderlineValue { style, color });

record!(docx_parse::vml::WatermarkBox {
    horizontal,
    horizontal_relative => "horizontalRelative",
    vertical,
    vertical_relative => "verticalRelative",
});
