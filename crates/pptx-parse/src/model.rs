use std::collections::BTreeMap;

pub use ooxml_drawingml::ShapeStyle;
use ooxml_drawingml::{
    ColorValue, GeometryPathCommand, ShapeEffects, ShapeFill, ShapeOutline, TableStyleList, Theme,
    ThemeFormatScheme,
};
use serde::{Deserialize, Serialize};

use crate::comments::{Comment, CommentAuthor, CommentFlavor};
use crate::relationships::Relationship;

pub use ooxml_drawingml::chart::{
    ChartAxes, ChartAxis, ChartDataLabels, ChartLegend, ChartMarker, ChartPlotGroup, ChartPoint,
    ChartPointLabel, ChartSeries, ChartSpace, ChartTextProperties,
};

/// Shape elements counted by source ordinals.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) enum ShapeElements {
    WithConnectors,
    #[default]
    WithoutConnectors,
}

impl ShapeElements {
    fn is_legacy(&self) -> bool {
        *self == Self::WithoutConnectors
    }

    pub(crate) fn contains(self, local: &str) -> bool {
        matches!(local, "sp" | "pic" | "graphicFrame" | "grpSp")
            || (self == Self::WithConnectors && local == "cxnSp")
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PptxPackage {
    pub presentation: Presentation,
    pub slides: Vec<Slide>,
    pub layouts: Vec<SlideLayout>,
    pub masters: Vec<SlideMaster>,
    pub themes: Vec<ThemePart>,
    /// Absent from packages serialized before charts were parsed.
    #[serde(default)]
    pub charts: Vec<ChartPart>,
    pub media: Vec<MediaPart>,
    /// Absent from packages serialized before table styles were parsed.
    #[serde(default, skip_serializing_if = "TableStyleList::is_empty")]
    pub table_styles: TableStyleList,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub comment_authors: Vec<CommentAuthor>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub comments: Vec<Comment>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub comment_flavor: Option<CommentFlavor>,
    pub relationships: BTreeMap<String, Vec<Relationship>>,
    #[serde(skip)]
    pub(crate) parts: Vec<PackagePart>,
    #[serde(default, skip_serializing_if = "ShapeElements::is_legacy")]
    pub(crate) shape_elements: ShapeElements,
}

impl PptxPackage {
    pub fn part_bytes(&self, path: &str) -> Option<&[u8]> {
        self.parts
            .iter()
            .find(|part| part.path == path)
            .map(|part| part.bytes.as_slice())
    }

    /// False for packages recovered from a collaboration update, which carry
    /// the parsed model but not the raw part bytes.
    pub fn has_parts(&self) -> bool {
        !self.parts.is_empty()
    }

    /// Whether source ordinals include connectors.
    pub fn models_connectors(&self) -> bool {
        self.shape_elements == ShapeElements::WithConnectors
    }

    pub fn replace_part(&mut self, path: &str, bytes: Vec<u8>) -> bool {
        let Some(part) = self.parts.iter_mut().find(|part| part.path == path) else {
            return false;
        };
        part.bytes = bytes;
        true
    }
}

#[derive(Clone, Debug, PartialEq)]
pub(crate) struct PackagePart {
    pub path: String,
    pub bytes: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Presentation {
    pub part_path: String,
    pub width_emu: i64,
    pub height_emu: i64,
    /// `p:presentation/@firstSlideNum`.
    #[serde(
        default = "default_first_slide_num",
        skip_serializing_if = "is_default_first_slide_num"
    )]
    pub first_slide_num: i32,
    pub slides: Vec<SlideReference>,
    pub master_part_paths: Vec<String>,
}

fn default_first_slide_num() -> i32 {
    1
}

fn is_default_first_slide_num(value: &i32) -> bool {
    *value == default_first_slide_num()
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideReference {
    pub id: u32,
    pub relationship_id: String,
    pub part_path: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Slide {
    pub part_path: String,
    pub name: Option<String>,
    pub layout_part_path: Option<String>,
    pub show_master_shapes: bool,
    pub background: Option<ShapeFill>,
    pub shapes: Vec<ShapeNode>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideLayout {
    pub part_path: String,
    pub name: Option<String>,
    pub layout_type: Option<String>,
    pub master_part_path: Option<String>,
    pub show_master_shapes: bool,
    pub background: Option<ShapeFill>,
    pub shapes: Vec<ShapeNode>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideMaster {
    pub part_path: String,
    pub name: Option<String>,
    pub theme_part_path: Option<String>,
    pub layout_part_paths: Vec<String>,
    pub background: Option<ShapeFill>,
    pub shapes: Vec<ShapeNode>,
    pub text_styles: TextStyleSet,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ThemePart {
    pub part_path: String,
    pub theme: Theme,
    /// Absent from packages serialized before `a:fmtScheme` was parsed.
    #[serde(default, skip_serializing_if = "ThemeFormatScheme::is_empty")]
    pub format_scheme: ThemeFormatScheme,
}

/// A chart part resolved against one referenced presentation theme.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ChartPart {
    pub part_path: String,
    pub theme_part_path: Option<String>,
    pub chart: ChartSpace,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MediaPart {
    pub part_path: String,
    pub content_type: String,
    pub bytes: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum ShapeNode {
    Shape(Shape),
    Picture(Picture),
    GraphicFrame(GraphicFrame),
    Group(GroupShape),
}

impl ShapeNode {
    pub fn id(&self) -> u32 {
        match self {
            Self::Shape(shape) => shape.base.id,
            Self::Picture(picture) => picture.base.id,
            Self::GraphicFrame(frame) => frame.base.id,
            Self::Group(group) => group.base.id,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeBase {
    pub id: u32,
    pub name: String,
    pub description: Option<String>,
    pub hidden: bool,
    pub placeholder: Option<Placeholder>,
    pub transform: ShapeTransform,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeTransform {
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
    pub rotation_deg: f64,
    pub flip_h: bool,
    pub flip_v: bool,
    pub child_x: Option<i64>,
    pub child_y: Option<i64>,
    pub child_width: Option<i64>,
    pub child_height: Option<i64>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Placeholder {
    pub placeholder_type: Option<String>,
    pub index: Option<u32>,
    pub orientation: Option<String>,
    pub size: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Shape {
    #[serde(flatten)]
    pub base: ShapeBase,
    pub geometry: String,
    /// Custom paths in shape-relative coordinates.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub paths: Vec<CustomGeometryPath>,
    /// Theme formatting and text defaults.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<Box<ShapeStyle>>,
    #[serde(default)]
    pub adjust_values: BTreeMap<String, f64>,
    pub fill: Option<ShapeFill>,
    /// The image behind an `a:blipFill`, when the fill is a stretched picture.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture_fill: Option<Box<PictureFill>>,
    pub outline: Option<ShapeOutline>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub effects: Option<ShapeEffects>,
    pub text: Option<TextBody>,
}

/// An `a:blipFill` on a shape: the image, and the box it stretches into.
#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PictureFill {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub relationship_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub media_part_path: Option<String>,
    #[serde(default, skip_serializing_if = "PictureCrop::is_whole")]
    pub crop: PictureCrop,
    /// `a:stretch/a:fillRect` insets, in thousandths of a percent of the box.
    #[serde(default, skip_serializing_if = "PictureCrop::is_whole")]
    pub fill_rect: PictureCrop,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CustomGeometryPath {
    pub commands: Vec<GeometryPathCommand>,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub no_fill: bool,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub no_stroke: bool,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Picture {
    #[serde(flatten)]
    pub base: ShapeBase,
    pub relationship_id: Option<String>,
    pub media_part_path: Option<String>,
    pub crop: PictureCrop,
    /// Bitmap effects in document order.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub effects: Vec<BlipEffect>,
    /// Preset mask; defaults to the frame rectangle.
    #[serde(default = "rect_geometry", skip_serializing_if = "is_rect")]
    pub geometry: String,
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub adjust_values: BTreeMap<String, f64>,
    pub fill: Option<ShapeFill>,
    pub outline: Option<ShapeOutline>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape_effects: Option<ShapeEffects>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<Box<ShapeStyle>>,
}

fn rect_geometry() -> String {
    "rect".to_owned()
}

fn is_rect(geometry: &str) -> bool {
    geometry == "rect"
}

/// Bitmap effects with unresolved colours.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum BlipEffect {
    /// `a:biLevel`: luminance below `threshold` becomes black, the rest white.
    BiLevel { threshold: f64 },
    /// `a:grayscl`.
    Grayscale,
    /// `a:lum`: brightness and contrast, each a fraction in `-1.0..=1.0`.
    Luminance { brightness: f64, contrast: f64 },
    /// `a:duotone`: luminance interpolates between the two colours.
    Duotone {
        shadow: Option<ColorValue>,
        highlight: Option<ColorValue>,
    },
    /// Exact colour replacement.
    ColorChange {
        from: Option<ColorValue>,
        to: Option<ColorValue>,
        #[serde(
            default = "default_use_alpha",
            skip_serializing_if = "use_alpha_is_default",
            rename = "useAlpha"
        )]
        use_alpha: bool,
    },
}

fn default_use_alpha() -> bool {
    true
}

fn use_alpha_is_default(value: &bool) -> bool {
    *value
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct PictureCrop {
    pub left: i32,
    pub top: i32,
    pub right: i32,
    pub bottom: i32,
}

impl PictureCrop {
    pub fn is_whole(&self) -> bool {
        *self == Self::default()
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GraphicFrame {
    #[serde(flatten)]
    pub base: ShapeBase,
    pub data: GraphicFrameData,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "type",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum GraphicFrameData {
    Table(Table),
    Chart {
        relationship_id: String,
        part_path: Option<String>,
    },
    Diagram {
        relationship_ids: Vec<String>,
    },
    Unknown {
        uri: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        picture: Option<Box<Picture>>,
    },
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct GroupShape {
    #[serde(flatten)]
    pub base: ShapeBase,
    pub children: Vec<ShapeNode>,
}

/// An `a:tbl`: its column grid, table-wide properties and rows.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Table {
    /// `a:gridCol/@w`, in EMU.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub grid: Vec<i64>,
    #[serde(default, skip_serializing_if = "TableProperties::is_default")]
    pub properties: TableProperties,
    #[serde(
        default,
        serialize_with = "serialize_table_rows",
        deserialize_with = "deserialize_table_rows"
    )]
    pub rows: Vec<TableRow>,
}

/// `a:tblPr`: which style parts apply, and the style they come from.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", default)]
pub struct TableProperties {
    pub first_row: bool,
    pub last_row: bool,
    pub first_col: bool,
    pub last_col: bool,
    pub band_row: bool,
    pub band_col: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,
}

impl TableProperties {
    pub fn is_default(&self) -> bool {
        *self == Self::default()
    }
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableRow {
    /// `a:tr/@h`, in EMU, a minimum rather than a fixed height.
    #[serde(default)]
    pub height: i64,
    pub cells: Vec<TableCell>,
}

impl TableRow {
    fn is_text_only(&self) -> bool {
        self.height == 0 && self.cells.iter().all(TableCell::is_text_only)
    }
}

/// An `a:tc`, with its `a:tcPr` anchoring and margins folded into `text`.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableCell {
    pub text: TextBody,
    #[serde(default = "unit_span", skip_serializing_if = "is_unit_span")]
    pub grid_span: u32,
    #[serde(default = "unit_span", skip_serializing_if = "is_unit_span")]
    pub row_span: u32,
    /// An `hMerge`/`vMerge` continuation: covered by an earlier origin cell.
    #[serde(default, skip_serializing_if = "is_false")]
    pub merged: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<ShapeFill>,
    #[serde(default, skip_serializing_if = "TableCellBorders::is_empty")]
    pub borders: TableCellBorders,
}

impl TableCell {
    pub fn from_text(text: TextBody) -> Self {
        Self {
            text,
            ..Self::default()
        }
    }

    /// Compared against a cell built from this text alone, so a field added
    /// later cannot be silently dropped by the released encoding.
    fn is_text_only(&self) -> bool {
        *self == Self::from_text(self.text.clone())
    }
}

impl Default for TableCell {
    fn default() -> Self {
        Self {
            text: TextBody::default(),
            grid_span: 1,
            row_span: 1,
            merged: false,
            fill: None,
            borders: TableCellBorders::default(),
        }
    }
}

/// `a:tcPr/a:lnL`, `a:lnT`, `a:lnR` and `a:lnB`.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", default)]
pub struct TableCellBorders {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<ShapeOutline>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<ShapeOutline>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<ShapeOutline>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<ShapeOutline>,
}

impl TableCellBorders {
    pub fn is_empty(&self) -> bool {
        *self == Self::default()
    }
}

fn unit_span() -> u32 {
    1
}

fn is_unit_span(span: &u32) -> bool {
    *span == 1
}

fn is_false(value: &bool) -> bool {
    !*value
}

/// Writes rows that hold nothing but cell text in the released encoding, so a
/// stored package keeps its bytes until a table gains geometry.
fn serialize_table_rows<S>(rows: &[TableRow], serializer: S) -> Result<S::Ok, S::Error>
where
    S: serde::Serializer,
{
    use serde::ser::SerializeSeq;

    if !rows.iter().all(TableRow::is_text_only) {
        return rows.serialize(serializer);
    }
    let mut sequence = serializer.serialize_seq(Some(rows.len()))?;
    for row in rows {
        let cells: Vec<_> = row.cells.iter().map(|cell| &cell.text).collect();
        sequence.serialize_element(&cells)?;
    }
    sequence.end()
}

/// Also accepts the released encoding's rows, which were bare `a:txBody` lists.
fn deserialize_table_rows<'de, D>(deserializer: D) -> Result<Vec<TableRow>, D::Error>
where
    D: serde::Deserializer<'de>,
{
    #[derive(Deserialize)]
    #[serde(untagged)]
    enum Row {
        Modern(TableRow),
        Legacy(Vec<TextBody>),
    }

    Ok(Vec::<Row>::deserialize(deserializer)?
        .into_iter()
        .map(|row| match row {
            Row::Modern(row) => row,
            Row::Legacy(cells) => TableRow {
                height: 0,
                cells: cells.into_iter().map(TableCell::from_text).collect(),
            },
        })
        .collect())
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextBody {
    pub anchor: Option<String>,
    pub vertical: Option<String>,
    /// Use a 1.2 em percentage pitch.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub compat_line_spacing: Option<bool>,
    pub autofit: Option<TextAutofit>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_overflow: Option<TextOverflow>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub horizontal_overflow: Option<TextOverflow>,
    pub inset_left: Option<i64>,
    pub inset_top: Option<i64>,
    pub inset_right: Option<i64>,
    pub inset_bottom: Option<i64>,
    /// List properties by outline level.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub list_style: Vec<ParagraphProperties>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_list_style: Option<Box<ParagraphProperties>>,
    pub paragraphs: Vec<TextParagraph>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TextOverflow {
    Overflow,
    Clip,
    Ellipsis,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "type",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum TextAutofit {
    None,
    Shape,
    Normal {
        font_scale: Option<f64>,
        line_space_reduction: Option<f64>,
    },
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextParagraph {
    pub properties: ParagraphProperties,
    pub runs: Vec<TextRun>,
    pub end_properties: Option<RunProperties>,
}

/// A spacing height, as a share of the text size or in points.
#[derive(Clone, Copy, Debug, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "type",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum LineSpacing {
    Percent { value: f64 },
    Points { value: f64 },
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphProperties {
    pub alignment: Option<String>,
    pub level: u32,
    pub margin_left: Option<i64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub margin_right: Option<i64>,
    pub indent: Option<i64>,
    pub bullet: Option<Bullet>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub line_spacing: Option<LineSpacing>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub space_before: Option<LineSpacing>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub space_after: Option<LineSpacing>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bullet_font: Option<BulletFont>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bullet_color: Option<BulletColor>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bullet_size: Option<BulletSize>,
    pub default_run: Option<RunProperties>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", content = "value", rename_all = "camelCase")]
pub enum BulletFont {
    FollowText,
    Typeface(String),
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", content = "value", rename_all = "camelCase")]
pub enum BulletColor {
    FollowText,
    Color(ColorValue),
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", content = "value", rename_all = "camelCase")]
pub enum BulletSize {
    FollowText,
    Percent(f64),
    Points(f64),
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(
    tag = "type",
    rename_all = "camelCase",
    rename_all_fields = "camelCase"
)]
pub enum Bullet {
    Character {
        value: String,
    },
    AutoNumber {
        scheme: String,
        start_at: u32,
        #[serde(default, skip_serializing_if = "std::ops::Not::not")]
        restart: bool,
    },
    None,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextRun {
    pub text: String,
    pub properties: RunProperties,
    pub field_id: Option<String>,
    pub field_type: Option<String>,
    pub line_break: bool,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RunProperties {
    pub font_size_pt: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub spacing_pt: Option<f64>,
    /// Baseline shift as a percentage of the font size; negative is subscript.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub baseline_pct: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<String>,
    pub font_family: Option<String>,
    pub color: Option<ColorValue>,
    pub language: Option<String>,
    pub hyperlink_relationship_id: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextStyleSet {
    pub title: Vec<ParagraphProperties>,
    pub body: Vec<ParagraphProperties>,
    pub other: Vec<ParagraphProperties>,
}
