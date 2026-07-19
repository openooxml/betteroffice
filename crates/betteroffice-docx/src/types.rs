use docx_layout::display_list::DisplayList;
use docx_layout::types::Layout;
use docx_parse::chart::Chart;
use docx_parse::document::DocumentBody;
use docx_parse::fonts::FontTable;
use docx_parse::header_footer::HeaderFooter;
use docx_parse::media::MediaFile;
use docx_parse::notes::Note;
use docx_parse::numbering::NumberingDefinitions;
use docx_parse::relationships::Relationship;
use docx_parse::settings::DocumentSettings;
use docx_parse::styles::StyleDefinitions;
use docx_parse::theme::Theme;

pub const DEFAULT_SERIALIZATION_TIME: &str = "1970-01-01T00:00:00.000Z";

#[derive(Clone, Debug, PartialEq)]
pub struct DocumentModel {
    pub body: DocumentBody,
    pub styles: Option<StyleDefinitions>,
    pub theme: Theme,
    pub numbering: NumberingDefinitions,
    pub settings: DocumentSettings,
    pub font_table: FontTable,
    pub headers: Vec<(String, HeaderFooter)>,
    pub footers: Vec<(String, HeaderFooter)>,
    pub footnotes: Vec<Note>,
    pub endnotes: Vec<Note>,
    pub footnote_separators: Vec<Note>,
    pub endnote_separators: Vec<Note>,
    pub relationships: Vec<(String, Relationship)>,
    pub media: Vec<(String, MediaFile)>,
    pub charts: Vec<(String, Chart)>,
    pub template_variables: Vec<String>,
    pub warnings: Vec<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct DocumentStructure {
    pub body_paragraphs: usize,
    pub body_tables: usize,
    pub sections: usize,
    pub headers: usize,
    pub footers: usize,
    pub footnotes: usize,
    pub endnotes: usize,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SaveOptions {
    pub now: String,
    pub update_modified_date: bool,
    pub modified_by: Option<String>,
}

impl Default for SaveOptions {
    fn default() -> Self {
        Self {
            now: DEFAULT_SERIALIZATION_TIME.to_owned(),
            update_modified_date: false,
            modified_by: None,
        }
    }
}

#[derive(Debug, Clone)]
pub struct LayoutResult {
    pub layout: Layout,
    pub display_list: DisplayList,
}
