use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result, bail};
use betteroffice_pptx::Presentation;
use betteroffice_xlsx::{SheetId, Workbook, viewport_for_used_range_within};
use docx_edit::{EngineSession, seed_from_docx};
use docx_layout::display_list::DisplayList;
use docx_parse::{S9PackageWire, S9ParseOptions, parse_docx_s9_wire};
use serde_json::{Value, json};

use crate::docx_scene::translate_document;
use crate::editing::{
    CaretGeometry, DeleteDirection, DocxEditor, MoveDirection, SceneChange, SelectionRect, TextLoc,
};
use crate::fonts::FontRegistry;
use crate::images::ImageRegistry;
use crate::pptx_scene::{PptxSlideSummary, translate_presentation};
use crate::scene_shared::PageScene;
use crate::xlsx_scene::translate_sheet;

pub struct DocumentView {
    pub source: PathBuf,
    pub reference: ReferenceDocument,
    pub pages: Vec<PageScene>,
}

pub enum ReferenceDocument {
    Docx(Box<DocxReference>),
    Xlsx(XlsxReference),
    Pptx(PptxReference),
}

pub struct DocxReference {
    pub display_list: DisplayList,
    pub fonts: FontRegistry,
    pub images: ImageRegistry,
    pub editor: DocxEditor,
}

pub struct XlsxReference {
    pub display_list: xlsx_render::DisplayList,
    pub sheet_index: usize,
    pub sheet_count: usize,
    pub sheet_name: String,
    pub chart_count: usize,
    pub chart_placeholders: usize,
}

pub struct PptxReference {
    pub slide_count: usize,
    pub summaries: Vec<PptxSlideSummary>,
    pub font_faces: Vec<String>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DocumentFormat {
    Docx,
    Xlsx,
    Pptx,
}

impl DocumentFormat {
    pub fn from_path(path: &Path) -> Result<Self> {
        match path
            .extension()
            .and_then(|extension| extension.to_str())
            .map(str::to_ascii_lowercase)
            .as_deref()
        {
            Some("docx") => Ok(Self::Docx),
            Some("xlsx") => Ok(Self::Xlsx),
            Some("pptx") => Ok(Self::Pptx),
            _ => bail!("--document must be a .docx, .xlsx, or .pptx file"),
        }
    }
}

impl DocumentView {
    pub fn scene_label(&self, index: usize) -> String {
        match &self.reference {
            ReferenceDocument::Docx(_) => format!("page {}", index + 1),
            ReferenceDocument::Xlsx(reference) => {
                format!("sheet {}", reference.sheet_index + 1)
            }
            ReferenceDocument::Pptx(_) => format!("slide {}", index + 1),
        }
    }

    pub fn title_summary(&self) -> String {
        match &self.reference {
            ReferenceDocument::Docx(_) => format!("{} pages", self.pages.len()),
            ReferenceDocument::Xlsx(reference) => format!(
                "sheet {} of {} ({})",
                reference.sheet_index + 1,
                reference.sheet_count,
                reference.sheet_name
            ),
            ReferenceDocument::Pptx(reference) => format!("{} slides", reference.slide_count),
        }
    }

    pub fn display_item_name(&self) -> &'static str {
        match &self.reference {
            ReferenceDocument::Docx(_) => "primitives",
            ReferenceDocument::Xlsx(_) => "commands",
            ReferenceDocument::Pptx(_) => "primitives",
        }
    }

    pub fn docx_hit_test(&self, page_index: usize, x: f64, y: f64) -> Result<Option<TextLoc>> {
        let ReferenceDocument::Docx(reference) = &self.reference else {
            return Ok(None);
        };
        reference.editor.hit_test(page_index, x, y)
    }

    pub fn docx_select_point(&mut self, loc: TextLoc, extend: bool, word: bool) -> Result<bool> {
        let ReferenceDocument::Docx(reference) = &mut self.reference else {
            return Ok(false);
        };
        reference.editor.select_point(loc, extend, word)?;
        Ok(true)
    }

    pub fn docx_extend_to(&mut self, loc: TextLoc) -> Result<bool> {
        let ReferenceDocument::Docx(reference) = &mut self.reference else {
            return Ok(false);
        };
        reference.editor.extend_to(loc)?;
        Ok(true)
    }

    pub fn docx_move_selection(&mut self, direction: MoveDirection, extend: bool) -> Result<bool> {
        let ReferenceDocument::Docx(reference) = &mut self.reference else {
            return Ok(false);
        };
        reference.editor.move_selection(direction, extend)
    }

    pub fn docx_insert_text(&mut self, text: &str) -> Result<bool> {
        self.apply_docx_edit(|editor| editor.insert_text(text))
    }

    pub fn docx_delete(&mut self, direction: DeleteDirection) -> Result<bool> {
        self.apply_docx_edit(|editor| editor.delete(direction))
    }

    pub fn docx_enter(&mut self) -> Result<bool> {
        self.apply_docx_edit(DocxEditor::enter)
    }

    pub fn docx_selection_rects(&self) -> &[SelectionRect] {
        let ReferenceDocument::Docx(reference) = &self.reference else {
            return &[];
        };
        reference.editor.selection_rects()
    }

    pub fn docx_caret_geometry(&self) -> Result<Option<CaretGeometry>> {
        let ReferenceDocument::Docx(reference) = &self.reference else {
            return Ok(None);
        };
        reference.editor.caret_geometry()
    }

    pub fn has_docx_caret(&self) -> bool {
        matches!(
            &self.reference,
            ReferenceDocument::Docx(reference) if reference.editor.has_collapsed_selection()
        )
    }

    pub fn save_docx_to(&mut self, path: &Path) -> Result<()> {
        let ReferenceDocument::Docx(reference) = &mut self.reference else {
            bail!("only DOCX documents are editable");
        };
        reference.editor.save_to(path)
    }

    pub fn edited_path(&self) -> Result<PathBuf> {
        let stem = self
            .source
            .file_stem()
            .and_then(|stem| stem.to_str())
            .context("DOCX input has no UTF-8 file stem")?;
        let parent = self.source.parent().unwrap_or_else(|| Path::new("."));
        Ok(parent.join(format!("{stem}-edited.docx")))
    }

    fn apply_docx_edit(
        &mut self,
        edit: impl FnOnce(&mut DocxEditor) -> Result<Option<SceneChange>>,
    ) -> Result<bool> {
        let ReferenceDocument::Docx(reference) = &mut self.reference else {
            return Ok(false);
        };
        let Some(change) = edit(&mut reference.editor)? else {
            return Ok(false);
        };
        let display_list = reference
            .editor
            .engine()
            .with_display_list(Clone::clone)
            .context("engine did not retain an edited display list")?;
        if change.rebuild_all || display_list.pages.len() != self.pages.len() {
            self.pages = translate_document(&display_list, &reference.fonts, &reference.images)?;
        } else {
            for page_index in change.changed_pages {
                let page = display_list
                    .pages
                    .get(page_index)
                    .with_context(|| format!("edited page {} is unavailable", page_index + 1))?;
                self.pages[page_index] =
                    crate::docx_scene::translate_page(page, &reference.fonts, &reference.images)?;
            }
        }
        reference.display_list = display_list;
        Ok(true)
    }
}

pub fn load_document(path: &Path, sheet_index: usize) -> Result<DocumentView> {
    match DocumentFormat::from_path(path)? {
        DocumentFormat::Docx => load_docx(path),
        DocumentFormat::Xlsx => load_xlsx(path, sheet_index),
        DocumentFormat::Pptx => load_pptx(path),
    }
}

fn load_pptx(path: &Path) -> Result<DocumentView> {
    let bytes = fs::read(path).with_context(|| format!("read PPTX {}", path.display()))?;
    let presentation =
        Presentation::open(&bytes).with_context(|| format!("open PPTX {}", path.display()))?;
    let slide_count = presentation.slides().len();
    let translation = translate_presentation(&presentation)?;
    Ok(DocumentView {
        source: path.to_owned(),
        reference: ReferenceDocument::Pptx(PptxReference {
            slide_count,
            summaries: translation.summaries,
            font_faces: translation.font_faces,
        }),
        pages: translation.pages,
    })
}

fn load_docx(path: &Path) -> Result<DocumentView> {
    let bytes = fs::read(path).with_context(|| format!("read DOCX {}", path.display()))?;
    let parsed = parse_docx_s9_wire(&bytes, S9ParseOptions::default())?;
    let package = parsed.document.package;
    let engine = EngineSession::new(0x0056_454c_4c4f);
    seed_from_docx(engine.doc(), &bytes).map_err(anyhow::Error::msg)?;

    let probe = region_request(&package, None)?;
    let requirements_json = engine
        .layout_font_requirements_json(&probe.to_string())
        .map_err(anyhow::Error::msg)?;
    let requirements: Value = serde_json::from_str(&requirements_json)?;
    let fonts = FontRegistry::load(&requirements)?;
    let request = region_request(&package, Some(&fonts.chain_ids))?;
    engine
        .layout_document_with_regions_json(&request.to_string())
        .map_err(anyhow::Error::msg)?;
    let extras = json!({ "fontChains": fonts.chain_ids }).to_string();
    engine
        .build_display_list_frame(&extras, 0)
        .map_err(anyhow::Error::msg)?;
    let initial_frame = engine
        .apply_and_layout("body", 0)
        .map_err(anyhow::Error::msg)?;
    let display_list = engine
        .with_display_list(Clone::clone)
        .context("engine did not retain a display list")?;
    let images = ImageRegistry::load(&bytes)?;
    let pages = translate_document(&display_list, &fonts, &images)?;
    let editor = DocxEditor::new(engine, bytes, package, &initial_frame)?;
    Ok(DocumentView {
        source: path.to_owned(),
        reference: ReferenceDocument::Docx(Box::new(DocxReference {
            display_list,
            fonts,
            images,
            editor,
        })),
        pages,
    })
}

fn load_xlsx(path: &Path, sheet_index: usize) -> Result<DocumentView> {
    let bytes = fs::read(path).with_context(|| format!("read XLSX {}", path.display()))?;
    let workbook =
        Workbook::open_for_read(&bytes).with_context(|| format!("open XLSX {}", path.display()))?;
    let sheet_count = workbook.sheet_count();
    if sheet_index >= sheet_count {
        bail!(
            "sheet {} is out of range for a workbook with {sheet_count} sheets",
            sheet_index + 1
        );
    }
    let sheet_id = SheetId(u32::try_from(sheet_index).context("sheet index is too large")?);
    let sheet = workbook.sheet(sheet_id)?;
    let sheet_name = sheet.name.clone();
    let viewport = viewport_for_used_range_within(sheet, |viewport| {
        viewport.width.is_finite()
            && viewport.height.is_finite()
            && viewport.width.ceil() <= 16_384.0
            && viewport.height.ceil() <= 16_384.0
    });
    let display_list = workbook.display_list_for(sheet_id, &viewport)?;
    let page = translate_sheet(&display_list)?;
    let chart_count = display_list.charts.len();
    let chart_placeholders = display_list
        .charts
        .iter()
        .filter(|chart| chart.placeholder)
        .count();
    Ok(DocumentView {
        source: path.to_owned(),
        reference: ReferenceDocument::Xlsx(XlsxReference {
            display_list,
            sheet_index,
            sheet_count,
            sheet_name,
            chart_count,
            chart_placeholders,
        }),
        pages: vec![page],
    })
}

fn region_request(
    package: &S9PackageWire,
    font_chains: Option<&BTreeMap<String, Vec<u32>>>,
) -> Result<Value> {
    let body = &package.document;
    let mut sections = body
        .sections
        .as_deref()
        .unwrap_or_default()
        .iter()
        .map(|section| {
            let properties = serde_json::to_value(&section.properties)?;
            let section_id = section
                .id
                .clone()
                .or_else(|| properties["sectionId"].as_str().map(str::to_owned));
            Ok(json!({ "sectionId": section_id, "properties": properties }))
        })
        .collect::<Result<Vec<_>>>()?;
    let final_properties = body
        .final_section_properties
        .as_ref()
        .map(serde_json::to_value)
        .transpose()?
        .unwrap_or_else(|| json!({}));
    sections.push(json!({
        "sectionId": final_properties["sectionId"].as_str(),
        "properties": final_properties
    }));
    let watermark = sections
        .last()
        .and_then(|section| section["properties"].get("watermark"))
        .cloned()
        .unwrap_or(Value::Null);
    let mut notes = Vec::new();
    push_notes(&mut notes, package.footnotes.as_deref(), "footnote");
    push_notes(&mut notes, package.endnotes.as_deref(), "endnote");
    let compatibility = &package.settings.compatibility_flags;
    let measurement = font_chains.map(|chains| {
        json!({
            "fontChains": chains,
            "defaults": { "fontSize": 11, "fontFamily": "Calibri" },
            "compat": {
                "noLeading": compatibility.no_leading,
                "doNotExpandShiftReturn": compatibility.do_not_expand_shift_return
            },
            "authoritativeShaping": true
        })
    });
    let mut request = json!({
        "bodyStory": "body",
        "options": { "pageGap": 24 },
        "regions": {
            "sections": sections,
            "settings": package.settings,
            "watermark": watermark
        },
        "notes": { "contents": notes },
        "renderEnv": {}
    });
    if let Some(measurement) = measurement {
        request["measurement"] = measurement;
    }
    Ok(request)
}

fn push_notes(target: &mut Vec<Value>, notes: Option<&[docx_parse::Note]>, note_kind: &str) {
    for note in notes.unwrap_or_default() {
        if !note.is_separator() {
            target.push(json!({
                "id": note.id,
                "noteKind": note_kind,
                "height": 0
            }));
        }
    }
}

#[cfg(test)]
mod tests {
    use std::sync::atomic::{AtomicU64, Ordering};

    use super::*;
    use crate::editing::TextLoc;

    static TEST_FILE_ID: AtomicU64 = AtomicU64::new(0);

    #[test]
    fn recognizes_supported_extensions() {
        assert_eq!(
            DocumentFormat::from_path(Path::new("document.DOCX")).unwrap(),
            DocumentFormat::Docx
        );
        assert_eq!(
            DocumentFormat::from_path(Path::new("workbook.XLSX")).unwrap(),
            DocumentFormat::Xlsx
        );
        assert_eq!(
            DocumentFormat::from_path(Path::new("slides.PPTX")).unwrap(),
            DocumentFormat::Pptx
        );
    }

    #[test]
    fn loads_showcase_with_native_chart_commands() {
        let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/showcase.xlsx");
        let document = load_document(&path, 0).unwrap();
        let ReferenceDocument::Xlsx(reference) = &document.reference else {
            panic!("expected XLSX reference");
        };
        assert_eq!(reference.sheet_name, "Dashboard");
        assert_eq!(reference.chart_count, 1);
        assert_eq!(reference.chart_placeholders, 0);
        assert_eq!(document.pages[0].skipped.total(), 0);
    }

    #[test]
    fn loads_demo_presentation_and_audits_positioned_glyphs() {
        let path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let document = load_document(&path, 0).unwrap();
        let ReferenceDocument::Pptx(reference) = &document.reference else {
            panic!("expected PPTX reference");
        };
        assert_eq!(reference.slide_count, 3);
        assert_eq!(document.pages.len(), 3);
        for summary in &reference.summaries {
            assert!(summary.glyph_audit.glyph_runs > 0);
            assert_eq!(summary.glyph_audit.drifted_caret_stops, 0);
            assert_eq!(summary.glyph_audit.missing_caret_stops, 0);
            assert_eq!(summary.glyph_audit.drifted_widths, 0);
        }
        assert_eq!(document.pages[0].skipped.total(), 0);
        assert_eq!(
            reference.summaries[0].structured(&document.pages[0].skipped)["primitives"]["image"]["translated"],
            1
        );
        assert_eq!(document.pages[1].skipped.counts["placeholder"], 1);
    }

    #[test]
    fn translates_chart_fixture_sub_primitives() {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../crates/pptx-parse/tests/fixtures/chart-deck.pptx");
        let document = load_document(&path, 0).unwrap();
        let ReferenceDocument::Pptx(reference) = &document.reference else {
            panic!("expected PPTX reference");
        };
        assert_eq!(reference.slide_count, 2);
        let summary = reference.summaries[0].structured(&document.pages[0].skipped);
        assert_eq!(summary["primitives"]["chart"]["translated"], 1);
        assert_eq!(summary["primitives"]["placeholder"]["skipped"], 1);
        assert!(reference.summaries[0].glyph_audit.missing_caret_stops > 0);
    }

    #[test]
    fn edits_saves_and_reopens_docx_through_the_viewer_path() {
        let source =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.docx");
        let source_before = fs::read(&source).unwrap();
        let mut document = load_document(&source, 0).unwrap();
        let (para_id, before, paragraph_count) = match &document.reference {
            ReferenceDocument::Docx(reference) => {
                let paragraphs = reference.editor.engine().doc().paragraphs("body").unwrap();
                let paragraph = paragraphs.first().unwrap();
                (
                    paragraph.para_id.clone(),
                    paragraph.text.clone(),
                    paragraphs.len(),
                )
            }
            _ => panic!("expected DOCX reference"),
        };
        assert_eq!(before, "Welcome to BetterOffice");

        let caret = match &document.reference {
            ReferenceDocument::Docx(reference) => reference
                .editor
                .engine()
                .resident_caret_snapshot(Some((&para_id, 7)))
                .unwrap()
                .caret_rect
                .unwrap(),
            _ => panic!("expected DOCX reference"),
        };
        let loc = document
            .docx_hit_test(caret.page_index, caret.x, caret.y + caret.height * 0.5)
            .unwrap()
            .unwrap();
        assert_eq!(loc.para_id, para_id);
        assert_eq!(loc.offset, 7);
        document.docx_select_point(loc, false, false).unwrap();
        assert!(document.docx_insert_text(" native").unwrap());

        let output = std::env::temp_dir().join(format!(
            "betteroffice-native-viewer-{}-{}.docx",
            std::process::id(),
            TEST_FILE_ID.fetch_add(1, Ordering::Relaxed)
        ));
        document.save_docx_to(&output).unwrap();
        let reopened = load_document(&output, 0).unwrap();
        let (after, reopened_count) = match &reopened.reference {
            ReferenceDocument::Docx(reference) => {
                let paragraphs = reference.editor.engine().doc().paragraphs("body").unwrap();
                (paragraphs.first().unwrap().text.clone(), paragraphs.len())
            }
            _ => panic!("expected DOCX reference"),
        };
        assert_eq!(after, "Welcome native to BetterOffice");
        assert_eq!(reopened_count, paragraph_count);
        assert_eq!(fs::read(&source).unwrap(), source_before);
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn applies_delete_and_enter_through_the_viewer_path() {
        let source =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.docx");
        let mut document = load_document(&source, 0).unwrap();
        let (para_id, paragraph_count) = match &document.reference {
            ReferenceDocument::Docx(reference) => {
                let paragraphs = reference.editor.engine().doc().paragraphs("body").unwrap();
                (paragraphs[0].para_id.clone(), paragraphs.len())
            }
            _ => panic!("expected DOCX reference"),
        };
        let caret = TextLoc { para_id, offset: 7 };

        document
            .docx_select_point(caret.clone(), false, false)
            .unwrap();
        let first_line = document.docx_caret_geometry().unwrap().unwrap();
        assert!(
            document
                .docx_move_selection(MoveDirection::Down, false)
                .unwrap()
        );
        let next_line = document.docx_caret_geometry().unwrap().unwrap();
        assert!(
            next_line.page_index > first_line.page_index
                || (next_line.page_index == first_line.page_index && next_line.y > first_line.y)
        );
        assert!(
            document
                .docx_move_selection(MoveDirection::Up, false)
                .unwrap()
        );
        document
            .docx_select_point(caret.clone(), false, false)
            .unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: caret.para_id.clone(),
                    offset: 10,
                },
                true,
                false,
            )
            .unwrap();
        assert!(!document.docx_selection_rects().is_empty());
        document
            .docx_select_point(
                TextLoc {
                    para_id: caret.para_id.clone(),
                    offset: 8,
                },
                false,
                true,
            )
            .unwrap();
        assert!(!document.docx_selection_rects().is_empty());
        document
            .docx_select_point(caret.clone(), false, false)
            .unwrap();
        assert!(
            document
                .docx_move_selection(MoveDirection::Home, false)
                .unwrap()
        );
        let home = document.docx_caret_geometry().unwrap().unwrap();
        assert!(
            document
                .docx_move_selection(MoveDirection::End, false)
                .unwrap()
        );
        let end = document.docx_caret_geometry().unwrap().unwrap();
        assert_eq!(home.page_index, end.page_index);
        assert!(end.x > home.x);
        document
            .docx_select_point(caret.clone(), false, false)
            .unwrap();
        assert!(document.docx_delete(DeleteDirection::Forward).unwrap());
        assert!(document.docx_insert_text(" ").unwrap());
        assert!(document.docx_insert_text("😀").unwrap());
        assert!(document.docx_delete(DeleteDirection::Backward).unwrap());
        document.docx_select_point(caret, false, false).unwrap();
        assert!(document.docx_enter().unwrap());
        assert!(document.docx_caret_geometry().unwrap().is_some());

        let paragraphs = match &document.reference {
            ReferenceDocument::Docx(reference) => {
                reference.editor.engine().doc().paragraphs("body").unwrap()
            }
            _ => panic!("expected DOCX reference"),
        };
        assert_eq!(paragraphs.len(), paragraph_count + 1);
        assert_eq!(paragraphs[0].text, "Welcome");
        assert_eq!(paragraphs[1].text, " to BetterOffice");

        let output = std::env::temp_dir().join(format!(
            "betteroffice-native-viewer-{}-{}.docx",
            std::process::id(),
            TEST_FILE_ID.fetch_add(1, Ordering::Relaxed)
        ));
        document.save_docx_to(&output).unwrap();
        let reopened = load_document(&output, 0).unwrap();
        let reopened_paragraphs = match &reopened.reference {
            ReferenceDocument::Docx(reference) => {
                reference.editor.engine().doc().paragraphs("body").unwrap()
            }
            _ => panic!("expected DOCX reference"),
        };
        assert_eq!(reopened_paragraphs.len(), paragraph_count + 1);
        assert_eq!(reopened_paragraphs[0].text, "Welcome");
        assert_eq!(reopened_paragraphs[1].text, " to BetterOffice");
        fs::remove_file(output).unwrap();

        assert!(document.docx_delete(DeleteDirection::Backward).unwrap());
        let paragraphs = match &document.reference {
            ReferenceDocument::Docx(reference) => {
                reference.editor.engine().doc().paragraphs("body").unwrap()
            }
            _ => panic!("expected DOCX reference"),
        };
        assert_eq!(paragraphs.len(), paragraph_count);
        assert_eq!(paragraphs[0].text, "Welcome to BetterOffice");
    }
}
