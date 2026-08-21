use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result, bail};
use betteroffice_pptx::Presentation;
use betteroffice_xlsx::CellRef;
use docx_edit::{EngineSession, SimpleFormat, seed_from_docx};
use docx_layout::display_list::DisplayList;
use docx_parse::{S9PackageWire, S9ParseOptions, parse_docx_s9_wire};
use serde_json::{Value, json};

use crate::chrome::{Alignment, EditingState};
use crate::docx_scene::translate_document;
use crate::editing::{
    CaretGeometry, DeleteDirection, DocxEditor, MoveDirection, SceneChange, SelectionRect, TextLoc,
};
use crate::fonts::FontRegistry;
use crate::images::ImageRegistry;
use crate::pptx_scene::{PptxSlideSummary, translate_presentation};
use crate::scene_shared::PageScene;
use crate::xlsx_editing::{CellMove, XlsxEditor};
use crate::xlsx_scene::translate_sheet;

pub struct DocumentView {
    pub source: PathBuf,
    pub reference: ReferenceDocument,
    pub pages: Vec<PageScene>,
    pub max_texture_dimension_2d: u32,
}

pub enum ReferenceDocument {
    Docx(Box<DocxReference>),
    Xlsx(Box<XlsxReference>),
    Pptx(PptxReference),
}

pub struct DocxReference {
    pub display_list: DisplayList,
    pub fonts: FontRegistry,
    pub images: ImageRegistry,
    pub editor: DocxEditor,
}

pub struct XlsxReference {
    pub editor: XlsxEditor,
    pub sheet_index: usize,
    pub sheet_count: usize,
    pub sheet_name: String,
    pub chart_count: usize,
    pub chart_placeholders: usize,
}

pub struct XlsxOverlay {
    pub rect: xlsx_render::Rect,
    pub draft: Option<String>,
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

    pub fn status_position(&self, page_index: usize) -> String {
        match &self.reference {
            ReferenceDocument::Docx(_) => {
                format!("Page {} of {}", page_index + 1, self.pages.len())
            }
            ReferenceDocument::Xlsx(reference) => {
                let cell = reference
                    .editor
                    .status()
                    .unwrap_or_else(|_| reference.editor.selection().to_a1());
                format!(
                    "Sheet {} of {} · {cell}",
                    reference.sheet_index + 1,
                    reference.sheet_count
                )
            }
            ReferenceDocument::Pptx(reference) => {
                format!("Slide {} of {}", page_index + 1, reference.slide_count)
            }
        }
    }

    pub fn editing_state(&self) -> Result<EditingState> {
        match &self.reference {
            ReferenceDocument::Docx(reference) => reference.editor.editing_state(),
            ReferenceDocument::Xlsx(reference) => Ok(EditingState::editable_without_selection(
                reference.editor.can_undo(),
                reference.editor.can_redo(),
            )),
            ReferenceDocument::Pptx(_) => Ok(EditingState::read_only()),
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

    pub fn docx_toggle_format(&mut self, format: SimpleFormat) -> Result<bool> {
        self.apply_docx_edit(|editor| editor.toggle_format(format))
    }

    pub fn docx_set_alignment(&mut self, alignment: Alignment) -> Result<bool> {
        self.apply_docx_edit(|editor| editor.set_alignment(alignment))
    }

    pub fn docx_undo(&mut self) -> Result<bool> {
        self.apply_docx_edit(DocxEditor::undo)
    }

    pub fn docx_redo(&mut self) -> Result<bool> {
        self.apply_docx_edit(DocxEditor::redo)
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

    pub fn xlsx_hit_test(&self, page_index: usize, x: f64, y: f64) -> Option<CellRef> {
        let ReferenceDocument::Xlsx(reference) = &self.reference else {
            return None;
        };
        (page_index == 0)
            .then(|| reference.editor.cell_at_point(x as f32, y as f32))
            .flatten()
    }

    pub fn xlsx_select_cell(&mut self, cell: CellRef) -> bool {
        let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
            return false;
        };
        reference.editor.select(cell)
    }

    pub fn xlsx_move_selection(&mut self, movement: CellMove) -> bool {
        let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
            return false;
        };
        reference.editor.move_selection(movement)
    }

    pub fn xlsx_begin_edit(&mut self, seed: Option<&str>) -> Result<bool> {
        let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
            return Ok(false);
        };
        reference.editor.begin_edit(seed)
    }

    pub fn xlsx_insert_text(&mut self, text: &str) -> bool {
        let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
            return false;
        };
        reference.editor.insert_text(text)
    }

    pub fn xlsx_backspace(&mut self) -> bool {
        let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
            return false;
        };
        reference.editor.backspace()
    }

    pub fn xlsx_cancel_edit(&mut self) -> bool {
        let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
            return false;
        };
        reference.editor.cancel()
    }

    pub fn xlsx_commit(&mut self, movement: Option<CellMove>) -> Result<bool> {
        let changed = {
            let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
                return Ok(false);
            };
            reference.editor.commit(movement)?
        };
        if changed {
            self.refresh_xlsx_scene()?;
        }
        Ok(changed)
    }

    pub fn xlsx_undo(&mut self) -> Result<bool> {
        let changed = {
            let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
                return Ok(false);
            };
            reference.editor.undo()?
        };
        if changed {
            self.refresh_xlsx_scene()?;
        }
        Ok(changed)
    }

    pub fn xlsx_redo(&mut self) -> Result<bool> {
        let changed = {
            let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
                return Ok(false);
            };
            reference.editor.redo()?
        };
        if changed {
            self.refresh_xlsx_scene()?;
        }
        Ok(changed)
    }

    pub fn xlsx_is_editing(&self) -> bool {
        matches!(
            &self.reference,
            ReferenceDocument::Xlsx(reference) if reference.editor.is_editing()
        )
    }

    pub fn xlsx_is_dirty(&self) -> bool {
        matches!(
            &self.reference,
            ReferenceDocument::Xlsx(reference) if reference.editor.is_dirty()
        )
    }

    pub fn xlsx_overlay(&self) -> Option<XlsxOverlay> {
        let ReferenceDocument::Xlsx(reference) = &self.reference else {
            return None;
        };
        Some(XlsxOverlay {
            rect: reference.editor.selection_rect()?,
            draft: reference.editor.draft_value().map(str::to_owned),
        })
    }

    pub fn save_xlsx_to(&self, path: &Path) -> Result<()> {
        let ReferenceDocument::Xlsx(reference) = &self.reference else {
            bail!("only XLSX workbooks use the XLSX save path");
        };
        reference.editor.save_to(path)
    }

    pub fn edited_path(&self) -> Result<PathBuf> {
        let extension = match &self.reference {
            ReferenceDocument::Docx(_) => "docx",
            ReferenceDocument::Xlsx(_) => "xlsx",
            ReferenceDocument::Pptx(_) => bail!("PPTX documents are read-only"),
        };
        let stem = self
            .source
            .file_stem()
            .and_then(|stem| stem.to_str())
            .context("input has no UTF-8 file stem")?;
        let parent = self.source.parent().unwrap_or_else(|| Path::new("."));
        Ok(parent.join(format!("{stem}-edited.{extension}")))
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

    fn refresh_xlsx_scene(&mut self) -> Result<()> {
        let ReferenceDocument::Xlsx(reference) = &mut self.reference else {
            return Ok(());
        };
        let display_list = reference.editor.display_list();
        let page = translate_sheet(display_list)?;
        reference.chart_count = display_list.charts.len();
        reference.chart_placeholders = display_list
            .charts
            .iter()
            .filter(|chart| chart.placeholder)
            .count();
        self.pages[0] = page;
        Ok(())
    }
}

pub fn load_document(
    path: &Path,
    sheet_index: usize,
    max_texture_dimension_2d: u32,
) -> Result<DocumentView> {
    load_document_with_xlsx_mode(path, sheet_index, max_texture_dimension_2d, true)
}

pub fn load_document_for_export(
    path: &Path,
    sheet_index: usize,
    max_texture_dimension_2d: u32,
) -> Result<DocumentView> {
    load_document_with_xlsx_mode(path, sheet_index, max_texture_dimension_2d, false)
}

fn load_document_with_xlsx_mode(
    path: &Path,
    sheet_index: usize,
    max_texture_dimension_2d: u32,
    recalculate_xlsx: bool,
) -> Result<DocumentView> {
    match DocumentFormat::from_path(path)? {
        DocumentFormat::Docx => load_docx(path, max_texture_dimension_2d),
        DocumentFormat::Xlsx => load_xlsx(
            path,
            sheet_index,
            max_texture_dimension_2d,
            recalculate_xlsx,
        ),
        DocumentFormat::Pptx => load_pptx(path, max_texture_dimension_2d),
    }
}

fn load_pptx(path: &Path, max_texture_dimension_2d: u32) -> Result<DocumentView> {
    let bytes = fs::read(path).with_context(|| format!("read PPTX {}", path.display()))?;
    let presentation =
        Presentation::open(&bytes).with_context(|| format!("open PPTX {}", path.display()))?;
    let slide_count = presentation.slides().len();
    let translation = translate_presentation(&presentation, max_texture_dimension_2d)?;
    Ok(DocumentView {
        source: path.to_owned(),
        reference: ReferenceDocument::Pptx(PptxReference {
            slide_count,
            summaries: translation.summaries,
            font_faces: translation.font_faces,
        }),
        pages: translation.pages,
        max_texture_dimension_2d,
    })
}

fn load_docx(path: &Path, max_texture_dimension_2d: u32) -> Result<DocumentView> {
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
        max_texture_dimension_2d,
    })
}

fn load_xlsx(
    path: &Path,
    sheet_index: usize,
    max_texture_dimension_2d: u32,
    recalculate: bool,
) -> Result<DocumentView> {
    let bytes = fs::read(path).with_context(|| format!("read XLSX {}", path.display()))?;
    let editor = XlsxEditor::open(bytes, sheet_index, max_texture_dimension_2d, recalculate)
        .map_err(|error| anyhow::anyhow!("open XLSX {}: {error:#}", path.display()))?;
    let sheet_count = editor.sheet_count();
    let sheet_name = editor.sheet_name()?.to_owned();
    let display_list = editor.display_list();
    let page = translate_sheet(display_list)?;
    let chart_count = display_list.charts.len();
    let chart_placeholders = display_list
        .charts
        .iter()
        .filter(|chart| chart.placeholder)
        .count();
    Ok(DocumentView {
        source: path.to_owned(),
        reference: ReferenceDocument::Xlsx(Box::new(XlsxReference {
            editor,
            sheet_index,
            sheet_count,
            sheet_name,
            chart_count,
            chart_placeholders,
        })),
        pages: vec![page],
        max_texture_dimension_2d,
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
                "id": note.id as i64,
                "noteKind": note_kind,
                "height": 0
            }));
        }
    }
}

#[cfg(test)]
mod tests {
    use std::sync::atomic::{AtomicU64, Ordering};

    use betteroffice_xlsx::{CellValue, DrawCmd, SheetId};

    use super::*;
    use crate::chrome::{Alignment, ToggleState};
    use crate::editing::TextLoc;
    use crate::test_fixtures;

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
    fn opens_document_with_footnote_and_endnote_regions() {
        let path = Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../crates/docx-edit/tests/fixtures/footnote-anchor.docx");
        let document = load_document(&path, 0, 8_192).unwrap();
        assert!(!document.pages.is_empty());
        assert!(
            document
                .pages
                .iter()
                .any(|page| !page.scene.encoding().is_empty())
        );
    }

    #[test]
    fn preserves_untouched_complex_paragraph_and_refuses_editing_it() {
        let source_bytes = test_fixtures::complex_docx();
        let complex_before = test_fixtures::paragraph(&source_bytes, "11111111");
        let source = test_fixtures::write_docx("complex-save", &source_bytes);
        let mut document = load_document(&source, 0, 8_192).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "22222222".to_owned(),
                    offset: 15,
                },
                false,
                false,
            )
            .unwrap();
        assert!(document.docx_insert_text(" edited").unwrap());
        let output = test_fixtures::write_docx("complex-output", b"replace me");
        document.save_docx_to(&output).unwrap();
        let saved = fs::read(&output).unwrap();
        assert_eq!(test_fixtures::paragraph(&saved, "11111111"), complex_before);
        let reopened = load_document(&output, 0, 8_192).unwrap();
        let ReferenceDocument::Docx(reference) = &reopened.reference else {
            panic!("expected DOCX reference");
        };
        let paragraphs = reference.editor.engine().doc().paragraphs("body").unwrap();
        assert_eq!(paragraphs[1].text, "Plain paragraph edited");

        let mut complex_edit = load_document(&source, 0, 8_192).unwrap();
        complex_edit
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 1,
                },
                false,
                false,
            )
            .unwrap();
        assert!(complex_edit.docx_insert_text("x").unwrap());
        let refused = test_fixtures::write_docx("complex-refused", b"sentinel");
        let error = complex_edit.save_docx_to(&refused).unwrap_err().to_string();
        assert!(error.contains("paragraph 1 (11111111)"));
        assert!(error.contains("hyperlinks"));
        assert!(error.contains("note references"));
        assert_eq!(fs::read(&refused).unwrap(), b"sentinel");

        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
        fs::remove_file(refused).unwrap();
    }

    #[test]
    fn derives_toolbar_marks_from_the_engine_selection_context() {
        let source = test_fixtures::write_docx("toolbar-state", &test_fixtures::complex_docx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 2,
                },
                false,
                false,
            )
            .unwrap();
        let caret = document.editing_state().unwrap();
        assert_eq!(caret.bold, ToggleState::On);
        assert!(!caret.inline_enabled);

        document
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 5,
                },
                true,
                false,
            )
            .unwrap();
        let mixed = document.editing_state().unwrap();
        assert_eq!(mixed.bold, ToggleState::Mixed);
        assert!(mixed.inline_enabled);
        fs::remove_file(source).unwrap();
    }

    #[test]
    fn format_preserves_selection_and_undo_updates_toolbar_state() {
        let source = test_fixtures::write_docx("toolbar-format", &test_fixtures::complex_docx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let start = TextLoc {
            para_id: "22222222".to_owned(),
            offset: 0,
        };
        let end = TextLoc {
            para_id: "22222222".to_owned(),
            offset: 5,
        };
        document.docx_select_point(start, false, false).unwrap();
        document.docx_select_point(end, true, false).unwrap();
        let before = match &document.reference {
            ReferenceDocument::Docx(reference) => reference.editor.selection_range().unwrap(),
            _ => unreachable!(),
        };
        let initial = document.editing_state().unwrap();
        assert_eq!(initial.bold, ToggleState::Off);
        assert!(!initial.can_undo);

        assert!(document.docx_toggle_format(SimpleFormat::Bold).unwrap());
        let formatted = document.editing_state().unwrap();
        assert_eq!(formatted.bold, ToggleState::On);
        assert!(formatted.can_undo);
        let after = match &document.reference {
            ReferenceDocument::Docx(reference) => reference.editor.selection_range().unwrap(),
            _ => unreachable!(),
        };
        assert_eq!(after, before);

        assert!(document.docx_toggle_format(SimpleFormat::Italic).unwrap());
        let combined = document.editing_state().unwrap();
        assert_eq!(combined.bold, ToggleState::On);
        assert_eq!(combined.italic, ToggleState::On);
        let after_second = match &document.reference {
            ReferenceDocument::Docx(reference) => reference.editor.selection_range().unwrap(),
            _ => unreachable!(),
        };
        assert_eq!(after_second, before);

        assert!(document.docx_undo().unwrap());
        let undone = document.editing_state().unwrap();
        assert_eq!(undone.bold, ToggleState::On);
        assert_eq!(undone.italic, ToggleState::Off);
        assert!(undone.can_undo);
        assert!(undone.can_redo);
        assert!(document.docx_undo().unwrap());
        let restored = document.editing_state().unwrap();
        assert_eq!(restored.bold, ToggleState::Off);
        assert!(!restored.can_undo);
        fs::remove_file(source).unwrap();
    }

    #[test]
    fn alignment_round_trips_through_save_and_reopen() {
        let source = test_fixtures::write_docx("toolbar-alignment", &test_fixtures::complex_docx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let caret = TextLoc {
            para_id: "22222222".to_owned(),
            offset: 4,
        };
        document
            .docx_select_point(caret.clone(), false, false)
            .unwrap();
        assert!(document.docx_set_alignment(Alignment::Center).unwrap());
        assert_eq!(
            document.editing_state().unwrap().alignment,
            Some(Alignment::Center)
        );

        let output = test_fixtures::write_docx("toolbar-alignment-output", b"sentinel");
        document.save_docx_to(&output).unwrap();
        let mut reopened = load_document(&output, 0, 8_192).unwrap();
        reopened.docx_select_point(caret, false, false).unwrap();
        assert_eq!(
            reopened.editing_state().unwrap().alignment,
            Some(Alignment::Center)
        );
        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn inline_format_round_trips_through_save_and_reopen() {
        let source =
            test_fixtures::write_docx("toolbar-inline-save", &test_fixtures::complex_docx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let start = TextLoc {
            para_id: "22222222".to_owned(),
            offset: 0,
        };
        let end = TextLoc {
            para_id: "22222222".to_owned(),
            offset: 5,
        };
        document
            .docx_select_point(start.clone(), false, false)
            .unwrap();
        document
            .docx_select_point(end.clone(), true, false)
            .unwrap();
        assert!(document.docx_toggle_format(SimpleFormat::Bold).unwrap());

        let output = test_fixtures::write_docx("toolbar-inline-output", b"sentinel");
        document.save_docx_to(&output).unwrap();
        let mut reopened = load_document(&output, 0, 8_192).unwrap();
        reopened.docx_select_point(start, false, false).unwrap();
        reopened.docx_select_point(end, true, false).unwrap();
        assert_eq!(reopened.editing_state().unwrap().bold, ToggleState::On);
        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn translates_embedded_and_rotated_image_fixtures() {
        for rotation in [None, Some(45.0)] {
            let source = test_fixtures::write_docx("image", &test_fixtures::image_docx(rotation));
            let document = load_document(&source, 0, 8_192).unwrap();
            let ReferenceDocument::Docx(reference) = &document.reference else {
                panic!("expected DOCX reference");
            };
            let image = reference
                .display_list
                .pages
                .iter()
                .flat_map(|page| &page.primitives)
                .find_map(|primitive| match primitive {
                    docx_layout::display_list::Primitive::Image(image) => Some(image),
                    _ => None,
                })
                .unwrap();
            assert!(image.rel_id.starts_with("data:image/png;base64,"));
            assert_eq!(image.attrs.content_frame.is_some(), rotation.is_some());
            assert_eq!(
                document
                    .pages
                    .iter()
                    .map(|page| page.skipped.total())
                    .sum::<usize>(),
                0
            );
            fs::remove_file(source).unwrap();
        }
    }

    #[test]
    fn loads_showcase_with_native_chart_commands() {
        let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/showcase.xlsx");
        let document = load_document(&path, 0, 8_192).unwrap();
        let ReferenceDocument::Xlsx(reference) = &document.reference else {
            panic!("expected XLSX reference");
        };
        assert_eq!(reference.sheet_name, "Dashboard");
        assert_eq!(reference.chart_count, 1);
        assert_eq!(reference.chart_placeholders, 0);
        assert_eq!(document.pages[0].skipped.total(), 0);
    }

    #[test]
    fn edits_a_selected_showcase_cell_and_rebuilds_the_sheet() {
        let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/showcase.xlsx");
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let cell = CellRef::parse_a1("B5").unwrap();
        assert!(document.xlsx_select_cell(cell));
        assert!(document.xlsx_move_selection(CellMove::Right));
        assert!(document.xlsx_move_selection(CellMove::Left));
        let selection = document.xlsx_overlay().unwrap().rect;
        assert!(selection.w > 0.0 && selection.h > 0.0);
        assert!(document.xlsx_begin_edit(Some("2048")).unwrap());
        assert_eq!(
            document.xlsx_overlay().unwrap().draft.as_deref(),
            Some("2048")
        );
        assert!(document.xlsx_commit(None).unwrap());

        let ReferenceDocument::Xlsx(reference) = &document.reference else {
            panic!("expected XLSX reference");
        };
        assert_eq!(
            reference
                .editor
                .workbook()
                .cell(reference.editor.sheet(), cell)
                .unwrap()
                .input,
            "2048"
        );
        assert!(
            reference
                .editor
                .display_list()
                .commands
                .iter()
                .any(|command| matches!(command, DrawCmd::Text { text, .. } if text == "2048"))
        );
        let state = document.editing_state().unwrap();
        assert!(state.can_save);
        assert!(state.can_undo);
        assert!(!state.inline_enabled);
        assert!(!state.alignment_enabled);
        assert!(document.status_position(0).contains("B5 = 2048"));
        assert!(document.xlsx_undo().unwrap());
        assert!(!document.xlsx_is_dirty());
        assert!(document.editing_state().unwrap().can_redo);
        let undone = test_fixtures::write_xlsx("undone-output", b"sentinel");
        document.save_xlsx_to(&undone).unwrap();
        assert_eq!(fs::read(&undone).unwrap(), fs::read(&path).unwrap());
        fs::remove_file(undone).unwrap();
        assert!(document.xlsx_redo().unwrap());
        assert!(document.xlsx_is_dirty());
        let ReferenceDocument::Xlsx(reference) = &document.reference else {
            panic!("expected XLSX reference");
        };
        assert_eq!(
            reference
                .editor
                .workbook()
                .cell(reference.editor.sheet(), cell)
                .unwrap()
                .input,
            "2048"
        );
    }

    #[test]
    fn formula_commit_uses_the_engine_recalculation() {
        let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/showcase.xlsx");
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let cell = CellRef::parse_a1("B5").unwrap();
        document.xlsx_select_cell(cell);
        document.xlsx_begin_edit(Some("=40+2")).unwrap();
        document.xlsx_commit(None).unwrap();

        let ReferenceDocument::Xlsx(reference) = &document.reference else {
            panic!("expected XLSX reference");
        };
        let calculated = reference
            .editor
            .workbook()
            .sheet(reference.editor.sheet())
            .unwrap()
            .cell(cell)
            .unwrap();
        assert_eq!(calculated.formula.as_deref(), Some("40+2"));
        assert_eq!(calculated.value, CellValue::Number { value: 42.0 });
        assert!(
            reference
                .editor
                .display_list()
                .commands
                .iter()
                .any(|command| matches!(command, DrawCmd::Text { text, .. } if text == "42"))
        );
        assert!(
            !reference
                .editor
                .display_list()
                .commands
                .iter()
                .any(|command| matches!(command, DrawCmd::Text { text, .. } if text == "=40+2"))
        );
    }

    #[test]
    fn escape_discards_the_cell_draft() {
        let path = Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/showcase.xlsx");
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let cell = CellRef::parse_a1("B5").unwrap();
        document.xlsx_select_cell(cell);
        let before = match &document.reference {
            ReferenceDocument::Xlsx(reference) => {
                reference
                    .editor
                    .workbook()
                    .cell(reference.editor.sheet(), cell)
                    .unwrap()
                    .input
            }
            _ => panic!("expected XLSX reference"),
        };
        document.xlsx_begin_edit(Some("discarded")).unwrap();
        assert!(document.xlsx_insert_text(" draft"));
        assert!(document.xlsx_cancel_edit());
        let after = match &document.reference {
            ReferenceDocument::Xlsx(reference) => {
                reference
                    .editor
                    .workbook()
                    .cell(reference.editor.sheet(), cell)
                    .unwrap()
                    .input
            }
            _ => panic!("expected XLSX reference"),
        };
        assert_eq!(after, before);
        assert!(document.xlsx_overlay().unwrap().draft.is_none());
    }

    #[test]
    fn saves_reopens_and_preserves_an_untouched_sheet() {
        let source = Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/showcase.xlsx");
        let source_bytes = fs::read(&source).unwrap();
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let untouched = match &document.reference {
            ReferenceDocument::Xlsx(reference) => reference
                .editor
                .workbook()
                .sheet(SheetId(2))
                .unwrap()
                .clone(),
            _ => panic!("expected XLSX reference"),
        };
        let cell = CellRef::parse_a1("B5").unwrap();
        document.xlsx_select_cell(cell);
        document.xlsx_begin_edit(Some("1776")).unwrap();
        document.xlsx_commit(None).unwrap();

        let output = test_fixtures::write_xlsx("edited-output", b"sentinel");
        document.save_xlsx_to(&output).unwrap();
        let reopened = load_document(&output, 0, 8_192).unwrap();
        let ReferenceDocument::Xlsx(reference) = &reopened.reference else {
            panic!("expected XLSX reference");
        };
        assert_eq!(
            reference
                .editor
                .workbook()
                .cell(reference.editor.sheet(), cell)
                .unwrap()
                .input,
            "1776"
        );
        assert_eq!(
            reference.editor.workbook().sheet(SheetId(2)).unwrap(),
            &untouched
        );
        let saved = fs::read(&output).unwrap();
        for part in [
            "xl/worksheets/sheet2.xml",
            "xl/worksheets/sheet3.xml",
            "xl/drawings/drawing1.xml",
            "xl/charts/chart1.xml",
            "xl/styles.xml",
            "xl/sharedStrings.xml",
        ] {
            assert_eq!(
                test_fixtures::part(&source_bytes, part),
                test_fixtures::part(&saved, part),
                "{part} changed"
            );
        }
        assert_eq!(fs::read(&source).unwrap(), source_bytes);
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn refuses_to_flatten_shared_formulas_on_save() {
        let source =
            test_fixtures::write_xlsx("shared-formula", &test_fixtures::shared_formula_xlsx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let cell = CellRef::parse_a1("C1").unwrap();
        document.xlsx_select_cell(cell);
        document.xlsx_begin_edit(Some("edited")).unwrap();
        document.xlsx_commit(None).unwrap();

        let output = test_fixtures::write_xlsx("shared-formula-refused", b"sentinel");
        let error = document.save_xlsx_to(&output).unwrap_err().to_string();
        assert!(error.contains("shared formulas"), "{error}");
        assert!(error.contains("xl/worksheets/sheet1.xml"), "{error}");
        assert_eq!(fs::read(&output).unwrap(), b"sentinel");
        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn rejects_sheet_size_over_device_limit() {
        let path = test_fixtures::write_xlsx("large", &test_fixtures::large_xlsx());
        let error = load_document(&path, 0, 8_192).err().unwrap().to_string();
        println!("{error}");
        assert!(error.contains("requested sheet size"));
        assert!(error.contains("ceiling 8192"));
        fs::remove_file(path).unwrap();
    }

    #[test]
    fn loads_demo_presentation_and_audits_positioned_glyphs() {
        let path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let document = load_document(&path, 0, 8_192).unwrap();
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
        let document = load_document(&path, 0, 8_192).unwrap();
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
        let mut document = load_document(&source, 0, 8_192).unwrap();
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
        let reopened = load_document(&output, 0, 8_192).unwrap();
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
        let mut document = load_document(&source, 0, 8_192).unwrap();
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
        fs::write(&output, b"sentinel").unwrap();
        let error = document.save_docx_to(&output).unwrap_err().to_string();
        assert!(error.contains("cannot save DOCX structural edits"));
        assert_eq!(fs::read(&output).unwrap(), b"sentinel");

        assert!(document.docx_delete(DeleteDirection::Backward).unwrap());
        let paragraphs = match &document.reference {
            ReferenceDocument::Docx(reference) => {
                reference.editor.engine().doc().paragraphs("body").unwrap()
            }
            _ => panic!("expected DOCX reference"),
        };
        assert_eq!(paragraphs.len(), paragraph_count);
        assert_eq!(paragraphs[0].text, "Welcome to BetterOffice");

        document.save_docx_to(&output).unwrap();
        let reopened = load_document(&output, 0, 8_192).unwrap();
        let reopened_paragraphs = match &reopened.reference {
            ReferenceDocument::Docx(reference) => {
                reference.editor.engine().doc().paragraphs("body").unwrap()
            }
            _ => panic!("expected DOCX reference"),
        };
        assert_eq!(reopened_paragraphs.len(), paragraph_count);
        assert_eq!(reopened_paragraphs[0].text, "Welcome to BetterOffice");
        fs::remove_file(output).unwrap();
    }
}
