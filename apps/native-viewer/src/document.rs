use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result, bail};
use betteroffice_xlsx::CellRef;
use docx_edit::{EngineSession, SimpleFormat, seed_from_docx};
use docx_layout::display_list::DisplayList;
use docx_parse::{S9PackageWire, S9ParseOptions, parse_docx_s9_wire};
use serde_json::{Value, json};
use vello::kurbo::Affine;

use crate::chrome::{Alignment, EditingState};
use crate::docx_scene::translate_document;
use crate::editing::{
    DeleteDirection, DocxEditor, MoveDirection, SceneChange, SelectionRect, TextLoc,
};
use crate::fonts::FontRegistry;
use crate::images::ImageRegistry;
use crate::pptx_editing::{PptxEditChange, PptxEditor, PptxHit, PptxTextHit};
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
    Pptx(Box<PptxReference>),
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
    pub editor: PptxEditor,
}

#[derive(Clone, Copy, Debug)]
pub struct ViewerCaretGeometry {
    pub page_index: usize,
    pub x: f64,
    pub y: f64,
    pub height: f64,
    pub transform: Affine,
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
            ReferenceDocument::Pptx(reference) => {
                format!("{} slides", reference.editor.slide_count())
            }
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
                format!(
                    "Slide {} of {}",
                    page_index + 1,
                    reference.editor.slide_count()
                )
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
            ReferenceDocument::Pptx(reference) => Ok(EditingState::editable_without_selection(
                reference.editor.can_undo(),
                reference.editor.can_redo(),
            )),
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

    pub fn caret_geometry(&self) -> Result<Option<ViewerCaretGeometry>> {
        match &self.reference {
            ReferenceDocument::Docx(reference) => {
                Ok(reference
                    .editor
                    .caret_geometry()?
                    .map(|caret| ViewerCaretGeometry {
                        page_index: caret.page_index,
                        x: caret.x,
                        y: caret.y,
                        height: caret.height,
                        transform: Affine::IDENTITY,
                    }))
            }
            ReferenceDocument::Pptx(reference) => {
                Ok(reference
                    .editor
                    .caret_geometry()
                    .map(|caret| ViewerCaretGeometry {
                        page_index: caret.page_index,
                        x: caret.x,
                        y: caret.y,
                        height: caret.height,
                        transform: caret.transform,
                    }))
            }
            ReferenceDocument::Xlsx(_) => Ok(None),
        }
    }

    pub fn has_text_caret(&self) -> bool {
        match &self.reference {
            ReferenceDocument::Docx(reference) => reference.editor.has_collapsed_selection(),
            ReferenceDocument::Pptx(reference) => reference.editor.has_caret(),
            ReferenceDocument::Xlsx(_) => false,
        }
    }

    pub fn save_docx_to(&mut self, path: &Path) -> Result<()> {
        let ReferenceDocument::Docx(reference) = &mut self.reference else {
            bail!("only DOCX documents are editable");
        };
        reference.editor.save_to(path)
    }

    pub fn pptx_hit_test(&self, slide_index: usize, x: f64, y: f64) -> Option<PptxHit> {
        let ReferenceDocument::Pptx(reference) = &self.reference else {
            return None;
        };
        reference.editor.hit_test(slide_index, x as f32, y as f32)
    }

    pub fn pptx_select_hit(&mut self, hit: PptxTextHit) -> bool {
        let ReferenceDocument::Pptx(reference) = &mut self.reference else {
            return false;
        };
        reference.editor.select_hit(hit)
    }

    pub fn pptx_clear_caret(&mut self) -> bool {
        let ReferenceDocument::Pptx(reference) = &mut self.reference else {
            return false;
        };
        reference.editor.clear_caret()
    }

    pub fn pptx_move_caret(&mut self, direction: MoveDirection) -> bool {
        let ReferenceDocument::Pptx(reference) = &mut self.reference else {
            return false;
        };
        reference.editor.move_caret(direction)
    }

    pub fn pptx_insert_text(&mut self, text: &str) -> Result<bool> {
        self.apply_pptx_edit(|editor| editor.insert_text(text))
    }

    pub fn pptx_delete(&mut self, direction: DeleteDirection) -> Result<bool> {
        self.apply_pptx_edit(|editor| editor.delete(direction))
    }

    pub fn pptx_enter(&mut self) -> Result<bool> {
        self.apply_pptx_edit(PptxEditor::enter)
    }

    pub fn pptx_undo(&mut self) -> Result<bool> {
        let pages = {
            let ReferenceDocument::Pptx(reference) = &mut self.reference else {
                return Ok(false);
            };
            reference.editor.undo()?
        };
        let Some(pages) = pages else {
            return Ok(false);
        };
        self.pages = pages;
        Ok(true)
    }

    pub fn pptx_redo(&mut self) -> Result<bool> {
        let pages = {
            let ReferenceDocument::Pptx(reference) = &mut self.reference else {
                return Ok(false);
            };
            reference.editor.redo()?
        };
        let Some(pages) = pages else {
            return Ok(false);
        };
        self.pages = pages;
        Ok(true)
    }

    pub fn pptx_is_dirty(&self) -> bool {
        matches!(
            &self.reference,
            ReferenceDocument::Pptx(reference) if reference.editor.is_dirty()
        )
    }

    pub fn save_pptx_to(&self, path: &Path) -> Result<()> {
        let ReferenceDocument::Pptx(reference) = &self.reference else {
            bail!("only PPTX decks use the PPTX save path");
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
            ReferenceDocument::Pptx(_) => "pptx",
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

    fn apply_pptx_edit(
        &mut self,
        edit: impl FnOnce(&mut PptxEditor) -> Result<Option<PptxEditChange>>,
    ) -> Result<bool> {
        let change = {
            let ReferenceDocument::Pptx(reference) = &mut self.reference else {
                return Ok(false);
            };
            edit(&mut reference.editor)?
        };
        let Some(change) = change else {
            return Ok(false);
        };
        self.pages[change.page_index] = change.page;
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
    let (editor, pages) = PptxEditor::open(bytes, max_texture_dimension_2d)
        .with_context(|| format!("open PPTX {}", path.display()))?;
    Ok(DocumentView {
        source: path.to_owned(),
        reference: ReferenceDocument::Pptx(Box::new(PptxReference { editor })),
        pages,
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

    use betteroffice_pptx::{Primitive as PptxPrimitive, Transform as PptxTransform};
    use betteroffice_xlsx::{CellValue, DrawCmd, SheetId};
    use vello::kurbo::Point;

    use super::*;
    use crate::chrome::{Alignment, ToggleState};
    use crate::editing::TextLoc;
    use crate::test_fixtures;

    static TEST_FILE_ID: AtomicU64 = AtomicU64::new(0);

    struct PptxProbe {
        hit: PptxTextHit,
        x: f32,
        y: f32,
        caret_x: f32,
        caret_y: f32,
        caret_height: f32,
        transform: Affine,
    }

    fn pptx_probe(document: &DocumentView, slide_index: usize) -> PptxProbe {
        pptx_probe_with_transform(document, slide_index, false)
    }

    fn pptx_probe_with_transform(
        document: &DocumentView,
        slide_index: usize,
        transformed: bool,
    ) -> PptxProbe {
        let ReferenceDocument::Pptx(reference) = &document.reference else {
            panic!("expected PPTX reference");
        };
        let rendered = reference.editor.rendered_slide(slide_index).unwrap();
        for primitive in rendered.display_list.primitives.iter().rev() {
            let PptxPrimitive::TextBox {
                shape_id: Some(shape_id),
                story_id: Some(story_id),
                x,
                y,
                w,
                h,
                lines,
                transform,
                ..
            } = primitive
            else {
                continue;
            };
            if (*transform != PptxTransform::default()) != transformed
                || reference.editor.story(story_id).is_err()
            {
                continue;
            }
            let affine = pptx_test_transform(*x, *y, *w, *h, *transform);
            for line in lines {
                for stop in &line.caret_stops {
                    let point = affine
                        * Point::new(f64::from(stop.x), f64::from(line.y + line.height / 2.0));
                    let Some(PptxHit::Text(hit)) =
                        reference
                            .editor
                            .hit_test(slide_index, point.x as f32, point.y as f32)
                    else {
                        continue;
                    };
                    if hit.shape_id == *shape_id
                        && hit.story_id == *story_id
                        && hit.position == stop.position
                    {
                        return PptxProbe {
                            hit,
                            x: point.x as f32,
                            y: point.y as f32,
                            caret_x: stop.x,
                            caret_y: line.y,
                            caret_height: line.height,
                            transform: affine,
                        };
                    }
                }
            }
        }
        panic!("demo PPTX has no hittable editable text stop");
    }

    fn pptx_test_transform(x: f32, y: f32, w: f32, h: f32, transform: PptxTransform) -> Affine {
        let center = Point::new(f64::from(x + w / 2.0), f64::from(y + h / 2.0));
        let flip = Affine::translate(center.to_vec2())
            * Affine::scale_non_uniform(
                if transform.flip_h { -1.0 } else { 1.0 },
                if transform.flip_v { -1.0 } else { 1.0 },
            )
            * Affine::translate(-center.to_vec2());
        Affine::rotate_about(f64::from(transform.rotation_deg).to_radians(), center) * flip
    }

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
        assert_eq!(
            document.edited_path().unwrap(),
            path.parent().unwrap().join("betteroffice-demo-edited.pptx")
        );
        let ReferenceDocument::Pptx(reference) = &document.reference else {
            panic!("expected PPTX reference");
        };
        assert_eq!(reference.editor.slide_count(), 3);
        assert_eq!(document.pages.len(), 3);
        for summary in reference.editor.summaries() {
            assert!(summary.glyph_audit.glyph_runs > 0);
            assert_eq!(summary.glyph_audit.drifted_caret_stops, 0);
            assert_eq!(summary.glyph_audit.missing_caret_stops, 0);
            assert_eq!(summary.glyph_audit.drifted_widths, 0);
        }
        assert_eq!(document.pages[0].skipped.total(), 0);
        assert_eq!(
            reference.editor.summaries()[0].structured(&document.pages[0].skipped)["primitives"]["image"]
                ["translated"],
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
        assert_eq!(reference.editor.slide_count(), 2);
        let summary = reference.editor.summaries()[0].structured(&document.pages[0].skipped);
        assert_eq!(summary["primitives"]["chart"]["translated"], 1);
        assert_eq!(summary["primitives"]["placeholder"]["skipped"], 1);
        assert!(
            reference.editor.summaries()[0]
                .glyph_audit
                .missing_caret_stops
                > 0
        );
    }

    #[test]
    fn pptx_hit_test_places_the_exact_engine_caret_and_types_into_its_story() {
        let path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        let Some(PptxHit::Text(hit)) = document.pptx_hit_test(
            probe.hit.slide_index,
            f64::from(probe.x),
            f64::from(probe.y),
        ) else {
            panic!("expected PPTX text hit");
        };
        assert_eq!(hit, probe.hit);
        assert!(document.pptx_select_hit(hit.clone()));
        let caret = document.caret_geometry().unwrap().unwrap();
        assert_eq!(caret.page_index, hit.slide_index);
        assert_eq!(caret.x.to_bits(), f64::from(probe.caret_x).to_bits());
        assert_eq!(caret.y.to_bits(), f64::from(probe.caret_y).to_bits());
        assert_eq!(
            caret.height.to_bits(),
            f64::from(probe.caret_height).to_bits()
        );
        assert_eq!(caret.transform, Affine::IDENTITY);

        assert!(document.pptx_insert_text("Native ").unwrap());
        let ReferenceDocument::Pptx(reference) = &document.reference else {
            panic!("expected PPTX reference");
        };
        let story = reference.editor.story(&hit.story_id).unwrap();
        assert!(story.plain_text().contains("Native "));
        assert_eq!(
            reference.editor.caret_position(),
            Some(hit.position + "Native ".encode_utf16().count() as u32)
        );
        let state = document.editing_state().unwrap();
        assert!(state.can_save);
        assert!(state.can_undo);
        assert!(!state.inline_enabled);
        assert!(!state.alignment_enabled);
    }

    #[test]
    fn pptx_transformed_caret_uses_the_exact_engine_stop() {
        let source =
            test_fixtures::write_pptx("pptx-transformed-caret", &test_fixtures::transformed_pptx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let probe = pptx_probe_with_transform(&document, 0, true);
        assert!(document.pptx_select_hit(probe.hit));
        let caret = document.caret_geometry().unwrap().unwrap();
        assert_eq!(caret.x.to_bits(), f64::from(probe.caret_x).to_bits());
        assert_eq!(caret.y.to_bits(), f64::from(probe.caret_y).to_bits());
        assert_eq!(
            caret.height.to_bits(),
            f64::from(probe.caret_height).to_bits()
        );
        assert_eq!(caret.transform, probe.transform);
        let actual_top = caret.transform * Point::new(caret.x, caret.y);
        let engine_top =
            probe.transform * Point::new(f64::from(probe.caret_x), f64::from(probe.caret_y));
        assert_eq!(actual_top.x.to_bits(), engine_top.x.to_bits());
        assert_eq!(actual_top.y.to_bits(), engine_top.y.to_bits());
        fs::remove_file(source).unwrap();
    }

    #[test]
    fn pptx_arrow_navigation_stops_at_text_box_edges() {
        let path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        for direction in [
            MoveDirection::Left,
            MoveDirection::Right,
            MoveDirection::Up,
            MoveDirection::Down,
        ] {
            assert!(document.pptx_select_hit(probe.hit.clone()));
            let mut stopped = false;
            for _ in 0..10_000 {
                if !document.pptx_move_caret(direction) {
                    stopped = true;
                    break;
                }
            }
            assert!(stopped);
            assert!(!document.pptx_move_caret(direction));
            let ReferenceDocument::Pptx(reference) = &document.reference else {
                panic!("expected PPTX reference");
            };
            assert_eq!(
                reference.editor.caret_story_id(),
                Some(probe.hit.story_id.as_str())
            );
        }
    }

    #[test]
    fn pptx_backspace_and_delete_remove_engine_caret_intervals() {
        let path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        let before = match &document.reference {
            ReferenceDocument::Pptx(reference) => {
                reference.editor.story(&probe.hit.story_id).unwrap()
            }
            _ => unreachable!(),
        };
        assert!(document.pptx_select_hit(probe.hit.clone()));
        assert!(document.pptx_delete(DeleteDirection::Forward).unwrap());
        let after_delete = match &document.reference {
            ReferenceDocument::Pptx(reference) => {
                reference.editor.story(&probe.hit.story_id).unwrap()
            }
            _ => unreachable!(),
        };
        assert!(after_delete.length < before.length);

        assert!(document.pptx_undo().unwrap());
        assert!(document.pptx_select_hit(probe.hit.clone()));
        assert!(document.pptx_move_caret(MoveDirection::Right));
        assert!(document.pptx_delete(DeleteDirection::Backward).unwrap());
        let after_backspace = match &document.reference {
            ReferenceDocument::Pptx(reference) => {
                reference.editor.story(&probe.hit.story_id).unwrap()
            }
            _ => unreachable!(),
        };
        assert!(after_backspace.length < before.length);
    }

    #[test]
    fn pptx_enter_splits_and_backspace_at_the_start_merges_the_paragraph() {
        let path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        let before = match &document.reference {
            ReferenceDocument::Pptx(reference) => {
                reference.editor.story(&probe.hit.story_id).unwrap()
            }
            _ => unreachable!(),
        };
        assert!(document.pptx_select_hit(probe.hit.clone()));
        assert!(document.pptx_enter().unwrap());
        let split = match &document.reference {
            ReferenceDocument::Pptx(reference) => {
                assert_eq!(
                    reference.editor.caret_position(),
                    Some(probe.hit.position + 1)
                );
                reference.editor.story(&probe.hit.story_id).unwrap()
            }
            _ => unreachable!(),
        };
        assert_eq!(split.paragraphs.len(), before.paragraphs.len() + 1);
        assert_eq!(
            split.plain_text().replace('\n', ""),
            before.plain_text().replace('\n', "")
        );

        assert!(document.pptx_delete(DeleteDirection::Backward).unwrap());
        let merged = match &document.reference {
            ReferenceDocument::Pptx(reference) => {
                reference.editor.story(&probe.hit.story_id).unwrap()
            }
            _ => unreachable!(),
        };
        assert_eq!(merged.paragraphs.len(), before.paragraphs.len());
        assert_eq!(merged.plain_text(), before.plain_text());
    }

    #[test]
    fn pptx_save_reopens_text_and_keeps_untouched_slides_byte_identical() {
        let path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let source = fs::read(&path).unwrap();
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        assert!(document.pptx_select_hit(probe.hit.clone()));
        assert!(document.pptx_insert_text("Persisted ").unwrap());

        let output = test_fixtures::write_pptx("pptx-edited-output", b"sentinel");
        document.save_pptx_to(&output).unwrap();
        let saved = fs::read(&output).unwrap();
        for slide in ["ppt/slides/slide2.xml", "ppt/slides/slide3.xml"] {
            assert_eq!(
                test_fixtures::part(&source, slide),
                test_fixtures::part(&saved, slide),
                "{slide}"
            );
        }
        assert_eq!(fs::read(&path).unwrap(), source);
        let reopened = load_document(&output, 0, 8_192).unwrap();
        let ReferenceDocument::Pptx(reference) = &reopened.reference else {
            panic!("expected PPTX reference");
        };
        assert!(
            reference
                .editor
                .story(&probe.hit.story_id)
                .unwrap()
                .plain_text()
                .contains("Persisted ")
        );
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn pptx_save_refuses_to_flatten_unchanged_run_color_metadata() {
        let source =
            test_fixtures::write_pptx("pptx-theme-linked", &test_fixtures::theme_linked_pptx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        assert!(document.pptx_select_hit(probe.hit));
        assert!(document.pptx_insert_text("Theme ").unwrap());
        let output = test_fixtures::write_pptx("pptx-theme-refused", b"sentinel");

        let error = document.save_pptx_to(&output).unwrap_err().to_string();
        assert!(error.contains("run color metadata"), "{error}");
        assert!(error.contains("unchanged text"), "{error}");
        assert_eq!(fs::read(&output).unwrap(), b"sentinel");
        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn pptx_save_round_trips_a_safe_paragraph_split() {
        let source = test_fixtures::write_pptx(
            "pptx-language-neutral",
            &test_fixtures::language_neutral_pptx(),
        );
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        let before = match &document.reference {
            ReferenceDocument::Pptx(reference) => {
                reference.editor.story(&probe.hit.story_id).unwrap()
            }
            _ => unreachable!(),
        };
        assert!(document.pptx_select_hit(probe.hit.clone()));
        assert!(document.pptx_move_caret(MoveDirection::Right));
        assert!(document.pptx_enter().unwrap());
        let output = test_fixtures::write_pptx("pptx-safe-split", b"sentinel");
        document.save_pptx_to(&output).unwrap();

        let reopened = load_document(&output, 0, 8_192).unwrap();
        let ReferenceDocument::Pptx(reference) = &reopened.reference else {
            panic!("expected PPTX reference");
        };
        let after = reference.editor.story(&probe.hit.story_id).unwrap();
        assert_eq!(after.paragraphs.len(), before.paragraphs.len() + 1);
        assert_eq!(
            after.plain_text().replace('\n', ""),
            before.plain_text().replace('\n', "")
        );
        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn pptx_save_refuses_a_paragraph_split_that_would_drop_run_language() {
        let path =
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/betteroffice-demo.pptx");
        let mut document = load_document(&path, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        assert!(document.pptx_select_hit(probe.hit));
        assert!(document.pptx_move_caret(MoveDirection::Right));
        assert!(document.pptx_enter().unwrap());
        let output = test_fixtures::write_pptx("pptx-language-refused", b"sentinel");

        let error = document.save_pptx_to(&output).unwrap_err().to_string();
        assert!(error.contains("run language"), "{error}");
        assert!(error.contains("unchanged text"), "{error}");
        assert_eq!(fs::read(&output).unwrap(), b"sentinel");
        fs::remove_file(output).unwrap();
    }

    #[test]
    fn pptx_save_refuses_a_digital_signature_before_touching_the_output() {
        let source = test_fixtures::write_pptx("pptx-signed", &test_fixtures::signed_pptx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let probe = pptx_probe(&document, 0);
        assert!(document.pptx_select_hit(probe.hit));
        assert!(document.pptx_insert_text("Signed ").unwrap());
        let output = test_fixtures::write_pptx("pptx-signed-output", b"sentinel");

        let error = document.save_pptx_to(&output).unwrap_err().to_string();
        assert!(error.contains("digital signature"), "{error}");
        assert!(error.contains("_xmlsignatures/sig1.xml"), "{error}");
        assert_eq!(fs::read(&output).unwrap(), b"sentinel");
        fs::remove_file(source).unwrap();
        fs::remove_file(output).unwrap();
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
        let first_line = document.caret_geometry().unwrap().unwrap();
        assert!(
            document
                .docx_move_selection(MoveDirection::Down, false)
                .unwrap()
        );
        let next_line = document.caret_geometry().unwrap().unwrap();
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
        let home = document.caret_geometry().unwrap().unwrap();
        assert!(
            document
                .docx_move_selection(MoveDirection::End, false)
                .unwrap()
        );
        let end = document.caret_geometry().unwrap().unwrap();
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
        assert!(document.caret_geometry().unwrap().is_some());

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
