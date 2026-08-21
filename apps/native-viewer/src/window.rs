use std::sync::Arc;
use std::time::{Duration, Instant};

use anyhow::{Context, Result, bail};
use betteroffice_xlsx::CellRef;
use docx_edit::SimpleFormat;
use vello::kurbo::{Affine, Point, Rect, Stroke};
use vello::peniko::{Color, Fill};
use vello::util::{RenderContext, RenderSurface};
use vello::{AaConfig, RenderParams, Renderer, RendererOptions, Scene};
use winit::application::ApplicationHandler;
use winit::dpi::{LogicalSize, PhysicalSize};
use winit::event::{ElementState, KeyEvent, MouseButton, MouseScrollDelta, WindowEvent};
use winit::event_loop::{ActiveEventLoop, ControlFlow, EventLoop};
use winit::keyboard::{Key, ModifiersState, NamedKey};
use winit::window::{Window, WindowId};

use crate::chrome::{
    Chrome, ChromeHit, ChromeState, STATUS_HEIGHT, TOOLBAR_HEIGHT, ToolbarCommand, ZOOM_MAX,
    ZOOM_MIN,
};
use crate::collaboration::{
    CollaborationClient, CollaborationConfig, TransportEvent, TransportEventReceiver,
    transport_event_channel,
};
use crate::collaboration_document;
use crate::document::{DocumentView, ReferenceDocument, load_document};
use crate::editing::{DeleteDirection, MoveDirection, TextLoc};
use crate::pptx_editing::PptxHit;
use crate::xlsx_editing::CellMove;
use crate::xlsx_scene::paint_cell_editor;

const MARGIN: f64 = 40.0;
const PAGE_GAP: f64 = 24.0;
const CARET_BLINK: Duration = Duration::from_millis(530);
const DOUBLE_CLICK: Duration = Duration::from_millis(500);
const DOUBLE_CLICK_DISTANCE: f64 = 5.0;
const STATUS_DURATION: Duration = Duration::from_secs(6);
const UNSAVED_EXIT_REASON: &str = "Unsaved changes: save or undo them before closing the viewer.";

#[derive(Clone, Copy, Debug)]
struct CollaborationWake;

pub fn run(
    document: DocumentView,
    context: RenderContext,
    collaboration: Option<CollaborationConfig>,
) -> Result<()> {
    for (index, page) in document.pages.iter().enumerate() {
        let label = document.scene_label(index);
        if let ReferenceDocument::Pptx(reference) = &document.reference {
            let summary = reference
                .editor
                .summaries()
                .get(index)
                .ok_or_else(|| anyhow::anyhow!("PPTX slide has no translation summary"))?;
            println!(
                "{label} PPTX summary: {}",
                serde_json::to_string(&summary.structured(&page.skipped))?
            );
            continue;
        }
        println!(
            "{label} skipped Vello {}: {} {:?}",
            document.display_item_name(),
            page.skipped.total(),
            page.skipped.counts
        );
        if !page.skipped.reasons.is_empty() {
            println!("{label} skip reasons: {:?}", page.skipped.reasons);
        }
    }
    let event_loop = EventLoop::<CollaborationWake>::with_user_event().build()?;
    let (event_sender, event_receiver) = transport_event_channel();
    let client = collaboration
        .map(|config| {
            let proxy = event_loop.create_proxy();
            let event_sender = event_sender.clone();
            CollaborationClient::start(config, move |event| {
                event_sender.try_send(event) && proxy.send_event(CollaborationWake).is_ok()
            })
        })
        .transpose()?;
    let mut app = Viewer::new(document, context, client)?;
    app.collaboration_events = Some(event_receiver);
    event_loop.run_app(&mut app)?;
    if let Some(error) = app.fatal {
        bail!(error);
    }
    Ok(())
}

struct Viewer {
    document: DocumentView,
    chrome: Chrome,
    context: RenderContext,
    window: Option<Arc<Window>>,
    surface: Option<RenderSurface<'static>>,
    renderer: Option<Renderer>,
    zoom: f64,
    scroll: f64,
    modifiers: ModifiersState,
    cursor: Option<(f64, f64)>,
    dragging: bool,
    last_click: Option<(Instant, f64, f64)>,
    caret_visible: bool,
    next_caret_blink: Instant,
    focused: bool,
    save_target: Option<std::path::PathBuf>,
    status_message: Option<StatusMessage>,
    collaboration: Option<CollaborationClient>,
    collaboration_events: Option<TransportEventReceiver>,
    fatal: Option<String>,
}

struct StatusMessage {
    text: String,
    expires_at: Instant,
}

#[derive(Debug)]
enum SaveFailure {
    Save(anyhow::Error),
    Verification(anyhow::Error),
}

impl SaveFailure {
    fn status(&self) -> String {
        match self {
            Self::Save(error) => format!("Save refused: {error:#}"),
            Self::Verification(error) => format!("Save verification failed: {error:#}"),
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum WindowErrorPath {
    Startup,
    Render,
    Selection,
    Edit,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum ErrorHandling {
    Fatal,
    Status,
    RecoverLayout,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum EditInput<'a> {
    Text(&'a str),
    Enter,
    Tab,
    Ignored,
}

fn edit_input<'a>(
    key: &Key<&str>,
    text: Option<&'a str>,
    command: bool,
    alt: bool,
) -> EditInput<'a> {
    match key {
        Key::Named(NamedKey::Enter) => EditInput::Enter,
        Key::Named(NamedKey::Tab) => EditInput::Tab,
        _ if !command && !alt => text
            .filter(|text| {
                !text.is_empty() && text.chars().all(|character| !character.is_control())
            })
            .map_or(EditInput::Ignored, EditInput::Text),
        _ => EditInput::Ignored,
    }
}

impl WindowErrorPath {
    fn handling(self) -> ErrorHandling {
        match self {
            Self::Startup => ErrorHandling::Fatal,
            Self::Render | Self::Selection => ErrorHandling::Status,
            Self::Edit => ErrorHandling::RecoverLayout,
        }
    }

    fn label(self) -> &'static str {
        match self {
            Self::Startup => "Startup failed",
            Self::Render => "Render failed",
            Self::Selection => "Selection failed",
            Self::Edit => "Edit failed",
        }
    }
}

#[derive(Clone, Copy)]
struct ViewportGeometry {
    width: f64,
    height: f64,
    zoom: f64,
    scroll: f64,
}

impl ViewportGeometry {
    fn document_height(self) -> f64 {
        (self.height - TOOLBAR_HEIGHT - STATUS_HEIGHT).max(1.0)
    }

    fn page_origin(self, pages: &[crate::scene_shared::PageScene], page_index: usize) -> Point {
        let page = &pages[page_index];
        let x = ((self.width - page.width * self.zoom) / 2.0).max(MARGIN);
        let preceding = pages
            .iter()
            .take(page_index)
            .map(|page| page.height * self.zoom + PAGE_GAP)
            .sum::<f64>();
        Point::new(x, TOOLBAR_HEIGHT + MARGIN - self.scroll + preceding)
    }

    fn document_point(
        self,
        pages: &[crate::scene_shared::PageScene],
        x: f64,
        y: f64,
        clamp: bool,
    ) -> Option<(usize, f64, f64)> {
        if !clamp && (y < TOOLBAR_HEIGHT || y >= (self.height - STATUS_HEIGHT).max(0.0)) {
            return None;
        }
        let mut best: Option<(f64, usize, f64, f64)> = None;
        for (page_index, page) in pages.iter().enumerate() {
            let origin = self.page_origin(pages, page_index);
            let page_width = page.width * self.zoom;
            let page_height = page.height * self.zoom;
            let inside = x >= origin.x
                && x <= origin.x + page_width
                && y >= origin.y
                && y <= origin.y + page_height;
            if inside {
                return Some((
                    page_index,
                    (x - origin.x) / self.zoom,
                    (y - origin.y) / self.zoom,
                ));
            }
            if clamp {
                let nearest_x = x.clamp(origin.x, origin.x + page_width);
                let nearest_y = y.clamp(origin.y, origin.y + page_height);
                let distance = (x - nearest_x).powi(2) + (y - nearest_y).powi(2);
                if best
                    .as_ref()
                    .is_none_or(|(best_distance, ..)| distance < *best_distance)
                {
                    best = Some((
                        distance,
                        page_index,
                        (nearest_x - origin.x) / self.zoom,
                        (nearest_y - origin.y) / self.zoom,
                    ));
                }
            }
        }
        best.map(|(_, page_index, local_x, local_y)| (page_index, local_x, local_y))
    }

    fn current_page(self, pages: &[crate::scene_shared::PageScene]) -> usize {
        let viewport_center = TOOLBAR_HEIGHT + self.document_height() / 2.0;
        pages
            .iter()
            .enumerate()
            .min_by(|(left_index, left), (right_index, right)| {
                let left_center =
                    self.page_origin(pages, *left_index).y + left.height * self.zoom / 2.0;
                let right_center =
                    self.page_origin(pages, *right_index).y + right.height * self.zoom / 2.0;
                (left_center - viewport_center)
                    .abs()
                    .total_cmp(&(right_center - viewport_center).abs())
            })
            .map_or(0, |(index, _)| index)
    }
}

enum PointerTarget {
    Chrome(Option<ToolbarCommand>),
    Docx(TextLoc),
    Xlsx(CellRef),
    Pptx(PptxHit),
    None,
}

fn pointer_target(
    document: &DocumentView,
    chrome: &Chrome,
    state: &ChromeState,
    geometry: ViewportGeometry,
    point: Point,
) -> Result<PointerTarget> {
    if let ChromeHit::Consumed(command) =
        chrome.hit_test(geometry.width, geometry.height, point, state)
    {
        return Ok(PointerTarget::Chrome(command));
    }
    let Some((page_index, local_x, local_y)) =
        geometry.document_point(&document.pages, point.x, point.y, false)
    else {
        return Ok(PointerTarget::None);
    };
    if let Some(loc) = document.docx_hit_test(page_index, local_x, local_y)? {
        return Ok(PointerTarget::Docx(loc));
    }
    if let Some(hit) = document.pptx_hit_test(page_index, local_x, local_y) {
        return Ok(PointerTarget::Pptx(hit));
    }
    Ok(document
        .xlsx_hit_test(page_index, local_x, local_y)
        .map_or(PointerTarget::None, PointerTarget::Xlsx))
}

impl Viewer {
    fn new(
        document: DocumentView,
        context: RenderContext,
        collaboration: Option<CollaborationClient>,
    ) -> Result<Self> {
        let save_target = Some(document.edited_path()?);
        Ok(Self {
            document,
            chrome: Chrome::new(),
            context,
            window: None,
            surface: None,
            renderer: None,
            zoom: 1.0,
            scroll: 0.0,
            modifiers: ModifiersState::empty(),
            cursor: None,
            dragging: false,
            last_click: None,
            caret_visible: true,
            next_caret_blink: Instant::now() + CARET_BLINK,
            focused: true,
            save_target,
            status_message: None,
            collaboration,
            collaboration_events: None,
            fatal: None,
        })
    }

    fn create_window(&mut self, event_loop: &ActiveEventLoop) -> Result<()> {
        if self.window.is_some() {
            return Ok(());
        }
        let window = Arc::new(
            event_loop.create_window(
                Window::default_attributes()
                    .with_title("BetterOffice Vello")
                    .with_inner_size(LogicalSize::new(1100.0, 900.0))
                    .with_min_inner_size(LogicalSize::new(520.0, 400.0)),
            )?,
        );
        let PhysicalSize { width, height } = window.inner_size();
        let surface = pollster::block_on(self.context.create_surface(
            window.clone(),
            width.max(1),
            height.max(1),
            wgpu::PresentMode::AutoVsync,
        ))?;
        let renderer = Renderer::new(
            &self.context.devices[surface.dev_id].device,
            RendererOptions::default(),
        )?;
        self.window = Some(window);
        self.surface = Some(surface);
        self.renderer = Some(renderer);
        self.update_title();
        self.request_redraw();
        Ok(())
    }

    fn render(&mut self) -> Result<()> {
        let selection_rects = self.document.docx_selection_rects().to_vec();
        let xlsx_overlay = self.document.xlsx_overlay();
        let caret = if self.focused && self.caret_visible {
            self.document.caret_geometry()?
        } else {
            None
        };
        let geometry = self.geometry().context("window geometry is unavailable")?;
        let chrome_state = self.chrome_state()?;
        let Some(window) = &self.window else {
            return Ok(());
        };
        let Some(surface) = &mut self.surface else {
            return Ok(());
        };
        let Some(renderer) = &mut self.renderer else {
            return Ok(());
        };
        let device_scale = window.scale_factor();
        let mut scene = Scene::new();
        let scale = self.zoom * device_scale;
        scene.push_clip_layer(
            Fill::NonZero,
            Affine::IDENTITY,
            &Rect::new(
                0.0,
                TOOLBAR_HEIGHT * device_scale,
                f64::from(surface.config.width),
                (geometry.height - STATUS_HEIGHT).max(TOOLBAR_HEIGHT) * device_scale,
            ),
        );
        for (page_index, page) in self.document.pages.iter().enumerate() {
            let origin = geometry.page_origin(&self.document.pages, page_index);
            let transform = Affine::translate((origin.x * device_scale, origin.y * device_scale))
                * Affine::scale(scale);
            scene.append(&page.background, Some(transform));
            for rect in selection_rects
                .iter()
                .filter(|rect| rect.page_index == page_index)
            {
                scene.fill(
                    Fill::NonZero,
                    transform,
                    Color::from_rgba8(66, 133, 244, 88),
                    None,
                    &Rect::new(rect.x, rect.y, rect.x + rect.width, rect.y + rect.height),
                );
            }
            scene.append(&page.scene, Some(transform));
            if page_index == 0
                && let Some(overlay) = &xlsx_overlay
            {
                if let Some(draft) = &overlay.draft {
                    let mut editor = Scene::new();
                    paint_cell_editor(&mut editor, overlay.rect, draft)?;
                    scene.append(&editor, Some(transform));
                }
                let rect = Rect::new(
                    f64::from(overlay.rect.x),
                    f64::from(overlay.rect.y),
                    f64::from(overlay.rect.x + overlay.rect.w),
                    f64::from(overlay.rect.y + overlay.rect.h),
                );
                scene.stroke(
                    &Stroke::new(2.0 / self.zoom),
                    transform,
                    Color::from_rgba8(33, 115, 70, 255),
                    None,
                    &rect,
                );
            }
            if let Some(caret) = caret
                .as_ref()
                .filter(|caret| caret.page_index == page_index)
            {
                let half_width = 0.75 / self.zoom;
                scene.fill(
                    Fill::NonZero,
                    transform * caret.transform,
                    Color::from_rgba8(32, 33, 36, 255),
                    None,
                    &Rect::new(
                        caret.x - half_width,
                        caret.y,
                        caret.x + half_width,
                        caret.y + caret.height,
                    ),
                );
            }
        }
        scene.pop_layer();
        self.chrome.paint(
            &mut scene,
            geometry.width,
            geometry.height,
            device_scale,
            &chrome_state,
        );
        let device = &self.context.devices[surface.dev_id].device;
        let queue = &self.context.devices[surface.dev_id].queue;
        renderer.render_to_texture(
            device,
            queue,
            &scene,
            &surface.target_view,
            &RenderParams {
                base_color: Color::from_rgba8(232, 234, 237, 255),
                width: surface.config.width,
                height: surface.config.height,
                antialiasing_method: AaConfig::Msaa16,
            },
        )?;
        let (frame, reconfigure) = match surface.surface.get_current_texture() {
            wgpu::CurrentSurfaceTexture::Success(frame) => (frame, false),
            wgpu::CurrentSurfaceTexture::Suboptimal(frame) => (frame, true),
            wgpu::CurrentSurfaceTexture::Lost
            | wgpu::CurrentSurfaceTexture::Outdated
            | wgpu::CurrentSurfaceTexture::Validation => {
                self.context.configure_surface(surface);
                window.request_redraw();
                return Ok(());
            }
            wgpu::CurrentSurfaceTexture::Timeout | wgpu::CurrentSurfaceTexture::Occluded => {
                window.request_redraw();
                return Ok(());
            }
        };
        let frame_view = frame
            .texture
            .create_view(&wgpu::TextureViewDescriptor::default());
        let mut encoder = device.create_command_encoder(&wgpu::CommandEncoderDescriptor {
            label: Some("BetterOffice Vello surface blit"),
        });
        surface
            .blitter
            .copy(device, &mut encoder, &surface.target_view, &frame_view);
        queue.submit(Some(encoder.finish()));
        frame.present();
        if reconfigure {
            self.context.configure_surface(surface);
            window.request_redraw();
        }
        Ok(())
    }

    fn resize(&mut self, size: PhysicalSize<u32>) {
        if size.width == 0 || size.height == 0 {
            return;
        }
        if let Some(surface) = &mut self.surface {
            self.context
                .resize_surface(surface, size.width, size.height);
        }
        self.clamp_scroll();
        self.request_redraw();
    }

    fn scroll_by(&mut self, delta: f64) {
        self.scroll -= delta;
        self.clamp_scroll();
        self.request_redraw();
    }

    fn zoom_by(&mut self, factor: f64) {
        let old_zoom = self.zoom;
        self.zoom = (self.zoom * factor).clamp(ZOOM_MIN, ZOOM_MAX);
        if (self.zoom - old_zoom).abs() < f64::EPSILON {
            return;
        }
        let viewport = self.viewport_height();
        let center = self.scroll + viewport / 2.0;
        self.scroll = center * (self.zoom / old_zoom) - viewport / 2.0;
        self.clamp_scroll();
        self.update_title();
        self.request_redraw();
    }

    fn reset_zoom(&mut self) {
        self.zoom = 1.0;
        self.scroll = 0.0;
        self.update_title();
        self.request_redraw();
    }

    fn geometry(&self) -> Option<ViewportGeometry> {
        let window = self.window.as_ref()?;
        let size = window.inner_size();
        let scale = window.scale_factor();
        Some(ViewportGeometry {
            width: f64::from(size.width) / scale,
            height: f64::from(size.height) / scale,
            zoom: self.zoom,
            scroll: self.scroll,
        })
    }

    fn chrome_state(&self) -> Result<ChromeState> {
        let geometry = self.geometry().context("window geometry is unavailable")?;
        let file_name = self
            .document
            .source
            .file_name()
            .and_then(|name| name.to_str())
            .unwrap_or("document")
            .to_owned();
        let page_index = geometry.current_page(&self.document.pages);
        let editing = self.document.editing_state()?;
        let persistent_message = match (
            editing.save_disabled_reason.as_deref(),
            self.collaboration.as_ref(),
        ) {
            (Some(save), Some(collaboration)) => {
                Some(format!("{save} · {}", collaboration.status_line()))
            }
            (Some(save), None) => Some(save.to_owned()),
            (None, Some(collaboration)) => Some(collaboration.status_line()),
            (None, None) => None,
        };
        let message = self
            .status_message
            .as_ref()
            .filter(|message| message.expires_at > Instant::now())
            .map(|message| message.text.clone())
            .or(persistent_message);
        Ok(ChromeState {
            editing,
            zoom: self.zoom,
            file_name,
            position: self.document.status_position(page_index),
            message,
        })
    }

    fn document_point(&self, x: f64, y: f64, clamp: bool) -> Option<(usize, f64, f64)> {
        self.geometry()?
            .document_point(&self.document.pages, x, y, clamp)
    }

    fn pointer_down(&mut self) -> Result<bool> {
        let Some((x, y)) = self.cursor else {
            return Ok(false);
        };
        let geometry = self.geometry().context("window geometry is unavailable")?;
        let state = self.chrome_state()?;
        let target = pointer_target(
            &self.document,
            &self.chrome,
            &state,
            geometry,
            Point::new(x, y),
        )?;
        let loc = match target {
            PointerTarget::Chrome(command) => {
                self.dragging = false;
                self.last_click = None;
                self.document.pptx_clear_caret();
                if let Some(command) = command {
                    self.execute_toolbar_command(command)?;
                }
                return Ok(true);
            }
            PointerTarget::Xlsx(cell) => {
                let committed = self.document.xlsx_commit(None)?;
                let selected = self.document.xlsx_select_cell(cell);
                self.dragging = false;
                self.last_click = None;
                self.update_title();
                return Ok(committed || selected);
            }
            PointerTarget::Pptx(PptxHit::Text(hit)) => {
                self.dragging = false;
                self.last_click = None;
                let selected = self.document.pptx_select_hit(hit);
                if selected {
                    self.reset_caret_blink();
                }
                return Ok(selected);
            }
            PointerTarget::Pptx(PptxHit::Other) => {
                self.dragging = false;
                self.last_click = None;
                return Ok(self.document.pptx_clear_caret());
            }
            PointerTarget::Docx(loc) => loc,
            PointerTarget::None => return Ok(self.document.pptx_clear_caret()),
        };
        let now = Instant::now();
        let word = self.last_click.is_some_and(|(last, last_x, last_y)| {
            now.duration_since(last) <= DOUBLE_CLICK
                && (x - last_x).hypot(y - last_y) <= DOUBLE_CLICK_DISTANCE
        });
        self.last_click = Some((now, x, y));
        self.dragging = true;
        let selected = self
            .document
            .docx_select_point(loc, self.modifiers.shift_key(), word)?;
        if selected {
            self.reset_caret_blink();
        }
        Ok(selected)
    }

    fn pointer_drag(&mut self) -> Result<bool> {
        if !self.dragging {
            return Ok(false);
        }
        let Some((x, y)) = self.cursor else {
            return Ok(false);
        };
        let Some((page_index, local_x, local_y)) = self.document_point(x, y, true) else {
            return Ok(false);
        };
        let Some(loc) = self.document.docx_hit_test(page_index, local_x, local_y)? else {
            return Ok(false);
        };
        let selected = self.document.docx_extend_to(loc)?;
        if selected {
            self.reset_caret_blink();
        }
        Ok(selected)
    }

    fn execute_toolbar_command(&mut self, command: ToolbarCommand) -> Result<()> {
        if command == ToolbarCommand::Save {
            let editing = self.document.editing_state()?;
            if !editing.can_save {
                self.set_status(
                    editing
                        .save_disabled_reason
                        .unwrap_or_else(|| "Save unavailable".to_owned()),
                );
                self.request_redraw();
                return Ok(());
            }
        }
        if self.document.xlsx_is_editing() {
            self.document.xlsx_commit(None)?;
        }
        let xlsx = matches!(&self.document.reference, ReferenceDocument::Xlsx(_));
        let pptx = matches!(&self.document.reference, ReferenceDocument::Pptx(_));
        let changed = match command {
            ToolbarCommand::Undo if xlsx => self.document.xlsx_undo()?,
            ToolbarCommand::Redo if xlsx => self.document.xlsx_redo()?,
            ToolbarCommand::Undo if pptx => self.document.pptx_undo()?,
            ToolbarCommand::Redo if pptx => self.document.pptx_redo()?,
            ToolbarCommand::Undo => self.document.docx_undo()?,
            ToolbarCommand::Redo => self.document.docx_redo()?,
            ToolbarCommand::Bold => self.document.docx_toggle_format(SimpleFormat::Bold)?,
            ToolbarCommand::Italic => self.document.docx_toggle_format(SimpleFormat::Italic)?,
            ToolbarCommand::Underline => {
                self.document.docx_toggle_format(SimpleFormat::Underline)?
            }
            ToolbarCommand::Align(alignment) => self.document.docx_set_alignment(alignment)?,
            ToolbarCommand::ZoomOut => {
                self.zoom_by(1.0 / 1.1);
                return Ok(());
            }
            ToolbarCommand::ResetZoom => {
                self.reset_zoom();
                return Ok(());
            }
            ToolbarCommand::ZoomIn => {
                self.zoom_by(1.1);
                return Ok(());
            }
            ToolbarCommand::Save => {
                self.save_with_status();
                return Ok(());
            }
        };
        if changed {
            self.reset_caret_blink();
            self.clamp_scroll();
            self.update_title();
        }
        self.request_redraw();
        Ok(())
    }

    fn keyboard_input(&mut self, event_loop: &ActiveEventLoop, event: &KeyEvent) -> Result<()> {
        if event.state != ElementState::Pressed {
            return Ok(());
        }
        let command = self.modifiers.super_key() || self.modifiers.control_key();
        let key = event.logical_key.as_ref();
        if matches!(key, Key::Named(NamedKey::Escape)) {
            if self.document.pptx_clear_caret() {
                self.request_redraw();
                return Ok(());
            }
            if self.document.xlsx_cancel_edit() {
                self.update_title();
                self.request_redraw();
                return Ok(());
            }
            if self.request_exit() {
                event_loop.exit();
            }
            return Ok(());
        }
        if command
            && !event.repeat
            && let Key::Character(character) = key
        {
            if character.eq_ignore_ascii_case("z") {
                self.execute_toolbar_command(if self.modifiers.shift_key() {
                    ToolbarCommand::Redo
                } else {
                    ToolbarCommand::Undo
                })?;
                return Ok(());
            }
            if character.eq_ignore_ascii_case("s") {
                self.execute_toolbar_command(ToolbarCommand::Save)?;
                return Ok(());
            }
            if matches!(character, "+" | "=") {
                self.execute_toolbar_command(ToolbarCommand::ZoomIn)?;
                return Ok(());
            }
            if character == "-" {
                self.execute_toolbar_command(ToolbarCommand::ZoomOut)?;
                return Ok(());
            }
            if character == "0" {
                self.execute_toolbar_command(ToolbarCommand::ResetZoom)?;
                return Ok(());
            }
        }
        match key {
            Key::Named(NamedKey::PageDown) => {
                self.scroll_by(-self.viewport_height() * 0.9);
                return Ok(());
            }
            Key::Named(NamedKey::PageUp) => {
                self.scroll_by(self.viewport_height() * 0.9);
                return Ok(());
            }
            _ => {}
        }
        let edit_input = edit_input(
            &key,
            event.text.as_deref(),
            command,
            self.modifiers.alt_key(),
        );
        if matches!(&self.document.reference, ReferenceDocument::Docx(_)) {
            let changed = match key {
                Key::Named(NamedKey::ArrowLeft) => self
                    .document
                    .docx_move_selection(MoveDirection::Left, self.modifiers.shift_key())?,
                Key::Named(NamedKey::ArrowRight) => self
                    .document
                    .docx_move_selection(MoveDirection::Right, self.modifiers.shift_key())?,
                Key::Named(NamedKey::ArrowUp) => self
                    .document
                    .docx_move_selection(MoveDirection::Up, self.modifiers.shift_key())?,
                Key::Named(NamedKey::ArrowDown) => self
                    .document
                    .docx_move_selection(MoveDirection::Down, self.modifiers.shift_key())?,
                Key::Named(NamedKey::Home) => self
                    .document
                    .docx_move_selection(MoveDirection::Home, self.modifiers.shift_key())?,
                Key::Named(NamedKey::End) => self
                    .document
                    .docx_move_selection(MoveDirection::End, self.modifiers.shift_key())?,
                Key::Named(NamedKey::Backspace) => {
                    self.document.docx_delete(DeleteDirection::Backward)?
                }
                Key::Named(NamedKey::Delete) => {
                    self.document.docx_delete(DeleteDirection::Forward)?
                }
                _ => match edit_input {
                    EditInput::Text(text) => self.document.docx_insert_text(text)?,
                    EditInput::Enter => self.document.docx_enter()?,
                    EditInput::Tab | EditInput::Ignored => false,
                },
            };
            if changed {
                self.reset_caret_blink();
                self.update_title();
                self.request_redraw();
            }
            return Ok(());
        }
        if matches!(&self.document.reference, ReferenceDocument::Xlsx(_)) {
            let editing = self.document.xlsx_is_editing();
            let changed = if editing {
                match key {
                    Key::Named(NamedKey::Backspace) => self.document.xlsx_backspace(),
                    _ => match edit_input {
                        EditInput::Text(text) => self.document.xlsx_insert_text(text),
                        EditInput::Enter => self.document.xlsx_commit(Some(CellMove::Down))?,
                        EditInput::Tab => self.document.xlsx_commit(Some(CellMove::Right))?,
                        EditInput::Ignored => false,
                    },
                }
            } else {
                match key {
                    Key::Named(NamedKey::ArrowLeft) => {
                        self.document.xlsx_move_selection(CellMove::Left)
                    }
                    Key::Named(NamedKey::ArrowRight) => {
                        self.document.xlsx_move_selection(CellMove::Right)
                    }
                    Key::Named(NamedKey::ArrowUp) => {
                        self.document.xlsx_move_selection(CellMove::Up)
                    }
                    Key::Named(NamedKey::ArrowDown) => {
                        self.document.xlsx_move_selection(CellMove::Down)
                    }
                    Key::Named(NamedKey::F2) => self.document.xlsx_begin_edit(None)?,
                    _ => match edit_input {
                        EditInput::Text(text) => self.document.xlsx_begin_edit(Some(text))?,
                        EditInput::Enter => self.document.xlsx_begin_edit(None)?,
                        EditInput::Tab => self.document.xlsx_move_selection(CellMove::Right),
                        EditInput::Ignored => false,
                    },
                }
            };
            if changed {
                self.update_title();
                self.request_redraw();
            }
            return Ok(());
        }
        if matches!(&self.document.reference, ReferenceDocument::Pptx(_))
            && self.document.has_text_caret()
        {
            let changed = match key {
                Key::Named(NamedKey::ArrowLeft) => {
                    self.document.pptx_move_caret(MoveDirection::Left)
                }
                Key::Named(NamedKey::ArrowRight) => {
                    self.document.pptx_move_caret(MoveDirection::Right)
                }
                Key::Named(NamedKey::ArrowUp) => self.document.pptx_move_caret(MoveDirection::Up),
                Key::Named(NamedKey::ArrowDown) => {
                    self.document.pptx_move_caret(MoveDirection::Down)
                }
                Key::Named(NamedKey::Home) => self.document.pptx_move_caret(MoveDirection::Home),
                Key::Named(NamedKey::End) => self.document.pptx_move_caret(MoveDirection::End),
                Key::Named(NamedKey::Backspace) => {
                    self.document.pptx_delete(DeleteDirection::Backward)?
                }
                Key::Named(NamedKey::Delete) => {
                    self.document.pptx_delete(DeleteDirection::Forward)?
                }
                _ => match edit_input {
                    EditInput::Text(text) => self.document.pptx_insert_text(text)?,
                    EditInput::Enter => self.document.pptx_enter()?,
                    EditInput::Tab | EditInput::Ignored => false,
                },
            };
            if changed {
                self.reset_caret_blink();
                self.update_title();
                self.request_redraw();
            }
            return Ok(());
        }
        if event.repeat {
            return Ok(());
        }
        match key {
            Key::Character("+" | "=") => self.zoom_by(1.1),
            Key::Character("-") => self.zoom_by(1.0 / 1.1),
            Key::Character("0") => self.reset_zoom(),
            Key::Named(NamedKey::Home) => {
                self.scroll = 0.0;
                self.request_redraw();
            }
            Key::Named(NamedKey::End) => {
                self.scroll = self.content_height();
                self.clamp_scroll();
                self.request_redraw();
            }
            _ => {}
        }
        Ok(())
    }

    fn save_with_status(&mut self) {
        match self.save_and_verify() {
            Ok(path) => {
                self.update_title();
                self.set_status(format!("Saved {}", path.display()));
            }
            Err(failure) => self.set_status(failure.status()),
        }
        self.request_redraw();
    }

    fn save_and_verify(&mut self) -> std::result::Result<std::path::PathBuf, SaveFailure> {
        if self.document.xlsx_is_editing() {
            self.document.xlsx_commit(None).map_err(SaveFailure::Save)?;
        }
        let (format, sheet_index) = match &self.document.reference {
            ReferenceDocument::Docx(_) => ("DOCX", 0),
            ReferenceDocument::Xlsx(reference) => ("XLSX", reference.sheet_index),
            ReferenceDocument::Pptx(_) => ("PPTX", 0),
        };
        let path = self
            .save_target
            .clone()
            .context("save target is unavailable")
            .map_err(SaveFailure::Save)?;
        match &self.document.reference {
            ReferenceDocument::Docx(_) => self.document.save_docx_to(&path),
            ReferenceDocument::Xlsx(_) => self.document.save_xlsx_to(&path),
            ReferenceDocument::Pptx(_) => self.document.save_pptx_to(&path),
        }
        .map_err(SaveFailure::Save)?;
        load_document(&path, sheet_index, self.document.max_texture_dimension_2d)
            .with_context(|| format!("reopen edited {format} {}", path.display()))
            .map_err(SaveFailure::Verification)?;
        println!("saved and verified edited {format}: {}", path.display());
        Ok(path)
    }

    fn request_exit(&mut self) -> bool {
        if !self.document.is_dirty() || self.document.is_remote_only_dirty() {
            return true;
        }
        self.set_status(UNSAVED_EXIT_REASON.to_owned());
        self.request_redraw();
        false
    }

    fn set_status(&mut self, text: String) {
        self.status_message = Some(StatusMessage {
            text,
            expires_at: Instant::now() + STATUS_DURATION,
        });
    }

    fn handle_collaboration_event(&mut self, event: TransportEvent) -> Result<()> {
        let repainted = match &mut self.collaboration {
            Some(collaboration) => collaboration_document::handle_transport_event(
                &mut self.document,
                collaboration,
                event,
            )?,
            None => false,
        };
        if repainted {
            self.refresh_after_collaboration_update();
        }
        self.request_redraw();
        Ok(())
    }

    #[cfg(test)]
    fn apply_collaboration_frame(&mut self, frame: &[u8]) -> Result<bool> {
        let Some(collaboration) = &mut self.collaboration else {
            return Ok(false);
        };
        let repainted =
            collaboration_document::apply_frame(&mut self.document, collaboration, frame)?;
        if repainted {
            self.refresh_after_collaboration_update();
        }
        Ok(repainted)
    }

    fn forward_local_updates(&mut self) {
        let Some(collaboration) = &mut self.collaboration else {
            return;
        };
        collaboration_document::forward_local_updates(&self.document, collaboration);
    }

    fn refresh_after_collaboration_update(&mut self) {
        self.reset_caret_blink();
        self.clamp_scroll();
        self.update_title();
    }

    fn reset_caret_blink(&mut self) {
        self.caret_visible = true;
        self.next_caret_blink = Instant::now() + CARET_BLINK;
    }

    fn content_height(&self) -> f64 {
        let page_height = self
            .document
            .pages
            .iter()
            .map(|page| page.height * self.zoom)
            .sum::<f64>();
        let gaps = PAGE_GAP * self.document.pages.len().saturating_sub(1) as f64;
        MARGIN * 2.0 + page_height + gaps
    }

    fn viewport_height(&self) -> f64 {
        self.geometry()
            .map_or(1.0, ViewportGeometry::document_height)
    }

    fn clamp_scroll(&mut self) {
        let maximum = (self.content_height() - self.viewport_height()).max(0.0);
        self.scroll = self.scroll.clamp(0.0, maximum);
    }

    fn request_redraw(&self) {
        if let Some(window) = &self.window {
            window.request_redraw();
        }
    }

    fn update_title(&self) {
        let Some(window) = &self.window else {
            return;
        };
        let name = self
            .document
            .source
            .file_name()
            .and_then(|name| name.to_str())
            .unwrap_or("document");
        let skipped = self
            .document
            .pages
            .iter()
            .map(|page| page.skipped.total())
            .sum::<usize>();
        let edit_state = match &self.document.reference {
            ReferenceDocument::Docx(reference) if reference.editor.is_dirty() => " — edited",
            ReferenceDocument::Docx(_) => " — editable",
            ReferenceDocument::Xlsx(_) if self.document.xlsx_is_dirty() => " — edited",
            ReferenceDocument::Xlsx(_) => " — editable",
            ReferenceDocument::Pptx(_) if self.document.pptx_is_dirty() => " — edited",
            ReferenceDocument::Pptx(_) => " — editable",
        };
        window.set_title(&format!(
            "{name} — {} — {:.0}% — {skipped} skipped{edit_state}",
            self.document.title_summary(),
            self.zoom * 100.0
        ));
    }

    fn handle_error(
        &mut self,
        event_loop: &ActiveEventLoop,
        path: WindowErrorPath,
        error: anyhow::Error,
    ) {
        let label = path.label();
        match path.handling() {
            ErrorHandling::Fatal => {
                self.fatal = Some(format!("{error:#}"));
                event_loop.exit();
            }
            ErrorHandling::Status => {
                self.set_status(format!("{label}: {error:#}"));
            }
            ErrorHandling::RecoverLayout => {
                let message = match self.document.recover_after_edit_error() {
                    Ok(()) => format!("{label}: {error:#}"),
                    Err(recovery) => {
                        format!("{label}: {error:#}; layout recovery failed: {recovery:#}")
                    }
                };
                self.dragging = false;
                self.clamp_scroll();
                self.update_title();
                self.set_status(message);
            }
        }
        if !matches!(path, WindowErrorPath::Render) {
            self.request_redraw();
        }
    }
}

impl ApplicationHandler<CollaborationWake> for Viewer {
    fn resumed(&mut self, event_loop: &ActiveEventLoop) {
        if let Err(error) = self.create_window(event_loop) {
            self.handle_error(event_loop, WindowErrorPath::Startup, error);
        }
    }

    fn window_event(
        &mut self,
        event_loop: &ActiveEventLoop,
        window_id: WindowId,
        event: WindowEvent,
    ) {
        if self
            .window
            .as_ref()
            .is_none_or(|window| window.id() != window_id)
        {
            return;
        }
        match event {
            WindowEvent::CloseRequested => {
                if self.request_exit() {
                    event_loop.exit();
                }
            }
            WindowEvent::Resized(size) => self.resize(size),
            WindowEvent::ScaleFactorChanged { .. } => {
                if let Some(window) = &self.window {
                    self.resize(window.inner_size());
                }
            }
            WindowEvent::RedrawRequested => {
                if let Err(error) = self.render() {
                    self.handle_error(event_loop, WindowErrorPath::Render, error);
                }
            }
            WindowEvent::ModifiersChanged(modifiers) => self.modifiers = modifiers.state(),
            WindowEvent::Focused(focused) => {
                self.focused = focused;
                self.dragging = false;
                self.reset_caret_blink();
                self.request_redraw();
            }
            WindowEvent::CursorMoved { position, .. } => {
                let scale = self
                    .window
                    .as_ref()
                    .map(|window| window.scale_factor())
                    .unwrap_or(1.0);
                self.cursor = Some((position.x / scale, position.y / scale));
                match self.pointer_drag() {
                    Ok(true) => self.request_redraw(),
                    Ok(false) => {}
                    Err(error) => self.handle_error(event_loop, WindowErrorPath::Selection, error),
                }
            }
            WindowEvent::MouseInput {
                state,
                button: MouseButton::Left,
                ..
            } => match state {
                ElementState::Pressed => match self.pointer_down() {
                    Ok(true) => self.request_redraw(),
                    Ok(false) => {}
                    Err(error) => self.handle_error(event_loop, WindowErrorPath::Edit, error),
                },
                ElementState::Released => self.dragging = false,
            },
            WindowEvent::MouseWheel { delta, .. } => {
                let delta = match delta {
                    MouseScrollDelta::LineDelta(_, y) => f64::from(y) * 48.0,
                    MouseScrollDelta::PixelDelta(position) => {
                        let scale = self
                            .window
                            .as_ref()
                            .map(|window| window.scale_factor())
                            .unwrap_or(1.0);
                        position.y / scale
                    }
                };
                if self.modifiers.super_key() || self.modifiers.control_key() {
                    self.zoom_by((delta * 0.002).exp());
                } else {
                    self.scroll_by(delta);
                }
            }
            WindowEvent::KeyboardInput { event, .. } => {
                if let Err(error) = self.keyboard_input(event_loop, &event) {
                    self.handle_error(event_loop, WindowErrorPath::Edit, error);
                }
            }
            _ => {}
        }
    }

    fn user_event(&mut self, event_loop: &ActiveEventLoop, _event: CollaborationWake) {
        let Some(event) = self
            .collaboration_events
            .as_ref()
            .and_then(TransportEventReceiver::try_recv)
        else {
            return;
        };
        if let Err(error) = self.handle_collaboration_event(event) {
            self.handle_error(event_loop, WindowErrorPath::Edit, error);
        }
    }

    fn about_to_wait(&mut self, event_loop: &ActiveEventLoop) {
        self.forward_local_updates();
        let now = Instant::now();
        if self
            .status_message
            .as_ref()
            .is_some_and(|message| message.expires_at <= now)
        {
            self.status_message = None;
            self.request_redraw();
        }
        let mut deadline = self
            .status_message
            .as_ref()
            .map(|message| message.expires_at);
        if self.focused && self.document.has_text_caret() {
            if now >= self.next_caret_blink {
                self.caret_visible = !self.caret_visible;
                self.next_caret_blink = now + CARET_BLINK;
                self.request_redraw();
            }
            deadline = Some(deadline.map_or(self.next_caret_blink, |current| {
                current.min(self.next_caret_blink)
            }));
        } else {
            self.caret_visible = false;
        }
        if let Some(deadline) = deadline {
            event_loop.set_control_flow(ControlFlow::WaitUntil(deadline));
        } else {
            event_loop.set_control_flow(ControlFlow::Wait);
        }
    }
}

#[cfg(test)]
mod tests {
    use std::path::Path;

    use super::*;
    use crate::collaboration::TransportCommand;
    use crate::collaboration_protocol::{
        ProtocolMessage, decode_messages, encode_update, encode_update_with_fingerprint,
    };
    use crate::document::load_collaborative_docx;
    use crate::editing::REMOTE_STRUCTURAL_SAVE_REASON;
    use crate::test_fixtures;

    fn next_sent_frame(commands: &mut tokio::sync::mpsc::Receiver<TransportCommand>) -> Vec<u8> {
        match commands.try_recv().unwrap() {
            TransportCommand::Send(frame) => frame,
            TransportCommand::Shutdown => panic!("collaboration client shut down"),
        }
    }

    #[test]
    fn only_startup_errors_are_fatal() {
        assert_eq!(WindowErrorPath::Startup.handling(), ErrorHandling::Fatal);
        assert_eq!(WindowErrorPath::Render.handling(), ErrorHandling::Status);
        assert_eq!(WindowErrorPath::Selection.handling(), ErrorHandling::Status);
        assert_eq!(
            WindowErrorPath::Edit.handling(),
            ErrorHandling::RecoverLayout
        );
    }

    #[test]
    fn named_space_uses_event_text_while_enter_and_tab_stay_commands() {
        let space: Key<&str> = Key::Named(NamedKey::Space);
        let enter: Key<&str> = Key::Named(NamedKey::Enter);
        let tab: Key<&str> = Key::Named(NamedKey::Tab);
        assert_eq!(
            edit_input(&space, Some(" "), false, false),
            EditInput::Text(" ")
        );
        assert_eq!(
            edit_input(&space, Some(" "), true, false),
            EditInput::Ignored
        );
        assert_eq!(
            edit_input(&enter, Some("\r"), false, false),
            EditInput::Enter
        );
        assert_eq!(edit_input(&tab, Some("\t"), false, false), EditInput::Tab);
    }

    #[test]
    fn collaboration_wiring_broadcasts_repaints_and_retains_offline_edits() {
        let bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:r><w:t>Shared</w:t></w:r></w:p>"#,
        );
        let source = test_fixtures::write_docx("window-collaboration", &bytes);
        let mut document = load_collaborative_docx(&source, 8_192, 101).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 6,
                },
                false,
                false,
            )
            .unwrap();
        let config = CollaborationConfig::new("wiring".to_owned(), "ws://127.0.0.1:9").unwrap();
        let (client, mut commands) = CollaborationClient::detached(config);
        let mut viewer = Viewer::new(document, RenderContext::new(), Some(client)).unwrap();

        viewer
            .handle_collaboration_event(TransportEvent::Connected)
            .unwrap();
        assert!(matches!(
            decode_messages(&next_sent_frame(&mut commands))
                .unwrap()
                .as_slice(),
            [ProtocolMessage::SyncStep1(_)]
        ));
        assert!(viewer.document.docx_insert_text(" local").unwrap());
        viewer.forward_local_updates();
        let local_frame = next_sent_frame(&mut commands);
        assert!(matches!(
            decode_messages(&local_frame).unwrap().as_slice(),
            [ProtocolMessage::Update(_)]
        ));

        let mut peer = load_collaborative_docx(&source, 8_192, 202).unwrap();
        let local_messages = decode_messages(&local_frame).unwrap();
        let [ProtocolMessage::Update(update)] = local_messages.as_slice() else {
            unreachable!();
        };
        peer.docx_apply_remote_update(update).unwrap();
        peer.docx_select_point(
            TextLoc {
                para_id: "11111111".to_owned(),
                offset: 0,
            },
            false,
            false,
        )
        .unwrap();
        assert!(peer.docx_insert_text("remote ").unwrap());
        let before_checksum = viewer.document.docx_canonical_checksum().unwrap();
        let before_glyphs = viewer.document.pages[0]
            .scene
            .encoding()
            .resources
            .glyphs
            .len();
        let mut repainted = false;
        for update in peer.docx_drain_local_updates() {
            repainted |= viewer
                .apply_collaboration_frame(&encode_update(&update).unwrap())
                .unwrap();
        }
        assert!(repainted);
        assert_ne!(
            viewer.document.docx_canonical_checksum().unwrap(),
            before_checksum
        );
        assert!(
            viewer.document.pages[0]
                .scene
                .encoding()
                .resources
                .glyphs
                .len()
                > before_glyphs
        );

        viewer
            .handle_collaboration_event(TransportEvent::Reconnecting {
                delay: Duration::from_millis(250),
                reason: "relay unavailable".to_owned(),
            })
            .unwrap();
        let offline_status = viewer.collaboration.as_ref().unwrap().status_line();
        assert!(offline_status.contains("offline, reconnecting"));
        assert!(offline_status.contains("relay unavailable"));
        assert!(viewer.document.docx_insert_text(" offline").unwrap());
        let offline_checksum = viewer.document.docx_canonical_checksum().unwrap();
        viewer.forward_local_updates();
        assert!(commands.try_recv().is_err());
        assert!(viewer.fatal.is_none());

        viewer
            .handle_collaboration_event(TransportEvent::Connected)
            .unwrap();
        assert!(matches!(
            decode_messages(&next_sent_frame(&mut commands))
                .unwrap()
                .as_slice(),
            [ProtocolMessage::SyncStep1(_)]
        ));
        viewer.forward_local_updates();
        assert!(matches!(
            decode_messages(&next_sent_frame(&mut commands))
                .unwrap()
                .as_slice(),
            [ProtocolMessage::Update(_)]
        ));
        assert_eq!(
            viewer.document.docx_canonical_checksum().unwrap(),
            offline_checksum
        );
        std::fs::remove_file(source).unwrap();
    }

    #[test]
    fn structural_save_command_is_disabled_and_dirty_exit_is_blocked() {
        let bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:r><w:t>Plain</w:t></w:r></w:p>"#,
        );
        let source = test_fixtures::write_docx("window-structural-save", &bytes);
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
        assert!(document.docx_enter().unwrap());
        let output = test_fixtures::write_docx("window-structural-output", b"sentinel");
        let mut viewer = Viewer::new(document, RenderContext::new(), None).unwrap();
        viewer.save_target = Some(output.clone());

        viewer
            .execute_toolbar_command(ToolbarCommand::Save)
            .unwrap();
        assert_eq!(
            viewer
                .status_message
                .as_ref()
                .map(|message| message.text.as_str()),
            Some(crate::editing::STRUCTURAL_SAVE_REASON)
        );
        assert_eq!(std::fs::read(&output).unwrap(), b"sentinel");
        assert!(!viewer.request_exit());
        assert_eq!(
            viewer
                .status_message
                .as_ref()
                .map(|message| message.text.as_str()),
            Some(UNSAVED_EXIT_REASON)
        );
        assert!(viewer.document.docx_undo().unwrap());
        assert!(viewer.request_exit());
        std::fs::remove_file(source).unwrap();
        std::fs::remove_file(output).unwrap();
    }

    #[test]
    fn remote_structural_dirty_state_does_not_block_exit() {
        let bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:r><w:t>Plain</w:t></w:r></w:p>"#,
        );
        let source = test_fixtures::write_docx("remote-structural-exit", &bytes);
        let local = load_collaborative_docx(&source, 8_192, 101).unwrap();
        let mut peer = load_collaborative_docx(&source, 8_192, 202).unwrap();
        peer.docx_select_point(
            TextLoc {
                para_id: "11111111".to_owned(),
                offset: 2,
            },
            false,
            false,
        )
        .unwrap();
        assert!(peer.docx_enter().unwrap());
        let fingerprint = local.docx_fingerprint().unwrap();
        let frame = peer
            .docx_drain_local_updates()
            .into_iter()
            .flat_map(|update| encode_update_with_fingerprint(&update, &fingerprint).unwrap())
            .collect::<Vec<_>>();
        let (client, _commands) =
            CollaborationClient::detached(CollaborationConfig::for_test("remote-structural", 101));
        let mut viewer = Viewer::new(local, RenderContext::new(), Some(client)).unwrap();

        assert!(viewer.apply_collaboration_frame(&frame).unwrap());
        assert!(viewer.document.is_dirty());
        assert!(viewer.document.is_remote_only_dirty());
        assert!(!viewer.document.docx_undo().unwrap());
        let editing = viewer.document.editing_state().unwrap();
        assert!(!editing.can_save);
        assert_eq!(
            editing.save_disabled_reason.as_deref(),
            Some(REMOTE_STRUCTURAL_SAVE_REASON)
        );
        viewer
            .execute_toolbar_command(ToolbarCommand::Save)
            .unwrap();
        assert_eq!(
            viewer
                .status_message
                .as_ref()
                .map(|message| message.text.as_str()),
            Some(REMOTE_STRUCTURAL_SAVE_REASON)
        );
        assert!(viewer.request_exit());
        std::fs::remove_file(source).unwrap();
    }

    #[test]
    fn verified_save_keeps_the_live_document_caret_history_and_scroll() {
        let bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:r><w:t>Plain</w:t></w:r></w:p>"#,
        );
        let source = test_fixtures::write_docx("window-save-state", &bytes);
        let mut document = load_document(&source, 0, 8_192).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 5,
                },
                false,
                false,
            )
            .unwrap();
        assert!(document.docx_insert_text("!").unwrap());
        let selection = match &document.reference {
            ReferenceDocument::Docx(reference) => reference.editor.selection_range().unwrap(),
            _ => unreachable!(),
        };
        let output = test_fixtures::write_docx("window-save-state-output", b"sentinel");
        let mut viewer = Viewer::new(document, RenderContext::new(), None).unwrap();
        viewer.save_target = Some(output.clone());
        viewer.scroll = 73.0;

        assert_eq!(viewer.save_and_verify().unwrap(), output);
        assert_eq!(viewer.document.source, source);
        assert_eq!(viewer.scroll, 73.0);
        assert!(viewer.document.has_text_caret());
        let after = match &viewer.document.reference {
            ReferenceDocument::Docx(reference) => {
                assert!(!reference.editor.is_dirty());
                assert!(reference.editor.editing_state().unwrap().can_undo);
                reference.editor.selection_range().unwrap()
            }
            _ => unreachable!(),
        };
        assert_eq!(after, selection);
        assert!(viewer.document.docx_undo().unwrap());
        assert!(viewer.document.is_dirty());
        std::fs::remove_file(source).unwrap();
        std::fs::remove_file(output).unwrap();
    }

    #[test]
    fn failed_reopen_is_reported_as_save_verification_failure() {
        let bytes = test_fixtures::editing_docx(
            r#"<w:p w14:paraId="11111111"><w:r><w:t>Plain</w:t></w:r></w:p>"#,
        );
        let source = test_fixtures::write_docx("window-save-verification", &bytes);
        let mut document = load_document(&source, 0, 8_192).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "11111111".to_owned(),
                    offset: 5,
                },
                false,
                false,
            )
            .unwrap();
        assert!(document.docx_insert_text("!").unwrap());
        let output = test_fixtures::write_docx("window-save-verification-output", b"sentinel");
        let invalid_output = output.with_extension("invalid");
        std::fs::rename(&output, &invalid_output).unwrap();
        let mut viewer = Viewer::new(document, RenderContext::new(), None).unwrap();
        viewer.save_target = Some(invalid_output.clone());

        viewer.save_with_status();
        let message = &viewer.status_message.as_ref().unwrap().text;
        assert!(
            message.starts_with("Save verification failed:"),
            "{message}"
        );
        assert!(!message.starts_with("Save refused:"), "{message}");
        assert_ne!(std::fs::read(&invalid_output).unwrap(), b"sentinel");
        assert_eq!(viewer.document.source, source);
        assert!(viewer.document.has_text_caret());
        std::fs::remove_file(source).unwrap();
        std::fs::remove_file(invalid_output).unwrap();
    }

    #[test]
    fn toolbar_hit_is_consumed_and_document_hit_moves_the_caret() {
        let source = test_fixtures::write_docx("chrome-hit", &test_fixtures::complex_docx());
        let mut document = load_document(&source, 0, 8_192).unwrap();
        document
            .docx_select_point(
                TextLoc {
                    para_id: "22222222".to_owned(),
                    offset: 0,
                },
                false,
                false,
            )
            .unwrap();
        let before = match &document.reference {
            ReferenceDocument::Docx(reference) => reference.editor.selection_range().unwrap(),
            _ => unreachable!(),
        };
        let geometry = ViewportGeometry {
            width: 1_100.0,
            height: 900.0,
            zoom: 1.0,
            scroll: 0.0,
        };
        let chrome = Chrome::new();
        let state = ChromeState {
            editing: document.editing_state().unwrap(),
            zoom: 1.0,
            file_name: "fixture.docx".to_owned(),
            position: "Page 1 of 1".to_owned(),
            message: None,
        };
        assert!(matches!(
            pointer_target(&document, &chrome, &state, geometry, Point::new(4.0, 4.0)).unwrap(),
            PointerTarget::Chrome(None)
        ));
        let unchanged = match &document.reference {
            ReferenceDocument::Docx(reference) => reference.editor.selection_range().unwrap(),
            _ => unreachable!(),
        };
        assert_eq!(unchanged, before);

        let caret = match &document.reference {
            ReferenceDocument::Docx(reference) => reference
                .editor
                .engine()
                .resident_caret_snapshot(Some(("22222222", 8)))
                .unwrap()
                .caret_rect
                .unwrap(),
            _ => unreachable!(),
        };
        let origin = geometry.page_origin(&document.pages, caret.page_index);
        let point = Point::new(origin.x + caret.x, origin.y + caret.y + caret.height / 2.0);
        let PointerTarget::Docx(loc) =
            pointer_target(&document, &chrome, &state, geometry, point).unwrap()
        else {
            panic!("expected document pointer target");
        };
        assert_eq!(loc.offset, 8);
        document.docx_select_point(loc, false, false).unwrap();
        let after = match &document.reference {
            ReferenceDocument::Docx(reference) => reference.editor.selection_range().unwrap(),
            _ => unreachable!(),
        };
        assert_ne!(after, before);
        std::fs::remove_file(source).unwrap();
    }

    #[test]
    fn worksheet_hit_selects_the_cell_under_the_pointer() {
        let source = Path::new(env!("CARGO_MANIFEST_DIR")).join("../demo/public/showcase.xlsx");
        let mut document = load_document(&source, 0, 8_192).unwrap();
        let expected = CellRef::parse_a1("B5").unwrap();
        document.xlsx_select_cell(expected);
        let rect = document.xlsx_overlay().unwrap().rect;
        document.xlsx_select_cell(CellRef::new(0, 0));
        let geometry = ViewportGeometry {
            width: 1_100.0,
            height: 900.0,
            zoom: 1.0,
            scroll: 0.0,
        };
        let origin = geometry.page_origin(&document.pages, 0);
        let point = Point::new(
            origin.x + f64::from(rect.x + rect.w / 2.0),
            origin.y + f64::from(rect.y + rect.h / 2.0),
        );
        let state = ChromeState {
            editing: document.editing_state().unwrap(),
            zoom: 1.0,
            file_name: "showcase.xlsx".to_owned(),
            position: document.status_position(0),
            message: None,
        };
        let PointerTarget::Xlsx(cell) =
            pointer_target(&document, &Chrome::new(), &state, geometry, point).unwrap()
        else {
            panic!("expected worksheet pointer target");
        };
        assert_eq!(cell, expected);
        assert!(document.xlsx_select_cell(cell));
        assert!(document.status_position(0).contains("B5 ="));
    }
}
