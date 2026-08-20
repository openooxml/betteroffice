use std::sync::Arc;
use std::time::{Duration, Instant};

use anyhow::{Context, Result, bail};
use vello::kurbo::{Affine, Rect};
use vello::peniko::{Color, Fill};
use vello::util::{RenderContext, RenderSurface};
use vello::{AaConfig, RenderParams, Renderer, RendererOptions, Scene};
use winit::application::ApplicationHandler;
use winit::dpi::{LogicalSize, PhysicalSize};
use winit::event::{ElementState, KeyEvent, MouseButton, MouseScrollDelta, WindowEvent};
use winit::event_loop::{ActiveEventLoop, ControlFlow, EventLoop};
use winit::keyboard::{Key, ModifiersState, NamedKey};
use winit::window::{Window, WindowId};

use crate::document::{DocumentView, ReferenceDocument, load_document};
use crate::editing::{DeleteDirection, MoveDirection};

const MARGIN: f64 = 40.0;
const PAGE_GAP: f64 = 24.0;
const CARET_BLINK: Duration = Duration::from_millis(530);
const DOUBLE_CLICK: Duration = Duration::from_millis(500);
const DOUBLE_CLICK_DISTANCE: f64 = 5.0;

pub fn run(document: DocumentView, context: RenderContext) -> Result<()> {
    for (index, page) in document.pages.iter().enumerate() {
        let label = document.scene_label(index);
        if let ReferenceDocument::Pptx(reference) = &document.reference {
            let summary = reference
                .summaries
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
    let event_loop = EventLoop::new()?;
    let mut app = Viewer::new(document, context)?;
    event_loop.run_app(&mut app)?;
    if let Some(error) = app.fatal {
        bail!(error);
    }
    Ok(())
}

struct Viewer {
    document: DocumentView,
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
    fatal: Option<String>,
}

impl Viewer {
    fn new(document: DocumentView, context: RenderContext) -> Result<Self> {
        let save_target = matches!(&document.reference, ReferenceDocument::Docx(_))
            .then(|| document.edited_path())
            .transpose()?;
        Ok(Self {
            document,
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
                    .with_inner_size(LogicalSize::new(1100.0, 900.0)),
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
        let caret = if self.focused && self.caret_visible {
            self.document.docx_caret_geometry()?
        } else {
            None
        };
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
        let mut y = MARGIN * device_scale - self.scroll * device_scale;
        let scale = self.zoom * device_scale;
        for (page_index, page) in self.document.pages.iter().enumerate() {
            let x = ((f64::from(surface.config.width) - page.width * scale) / 2.0)
                .max(MARGIN * device_scale);
            let transform = Affine::translate((x, y)) * Affine::scale(scale);
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
            if let Some(caret) = caret
                .as_ref()
                .filter(|caret| caret.page_index == page_index)
            {
                let half_width = 0.75 / self.zoom;
                scene.fill(
                    Fill::NonZero,
                    transform,
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
            y += (page.height * self.zoom + PAGE_GAP) * device_scale;
        }
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
        self.zoom = (self.zoom * factor).clamp(0.25, 5.0);
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

    fn document_point(&self, x: f64, y: f64, clamp: bool) -> Option<(usize, f64, f64)> {
        let window = self.window.as_ref()?;
        let viewport_width = f64::from(window.inner_size().width) / window.scale_factor();
        let mut page_top = MARGIN - self.scroll;
        let mut best: Option<(f64, usize, f64, f64)> = None;
        for (page_index, page) in self.document.pages.iter().enumerate() {
            let page_width = page.width * self.zoom;
            let page_height = page.height * self.zoom;
            let page_left = ((viewport_width - page_width) / 2.0).max(MARGIN);
            let inside = x >= page_left
                && x <= page_left + page_width
                && y >= page_top
                && y <= page_top + page_height;
            if inside {
                return Some((
                    page_index,
                    (x - page_left) / self.zoom,
                    (y - page_top) / self.zoom,
                ));
            }
            if clamp {
                let nearest_x = x.clamp(page_left, page_left + page_width);
                let nearest_y = y.clamp(page_top, page_top + page_height);
                let distance = (x - nearest_x).powi(2) + (y - nearest_y).powi(2);
                if best
                    .as_ref()
                    .is_none_or(|(best_distance, ..)| distance < *best_distance)
                {
                    best = Some((
                        distance,
                        page_index,
                        (nearest_x - page_left) / self.zoom,
                        (nearest_y - page_top) / self.zoom,
                    ));
                }
            }
            page_top += page_height + PAGE_GAP;
        }
        best.map(|(_, page_index, local_x, local_y)| (page_index, local_x, local_y))
    }

    fn pointer_down(&mut self) -> Result<bool> {
        let Some((x, y)) = self.cursor else {
            return Ok(false);
        };
        let Some((page_index, local_x, local_y)) = self.document_point(x, y, false) else {
            return Ok(false);
        };
        let Some(loc) = self.document.docx_hit_test(page_index, local_x, local_y)? else {
            return Ok(false);
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

    fn keyboard_input(&mut self, event_loop: &ActiveEventLoop, event: &KeyEvent) -> Result<()> {
        if event.state != ElementState::Pressed {
            return Ok(());
        }
        let command = self.modifiers.super_key() || self.modifiers.control_key();
        let key = event.logical_key.as_ref();
        if matches!(key, Key::Named(NamedKey::Escape)) {
            event_loop.exit();
            return Ok(());
        }
        if command
            && !event.repeat
            && let Key::Character(character) = key
        {
            if character.eq_ignore_ascii_case("s") {
                self.save_and_reopen()?;
                return Ok(());
            }
            if matches!(character, "+" | "=") {
                self.zoom_by(1.1);
                return Ok(());
            }
            if character == "-" {
                self.zoom_by(1.0 / 1.1);
                return Ok(());
            }
            if character == "0" {
                self.reset_zoom();
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
                Key::Named(NamedKey::Enter) => self.document.docx_enter()?,
                Key::Character(text)
                    if !command
                        && !self.modifiers.alt_key()
                        && text.chars().all(|character| !character.is_control()) =>
                {
                    self.document.docx_insert_text(text)?
                }
                _ => false,
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

    fn save_and_reopen(&mut self) -> Result<()> {
        if !matches!(&self.document.reference, ReferenceDocument::Docx(_)) {
            return Ok(());
        }
        let path = self
            .save_target
            .clone()
            .context("DOCX save target is unavailable")?;
        self.document.save_docx_to(&path)?;
        let reopened = load_document(&path, 0, self.document.max_texture_dimension_2d)
            .with_context(|| format!("reopen edited DOCX {}", path.display()))?;
        println!("saved and reopened edited DOCX: {}", path.display());
        self.document = reopened;
        self.scroll = 0.0;
        self.dragging = false;
        self.reset_caret_blink();
        self.clamp_scroll();
        self.update_title();
        self.request_redraw();
        Ok(())
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
        self.window
            .as_ref()
            .map(|window| f64::from(window.inner_size().height) / window.scale_factor())
            .unwrap_or(1.0)
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
            _ => "",
        };
        window.set_title(&format!(
            "{name} — {} — {:.0}% — {skipped} skipped{edit_state}",
            self.document.title_summary(),
            self.zoom * 100.0
        ));
    }

    fn fail(&mut self, event_loop: &ActiveEventLoop, error: impl std::fmt::Display) {
        self.fatal = Some(error.to_string());
        event_loop.exit();
    }
}

impl ApplicationHandler for Viewer {
    fn resumed(&mut self, event_loop: &ActiveEventLoop) {
        if let Err(error) = self.create_window(event_loop) {
            self.fail(event_loop, error);
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
            WindowEvent::CloseRequested => event_loop.exit(),
            WindowEvent::Resized(size) => self.resize(size),
            WindowEvent::ScaleFactorChanged { .. } => {
                if let Some(window) = &self.window {
                    self.resize(window.inner_size());
                }
            }
            WindowEvent::RedrawRequested => {
                if let Err(error) = self.render() {
                    self.fail(event_loop, error);
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
                    Err(error) => self.fail(event_loop, error),
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
                    Err(error) => self.fail(event_loop, error),
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
                    self.fail(event_loop, error);
                }
            }
            _ => {}
        }
    }

    fn about_to_wait(&mut self, event_loop: &ActiveEventLoop) {
        if self.focused && self.document.has_docx_caret() {
            let now = Instant::now();
            if now >= self.next_caret_blink {
                self.caret_visible = !self.caret_visible;
                self.next_caret_blink = now + CARET_BLINK;
                self.request_redraw();
            }
            event_loop.set_control_flow(ControlFlow::WaitUntil(self.next_caret_blink));
        } else {
            self.caret_visible = false;
            event_loop.set_control_flow(ControlFlow::Wait);
        }
    }
}
