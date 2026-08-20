use std::sync::Arc;

use anyhow::{Result, bail};
use vello::kurbo::Affine;
use vello::peniko::Color;
use vello::util::{RenderContext, RenderSurface};
use vello::{AaConfig, RenderParams, Renderer, RendererOptions, Scene};
use winit::application::ApplicationHandler;
use winit::dpi::{LogicalSize, PhysicalSize};
use winit::event::{ElementState, MouseScrollDelta, WindowEvent};
use winit::event_loop::{ActiveEventLoop, EventLoop};
use winit::keyboard::{Key, ModifiersState, NamedKey};
use winit::window::{Window, WindowId};

use crate::document::{DocumentView, ReferenceDocument};

const MARGIN: f64 = 40.0;
const PAGE_GAP: f64 = 24.0;

pub fn run(document: DocumentView) -> Result<()> {
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
    let mut app = Viewer::new(document);
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
    fatal: Option<String>,
}

impl Viewer {
    fn new(document: DocumentView) -> Self {
        Self {
            document,
            context: RenderContext::new(),
            window: None,
            surface: None,
            renderer: None,
            zoom: 1.0,
            scroll: 0.0,
            modifiers: ModifiersState::empty(),
            fatal: None,
        }
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
        for page in &self.document.pages {
            let x = ((f64::from(surface.config.width) - page.width * scale) / 2.0)
                .max(MARGIN * device_scale);
            scene.append(
                &page.scene,
                Some(Affine::translate((x, y)) * Affine::scale(scale)),
            );
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
        window.set_title(&format!(
            "{name} — {} — {:.0}% — {skipped} skipped",
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
            WindowEvent::KeyboardInput { event, .. }
                if event.state == ElementState::Pressed && !event.repeat =>
            {
                match event.logical_key.as_ref() {
                    Key::Character("+" | "=") => self.zoom_by(1.1),
                    Key::Character("-") => self.zoom_by(1.0 / 1.1),
                    Key::Character("0") => self.reset_zoom(),
                    Key::Named(NamedKey::Escape) => event_loop.exit(),
                    Key::Named(NamedKey::PageDown) => self.scroll_by(-self.viewport_height() * 0.9),
                    Key::Named(NamedKey::PageUp) => self.scroll_by(self.viewport_height() * 0.9),
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
            }
            _ => {}
        }
    }
}
