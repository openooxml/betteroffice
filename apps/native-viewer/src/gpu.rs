use std::fs;
use std::path::{Path, PathBuf};
use std::sync::mpsc;

use anyhow::{Context, Result, anyhow, bail};
use docx_raster::RenderResources;
use image::{DynamicImage, ImageFormat};
use vello::kurbo::Affine;
use vello::peniko::Color;
use vello::util::RenderContext;
use vello::{AaConfig, RenderParams, Renderer, RendererOptions, Scene};

use crate::document::{DocumentView, ReferenceDocument};
use crate::scene_shared::PageScene;

const DIFFERENCE_THRESHOLD: u8 = 8;

pub fn render_comparison(
    document: &DocumentView,
    page_index: usize,
    output: &Path,
    scale: f64,
) -> Result<()> {
    let page = document
        .pages
        .get(page_index)
        .with_context(|| format!("page {} is out of range", page_index + 1))?;
    let rendered = render_page_gpu(page, scale)?;
    image::save_buffer_with_format(
        output,
        &rendered.rgba,
        rendered.width,
        rendered.height,
        image::ColorType::Rgba8,
        ImageFormat::Png,
    )
    .with_context(|| format!("write Vello PNG {}", output.display()))?;

    let (raster_bytes, skipped_raster_images) = match &document.reference {
        ReferenceDocument::Docx(reference) => {
            let resources = RenderResources::new(
                &reference.fonts.store,
                &reference.fonts.chains,
                &reference.images.raw,
            );
            let raster = docx_raster::render_page(&reference.display_list, page_index, &resources)
                .map_err(anyhow::Error::msg)?;
            (raster.bytes, Some(raster.skipped_images))
        }
        ReferenceDocument::Xlsx(reference) => (
            xlsx_raster::render_png(&reference.display_list).map_err(anyhow::Error::msg)?,
            None,
        ),
    };
    let reference_path = raster_path(output);
    fs::write(&reference_path, &raster_bytes)
        .with_context(|| format!("write raster PNG {}", reference_path.display()))?;
    let reference = image::load_from_memory_with_format(&raster_bytes, ImageFormat::Png)?;
    let metrics = compare(&rendered.rgba, rendered.width, rendered.height, reference)?;

    println!("document: {}", document.source.display());
    match &document.reference {
        ReferenceDocument::Docx(reference) => {
            println!("page: {} of {}", page_index + 1, document.pages.len());
            println!("gpu: {}", rendered.adapter);
            println!("font requirements: {:?}", reference.fonts.requirements);
        }
        ReferenceDocument::Xlsx(reference) => {
            println!(
                "sheet: {} of {} ({})",
                reference.sheet_index + 1,
                reference.sheet_count,
                reference.sheet_name
            );
            println!("gpu: {}", rendered.adapter);
            println!(
                "charts: {}, placeholders: {}",
                reference.chart_count, reference.chart_placeholders
            );
        }
    }
    println!("vello PNG: {}", output.display());
    println!("raster PNG: {}", reference_path.display());
    let skipped_label = match &document.reference {
        ReferenceDocument::Docx(_) => "primitives",
        ReferenceDocument::Xlsx(_) => "commands",
    };
    println!(
        "skipped Vello {skipped_label}: {} {:?}",
        page.skipped.total(),
        page.skipped.counts
    );
    if !page.skipped.reasons.is_empty() {
        println!("skip reasons: {:?}", page.skipped.reasons);
    }
    if let Some(skipped_images) = skipped_raster_images {
        println!("skipped raster images: {skipped_images}");
    }
    println!(
        "mean absolute difference RGBA: {:.4}, {:.4}, {:.4}, {:.4}",
        metrics.mean[0], metrics.mean[1], metrics.mean[2], metrics.mean[3]
    );
    println!(
        "pixels differing above threshold {}: {:.4}%",
        DIFFERENCE_THRESHOLD,
        metrics.fraction * 100.0
    );
    Ok(())
}

struct GpuImage {
    rgba: Vec<u8>,
    width: u32,
    height: u32,
    adapter: String,
}

fn render_page_gpu(page: &PageScene, scale: f64) -> Result<GpuImage> {
    let width = scaled_dimension(page.width, scale)?;
    let height = scaled_dimension(page.height, scale)?;
    let mut scene = Scene::new();
    scene.append(&page.scene, Some(Affine::scale(scale)));
    let mut context = RenderContext::new();
    let device_id = pollster::block_on(context.device(None)).context("no compatible GPU device")?;
    let handle = &context.devices[device_id];
    let adapter = handle.adapter().get_info();
    let mut renderer = Renderer::new(&handle.device, RendererOptions::default())?;
    let texture = handle.device.create_texture(&wgpu::TextureDescriptor {
        label: Some("BetterOffice Vello headless target"),
        size: wgpu::Extent3d {
            width,
            height,
            depth_or_array_layers: 1,
        },
        mip_level_count: 1,
        sample_count: 1,
        dimension: wgpu::TextureDimension::D2,
        format: wgpu::TextureFormat::Rgba8Unorm,
        usage: wgpu::TextureUsages::STORAGE_BINDING
            | wgpu::TextureUsages::TEXTURE_BINDING
            | wgpu::TextureUsages::COPY_SRC,
        view_formats: &[],
    });
    let view = texture.create_view(&wgpu::TextureViewDescriptor::default());
    renderer.render_to_texture(
        &handle.device,
        &handle.queue,
        &scene,
        &view,
        &RenderParams {
            base_color: Color::WHITE,
            width,
            height,
            antialiasing_method: AaConfig::Msaa16,
        },
    )?;
    let rgba = read_texture(&handle.device, &handle.queue, &texture, width, height)?;
    Ok(GpuImage {
        rgba,
        width,
        height,
        adapter: format!("{} ({:?})", adapter.name, adapter.backend),
    })
}

fn read_texture(
    device: &wgpu::Device,
    queue: &wgpu::Queue,
    texture: &wgpu::Texture,
    width: u32,
    height: u32,
) -> Result<Vec<u8>> {
    let row_bytes = width.checked_mul(4).context("row byte count overflow")?;
    let padded_row_bytes =
        row_bytes.div_ceil(wgpu::COPY_BYTES_PER_ROW_ALIGNMENT) * wgpu::COPY_BYTES_PER_ROW_ALIGNMENT;
    let buffer_size = u64::from(padded_row_bytes)
        .checked_mul(u64::from(height))
        .context("readback buffer size overflow")?;
    let buffer = device.create_buffer(&wgpu::BufferDescriptor {
        label: Some("BetterOffice Vello readback"),
        size: buffer_size,
        usage: wgpu::BufferUsages::COPY_DST | wgpu::BufferUsages::MAP_READ,
        mapped_at_creation: false,
    });
    let mut encoder = device.create_command_encoder(&wgpu::CommandEncoderDescriptor {
        label: Some("BetterOffice Vello readback encoder"),
    });
    encoder.copy_texture_to_buffer(
        wgpu::TexelCopyTextureInfo {
            texture,
            mip_level: 0,
            origin: wgpu::Origin3d::ZERO,
            aspect: wgpu::TextureAspect::All,
        },
        wgpu::TexelCopyBufferInfo {
            buffer: &buffer,
            layout: wgpu::TexelCopyBufferLayout {
                offset: 0,
                bytes_per_row: Some(padded_row_bytes),
                rows_per_image: Some(height),
            },
        },
        wgpu::Extent3d {
            width,
            height,
            depth_or_array_layers: 1,
        },
    );
    queue.submit(Some(encoder.finish()));
    let slice = buffer.slice(..);
    let (sender, receiver) = mpsc::channel();
    slice.map_async(wgpu::MapMode::Read, move |result| {
        let _ = sender.send(result);
    });
    device.poll(wgpu::PollType::wait_indefinitely())?;
    receiver
        .recv()
        .context("GPU map callback did not run")?
        .map_err(|error| anyhow!(error))?;
    let mapped = slice.get_mapped_range();
    let mut rgba = Vec::with_capacity(row_bytes as usize * height as usize);
    for row in mapped.chunks_exact(padded_row_bytes as usize) {
        rgba.extend_from_slice(&row[..row_bytes as usize]);
    }
    drop(mapped);
    buffer.unmap();
    Ok(rgba)
}

fn scaled_dimension(value: f64, scale: f64) -> Result<u32> {
    let value = (value * scale).ceil();
    if !value.is_finite() || !(1.0..=16_384.0).contains(&value) {
        bail!("scaled page dimension is outside 1..=16384");
    }
    Ok(value as u32)
}

struct DifferenceMetrics {
    mean: [f64; 4],
    fraction: f64,
}

fn compare(
    vello: &[u8],
    width: u32,
    height: u32,
    reference: DynamicImage,
) -> Result<DifferenceMetrics> {
    let reference = if reference.width() != width || reference.height() != height {
        reference.resize_exact(width, height, image::imageops::FilterType::Lanczos3)
    } else {
        reference
    }
    .into_rgba8();
    let expected = width as usize * height as usize * 4;
    if vello.len() != expected || reference.as_raw().len() != expected {
        bail!("comparison image dimensions are inconsistent");
    }
    let mut sums = [0u64; 4];
    let mut different = 0u64;
    for (left, right) in vello
        .chunks_exact(4)
        .zip(reference.as_raw().chunks_exact(4))
    {
        let mut pixel_differs = false;
        for channel in 0..4 {
            let difference = left[channel].abs_diff(right[channel]);
            sums[channel] += u64::from(difference);
            pixel_differs |= difference > DIFFERENCE_THRESHOLD;
        }
        different += u64::from(pixel_differs);
    }
    let pixels = u64::from(width) * u64::from(height);
    Ok(DifferenceMetrics {
        mean: sums.map(|sum| sum as f64 / pixels as f64),
        fraction: different as f64 / pixels as f64,
    })
}

fn raster_path(output: &Path) -> PathBuf {
    let stem = output
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("page");
    output.with_file_name(format!("{stem}.raster.png"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn derives_reference_path() {
        assert_eq!(
            raster_path(Path::new("out/page.png")),
            Path::new("out/page.raster.png")
        );
    }
}
