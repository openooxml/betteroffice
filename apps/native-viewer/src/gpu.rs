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
const MAX_HEADLESS_PIXELS: u64 = 33_000_000;

pub fn create_render_context() -> Result<(RenderContext, u32)> {
    let mut context = RenderContext::new();
    let device_id = pollster::block_on(context.device(None)).context("no compatible GPU device")?;
    let limit = context.devices[device_id]
        .device
        .limits()
        .max_texture_dimension_2d;
    Ok((context, limit))
}

pub fn render_comparison(
    context: &mut RenderContext,
    document: &DocumentView,
    page_index: usize,
    output: &Path,
    scale: f64,
) -> Result<()> {
    let selection = if matches!(&document.reference, ReferenceDocument::Pptx(_)) {
        "slide"
    } else {
        "page"
    };
    let page = document
        .pages
        .get(page_index)
        .with_context(|| format!("{selection} {} is out of range", page_index + 1))?;
    let rendered = render_page_gpu(context, page, scale)?;
    image::save_buffer_with_format(
        output,
        &rendered.rgba,
        rendered.width,
        rendered.height,
        image::ColorType::Rgba8,
        ImageFormat::Png,
    )
    .with_context(|| format!("write Vello PNG {}", output.display()))?;

    if let ReferenceDocument::Pptx(reference) = &document.reference {
        let summary = reference
            .editor
            .summaries()
            .get(page_index)
            .context("PPTX slide has no translation summary")?;
        println!("document: {}", document.source.display());
        println!(
            "slide: {} of {}",
            page_index + 1,
            reference.editor.slide_count()
        );
        println!("gpu: {}", rendered.adapter);
        println!("font faces: {:?}", reference.editor.font_faces());
        println!("vello PNG: {}", output.display());
        println!("raster PNG: not produced (PPTX has no raster backend)");
        println!(
            "pptx summary: {}",
            serde_json::to_string(&summary.structured(&page.skipped))?
        );
        return Ok(());
    }

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
            xlsx_raster::render_png(reference.editor.display_list()).map_err(anyhow::Error::msg)?,
            None,
        ),
        ReferenceDocument::Pptx(_) => unreachable!(),
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
        ReferenceDocument::Pptx(_) => unreachable!(),
    }
    println!("vello PNG: {}", output.display());
    println!("raster PNG: {}", reference_path.display());
    let skipped_label = match &document.reference {
        ReferenceDocument::Docx(_) => "primitives",
        ReferenceDocument::Xlsx(_) => "commands",
        ReferenceDocument::Pptx(_) => unreachable!(),
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

fn render_page_gpu(context: &mut RenderContext, page: &PageScene, scale: f64) -> Result<GpuImage> {
    let width = scaled_dimension(page.width, scale)?;
    let height = scaled_dimension(page.height, scale)?;
    let mut scene = Scene::new();
    scene.append(&page.background, Some(Affine::scale(scale)));
    scene.append(&page.scene, Some(Affine::scale(scale)));
    let device_id = pollster::block_on(context.device(None)).context("no compatible GPU device")?;
    let handle = &context.devices[device_id];
    validate_target_size(
        width,
        height,
        handle.device.limits().max_texture_dimension_2d,
    )?;
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
    validate_readback(&rgba, width, height)?;
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
    if !value.is_finite() || !(1.0..=f64::from(u32::MAX)).contains(&value) {
        bail!("scaled page dimension is outside the u32 range");
    }
    Ok(value as u32)
}

fn validate_target_size(width: u32, height: u32, dimension_limit: u32) -> Result<()> {
    if width > dimension_limit || height > dimension_limit {
        bail!(
            "requested headless target {width}x{height} exceeds GPU texture dimension ceiling {dimension_limit}"
        );
    }
    let pixels = u64::from(width) * u64::from(height);
    if pixels > MAX_HEADLESS_PIXELS {
        bail!(
            "requested headless target {width}x{height} has {pixels} pixels, exceeding the {MAX_HEADLESS_PIXELS}-pixel rendering ceiling"
        );
    }
    Ok(())
}

fn validate_readback(rgba: &[u8], width: u32, height: u32) -> Result<()> {
    if rgba.as_chunks::<4>().0.iter().all(|pixel| pixel[3] == 0) {
        bail!("Vello rendered an all-transparent {width}x{height} target");
    }
    Ok(())
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
    let (vello_pixels, _) = vello.as_chunks::<4>();
    let (reference_pixels, _) = reference.as_raw().as_chunks::<4>();
    for (left, right) in vello_pixels.iter().zip(reference_pixels) {
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
    use crate::document::{ReferenceDocument, load_document};
    use crate::test_fixtures;

    #[test]
    fn derives_reference_path() {
        assert_eq!(
            raster_path(Path::new("out/page.png")),
            Path::new("out/page.raster.png")
        );
    }

    #[test]
    fn rejects_targets_over_device_dimension_limit() {
        let error = validate_target_size(100, 8_200, 8_192).unwrap_err();
        assert!(error.to_string().contains("100x8200"));
        assert!(error.to_string().contains("8192"));
    }

    #[test]
    fn rejects_targets_over_area_limit() {
        let error = validate_target_size(5_075, 6_567, 8_192).unwrap_err();
        assert!(error.to_string().contains("5075x6567"));
        assert!(error.to_string().contains("33000000"));
        validate_target_size(5_018, 6_487, 8_192).unwrap();
    }

    #[test]
    fn rejects_all_transparent_readback() {
        let error = validate_readback(&[0; 16], 2, 2).unwrap_err();
        assert!(error.to_string().contains("all-transparent 2x2"));
        validate_readback(&[0, 0, 0, 1], 1, 1).unwrap();
    }

    #[test]
    fn embedded_and_rotated_image_pixels_agree_with_raster() {
        let Ok((mut context, limit)) = create_render_context() else {
            return;
        };
        for rotation in [None, Some(45.0)] {
            let path = test_fixtures::write_docx("gpu-image", &test_fixtures::image_docx(rotation));
            let document = load_document(&path, 0, limit).unwrap();
            let rendered = render_page_gpu(&mut context, &document.pages[0], 1.0).unwrap();
            let ReferenceDocument::Docx(reference) = &document.reference else {
                panic!("expected DOCX reference");
            };
            let resources = RenderResources::new(
                &reference.fonts.store,
                &reference.fonts.chains,
                &reference.images.raw,
            );
            let raster = docx_raster::render_page(&reference.display_list, 0, &resources).unwrap();
            let raster = image::load_from_memory_with_format(&raster.bytes, ImageFormat::Png)
                .unwrap()
                .into_rgba8();
            assert_eq!((rendered.width, rendered.height), raster.dimensions());
            let vello_green = green_pixels(&rendered.rgba);
            let raster_green = green_pixels(raster.as_raw());
            println!(
                "rotation {rotation:?}: Vello green pixels {vello_green}, raster green pixels {raster_green}"
            );
            assert!(raster_green > 1_000);
            let difference = vello_green.abs_diff(raster_green);
            assert!(
                difference * 100 <= raster_green,
                "rotation {rotation:?}: Vello {vello_green}, raster {raster_green}"
            );
            fs::remove_file(path).unwrap();
        }
    }

    fn green_pixels(rgba: &[u8]) -> usize {
        rgba.as_chunks::<4>()
            .0
            .iter()
            .filter(|pixel| pixel[0] < 64 && pixel[1] > 192 && pixel[2] < 64 && pixel[3] > 192)
            .count()
    }
}
