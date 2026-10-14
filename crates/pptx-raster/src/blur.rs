//! Separable box blur over premultiplied pixels, standing in for the Gaussian
//! blur tiny-skia does not ship.

use tiny_skia::Pixmap;

/// Box radius approximating a Gaussian of `sigma`; 0 when nothing would soften.
pub(crate) fn box_radius(sigma: f32) -> u32 {
    if !sigma.is_finite() || sigma <= 0.5 {
        return 0;
    }
    sigma.round().min(128.0) as u32
}

/// How far three box passes of `radius` carry ink.
pub(crate) fn spread(radius: u32) -> u32 {
    radius * 3
}

/// Three box passes per axis, which approximate a Gaussian of sigma `radius`.
pub(crate) fn blur(pixmap: &mut Pixmap, radius: u32) {
    let radius = radius as usize;
    let (width, height) = (pixmap.width() as usize, pixmap.height() as usize);
    if radius == 0 || width == 0 || height == 0 {
        return;
    }
    let mut scratch = vec![0u8; width * height * 4];
    for _ in 0..3 {
        blur_rows(pixmap.data_mut(), &mut scratch, width, height, radius);
        transpose(&scratch, pixmap.data_mut(), width, height);
        blur_rows(pixmap.data_mut(), &mut scratch, height, width, radius);
        transpose(&scratch, pixmap.data_mut(), height, width);
    }
}

fn blur_rows(src: &[u8], dst: &mut [u8], width: usize, height: usize, radius: usize) {
    let window = (radius * 2 + 1) as u32;
    for y in 0..height {
        let row = y * width * 4;
        let at = |x: usize| {
            let x = x.min(width - 1);
            let offset = row + x * 4;
            [
                src[offset] as u32,
                src[offset + 1] as u32,
                src[offset + 2] as u32,
                src[offset + 3] as u32,
            ]
        };
        let first = at(0);
        let mut sums = [0u32; 4];
        for (channel, sum) in sums.iter_mut().enumerate() {
            *sum = first[channel] * (radius as u32 + 1);
        }
        for x in 1..=radius {
            let pixel = at(x);
            for channel in 0..4 {
                sums[channel] += pixel[channel];
            }
        }
        for x in 0..width {
            let offset = row + x * 4;
            for channel in 0..4 {
                dst[offset + channel] = ((sums[channel] + window / 2) / window) as u8;
            }
            let leaving = at(x.saturating_sub(radius));
            let entering = at(x + radius + 1);
            for channel in 0..4 {
                sums[channel] = sums[channel] + entering[channel] - leaving[channel];
            }
        }
    }
}

fn transpose(src: &[u8], dst: &mut [u8], width: usize, height: usize) {
    for y in 0..height {
        for x in 0..width {
            let from = (y * width + x) * 4;
            let to = (x * height + y) * 4;
            dst[to..to + 4].copy_from_slice(&src[from..from + 4]);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tiny_skia::IntSize;

    fn opaque_dot(size: u32) -> Pixmap {
        let mut pixmap = Pixmap::from_vec(
            vec![0; (size * size * 4) as usize],
            IntSize::from_wh(size, size).unwrap(),
        )
        .unwrap();
        let center = (size / 2 * size + size / 2) as usize * 4;
        pixmap.data_mut()[center..center + 4].copy_from_slice(&[0, 0, 0, 255]);
        pixmap
    }

    fn alpha(pixmap: &Pixmap, x: usize, y: usize, width: usize) -> u8 {
        pixmap.data()[(y * width + x) * 4 + 3]
    }

    #[test]
    fn a_blur_spreads_ink_off_the_pixel_it_started_on() {
        let mut pixmap = opaque_dot(21);
        blur(&mut pixmap, 2);
        assert!(alpha(&pixmap, 10, 10, 21) < 255);
        assert!(alpha(&pixmap, 13, 10, 21) > 0);
        assert!(alpha(&pixmap, 10, 13, 21) > 0);
        assert_eq!(alpha(&pixmap, 0, 0, 21), 0);
    }

    #[test]
    fn a_zero_radius_leaves_the_pixmap_alone() {
        let mut pixmap = opaque_dot(9);
        let before = pixmap.data().to_vec();
        blur(&mut pixmap, 0);
        assert_eq!(pixmap.data(), before.as_slice());
    }
}
