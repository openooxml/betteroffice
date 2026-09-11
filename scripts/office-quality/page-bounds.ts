export type OfficePageBounds = {
  kind: 'office-page-bounds';
  pages: {
    width_pt: number;
    height_pt: number;
    width_px: number;
    height_px: number;
  }[];
};

const maximumPixels = 8192;

export function validatePageBounds(input: unknown): OfficePageBounds | null {
  if (input == null) return null;
  const profile = input as OfficePageBounds;
  if (
    profile.kind !== 'office-page-bounds' ||
    !Array.isArray(profile.pages) ||
    !profile.pages.length ||
    profile.pages.length > 100
  )
    throw new Error('Invalid Office page bounds profile');
  return {
    kind: 'office-page-bounds',
    pages: profile.pages.map((page) => {
      if (
        !page ||
        [page.width_pt, page.height_pt].some(
          (value) =>
            !Number.isFinite(value) || value <= 0 || value > (maximumPixels * 72) / 150
        ) ||
        [page.width_px, page.height_px].some(
          (value) => !Number.isInteger(value) || value < 1 || value > maximumPixels
        ) ||
        page.width_px !== Math.ceil((page.width_pt * 150) / 72) ||
        page.height_px !== Math.ceil((page.height_pt * 150) / 72)
      )
        throw new Error('Office page bounds must match the PDF extent at 150 DPI');
      return {
        width_pt: page.width_pt,
        height_pt: page.height_pt,
        width_px: page.width_px,
        height_px: page.height_px,
      };
    }),
  };
}

export function capturePageExtent(
  profile: OfficePageBounds | null,
  index: number,
  width: number,
  height: number
) {
  if (
    !Number.isInteger(index) ||
    index < 0 ||
    [width, height].some((value) => !Number.isInteger(value) || value < 1)
  )
    throw new Error('Invalid native canvas extent');
  const target = profile?.pages[index];
  const raw = { width_px: width, height_px: height };
  const output = target
    ? { width_px: target.width_px, height_px: target.height_px }
    : { ...raw };
  if (Math.abs(output.width_px - width) > 1 || Math.abs(output.height_px - height) > 1)
    throw new Error(`Page ${index + 1}: Office and native canvas extents differ by more than 1 pixel`);
  return {
    page: index + 1,
    raw,
    output,
    adjustment:
      output.width_px === width && output.height_px === height
        ? 'none'
        : 'office-page-bounds',
  };
}
