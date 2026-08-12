/**
 * CJK font binaries for `@betteroffice/fonts` — an optional add-on, installed
 * only by hosts that open Chinese, Japanese or Korean documents.
 *
 * This package is bytes, not policy: the face manifest, the Word-family alias
 * table and the whole resolution chain live in `@betteroffice/fonts`, which
 * reaches these assets through an optional dynamic import. Depending on this
 * package directly buys nothing — install it alongside `@betteroffice/fonts`
 * and the CJK faces become resolvable automatically.
 *
 * It exists as a separate package purely for install size. npm has no
 * partial-tarball fetch, so a subpath export inside `@betteroffice/fonts`
 * would still put these 33 MB in every consumer's `node_modules`; only a
 * package boundary keeps them out.
 *
 * The binaries are static CFF (OTTO) Regulars from noto-cjk's SubsetOTF
 * distribution, NOT the google/fonts variable TTFs: those VFs default to the
 * Thin (wght=100) instance, so the Rust engine (which reads default-instance
 * advances) and the browser (which applies wght=400) would disagree on the
 * same bytes. The statics keep both sides identical.
 */

/**
 * Asset filename -> URL resolver, keyed exactly as `@betteroffice/fonts`'
 * face manifest names these files.
 *
 * Every entry is a `new URL()` with a STRING LITERAL specifier: bundlers only
 * statically resolve (and therefore emit) the asset when the specifier is a
 * literal — a template expression collapses to a single wrong asset under
 * webpack/Turbopack.
 */
export const CJK_FONT_ASSET_URLS: Record<string, () => URL> = {
  'NotoSansJP-Regular.otf': () => new URL('../assets/NotoSansJP-Regular.otf', import.meta.url),
  'NotoSansKR-Regular.otf': () => new URL('../assets/NotoSansKR-Regular.otf', import.meta.url),
  'NotoSansSC-Regular.otf': () => new URL('../assets/NotoSansSC-Regular.otf', import.meta.url),
  'NotoSansTC-Regular.otf': () => new URL('../assets/NotoSansTC-Regular.otf', import.meta.url),
  'NotoSerifSC-Regular.otf': () => new URL('../assets/NotoSerifSC-Regular.otf', import.meta.url),
};
