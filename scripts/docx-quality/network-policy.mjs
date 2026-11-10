import { readFileSync } from 'node:fs';

const version = (name) =>
  JSON.parse(readFileSync(new URL(`../../packages/${name}/package.json`, import.meta.url), 'utf8'))
    .version.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const fontUrl = new RegExp(
  `^https://cdn\\.jsdelivr\\.net/npm/@betteroffice/(?:fonts@${version('fonts')}/dist/decoder/fontpack\\.wasm|fonts-latin@${version('fonts-latin')}/assets/[A-Za-z0-9-]+\\.ttf\\.zst|fonts-cjk@${version('fonts-cjk')}/assets/[A-Za-z0-9-]+\\.otf\\.zst|fonts-intl@${version('fonts-intl')}/assets/[A-Za-z0-9-]+\\.ttf\\.zst)$`
);

const STABLE_VERSION = /^\d+\.\d+\.\d+$/;

function publishedFontUrl(version) {
  const escaped = version.replace(/\./g, '\\.');
  return new RegExp(
    `^https://cdn\\.jsdelivr\\.net/npm/@betteroffice/(?:fonts@${escaped}/assets/[A-Za-z0-9-]+\\.(?:ttf|otf)|fonts-cjk@\\d+\\.\\d+\\.\\d+/assets/[A-Za-z0-9-]+\\.(?:ttf|otf)(?:\\.zst)?)$`
  );
}

/**
 * Allow only the jsDelivr font binaries the capture harness requests, fetched without credentials.
 * A published-channel capture also loads the fonts of the published release.
 */
export function isAllowedFontRequest(request, environment = process.env) {
  const published = environment.QUALITY_PUBLISHED_FONTS?.trim() ?? '';
  const allowed = fontUrl.test(request.url) || (STABLE_VERSION.test(published) && publishedFontUrl(published).test(request.url));
  return (
    allowed &&
    request.method === 'GET' &&
    request.bodyBytes === 0 &&
    !request.referer &&
    !request.cookie
  );
}
