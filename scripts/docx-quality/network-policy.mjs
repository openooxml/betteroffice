import { readFileSync } from 'node:fs';

const version = (name) =>
  JSON.parse(readFileSync(new URL(`../../packages/${name}/package.json`, import.meta.url), 'utf8'))
    .version.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const fontUrl = new RegExp(
  `^https://cdn\\.jsdelivr\\.net/npm/@betteroffice/(?:fonts@${version('fonts')}|fonts-cjk@${version('fonts-cjk')})/assets/[A-Za-z0-9-]+\\.(?:ttf|otf)$`
);

/** Allow only the jsDelivr font binaries the capture harness requests, fetched without credentials. */
export function isAllowedFontRequest(request) {
  return (
    fontUrl.test(request.url) &&
    request.method === 'GET' &&
    request.bodyBytes === 0 &&
    !request.referer &&
    !request.cookie
  );
}
