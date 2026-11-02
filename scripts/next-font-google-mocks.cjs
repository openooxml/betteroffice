const { join } = require('node:path');

const font = join(__dirname, '../crates/docx-edit/tests/fonts/LiberationSans-Regular.woff2');

function face(url) {
  const family = new URL(url).searchParams.getAll('family').map((value) => value.split(':')[0])[0] ?? 'Fallback';
  return `@font-face { font-family: '${family}'; font-style: normal; font-weight: 100 900; font-display: swap; src: url(${font}) format('woff2'); }`;
}

module.exports = new Proxy({}, {
  get: (_target, url) => (typeof url === 'string' && url.startsWith('https://') ? face(url) : undefined),
  has: (_target, url) => typeof url === 'string' && url.startsWith('https://'),
});
