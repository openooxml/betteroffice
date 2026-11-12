/**
 * Faces a capture's font provider advertises and how each load ended. A load is recorded only when the provider
 * offered a face for the request, so a document asking for a family the pack does not carry records nothing and
 * falls back without a failure. A recorded load that failed is an advertised face that did not load.
 */

const fileOf = (url) => {
  try {
    return decodeURIComponent(new URL(url).pathname.split('/').pop() ?? '');
  } catch {
    return '';
  }
};

export function failureReason(load, blockedFiles = new Set()) {
  const http = /\bHTTP (\d{3})\b/.exec(load.error ?? '');
  if (http) return `http ${http[1]}`;
  if (load.face && blockedFiles.has(load.face)) return 'blocked';
  return 'error';
}

export function fontFaceSummary(fontLoads = [], advertised = [], networkViolations = []) {
  const blocked = new Set(networkViolations.map((request) => fileOf(request.url)).filter(Boolean));
  const loaded = new Set();
  const failed = {};
  for (const load of fontLoads) {
    const face = load.face ?? `${load.kind}:${load.family}`;
    if (load.ok) loaded.add(face);
    else if (!(face in failed)) failed[face] = failureReason(load, blocked);
  }
  return {
    advertised: new Set(advertised).size,
    loaded_faces: [...loaded].filter((face) => !(face in failed)).sort(),
    failed_faces: Object.fromEntries(Object.entries(failed).sort(([a], [b]) => a.localeCompare(b))),
  };
}
