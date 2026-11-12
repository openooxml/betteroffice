/** Summary of the faces a capture's font provider advertises and how each load of them ended. */

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

export function fontFaceSummary(fontLoads = [], advertised = [], networkViolations = [], baseLatin = []) {
  const offered = new Set([...advertised, ...baseLatin]);
  const blocked = new Set(networkViolations.map((request) => fileOf(request.url)).filter(Boolean));
  const loaded = new Set();
  const failed = {};
  let unattributedLoaded = 0;
  let unattributedFailed = 0;
  for (const load of fontLoads) {
    if (!load.face || !offered.has(load.face)) {
      if (load.ok) unattributedLoaded++;
      else unattributedFailed++;
      continue;
    }
    if (load.ok) loaded.add(load.face);
    else if (!(load.face in failed)) failed[load.face] = failureReason(load, blocked);
  }
  return {
    advertised: offered.size,
    base_latin: new Set(baseLatin).size,
    loaded_faces: [...loaded].filter((face) => !(face in failed)).sort(),
    failed_faces: Object.fromEntries(Object.entries(failed).sort(([a], [b]) => a.localeCompare(b))),
    unattributed: { loaded: unattributedLoaded, failed: unattributedFailed },
  };
}
