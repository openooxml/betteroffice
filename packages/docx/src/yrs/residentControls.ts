import type { DocxContentControl } from './contentControls';
import type { YrsContentControlValue } from './index';
import type { ResidentEngineSession } from './residentEngineSession';

export interface ResidentControlPosition { story: string; index: number; encoded: number[] }
export interface ResidentControlTarget { control: DocxContentControl; at: ResidentControlPosition }
export interface ResidentControlSnapshot { version: string; controls: { handle: string; control: DocxContentControl }[] }
interface ControlHandle { position: ResidentControlPosition }
const handles = new WeakMap<ResidentEngineSession, Map<string, ControlHandle>>();

function identity(position: ResidentControlPosition): string { return JSON.stringify([position.story, Array.from(position.encoded)]); }

export function readControlHandles(engine: ResidentEngineSession): ResidentControlSnapshot {
  let registry = handles.get(engine);
  if (!registry) { registry = new Map(); handles.set(engine, registry); }
  const controls = engine.controlTargets().map(({ at, control }) => {
    const position = at;
    const known = [...registry!].find(([, entry]) => identity(entry.position) === identity(position));
    const handle = known?.[0] ?? crypto.randomUUID();
    if (!known) registry!.set(handle, { position });
    return { handle, control };
  });
  return { version: engine.proposalEngine.version(), controls };
}

export function resolveControlHandle(engine: ResidentEngineSession, handle: string): ResidentControlTarget | null {
  const entry = handles.get(engine)?.get(handle);
  if (!entry) return null;
  return engine.controlTargets().find(target => identity(target.at) === identity(entry.position)) ?? null;
}

export function validControlValue(value: unknown): value is YrsContentControlValue {
  if (typeof value === 'string') return !/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/u.test(value);
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const data = value as Record<string, unknown>;
  if (data.kind === 'checkbox') return Object.keys(data).every(key => ['kind', 'checked'].includes(key)) && typeof data.checked === 'boolean';
  if (data.kind === 'dropdown') return Object.keys(data).every(key => ['kind', 'value'].includes(key)) && typeof data.value === 'string';
  if (data.kind === 'date') return Object.keys(data).every(key => ['kind', 'date'].includes(key)) && typeof data.date === 'string';
  return false;
}
