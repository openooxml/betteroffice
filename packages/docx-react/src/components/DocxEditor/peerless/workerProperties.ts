import type { createPeerlessShellController, PeerlessWorkerQuery } from './shellController';
import type { PeerlessInputIntent } from './inputEvents';

export type WorkerShellController = ReturnType<typeof createPeerlessShellController>;
export type WorkerPropertyQuery = PeerlessWorkerQuery<'properties'>;

export async function readWorkerSection(controller: WorkerShellController): Promise<WorkerPropertyQuery> {
  const parts = await controller.query({ kind: 'partSnapshot' });
  const pageIndex = controller.ref.getCurrentPage() - 1;
  const page = controller.getSnapshot().view?.geometry?.pages.find(page => page.page.pageIndex === pageIndex);
  const source = parts.value.metadata.body.sections ?? [];
  const index = page?.displayPage?.sectionIndex;
  const section = index !== undefined && index >= 0 && index < source.length - 1 ? source[index]?.id ?? String(index) : 'final';
  return controller.query({ kind: 'properties', target: { kind: 'section', section } });
}

export async function applyWorkerProperty(controller: WorkerShellController, query: PeerlessWorkerQuery, intent: PeerlessInputIntent,
  onError?: (error: Error) => void): Promise<boolean> {
  try {
    const env = controller.environment();
    if (env.readOnly || env.mode !== 'editing') throw new Error('Document is not editable');
    const result = await controller.applyQuery(query, omitUndefined(intent));
    if (!result.ok) throw new Error(result.failure.message);
    controller.ref.focus();
    return true;
  } catch (error) { onError?.(error instanceof Error ? error : new Error(String(error))); return false; }
}

function omitUndefined<T>(value: T): T {
  if (Array.isArray(value)) return value.map(omitUndefined) as T;
  if (value && typeof value === 'object' && Object.getPrototypeOf(value) === Object.prototype) {
    return Object.fromEntries(Object.entries(value).filter(([, value]) => value !== undefined).map(([key, value]) => [key, omitUndefined(value)])) as T;
  }
  return value;
}
