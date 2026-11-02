import type { SessionScope } from '../../../../shared/office-session/transport';

export { compileWasm as compileSessionWasm } from '../../../../shared/office-session/wasm';

export function sessionWorkerScope(): SessionScope {
  return self as unknown as SessionScope;
}
