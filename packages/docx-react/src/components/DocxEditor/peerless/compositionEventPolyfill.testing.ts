type EventScope = Record<string, unknown>;

function carriesData(ctor: unknown): boolean {
  if (typeof ctor !== 'function') return false;
  try {
    return (new (ctor as typeof CompositionEvent)('compositionend', { data: 'x' })).data === 'x';
  } catch {
    return false;
  }
}

export function installCompositionEventPolyfill(): () => void {
  const scopes = [...new Set([globalThis, globalThis.document?.defaultView].filter(Boolean))] as EventScope[];
  const targets = scopes.filter((scope) => !carriesData(scope.CompositionEvent));
  if (targets.length === 0) return () => {};
  const previous = targets.map((scope) => Object.getOwnPropertyDescriptor(scope, 'CompositionEvent'));
  const Base = ((globalThis.document?.defaultView as EventScope | null)?.UIEvent ?? globalThis.UIEvent) as typeof UIEvent;
  class CompositionEventPolyfill extends Base {
    data: string;
    constructor(type: string, init: CompositionEventInit = {}) {
      super(type, init);
      this.data = init.data ?? '';
    }
    initCompositionEvent(type: string, bubbles = false, cancelable = false, view: Window | null = null, data = '') {
      this.initEvent(type, bubbles, cancelable);
      Object.defineProperties(this, {
        view: { configurable: true, value: view },
        detail: { configurable: true, value: 0 },
      });
      this.data = data;
    }
  }
  for (const scope of targets) {
    Object.defineProperty(scope, 'CompositionEvent', { configurable: true, writable: true, value: CompositionEventPolyfill });
  }
  return () => {
    targets.forEach((scope, index) => {
      const descriptor = previous[index];
      if (descriptor) Object.defineProperty(scope, 'CompositionEvent', descriptor);
      else Reflect.deleteProperty(scope, 'CompositionEvent');
    });
  };
}
