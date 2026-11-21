const activations = new WeakMap<object, string>();
let nextActivation = 0;
let current: {
  activationId: string;
  event: Event;
  currentEvent(): Event | undefined;
  close(): void;
} | null = null;

export function pluginActivationId(activation: object): string {
  let id = activations.get(activation);
  if (!id) {
    id = `plugin-activation-${++nextActivation}`;
    activations.set(activation, id);
  }
  return id;
}

export function gestureFor(activationId: string): boolean {
  return current?.activationId === activationId && current.event.eventPhase !== Event.NONE &&
    current.currentEvent() === current.event;
}

export function trackPluginGestures(
  root: HTMLElement,
  trusted: (event: Event) => boolean = (event) => event.isTrusted,
  currentEvent: () => Event | undefined = () => root.ownerDocument.defaultView?.event
): () => void {
  const window = root.ownerDocument.defaultView;
  const types = ['click', 'keydown', 'pointerup'] as const;
  let close: (() => void) | null = null;
  const capture = (event: Event) => {
    current?.close();
    if (!trusted(event)) return;
    const marker = event.composedPath().find(
      (target): target is Element =>
        target instanceof Element && target.hasAttribute('data-plugin-activation')
    );
    const activationId = marker?.getAttribute('data-plugin-activation');
    if (!marker || !root.contains(marker) || !activationId) return;
    const timer = setTimeout(() => closeGesture(), 0);
    const closeGesture = () => {
      clearTimeout(timer);
      if (current?.event === event) current = null;
      close = null;
    };
    close = closeGesture;
    current = { activationId, event, currentEvent, close: closeGesture };
  };
  const bubble = (event: Event) => {
    if (current?.event === event) current.close();
  };
  const interrupt = (event: Event) => {
    if (current?.event !== event) current?.close();
  };
  for (const type of types) {
    root.addEventListener(type, capture, true);
    window?.addEventListener(type, interrupt, true);
    window?.addEventListener(type, bubble);
  }
  return () => {
    close?.();
    for (const type of types) {
      root.removeEventListener(type, capture, true);
      window?.removeEventListener(type, interrupt, true);
      window?.removeEventListener(type, bubble);
    }
  };
}
