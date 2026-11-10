export function trackFocus(
  roots: () => readonly Element[],
  onChange: (focused: boolean) => void
): () => void {
  const document = roots()[0]?.ownerDocument ?? globalThis.document;
  const window = document.defaultView!;
  let focused: boolean | null = null;
  let active = true;
  let windowFocused = true;
  let queued = false;
  let observed: readonly Element[] = [];
  const observer = new MutationObserver(() => check());
  const observeRoots = () => {
    const next = roots().filter((root) => root.isConnected);
    if (next.length === observed.length && next.every((root, index) => root === observed[index])) return;
    observer.disconnect();
    observed = next;
    for (const root of observed) observer.observe(root, { childList: true, subtree: true });
  };
  const inside = (target: EventTarget | null) =>
    target instanceof Node && target.isConnected &&
    roots().some((root) => root.isConnected && root.contains(target));
  const change = (next: boolean) => {
    if (!active) return;
    if (next) observeRoots();
    else {
      observer.disconnect();
      observed = [];
    }
    if (focused === next) return;
    focused = next;
    onChange(next);
  };
  const check = () => change(windowFocused && inside(document.activeElement));
  const recheck = () => {
    if (queued) return;
    queued = true;
    queueMicrotask(() => { queued = false; if (active) check(); });
  };
  const focusIn = (event: FocusEvent) => change(windowFocused && inside(event.target));
  const focusOut = (event: FocusEvent) => {
    if (inside(event.relatedTarget)) change(windowFocused);
    else recheck();
  };
  const blur = () => { windowFocused = false; change(false); };
  const focus = () => { windowFocused = true; recheck(); };
  document.addEventListener('focusin', focusIn, true);
  document.addEventListener('focusout', focusOut, true);
  window.addEventListener('blur', blur);
  window.addEventListener('focus', focus);
  check();
  return () => {
    active = false;
    observer.disconnect();
    document.removeEventListener('focusin', focusIn, true);
    document.removeEventListener('focusout', focusOut, true);
    window.removeEventListener('blur', blur);
    window.removeEventListener('focus', focus);
  };
}
