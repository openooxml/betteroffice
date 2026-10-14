import { useEffect, useState } from 'react';

const CARET_BLINK_MS = 530;

export function useCaretBlink(eligible: boolean, projection: unknown): boolean {
  const [pageVisible, setPageVisible] = useState(
    () => typeof document === 'undefined' || document.visibilityState === 'visible'
  );
  const [caretVisible, setCaretVisible] = useState(false);

  useEffect(() => {
    if (typeof document === 'undefined') return;
    const update = () => setPageVisible(document.visibilityState === 'visible');
    document.addEventListener('visibilitychange', update);
    return () => document.removeEventListener('visibilitychange', update);
  }, []);

  useEffect(() => {
    const blinking = eligible && pageVisible;
    setCaretVisible(blinking);
    if (!blinking) return;
    const timer = window.setInterval(
      () => setCaretVisible((visible) => !visible),
      CARET_BLINK_MS
    );
    return () => window.clearInterval(timer);
  }, [eligible, pageVisible, projection]);

  return caretVisible;
}
