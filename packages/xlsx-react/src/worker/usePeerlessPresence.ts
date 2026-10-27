import { normalizeRange } from '@betteroffice/xlsx';
import { useEffect, useLayoutEffect, useState } from 'react';
import type { Selection, WorkbookSheetSummary } from '@betteroffice/xlsx';
import type { AwarenessPeer, CollaborationAwareness } from '@betteroffice/xlsx/collaboration';
import type { PeerlessWorkbookSession } from './usePeerlessSessionWorkbook';

export function usePeerlessPresence(run: PeerlessWorkbookSession | null,
  provider: CollaborationAwareness | null | undefined, active: number,
  selection: Selection | null, sheets: WorkbookSheetSummary[]): readonly AwarenessPeer[] {
  const [presence, setPresence] = useState<{ run: PeerlessWorkbookSession | null;
    provider: typeof provider; peers: readonly AwarenessPeer[] }>({ run: null, provider: null, peers: [] });
  useLayoutEffect(() => {
    setPresence({ run, provider, peers: [] });
    if (!run?.current || run.failure || !provider) return;
    const off = provider.onAwareness((peers) => {
      if (run.current && !run.retiring && !run.failure) setPresence({ run, provider, peers: [...peers] });
    });
    return () => { off(); provider.setCursor(null); };
  }, [run, run?.failure, provider]);
  useEffect(() => {
    if (!provider) return;
    if (!run?.ready) { provider.setCursor(null); return; }
    const sheet = sheets.find((entry) => entry.index === active)?.id;
    const range = selection ? normalizeRange(selection) : null;
    provider.setCursor(sheet && selection && range ? {
      sheet,
      anchor: selection.range ? { row: range.top, col: range.left } : { ...selection.anchor },
      head: selection.range ? { row: range.bottom, col: range.right } : { ...selection.focus },
    } : null);
  }, [run, run?.ready, provider, active, selection, sheets]);
  return run?.current && !run.retiring && !run.failure && presence.run === run && presence.provider === provider ? presence.peers : [];
}
