import { openWorkbookSession } from '@betteroffice/xlsx';
import { createWorkbookEditPeer, hydratePeer } from '@betteroffice/xlsx/worker';

export const editableWorkbookSessionBackend = {
  open: openWorkbookSession,
  hydrate: hydratePeer,
  attach: createWorkbookEditPeer,
};
