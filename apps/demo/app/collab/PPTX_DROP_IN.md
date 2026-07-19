# PPTX collaboration drop-in

TODO after PR #58 merges:

1. Generate `public/seeds/pptx.bin` beside the DOCX and XLSX seeds.
2. Import the PPTX editor stylesheet from its route entry.
3. In the PPTX client, reuse `useDemoRoom()` and `useCollabRoom()` with the PPTX `CollaborationProvider`.
4. Fetch the seed and pass `{ clientId, initialUpdate, onReplica }` through the editor's `collaboration` prop.
5. Render `CollaborationControls` in the existing page header.

```tsx
const room = useDemoRoom();
const createProvider = useCallback(
  (replica, transport) => new CollaborationProvider(replica, transport),
  [],
);
const collab = useCollabRoom(COLLAB_RELAY_ORIGIN, room, createProvider);

<PptxEditor
  collaboration={
    seed && collab.clientId
      ? { clientId: collab.clientId, initialUpdate: seed, onReplica: collab.onReplica }
      : undefined
  }
/>;
```
