import { expect, test } from 'bun:test';
import { createYrsSession } from '../../../packages/docx/src/yrs';
import { decodeMessages, encodeUpdate } from '../../../packages/docx/src/collaboration/protocol';
import { RetainedUpdateLog } from '../src/retention';

test('rehydrates real yrs text, formatting and identities after repeated checkpoints', async () => {
  const source = await createYrsSession();
  const offline = await createYrsSession();
  const restored = await createYrsSession();
  try {
    const { paraId } = source.createStory('body', 'Seed text');
    offline.loadState(source.encodeState());
    const log = new RetainedUpdateLog(8, 1024 * 1024);
    log.retain(encodeUpdate(source.encodeState()));
    source.onUpdate(update => log.retain(encodeUpdate(update)));
    for (let i = 0; i < 40; i++) source.insertText({ story: 'body', paraId, offset: 9 + i }, String(i % 10));
    source.formatRange({ story: 'body', start: { paraId, offset: 0 }, end: { paraId, offset: 4 } }, { bold: true });
    offline.insertText({ story: 'body', paraId, offset: 0 }, 'Offline ');
    log.retain(encodeUpdate(offline.encodeState()));
    log.checkpoint();
    for (const frame of log.snapshot()) for (const message of decodeMessages(frame)) {
      if (message.type === 'update' || message.type === 'sync-step-2') restored.applyUpdate(message.update);
    }
    source.applyUpdate(offline.encodeState());
    expect(restored.paragraphs('body')).toEqual(source.paragraphs('body'));
    expect(restored.storySegments('body')).toEqual(source.storySegments('body'));
    expect(restored.paragraphs('body')[0].text).toStartWith('Offline Seed text');
  } finally { source.destroy(); offline.destroy(); restored.destroy(); }
});
