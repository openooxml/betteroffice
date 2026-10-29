import { createEmptyDocx } from '../docx/rezip/createEmpty';
import { createResidentSaveWarmupSession } from './residentEngineSession';

export async function warmResidentSave(yieldSave: () => Promise<void>): Promise<void> {
  await yieldSave();
  const source = new Uint8Array(await createEmptyDocx());
  await yieldSave();
  const synthetic = createResidentSaveWarmupSession();
  try {
    await yieldSave();
    const hostJson = synthetic.openDocx(source);
    await yieldSave();
    const paragraph = synthetic.paragraphs('body')[0];
    if (!paragraph) return;
    await yieldSave();
    synthetic.insertText({ story: 'body', paraId: paragraph.paraId, offset: 0 }, 'warm');
    await yieldSave();
    await synthetic.save(source, hostJson, undefined, [], { full: false }, undefined, yieldSave);
  } finally {
    synthetic.destroy();
  }
}
