import { constants } from 'node:fs';
import { open, realpath, stat } from 'node:fs/promises';
import { isAbsolute, relative, resolve, sep } from 'node:path';
import { imageSize } from 'image-size';
import { DocumentToolError } from './types';

export async function readPptxImage(root: string, path: string) {
  function contained(path: string) {
    const file = resolve(root, path);
    const rel = relative(root, file);
    if (rel === '..' || rel.startsWith(`..${sep}`) || isAbsolute(rel)) throw new DocumentToolError('OUTSIDE_WORKSPACE', 'Choose an image path inside the configured workspace root.');
    return file;
  }
  const lexical = contained(path);
  const file = contained(await realpath(lexical));
  const handle = await open(file, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0) | (constants.O_NONBLOCK ?? 0));
  try {
    const info = await handle.stat();
    const current = await stat(contained(await realpath(lexical)));
    if (info.dev !== current.dev || info.ino !== current.ino) throw new DocumentToolError('FILE_CHANGED', 'Image changed during access. Retry with a stable workspace.');
    if (!info.isFile() || info.size > 8 * 1024 * 1024) throw new DocumentToolError('INVALID_IMAGE', 'Choose a raster image file up to 8 MiB.');
    const bytes = new Uint8Array(info.size + 1);
    let length = 0;
    while (length < bytes.length) {
      const read = await handle.read(bytes, length, bytes.length - length, length);
      if (!read.bytesRead) break;
      length += read.bytesRead;
    }
    const after = await handle.stat();
    if (length !== info.size || after.size !== info.size || after.mtimeMs !== info.mtimeMs) throw new DocumentToolError('FILE_CHANGED', 'Image changed during access. Retry with a stable workspace.');
    const data = bytes.subarray(0, length);
    let type;
    try { type = imageSize(data).type; }
    catch { throw new DocumentToolError('INVALID_IMAGE', 'Choose a PNG, JPEG, GIF, or WebP image.'); }
    const types = { png: 'image/png', jpg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp' } as const;
    const contentType = types[type as keyof typeof types];
    if (!contentType) throw new DocumentToolError('INVALID_IMAGE', 'Choose a PNG, JPEG, GIF, or WebP image.');
    return { base64: Buffer.from(data).toString('base64'), contentType };
  } finally { await handle.close(); }
}
