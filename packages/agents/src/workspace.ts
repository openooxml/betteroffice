import { constants } from 'node:fs';
import { lstat, open, readdir, realpath, stat, type FileHandle } from 'node:fs/promises';
import { basename, dirname, extname, isAbsolute, relative, resolve, sep } from 'node:path';
import { openDocx, type DocxAgentDocument } from './document';
import { openXlsx, type XlsxAgentWorkbook } from './xlsx';
import { openPptx, type PptxAgentPresentation } from './pptx';
import { DocumentToolError, type DocumentRenderer } from './types';

export class FileWorkspace {
  private readonly documents = new Map<string, { path: string; document: DocxAgentDocument | XlsxAgentWorkbook | PptxAgentPresentation }>();
  private nextId = 1;

  private constructor(readonly root: string, private readonly renderer?: DocumentRenderer) {}

  static async create(root: string, renderer?: DocumentRenderer) {
    const canonical = await realpath(root);
    if (!(await stat(canonical)).isDirectory()) throw new DocumentToolError('INVALID_ROOT', 'Workspace root must be a directory.');
    return new FileWorkspace(canonical, renderer);
  }

  async files(directory = '.', offset = 0) {
    const folder = await this.path(directory);
    const entries = (await readdir(folder, { withFileTypes: true }))
      .filter(entry => entry.isDirectory() || entry.isFile() && ['.docx', '.xlsx', '.pptx'].includes(extname(entry.name).toLowerCase()))
      .sort((a, b) => a.name.localeCompare(b.name));
    return {
      files: entries.slice(offset, offset + 100).map(entry => ({ path: relative(this.root, resolve(folder, entry.name)), kind: entry.isDirectory() ? 'directory' : extname(entry.name).slice(1).toLowerCase() })),
      nextOffset: offset + 100 < entries.length ? offset + 100 : null,
      open: [...this.documents].map(([document, entry]) => ({ document, path: relative(this.root, entry.path) })),
    };
  }

  async open(path: string) {
    const file = await this.path(path);
    const format = extname(file).toLowerCase();
    if (!['.docx', '.xlsx', '.pptx'].includes(format)) throw new DocumentToolError('UNSUPPORTED_FORMAT', 'Open a DOCX, XLSX, or PPTX file.');
    for (const [id, entry] of this.documents) if (entry.path === file) return { document: id, ...entry.document.overview() };
    if (this.documents.size >= 10) throw new DocumentToolError('DOCUMENT_LIMIT', 'Close a document before opening more than 10.');
    const handle = await open(file, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0) | (constants.O_NONBLOCK ?? 0));
    let bytes: Uint8Array;
    try {
      const info = await this.validateHandle(file, handle);
      if (!info.isFile() || info.size > 64 * 1024 * 1024) throw new DocumentToolError('FILE_TOO_LARGE', 'Open an Office file up to 64 MiB.');
      const buffer = new Uint8Array(info.size + 1);
      let length = 0;
      while (length < buffer.length) {
        const { bytesRead } = await handle.read(buffer, length, buffer.length - length, length);
        if (!bytesRead) break;
        length += bytesRead;
      }
      if (length !== info.size) throw new DocumentToolError('FILE_CHANGED', 'The file changed during access. Retry with a stable workspace.');
      bytes = buffer.subarray(0, length);
    } finally { await handle.close(); }
    const options = { name: basename(file) };
    const document = format === '.xlsx' ? await openXlsx(bytes, options) : format === '.pptx' ? await openPptx(bytes, options) : await openDocx(bytes, { ...options, renderer: this.renderer });
    const id = `doc${this.nextId++}`;
    this.documents.set(id, { path: file, document });
    return { document: id, ...document.overview() };
  }

  get(id: string) {
    const entry = this.documents.get(id);
    if (!entry) throw new DocumentToolError('UNKNOWN_DOCUMENT', 'Use office_open first, or office_files to list open document IDs.', { document: id });
    return entry.document;
  }

  async export(id: string, path: string, proposal?: string) {
    const lexical = this.contained(path);
    const document = this.get(id);
    const format = document.overview().format;
    if (extname(lexical).toLowerCase() !== `.${format}`) throw new DocumentToolError('UNSUPPORTED_FORMAT', `Export path must end in .${format}.`);
    const bytes = await document.export(proposal);
    const parent = await realpath(dirname(lexical));
    this.contained(parent);
    const file = resolve(parent, basename(lexical));
    const handle = await open(file, 'wx', 0o600);
    try {
      await this.validateHandle(file, handle);
      await handle.writeFile(bytes);
    } catch (error) {
      throw new DocumentToolError('EXPORT_INCOMPLETE', 'Export failed after creating its destination. Inspect the incomplete file and choose a new filename before retrying.', {
        path: relative(this.root, file), cause: error instanceof Error ? error.message : String(error),
      });
    } finally { await handle.close(); }
    return { path: relative(this.root, file), bytes: bytes.length, ...(proposal ? { proposal, accepted: false } : {}) };
  }

  close(id: string) {
    this.get(id).close();
    this.documents.delete(id);
    return { document: id, closed: true };
  }

  dispose() {
    for (const entry of this.documents.values()) entry.document.close();
    this.documents.clear();
  }

  private contained(path: string) {
    const absolute = resolve(this.root, path);
    const rel = relative(this.root, absolute);
    if (rel === '..' || rel.startsWith(`..${sep}`) || isAbsolute(rel)) {
      throw new DocumentToolError('OUTSIDE_WORKSPACE', 'Choose a path inside the configured workspace root.');
    }
    return absolute;
  }

  private async path(path: string) {
    return this.contained(await realpath(this.contained(path)));
  }

  private async validateHandle(file: string, handle: FileHandle) {
    const canonical = await this.path(file);
    const current = await lstat(canonical);
    const opened = await handle.stat();
    if (current.isSymbolicLink() || current.dev !== opened.dev || current.ino !== opened.ino) {
      throw new DocumentToolError('FILE_CHANGED', 'The file changed during access. Retry with a stable workspace.');
    }
    return opened;
  }
}
