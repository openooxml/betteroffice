import { readdir, readFile, realpath, stat, writeFile } from 'node:fs/promises';
import { basename, dirname, extname, isAbsolute, relative, resolve, sep } from 'node:path';
import { openDocx, type DocxAgentDocument } from './document';
import { DocumentToolError, type DocumentRenderer } from './types';

export class FileWorkspace {
  private readonly documents = new Map<string, { path: string; document: DocxAgentDocument }>();
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
      .filter(entry => entry.isDirectory() || entry.isFile() && extname(entry.name).toLowerCase() === '.docx')
      .sort((a, b) => a.name.localeCompare(b.name));
    return {
      files: entries.slice(offset, offset + 100).map(entry => ({ path: relative(this.root, resolve(folder, entry.name)), kind: entry.isDirectory() ? 'directory' : 'docx' })),
      nextOffset: offset + 100 < entries.length ? offset + 100 : null,
      open: [...this.documents].map(([document, entry]) => ({ document, path: relative(this.root, entry.path) })),
    };
  }

  async open(path: string) {
    const file = await this.path(path);
    if (extname(file).toLowerCase() !== '.docx') throw new DocumentToolError('UNSUPPORTED_FORMAT', 'This release supports DOCX.');
    for (const [id, entry] of this.documents) if (entry.path === file) return { document: id, ...entry.document.overview() };
    if (this.documents.size >= 10) throw new DocumentToolError('DOCUMENT_LIMIT', 'Close a document before opening more than 10.');
    const info = await stat(file);
    if (!info.isFile() || info.size > 64 * 1024 * 1024) throw new DocumentToolError('FILE_TOO_LARGE', 'Open a DOCX file up to 64 MiB.');
    const document = await openDocx(await readFile(file), { name: basename(file), renderer: this.renderer });
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
    const parent = await realpath(dirname(lexical));
    this.contained(parent);
    const file = resolve(parent, basename(lexical));
    if (extname(file).toLowerCase() !== '.docx') throw new DocumentToolError('UNSUPPORTED_FORMAT', 'Export path must end in .docx.');
    const bytes = await this.get(id).export(proposal);
    await writeFile(file, bytes, { flag: 'wx' });
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
}
