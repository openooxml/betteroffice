import { ResidentWorkerFailureError, type YrsSession } from '@betteroffice/docx/yrs';
import type { WorkerPageExportResult } from './workerOpenExport';

let nextGeneration = 0;

export class WorkerExportVersions {
  private readonly generation = ++nextGeneration;
  private readonly layouts = new Map<string, string>();

  workerLayoutVersion(outward: string): string {
    return this.layouts.get(outward) ?? outward;
  }

  adapt(result: WorkerPageExportResult, peer: string, worker: string): WorkerPageExportResult {
    if (result.version !== worker) {
      throw new ResidentWorkerFailureError('Resident engine worker returned an unexpected export version');
    }
    if (!result.ok) return { ...result, version: peer };
    const layout = result.content.layout;
    if (layout.documentVersion !== worker || !layout.layoutVersion.startsWith(`${worker}:`)) {
      throw new ResidentWorkerFailureError('Resident engine worker returned an unexpected layout version');
    }
    const outward = `${peer}:${this.generation}` + layout.layoutVersion.slice(worker.length);
    const previous = this.layouts.get(outward);
    if (previous !== undefined && previous !== layout.layoutVersion) {
      throw new ResidentWorkerFailureError('Resident engine worker reused an export layout token');
    }
    this.layouts.set(outward, layout.layoutVersion);
    return {
      ...result,
      version: peer,
      content: {
        ...result.content,
        layout: { ...layout, documentVersion: peer, layoutVersion: outward },
      },
    };
  }
}

const versions = new WeakMap<YrsSession, { owner: object; load: number; adapter: WorkerExportVersions }>();

export function clearWorkerExportVersions(peer: YrsSession): void {
  versions.delete(peer);
}

export function workerExportVersions(peer: YrsSession, owner: object, load: number): WorkerExportVersions {
  const current = versions.get(peer);
  if (current?.owner === owner && current.load === load) return current.adapter;
  const adapter = new WorkerExportVersions();
  versions.set(peer, { owner, load, adapter });
  return adapter;
}
