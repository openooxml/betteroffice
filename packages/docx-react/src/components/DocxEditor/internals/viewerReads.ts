import type { ResidentDocumentRead, ResidentEngineWorkerClient } from '@betteroffice/docx/yrs';
import type { ResidentDocumentReadValues } from '@betteroffice/docx/yrs/residentEngineWorkerProtocol';

export type ViewerReadOutcome<V> =
  | { status: 'ok'; version: string; value: V }
  | { status: 'superseded' };

type VersionedRead = Extract<ResidentDocumentRead, { expectVersion: string }>;

export async function readAt<K extends VersionedRead['kind']>(
  read: ResidentEngineWorkerClient['documentRead'],
  request: VersionedRead & { kind: K }
): Promise<ViewerReadOutcome<ResidentDocumentReadValues[K]>> {
  try {
    const reply = await read<K>(request);
    return reply.version === request.expectVersion
      ? { status: 'ok', version: reply.version, value: reply.value }
      : { status: 'superseded' };
  } catch {
    return { status: 'superseded' };
  }
}
