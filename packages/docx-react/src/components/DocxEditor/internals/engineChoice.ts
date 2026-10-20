import { useEffect, useMemo, useState } from 'react';
import type { DocxEditorProps } from '../../DocxEditor';
import { LEGACY_SYNC_API_PROP } from '../peerless/contracts';

type WorkerCapability =
  | 'Worker'
  | 'OffscreenCanvas'
  | 'HTMLCanvasElement.prototype.transferControlToOffscreen'
  | 'createImageBitmap';
type EngineProps = Pick<
  DocxEditorProps,
  'experimentalWorkerOpen' | 'mediaTokens' | 'collaboration' | 'document' | 'documentBuffer' | 'readOnly' | 'mode'
>;

const warned = new Set<string>();
let missingCapabilitiesOverride: readonly WorkerCapability[] | null = null;

function missingCapabilities(): readonly WorkerCapability[] {
  if (missingCapabilitiesOverride !== null) return missingCapabilitiesOverride;
  const missing: WorkerCapability[] = [];
  if (typeof Worker !== 'function') missing.push('Worker');
  if (typeof OffscreenCanvas !== 'function') missing.push('OffscreenCanvas');
  if (
    typeof HTMLCanvasElement !== 'function' ||
    typeof HTMLCanvasElement.prototype.transferControlToOffscreen !== 'function'
  ) {
    missing.push('HTMLCanvasElement.prototype.transferControlToOffscreen');
  }
  if (typeof createImageBitmap !== 'function') missing.push('createImageBitmap');
  return missing;
}

function chooseEngine(props: EngineProps): { workerOpen: boolean; warnings: string[] } {
  if (props.experimentalWorkerOpen === false) {
    return {
      workerOpen: false,
      warnings: ['[DocxEditor] experimentalWorkerOpen={false} selects the deprecated in-thread engine; omit experimentalWorkerOpen to use the default worker-owned editor.'],
    };
  }
  const warnings: string[] = [];
  const missing = missingCapabilities();
  if (missing.length > 0) {
    warnings.push(capabilityWarning(missing));
  }
  if (props.mediaTokens === true) {
    warnings.push('[DocxEditor] mediaTokens requires the in-thread engine; omit mediaTokens to use the worker-owned editor.');
  }
  if (missing.length === 0 && props.collaboration?.initialUpdate !== undefined) {
    warnings.push('[DocxEditor] collaboration.initialUpdate requires the in-thread engine; omit collaboration.initialUpdate to open the source in the worker-owned editor.');
  }
  if (props.document && !props.documentBuffer && !props.readOnly && props.mode !== 'viewing') {
    warnings.push('[DocxEditor] An editable document source requires the in-thread engine; pass documentBuffer to use the worker-owned editor.');
  }
  return { workerOpen: warnings.length === 0, warnings };
}

function capabilityWarning(missing: readonly WorkerCapability[]) {
  return `[DocxEditor] Using the in-thread engine because these browser features are unavailable: ${missing.join(', ')}; use a browser with these features for the worker-owned editor.`;
}

type RouteProps = Pick<DocxEditorProps, typeof LEGACY_SYNC_API_PROP | 'experimentalWorkerOpen' | 'collaboration' | 'onChange'>;
type EditorRoute = { route: 'hybrid' | 'in-thread' | 'capability-fallback' | 'sync-api-error' | 'peerless'; warning?: string };

export function chooseDocxEditorRoute(props: RouteProps): EditorRoute {
  if (props[LEGACY_SYNC_API_PROP] === true) {
    return { route: 'hybrid', warning: `[DocxEditor] ${LEGACY_SYNC_API_PROP}={true} selects the deprecated hybrid synchronous API, removed in 0.6.0.` };
  }
  if (props.experimentalWorkerOpen === false) return { route: 'in-thread' };
  if (missingCapabilities().length) return { route: 'capability-fallback' };
  if (props.onChange !== undefined) return { route: 'sync-api-error' };
  return { route: 'peerless' };
}

export function useDocxEditorRoute(props: RouteProps): EditorRoute {
  const legacySync = props[LEGACY_SYNC_API_PROP];
  const inThread = props.experimentalWorkerOpen === false;
  const onChange = props.onChange !== undefined;
  const choice = useMemo(() => chooseDocxEditorRoute(props), [legacySync, inThread, onChange]);
  useEffect(() => {
    const warning = choice.warning;
    if (!warning || warned.has(warning)) return;
    warned.add(warning);
    console.warn(warning);
  }, [choice]);
  return choice;
}

export function useEditorEngineChoice(props: EngineProps): boolean {
  const [choice] = useState(() => chooseEngine(props));
  const initialUpdate = props.collaboration?.initialUpdate;
  useEffect(() => {
    const warnings = [...choice.warnings];
    if (choice.workerOpen) {
      if (props.mediaTokens === true) {
        warnings.push('[DocxEditor] mediaTokens changes take effect only on remount; this editor keeps the worker-owned engine.');
      }
      if (initialUpdate !== undefined) {
        warnings.push('[DocxEditor] collaboration.initialUpdate changes take effect only on remount; this editor keeps the worker-owned engine.');
      }
    }
    for (const warning of warnings) {
      if (warned.has(warning)) continue;
      warned.add(warning);
      console.warn(warning);
    }
  }, [choice, props.mediaTokens, initialUpdate]);
  return choice.workerOpen;
}

export function setMissingWorkerCapabilitiesForTests(missing: readonly WorkerCapability[] | null): void {
  missingCapabilitiesOverride = missing;
}

export function resetEngineChoiceForTests(): void {
  missingCapabilitiesOverride = null;
  warned.clear();
}
