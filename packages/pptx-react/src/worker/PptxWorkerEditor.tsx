import type { PptxInThreadEditorProps } from '../PptxEditor';
import { LEGACY_SYNC_API } from '../runtimeOptions';
import { PptxEditorContent } from '../PptxEditor';
import type { PptxWorkerEditorApi } from './createWorkerEditorApi';
import { useEditableSessionPresentation } from './useEditableSessionPresentation';

export type PptxWorkerEditorProps = Omit<PptxInThreadEditorProps, 'onReady' | 'readOnly' | 'experimentalWorkerOpen' | 'collaboration'> & {
  /** @deprecated Removed in 0.6.0. Omit for the async editor. */
  [LEGACY_SYNC_API]: true;
  collaboration?: undefined;
  /** Keeps worker opening enabled in the legacy runtime. */
  experimentalWorkerOpen?: true;
  readOnly?: false;
  /** Without `Worker`, falls back to the in-thread engine and passes `PptxEditorApi`. */
  onReady?: (api: PptxWorkerEditorApi) => void;
};

export function PptxWorkerEditor(props: PptxWorkerEditorProps) {
  const backend = useEditableSessionPresentation(props);
  const { onReady, ...content } = props;
  return <PptxEditorContent {...content} backend={backend} onWorkerReady={onReady} />;
}
