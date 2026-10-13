import type { DocxEditRequest } from './edits';
import type { DocxCommandPayloads, ResidentDocxCommandEngine, ResidentDocxCommandState } from './residentCommandProtocol';

export function commandBatchRequest(payload: DocxCommandPayloads['applyEdits']): DocxEditRequest {
  const { request, suggesting } = payload;
  return { ...request, ...(suggesting ? { steps: request.steps.map(step => 'suggest' in step && step.suggest ? step : { ...step, suggest: { author: suggesting.name, date: suggesting.date } }) } : {}) };
}

export function applyCommandBatch(engine: ResidentDocxCommandEngine, state: ResidentDocxCommandState, payload: DocxCommandPayloads['applyEdits']) {
  const request = commandBatchRequest(payload);
  engine.beginUndoCapture();
  const result = engine.proposalEngine.applyEdits(request);
  if (result.ok && result.applied) state.storedMarks = null;
  return result;
}
