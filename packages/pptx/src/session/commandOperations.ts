import type { PptxEditStep } from '../edits';
import type { ProposalEdit } from '../proposals';
import type { PresentationHandle } from '../wasm/loader';
import type { PresentationCommandMutation } from './commandProtocol';
import { findTopLevelShape } from './commandShapes';

export function prepareCommandOperation(handle: PresentationHandle, mutation: PresentationCommandMutation): (() => unknown) | null {
  if (!mutation || !Array.isArray(mutation.args) || mutation.args.length > 5) return null;
  const { method, args } = mutation;
  const snapshot = handle.snapshot();
  const text = (value: unknown, limit = 1e6): value is string => typeof value === 'string' && value.length <= limit &&
    !/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ud800-\udfff\ufffe\uffff]/u.test(value);
  const position = (value: unknown): value is { xEmu: number; yEmu: number } => Boolean(value && typeof value === 'object' &&
    ['xEmu', 'yEmu'].every((key) => Number.isSafeInteger((value as Record<string, unknown>)[key])));
  const commentDraft = (value: unknown) => Boolean(value && typeof value === 'object' &&
    text((value as Record<string, unknown>).author, 1024) && text((value as Record<string, unknown>).text) &&
    (value as { text: string }).text.length > 0 &&
    text((value as Record<string, unknown>).created, 128) && Number.isFinite(Date.parse((value as { created: string }).created)) &&
    ((value as { initials?: string }).initials === undefined || text((value as { initials: string }).initials, 128)));
  if (method === 'insertTableRow' || method === 'deleteTableRow' || method === 'insertTableColumn' || method === 'deleteTableColumn') {
    const insert = method === 'insertTableRow' || method === 'insertTableColumn';
    if (args.length < (insert ? 4 : 3) || args.length > (insert ? 5 : 3)) return null;
    const [slide, shape, index] = args;
    if (!text(slide, 1024) || !text(shape, 1024) || !Number.isSafeInteger(index) || index < 0) return null;
    const admitted = (position?: 'before' | 'after', cells?: readonly string[]) => {
      try { handle.checkTableEdit(method, slide, shape, index, position, cells); return true; }
      catch { return false; }
    };
    if (method === 'insertTableRow' || method === 'insertTableColumn') {
      const position = args[3];
      const cells = args[4];
      if (position !== 'before' && position !== 'after' || cells !== undefined && (!Array.isArray(cells) || cells.length > 1000 || cells.some(cell => !text(cell, 16000)))) return null;
      return admitted(position, cells) ? () => handle[method](slide, shape, index, position, cells) : null;
    }
    return admitted() ? () => handle[method](slide, shape, index) : null;
  }
  if (method === 'setCommentFlavor') return ['legacy', 'modern'].includes(args[0]) && !snapshot.comments?.length ? () => handle.setCommentFlavor(args[0]) : null;
  if (method === 'replyToComment' || method === 'setCommentStatus' || method === 'setCommentPosition' || method === 'removeComment') {
    const comment = snapshot.comments?.find((comment) => comment.id === args[0]);
    if (!comment) return null;
    switch (method) {
      case 'replyToComment': return snapshot.commentFlavor === 'modern' && !comment.parentId && commentDraft(args[1]) ? () => handle.replyToComment(comment.id, args[1]) : null;
      case 'setCommentStatus': return snapshot.commentFlavor === 'modern' && typeof args[1] === 'boolean' ? () => handle.setCommentStatus(comment.id, args[1]) : null;
      case 'setCommentPosition': return !comment.parentId && position(args[1]) ? () => handle.setCommentPosition(comment.id, args[1]) : null;
      case 'removeComment': return () => handle.removeComment(comment.id);
    }
  }
  if (method === 'propose') {
    const [agent, note, edits] = args;
    if (!text(agent, 1024) || !agent.trim() || note !== null && !text(note) ||
      !Array.isArray(edits) || !edits.length || edits.length > 256) return null;
    try { if (JSON.stringify(edits).length > 1e6) return null; }
    catch { return null; }
    if (handle.listProposals().length >= 64) return null;
    const fields = new Map<ProposalEdit['type'], readonly string[]>([
      ['replaceText', ['storyId', 'start', 'end', 'text', 'style']],
      ['formatText', ['storyId', 'start', 'end', 'patch']],
      ['setParagraphAlignment', ['storyId', 'start', 'end', 'alignment']],
      ['setShapeRect', ['slideId', 'shapeId', 'rect']],
      ['setShapeFill', ['slideId', 'shapeId', 'color']],
      ['setShapeStroke', ['slideId', 'shapeId', 'stroke']],
      ['setShapeAdjust', ['slideId', 'shapeId', 'adjustments']],
      ['setSlideNotes', ['slideId', 'text']],
    ]);
    const onlyKeys = (value: unknown, keys: readonly string[]) => Boolean(value && typeof value === 'object' &&
      !Array.isArray(value) && Object.keys(value).every((key) => keys.includes(key)));
    const validValues = (value: unknown, depth = 0): boolean => {
      if (typeof value === 'string') return text(value);
      if (typeof value === 'number') return Number.isFinite(value);
      return !value || typeof value !== 'object' || depth < 16 &&
        Object.entries(value).every(([key, entry]) => text(key) && validValues(entry, depth + 1));
    };
    const style = (value: unknown, extended = false) => {
      const keys = ['bold', 'italic', 'fontFamily', 'underline', 'color', 'fontSizePt', 'spacingPt', 'baselinePct'];
      if (extended) keys.push('kernPt', 'caps');
      if (!onlyKeys(value, keys)) return false;
      return Object.entries(value as Record<string, unknown>).every(([key, entry]) => {
        if (entry === undefined || entry === null) return true;
        switch (key) {
          case 'bold':
          case 'italic': return typeof entry === 'boolean';
          case 'fontFamily': return text(entry);
          case 'underline': return typeof entry === 'string' && ['none', 'words', 'sng', 'dbl', 'heavy', 'dotted',
            'dottedHeavy', 'dash', 'dashHeavy', 'dashLong', 'dashLongHeavy', 'dotDash', 'dotDashHeavy',
            'dotDotDash', 'dotDotDashHeavy', 'wavy', 'wavyHeavy', 'wavyDbl'].includes(entry);
          case 'color': return typeof entry === 'string' && /^#?[0-9a-f]{6}$/i.test(entry);
          case 'fontSizePt': return typeof entry === 'number' && Number.isFinite(entry) && entry >= 1 && entry <= 4000;
          case 'spacingPt': return typeof entry === 'number' && Number.isFinite(entry) && Math.abs(entry) <= 4000;
          case 'baselinePct': return typeof entry === 'number' && Number.isFinite(entry) &&
            entry * 1000 >= -2147483648 && entry * 1000 <= 2147483647;
          case 'kernPt': return typeof entry === 'number' && Number.isFinite(entry) && entry >= 0 && entry <= 4000;
          case 'caps': return typeof entry === 'string' && ['none', 'small', 'all'].includes(entry);
          default: return false;
        }
      });
    };
    const content = handle.readContent();
    if (!content.ok) return null;
    const steps: PptxEditStep[] = [];
    const canonicalEdits: ProposalEdit[] = [];
    for (const candidate of edits) {
      if (!candidate || typeof candidate !== 'object') return null;
      const keys = fields.get(candidate.type);
      if (!keys || !onlyKeys(candidate, ['type', ...keys])) return null;
      const source = candidate as unknown as Record<string, unknown>;
      const edit = Object.fromEntries(['type', ...keys].map((key) => [key, source[key]])) as ProposalEdit;
      if (!validValues(edit) || edit.type === 'replaceText' && edit.style != null && !style(edit.style, true) ||
        edit.type === 'formatText' && !style(edit.patch) ||
        edit.type === 'setShapeRect' && !onlyKeys(edit.rect, ['x', 'y', 'width', 'height']) ||
        edit.type === 'setShapeStroke' && !onlyKeys(edit.stroke, ['color', 'widthPt'])) return null;
      canonicalEdits.push(edit);
      if ('storyId' in edit) {
        const story = content.stories.find((story) => story.storyId === edit.storyId);
        if (!story) return null;
        const target = { kind: 'range' as const, slideId: story.slideId, shapeId: story.shapeId, storyId: story.storyId, start: edit.start, end: edit.end };
        if (edit.type === 'replaceText') steps.push({ op: 'replaceText', target, text: edit.text });
        else if (edit.type === 'formatText') {
          const patch = Object.fromEntries(Object.entries(edit.patch).filter(([, value]) => value != null));
          edit.patch = patch;
          steps.push({ op: 'formatText', target, patch });
        } else if (edit.type === 'setParagraphAlignment') steps.push({ op: 'setParagraphAlignment', target, alignment: edit.alignment as 'l' | 'ctr' | 'r' | 'just' | null });
        else return null;
      } else {
        if (!['setShapeRect', 'setShapeFill', 'setShapeStroke', 'setShapeAdjust', 'setSlideNotes'].includes(edit.type)) return null;
        const { type, ...fields } = edit;
        const { slideId, ...rest } = fields;
        const shapeId = 'shapeId' in rest ? rest.shapeId : undefined;
        if ('shapeId' in rest) delete (rest as { shapeId?: string }).shapeId;
        steps.push({ op: type, target: shapeId === undefined ? { slideId } : { slideId, shapeId }, ...rest } as PptxEditStep);
      }
    }
    const expectVersion = handle.version();
    try { if (!handle.validateEdits({ expectVersion, steps }).ok) return null; }
    catch (error) {
      if (error && typeof error === 'object' && 'name' in error && error.name === 'PptxEditDecodeError') return null;
      throw error;
    }
    return () => handle.propose(agent, note, canonicalEdits);
  }
  if (method === 'acceptProposal' || method === 'rejectProposal') {
    const proposal = handle.listProposals().find((proposal) => proposal.id === args[0]);
    if (!proposal) return null;
    if (method === 'rejectProposal') return () => handle.rejectProposal(proposal.id);
    if (args[1]?.force !== undefined && typeof args[1].force !== 'boolean' || proposal.staleTargets.length && !args[1]?.force) return null;
    return () => handle.acceptProposal(proposal.id, args[1]);
  }
  if (method === 'insertSlide') {
    const [index, layout] = args;
    if (!Number.isSafeInteger(index) || index < 0 || index > snapshot.slides.length ||
      layout !== undefined && (typeof layout !== 'string' || !snapshot.slides.some((slide) => slide.layoutPartPath === layout))) return null;
    return () => handle.insertSlide(index, layout);
  }
  const slide = snapshot.slides.find((slide) => slide.id === args[0]);
  if (!slide) return null;
  if (method === 'setSlideNotes') return text(args[1]) ? () => handle.setSlideNotes(slide.id, args[1]) : null;
  if (method === 'addComment') {
    const draft = args[1];
    return commentDraft(draft) && (draft.xEmu === undefined && draft.yEmu === undefined || position(draft))
      ? () => handle.addComment(slide.id, draft) : null;
  }
  if (method === 'deleteSlide') return () => handle.deleteSlide(slide.id);
  if (method === 'moveSlide') {
    const index = args[1];
    return Number.isSafeInteger(index) && index >= 0 && index < snapshot.slides.length ? () => handle.moveSlide(slide.id, index) : null;
  }
  const shape = findTopLevelShape(slide, args[1]);
  if (!shape || shape.id !== args[1]) return null;
  switch (method) {
    case 'setShapeFill': {
      const color = args[2];
      return shape.kind === 'shape' && (color === null || /^#[0-9a-f]{6}$/i.test(color)) ? () => handle.setShapeFill(slide.id, shape.id, color) : null;
    }
    case 'setShapeStroke': {
      const stroke = args[2];
      return shape.kind === 'shape' && stroke && (stroke.color === undefined || typeof stroke.color === 'string' && /^#[0-9a-f]{6}$/i.test(stroke.color)) &&
        (stroke.widthPt === undefined || Number.isFinite(stroke.widthPt) && stroke.widthPt >= 0 && stroke.widthPt <= 1584)
        ? () => handle.setShapeStroke(slide.id, shape.id, stroke) : null;
    }
    case 'setShapeAdjust': {
      const adjustments = args[2];
      return shape.kind === 'shape' && adjustments && typeof adjustments === 'object' && !Array.isArray(adjustments) &&
        Object.keys(adjustments).length <= 64 && Object.entries(adjustments).every(([key, value]) =>
          /^[a-zA-Z][a-zA-Z0-9]*$/.test(key) && Number.isFinite(value) && Math.abs(value) <= 1e6)
        ? () => handle.setShapeAdjust(slide.id, shape.id, adjustments) : null;
    }
    case 'removeShape': return () => handle.removeShape(slide.id, shape.id);
    case 'bringShapeToFront': return () => handle.bringShapeToFront(slide.id, shape.id);
    case 'sendShapeToBack': return () => handle.sendShapeToBack(slide.id, shape.id);
    case 'bringShapeForward': return () => handle.bringShapeForward(slide.id, shape.id);
    case 'sendShapeBackward': return () => handle.sendShapeBackward(slide.id, shape.id);
    default: return null;
  }
}
