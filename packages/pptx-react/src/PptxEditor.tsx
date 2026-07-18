import {
  initWasm,
  openPresentation,
  paintSlide,
  sizeCanvasForSlide,
} from '@betteroffice/pptx';
import type {
  CanvasImageResolver,
  DeckSnapshot,
  PptxFontFace,
  PresentationHandle,
  SlideDisplayList,
  StorySnapshot,
  TextBoxPrimitive,
  TextStyle,
  TextStylePatch,
} from '@betteroffice/pptx';
import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import type {
  CSSProperties,
  ChangeEvent,
  KeyboardEvent,
  PointerEvent,
} from 'react';

export interface PptxTextSelection {
  shapeId: string;
  storyId: string;
  anchor: number;
  focus: number;
}

export interface PptxEditorApi {
  handle: PresentationHandle;
  refresh: () => void;
}

export interface PptxEditorProps {
  file?: Uint8Array;
  fonts: ReadonlyArray<PptxFontFace>;
  clientId?: number;
  className?: string;
  onReady?: (api: PptxEditorApi) => void;
  onChange?: (snapshot: DeckSnapshot) => void;
  onError?: (error: Error) => void;
}

interface EditorModel {
  snapshot: DeckSnapshot;
  slideIndex: number;
  frame: SlideDisplayList | null;
  thumbnails: Map<string, SlideDisplayList>;
}

const initialStyle: Required<Pick<TextStyle, 'bold' | 'italic' | 'fontSizePt' | 'color'>> = {
  bold: false,
  italic: false,
  fontSizePt: 24,
  color: '#111827',
};

export function PptxEditor({
  file,
  fonts,
  clientId,
  className,
  onReady,
  onChange,
  onError,
}: PptxEditorProps) {
  const handleRef = useRef<PresentationHandle | null>(null);
  const modelRef = useRef<EditorModel | null>(null);
  const onReadyRef = useRef(onReady);
  const onChangeRef = useRef(onChange);
  const onErrorRef = useRef(onError);
  const stageRef = useRef<HTMLDivElement>(null);
  const canvasHostRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const imageCacheRef = useRef(new Map<string, Promise<CanvasImageSource | null>>());
  const [model, setModel] = useState<EditorModel | null>(null);
  const [selection, setSelection] = useState<PptxTextSelection | null>(null);
  const [textStyle, setTextStyle] = useState(initialStyle);
  const [viewport, setViewport] = useState({ width: 0, height: 0 });
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  onReadyRef.current = onReady;
  onChangeRef.current = onChange;
  onErrorRef.current = onError;
  modelRef.current = model;

  const reportError = useCallback((value: unknown) => {
    const next = value instanceof Error ? value : new Error(String(value));
    setError(next.message);
    onErrorRef.current?.(next);
  }, []);

  const refreshAt = useCallback(
    (requestedIndex?: number, notify = false, refreshAll = false): EditorModel | null => {
      const handle = handleRef.current;
      if (!handle) return null;
      try {
        const snapshot = handle.snapshot();
        const index = clampSlideIndex(
          requestedIndex ?? modelRef.current?.slideIndex ?? 0,
          snapshot.slides.length
        );
        const thumbnails = new Map<string, SlideDisplayList>();
        for (let slideIndex = 0; slideIndex < snapshot.slides.length; slideIndex += 1) {
          const slide = snapshot.slides[slideIndex];
          const cached = modelRef.current?.thumbnails.get(slide.id);
          if (slideIndex !== index && cached && !refreshAll) thumbnails.set(slide.id, cached);
          else if (slideIndex !== index) thumbnails.set(slide.id, handle.layoutSlide(slideIndex));
        }
        const frame = snapshot.slides.length > 0 ? handle.layoutSlide(index) : null;
        if (frame) thumbnails.set(snapshot.slides[index].id, frame);
        const next = { snapshot, slideIndex: index, frame, thumbnails };
        modelRef.current = next;
        setModel(next);
        setError(null);
        if (notify) onChangeRef.current?.(snapshot);
        return next;
      } catch (value) {
        reportError(value);
        return null;
      }
    },
    [reportError]
  );

  const refresh = useCallback(() => {
    refreshAt(undefined, false, true);
  }, [refreshAt]);

  useEffect(() => {
    let disposed = false;
    let handle: PresentationHandle | null = null;
    let browserFaces: FontFace[] = [];
    let unsubscribeUpdates = () => {};
    handleRef.current?.dispose();
    handleRef.current = null;
    modelRef.current = null;
    setModel(null);
    setSelection(null);
    setError(null);
    imageCacheRef.current.clear();
    if (!file) return;
    setLoading(true);
    void Promise.all([initWasm(), installBrowserFonts(fonts)]).then(
      ([, installed]) => {
        browserFaces = installed;
        if (disposed) {
          removeBrowserFonts(browserFaces);
          return;
        }
        try {
          handle = openPresentation(file, { clientId, fonts });
          handleRef.current = handle;
          unsubscribeUpdates = handle.onUpdate((_update, origin) => {
            if (origin === 'remote') refreshAt(undefined, true, true);
          });
          refreshAt(0);
          setLoading(false);
          onReadyRef.current?.({ handle, refresh });
        } catch (value) {
          setLoading(false);
          reportError(value);
        }
      },
      (value: unknown) => {
        if (disposed) return;
        setLoading(false);
        reportError(value);
      }
    );
    return () => {
      disposed = true;
      unsubscribeUpdates();
      handle?.dispose();
      if (handleRef.current === handle) handleRef.current = null;
      removeBrowserFonts(browserFaces);
    };
  }, [clientId, file, fonts, refresh, refreshAt, reportError]);

  useEffect(() => {
    const host = canvasHostRef.current;
    if (!host) return;
    const update = () => setViewport({ width: host.clientWidth, height: host.clientHeight });
    update();
    const observer = new ResizeObserver(update);
    observer.observe(host);
    return () => observer.disconnect();
  }, []);

  const scale = useMemo(() => {
    if (!model?.frame || viewport.width <= 0 || viewport.height <= 0) return 1;
    return Math.min(
      (viewport.width - 40) / model.frame.width,
      (viewport.height - 40) / model.frame.height,
      1
    );
  }, [model?.frame, viewport]);

  useEffect(() => {
    const canvas = canvasRef.current;
    const frame = model?.frame;
    if (!canvas || !frame) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const dpr = window.devicePixelRatio || 1;
    sizeCanvasForSlide(canvas, frame, dpr, scale);
    let cancelled = false;
    void paintSlide(ctx, frame, dpr, scale, {
      resolveImage: (assetId) => resolveImage(assetId, handleRef, imageCacheRef),
    }).then(
      () => {
        if (!cancelled) paintSelection(ctx, frame, selection, dpr, scale);
      },
      (value: unknown) => {
        if (!cancelled) reportError(value);
      }
    );
    return () => {
      cancelled = true;
    };
  }, [model?.frame, reportError, scale, selection]);

  const selectSlide = (index: number) => {
    setSelection(null);
    refreshAt(index);
    stageRef.current?.focus();
  };

  const pointerDown = (event: PointerEvent<HTMLCanvasElement>) => {
    const handle = handleRef.current;
    const current = modelRef.current;
    const canvas = canvasRef.current;
    if (!handle || !current?.frame || !canvas) return;
    const rect = canvas.getBoundingClientRect();
    const x = ((event.clientX - rect.left) * current.frame.width) / rect.width;
    const y = ((event.clientY - rect.top) * current.frame.height) / rect.height;
    try {
      handle.layoutSlide(current.slideIndex);
      const hit = handle.hitTest(x, y);
      if (hit?.kind === 'text') {
        try {
          handle.story(hit.storyId);
          setSelection({
            shapeId: hit.shapeId,
            storyId: hit.storyId,
            anchor: hit.position,
            focus: hit.position,
          });
        } catch {
          setSelection(null);
        }
      } else {
        setSelection(null);
      }
      stageRef.current?.focus();
    } catch (value) {
      reportError(value);
    }
  };

  const commit = (nextSelection: PptxTextSelection | null) => {
    setSelection(nextSelection);
    refreshAt(undefined, true);
  };

  const keyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    const handle = handleRef.current;
    if (!handle || !selection) return;
    if ((event.metaKey || event.ctrlKey) && (event.key === 'b' || event.key === 'B')) {
      event.preventDefault();
      applyFormatting({ bold: !textStyle.bold });
      return;
    }
    if ((event.metaKey || event.ctrlKey) && (event.key === 'i' || event.key === 'I')) {
      event.preventDefault();
      applyFormatting({ italic: !textStyle.italic });
      return;
    }
    const start = Math.min(selection.anchor, selection.focus);
    const end = Math.max(selection.anchor, selection.focus);
    try {
      if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
        event.preventDefault();
        const story = handle.story(selection.storyId);
        const delta = event.key === 'ArrowLeft' ? -1 : 1;
        const focus = Math.max(0, Math.min(story.length, selection.focus + delta));
        setSelection({
          ...selection,
          anchor: event.shiftKey ? selection.anchor : focus,
          focus,
        });
        return;
      }
      if (event.key === 'Backspace') {
        event.preventDefault();
        if (start !== end) {
          handle.deleteText(selection.storyId, start, end);
          commit({ ...selection, anchor: start, focus: start });
          return;
        }
        const previous = previousTextIndex(handle.story(selection.storyId), start);
        if (previous < start) {
          handle.deleteText(selection.storyId, previous, start);
          commit({ ...selection, anchor: previous, focus: previous });
        }
        return;
      }
      if (event.key === 'Delete') {
        event.preventDefault();
        if (start !== end) {
          handle.deleteText(selection.storyId, start, end);
          commit({ ...selection, anchor: start, focus: start });
          return;
        }
        const next = nextTextIndex(handle.story(selection.storyId), end);
        if (next > end) {
          handle.deleteText(selection.storyId, end, next);
          commit({ ...selection, anchor: end, focus: end });
        }
        return;
      }
      if (event.key === 'Enter') {
        event.preventDefault();
        if (start !== end) handle.deleteText(selection.storyId, start, end);
        handle.insertParagraphBreak(selection.storyId, start);
        commit({ ...selection, anchor: start + 1, focus: start + 1 });
        return;
      }
      if (
        event.key.length > 0 &&
        !event.metaKey &&
        !event.ctrlKey &&
        !event.altKey &&
        Array.from(event.key).length === 1
      ) {
        event.preventDefault();
        if (start !== end) handle.deleteText(selection.storyId, start, end);
        handle.insertText(selection.storyId, start, event.key, textStyle);
        const next = start + event.key.length;
        commit({ ...selection, anchor: next, focus: next });
      }
    } catch (value) {
      reportError(value);
    }
  };

  const applyFormatting = (patch: TextStylePatch) => {
    setTextStyle((current) => ({ ...current, ...patch }));
    const handle = handleRef.current;
    if (!handle || !selection || selection.anchor === selection.focus) return;
    try {
      handle.formatText(
        selection.storyId,
        Math.min(selection.anchor, selection.focus),
        Math.max(selection.anchor, selection.focus),
        patch
      );
      refreshAt(undefined, true);
    } catch (value) {
      reportError(value);
    }
  };

  const addSlide = () => {
    const handle = handleRef.current;
    const current = modelRef.current;
    if (!handle || !current) return;
    try {
      const index = current.slideIndex + 1;
      const layout = current.snapshot.slides[current.slideIndex]?.layoutPartPath ?? undefined;
      handle.insertSlide(index, layout);
      setSelection(null);
      refreshAt(index, true, true);
    } catch (value) {
      reportError(value);
    }
  };

  const deleteSlide = () => {
    const handle = handleRef.current;
    const current = modelRef.current;
    if (!handle || !current || current.snapshot.slides.length <= 1) return;
    try {
      handle.deleteSlide(current.snapshot.slides[current.slideIndex].id);
      setSelection(null);
      refreshAt(Math.max(0, current.slideIndex - 1), true, true);
    } catch (value) {
      reportError(value);
    }
  };

  const addTextBox = () => {
    const handle = handleRef.current;
    const current = modelRef.current;
    if (!handle || !current) return;
    const slide = current.snapshot.slides[current.slideIndex];
    if (!slide) return;
    try {
      const receipt = handle.addTextBox(slide.id, {
        name: 'Text Box',
        rect: {
          x: Math.round(current.snapshot.widthEmu * 0.15),
          y: Math.round(current.snapshot.heightEmu * 0.25),
          width: Math.round(current.snapshot.widthEmu * 0.7),
          height: Math.round(current.snapshot.heightEmu * 0.3),
        },
        text: '',
        style: textStyle,
      });
      const next = refreshAt(undefined, true);
      const shape = next?.snapshot.slides[next.slideIndex]?.shapes.find(
        (candidate) => candidate.id === receipt.shapeId
      );
      const story = shape?.textStories[0];
      if (story) {
        setSelection({ shapeId: shape.id, storyId: story.id, anchor: 0, focus: 0 });
        stageRef.current?.focus();
      }
    } catch (value) {
      reportError(value);
    }
  };

  const history = (direction: 'undo' | 'redo') => {
    const handle = handleRef.current;
    if (!handle) return;
    try {
      if (direction === 'undo') handle.undo();
      else handle.redo();
      setSelection(null);
      refreshAt(undefined, true);
    } catch (value) {
      reportError(value);
    }
  };

  const slideCount = model?.snapshot.slides.length ?? 0;
  const currentSlide = model?.slideIndex ?? 0;

  return (
    <div className={className} style={styles.root}>
      <div style={styles.toolbar} aria-label="Presentation formatting toolbar">
        <button
          type="button"
          style={toolbarButton(handleRef.current?.canUndo() ?? false)}
          disabled={!handleRef.current?.canUndo()}
          onClick={() => history('undo')}
        >
          Undo
        </button>
        <button
          type="button"
          style={toolbarButton(handleRef.current?.canRedo() ?? false)}
          disabled={!handleRef.current?.canRedo()}
          onClick={() => history('redo')}
        >
          Redo
        </button>
        <span style={styles.divider} />
        <button
          type="button"
          aria-pressed={textStyle.bold}
          style={formatButton(textStyle.bold)}
          onClick={() => applyFormatting({ bold: !textStyle.bold })}
        >
          B
        </button>
        <button
          type="button"
          aria-pressed={textStyle.italic}
          style={{ ...formatButton(textStyle.italic), fontStyle: 'italic' }}
          onClick={() => applyFormatting({ italic: !textStyle.italic })}
        >
          I
        </button>
        <label style={styles.fieldLabel}>
          Size
          <input
            aria-label="Font size"
            type="number"
            min={6}
            max={400}
            value={textStyle.fontSizePt}
            style={styles.numberInput}
            onChange={(event: ChangeEvent<HTMLInputElement>) => {
              const value = Number(event.target.value);
              if (Number.isFinite(value)) applyFormatting({ fontSizePt: value });
            }}
          />
        </label>
        <label style={styles.fieldLabel}>
          Color
          <input
            aria-label="Text color"
            type="color"
            value={textStyle.color}
            style={styles.colorInput}
            onChange={(event: ChangeEvent<HTMLInputElement>) =>
              applyFormatting({ color: event.target.value })
            }
          />
        </label>
        <span style={styles.divider} />
        <button type="button" style={toolbarButton(true)} onClick={addSlide}>
          Add slide
        </button>
        <button
          type="button"
          style={toolbarButton(slideCount > 1)}
          disabled={slideCount <= 1}
          onClick={deleteSlide}
        >
          Delete slide
        </button>
        <button type="button" style={toolbarButton(slideCount > 0)} disabled={slideCount === 0} onClick={addTextBox}>
          Add text box
        </button>
      </div>
      <div style={styles.workspace}>
        <aside style={styles.slideStrip} aria-label="Slides">
          {model?.snapshot.slides.map((slide, index) => (
            <button
              type="button"
              key={slide.id}
              aria-current={index === currentSlide ? 'page' : undefined}
              style={slideButton(index === currentSlide)}
              onClick={() => selectSlide(index)}
            >
              <span style={styles.slideNumber}>{index + 1}</span>
              <span style={styles.slidePreview}>
                {model.thumbnails.get(slide.id) ? (
                  <SlideThumbnail
                    frame={model.thumbnails.get(slide.id)!}
                    resolveImage={(assetId) => resolveImage(assetId, handleRef, imageCacheRef)}
                  />
                ) : (
                  <span style={styles.slideTitle}>{slideTitle(slide.shapes) || `Slide ${index + 1}`}</span>
                )}
              </span>
            </button>
          ))}
        </aside>
        <div
          ref={stageRef}
          style={styles.stage}
          tabIndex={0}
          role="application"
          aria-label="Presentation slide editor"
          onKeyDown={keyDown}
        >
          <div ref={canvasHostRef} style={styles.canvasHost}>
            {model?.frame ? (
              <canvas
                ref={canvasRef}
                style={styles.canvas}
                onPointerDown={pointerDown}
                aria-label={`Slide ${currentSlide + 1} of ${slideCount}`}
              />
            ) : (
              <div style={styles.empty}>
                {loading ? 'Opening presentation…' : file ? 'No slides' : 'Open a PPTX file'}
              </div>
            )}
          </div>
          {error ? <div style={styles.error}>{error}</div> : null}
        </div>
      </div>
    </div>
  );
}

function SlideThumbnail({
  frame,
  resolveImage,
}: {
  frame: SlideDisplayList;
  resolveImage: CanvasImageResolver;
}) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const scale = 128 / frame.width;
    const dpr = window.devicePixelRatio || 1;
    sizeCanvasForSlide(canvas, frame, dpr, scale);
    void paintSlide(ctx, frame, dpr, scale, { resolveImage }).catch(() => undefined);
  }, [frame, resolveImage]);
  return <canvas ref={canvasRef} style={styles.thumbnailCanvas} aria-hidden="true" />;
}

function clampSlideIndex(index: number, count: number): number {
  if (count === 0) return 0;
  return Math.max(0, Math.min(count - 1, index));
}

function storyText(story: StorySnapshot): string {
  return story.paragraphs
    .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
    .join('\n');
}

function previousTextIndex(story: StorySnapshot, index: number): number {
  const prefix = storyText(story).slice(0, index);
  const characters = Array.from(prefix);
  const character = characters[characters.length - 1];
  return character ? index - character.length : index;
}

function nextTextIndex(story: StorySnapshot, index: number): number {
  const character = Array.from(storyText(story).slice(index))[0];
  return character ? index + character.length : index;
}

async function installBrowserFonts(fonts: ReadonlyArray<PptxFontFace>): Promise<FontFace[]> {
  if (typeof FontFace === 'undefined' || typeof document === 'undefined') return [];
  const installed = await Promise.all(
    fonts.map(async (font) => {
      const source = font.bytes.slice().buffer as ArrayBuffer;
      const face = new FontFace(font.family, source, {
        style: font.italic ? 'italic' : 'normal',
        weight: font.bold ? '700' : '400',
      });
      const loaded = await face.load();
      document.fonts.add(loaded);
      return loaded;
    })
  );
  return installed;
}

function removeBrowserFonts(fonts: FontFace[]): void {
  if (typeof document === 'undefined') return;
  for (const font of fonts) document.fonts.delete(font);
}

function resolveImage(
  assetId: string,
  handleRef: { current: PresentationHandle | null },
  cacheRef: { current: Map<string, Promise<CanvasImageSource | null>> }
): Promise<CanvasImageSource | null> {
  const cached = cacheRef.current.get(assetId);
  if (cached) return cached;
  const pending = decodeImage(handleRef.current?.mediaBytes(assetId));
  cacheRef.current.set(assetId, pending);
  return pending;
}

async function decodeImage(bytes: Uint8Array | undefined): Promise<CanvasImageSource | null> {
  if (!bytes) return null;
  const blob = new Blob([bytes.slice()]);
  if (typeof createImageBitmap === 'function') return createImageBitmap(blob);
  const url = URL.createObjectURL(blob);
  try {
    return await new Promise<HTMLImageElement>((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = () => reject(new Error('could not decode slide image'));
      image.src = url;
    });
  } finally {
    URL.revokeObjectURL(url);
  }
}

function paintSelection(
  ctx: CanvasRenderingContext2D,
  frame: SlideDisplayList,
  selection: PptxTextSelection | null,
  dpr: number,
  scale: number
): void {
  if (!selection) return;
  const textBox = frame.primitives.find(
    (primitive): primitive is TextBoxPrimitive =>
      primitive.kind === 'textBox' &&
      primitive.storyId === selection.storyId &&
      primitive.shapeId === selection.shapeId
  );
  if (!textBox) return;
  const start = Math.min(selection.anchor, selection.focus);
  const end = Math.max(selection.anchor, selection.focus);
  ctx.save();
  ctx.setTransform(dpr * scale, 0, 0, dpr * scale, 0, 0);
  if (start !== end) {
    ctx.fillStyle = 'rgba(59, 130, 246, 0.26)';
    for (const line of textBox.lines) {
      const lineStart = Math.max(start, line.start);
      const lineEnd = Math.min(end, line.end);
      if (lineStart >= lineEnd) continue;
      const x1 = caretX(line, lineStart);
      const x2 = caretX(line, lineEnd);
      ctx.fillRect(Math.min(x1, x2), line.y, Math.max(1, Math.abs(x2 - x1)), line.height);
    }
  } else {
    const line =
      textBox.lines.find((candidate) => start >= candidate.start && start <= candidate.end) ??
      textBox.lines[textBox.lines.length - 1];
    if (line) {
      const x = caretX(line, start);
      ctx.fillStyle = '#1d4ed8';
      ctx.fillRect(x, line.y, 1.5, line.height);
    }
  }
  ctx.restore();
}

function caretX(line: TextBoxPrimitive['lines'][number], position: number): number {
  const first = line.caretStops[0];
  if (!first) return line.x;
  return line.caretStops.reduce(
    (nearest, stop) =>
      Math.abs(stop.position - position) < Math.abs(nearest.position - position) ? stop : nearest,
    first
  ).x;
}

function slideTitle(shapes: DeckSnapshot['slides'][number]['shapes']): string {
  for (const shape of shapes) {
    for (const story of shape.textStories) {
      const value = storyText(story).trim();
      if (value) return value.slice(0, 40);
    }
  }
  return '';
}

const styles: Record<string, CSSProperties> = {
  root: {
    display: 'flex',
    flexDirection: 'column',
    width: '100%',
    height: '100%',
    minHeight: 480,
    overflow: 'hidden',
    color: '#172033',
    background: '#f3f5f8',
    fontFamily: 'Inter, ui-sans-serif, system-ui, sans-serif',
  },
  toolbar: {
    display: 'flex',
    alignItems: 'center',
    gap: 6,
    minHeight: 48,
    padding: '7px 12px',
    background: '#ffffff',
    borderBottom: '1px solid #d8dee9',
    boxSizing: 'border-box',
    flexWrap: 'wrap',
  },
  divider: { width: 1, alignSelf: 'stretch', background: '#e0e5ed', margin: '0 4px' },
  fieldLabel: { display: 'inline-flex', alignItems: 'center', gap: 5, fontSize: 12 },
  numberInput: { width: 54, height: 30, border: '1px solid #ccd3df', borderRadius: 5, padding: '0 5px' },
  colorInput: { width: 32, height: 30, padding: 2, border: '1px solid #ccd3df', borderRadius: 5, background: '#fff' },
  workspace: { display: 'flex', flex: 1, minHeight: 0 },
  slideStrip: {
    width: 184,
    padding: '14px 10px',
    overflowY: 'auto',
    background: '#eef1f5',
    borderRight: '1px solid #d8dee9',
    boxSizing: 'border-box',
  },
  slideNumber: { width: 18, flex: '0 0 auto', paddingTop: 3, fontSize: 11, color: '#647087', textAlign: 'right' },
  slidePreview: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: 132,
    aspectRatio: '16 / 9',
    padding: 8,
    overflow: 'hidden',
    background: '#ffffff',
    boxSizing: 'border-box',
    boxShadow: '0 1px 4px rgba(20, 31, 50, 0.16)',
  },
  slideTitle: { fontSize: 9, lineHeight: 1.25, color: '#39445a', textAlign: 'center' },
  thumbnailCanvas: { display: 'block', maxWidth: '100%', height: 'auto' },
  stage: { position: 'relative', flex: 1, minWidth: 0, outline: 'none', overflow: 'hidden' },
  canvasHost: { display: 'flex', alignItems: 'center', justifyContent: 'center', width: '100%', height: '100%', overflow: 'auto' },
  canvas: { display: 'block', flex: '0 0 auto', background: '#fff', boxShadow: '0 8px 32px rgba(27, 39, 61, 0.2)', touchAction: 'none' },
  empty: { margin: 'auto', color: '#6b7587', fontSize: 14 },
  error: { position: 'absolute', left: 16, right: 16, bottom: 14, padding: '9px 12px', color: '#8b1e2d', background: '#fff0f2', border: '1px solid #efb8c0', borderRadius: 6, fontSize: 12 },
};

function toolbarButton(enabled: boolean): CSSProperties {
  return {
    height: 32,
    padding: '0 10px',
    border: '1px solid #ccd3df',
    borderRadius: 5,
    color: enabled ? '#273248' : '#9aa3b1',
    background: '#ffffff',
    cursor: enabled ? 'pointer' : 'default',
    fontSize: 12,
  };
}

function formatButton(active: boolean): CSSProperties {
  return {
    ...toolbarButton(true),
    width: 34,
    padding: 0,
    fontWeight: 700,
    color: active ? '#ffffff' : '#273248',
    background: active ? '#325ee6' : '#ffffff',
    borderColor: active ? '#325ee6' : '#ccd3df',
  };
}

function slideButton(active: boolean): CSSProperties {
  return {
    display: 'flex',
    alignItems: 'flex-start',
    gap: 7,
    width: '100%',
    marginBottom: 12,
    padding: 4,
    border: active ? '2px solid #325ee6' : '2px solid transparent',
    borderRadius: 5,
    background: 'transparent',
    cursor: 'pointer',
  };
}
