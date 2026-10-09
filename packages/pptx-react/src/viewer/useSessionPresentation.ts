import { openPresentationSession } from '@betteroffice/pptx';
import type { PresentationFrame, PresentationSession } from '@betteroffice/pptx';
import { useCallback, useEffect, useRef, useState } from 'react';
import type { PptxWorkerViewerApi, PptxWorkerViewerProps } from '../PptxEditor';
import type { PptxCommandStore } from '../commands/types';
import { frameImages, installFonts, useStableFonts } from './sessionPaint';

export const presentationSessionOpener = { open: openPresentationSession };

export function currentPresentationFrame(
  frame: PresentationFrame | undefined, active: number, expected: PresentationFrame | undefined
): boolean {
  return !!frame && frame === expected && frame.slideIndex === active;
}

export class ViewerSession {
  alive = true;
  failed = false;
  active: number;
  navigation = 0;
  readonly images = frameImages();
  readonly visible = new Set<number>();
  private readonly cache = new Map<string, PresentationFrame>();
  private readonly pending = new Set<number>();
  private readonly waiters = new Set<(painted: boolean) => void>();
  private running = false;
  private painted: PresentationFrame | undefined;
  private ready = false;

  constructor(
    readonly session: PresentationSession,
    initialSlide: number | undefined,
    private readonly changed: () => void,
    private readonly report: (error: unknown) => void,
    private readonly onReady: () => void,
    private readonly isCurrent: () => boolean = () => true
  ) {
    this.active = Math.max(0, Math.min(session.state.slides.length - 1,
      Number.isInteger(initialSlide) ? initialSlide! - 1 : 0));
  }

  get current(): boolean { return this.alive && !this.failed && this.isCurrent(); }

  require(): PresentationSession {
    if (!this.current) throw new Error('Presentation is no longer open');
    return this.session;
  }

  frame(index: number): PresentationFrame | undefined {
    const id = this.session.state.slides[index]?.id;
    for (const [key, frame] of this.cache) {
      if (key === JSON.stringify([id, frame.version])) return frame;
    }
    return undefined;
  }

  start(): void {
    if (this.session.state.slides.length) this.pending.add(this.active);
    void this.drain();
  }

  show(slide: number): boolean {
    if (!this.current || !Number.isInteger(slide) || slide < 1 ||
      slide > this.session.state.slides.length) return false;
    if (this.active !== slide - 1) {
      this.finish(false);
      this.active = slide - 1;
      this.navigation += 1;
      this.painted = undefined;
      const frame = this.frame(this.active);
      if (frame) {
        const key = JSON.stringify([this.session.state.slides[this.active].id, frame.version]);
        this.cache.delete(key);
        this.cache.set(key, frame);
      } else this.pending.add(this.active);
      this.trim();
      this.changed();
    }
    void this.drain();
    return true;
  }

  async showAsync(slide: number): Promise<boolean> {
    if (!this.show(slide)) return false;
    if (this.painted && this.painted === this.frame(this.active)) return true;
    return new Promise((resolve) => this.waiters.add(resolve));
  }

  visibility(index: number, visible: boolean): void {
    if (!this.current) return;
    if (visible) {
      if (!this.visible.has(index) && !this.frame(index)) this.pending.add(index);
      this.visible.add(index);
    } else {
      this.visible.delete(index);
      this.trim();
      if (index !== this.active) this.pending.delete(index);
    }
    this.changed();
    void this.drain();
  }

  didPaint(frame: PresentationFrame): void {
    if (!this.current || !currentPresentationFrame(frame, this.active, this.frame(this.active))) return;
    this.painted = frame;
    this.finish(true);
    if (!this.ready) {
      this.ready = true;
      this.onReady();
    }
    this.changed();
    void this.drain();
  }

  isPainted(frame: PresentationFrame): boolean { return this.painted === frame; }

  fail(error: unknown): void {
    if (!this.current) return;
    this.failed = true;
    this.pending.clear();
    this.finish(false);
    this.report(error);
    this.changed();
  }

  dispose(): void {
    this.alive = false;
    this.finish(false);
    this.cache.clear();
    this.pending.clear();
    this.visible.clear();
    this.painted = undefined;
    this.images.dispose();
    void this.session.dispose().catch(() => {});
  }

  private finish(value: boolean): void {
    for (const resolve of this.waiters) resolve(value);
    this.waiters.clear();
  }

  private trim(): void {
    while (this.cache.size > 25) {
      const oldest = [...this.cache].find(([, cached]) =>
        cached.slideIndex !== this.active && !this.visible.has(cached.slideIndex));
      if (!oldest) break;
      this.cache.delete(oldest[0]);
    }
  }

  private async drain(): Promise<void> {
    if (this.running || !this.current) return;
    this.running = true;
    try {
      while (this.current) {
        const index = this.pending.has(this.active) ? this.active :
          this.painted === this.frame(this.active) ? this.pending.values().next().value : undefined;
        if (index === undefined) break;
        this.pending.delete(index);
        if (index !== this.active && !this.visible.has(index)) continue;
        if (this.frame(index)) continue;
        try {
          const frame = await this.session.call.frame(index);
          if (!this.current) break;
          if (frame.slideIndex !== index) throw new Error('Unexpected slide frame');
          const id = this.session.state.slides[index].id;
          this.cache.set(JSON.stringify([id, frame.version]), frame);
          this.trim();
          this.changed();
        } catch (error) {
          if (!this.current) break;
          if (error instanceof Error && error.name === 'SessionSuperseded') {
            if (index === this.active || this.visible.has(index)) this.pending.add(index);
          } else this.fail(error);
        }
      }
    } finally { this.running = false; }
  }
}

export function useSessionPresentation(
  props: PptxWorkerViewerProps, commands: PptxCommandStore, focus: () => void
) {
  const fonts = useStableFonts(props.fonts);
  const latest = useRef({ props, focus, fonts });
  latest.current = { props, focus, fonts };
  const generation = useRef(0);
  const [run, setRun] = useState<ViewerSession | null>(null);
  const [, setRevision] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [notes, setNotes] = useState<{ id: string; text: string } | null>(null);
  const reportError = useCallback((value: unknown) => {
    const error = value instanceof Error ? value : new Error(String(value));
    setError(error.message);
    latest.current.props.onError?.(error);
  }, []);

  useEffect(() => {
    const token = ++generation.current;
    let disposed = false;
    let openingFailed = false;
    let opened: ViewerSession | undefined;
    let offFailure = () => {};
    const browserFonts: FontFace[] = [];
    const current = () => !disposed && token === generation.current &&
      latest.current.props.file === props.file && latest.current.fonts === fonts &&
      latest.current.props.clientId === props.clientId;
    const changed = () => { if (current()) setRevision((revision) => revision + 1); };
    setRun(null);
    setNotes(null);
    setError(null);
    setLoading(Boolean(props.file));
    if (props.file) void (async () => {
      try {
        const fontsReady = installFonts(fonts, browserFonts,
          () => current() && !openingFailed && (opened?.current ?? true));
        void fontsReady.catch(() => {});
        const session = await presentationSessionOpener.open(props.file!, {
          fonts, clientId: props.clientId,
        });
        if (!current()) { void session.dispose().catch(() => {}); return; }
        opened = new ViewerSession(session, latest.current.props.initialSlide, changed,
          (error) => { if (current()) reportError(error); }, () => {
            if (current()) latest.current.props.onReady?.(api);
          }, current);
        const viewer = opened;
        const require = () => viewer.require();
        const afterSession = async <T,>(operation: (session: PresentationSession) => Promise<T>): Promise<T | null> => {
          if (!viewer.current) return null;
          try {
            const result = await operation(require());
            return viewer.current ? result : null;
          } catch (error) {
            if (!viewer.current) return null;
            throw error;
          }
        };
        const refusal = () => afterSession(async (session) => ({
          ok: false as const, version: await session.call.version(),
          failure: { code: 'read-only' as const, message: 'The editor is read-only' },
        }));
        const api: PptxWorkerViewerApi = {
          handle: null, commands, save: () => null, getPositionAtPoint: () => null,
          selectText: () => false, clearSelection: () => {}, refreshProposals: () => {},
          refresh: () => {}, flushPendingInput: async () => { require(); },
          focus: () => { if (viewer.current) latest.current.focus(); },
          goToSlide: (slide) => {
            const accepted = viewer.show(slide);
            if (accepted) latest.current.focus();
            return accepted;
          },
          goToSlideAsync: async (slide) => {
            const painted = await viewer.showAsync(slide);
            return viewer.current && painted;
          },
          saveAsync: () => afterSession((session) => session.save()),
          version: () => afterSession((session) => session.call.version()),
          readContent: (request) => afterSession((session) => session.call.readContent(request)),
          findText: (request) => afterSession((session) => session.call.findText(request)),
          validateEdits: refusal, applyEdits: refusal,
        };
        offFailure = session.onFailure((failure) => viewer.fail(failure));
        if (session.failure) viewer.fail(session.failure);
        await fontsReady;
        if (!current()) return;
        setRun(viewer);
        setLoading(false);
        viewer.start();
      } catch (error) {
        openingFailed = true;
        if (!current()) return;
        opened?.fail(error);
        if (!opened) {
          for (const font of browserFonts.splice(0)) document.fonts.delete(font);
          reportError(error);
        }
        setLoading(false);
      }
    })();
    return () => {
      disposed = true;
      generation.current += 1;
      offFailure();
      opened?.dispose();
      for (const font of browserFonts) document.fonts.delete(font);
    };
  }, [props.file, props.clientId, fonts, commands, reportError]);

  const slide = run?.current ? run.session.state.slides[run.active] : undefined;
  useEffect(() => {
    setNotes(null);
    if (!run?.current || !slide) return;
    let cancelled = false;
    void run.session.call.readContent({ slideIds: [slide.id] }).then((result) => {
      if (cancelled || !run.current) return;
      if (!result.ok) throw new Error(result.failure.message);
      setNotes({ id: slide.id, text: result.slides[0]?.notes ?? '' });
    }).catch((error) => { if (!cancelled && run.current) reportError(error); });
    return () => { cancelled = true; };
  }, [run, slide?.id, reportError]);

  return { run, error, loading, reportError, notes: notes?.id === slide?.id ? notes?.text ?? '' : '' };
}
