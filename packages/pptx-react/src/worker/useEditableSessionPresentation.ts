import { createPptxWorkerEditorSession, PptxWorkerEditorDisposedError } from '@betteroffice/pptx';
import type {
  DeckSnapshot, PptxWorkerEditorAccess, PptxWorkerEditorFrame, PptxWorkerEditorSession,
  SlideDisplayList,
} from '@betteroffice/pptx';
import { useCallback, useRef } from 'react';
import type { PptxWorkerEditorProps } from './PptxWorkerEditor';
import type { WorkerInputCoordinator } from '../commands/workerInputCoordinator';
import { frameImages, installFonts, useStableFonts } from '../viewer/sessionPaint';
import { currentPresentationFrame } from '../viewer/useSessionPresentation';

export const workerEditorSessionOpener = { open: createPptxWorkerEditorSession };

export class EditablePresentation {
  readonly owner: PptxWorkerEditorSession;
  readonly images = frameImages();
  access: PptxWorkerEditorAccess | null = null;
  active = 0;
  navigation = 0;
  alive = true;
  onFirstPaint: (() => void) | undefined;
  private ready = false;
  private readonly cache = new Map<string, PptxWorkerEditorFrame>();
  private readonly lists = new WeakMap<SlideDisplayList, PptxWorkerEditorFrame>();
  private readonly waiters = new Set<{ slide: number; resolve(value: boolean): void; reject(error: unknown): void }>();
  private painted: PptxWorkerEditorFrame | undefined;
  private requesting = false;
  private requested = false;
  private readonly off: () => void;
  private readonly fonts: FontFace[] = [];
  private fontsReady: Promise<void>;

  constructor(props: PptxWorkerEditorProps, private readonly changed: () => void,
    private readonly isCurrent: () => boolean = () => true) {
    this.active = Number.isInteger(props.initialSlide) ? Math.max(0, props.initialSlide! - 1) : 0;
    this.owner = workerEditorSessionOpener.open(props.file!, {
      fonts: props.fonts, clientId: props.clientId, collaboration: props.collaboration,
      onError: (error) => { if (this.current) props.onError?.(error); },
    });
    this.off = this.owner.subscribe(() => {
      if (!this.current) return;
      if (this.owner.failure) this.finish(this.owner.failure);
      changed();
      void this.request();
    });
    this.fontsReady = installFonts(props.fonts, this.fonts, () => this.current);
    void this.fontsReady.catch(() => {});
  }

  get current(): boolean { return this.alive && this.isCurrent(); }
  get hydrated(): boolean { return this.current && this.owner.hydrated && this.access !== null; }
  get overlays(): boolean {
    const frame = this.frame(this.active);
    return this.hydrated && !!frame && this.painted === frame && this.matches(frame);
  }
  require(): void {
    if (!this.current) throw new PptxWorkerEditorDisposedError();
    if (this.owner.failure) throw this.owner.failure;
  }
  async open(coordinator: WorkerInputCoordinator): Promise<PptxWorkerEditorAccess> {
    const access = await this.owner.handleAsync();
    await this.fontsReady;
    this.require();
    this.access = coordinator.access(access);
    this.changed();
    void this.request();
    return this.access;
  }
  snapshot(): DeckSnapshot | null {
    if (this.hydrated) return this.access!.snapshot();
    const projection = this.owner.state.projection;
    return projection ? {
      widthEmu: projection.size.width, heightEmu: projection.size.height,
      slides: projection.slides.map((slide) => ({ ...slide, shapes: [], sourcePartPath: null })),
    } : null;
  }
  frame(index: number): PptxWorkerEditorFrame | undefined {
    const id = this.owner.state.projection?.slides[index]?.id;
    return id ? this.cache.get(id) : undefined;
  }
  matches(frame: PptxWorkerEditorFrame): boolean {
    const state = this.owner.state;
    return this.current && !state.failure && frame.sequence === state.sequence && frame.version === state.version &&
      state.projection?.slides[frame.slideIndex]?.id === frame.slideId && this.cache.get(frame.slideId) === frame;
  }
  currentPaint(frame: PptxWorkerEditorFrame, navigation: number): boolean {
    return this.matches(frame) && navigation === this.navigation &&
      currentPresentationFrame(frame, this.active, this.frame(this.active));
  }
  resolveImage = async (id: string) => {
    const frame = this.frame(this.active);
    if (!frame) return null;
    const resolve = this.images.resolve(frame);
    try { return await resolve(id); } finally { resolve.release(); }
  };
  thumbnailImages(list: SlideDisplayList) {
    const frame = this.lists.get(list);
    if (!frame || !this.matches(frame)) return null;
    return this.images.resolve(frame);
  }
  isThumbnailCurrent(list: SlideDisplayList): boolean {
    const frame = this.lists.get(list);
    return !!frame && this.matches(frame);
  }
  show(slide: number): boolean {
    this.require();
    const count = this.access && this.owner.hydrated ? this.access.snapshot().slides.length :
      this.owner.state.projection?.slides.length ?? 0;
    if (!Number.isInteger(slide) || slide < 1 || slide > count) return false;
    if (this.active !== slide - 1) {
      this.finish();
      this.active = slide - 1;
      this.navigation += 1;
      this.painted = undefined;
      this.changed();
    }
    void this.request();
    return true;
  }
  async showAsync(slide: number): Promise<boolean> {
    if (!this.show(slide)) return false;
    if (this.overlays) return true;
    return new Promise((resolve, reject) => this.waiters.add({ slide, resolve, reject }));
  }
  didPaint(frame: PptxWorkerEditorFrame, navigation: number): void {
    if (!this.currentPaint(frame, navigation)) return;
    this.painted = frame;
    for (const waiter of this.waiters) waiter.resolve(waiter.slide === this.active + 1);
    this.waiters.clear();
    if (this.hydrated && !this.ready && this.onFirstPaint) {
      this.ready = true;
      this.onFirstPaint();
    }
    this.changed();
  }
  private finish(error?: unknown): void {
    for (const waiter of this.waiters) error ? waiter.reject(error) : waiter.resolve(false);
    this.waiters.clear();
  }
  async request(): Promise<void> {
    if (this.requesting) { this.requested = true; return; }
    if (!this.current || this.owner.failure || !this.owner.state.projection ||
        this.owner.state.sequence !== this.owner.state.acknowledgedSequence) return;
    this.requesting = true;
    const navigation = this.navigation;
    const sequence = this.owner.state.sequence;
    try {
      const slides = this.owner.state.projection.slides;
      this.active = Math.min(this.active, Math.max(0, slides.length - 1));
      const order = [this.active, ...slides.map((_, index) => index).filter((index) => index !== this.active)];
      for (const index of order) {
        const slide = slides[index];
        if (!slide || !this.current || navigation !== this.navigation || sequence !== this.owner.state.sequence) break;
        const previous = this.cache.get(slide.id);
        if (previous && this.matches(previous)) continue;
        const frame = await this.owner.frame(slide.id);
        if (!this.current || this.owner.failure || navigation !== this.navigation || sequence !== this.owner.state.sequence) break;
        if (frame.slideId !== slide.id || frame.slideIndex !== index || frame.sequence !== sequence ||
            frame.version !== this.owner.state.version || !Number.isInteger(frame.epoch) ||
            frame.epoch <= (previous?.epoch ?? 0)) continue;
        this.cache.set(slide.id, frame);
        this.lists.set(frame.displayList, frame);
        this.changed();
      }
    } catch (error) {
      if (this.current && !this.owner.failure) this.finish(error);
    } finally {
      this.requesting = false;
      const requested = this.requested;
      this.requested = false;
      if (this.current && (requested || navigation !== this.navigation || sequence !== this.owner.state.sequence)) void this.request();
    }
  }
  dispose(): void {
    this.alive = false;
    this.finish(new PptxWorkerEditorDisposedError());
    this.off();
    this.cache.clear();
    this.painted = undefined;
    this.images.dispose();
    for (const font of this.fonts) document.fonts.delete(font);
    void this.owner.dispose().catch(() => {});
  }
}

export function useEditableSessionPresentation(props: PptxWorkerEditorProps) {
  const fonts = useStableFonts(props.fonts);
  const latest = useRef(props);
  const currentFonts = useRef(fonts);
  latest.current = props;
  currentFonts.current = fonts;
  return useCallback((changed: () => void) => {
    const opened = latest.current;
    return new EditablePresentation({ ...opened, fonts, onError: (error) => latest.current.onError?.(error) }, changed,
      () => latest.current.file === opened.file && currentFonts.current === fonts &&
        latest.current.clientId === opened.clientId && latest.current.collaboration === opened.collaboration);
  },
    [props.file, props.clientId, props.collaboration, fonts]);
}
