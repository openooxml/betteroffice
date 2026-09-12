import { paintSlide, sizeCanvasForSlide } from '@betteroffice/pptx';
import type { CanvasImageResolver, PresentationHandle } from '@betteroffice/pptx';
import { useEffect, useRef, useState } from 'react';
import type { CSSProperties } from 'react';

export function PresentationOverlay({
  handle,
  slideCount,
  startIndex,
  resolveImage,
  label,
  counterLabel,
  exitLabel,
  previousLabel,
  nextLabel,
  onExit,
  onError,
}: {
  handle: PresentationHandle;
  slideCount: number;
  startIndex: number;
  resolveImage: CanvasImageResolver;
  label: string;
  counterLabel: (current: number, total: number) => string;
  exitLabel: string;
  previousLabel: string;
  nextLabel: string;
  onExit: () => void;
  onError: (error: unknown) => void;
}) {
  const containerRef = useRef<HTMLDivElement>(null);
  const dialogRef = useRef<HTMLDialogElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const callbacks = useRef({ onExit, onError });
  callbacks.current = { onExit, onError };
  const [requestedIndex, setIndex] = useState(startIndex);
  const index = Math.max(0, Math.min(requestedIndex, slideCount - 1));
  const [viewport, setViewport] = useState({ width: 0, height: 0 });
  const step = (delta: number) =>
    setIndex(Math.max(0, Math.min(index + delta, slideCount - 1)));

  useEffect(() => {
    setIndex((current) => Math.max(0, Math.min(current, slideCount - 1)));
    if (slideCount === 0) callbacks.current.onExit();
  }, [slideCount]);

  useEffect(() => {
    const dialog = dialogRef.current!;
    const container = containerRef.current!;
    const previousFocus = document.activeElement;
    let active = true;
    dialog.showModal();
    dialog.focus();
    void container
      .requestFullscreen?.()
      .then(() => {
        if (active) {
          dialog.close();
          dialog.showModal();
          dialog.focus();
        } else if (document.fullscreenElement === container) {
          void document.exitFullscreen().catch(() => undefined);
        }
      })
      .catch(() => undefined);
    const onFullscreenChange = () => {
      if (!document.fullscreenElement) callbacks.current.onExit();
    };
    const onResize = () =>
      setViewport({ width: window.innerWidth, height: window.innerHeight });
    onResize();
    window.addEventListener('resize', onResize);
    document.addEventListener('fullscreenchange', onFullscreenChange);
    return () => {
      active = false;
      window.removeEventListener('resize', onResize);
      document.removeEventListener('fullscreenchange', onFullscreenChange);
      if (document.fullscreenElement === container) {
        void document.exitFullscreen().catch(() => undefined);
      }
      dialog.close();
      if (previousFocus instanceof HTMLElement && previousFocus.isConnected)
        previousFocus.focus();
    };
  }, []);

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas || slideCount === 0 || viewport.width <= 0 || viewport.height <= 0)
      return;
    let cancelled = false;
    const paint = async () => {
      const frame = handle.layoutSlide(index);
      const buffer = document.createElement('canvas');
      const context = buffer.getContext('2d');
      if (!context) return;
      const scale = Math.min(
        viewport.width / frame.width,
        viewport.height / frame.height
      );
      const dpr = window.devicePixelRatio || 1;
      sizeCanvasForSlide(buffer, frame, dpr, scale);
      await paintSlide(context, frame, dpr, scale, { resolveImage });
      if (cancelled) return;
      sizeCanvasForSlide(canvas, frame, dpr, scale);
      canvas.getContext('2d')?.drawImage(buffer, 0, 0);
    };
    void paint().catch((error: unknown) => {
      if (!cancelled) {
        callbacks.current.onError(error);
        callbacks.current.onExit();
      }
    });
    return () => {
      cancelled = true;
    };
  }, [handle, index, slideCount, resolveImage, viewport]);

  return (
    <div ref={containerRef}>
      <dialog
        ref={dialogRef}
        aria-label={label}
        tabIndex={-1}
        style={styles.presentationOverlay}
        onCancel={(event) => {
          event.preventDefault();
          onExit();
        }}
        onClick={(event) => {
          if (event.target === dialogRef.current) step(1);
        }}
        onKeyDown={(event) => {
          if (event.altKey || event.ctrlKey || event.metaKey) return;
          switch (event.key) {
            case ' ':
              if (event.target instanceof HTMLElement && event.target.closest('button'))
                return;
              step(1);
              break;
            case 'ArrowRight':
            case 'ArrowDown':
            case 'PageDown':
              step(1);
              break;
            case 'ArrowLeft':
            case 'ArrowUp':
            case 'PageUp':
              step(-1);
              break;
            case 'Home':
              setIndex(0);
              break;
            case 'End':
              setIndex(Math.max(0, slideCount - 1));
              break;
            case 'Escape':
              onExit();
              break;
            default:
              return;
          }
          event.preventDefault();
          event.stopPropagation();
        }}
      >
        <canvas ref={canvasRef} style={styles.presentationCanvas} />
        <button
          type="button"
          onClick={onExit}
          aria-label={exitLabel}
          title={exitLabel}
          style={styles.presentationExit}
        >
          <svg
            width="16"
            height="16"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
            aria-hidden="true"
          >
            <path d="M6 6 18 18M18 6 6 18" />
          </svg>
        </button>
        <div style={styles.presentationControls}>
          <button
            type="button"
            onClick={() => step(-1)}
            disabled={index === 0}
            aria-label={previousLabel}
            style={styles.presentationNavButton}
          >
            <svg
              width="14"
              height="14"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.4"
              aria-hidden="true"
            >
              <path d="m15 6-6 6 6 6" />
            </svg>
          </button>
          <span style={styles.presentationCounter} aria-live="polite">
            {counterLabel(index + 1, slideCount)}
          </span>
          <button
            type="button"
            onClick={() => step(1)}
            disabled={index >= slideCount - 1}
            aria-label={nextLabel}
            style={styles.presentationNavButton}
          >
            <svg
              width="14"
              height="14"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.4"
              aria-hidden="true"
            >
              <path d="m9 6 6 6-6 6" />
            </svg>
          </button>
        </div>
      </dialog>
    </div>
  );
}

const styles = {
  presentationOverlay: {
    position: 'fixed',
    margin: 0,
    padding: 0,
    border: 0,
    width: '100vw',
    height: '100vh',
    maxWidth: 'none',
    maxHeight: 'none',
    inset: 0,
    zIndex: 2147483000,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    background: '#000000',
    cursor: 'pointer',
  },
  presentationCanvas: { display: 'block', pointerEvents: 'none' },
  presentationExit: {
    position: 'absolute',
    top: 16,
    right: 16,
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: 36,
    height: 36,
    border: 0,
    borderRadius: 999,
    background: 'rgba(255, 255, 255, 0.14)',
    color: '#ffffff',
    cursor: 'pointer',
  },
  presentationControls: {
    position: 'absolute',
    bottom: 20,
    left: '50%',
    transform: 'translateX(-50%)',
    display: 'flex',
    alignItems: 'center',
    gap: 12,
    padding: '6px 14px',
    borderRadius: 999,
    background: 'rgba(255, 255, 255, 0.12)',
  },
  presentationNavButton: {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: 28,
    height: 28,
    border: 0,
    borderRadius: 999,
    background: 'rgba(255, 255, 255, 0.14)',
    color: '#ffffff',
    cursor: 'pointer',
  },
  presentationCounter: {
    minWidth: 64,
    color: '#ffffff',
    fontSize: 12,
    fontWeight: 600,
    textAlign: 'center',
  },
} satisfies Record<string, CSSProperties>;
