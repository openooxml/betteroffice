import { createPortal } from 'react-dom';
import type { DocxPluginActivation, DocxPluginHost } from './createDocxPluginHost';
import { PluginRenderScope } from './PluginRenderScope';

/**
 * The unscaled layer managed overlays draw in, above the host's raw overlays. It ignores the
 * pointer; interactive overlay elements opt back in with `pointer-events: auto`.
 */
export function PluginOverlays({
  host,
  activations,
  target,
  layerRef,
}: {
  host: DocxPluginHost;
  activations: readonly DocxPluginActivation[];
  target: HTMLElement | null;
  layerRef: (element: HTMLDivElement | null) => void;
}) {
  if (!target) return null;
  return createPortal(
    <div
      ref={layerRef}
      className="docx-plugin-overlays"
      data-testid="plugin-overlays"
      style={{
        position: 'absolute',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        pointerEvents: 'none',
        overflow: 'visible',
        zIndex: 9,
      }}
    >
      {activations.map((activation) => {
        const Overlay = activation.plugin.overlay;
        const geometry = activation.context.geometry;
        return Overlay && geometry ? (
          <PluginRenderScope key={activation.key} host={host} activation={activation}>
            <Overlay context={activation.context} geometry={geometry} />
          </PluginRenderScope>
        ) : null;
      })}
    </div>,
    target
  );
}
