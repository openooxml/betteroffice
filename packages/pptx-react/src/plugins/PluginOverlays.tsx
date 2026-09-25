import { PLUGIN_CHROME } from '../commands/pluginEvents';
import type { PptxPluginActivation, PptxPluginHost } from './createPptxPluginHost';
import { PluginRenderScope } from './PluginRenderScope';

/**
 * The unscaled layer managed overlays draw in, aligned with the slide canvas and below the
 * editor's selection, handles and proposal controls. It ignores the pointer; interactive overlay
 * elements opt back in with `pointer-events: auto`.
 */
export function PluginOverlays({
  host,
  activations,
  layerRef,
}: {
  host: PptxPluginHost;
  activations: readonly PptxPluginActivation[];
  layerRef: (element: HTMLDivElement | null) => void;
}) {
  return (
    <div
      ref={layerRef}
      {...PLUGIN_CHROME}
      className="pptx-plugin-overlays"
      data-testid="plugin-overlays"
      style={{ position: 'absolute', inset: 0, pointerEvents: 'none', overflow: 'visible' }}
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
    </div>
  );
}
