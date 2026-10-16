import { PLUGIN_CHROME } from '../commands/pluginEvents';
import type { XlsxPluginActivation, XlsxPluginHost } from './createXlsxPluginHost';
import { PluginRenderScope } from './PluginRenderScope';

/**
 * The layer managed overlays draw in, on the grid canvas and clipped to it, below the editor's
 * selection, charts outline and cell editor. It ignores the pointer; interactive overlay elements
 * opt back in with `pointer-events: auto`.
 */
export function PluginOverlays({
  host,
  activations,
  layerRef,
  width,
  height,
}: {
  host: XlsxPluginHost;
  activations: readonly XlsxPluginActivation[];
  layerRef: (element: HTMLDivElement | null) => void;
  width: number;
  height: number;
}) {
  return (
    <div
      ref={layerRef}
      {...PLUGIN_CHROME}
      className="xlsx-plugin-overlays"
      data-testid="plugin-overlays"
      style={{
        position: 'absolute',
        top: 0,
        left: 0,
        width,
        height,
        overflow: 'hidden',
        pointerEvents: 'none',
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
    </div>
  );
}
