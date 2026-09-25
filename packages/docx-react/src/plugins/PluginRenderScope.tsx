import { Component, type ReactNode } from 'react';
import { DocxCommandContext } from '../commands/DocxCommandProvider';
import type { DocxPluginActivation, DocxPluginHost } from './createDocxPluginHost';
import type { DocxPluginErrorPhase } from './types';

interface BoundaryProps {
  onError(error: unknown): void;
  children?: ReactNode;
}

/** Stops a failing contribution at its own subtree and reports it once. */
export class PluginContributionBoundary extends Component<BoundaryProps, { failed: boolean }> {
  override state = { failed: false };

  static getDerivedStateFromError(): { failed: boolean } {
    return { failed: true };
  }

  override componentDidCatch(error: unknown): void {
    this.props.onError(error);
  }

  override render(): ReactNode {
    return this.state.failed ? null : this.props.children;
  }
}

/**
 * Renders one plugin contribution: behind its own error boundary and with the plugin's restricted
 * commands in place of the editor's, so public command hooks inside it act for the plugin.
 */
export function PluginRenderScope({
  host,
  activation,
  phase = 'render',
  children,
}: {
  host: DocxPluginHost;
  activation: DocxPluginActivation;
  phase?: DocxPluginErrorPhase;
  children: ReactNode;
}) {
  return (
    <PluginContributionBoundary onError={(error) => host.fail(activation.pluginId, phase, error)}>
      <DocxCommandContext.Provider value={host.commandStore(activation)}>
        {children}
      </DocxCommandContext.Provider>
    </PluginContributionBoundary>
  );
}
