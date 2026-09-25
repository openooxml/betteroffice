import { useCallback, useSyncExternalStore } from 'react';
import { xlsxCommandController } from '../commands/createXlsxCommandStore';
import { useXlsxCommands } from '../commands/hooks';
import type { XlsxPluginCommandId } from '../commands/types';
import { ToolbarCommandButton } from '../components/toolbar/ToolbarCommand';
import { ToolbarGroup } from '../components/ui/ToolbarPrimitives';
import { useTranslation } from '../i18n';

const NONE: readonly XlsxPluginCommandId[] = Object.freeze([]);

/**
 * The toolbar commands installed plugins contribute, in plugin and declaration order. The
 * default toolbar includes it; place it in replacement chrome to keep those commands reachable.
 */
export function XlsxPluginToolbar(): React.ReactElement | null {
  const store = useXlsxCommands();
  const controller = xlsxCommandController(store);
  const read = useCallback(() => controller?.pluginToolbar() ?? NONE, [controller]);
  const ids = useSyncExternalStore(store.subscribe, read, read);
  const { t } = useTranslation();
  if (ids.length === 0) return null;
  return (
    <ToolbarGroup label={t('plugins.toolbarGroup')}>
      {ids.map((id) => (
        <ToolbarCommandButton key={id} id={id} />
      ))}
    </ToolbarGroup>
  );
}
