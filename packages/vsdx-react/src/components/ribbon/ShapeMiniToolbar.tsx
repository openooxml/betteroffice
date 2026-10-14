import type { CSSProperties, RefObject } from 'react';
import type { TFunction } from '@betteroffice/vsdx-i18n';
import { RibbonIcon } from './RibbonIcon';
import type { RibbonIconName } from './RibbonIcon';
import { useRibbonCommands } from './commands';
import type { RibbonCommand, RibbonCommandId } from './commands';

export const MINI_TOOLBAR_CANDIDATES: ReadonlyArray<RibbonCommandId> = ['fillColor', 'lineColor'];

const MINI_TOOLBAR_ICONS: Record<string, RibbonIconName> = { fillColor: 'fill', lineColor: 'line' };

export function resolveMiniToolbarIds(commands: Partial<Record<RibbonCommandId, RibbonCommand | undefined>>): RibbonCommandId[] {
  return MINI_TOOLBAR_CANDIDATES.filter((id) => commands[id]?.enabled === true);
}

export interface MiniToolbarRect {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export interface MiniToolbarSize {
  width: number;
  height: number;
}

export interface MiniToolbarViewport {
  width: number;
  height: number;
}

export function miniToolbarPosition(
  menu: MiniToolbarRect,
  toolbar: MiniToolbarSize,
  viewport: MiniToolbarViewport,
  margin = 4,
  gap = 4,
): { top: number; left: number; below: boolean } {
  const left = Math.max(margin, Math.min(menu.left, viewport.width - toolbar.width - margin));
  const above = menu.top - gap - toolbar.height;
  if (above >= margin) return { top: above, left, below: false };
  const belowTop = menu.bottom + gap;
  return { top: Math.max(margin, Math.min(belowTop, viewport.height - toolbar.height - margin)), left, below: true };
}

export interface ShapeMiniToolbarProps {
  t: TFunction;
  toolbarRef: RefObject<HTMLDivElement | null>;
  style?: CSSProperties;
  below?: boolean;
}

export function ShapeMiniToolbar({ t, toolbarRef, style, below = false }: ShapeMiniToolbarProps) {
  const commands = useRibbonCommands();
  const ids = resolveMiniToolbarIds(commands);
  if (ids.length === 0) return null;
  return (
    <div ref={toolbarRef} role="toolbar" aria-label={t('contextMenu.miniToolbarLabel')} data-below={below} style={{ ...styles.toolbar, ...style }}>
      {ids.map((id) => {
        const command = commands[id];
        const label = t(`ribbon.commands.${id}`);
        return (
          <label key={id} title={label} style={{ ...styles.button, position: 'relative' }}>
            <RibbonIcon name={MINI_TOOLBAR_ICONS[id] ?? 'fill'} size={20} />
            <span aria-hidden="true" style={{ position: 'absolute', bottom: 4, width: 16, height: 3, borderRadius: 1, background: command.value ?? '#000000' }} />
            <input type="color" value={command.value ?? '#000000'} aria-label={label} data-command-id={id} onChange={(event) => command.run(event.target.value)} style={styles.colorInput} />
          </label>
        );
      })}
    </div>
  );
}

const styles: Record<string, CSSProperties> = {
  toolbar: { position: 'fixed', display: 'flex', alignItems: 'center', gap: 2, padding: 4, backgroundColor: '#ffffff', border: '1px solid #e0e0e0', borderRadius: 4, boxShadow: '0 4px 12px rgba(0,0,0,0.14)', zIndex: 10000 },
  button: { appearance: 'none', display: 'inline-grid', placeItems: 'center', width: 32, height: 32, padding: 0, border: 0, borderRadius: 4, boxSizing: 'border-box', color: '#242424', cursor: 'pointer' },
  colorInput: { position: 'absolute', inset: 0, opacity: 0, width: '100%', height: '100%', cursor: 'inherit' },
};
