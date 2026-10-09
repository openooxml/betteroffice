import { useEffect, useRef, useState } from 'react';
import type { CSSProperties } from 'react';
import type { TFunction } from '@betteroffice/vsdx-i18n';
import { CommandMenu } from './CommandMenu';
import type { CommandMenuEntry } from './CommandMenu';
import type { RibbonCommandId } from './commands';
import { ShapeMiniToolbar, miniToolbarPosition } from './ShapeMiniToolbar';

export interface ShapeContextMenuProps {
  t: TFunction;
  position: { top: number; left: number };
  onClose: () => void;
  onCloseAndFocus: () => void;
}

/** Shape operations offered on right-click, in Visio's relative order. */
export const SHAPE_CONTEXT_ENTRIES: ReadonlyArray<CommandMenuEntry> = [
  { id: 'delete', icon: 'delete' },
  {
    id: 'bringToFront', icon: 'front', children: [
      { id: 'bringToFront', icon: 'front' },
      { id: 'bringForward', icon: 'forward' },
    ],
  },
  {
    id: 'sendToBack', icon: 'back', children: [
      { id: 'sendBackward', icon: 'backward' },
      { id: 'sendToBack', icon: 'back' },
    ],
  },
  {
    id: 'rotateRight', icon: 'rotateRight', children: [
      { id: 'rotateRight', icon: 'rotateRight' },
      { id: 'rotateLeft', icon: 'rotateLeft' },
      { id: 'flipHorizontal', icon: 'flipHorizontal' },
      { id: 'flipVertical', icon: 'flipVertical' },
    ],
  },
];

const SHAPE_CONTEXT_DIVIDERS: ReadonlySet<RibbonCommandId> = new Set(['delete']);

/** Right-click menu for the selected shape, sharing the ribbon menus' keyboard behaviour. */
export function ShapeContextMenu({ t, position, onClose, onCloseAndFocus }: ShapeContextMenuProps) {
  const wrapRef = useRef<HTMLDivElement | null>(null);
  const toolbarRef = useRef<HTMLDivElement | null>(null);
  const [toolbarStyle, setToolbarStyle] = useState<CSSProperties>({ top: position.top, left: position.left, visibility: 'hidden' });
  const [below, setBelow] = useState(false);
  const menuLabel = t('contextMenu.label');
  useEffect(() => {
    const wrap = wrapRef.current;
    const toolbar = toolbarRef.current;
    if (!wrap || !toolbar) return;
    const update = () => {
      const menu = wrap.querySelector('[role="menu"]:not([data-submenu])') as HTMLElement | null;
      if (!menu) return;
      const menuRect = menu.getBoundingClientRect();
      const toolbarRect = toolbar.getBoundingClientRect();
      const next = miniToolbarPosition(
        { top: menuRect.top, left: menuRect.left, bottom: menuRect.bottom, right: menuRect.right },
        { width: toolbarRect.width, height: toolbarRect.height },
        { width: window.innerWidth, height: window.innerHeight },
      );
      setToolbarStyle((previous) => {
        const style = { top: next.top, left: next.left, visibility: 'visible' as const };
        return previous.top === style.top && previous.left === style.left && previous.visibility === style.visibility ? previous : style;
      });
      setBelow((previous) => (previous === next.below ? previous : next.below));
    };
    update();
    const menu = wrap.querySelector('[role="menu"]:not([data-submenu])');
    const observer = typeof MutationObserver === 'undefined' ? null : new MutationObserver(update);
    if (menu && observer) observer.observe(menu, { attributes: true, attributeFilter: ['style'] });
    window.addEventListener('resize', update);
    return () => {
      observer?.disconnect();
      window.removeEventListener('resize', update);
    };
  }, [position, menuLabel]);
  return (
    <div ref={wrapRef}>
      <ShapeMiniToolbar t={t} toolbarRef={toolbarRef} style={toolbarStyle} below={below} />
      <CommandMenu menuLabel={menuLabel} entries={SHAPE_CONTEXT_ENTRIES} position={position} anchorRef={toolbarRef} dividerAfter={SHAPE_CONTEXT_DIVIDERS} label={(id) => t(`ribbon.commands.${id}`)} onClose={onClose} onCloseAndFocus={onCloseAndFocus} />
    </div>
  );
}
