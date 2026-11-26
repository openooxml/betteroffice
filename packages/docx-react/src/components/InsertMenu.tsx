import { useRef } from 'react';
import type { OverflowMenuEntry } from '../../../../shared/react-toolbar/OverflowMenu';
import { useDocxCommand, useDocxCommands } from '../commands/hooks';
import type { DocxCommandId, DocxCommandState } from '../commands/types';
import { useTranslation } from '../i18n';
import { cn } from '../lib/utils';
import { hiddenInMenu, useCommandMenuItem, visible } from './commandMenuItem';
import { commandOverflowEntry } from './toolbar/ToolbarCommand';
import { useOverflowSource } from './toolbar/overflowRegistry';
import { MaterialSymbol } from './ui/MaterialSymbol';
import { MenuDropdown } from './ui/MenuDropdown';
import type { MenuEntry } from './ui/MenuDropdown';
import { TableGridInline } from './ui/TableGridInline';

interface BreakSubmenuItem {
  icon: string;
  label: string;
  onClick?: () => void;
  description?: string;
}

function BreakSubmenu({ items, closeMenu }: { items: BreakSubmenuItem[]; closeMenu: () => void }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', minWidth: 220 }}>
      {items.map((item) => {
        const disabled = !item.onClick;
        return (
          <button
            key={item.label}
            type="button"
            tabIndex={-1}
            role="menuitem"
            aria-disabled={disabled || undefined}
            title={disabled ? item.description : undefined}
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: 8,
              padding: '6px 12px',
              border: 'none',
              background: 'transparent',
              cursor: disabled ? 'default' : 'pointer',
              fontSize: 13,
              color: 'var(--doc-text)',
              width: '100%',
              textAlign: 'left',
              whiteSpace: 'nowrap',
              opacity: disabled ? 0.4 : 1,
            }}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              if (disabled) return;
              item.onClick?.();
              closeMenu();
            }}
            onMouseOver={(e) => {
              if (!disabled) {
                (e.currentTarget as HTMLButtonElement).style.backgroundColor =
                  'var(--doc-bg-hover)';
              }
            }}
            onMouseOut={(e) => {
              (e.currentTarget as HTMLButtonElement).style.backgroundColor = 'transparent';
            }}
          >
            <MaterialSymbol name={item.icon} size={18} />
            <span>{item.label}</span>
          </button>
        );
      })}
    </div>
  );
}

const BREAKS = [
  'insertPageBreak',
  'insertSectionBreakNextPage',
  'insertSectionBreakContinuous',
] as const satisfies readonly DocxCommandId[];

/** @experimental */
export interface InsertMenuProps {
  className?: string;
}

/** Roving focus a menu bar supplies to its menus. */
export interface MenuBarSlot {
  tabIndex: number;
  onFocus(): void;
}

/**
 * The Insert menu: image, table, breaks, table of contents and watermark.
 * In a formatting rail it moves into the overflow menu at narrow widths.
 * @experimental
 */
export function InsertMenu({
  className,
  menuBar,
}: InsertMenuProps & {
  /** @internal Set by `MenuBar`. */
  menuBar?: MenuBarSlot;
}) {
  const { t } = useTranslation();
  const store = useDocxCommands();
  const wrapperRef = useRef<HTMLSpanElement>(null);
  const image = useCommandMenuItem('insertImage', 'image');
  const insertTable = useDocxCommand('insertTable');
  const pageBreak = useCommandMenuItem('insertPageBreak', 'page_break');
  const nextPage = useCommandMenuItem('insertSectionBreakNextPage', 'horizontal_rule');
  const continuous = useCommandMenuItem('insertSectionBreakContinuous', 'border_horizontal');
  const toc = useCommandMenuItem('insertTOC', 'format_list_numbered');
  const watermark = useCommandMenuItem('watermark', 'branding_watermark');
  const label = t('toolbar.insert');
  const breakLabel = t('toolbar.break');

  useOverflowSource(wrapperRef, () => {
    const shown = (id: DocxCommandId) => !hiddenInMenu(store.getState(id) as DocxCommandState);
    const entry = (id: DocxCommandId) => commandOverflowEntry(store, t, id);
    return [
      {
        kind: 'submenu',
        id: 'insertMenu',
        label,
        entries: [
          ...(shown('insertImage') ? [entry('insertImage')] : []),
          entry('insertTable'),
          { kind: 'separator', id: 'insertMenu-separator' },
          { kind: 'submenu', id: 'insertMenu-break', label: breakLabel, entries: BREAKS.map(entry) },
          ...(['insertTOC', 'watermark'] as const).filter(shown).map(entry),
        ],
      } satisfies OverflowMenuEntry,
    ];
  });

  const tableEntry: MenuEntry = insertTable.state.enabled
    ? {
        icon: 'grid_on',
        label: insertTable.label,
        submenuRole: 'dialog',
        submenuContent: (closeMenu: () => void) => (
          <TableGridInline
            onInsert={(rows: number, columns: number) => {
              void insertTable.execute({ rows, columns });
              closeMenu();
            }}
          />
        ),
      }
    : {
        icon: 'grid_on',
        label: insertTable.label,
        disabled: true,
        description: insertTable.state.disabledReason.message,
      };
  const items: MenuEntry[] = [
    ...visible(image),
    tableEntry,
    { type: 'separator' },
    {
      icon: 'page_break',
      label: breakLabel,
      submenuContent: (closeMenu: () => void) => (
        <BreakSubmenu
          closeMenu={closeMenu}
          items={[pageBreak, nextPage, continuous].map(({ entry }) => ({
            icon: entry.icon ?? 'page_break',
            label: entry.label,
            onClick: entry.disabled ? undefined : entry.onClick,
            description: entry.description,
          }))}
        />
      ),
    },
    ...visible(toc, watermark),
  ];

  if (menuBar) {
    return (
      <MenuDropdown
        label={label}
        tabIndex={menuBar.tabIndex}
        onFocus={menuBar.onFocus}
        items={items}
      />
    );
  }
  return (
    <span ref={wrapperRef} className={cn('inline-flex flex-shrink-0', className)}>
      <MenuDropdown
        label={label}
        trigger={label}
        triggerClassName="h-7 text-sm gap-0.5 px-2"
        showChevron
        items={items}
      />
    </span>
  );
}
