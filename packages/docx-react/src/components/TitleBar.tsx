/**
 * TitleBar and sub-components for the Google Docs-style 2-level toolbar.
 *
 * - TitleBar: two-row layout (row 1: logo + doc name + right actions, row 2: menu bar)
 * - Logo: renders custom logo content left-aligned
 * - DocumentName: editable document name input
 * - MenuBar: File/Format/Insert/Help menus built from the editor's commands
 * - TitleBarRight: right-aligned actions slot
 */

import React, { useCallback, useState, Children, isValidElement } from 'react';
import type { ReactNode } from 'react';
import { MenuDropdown } from './ui/MenuDropdown';
import type { MenuEntry } from './ui/MenuDropdown';
import { LayoutMenu } from './LayoutMenu';
import { InsertMenu } from './InsertMenu';
import { useCommandMenuItem, visible } from './commandMenuItem';
import { useTranslation } from '../i18n';

// ============================================================================
// Default Doc Icon (shown when no Logo is provided)
// ============================================================================

function DefaultDocIcon() {
  return (
    <svg width="32" height="40" viewBox="0 0 32 40" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path
        d="M2 0C0.9 0 0 0.9 0 2V38C0 39.1 0.9 40 2 40H30C31.1 40 32 39.1 32 38V10L22 0H2Z"
        fill="#cbd5e1"
      />
      <path d="M22 0L32 10H24C22.9 10 22 9.1 22 8V0Z" fill="#94a3b8" />
      <rect x="7" y="18" width="18" height="2" rx="1" fill="#fff" />
      <rect x="7" y="23" width="18" height="2" rx="1" fill="#fff" />
      <rect x="7" y="28" width="12" height="2" rx="1" fill="#fff" />
    </svg>
  );
}

// ============================================================================
// Logo
// ============================================================================

/** @experimental */
export interface LogoProps {
  children: ReactNode;
}

export function Logo({ children }: LogoProps) {
  return <div className="flex items-center flex-shrink-0">{children}</div>;
}

// ============================================================================
// DocumentName
// ============================================================================

/** @experimental */
export interface DocumentNameProps {
  value: string;
  onChange?: (value: string) => void;
  placeholder?: string;
  editable?: boolean;
}

function stripExtension(name: string): string {
  return name.replace(/\.docx$/i, '');
}

export function DocumentName({ value, onChange, placeholder, editable = true }: DocumentNameProps) {
  const { t } = useTranslation();
  const resolvedPlaceholder = placeholder ?? t('titleBar.untitled');
  const displayName = stripExtension(value) ?? '';

  if (!editable) {
    return (
      <span className="text-base font-normal text-foreground px-2 py-0 min-w-[100px] max-w-[300px] truncate leading-tight">
        {displayName || resolvedPlaceholder}
      </span>
    );
  }
  return (
    <input
      type="text"
      value={displayName}
      onChange={(e) => {
        const raw = e.target.value;
        onChange?.(raw.endsWith('.docx') ? raw : raw + '.docx');
      }}
      placeholder={resolvedPlaceholder}
      className="text-base font-normal text-foreground bg-transparent border-0 outline-none px-2 py-0 rounded hover:bg-muted focus:bg-doc-bg-input focus:ring-1 focus:ring-ring min-w-[100px] max-w-[300px] truncate leading-tight"
      aria-label={t('titleBar.documentNameAriaLabel')}
    />
  );
}

// ============================================================================
// TitleBarRight
// ============================================================================

/** @experimental */
export interface TitleBarRightProps {
  children: ReactNode;
}

export function TitleBarRight({ children }: TitleBarRightProps) {
  return <div className="flex items-center gap-2 ml-auto flex-shrink-0">{children}</div>;
}

// ============================================================================
// MenuBar
// ============================================================================

export function MenuBar() {
  const { t } = useTranslation();
  const open = useCommandMenuItem('open', 'file_upload');
  const save = useCommandMenuItem('save', 'file_download');
  const print = useCommandMenuItem('print', 'print');
  const pageSetup = useCommandMenuItem('pageSetup', 'settings');
  const ltr = useCommandMenuItem('setLtr', 'format_textdirection_l_to_r');
  const rtl = useCommandMenuItem('setRtl', 'format_textdirection_r_to_l');
  const reportIssue = useCommandMenuItem('reportIssue', undefined);

  const [activeMenu, setActiveMenu] = useState('file');
  const fileItems = visible(open, save);
  const printItems = visible(print, pageSetup);
  const file: MenuEntry[] = [
    ...fileItems,
    ...(fileItems.length > 0 && printItems.length > 0 ? [{ type: 'separator' as const }] : []),
    ...printItems,
  ];
  const focusedMenu =
    (activeMenu === 'file' && file.length === 0) || (activeMenu === 'help' && reportIssue.hidden)
      ? 'format'
      : activeMenu;

  return (
    <div className="flex items-center" role="menubar" aria-label={t('titleBar.menuBarAriaLabel')}>
      {file.length > 0 && (
        <MenuDropdown
          label={t('toolbar.file')}
          items={file}
          tabIndex={focusedMenu === 'file' ? 0 : -1}
          onFocus={() => setActiveMenu('file')}
        />
      )}

      <MenuDropdown
        label={t('toolbar.format')}
        items={visible(ltr, rtl)}
        tabIndex={focusedMenu === 'format' ? 0 : -1}
        onFocus={() => setActiveMenu('format')}
      />

      <InsertMenu
        menuBar={{
          tabIndex: focusedMenu === 'insert' ? 0 : -1,
          onFocus: () => setActiveMenu('insert'),
        }}
      />

      <LayoutMenu
        tabIndex={focusedMenu === 'layout' ? 0 : -1}
        onFocus={() => setActiveMenu('layout')}
      />

      {!reportIssue.hidden && (
        <MenuDropdown
          label={t('toolbar.help')}
          items={visible(reportIssue)}
          tabIndex={focusedMenu === 'help' ? 0 : -1}
          onFocus={() => setActiveMenu('help')}
        />
      )}
    </div>
  );
}

// ============================================================================
// TitleBar
// ============================================================================

/** @experimental */
export interface TitleBarProps {
  children: ReactNode;
}

/**
 * TitleBar layout (Google Docs style):
 *
 *   ┌──────────┬────────────────────────────┬──────────────────┐
 *   │          │ Document Name              │                  │
 *   │  Logo    │                            │  Right Actions   │
 *   │          │ File  Format  Insert       │                  │
 *   └──────────┴────────────────────────────┴──────────────────┘
 *
 * Logo and TitleBarRight span full height. DocumentName + MenuBar
 * stack vertically in the center column.
 */
export function TitleBar({ children }: TitleBarProps) {
  let logoItem: ReactNode = null;
  let rightItem: ReactNode = null;
  const middleTopItems: ReactNode[] = [];
  const menuBarItems: ReactNode[] = [];

  Children.forEach(children, (child) => {
    if (!isValidElement(child)) return;
    if (child.type === Logo) {
      logoItem = child;
    } else if (child.type === TitleBarRight) {
      rightItem = child;
    } else if (child.type === MenuBar) {
      menuBarItems.push(child);
    } else {
      middleTopItems.push(child);
    }
  });

  const handleMouseDown = useCallback((e: React.MouseEvent) => {
    const target = e.target as HTMLElement;
    const isInteractive =
      target.tagName === 'INPUT' ||
      target.tagName === 'TEXTAREA' ||
      target.tagName === 'SELECT' ||
      target.tagName === 'OPTION';

    if (!isInteractive) {
      e.preventDefault();
    }
  }, []);

  return (
    <div
      className="flex items-stretch bg-doc-surface pt-2 pb-1"
      onMouseDown={handleMouseDown}
      data-testid="title-bar"
    >
      {/* Left: Logo spanning full height (default doc icon if none provided) */}
      <div className="flex items-center flex-shrink-0 pl-3 pr-1">
        {logoItem || <DefaultDocIcon />}
      </div>

      {/* Center: doc name on top, menus below */}
      <div className="flex flex-col justify-center flex-1 min-w-0 py-1">
        {middleTopItems.length > 0 && (
          <div className="flex items-center gap-2 px-1">{middleTopItems}</div>
        )}
        {menuBarItems.length > 0 && <div className="flex items-center px-1">{menuBarItems}</div>}
      </div>

      {/* Right: actions spanning full height */}
      {rightItem && <div className="flex items-center flex-shrink-0 px-3">{rightItem}</div>}
    </div>
  );
}
