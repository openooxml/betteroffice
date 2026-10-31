import { useId, useRef, useState } from 'react';
import type { CSSProperties } from 'react';
import type { SectionProperties } from '@betteroffice/docx/types/document';
import { useDocxCommand } from '../commands/hooks';
import { useTranslation } from '../i18n';
import { useEditorChrome } from './EditorToolbarContext';
import { MenuDropdown, type MenuEntry } from './ui/MenuDropdown';
import {
  formatLayoutLength,
  formatLayoutMargins,
  MARGIN_PRESETS,
  PAGE_SIZES,
  pageLayoutOrientation,
  pageOrientationPatch,
} from './pageLayout';

interface LayoutChoice {
  label: string;
  secondaryText?: string;
  checked?: boolean;
  disabled?: boolean;
  description?: string;
  onClick(): void;
}

const choiceStyle: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'flex-start',
  gap: 2,
  padding: '6px 12px',
  width: '100%',
  border: 'none',
  background: 'transparent',
  color: 'var(--doc-text)',
  cursor: 'pointer',
  fontSize: 13,
  textAlign: 'left',
  whiteSpace: 'nowrap',
};

function LayoutChoices({ choices, disabled, closeMenu }: {
  choices: LayoutChoice[];
  disabled: boolean;
  closeMenu(): void;
}) {
  const id = useId();
  return (
    <>
      {choices.map((choice, index) => {
        const unavailable = disabled || choice.disabled;
        const descriptionId = choice.secondaryText ? `${id}-${index}` : undefined;
        return (
          <button
            key={choice.label}
            type="button"
            role={choice.checked === undefined ? 'menuitem' : 'menuitemradio'}
            aria-checked={choice.checked}
            aria-label={choice.label}
            aria-describedby={descriptionId}
            aria-disabled={unavailable || undefined}
            title={choice.description}
            tabIndex={-1}
            style={{
              ...choiceStyle,
              opacity: unavailable ? 0.4 : 1,
              cursor: unavailable ? 'default' : 'pointer',
            }}
            onMouseDown={(event) => event.preventDefault()}
            onMouseOver={(event) => {
              if (!unavailable) event.currentTarget.style.backgroundColor = 'var(--doc-bg-hover)';
            }}
            onMouseOut={(event) => {
              event.currentTarget.style.backgroundColor = 'transparent';
            }}
            onClick={() => {
              if (unavailable) return;
              choice.onClick();
              closeMenu();
            }}
          >
            <span><span aria-hidden="true">{choice.checked ? '✓ ' : ''}</span>{choice.label}</span>
            {choice.secondaryText && (
              <span id={descriptionId} style={{ color: 'var(--doc-text-muted)', fontSize: 11 }}>
                {choice.secondaryText}
              </span>
            )}
          </button>
        );
      })}
    </>
  );
}

export function LayoutMenu({ tabIndex, onFocus }: { tabIndex: number; onFocus(): void }) {
  const { t, lang } = useTranslation();
  const layout = useEditorChrome()?.pageLayout;
  const pageSetup = useDocxCommand('pageSetup');
  const indent = useDocxCommand('indent');
  const outdent = useDocxCommand('outdent');
  const [loading, setLoading] = useState(false);
  const readRequest = useRef(0);
  const readCurrent = () => {
    if (!layout?.readCurrent || !pageSetup.state.enabled) return;
    const request = ++readRequest.current;
    setLoading(true);
    void layout.readCurrent().then(
      () => { if (request === readRequest.current) setLoading(false); },
      () => { if (request === readRequest.current) setLoading(false); }
    );
  };
  const properties = layout?.currentProps;
  const unit = layout?.unit ?? 'inch';
  const orientation = pageLayoutOrientation(properties);
  const openDialog = () => { void pageSetup.execute(); };
  const apply = (patch: Partial<SectionProperties>) => {
    if (pageSetup.state.enabled) void layout?.onApply(patch);
  };
  const submenu = (label: string, choices: LayoutChoice[], pageItem = true): MenuEntry => {
    const disabled = pageItem && (!pageSetup.state.enabled || !layout || loading || (!!layout.readCurrent && !properties));
    return {
      label,
      disabled,
      description: !pageItem || pageSetup.state.enabled ? undefined : pageSetup.state.disabledReason.message,
      submenuContent: (closeMenu) => (
        <LayoutChoices choices={choices} disabled={disabled} closeMenu={closeMenu} />
      ),
    };
  };
  const margins = MARGIN_PRESETS.map(({ labelKey, ...preset }) => ({
    label: t(labelKey),
    secondaryText: formatLayoutMargins(preset, unit, t, lang),
    checked: (Object.keys(preset) as (keyof typeof preset)[]).every(
      (key) => (properties?.[key] ?? 1440) === preset[key]
    ),
    onClick: () => apply(preset),
  }));
  const sizes = PAGE_SIZES.map((size) => {
    const pageWidth = orientation === 'landscape' ? size.height : size.width;
    const pageHeight = orientation === 'landscape' ? size.width : size.height;
    return {
      label: t(size.labelKey).replace(/\s*\([^()]*×[^()]*\)\s*$/, ''),
      secondaryText: `${formatLayoutLength(pageWidth, unit, t, lang)} × ${formatLayoutLength(pageHeight, unit, t, lang)}`,
      checked: Math.abs((properties?.pageWidth || 12240) - pageWidth) < 20 &&
        Math.abs((properties?.pageHeight || 15840) - pageHeight) < 20,
      onClick: () => apply({ pageWidth, pageHeight, orientation }),
    };
  });
  return (
    <MenuDropdown
      label={t('toolbar.layout')}
      tabIndex={tabIndex}
      onFocus={onFocus}
      onOpen={readCurrent}
      items={[
        submenu(t('dialogs.pageSetup.margins'), [
          ...margins,
          { label: t('layout.customMargins'), onClick: openDialog },
        ]),
        submenu(t('dialogs.pageSetup.orientation'), (['portrait', 'landscape'] as const).map((value) => ({
          label: t(value === 'portrait' ? 'dialogs.pageSetup.portrait' : 'dialogs.pageSetup.landscape'),
          checked: orientation === value,
          onClick: () => apply(pageOrientationPatch(properties, value)),
        }))),
        submenu(t('dialogs.pageSetup.sizeLabel'), [
          ...sizes,
          { label: t('layout.morePaperSizes'), onClick: openDialog },
        ]),
        { type: 'separator' },
        submenu(t('layout.indent'), [
          {
            label: t('lists.increaseIndent'),
            disabled: !indent.state.enabled,
            description: indent.state.enabled ? undefined : indent.state.disabledReason.message,
            onClick: () => { void indent.execute(); },
          },
          {
            label: t('lists.decreaseIndent'),
            disabled: !outdent.state.enabled,
            description: outdent.state.enabled ? undefined : outdent.state.disabledReason.message,
            onClick: () => { void outdent.execute(); },
          },
        ], false),
      ]}
    />
  );
}
