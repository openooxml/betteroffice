import { MenuDropdown, type MenuEntry } from './MenuDropdown';
import { cn } from '../../lib/utils';
import { IconLineSpacing } from './Icons';
import { useTranslation } from '../../i18n';
import { createT, en } from '@betteroffice/docx-i18n';
import type { TFunction, TranslationKey } from '@betteroffice/docx-i18n';

function formatMultiplier(multiplier: number, lang: string): string {
  return new Intl.NumberFormat(lang, { maximumFractionDigits: 2 }).format(multiplier);
}

// ============================================================================
// TYPES
// ============================================================================

export interface LineSpacingOption {
  label: string;
  labelKey?: TranslationKey;
  value: number;
  twipsValue: number;
}

export interface LineSpacingPickerProps {
  value?: number;
  onChange?: (twipsValue: number) => void;
  spaceBefore?: number;
  spaceAfter?: number;
  onParagraphSpacingChange?: (spacing: { before?: number; after?: number }) => void;
  onOpenOptions?: () => void;
  options?: LineSpacingOption[];
  disabled?: boolean;
  /** Why the picker is disabled. */
  description?: string;
  className?: string;
  width?: number | string;
}

// ============================================================================
// CONSTANTS
// ============================================================================

/**
 * Standard line spacing options (Google Docs style)
 * OOXML uses twips for line spacing when lineRule="auto"
 * 240 twips = 1.0 line spacing (single)
 */
const DEFAULT_OPTIONS: LineSpacingOption[] = [
  { label: 'Single', labelKey: 'lineSpacing.single', value: 1.0, twipsValue: 240 },
  { label: '1.15', value: 1.15, twipsValue: 276 },
  { label: '1.5', value: 1.5, twipsValue: 360 },
  { label: 'Double', labelKey: 'lineSpacing.double', value: 2.0, twipsValue: 480 },
];

// ============================================================================
// COMPONENT
// ============================================================================

export function LineSpacingPicker({
  value,
  onChange,
  spaceBefore = 0,
  spaceAfter = 0,
  onParagraphSpacingChange,
  onOpenOptions,
  options = DEFAULT_OPTIONS,
  disabled = false,
  description,
  className,
}: LineSpacingPickerProps) {
  const { t, lang } = useTranslation();
  const getOptionLabel = (option: LineSpacingOption) =>
    option.labelKey ? t(option.labelKey) : formatMultiplier(option.value, lang);

  const currentOption = options.find((option) => option.twipsValue === value) ?? options[0];
  const items: MenuEntry[] = options.map((option) => ({
    label: getOptionLabel(option),
    icon: option.twipsValue === currentOption.twipsValue ? 'check' : undefined,
    checked: option.twipsValue === currentOption.twipsValue,
    onClick: () => onChange?.(option.twipsValue),
  }));
  items.push(
    { type: 'separator' },
    {
      label: t(spaceBefore !== 0 ? 'lineSpacing.removeBefore' : 'lineSpacing.addBefore'),
      disabled: !onParagraphSpacingChange,
      onClick: () => onParagraphSpacingChange?.({ before: spaceBefore !== 0 ? 0 : 240 }),
    },
    {
      label: t(spaceAfter !== 0 ? 'lineSpacing.removeAfter' : 'lineSpacing.addAfter'),
      disabled: !onParagraphSpacingChange,
      onClick: () => onParagraphSpacingChange?.({ after: spaceAfter !== 0 ? 0 : 240 }),
    },
    { label: t('lineSpacing.options'), disabled: !onOpenOptions, onClick: onOpenOptions }
  );

  return (
    <MenuDropdown
      label={t('lineSpacing.label')}
      trigger={<IconLineSpacing className="h-5 w-5 shrink-0" />}
      triggerClassName={cn('h-7 text-sm gap-0.5 px-2', disabled && 'opacity-[0.38]', className)}
      title={description ?? t('lineSpacing.lineSpacingTitle', { label: getOptionLabel(currentOption) })}
      description={disabled ? description : undefined}
      showChevron
      disabled={disabled}
      items={items}
    />
  );
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

export function getDefaultLineSpacingOptions(): LineSpacingOption[] {
  return [...DEFAULT_OPTIONS];
}

export function lineSpacingMultiplierToTwips(multiplier: number): number {
  return Math.round(multiplier * 240);
}

export function twipsToLineSpacingMultiplier(twips: number): number {
  return twips / 240;
}

export function getLineSpacingLabel(
  twips: number,
  t: TFunction = createT(en),
  lang = 'en'
): string {
  const option = DEFAULT_OPTIONS.find((opt) => opt.twipsValue === twips);
  if (option) return option.labelKey ? t(option.labelKey) : formatMultiplier(option.value, lang);
  const multiplier = twipsToLineSpacingMultiplier(twips);
  return formatMultiplier(multiplier, lang);
}

export function isStandardLineSpacing(twips: number): boolean {
  return DEFAULT_OPTIONS.some((opt) => opt.twipsValue === twips);
}

export function nearestStandardLineSpacing(twips: number): LineSpacingOption {
  let nearest = DEFAULT_OPTIONS[0];
  let minDiff = Math.abs(twips - nearest.twipsValue);

  for (const option of DEFAULT_OPTIONS) {
    const diff = Math.abs(twips - option.twipsValue);
    if (diff < minDiff) {
      minDiff = diff;
      nearest = option;
    }
  }

  return nearest;
}

export function createLineSpacingOption(multiplier: number, lang = 'en'): LineSpacingOption {
  const twipsValue = lineSpacingMultiplierToTwips(multiplier);
  const label = formatMultiplier(multiplier, lang);
  return { label, value: multiplier, twipsValue };
}

export default LineSpacingPicker;
