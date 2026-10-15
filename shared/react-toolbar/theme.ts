import { useState } from 'react';
import type { CSSProperties, FocusEvent } from 'react';

/**
 * Toolbar palette as CSS custom properties. A host restyles the toolbars by
 * setting any `--bo-toolbar-*` property on an ancestor; the second value is
 * the neutral default. The docx stylesheet reads the same properties.
 */
const TOKENS = [
  'ink',
  'ink-muted',
  'rail',
  'surface',
  'hover',
  'selected',
  'line',
  'accent',
  'accent-ink',
] as const;

export const toolbarTheme = {
  ink: 'var(--bo-toolbar-ink, #1f1f1f)',
  inkMuted: 'var(--bo-toolbar-ink-muted, #444746)',
  rail: 'var(--bo-toolbar-rail, #f1f1f1)',
  surface: 'var(--bo-toolbar-surface, #ffffff)',
  hover: 'var(--bo-toolbar-hover, rgba(31, 31, 31, 0.08))',
  selected: 'var(--bo-toolbar-selected, rgba(31, 31, 31, 0.12))',
  line: 'var(--bo-toolbar-line, rgba(31, 31, 31, 0.16))',
  accent: 'var(--bo-toolbar-accent, #1f1f1f)',
  accentInk: 'var(--bo-toolbar-accent-ink, #ffffff)',
} as const;

export const TOOLBAR_CONTROL_SIZE = 28;
export const TOOLBAR_ICON_SIZE = 18;
export const TOOLBAR_RADIUS = 6;
export const TOOLBAR_DISABLED_OPACITY = 0.38;
export const TOOLBAR_FONT_FAMILY =
  'ui-sans-serif, system-ui, -apple-system, "Segoe UI", Roboto, Helvetica, Arial, sans-serif';
export const TOOLBAR_MENU_SHADOW = '0 4px 16px rgba(0, 0, 0, 0.12), 0 1px 3px rgba(0, 0, 0, 0.08)';

export interface ToolbarControlState {
  disabled?: boolean;
  /** Pressed toggle: ghost selected tint. */
  pressed?: boolean;
  /** Open popup trigger: hover fill. */
  open?: boolean;
  hovered?: boolean;
  /** Keyboard focus from `useFocusRing`; omitted, the browser's own focus ring shows. */
  focusRing?: boolean;
}

/** Focus ring for keyboard focus: a 2px accent outline just outside the control. */
export function toolbarFocusStyle(visible: boolean | undefined, offset = 1): CSSProperties {
  if (visible === undefined) return {};
  return visible
    ? { outline: `2px solid ${toolbarTheme.accent}`, outlineOffset: offset }
    : { outline: 'none' };
}

/**
 * The host's `--bo-toolbar-*` values at `element`, for popups portaled
 * outside the editor tree where they would not inherit.
 */
export function inheritedToolbarTokens(element: Element | null): CSSProperties {
  if (!element) return {};
  const computed = getComputedStyle(element);
  const tokens: Record<string, string> = {};
  for (const token of TOKENS) {
    const name = `--bo-toolbar-${token}`;
    const value = computed.getPropertyValue(name).trim();
    if (value) tokens[name] = value;
  }
  return tokens as CSSProperties;
}

/** Icon and text buttons on the rail. */
export function toolbarButtonStyle({
  disabled = false,
  pressed = false,
  open = false,
  hovered = false,
  focusRing,
}: ToolbarControlState): CSSProperties {
  return {
    appearance: 'none',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    gap: 2,
    minWidth: TOOLBAR_CONTROL_SIZE,
    height: TOOLBAR_CONTROL_SIZE,
    padding: '0 4px',
    border: 0,
    borderRadius: TOOLBAR_RADIUS,
    background: pressed
      ? toolbarTheme.selected
      : (hovered || open) && !disabled
        ? toolbarTheme.hover
        : 'transparent',
    color: pressed
      ? toolbarTheme.ink
      : hovered || open
        ? toolbarTheme.ink
        : toolbarTheme.inkMuted,
    cursor: disabled ? 'default' : 'pointer',
    opacity: disabled ? TOOLBAR_DISABLED_OPACITY : 1,
    font: `500 13px ${TOOLBAR_FONT_FAMILY}`,
    lineHeight: 1,
    whiteSpace: 'nowrap',
    boxSizing: 'border-box',
    ...toolbarFocusStyle(focusRing),
  };
}

/** Solid primary action button. */
export function toolbarPrimaryButtonStyle({
  disabled = false,
  focusRing,
}: Pick<ToolbarControlState, 'disabled' | 'focusRing'>): CSSProperties {
  return {
    ...toolbarButtonStyle({ disabled, focusRing }),
    background: toolbarTheme.accent,
    color: toolbarTheme.accentInk,
  };
}

/** Text fields and comboboxes on the rail. */
export function toolbarFieldStyle({
  disabled = false,
  focusRing,
}: Pick<ToolbarControlState, 'disabled' | 'focusRing'>): CSSProperties {
  return {
    height: TOOLBAR_CONTROL_SIZE,
    border: `1px solid ${toolbarTheme.line}`,
    borderRadius: TOOLBAR_RADIUS,
    background: toolbarTheme.surface,
    color: toolbarTheme.ink,
    opacity: disabled ? TOOLBAR_DISABLED_OPACITY : 1,
    font: `400 13px ${TOOLBAR_FONT_FAMILY}`,
    boxSizing: 'border-box',
    ...toolbarFocusStyle(focusRing),
  };
}

/** The rounded rail that holds the controls. */
export const toolbarRailStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 0,
  minHeight: 36,
  padding: '4px 8px',
  borderRadius: 18,
  background: toolbarTheme.rail,
  color: toolbarTheme.ink,
  font: `400 13px ${TOOLBAR_FONT_FAMILY}`,
  boxSizing: 'border-box',
};

/** Vertical divider between groups. */
export const toolbarSeparatorStyle: CSSProperties = {
  width: 1,
  height: 20,
  margin: '0 6px',
  background: toolbarTheme.line,
  flex: '0 0 auto',
};

/** Popup menus and dialogs opened from the rail. */
export const toolbarMenuStyle: CSSProperties = {
  padding: 4,
  border: `1px solid ${toolbarTheme.line}`,
  borderRadius: 8,
  background: toolbarTheme.surface,
  color: toolbarTheme.ink,
  boxShadow: TOOLBAR_MENU_SHADOW,
  boxSizing: 'border-box',
  outline: 'none',
};

/** One row of a popup menu. */
export function toolbarMenuItemStyle({
  disabled = false,
  hovered = false,
  focusRing,
}: Pick<ToolbarControlState, 'disabled' | 'hovered' | 'focusRing'>): CSSProperties {
  return {
    appearance: 'none',
    display: 'flex',
    alignItems: 'center',
    gap: 10,
    width: '100%',
    minHeight: 32,
    padding: '4px 10px',
    border: 0,
    borderRadius: TOOLBAR_RADIUS,
    background: (hovered || focusRing) && !disabled ? toolbarTheme.hover : 'transparent',
    color: toolbarTheme.ink,
    cursor: disabled ? 'default' : 'pointer',
    opacity: disabled ? TOOLBAR_DISABLED_OPACITY : 1,
    font: `400 13px ${TOOLBAR_FONT_FAMILY}`,
    textAlign: 'left',
    boxSizing: 'border-box',
    ...toolbarFocusStyle(focusRing, -2),
  };
}

export const toolbarMenuSeparatorStyle: CSSProperties = {
  height: 1,
  margin: '4px 2px',
  background: toolbarTheme.line,
};

function matchesFocusVisible(element: Element): boolean {
  try {
    return element.matches(':focus-visible');
  } catch {
    return true;
  }
}

/**
 * Tracks keyboard focus for controls styled inline, where `:focus-visible`
 * cannot be expressed. Spread `focusProps` on the control.
 */
export function useFocusRing<T extends Element = HTMLElement>(
  handlers: {
    onFocus?: (event: FocusEvent<T>) => void;
    onBlur?: (event: FocusEvent<T>) => void;
  } = {}
) {
  const [focusRing, setFocusRing] = useState(false);
  return {
    focusRing,
    focusProps: {
      onFocus: (event: FocusEvent<T>) => {
        setFocusRing(matchesFocusVisible(event.currentTarget));
        handlers.onFocus?.(event);
      },
      onBlur: (event: FocusEvent<T>) => {
        setFocusRing(false);
        handlers.onBlur?.(event);
      },
    },
  };
}
