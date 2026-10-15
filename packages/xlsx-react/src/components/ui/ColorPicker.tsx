import { useId, useState } from 'react';
import { ToolbarIcon } from './ToolbarIcon';
import {
  TOOLBAR_DISABLED_OPACITY,
  toolbarFocusStyle,
  toolbarTheme,
  useFocusRing,
} from '../../../../../shared/react-toolbar/theme';

export interface ColorPickerProps {
  mode: 'text' | 'fill' | 'border';
  value: string;
  label: string;
  onChange?: (value: string) => void;
  /** Called as the platform picker opens, before any `onChange`. */
  onOpen?: () => void;
  disabled?: boolean;
  /** Why the picker is disabled; announced and shown on hover, and it stays focusable. */
  description?: string;
}

export function ColorPicker({
  mode,
  value,
  label,
  onChange,
  onOpen,
  disabled = false,
  description,
}: ColorPickerProps) {
  const [hovered, setHovered] = useState(false);
  const focus = useFocusRing<HTMLInputElement>();
  const id = useId();
  const reason = disabled && description ? description : undefined;
  const title = reason ? `${label}: ${reason}` : label;
  return (
    <label
      title={title}
      aria-label={label}
      style={{
        position: 'relative',
        display: 'inline-grid',
        placeItems: 'center',
        width: 28,
        height: 28,
        borderRadius: 6,
        background: hovered && !disabled ? toolbarTheme.hover : 'transparent',
        color: hovered && !disabled ? toolbarTheme.ink : toolbarTheme.inkMuted,
        cursor: disabled ? 'default' : 'pointer',
        opacity: disabled ? TOOLBAR_DISABLED_OPACITY : 1,
        boxSizing: 'border-box',
        ...toolbarFocusStyle(focus.focusRing),
      }}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
    >
      <ToolbarIcon
        name={mode === 'text' ? 'textColor' : mode === 'fill' ? 'fillColor' : 'borderColor'}
      />
      <span
        aria-hidden="true"
        style={{
          position: 'absolute',
          left: 7,
          right: 7,
          bottom: 5,
          height: 3,
          borderRadius: 1,
          boxShadow: `inset 0 0 0 1px ${toolbarTheme.line}`,
          background: value,
        }}
      />
      <input
        type="color"
        {...focus.focusProps}
        value={value}
        disabled={disabled && !reason}
        aria-disabled={reason ? true : undefined}
        aria-describedby={reason ? id : undefined}
        aria-label={label}
        onClick={(event) => {
          if (disabled) event.preventDefault();
          else onOpen?.();
        }}
        onKeyDown={(event) => {
          if (disabled && (event.key === 'Enter' || event.key === ' ')) event.preventDefault();
        }}
        onChange={(event) => {
          if (!disabled) onChange?.(event.target.value);
        }}
        style={{
          position: 'absolute',
          inset: 0,
          width: '100%',
          height: '100%',
          opacity: 0,
        }}
      />
      {reason && (
        <span id={id} hidden>
          {reason}
        </span>
      )}
    </label>
  );
}
