import { useId, useState } from 'react';
import { ToolbarIcon } from './ToolbarIcon';
import { toolbarColors } from './ToolbarPrimitives';

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
        borderRadius: 4,
        background: hovered && !disabled ? toolbarColors.hover : 'transparent',
        color: disabled ? toolbarColors.disabled : toolbarColors.text,
        cursor: disabled ? 'default' : 'pointer',
        opacity: disabled ? 0.48 : 1,
        boxSizing: 'border-box',
      }}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
    >
      <ToolbarIcon
        name={mode === 'text' ? 'textColor' : mode === 'fill' ? 'fillColor' : 'borders'}
        size={20}
      />
      <span
        aria-hidden="true"
        style={{
          position: 'absolute',
          left: 5,
          right: 5,
          bottom: 2,
          height: 3,
          borderRadius: 2,
          background: value,
        }}
      />
      <input
        type="color"
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
