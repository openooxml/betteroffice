import { useState } from 'react';
import { ToolbarIcon } from './ToolbarIcon';
import { toolbarColors } from './ToolbarPrimitives';

export interface ColorPickerProps {
  value: string;
  label: string;
  onChange?: (value: string) => void;
  disabled?: boolean;
  testId?: string;
}

export function ColorPicker({
  value,
  label,
  onChange,
  disabled = false,
  testId,
}: ColorPickerProps) {
  const [hovered, setHovered] = useState(false);
  return (
    <label
      title={label}
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
      <ToolbarIcon name="textColor" size={20} />
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
        data-testid={testId}
        type="color"
        value={value}
        disabled={disabled}
        aria-label={label}
        onChange={(event) => onChange?.(event.target.value)}
        style={{
          position: 'absolute',
          inset: 0,
          width: '100%',
          height: '100%',
          opacity: 0,
        }}
      />
    </label>
  );
}
