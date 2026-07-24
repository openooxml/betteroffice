import { useState } from 'react';
import { ToolbarIcon } from './ToolbarIcon';
import type { ToolbarIconName } from './ToolbarIcon';
import { toolbarColors } from './ToolbarPrimitives';

export interface ColorPickerProps {
  value: string;
  label: string;
  onChange?: (value: string) => void;
  onClear?: () => void;
  icon?: ToolbarIconName;
  none?: boolean;
  clearLabel?: string;
  disabled?: boolean;
  testId?: string;
}

export function ColorPicker({
  value,
  label,
  onChange,
  onClear,
  icon = 'textColor',
  none = false,
  clearLabel,
  disabled = false,
  testId,
}: ColorPickerProps) {
  const [hovered, setHovered] = useState(false);
  return (
    <span
      style={{
        display: 'inline-flex',
        alignItems: 'center',
        height: 28,
      }}
    >
      <label
        title={label}
        aria-label={label}
        style={{
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
          position: 'relative',
        }}
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
      >
        <ToolbarIcon name={icon} size={20} />
        <span
          aria-hidden="true"
          style={{
            position: 'absolute',
            left: 5,
            right: 5,
            bottom: 2,
            height: 3,
            borderRadius: 2,
            background: none ? 'transparent' : value,
            border: none ? `1px solid ${toolbarColors.disabled}` : 0,
            boxSizing: 'border-box',
          }}
        />
        {none ? (
          <span
            aria-hidden="true"
            style={{
              position: 'absolute',
              left: 5,
              bottom: 3,
              width: 18,
              height: 1,
              background: '#d93025',
              transform: 'rotate(-18deg)',
            }}
          />
        ) : null}
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
      {onClear ? (
        <button
          type="button"
          disabled={disabled}
          aria-label={clearLabel ?? label}
          title={clearLabel ?? label}
          onMouseDown={(event) => event.preventDefault()}
          onClick={onClear}
          style={{
            appearance: 'none',
            width: 15,
            height: 28,
            padding: 0,
            border: 0,
            borderRadius: 3,
            background: 'transparent',
            color: disabled ? toolbarColors.disabled : toolbarColors.muted,
            cursor: disabled ? 'default' : 'pointer',
            opacity: disabled ? 0.48 : 1,
            font: '500 14px ui-sans-serif, system-ui, sans-serif',
          }}
        >
          ×
        </button>
      ) : null}
    </span>
  );
}
