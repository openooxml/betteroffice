import type {
  BorderPreset,
  BorderStyle,
  HorizontalAlignment,
  VerticalAlignment,
} from '../Toolbar';
import { ToolbarSymbol } from '../../../../../shared/react-toolbar/icons';
import {
  borderAll,
  borderBottom,
  borderClear,
  borderHorizontal,
  borderInner,
  borderLeft,
  borderOuter,
  borderRight,
  borderTop,
  borderVertical,
  formatAlignCenter,
  formatAlignLeft,
  formatAlignRight,
  verticalAlignBottom,
  verticalAlignCenter,
  verticalAlignTop,
} from '../../../../../shared/react-toolbar/symbols';

const BORDER_SYMBOLS: Record<BorderPreset, string> = {
  all: borderAll,
  inner: borderInner,
  horizontal: borderHorizontal,
  vertical: borderVertical,
  outer: borderOuter,
  left: borderLeft,
  top: borderTop,
  right: borderRight,
  bottom: borderBottom,
  none: borderClear,
};

const HORIZONTAL_SYMBOLS: Record<HorizontalAlignment, string> = {
  left: formatAlignLeft,
  center: formatAlignCenter,
  right: formatAlignRight,
};

const VERTICAL_SYMBOLS: Record<VerticalAlignment, string> = {
  top: verticalAlignTop,
  middle: verticalAlignCenter,
  bottom: verticalAlignBottom,
};

export function BorderGlyph({ preset }: { preset: BorderPreset }) {
  return <ToolbarSymbol path={BORDER_SYMBOLS[preset]} />;
}

export function HorizontalAlignmentGlyph({ value }: { value: HorizontalAlignment }) {
  return <ToolbarSymbol path={HORIZONTAL_SYMBOLS[value]} />;
}

export function VerticalAlignmentGlyph({ value }: { value: VerticalAlignment }) {
  return <ToolbarSymbol path={VERTICAL_SYMBOLS[value]} />;
}

export function BorderStyleGlyph({ value }: { value: BorderStyle }) {
  return (
    <svg
      width="44"
      height="12"
      viewBox="0 0 44 12"
      fill="none"
      stroke="currentColor"
      aria-hidden="true"
    >
      {value === 'double' ? (
        <path d="M2 4h40M2 8h40" />
      ) : (
        <path
          d="M2 6h40"
          strokeWidth="2"
          strokeDasharray={value === 'dashed' ? '7 4' : value === 'dotted' ? '1 4' : undefined}
        />
      )}
    </svg>
  );
}

export function ColorSwatch({ color }: { color: string }) {
  return (
    <span
      aria-hidden="true"
      style={{ display: 'inline-block', width: 20, height: 5, borderRadius: 2, background: color }}
    />
  );
}
