/** Interactive checkbox, dropdown, and date controls authored through Yrs. */

import { useCallback, useEffect, useRef, useState } from 'react';
import type { CSSProperties, KeyboardEvent as ReactKeyboardEvent, RefObject } from 'react';
import type { YrsContentControlValue } from '@betteroffice/docx/yrs';

const WIDGET_SELECTOR = '.layout-sdt-widget, .layout-inline-sdt-widget';

type ControlTarget = {
  pos: number | null;
  tag?: string;
  id?: number;
};

type Popup =
  | {
      kind: 'dropdown';
      target: ControlTarget;
      items: { displayText: string; value: string }[];
      current: string;
      rect: DOMRect;
    }
  | { kind: 'date'; target: ControlTarget; current: string; rect: DOMRect };

export interface ContentControlWidgetsProps {
  containerRef: RefObject<HTMLElement | null>;
  applyYrsValue: (pos: number, value: YrsContentControlValue, embedId?: string) => boolean;
}

function positionFromTrigger(trigger: HTMLElement): number | null {
  const group = /^sdt@(\d+)$/.exec(trigger.dataset.sdtGroupId ?? '');
  const raw = group?.[1] ?? trigger.dataset.sdtPos;
  const position = raw == null || raw === '' ? Number.NaN : Number(raw);
  return Number.isFinite(position) ? position : null;
}

function targetFromTrigger(trigger: HTMLElement): ControlTarget {
  const rawId = Number(trigger.dataset.sdtControlId);
  return {
    pos: positionFromTrigger(trigger),
    ...(trigger.dataset.sdtTag ? { tag: trigger.dataset.sdtTag } : {}),
    ...(Number.isFinite(rawId) ? { id: rawId } : {}),
  };
}

function triggerListItems(trigger: HTMLElement): Array<{ displayText: string; value: string }> {
  try {
    const parsed = JSON.parse(trigger.dataset.sdtListItems ?? '[]') as unknown;
    if (!Array.isArray(parsed)) return [];
    return parsed.flatMap((item) => {
      if (!item || typeof item !== 'object') return [];
      const candidate = item as { displayText?: unknown; value?: unknown };
      if (typeof candidate.value !== 'string') return [];
      return [
        {
          displayText:
            typeof candidate.displayText === 'string' ? candidate.displayText : candidate.value,
          value: candidate.value,
        },
      ];
    });
  } catch {
    return [];
  }
}

export function ContentControlWidgets({
  containerRef,
  applyYrsValue,
}: ContentControlWidgetsProps): React.ReactElement | null {
  const [popup, setPopup] = useState<Popup | null>(null);
  const popupRef = useRef<HTMLDivElement>(null);

  const apply = useCallback(
    (target: ControlTarget, value: YrsContentControlValue) => {
      if (target.pos == null) return;
      try {
        applyYrsValue(target.pos, value, target.id === undefined ? undefined : String(target.id));
      } catch {
        // Locked, bound, or invalid values remain inert in the UI layer.
      }
      setPopup(null);
    },
    [applyYrsValue]
  );

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    const activate = (trigger: HTMLElement) => {
      const kind = trigger.dataset.sdtWidget;
      if (!kind) return;
      const target = targetFromTrigger(trigger);
      if (kind === 'checkbox') {
        const checked = trigger.getAttribute('aria-checked') === 'true';
        apply(target, { kind: 'checkbox', checked: !checked });
      } else if (kind === 'dropdown') {
        setPopup({
          kind: 'dropdown',
          target,
          items: triggerListItems(trigger),
          current: trigger.dataset.sdtValue ?? '',
          rect: trigger.getBoundingClientRect(),
        });
      } else if (kind === 'date') {
        setPopup({
          kind: 'date',
          target,
          current: trigger.dataset.sdtValue?.slice(0, 10) ?? '',
          rect: trigger.getBoundingClientRect(),
        });
      }
    };

    const onMouseDown = (event: MouseEvent) => {
      const trigger = (event.target as HTMLElement | null)?.closest?.(WIDGET_SELECTOR);
      if (trigger) event.preventDefault();
    };
    const onClick = (event: MouseEvent) => {
      const trigger = (event.target as HTMLElement | null)?.closest?.(
        WIDGET_SELECTOR
      ) as HTMLElement | null;
      if (!trigger) return;
      event.preventDefault();
      event.stopPropagation();
      activate(trigger);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== 'Enter' && event.key !== ' ') return;
      const trigger = (event.target as HTMLElement | null)?.closest?.(
        WIDGET_SELECTOR
      ) as HTMLElement | null;
      if (!trigger) return;
      event.preventDefault();
      activate(trigger);
    };

    container.addEventListener('mousedown', onMouseDown);
    container.addEventListener('click', onClick);
    container.addEventListener('keydown', onKeyDown);
    return () => {
      container.removeEventListener('mousedown', onMouseDown);
      container.removeEventListener('click', onClick);
      container.removeEventListener('keydown', onKeyDown);
    };
  }, [apply, containerRef]);

  useEffect(() => {
    if (!popup) return;
    const onMouseDown = (event: MouseEvent) => {
      if (!popupRef.current?.contains(event.target as Node)) setPopup(null);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setPopup(null);
    };
    document.addEventListener('mousedown', onMouseDown);
    document.addEventListener('keydown', onKeyDown);
    return () => {
      document.removeEventListener('mousedown', onMouseDown);
      document.removeEventListener('keydown', onKeyDown);
    };
  }, [popup]);

  useEffect(() => {
    if (popup?.kind !== 'dropdown') return;
    const options = popupRef.current?.querySelectorAll<HTMLElement>(
      '.layout-sdt-widget-option'
    );
    if (!options?.length) return;
    ([...options].find((option) => option.getAttribute('aria-selected') === 'true') ?? options[0])
      .focus();
  }, [popup]);

  const handlePopupKeyDown = (event: ReactKeyboardEvent) => {
    if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
    const options = [
      ...(popupRef.current?.querySelectorAll<HTMLElement>('.layout-sdt-widget-option') ?? []),
    ];
    if (options.length === 0) return;
    event.preventDefault();
    const index = options.indexOf(document.activeElement as HTMLElement);
    const next =
      event.key === 'ArrowDown'
        ? (index + 1) % options.length
        : (index - 1 + options.length) % options.length;
    options[next].focus();
  };

  if (!popup) return null;
  const style: CSSProperties = {
    position: 'fixed',
    top: popup.rect.bottom + 2,
    left: popup.rect.left,
    zIndex: 1000,
  };

  return (
    <div
      ref={popupRef}
      className="layout-sdt-widget-popup"
      style={style}
      role={popup.kind === 'dropdown' ? 'listbox' : undefined}
      onKeyDown={handlePopupKeyDown}
      onMouseDown={(event) => event.preventDefault()}
    >
      {popup.kind === 'dropdown' ? (
        popup.items.length === 0 ? (
          <div className="layout-sdt-widget-empty">No options</div>
        ) : (
          popup.items.map((item) => (
            <button
              key={item.value}
              type="button"
              role="option"
              aria-selected={item.displayText === popup.current}
              className={`layout-sdt-widget-option${
                item.displayText === popup.current ? ' is-selected' : ''
              }`}
              onClick={() => apply(popup.target, { kind: 'dropdown', value: item.value })}
            >
              {item.displayText}
            </button>
          ))
        )
      ) : (
        <input
          type="date"
          className="layout-sdt-widget-date"
          autoFocus
          defaultValue={popup.current}
          onChange={(event) => {
            if (event.target.value) {
              apply(popup.target, { kind: 'date', date: event.target.value });
            }
          }}
        />
      )}
    </div>
  );
}
