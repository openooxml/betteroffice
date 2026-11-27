import { useMemo } from 'react';
import type { SlideDisplayList } from '@betteroffice/pptx';
import { useTranslation } from '../i18n';
import { emptyPlaceholderGuides, type EmptyPlaceholderGuide } from '../slideLayouts';

function verticalAlign(anchor: EmptyPlaceholderGuide['anchor']): 'flex-start' | 'center' | 'flex-end' {
  if (anchor === 'center') return 'center';
  if (anchor === 'bottom') return 'flex-end';
  return 'flex-start';
}

function horizontalAlign(guide: EmptyPlaceholderGuide): 'flex-start' | 'center' | 'flex-end' {
  if (guide.align === 'center') return 'center';
  if (guide.align === 'right') return 'flex-end';
  if (guide.align === 'left' || guide.align === 'justify') return 'flex-start';
  return guide.placeholderType === 'ctrTitle' || guide.placeholderType === 'subTitle'
    ? 'center'
    : 'flex-start';
}

function promptBaseFontSize(placeholderType: string): number {
  if (placeholderType === 'title' || placeholderType === 'ctrTitle') return 22;
  if (placeholderType === 'subTitle') return 18;
  return 16;
}

function cssTransform(guide: EmptyPlaceholderGuide): string | undefined {
  const rotationDeg = guide.transform?.rotationDeg;
  return rotationDeg ? `rotate(${rotationDeg}deg)` : undefined;
}

export function PlaceholderGuides({
  displayList,
  scale,
  selectedShapeId,
  editingShapeId,
  hiddenShapeId,
}: {
  displayList: SlideDisplayList;
  scale: number;
  selectedShapeId?: string | null;
  editingShapeId?: string | null;
  hiddenShapeId?: string | null;
}) {
  const { t } = useTranslation();
  const guides = useMemo(() => emptyPlaceholderGuides(displayList, t), [displayList, t]);
  if (!guides.length) return null;
  return (
    <>
      {guides.map((guide) => {
        if (hiddenShapeId === guide.shapeId || editingShapeId === guide.shapeId) return null;
        const active = selectedShapeId === guide.shapeId;
        const justify = horizontalAlign(guide);
        const transform = cssTransform(guide);
        return (
          <div
            key={guide.key}
            data-testid="pptx-placeholder-guide"
            data-placeholder-type={guide.placeholderType}
            aria-hidden="true"
            style={{
              position: 'absolute',
              pointerEvents: 'none',
              boxSizing: 'border-box',
              overflow: 'hidden',
              left: guide.x * scale,
              top: guide.y * scale,
              width: Math.max(1, guide.w * scale),
              height: Math.max(1, guide.h * scale),
              border: active ? 'none' : '1px dashed rgba(100, 116, 139, 0.55)',
              display: 'flex',
              alignItems: verticalAlign(guide.anchor),
              justifyContent: justify,
              textAlign: justify === 'center' ? 'center' : justify === 'flex-end' ? 'right' : 'left',
              padding: Math.max(4, Math.round(8 * scale)),
              color: '#94a3b8',
              fontSize: Math.max(11, Math.round(promptBaseFontSize(guide.placeholderType) * scale)),
              lineHeight: 1.3,
              userSelect: 'none',
              ...(transform ? { transform, transformOrigin: 'center center' } : {}),
            }}
          >
            {guide.prompt}
          </div>
        );
      })}
    </>
  );
}
