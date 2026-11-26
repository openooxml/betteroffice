import type {
  PrimitiveTransform,
  SlideDisplayList,
  TextBoxPrimitive,
} from '@betteroffice/pptx';
import type { PresentationSlideLayoutSummary } from '@betteroffice/pptx/worker';
import type { TFunction, TranslationKey } from '@betteroffice/pptx-i18n';

const LAYOUT_TYPE_KEYS: Record<string, TranslationKey> = {
  title: 'toolbar.layouts.title',
  obj: 'toolbar.layouts.obj',
  secHead: 'toolbar.layouts.secHead',
  twoObj: 'toolbar.layouts.twoObj',
  twoTxTwoObj: 'toolbar.layouts.twoTxTwoObj',
  titleOnly: 'toolbar.layouts.titleOnly',
  blank: 'toolbar.layouts.blank',
  objTx: 'toolbar.layouts.objTx',
  picTx: 'toolbar.layouts.picTx',
  vertTx: 'toolbar.layouts.vertTx',
  vertTitleAndTx: 'toolbar.layouts.vertTitleAndTx',
};

const STANDARD_LAYOUT_NAME_KEYS: Record<string, TranslationKey> = {
  'title slide': 'toolbar.layouts.title',
  'title and content': 'toolbar.layouts.obj',
  'section header': 'toolbar.layouts.secHead',
  'two content': 'toolbar.layouts.twoObj',
  comparison: 'toolbar.layouts.twoTxTwoObj',
  'title only': 'toolbar.layouts.titleOnly',
  blank: 'toolbar.layouts.blank',
  'content with caption': 'toolbar.layouts.objTx',
  'picture with caption': 'toolbar.layouts.picTx',
  'title and vertical text': 'toolbar.layouts.vertTx',
  'vertical title and text': 'toolbar.layouts.vertTitleAndTx',
};

const EDITABLE_PLACEHOLDER_TYPES = new Set(['title', 'ctrTitle', 'subTitle', 'body']);

/** Resolve which layout part path clicking the main "New Slide" action uses. */
export function defaultInsertLayoutPath(
  layouts: readonly PresentationSlideLayoutSummary[] | undefined,
  activeLayoutPartPath: string | null | undefined,
): string | undefined {
  if (layouts?.length) {
    const active = activeLayoutPartPath
      ? layouts.find((layout) => layout.partPath === activeLayoutPartPath)
      : undefined;
    if (active && active.layoutType !== 'blank' && active.layoutType !== 'title') {
      return active.partPath;
    }
    const masterScoped = active?.masterPartPath
      ? layouts.filter((layout) => layout.masterPartPath === active.masterPartPath)
      : layouts;
    const pool = masterScoped.length ? masterScoped : layouts;
    const content =
      pool.find((layout) => layout.layoutType === 'obj') ??
      pool.find((layout) => layout.layoutType !== 'blank' && layout.layoutType !== 'title') ??
      pool.find((layout) => layout.layoutType !== 'blank') ??
      pool[0];
    return content?.partPath;
  }
  return activeLayoutPartPath ?? undefined;
}

/** Build localized toolbar/menu options for every layout available in the deck. */
export function slideLayoutChoices(
  layouts: readonly PresentationSlideLayoutSummary[] | undefined,
  fallbackSlideLayouts: readonly (string | null | undefined)[] | undefined,
  t: TFunction,
): { value: string; label: string; args: { layoutPartPath: string } }[] {
  if (layouts?.length) {
    return layouts.map((layout, index) => {
      const rawName = layout.name?.trim();
      const standardKey = rawName ? STANDARD_LAYOUT_NAME_KEYS[rawName.toLowerCase()] : undefined;
      const typeKey = layout.layoutType ? LAYOUT_TYPE_KEYS[layout.layoutType] : undefined;
      const label = standardKey
        ? t(standardKey)
        : rawName
          ? rawName
          : typeKey
            ? t(typeKey)
            : t('toolbar.layoutOption', { number: index + 1 });
      return { value: layout.partPath, label, args: { layoutPartPath: layout.partPath } };
    });
  }
  const unique = [
    ...new Set((fallbackSlideLayouts ?? []).filter((path): path is string => Boolean(path))),
  ];
  return unique.map((layoutPartPath, index) => ({
    value: layoutPartPath,
    label: t('toolbar.layoutOption', { number: index + 1 }),
    args: { layoutPartPath },
  }));
}

export interface EmptyPlaceholderGuide {
  key: string;
  shapeId: string;
  x: number;
  y: number;
  w: number;
  h: number;
  placeholderType: string;
  prompt: string;
  anchor: TextBoxPrimitive['anchor'];
  align?: TextBoxPrimitive['paragraphs'][number]['align'];
  transform?: PrimitiveTransform;
}

function isEmptyTextBox(box: TextBoxPrimitive): boolean {
  return (
    box.paragraphs.length <= 1 &&
    !box.paragraphs.some((paragraph) => paragraph.runs.some((run) => run.text.length > 0))
  );
}

/** Whether the targeted text box is an empty editable placeholder ready for first-click typing. */
export function isEmptyEditablePlaceholder(
  list: SlideDisplayList,
  shapeId: string,
  storyId?: string,
): boolean {
  return list.primitives.some(
    (primitive) =>
      primitive.kind === 'textBox' &&
      primitive.shapeId === shapeId &&
      (storyId === undefined || primitive.storyId === storyId) &&
      Boolean(
        primitive.placeholderType && EDITABLE_PLACEHOLDER_TYPES.has(primitive.placeholderType),
      ) &&
      isEmptyTextBox(primitive),
  );
}

function placeholderPromptKey(placeholderType: string): TranslationKey {
  if (placeholderType === 'title' || placeholderType === 'ctrTitle') {
    return 'editor.placeholders.title';
  }
  if (placeholderType === 'subTitle') {
    return 'editor.placeholders.subtitle';
  }
  return 'editor.placeholders.body';
}

/** Return visual placeholder guides for empty editable placeholder boxes on the slide. */
export function emptyPlaceholderGuides(
  list: SlideDisplayList,
  t: TFunction,
): EmptyPlaceholderGuide[] {
  const guides: EmptyPlaceholderGuide[] = [];
  for (const primitive of list.primitives) {
    if (
      primitive.kind !== 'textBox' ||
      !primitive.shapeId ||
      !primitive.placeholderType ||
      !EDITABLE_PLACEHOLDER_TYPES.has(primitive.placeholderType) ||
      !isEmptyTextBox(primitive)
    ) {
      continue;
    }
    const placeholderType = primitive.placeholderType;
    guides.push({
      key: `${primitive.shapeId}:${primitive.storyId ?? ''}`,
      shapeId: primitive.shapeId,
      x: primitive.x,
      y: primitive.y,
      w: primitive.w,
      h: primitive.h,
      placeholderType,
      prompt: t(placeholderPromptKey(placeholderType)),
      anchor: primitive.anchor,
      align: primitive.paragraphs[0]?.align,
      transform: primitive.transform,
    });
  }
  return guides;
}
