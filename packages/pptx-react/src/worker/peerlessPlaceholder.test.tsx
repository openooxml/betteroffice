import { expect, test } from 'bun:test';
import { act, fireEvent, waitFor } from '@testing-library/react';
import JSZip from 'jszip';
import { createT, en, zhCN } from '@betteroffice/pptx-i18n';
import { decodeCommandDisplayList } from '@betteroffice/pptx/worker';
import * as wasm from '../../../pptx/src/wasm/loader';
import { emptyPlaceholderGuides } from '../slideLayouts';
import { open } from './peerlessTestHarness';

function compose(input: HTMLElement, type: string, data = '') {
  const event = new Event(type, { bubbles: true });
  Object.defineProperty(event, 'data', { value: data });
  fireEvent(input, event);
}

async function deckWithFooterAndContentLayout(options: {
  existingContent?: boolean;
  customTitle?: boolean;
} = {}): Promise<Uint8Array> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-officedocument.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${rel}/slide" Target="slides/slide1.xml"/></Relationships>`);
  zip.file('ppt/slideMasters/slideMaster1.xml', `<p:sldMaster ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/><p:sldLayoutId id="2147483651" r:id="rId3"/></p:sldLayoutIdLst></p:sldMaster>`);
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${rel}/slideLayout" Target="../slideLayouts/slideLayout2.xml"/><Relationship Id="rId3" Type="${rel}/slideLayout" Target="../slideLayouts/slideLayout3.xml"/></Relationships>`);
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<p:sldLayout ${ns} type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
  zip.file('ppt/slideLayouts/slideLayout2.xml', `<p:sldLayout ${ns} type="obj"><p:cSld name="Title and Content"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="800000" y="450000"/><a:ext cx="10000000" cy="1100000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="800000" y="1750000"/><a:ext cx="10000000" cy="4400000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="t"/><a:lstStyle/></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout2.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
  zip.file('ppt/slideLayouts/slideLayout3.xml', `<p:sldLayout ${ns} type="obj"><p:cSld name="季度汇报模板"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout3.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="4" name="Footer 1"/><p:cNvSpPr/><p:nvPr><p:ph type="ftr" idx="11"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3000000" y="6200000"/><a:ext cx="4000000" cy="400000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`);
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`);
  if (options.existingContent) {
    const path = 'ppt/slides/slide1.xml';
    const slide = await zip.file(path)!.async('string');
    const shapes = ['type="title"', 'type="body" idx="1"'].map((ph, index) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${index + 2}" name="Placeholder ${index + 2}"/><p:cNvSpPr/><p:nvPr><p:ph ${ph}/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>`);
    zip.file(path, slide.replace('</p:spTree>', `${shapes.join('')}</p:spTree>`));
    const relPath = 'ppt/slides/_rels/slide1.xml.rels';
    const relationship = await zip.file(relPath)!.async('string');
    zip.file(relPath, relationship.replace('slideLayout1.xml', 'slideLayout2.xml'));
  }
  if (options.customTitle) {
    const path = 'ppt/slideLayouts/slideLayout2.xml';
    const layout = await zip.file(path)!.async('string');
    zip.file(path, layout
      .replace('type="title"', 'type="title" hasCustomPrompt="1"')
      .replace('<a:bodyPr anchor="ctr"/><a:lstStyle/>',
        '<a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:r><a:t>Enter </a:t></a:r><a:r><a:t>your headline</a:t></a:r></a:p>')
      .replace('<a:bodyPr anchor="t"/><a:lstStyle/>',
        '<a:bodyPr anchor="t"/><a:lstStyle/><a:p><a:r><a:t>Unflagged layout text</a:t></a:r></a:p>'));
  }
  return zip.generateAsync({ type: 'uint8array' });
}

test('existing empty title and body placeholders show their default guides', async () => {
  const file = await deckWithFooterAndContentLayout({ existingContent: true });
  const value = await open({ file });
  const list = decodeCommandDisplayList(value.client.state.frame!);
  expect(emptyPlaceholderGuides(list, createT(en)).map(({ placeholderType, prompt }) =>
    [placeholderType, prompt])).toEqual([
    ['title', 'Click to add title'],
    ['body', 'Click to add text'],
  ]);
});

test('placeholder guides use flagged layout text and keep the default for unflagged layout text', async () => {
  const file = await deckWithFooterAndContentLayout({ existingContent: true, customTitle: true });
  const value = await open({ file });
  const prompts = () => value.view.getAllByTestId('pptx-placeholder-guide').map((guide) => guide.textContent);
  await waitFor(() => expect(prompts()).toEqual(['Enter your headline', 'Click to add text']));

  await act(async () => {
    expect(await value.api.commands.execute('insertSlide', {})).toMatchObject({ ok: true });
  });
  await waitFor(() => {
    expect(value.client.state.slides).toHaveLength(2);
    expect(value.client.state.frame?.slideId).toBe(value.client.state.slides[1]!.id);
    expect(prompts()).toEqual(['Enter your headline', 'Click to add text']);
  });
});

test('clicking deep inside an empty body placeholder enters text editing, hides guide during IME, and skips footer guides', async () => {
  const bytes = await deckWithFooterAndContentLayout();
  const value = await open({ file: bytes, i18n: zhCN });

  // Existing empty footer placeholder on slide 1 must not show a placeholder guide.
  expect(value.view.queryAllByTestId('pptx-placeholder-guide')).toHaveLength(0);
  // Custom layout name stays intact in zh-CN alongside localized built-in names.
  expect(value.api.commands.getState('insertSlide').options?.map((o) => o.label)).toEqual([
    '空白',
    '标题和内容',
    '季度汇报模板',
  ]);

  await act(async () => {
    expect(await value.api.commands.execute('insertSlide', {})).toMatchObject({ ok: true });
  });
  await waitFor(() => expect(value.view.getAllByTestId('pptx-placeholder-guide')).toHaveLength(2));

  const list = decodeCommandDisplayList(value.client.state.frame!);
  const bodyGuide = list.primitives.find(
    (primitive) => primitive.kind === 'textBox' && primitive.placeholderType === 'body',
  )!;
  const bounds = value.canvas.getBoundingClientRect();
  const scaleX = bounds.width / list.width;
  const scaleY = bounds.height / list.height;

  // Click in the bottom third of the empty body placeholder, well below line 1.
  const clickX = bounds.left + (bodyGuide.x + bodyGuide.w * 0.5) * scaleX;
  const clickY = bounds.top + (bodyGuide.y + bodyGuide.h * 0.75) * scaleY;
  await act(async () => {
    fireEvent.pointerDown(value.stage, { button: 0, pointerId: 1, isPrimary: true, clientX: clickX, clientY: clickY });
    fireEvent.pointerUp(value.stage, { button: 0, pointerId: 1, isPrimary: true, clientX: clickX, clientY: clickY });
    await value.api.flushPendingInput();
  });

  // Entering text editing hides the body guide immediately (only the title guide remains).
  await waitFor(() =>
    expect(value.view.getAllByTestId('pptx-placeholder-guide').map((el) => el.dataset.placeholderType)).toEqual([
      'title',
    ]),
  );

  // Chinese IME composition on the empty placeholder stays clear of the guide prompt.
  const input = value.view.getByLabelText('幻灯片文本');
  await act(async () => {
    input.focus();
    compose(input, 'compositionstart');
    await value.api.flushPendingInput();
  });
  await act(async () => {
    compose(input, 'compositionupdate', '你好');
  });
  await waitFor(() => expect(value.view.queryByTestId('pptx-preedit')?.textContent).toBe('你好'));
  expect(value.view.getAllByTestId('pptx-placeholder-guide').map((el) => el.dataset.placeholderType)).toEqual([
    'title',
  ]);

  await act(async () => {
    compose(input, 'compositionend', '你好世界');
    await value.api.flushPendingInput();
  });

  // Clearing selection leaves only the empty title guide visible (body now has text).
  await act(async () => {
    await value.api.clearSelection();
  });
  await waitFor(() =>
    expect(value.view.getAllByTestId('pptx-placeholder-guide').map((el) => el.textContent)).toEqual([
      '单击此处添加标题',
    ]),
  );

  // Save and reopen keeps <p:ph> on both placeholders and text on body.
  const saved = wasm.openPresentation(await value.api.saveAsync());
  try {
    const slide = saved.snapshot().slides[1]!;
    expect(slide.shapes).toHaveLength(2);
    expect(slide.shapes[0]!.placeholder).toMatchObject({ placeholderType: 'title' });
    expect(slide.shapes[1]!.placeholder).toMatchObject({ index: 1 });
    expect(slide.shapes[1]!.textStories[0]!.paragraphs[0]!.runs.map((r) => r.text).join('')).toBe('你好世界');
  } finally {
    saved.dispose();
  }
});
