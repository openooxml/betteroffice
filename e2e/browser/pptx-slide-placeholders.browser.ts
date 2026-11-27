import { test, expect, type Page } from 'playwright/test';
import JSZip from 'jszip';
import { inputReady, slideCanvas } from './pptx-async-helpers';

const thumbnails = (page: Page) => page.locator('aside[aria-label="Slides"] button');

async function buildDeck(): Promise<Buffer> {
  const zip = new JSZip();
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ns = `xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${rel}"`;
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-officedocument.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideLayouts/slideLayout2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
  zip.file('_rels/.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`);
  zip.file('ppt/presentation.xml', `<p:presentation ${ns}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
  zip.file('ppt/_rels/presentation.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${rel}/slide" Target="slides/slide1.xml"/></Relationships>`);
  zip.file('ppt/slideMasters/slideMaster1.xml', `<p:sldMaster ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst></p:sldMaster>`);
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${rel}/slideLayout" Target="../slideLayouts/slideLayout2.xml"/></Relationships>`);
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<p:sldLayout ${ns} type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
  zip.file('ppt/slideLayouts/slideLayout2.xml', `<p:sldLayout ${ns} type="obj"><p:cSld name="Title and Content"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="800000" y="450000"/><a:ext cx="10000000" cy="1100000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="800000" y="1750000"/><a:ext cx="10000000" cy="4400000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="t"/><a:lstStyle/></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout2.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
  zip.file('ppt/slides/slide1.xml', `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>`);
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`);
  return Buffer.from(await zip.generateAsync({ type: 'uint8array' }));
}

async function openBuffer(page: Page, buffer: Buffer, name = 'placeholders.pptx') {
  await page.goto('/?format=pptx');
  await page.locator('input[type=file]').last().setInputFiles({
    name,
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer,
  });
  await expect(slideCanvas(page)).toBeVisible();
  await inputReady(page);
  await expect(page.getByTestId('pptx-worker-status')).toHaveCount(0);
}

test('pptx: new slide instantiates layout placeholders, clicks to type, undoes, and preserves <p:ph> on save/reopen', async ({ page }) => {
  test.setTimeout(60_000);
  await openBuffer(page, await buildDeck());
  await expect(thumbnails(page)).toHaveCount(1);

  await page.getByTestId('pptx-new-slide').click();
  await expect(thumbnails(page)).toHaveCount(2);

  const guides = page.getByTestId('pptx-placeholder-guide');
  await expect(guides).toHaveCount(2);
  await expect(guides.nth(0)).toHaveText('Click to add title');
  await expect(guides.nth(1)).toHaveText('Click to add text');

  // Click deep inside the body placeholder (well below the single empty top line).
  const bodyBox = await guides.nth(1).boundingBox();
  expect(bodyBox).not.toBeNull();
  await page.mouse.click(bodyBox!.x + bodyBox!.width * 0.5, bodyBox!.y + bodyBox!.height * 0.75);

  await expect(guides).toHaveCount(1);
  await page.keyboard.type('Roadmap notes');

  // Click empty slide margin to clear selection; title guide stays, filled body guide stays hidden.
  const canvasBox = await slideCanvas(page).boundingBox();
  expect(canvasBox).not.toBeNull();
  await page.mouse.click(canvasBox!.x + 4, canvasBox!.y + 4);
  await expect(guides).toHaveCount(1);
  await expect(guides.first()).toHaveText('Click to add title');

  // Save and verify <p:ph> elements are preserved in the saved OOXML archive.
  const downloadPromise = page.waitForEvent('download');
  await page.getByTestId('pptx-save').click();
  const download = await downloadPromise;
  const savedStream = await download.createReadStream();
  const chunks: Buffer[] = [];
  for await (const chunk of savedStream) chunks.push(Buffer.from(chunk));
  const savedBytes = Buffer.concat(chunks);

  const savedZip = await JSZip.loadAsync(savedBytes);
  const slide2Xml = await savedZip.file('ppt/slides/slide2.xml')?.async('string');
  expect(slide2Xml).toBeTruthy();
  expect(slide2Xml).toContain('<p:ph type="title"');
  expect(slide2Xml).toContain('<p:ph idx="1"');
  expect(slide2Xml).toContain('Roadmap notes');

  // Undo text edit -> body guide reappears; undo slide insert -> back to 1 slide.
  await page.getByTestId('pptx-undo').click();
  await expect(guides).toHaveCount(2);
  await page.getByTestId('pptx-undo').click();
  await expect(thumbnails(page)).toHaveCount(1);

  // Reopen the saved deck and confirm slide 2 shows only the still-empty title guide.
  await openBuffer(page, savedBytes, 'saved-placeholders.pptx');
  await expect(thumbnails(page)).toHaveCount(2);
  await thumbnails(page).nth(1).click();
  await expect(guides).toHaveCount(1);
  await expect(guides.first()).toHaveText('Click to add title');
});
