const wait = async (predicate, label) => {
  const deadline = Date.now() + 90000;
  while (Date.now() < deadline) {
    const error = document.querySelector('.notice');
    if (error) throw new Error(error.textContent);
    if (predicate()) return;
    await new Promise(resolve => setTimeout(resolve, 100));
  }
  throw new Error(label + ': ' + document.body.innerText.slice(0, 800));
};
await wait(() => [...document.querySelectorAll('.editor-stage canvas')].some(canvas => canvas.width > 100), 'Editor did not render');
const configuration = await window.webkit.messageHandlers.desktop.postMessage({type: 'configuration'});
if (configuration.format === 'docx') {
  const input = document.querySelector('[data-testid="yrs-input"]');
  input.focus();
  input.value = 'Native smoke ';
  input.dispatchEvent(new InputEvent('input', {bubbles: true, inputType: 'insertText', data: 'Native smoke '}));
} else if (configuration.format === 'xlsx') {
  const grid = document.querySelector('[data-testid="xlsx-scroll"]');
  grid.focus();
  grid.dispatchEvent(new KeyboardEvent('keydown', {key: 'N', bubbles: true}));
  await wait(() => document.querySelector('[data-testid="xlsx-cell-editor"]'), 'Cell editor did not open');
  document.querySelector('[data-testid="xlsx-cell-editor"]').dispatchEvent(new KeyboardEvent('keydown', {key: 'Enter', bubbles: true}));
} else {
  document.querySelector('[data-testid="pptx-new-slide"]').click();
}
await wait(() => document.querySelector('.document-name')?.textContent.includes('Edited'), 'Edit did not mark the file dirty');
await new Promise(resolve => setTimeout(resolve, 600));
window.desktop.command('save');
await wait(() => !document.querySelector('.document-name')?.textContent.includes('Edited'), 'Native save did not finish');
return {format: configuration.format, saved: true, canvases: document.querySelectorAll('canvas').length, controls: [...document.querySelectorAll('button')].slice(0, 25).map(button => button.getAttribute('aria-label') || button.textContent)};
