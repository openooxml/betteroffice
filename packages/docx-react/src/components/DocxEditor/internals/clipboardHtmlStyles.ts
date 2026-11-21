type Declaration = { value: string; order: number };
type Rules = Map<string, Map<string, Declaration>>;

const STYLE_TEXT_LIMIT = 200_000;
const DECLARATION_LIMIT = 512;
const SELECTOR = /^([a-z][\w-]*)?(\.[\w-]+)?$/i;
const RUN_PROPERTIES = new Set(['font-weight', 'font-style', 'font-size', 'font-family', 'text-decoration-line',
  'color', 'background-color', 'vertical-align', 'white-space', 'mso-highlight', 'mso-list']);

function* delimiters(css: string, wanted: string) {
  let parentheses = 0;
  for (const token of css.matchAll(/"(?:\\[\s\S]|[^"\\])*(?:"|$)|'(?:\\[\s\S]|[^'\\])*(?:'|$)|\\[\s\S]|[(){};,]/g)) {
    if (token[0] === '(') parentheses++;
    else if (token[0] === ')') parentheses = Math.max(0, parentheses - 1);
    else if (!parentheses && wanted.includes(token[0])) yield { at: token.index!, token: token[0] };
  }
}

function withoutComments(css: string): string {
  return css.replace(/"(?:\\[\s\S]|[^"\\])*(?:"|$)|'(?:\\[\s\S]|[^'\\])*(?:'|$)|\/\*[\s\S]*?(?:\*\/|$)/g,
    token => token.startsWith('/*') ? ' ' : token);
}

function readDeclarations(css: string): Map<string, string> {
  const declarations = new Map<string, string>();
  let start = 0;
  for (const { at } of delimiters(css + ';', ';')) {
    const entry = css.slice(start, at);
    start = at + 1;
    const colon = entry.indexOf(':');
    if (colon < 0) continue;
    let name = entry.slice(0, colon).trim().toLowerCase();
    const value = entry.slice(colon + 1).trim().replace(/!\s*important$/i, '').trimEnd();
    if (!name || !value || value.length > DECLARATION_LIMIT) continue;
    if (name === 'text-decoration') name = 'text-decoration-line';
    if (name === 'background') name = 'background-color';
    if (name === 'font') {
      const font = /^((?:(?:normal|italic|oblique|bold|[1-9]00|small-caps)\s+)*)((?:\d+(?:\.\d*)?|\.\d+)(?:pt|px))(?:\s*\/\s*\S+)?\s+(\S[\s\S]*)$/i.exec(value);
      if (!font) continue;
      const modifiers = font[1]!.trim().toLowerCase().split(/\s+/);
      declarations.set('font-weight', modifiers.find(word => /^(bold|[1-9]00)$/.test(word)) ?? 'normal');
      declarations.set('font-style', modifiers.find(word => word === 'italic' || word === 'oblique') ?? 'normal');
      declarations.set('font-size', font[2]!);
      declarations.set('font-family', font[3]!);
    } else if (RUN_PROPERTIES.has(name)) declarations.set(name, value);
  }
  return declarations;
}

const GENERIC_FAMILY = /^(?:-|ui-)|^(?:serif|sans-serif|monospace|cursive|fantasy|system-ui|emoji|math|fangsong|inherit|initial|unset|revert)$/i;

export function formattedFontFamily(value: string): string | null {
  let start = 0;
  for (const { at } of delimiters(value + ',', ',')) {
    const family = value.slice(start, at).trim().replace(/^(["'])([\s\S]*)\1$/, '$2').trim();
    start = at + 1;
    if (family && !GENERIC_FAMILY.test(family)) return family.length <= 31 ? family : null;
  }
  return null;
}

export function clipboardHtmlStyles(doc: Document) {
  const rules: Rules = new Map();
  let remaining = STYLE_TEXT_LIMIT;
  let order = 0;
  for (const element of doc.querySelectorAll('style')) {
    if (!remaining) break;
    const text = (element.textContent ?? '').slice(0, remaining);
    remaining -= text.length;
    const css = withoutComments(text).replace(/^\s*<!--|-->\s*$/g, '');
    let start = 0;
    let depth = 0;
    let selectors: string[] = [];
    for (const { at, token } of delimiters(css, '{};')) {
      if (token === '{') {
        selectors = depth === 0 ? css.slice(start, at).trim().split(',').map(s => s.trim()) : [];
        depth++;
        start = at + 1;
      } else if (token === '}' && depth > 0) {
        depth--;
        if (!depth && selectors.length && selectors.every(s => s && SELECTOR.test(s))) {
          const targets = selectors.map(s => {
            const [, tag, cls] = SELECTOR.exec(s)!;
            const key = (tag ?? '').toLowerCase() + (cls ?? '');
            let target = rules.get(key);
            if (!target) {
              target = new Map();
              rules.set(key, target);
            }
            return target;
          });
          for (const [name, value] of readDeclarations(css.slice(start, at))) {
            if (name === 'mso-list') continue;
            const declaration = { value, order: ++order };
            for (const target of targets) target.set(name, declaration);
          }
        }
        start = at + 1;
      } else if (!depth) start = at + 1;
    }
  }
  return (element: Element): Map<string, string> => {
    const style = new Map<string, string>();
    const tag = element.tagName.toLowerCase();
    for (const [name, declaration] of rules.get(tag) ?? []) style.set(name, declaration.value);
    for (const prefix of ['', tag]) {
      const orders = new Map<string, number>();
      for (const cls of element.classList) {
        for (const [name, declaration] of rules.get(prefix + '.' + cls) ?? []) {
          if (declaration.order > (orders.get(name) ?? -1)) {
            style.set(name, declaration.value);
            orders.set(name, declaration.order);
          }
        }
      }
    }
    const inline = element.getAttribute('style');
    if (inline) {
      for (const [name, value] of readDeclarations(withoutComments(inline))) style.set(name, value);
    }
    return style;
  };
}
