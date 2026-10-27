import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';
import ts from 'typescript';

export interface HardcodedString {
  file: string;
  line: number;
  kind: 'jsx-text' | 'jsx-attribute' | 'named' | 'literal' | 'format';
  name: string;
  text: string;
}

const USER_ATTRIBUTES = new Set([
  'aria-label',
  'aria-description',
  'aria-roledescription',
  'aria-valuetext',
  'aria-placeholder',
  'title',
  'placeholder',
  'alt',
]);

const USER_NAME = /(label|title|tooltip|placeholder|description|heading|caption|hint|text|content|announcement|status)$/i;
const NON_USER_NAME = /^(className|style|id|key|role|type|kind|code|href|src|testId|fontFamily|font|ascii|textContent|contentType|mimeType|text-anchor)$/i;
const DIAGNOSTIC_CALLS = new Set([
  'super', 'log', 'describe', 'test', 'it', 'expect', 'require', 'Symbol', 'RegExp', 'postMessage', 'dispatchEvent',
  'CustomEvent', 'querySelector', 'querySelectorAll', 'matches', 'closest', 'setAttribute', 'getAttribute',
  'removeAttribute', 'createElement', 'addEventListener', 'removeEventListener', 'fetch', 'has', 'get', 'set', 'delete',
  'add', 'includes', 'startsWith', 'endsWith', 'indexOf', 'split', 'join', 'replace', 'replaceAll', 'match',
  'localeCompare', 'getPropertyValue', 'setProperty', 'mark', 'measure', 'lost', 'replaced',
]);
const DIAGNOSTIC_CALL = /^(?!set|show)\w*(fail|refus|reject|unavailable|assert|invariant|abort|warn)|^(?!set|show)\w*Error$/i;
const DIAGNOSTIC_PROPERTY = /^(message|reason|detail|details|error|cause)$/;

function words(text: string): boolean {
  return /\p{L}{2,}/u.test(text);
}

function sentence(text: string): boolean {
  return /\p{L}{2,}[  ]+\p{L}{2,}/u.test(text) || /^\p{Lu}\p{Ll}+[.!?…:]?$/u.test(text.trim());
}

function phrase(text: string): boolean {
  const trimmed = text.trim();
  if (/[{};]|:\s*\S+;|^[\w-]+:|\/\/|=>|^\.|^#|^@|\$\{/.test(trimmed)) return false;
  if (/^[\p{Ll}\d_-]+(\.[\p{L}\d_-]+)+$/u.test(trimmed)) return false;
  return /^\p{Lu}\p{Ll}*[,']?(?:[  ]+[\p{L}\d'’,()/-]+)+[.!?…:]?$/u.test(trimmed);
}

function literalParts(node: ts.Expression): string[] | undefined {
  if (ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node)) return [node.text];
  if (ts.isTemplateExpression(node)) return [node.head.text, ...node.templateSpans.map((span) => span.literal.text)];
  if (ts.isParenthesizedExpression(node) || ts.isAsExpression(node) || ts.isSatisfiesExpression(node)) return literalParts(node.expression);
  if (ts.isConditionalExpression(node)) {
    const parts = [...(literalParts(node.whenTrue) ?? []), ...(literalParts(node.whenFalse) ?? [])];
    return parts.length > 0 ? parts : undefined;
  }
  if (ts.isBinaryExpression(node)) {
    const operator = node.operatorToken.kind;
    if (operator === ts.SyntaxKind.BarBarToken || operator === ts.SyntaxKind.QuestionQuestionToken) return literalParts(node.right);
    if (operator === ts.SyntaxKind.PlusToken) {
      const parts = [...(literalParts(node.left) ?? []), ...(literalParts(node.right) ?? [])];
      return parts.length > 0 ? parts : undefined;
    }
  }
  return undefined;
}

function nameOf(name: ts.Node | undefined): string | undefined {
  if (!name) return undefined;
  if (ts.isIdentifier(name) || ts.isStringLiteral(name) || ts.isPrivateIdentifier(name)) return name.text;
  if (ts.isJsxNamespacedName(name)) return `${name.namespace.text}:${name.name.text}`;
  return undefined;
}

function calleeName(node: ts.CallExpression | ts.NewExpression): string {
  const callee = node.expression;
  if (ts.isIdentifier(callee)) return callee.text;
  if (ts.isPropertyAccessExpression(callee)) {
    const target = callee.expression;
    if (ts.isIdentifier(target) && target.text === 'console') return 'log';
    return callee.name.text;
  }
  if (callee.kind === ts.SyntaxKind.SuperKeyword) return 'super';
  return '';
}

function inDiagnostic(node: ts.Node): boolean {
  for (let current: ts.Node = node; current.parent; current = current.parent) {
    const parent = current.parent;
    if (ts.isThrowStatement(parent)) return true;
    if ((ts.isCallExpression(parent) || ts.isNewExpression(parent)) && parent.expression !== current) {
      const callee = calleeName(parent);
      if (DIAGNOSTIC_CALLS.has(callee) || DIAGNOSTIC_CALL.test(callee)) return true;
    }
    if (ts.isPropertyAssignment(parent) && parent.initializer === current && ts.isObjectLiteralExpression(parent.parent)) {
      const siblings = parent.parent.properties.map((property) => (property.name ? nameOf(property.name) : undefined) ?? '');
      if (DIAGNOSTIC_PROPERTY.test(nameOf(parent.name) ?? '') && siblings.some((name) => /^(code|ok|failure)$/.test(name))) return true;
    }
    if (ts.isFunctionLike(parent) || ts.isSourceFile(parent)) return false;
  }
  return false;
}

function keyedSibling(node: ts.Node): boolean {
  const assignment = node.parent;
  if (!ts.isPropertyAssignment(assignment) || !ts.isObjectLiteralExpression(assignment.parent)) return false;
  const name = nameOf(assignment.name);
  const siblings = assignment.parent.properties.map((property) => (property.name ? nameOf(property.name) : undefined) ?? '');
  return siblings.includes(`${name}Key`) || siblings.some((sibling) => /^[a-z]+(-[a-z]+)+$/.test(sibling));
}

function diagnosticVariable(node: ts.Node): boolean {
  for (let current: ts.Node = node; current.parent; current = current.parent) {
    const parent = current.parent;
    if (ts.isVariableDeclaration(parent) && parent.initializer === current) return DIAGNOSTIC_PROPERTY.test(nameOf(parent.name) ?? '');
    if (!(ts.isConditionalExpression(parent) || ts.isBinaryExpression(parent) || ts.isParenthesizedExpression(parent))) return false;
  }
  return false;
}

function nonValuePosition(node: ts.StringLiteralLike): boolean {
  const parent = node.parent;
  if (ts.isLiteralTypeNode(parent) || ts.isImportDeclaration(parent) || ts.isExportDeclaration(parent)) return true;
  if (ts.isExternalModuleReference(parent) || ts.isImportTypeNode(parent.parent)) return true;
  if ((ts.isPropertyAssignment(parent) || ts.isPropertySignature(parent) || ts.isMethodDeclaration(parent)) && parent.name === node) return true;
  if (ts.isElementAccessExpression(parent) && parent.argumentExpression === node) return true;
  if (ts.isCaseClause(parent)) return true;
  if (ts.isBinaryExpression(parent)) {
    const operator = parent.operatorToken.kind;
    if (operator === ts.SyntaxKind.EqualsEqualsEqualsToken || operator === ts.SyntaxKind.ExclamationEqualsEqualsToken
      || operator === ts.SyntaxKind.EqualsEqualsToken || operator === ts.SyntaxKind.ExclamationEqualsToken
      || operator === ts.SyntaxKind.InKeyword) return true;
  }
  if (ts.isJsxAttribute(parent)) return true;
  return false;
}

function enclosingName(node: ts.Node): string | undefined {
  for (let current: ts.Node = node; current.parent; current = current.parent) {
    const parent = current.parent;
    if (ts.isPropertyAssignment(parent) && parent.initializer === current) return nameOf(parent.name);
    if (ts.isVariableDeclaration(parent) && parent.initializer === current) return nameOf(parent.name);
    if (ts.isBindingElement(parent) && parent.initializer === current) return nameOf(parent.name);
    if (ts.isParameter(parent) && parent.initializer === current) return nameOf(parent.name);
    if (ts.isJsxAttribute(parent)) return nameOf(parent.name);
    if (ts.isReturnStatement(parent)) {
      for (let fn: ts.Node | undefined = parent; fn; fn = fn.parent) {
        if (ts.isFunctionDeclaration(fn) || ts.isMethodDeclaration(fn)) return nameOf(fn.name);
        if ((ts.isArrowFunction(fn) || ts.isFunctionExpression(fn)) && ts.isVariableDeclaration(fn.parent)) return nameOf(fn.parent.name);
      }
      return undefined;
    }
    if (ts.isArrowFunction(parent) && parent.body === current && ts.isVariableDeclaration(parent.parent)) return nameOf(parent.parent.name);
    if (!(ts.isConditionalExpression(parent) || ts.isBinaryExpression(parent) || ts.isParenthesizedExpression(parent)
      || ts.isAsExpression(parent) || ts.isSatisfiesExpression(parent) || ts.isTemplateSpan(parent) || ts.isTemplateExpression(parent)
      || ts.isJsxExpression(parent))) return undefined;
  }
  return undefined;
}

function insideStyle(node: ts.Node): boolean {
  for (let current: ts.Node | undefined = node.parent; current; current = current.parent) {
    if (ts.isJsxElement(current)) {
      const tag = current.openingElement.tagName.getText();
      if (tag === 'style' || tag === 'script' || tag === 'code') return true;
    }
  }
  return false;
}

function ignored(source: ts.SourceFile, node: ts.Node): boolean {
  const { line } = source.getLineAndCharacterOfPosition(node.getStart(source));
  const lines = source.text.split('\n');
  return [lines[line], lines[line - 1]].some((text) => text?.includes('i18n-ignore'));
}

export function scanSource(file: string, text: string): HardcodedString[] {
  const source = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true, file.endsWith('.tsx') ? ts.ScriptKind.TSX : ts.ScriptKind.TS);
  const found: HardcodedString[] = [];
  const seen = new Set<ts.Node>();
  const reported: ts.Node[] = [];
  const report = (node: ts.Node, kind: HardcodedString['kind'], name: string, value: string) => {
    if (seen.has(node) || ignored(source, node)) return;
    if (reported.some((outer) => node.pos >= outer.pos && node.end <= outer.end)) return;
    seen.add(node);
    reported.push(node);
    const line = source.getLineAndCharacterOfPosition(node.getStart(source)).line + 1;
    found.push({ file, line, kind, name, text: value.trim().replace(/\s+/g, ' ') });
  };
  const checkLiteral = (node: ts.StringLiteral | ts.NoSubstitutionTemplateLiteral | ts.TemplateExpression) => {
    if (!ts.isTemplateExpression(node) && nonValuePosition(node)) return;
    const value = literalParts(node)?.join(' ') ?? '';
    if (!words(value) || inDiagnostic(node) || keyedSibling(node) || diagnosticVariable(node)) return;
    const name = enclosingName(node);
    if (name && (NON_USER_NAME.test(name) || name.includes('-') || /family|font/i.test(name))) return;
    if (name && USER_NAME.test(name) && sentence(value)) report(node, 'named', name, value);
    else if (phrase(value)) report(node, 'literal', name ?? '', value);
  };
  const visit = (node: ts.Node) => {
    if (ts.isJsxText(node)) {
      const value = node.text.replace(/&\w+;/g, '');
      if (words(value) && !insideStyle(node)) report(node, 'jsx-text', '', node.text);
    } else if (ts.isJsxAttribute(node) && node.initializer) {
      const name = nameOf(node.name) ?? '';
      const init = node.initializer;
      const expression = ts.isStringLiteral(init) ? init : ts.isJsxExpression(init) ? init.expression : undefined;
      const value = expression ? literalParts(expression)?.join(' ') ?? '' : '';
      if (USER_ATTRIBUTES.has(name) ? words(value) : USER_NAME.test(name) && !NON_USER_NAME.test(name) && words(value) && sentence(value)) {
        report(node, 'jsx-attribute', name, value);
      }
    } else if (ts.isJsxExpression(node) && node.expression && (ts.isJsxElement(node.parent) || ts.isJsxFragment(node.parent))) {
      const value = literalParts(node.expression)?.join(' ') ?? '';
      if (words(value) && sentence(value) && !insideStyle(node)) report(node, 'jsx-text', '', value);
    } else if ((ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node) || ts.isTemplateExpression(node)) && !ts.isJsxAttribute(node.parent)) {
      checkLiteral(node);
    }
    if (ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression)) {
      const method = node.expression.name.text;
      const parent = node.parent;
      const shown = ts.isTemplateSpan(parent) || (ts.isBinaryExpression(parent) && parent.operatorToken.kind === ts.SyntaxKind.PlusToken);
      if (method === 'toFixed' && shown) report(node, 'format', method, node.getText(source));
      if (/^toLocale(Date|Time)?String$/.test(method) && !node.arguments.some((argument) => ts.isStringLiteral(argument) || ts.isIdentifier(argument))) {
        report(node, 'format', method, node.getText(source));
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(source);
  return found;
}

const SKIPPED_FILE = /(\.test|\.spec|\.stories|\.bench|\.fixture|\.testing)\.tsx?$|\.d\.ts$|^testing\.ts$|TestHarness\.tsx?$/;
const SKIPPED_DIR = new Set(['node_modules', 'generated', 'dist', 'testing', '__tests__', '__fixtures__', 'fixtures', 'test', 'tests']);

export function sourceFiles(directory: string): string[] {
  if (statSync(directory).isFile()) return [directory];
  const files: string[] = [];
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const path = join(directory, entry.name);
    if (entry.isDirectory()) {
      if (!SKIPPED_DIR.has(entry.name)) files.push(...sourceFiles(path));
    } else if (/\.tsx?$/.test(entry.name) && !SKIPPED_FILE.test(entry.name)) {
      files.push(path);
    }
  }
  return files.sort();
}

export interface AllowedString {
  text: string;
  file?: string;
  reason: string;
}

export function allowedStrings(root: string): AllowedString[] {
  return JSON.parse(readFileSync(join(root, 'scripts/hardcoded-ui-strings.allowlist.json'), 'utf8')) as AllowedString[];
}

export function allows(entry: AllowedString, item: HardcodedString): boolean {
  return entry.text === item.text && (entry.file === undefined || entry.file === item.file);
}

export function scanDirectories(root: string, directories: string[], allowlist: AllowedString[] = []): HardcodedString[] {
  return directories
    .flatMap((directory) =>
      sourceFiles(join(root, directory)).flatMap((file) => scanSource(relative(root, file), readFileSync(file, 'utf8'))),
    )
    .filter((item) => !allowlist.some((entry) => allows(entry, item)));
}

if (import.meta.main) {
  const root = join(import.meta.dir, '..');
  const found = scanDirectories(root, process.argv.slice(2), allowedStrings(root));
  for (const item of found) console.log(`${item.file}:${item.line}\t${item.kind}\t${item.name}\t${item.text}`);
  console.error(`${found.length} hardcoded strings`);
}
