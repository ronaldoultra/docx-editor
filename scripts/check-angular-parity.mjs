#!/usr/bin/env node
// Check the Angular adapter parity ledger.
//
//   bun run check:angular-parity
//
// `openspec/changes/angular-adapter-parity/parity-ledger.md` has one row for each Vue export
// that is not deprecated and not internal, and one row for each Vue test file. Each row states
// its Angular counterpart and a status: `todo`, `ported`, or `dropped`. The Vue package can leave
// only when no row is `todo`.
//
// The check fails when:
// - a Vue export or a Vue test file has no ledger row,
// - a row has a status that is not `todo`, `ported`, or `dropped`,
// - a `ported` row names no Angular counterpart, or a `dropped` row gives no reason,
// - an export of `@docx-editor.dev/angular` has no TSDoc release tag.
//
// `--no-todo` also fails on each `todo` row. Phase 10 uses it as the removal gate.

import { existsSync, readFileSync, readdirSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
export const LEDGER = 'openspec/changes/angular-adapter-parity/parity-ledger.md';
export const VUE_API = 'docs/api/docx-editor-vue/index.api.md';
export const VUE_TESTS = 'packages/vue/test';
export const ANGULAR_ENTRY = 'packages/angular/src/public-api.ts';

export const STATUSES = ['todo', 'ported', 'dropped'];
const RELEASE_TAGS = ['@public', '@beta', '@alpha', '@internal'];
const DECLARATION =
  /^export (?:declare )?(?:abstract )?(const|function|interface|type|class|enum|namespace) ([A-Za-z_$][\w$]*)/;
const REEXPORT = /^export \{ ([A-Za-z_$][\w$]*) \}$/;

/**
 * The exports in an API Extractor report, in report order. Each entry has `name`, `kind`, and
 * `tags`, the TSDoc tags from the `// @…` line above a declaration. A re-export has kind
 * `re-export` and no tags, because the report does not print the tags of a re-exported symbol.
 */
export function parseApiReport(text) {
  const exports = [];
  const seen = new Set();
  let tags = [];
  for (const line of text.split(/\r?\n/)) {
    if (line.startsWith('// @')) {
      tags = line
        .slice(3)
        .split(/\s+/)
        .filter((word) => word.startsWith('@'));
      continue;
    }
    const declared = DECLARATION.exec(line);
    const reexported = declared ? null : REEXPORT.exec(line);
    const name = declared?.[2] ?? reexported?.[1];
    if (name && !seen.has(name)) {
      seen.add(name);
      exports.push({ name, kind: declared ? declared[1] : 're-export', tags });
    }
    if (line.trim() !== '') tags = [];
  }
  return exports;
}

/** The exports that need a ledger row: not `@deprecated` and not `@internal`. */
export function ledgerExports(exports) {
  return exports.filter(
    (entry) => !entry.tags.includes('@deprecated') && !entry.tags.includes('@internal')
  );
}

/** The repository-relative test files under `directory`, sorted, with `/` separators. */
export function testFiles(directory, root = ROOT) {
  const found = [];
  const walk = (folder) => {
    for (const entry of readdirSync(folder, { withFileTypes: true })) {
      const path = join(folder, entry.name);
      if (entry.isDirectory()) walk(path);
      else if (/\.test\.(ts|tsx)$/.test(entry.name)) {
        found.push(relative(root, path).replaceAll('\\', '/'));
      }
    }
  };
  walk(join(root, directory));
  return found.sort();
}

function tableCells(line) {
  const trimmed = line.trim();
  return trimmed
    .slice(1, trimmed.endsWith('|') ? -1 : undefined)
    .split('|')
    .map((cell) => cell.trim());
}

/**
 * The ledger rows, keyed by section (`## Exports` and `## Tests`). A row is a table line whose
 * first cell is a code span. The columns are found by their header: `Angular…`, `Status`, and
 * `Notes`. Returns `{ exports: Map<key, row>, tests: Map<key, row>, duplicates: string[] }`.
 */
export function parseLedger(text) {
  const sections = { exports: new Map(), tests: new Map() };
  const duplicates = [];
  let section = null;
  let columns = null;
  for (const line of text.split(/\r?\n/)) {
    const heading = /^## (.+)$/.exec(line);
    if (heading) {
      const title = heading[1].trim().toLowerCase();
      section = title.startsWith('export') ? 'exports' : title.startsWith('test') ? 'tests' : null;
      columns = null;
      continue;
    }
    if (!section || !line.trim().startsWith('|')) continue;
    const cells = tableCells(line);
    const key = /^`([^`]+)`$/.exec(cells[0] ?? '')?.[1];
    if (!key) {
      const headers = cells.map((cell) => cell.toLowerCase());
      if (headers.includes('status')) {
        columns = {
          angular: headers.findIndex((cell) => cell.startsWith('angular')),
          status: headers.indexOf('status'),
          notes: headers.indexOf('notes'),
        };
      }
      continue;
    }
    if (!columns) continue;
    const cell = (index) => (index >= 0 ? (cells[index] ?? '') : '');
    const row = {
      key,
      angular: cell(columns.angular),
      status: cell(columns.status),
      notes: cell(columns.notes),
    };
    if (sections[section].has(key)) duplicates.push(key);
    sections[section].set(key, row);
  }
  return { ...sections, duplicates };
}

/** A cell that says nothing: empty, or a dash. */
function blank(cell) {
  return cell === '' || cell === '-' || cell === '—';
}

/** The problems in one ledger row, as messages. */
function rowProblems(section, row, noTodo) {
  const where = `${section} row \`${row.key}\``;
  if (!STATUSES.includes(row.status)) {
    return [`${where} has status "${row.status}", not one of ${STATUSES.join(', ')}`];
  }
  if (row.status === 'ported' && blank(row.angular)) {
    return [`${where} is ported but names no Angular counterpart`];
  }
  if (row.status === 'dropped' && blank(row.notes)) {
    return [`${where} is dropped but gives no reason`];
  }
  if (row.status === 'todo' && noTodo) return [`${where} is still todo`];
  return [];
}

/**
 * Compare the ledger with the Vue exports and test files. Returns a list of messages; an empty
 * list means the ledger is complete.
 */
export function checkLedger({ exports, tests, ledger, noTodo = false }) {
  const problems = ledger.duplicates.map((key) => `ledger has two rows for \`${key}\``);
  for (const entry of exports) {
    if (!ledger.exports.has(entry.name)) problems.push(`Vue export \`${entry.name}\` has no row`);
  }
  for (const file of tests) {
    if (!ledger.tests.has(file)) problems.push(`Vue test \`${file}\` has no row`);
  }
  for (const row of ledger.exports.values()) problems.push(...rowProblems('export', row, noTodo));
  for (const row of ledger.tests.values()) problems.push(...rowProblems('test', row, noTodo));
  return problems;
}

/** The TSDoc release tags in the doc comments directly above `node`. */
function releaseTags(node, source) {
  const tags = [];
  for (const range of ts.getLeadingCommentRanges(source.text, node.getFullStart()) ?? []) {
    const comment = source.text.slice(range.pos, range.end);
    if (!comment.startsWith('/**')) continue;
    for (const tag of RELEASE_TAGS) if (new RegExp(`${tag}\\b`).test(comment)) tags.push(tag);
  }
  return tags;
}

/** The top-level declarations of a file, by name. */
function declarations(source) {
  const found = new Map();
  for (const statement of source.statements) {
    if (ts.isVariableStatement(statement)) {
      for (const declaration of statement.declarationList.declarations) {
        if (ts.isIdentifier(declaration.name)) found.set(declaration.name.text, statement);
      }
    } else if ('name' in statement && statement.name && ts.isIdentifier(statement.name)) {
      found.set(statement.name.text, statement);
    }
  }
  return found;
}

function readSource(file) {
  return ts.createSourceFile(file, readFileSync(file, 'utf8'), ts.ScriptTarget.Latest, true);
}

/** `specifier` from `from`, as a `.ts` file, or null when it leaves the package. */
function resolveLocal(from, specifier) {
  if (!specifier.startsWith('.')) return null;
  const base = resolve(dirname(from), specifier).replace(/\.(m?js)$/, '');
  for (const candidate of [`${base}.ts`, join(base, 'index.ts')]) {
    if (existsSync(candidate)) return candidate;
  }
  return null;
}

/**
 * The exports of the Angular entry that carry no TSDoc release tag. Re-exports from other
 * packages are skipped: their own API reports hold their tags.
 */
export function untaggedAngularExports(entry = join(ROOT, ANGULAR_ENTRY)) {
  const untagged = [];
  const visit = (file, names, trail) => {
    const source = readSource(file);
    const local = declarations(source);
    for (const statement of source.statements) {
      if (!ts.isExportDeclaration(statement)) continue;
      const specifier = statement.moduleSpecifier;
      const target =
        specifier && ts.isStringLiteral(specifier) ? resolveLocal(file, specifier.text) : null;
      if (specifier && !target) continue;
      const clause = statement.exportClause;
      if (!clause) {
        visit(target, names, trail);
        continue;
      }
      if (!ts.isNamedExports(clause)) continue;
      for (const element of clause.elements) {
        const exported = element.name.text;
        if (names && !names.has(exported)) continue;
        const original = (element.propertyName ?? element.name).text;
        if (target) visit(target, new Set([original]), [...trail, exported]);
        else check(local.get(original), source, exported, file);
      }
    }
    for (const [name, node] of local) {
      const isExported = node.modifiers?.some(
        (modifier) => modifier.kind === ts.SyntaxKind.ExportKeyword
      );
      if (isExported && (!names || names.has(name))) check(node, source, name, file);
    }
  };
  const check = (node, source, name, file) => {
    if (!node || releaseTags(node, source).length > 0) return;
    untagged.push(`${name} (${relative(ROOT, file).replaceAll('\\', '/')})`);
  };
  visit(entry, null, []);
  return [...new Set(untagged)].sort();
}

function main(argv) {
  const noTodo = argv.includes('--no-todo');
  const exports = ledgerExports(parseApiReport(readFileSync(join(ROOT, VUE_API), 'utf8')));
  const tests = testFiles(VUE_TESTS);
  const ledger = parseLedger(readFileSync(join(ROOT, LEDGER), 'utf8'));
  const problems = checkLedger({ exports, tests, ledger, noTodo });
  for (const name of untaggedAngularExports()) {
    problems.push(`Angular export ${name} has no TSDoc release tag`);
  }
  const rows = [...ledger.exports.values(), ...ledger.tests.values()];
  const count = (status) => rows.filter((row) => row.status === status).length;
  console.log(
    `Parity ledger: ${rows.length} rows (${count('ported')} ported, ${count('dropped')} dropped, ` +
      `${count('todo')} todo) for ${exports.length} Vue exports and ${tests.length} Vue tests.`
  );
  if (problems.length === 0) return 0;
  for (const problem of problems) console.error(`  ${problem}`);
  console.error(`check:angular-parity found ${problems.length} problems.`);
  return 1;
}

if (process.argv[1] === fileURLToPath(import.meta.url)) process.exit(main(process.argv.slice(2)));
