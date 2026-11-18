import { expect, test } from 'bun:test';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  checkLedger,
  ledgerExports,
  parseApiReport,
  parseLedger,
  untaggedAngularExports,
} from '../check-angular-parity.mjs';

const REPORT = [
  '```ts',
  "import { blankDocumentBytes } from '@docx-editor.dev/core/editor';",
  '',
  'export { blankDocumentBytes }',
  '',
  '// @public',
  'export const DocxEditor: DocxEditorNamespace;',
  '',
  '// @public @deprecated (undocumented)',
  'export const Toolbar: unknown;',
  '',
  '// @internal',
  'export function notificationYieldsToTask(): boolean;',
  '',
  '// @public (undocumented)',
  'export interface DocxEditorProps {',
  '    document?: ArrayBuffer;',
  '}',
  '```',
].join('\n');

function ledger(exportRows: string[], testRows: string[]) {
  return parseLedger(
    [
      '# Ledger',
      '',
      '## Exports',
      '',
      '| Vue export | Kind | Angular counterpart | Status | Notes |',
      '| --- | --- | --- | --- | --- |',
      ...exportRows,
      '',
      '## Tests',
      '',
      '| Vue test | Angular spec | Status | Notes |',
      '| --- | --- | --- | --- |',
      ...testRows,
    ].join('\n')
  );
}

const exports = ledgerExports(parseApiReport(REPORT));
const tests = ['packages/vue/test/live-props.test.ts'];
const complete = {
  exports: [
    '| `blankDocumentBytes` | re-export from core | | todo | |',
    '| `DocxEditor` | const | `DocxEditorComponent` | ported | |',
    '| `DocxEditorProps` | interface | | dropped | Angular inputs replace the props object. |',
  ],
  tests: ['| `packages/vue/test/live-props.test.ts` | | todo | Task 3.1 |'],
};

test('the API report gives the exports that need a row, without deprecated or internal ones', () => {
  expect(exports).toEqual([
    { name: 'blankDocumentBytes', kind: 're-export', tags: [] },
    { name: 'DocxEditor', kind: 'const', tags: ['@public'] },
    { name: 'DocxEditorProps', kind: 'interface', tags: ['@public'] },
  ]);
});

test('a complete ledger passes, and --no-todo fails on each todo row', () => {
  const parsed = ledger(complete.exports, complete.tests);
  expect(checkLedger({ exports, tests, ledger: parsed })).toEqual([]);
  expect(checkLedger({ exports, tests, ledger: parsed, noTodo: true })).toEqual([
    'export row `blankDocumentBytes` is still todo',
    'test row `packages/vue/test/live-props.test.ts` is still todo',
  ]);
});

test('a missing export row and a missing test row fail', () => {
  const parsed = ledger(complete.exports.slice(1), []);
  expect(checkLedger({ exports, tests, ledger: parsed })).toEqual([
    'Vue export `blankDocumentBytes` has no row',
    'Vue test `packages/vue/test/live-props.test.ts` has no row',
  ]);
});

test('a dropped row without a reason fails', () => {
  const rows = [...complete.exports.slice(0, 2), '| `DocxEditorProps` | interface | | dropped | |'];
  expect(checkLedger({ exports, tests, ledger: ledger(rows, complete.tests) })).toEqual([
    'export row `DocxEditorProps` is dropped but gives no reason',
  ]);
});

test('a ported row without a counterpart, an unknown status, and a duplicate row fail', () => {
  const rows = [
    '| `blankDocumentBytes` | re-export from core | | done | |',
    '| `DocxEditor` | const | — | ported | |',
    complete.exports[2],
    complete.exports[2],
  ];
  expect(checkLedger({ exports, tests, ledger: ledger(rows, complete.tests) })).toEqual([
    'ledger has two rows for `DocxEditorProps`',
    'export row `blankDocumentBytes` has status "done", not one of todo, ported, dropped',
    'export row `DocxEditor` is ported but names no Angular counterpart',
  ]);
});

test('an Angular export without a TSDoc release tag fails', () => {
  const root = mkdtempSync(join(tmpdir(), 'angular-parity-'));
  try {
    mkdirSync(join(root, 'lib'));
    writeFileSync(
      join(root, 'public-api.ts'),
      [
        "export { Tagged, Untagged, type Shape } from './lib/parts';",
        "export * from './lib/more';",
        "export type { DocumentChange } from '@docx-editor.dev/core/contracts/editor';",
      ].join('\n')
    );
    writeFileSync(
      join(root, 'lib', 'parts.ts'),
      [
        '/**',
        ' * Tagged.',
        ' *',
        ' * @public',
        ' */',
        'export class Tagged {}',
        '/** No tag. */',
        'export class Untagged {}',
        'export type Shape = string;',
        'export class NotReexported {}',
      ].join('\n')
    );
    writeFileSync(
      join(root, 'lib', 'more.ts'),
      ['/** @beta */', 'export const tagged = 1;', 'export function untagged() {}'].join('\n')
    );
    const names = untaggedAngularExports(join(root, 'public-api.ts')).map(
      (entry: string) => entry.split(' ')[0]
    );
    expect(names).toEqual(['Shape', 'Untagged', 'untagged']);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});
