import { expect, test } from 'bun:test';
import { fileURLToPath } from 'node:url';
import {
  declarationCandidates,
  packageName,
  publishedEntries,
} from '../build-core-declarations.mjs';
import { siblingsOf } from '../check-built-siblings.mjs';
import { declarationCompilerOptions } from '../declaration-options.mjs';

const paths = {
  '@docx-editor.dev/core': ['./src/index.ts'],
  '@docx-editor.dev/core/store': ['./src/store/index.ts'],
};

test('each export with a top-level types condition becomes one declaration entry', () => {
  const entries = publishedEntries(
    {
      exports: {
        '.': { types: './dist/index.d.ts', import: './dist/index.mjs' },
        './store': { types: './dist/store.d.ts', import: './dist/store.mjs' },
        './editor.css': './dist/editor.css',
        './package.json': './package.json',
      },
    },
    paths
  );
  expect(entries.map(({ specifier, name }) => [specifier, name])).toEqual([
    ['@docx-editor.dev/core', 'index'],
    ['@docx-editor.dev/core/store', 'store'],
  ]);
  expect(entries[1]!.source.endsWith('/packages/core/src/store/index.ts')).toBe(true);
});

test('a subpath without a source path or with nested types fails instead of shipping no types', () => {
  expect(() =>
    publishedEntries({ exports: { './layout': { types: './dist/layout.d.ts' } } }, paths)
  ).toThrow('has no path for its source');
  expect(() =>
    publishedEntries(
      { exports: { './store': { import: { types: './dist/store.d.ts', default: './x.mjs' } } } },
      paths
    )
  ).toThrow('nests its types condition');
});

test('a bare import is checked against the package that owns it', () => {
  expect(packageName('prosemirror-model')).toBe('prosemirror-model');
  expect(packageName('fflate/browser')).toBe('fflate');
  expect(packageName('@docx-editor.dev/i18n/locales')).toBe('@docx-editor.dev/i18n');
});

test('an import keeps its module kind when it resolves to a declaration file', () => {
  expect(declarationCandidates('/o/foo.ts')).toEqual(['/o/foo.d.ts']);
  expect(declarationCandidates('/o/foo.js')).toEqual(['/o/foo.d.ts']);
  expect(declarationCandidates('/o/foo.mts')).toEqual(['/o/foo.d.mts']);
  expect(declarationCandidates('/o/foo.cjs')).toEqual(['/o/foo.d.cts']);
  expect(declarationCandidates('/o/foo')).toEqual(['/o/foo.d.ts', '/o/foo/index.d.ts']);
});

test('declaration builds drop sibling source paths and keep local aliases', () => {
  // Vue reads core and i18n as siblings, so it reads both from dist.
  const vue = declarationCompilerOptions(
    new URL('../../packages/vue/tsup.config.ts', import.meta.url)
  );
  expect(vue.paths).toEqual({});
  // Fonts has core only as a development dependency, so it keeps core's source paths.
  const fonts = declarationCompilerOptions(
    new URL('../../packages/fonts/tsup.config.ts', import.meta.url)
  );
  expect(Object.keys(fonts.paths)).toContain('@docx-editor.dev/core');
});

test('a package checks every sibling it reads, including those behind core', () => {
  const siblings = (name: string) =>
    [
      ...siblingsOf(fileURLToPath(new URL(`../../packages/${name}`, import.meta.url))).keys(),
    ].sort();
  expect(siblings('vue')).toEqual(['@docx-editor.dev/core', '@docx-editor.dev/i18n']);
  expect(siblings('fonts')).toEqual([]);
});
