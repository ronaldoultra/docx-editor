import { expect, test } from 'bun:test';
import { angularReminder, selectTests } from '../test/changed.mjs';
import { ownedByOtherRunner } from '../test/other-runners.mjs';

// Bun cannot compile Angular components, so `bun run test` and `bun run test:changed` leave the
// specs in `packages/angular` to `bun run test:angular`.

test('the Angular specs belong to another runner, on either path separator', () => {
  expect(ownedByOtherRunner('packages/angular/src/lib/docx-editor.component.spec.ts')).toBe(true);
  expect(ownedByOtherRunner('packages\\angular\\src\\lib\\docx-editor.component.spec.ts')).toBe(
    true
  );
  expect(ownedByOtherRunner('packages/core/src/__tests__/a.test.ts')).toBe(false);
  expect(ownedByOtherRunner('packages/angular-extra/a.test.ts')).toBe(false);
});

test('test:changed never selects an Angular spec', () => {
  const files = [
    'packages/core/src/editor/a.ts',
    'packages/core/src/editor/__tests__/a.test.ts',
    'packages/angular/src/lib/docx-editor.component.spec.ts',
  ];
  const reverse = new Map([
    [
      'packages/core/src/editor/a.ts',
      new Set([files[1], 'packages/angular/src/lib/docx-editor.component.spec.ts']),
    ],
  ]);
  expect(selectTests(['packages/core/src/editor/a.ts'], files, reverse, () => '')).toEqual({
    all: false,
    tests: ['packages/core/src/editor/__tests__/a.test.ts'],
  });
});

test('test:changed names test:angular for Angular and core editor changes only', () => {
  expect(angularReminder(['packages/angular/src/public-api.ts'])).toContain('bun run test:angular');
  expect(angularReminder(['packages/core/src/editor/docx-editor.ts'])).toContain(
    'bun run test:angular'
  );
  expect(angularReminder(['packages/core/src/layout/semantic-layout.ts'])).toBeNull();
  expect(angularReminder(['docs/guide.md'])).toBeNull();
});
