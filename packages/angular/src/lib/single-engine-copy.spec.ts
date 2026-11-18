import * as root from '@docx-editor.dev/core';
import * as editor from '@docx-editor.dev/core/editor';
import * as output from '@docx-editor.dev/core/output';

// The engine holds module-level state: the HarfBuzz shaper and its cache budget, and layout
// caches keyed by object identity. A second copy in the test bundle does not crash. It loads the
// shaper twice and misses every identity-keyed cache, so the specs would test a setup that no
// application has. Each pair below is one function that two core subpaths export from the same
// source file. One copy of the engine gives one object for each pair.
//
// Do not add `createLayoutShaping` here: `/editor` and `/layout` export two different functions
// with that name.
describe('the Angular tests load one copy of @docx-editor.dev/core', () => {
  it('shares the editor entry with the package root', () => {
    expect(editor.createDocxEditor).toBe(root.createDocxEditor);
    expect(editor.blankDocumentBytes).toBe(root.blankDocumentBytes);
    expect(editor.commandForSlot).toBe(root.commandForSlot);
  });

  it('shares the paint lane with the editor entry', () => {
    expect(editor.observeUnsupportedElements).toBe(output.observeUnsupportedElements);
  });
});
