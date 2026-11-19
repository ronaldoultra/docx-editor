import type { EditorModule } from '@docx-editor.dev/core/editor';

/**
 * A minimal review module. The `'suggesting'` mode needs a review module, and this one is enough
 * to record edits as tracked changes. It lists no review items, so the example shows no review
 * pane.
 */
export const EXAMPLE_REVIEW_MODULE: EditorModule = {
  id: 'review',
  review: {
    displayModes: ['all-markup', 'proposed', 'original'],
    collectReviewItems: () => [],
    revisionItemsOfParagraph: () => [],
  },
};
