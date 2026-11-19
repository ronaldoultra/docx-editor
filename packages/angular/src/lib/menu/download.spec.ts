// Task 5.1: the claims of packages/vue/test/menu-download-name.test.ts.
import { describe, expect, test } from 'vitest';
import { downloadName } from './download';

describe('downloadName', () => {
  test('a long run of dots and spaces before the end stays linear', () => {
    expect(downloadName(`a${'. '.repeat(100_000)}b`)).toBe('a.docx');
    expect(downloadName(`a${'. '.repeat(100_000)}`)).toBe('a.docx');
  });

  test('a device name with an extension and a padded suffix are handled', () => {
    expect(downloadName('NUL')).toBe('document.docx');
    expect(downloadName('NUL.tar')).toBe('_NUL.tar.docx');
    expect(downloadName('report.docx ')).toBe('report.docx');
  });

  test('truncation that ends on a dot or a space does not leave one', () => {
    expect(downloadName(`${'a'.repeat(199)}.b`)).toBe(`${'a'.repeat(199)}.docx`);
    expect(downloadName(`${'a'.repeat(199)} b`)).toBe(`${'a'.repeat(199)}.docx`);
  });
});
