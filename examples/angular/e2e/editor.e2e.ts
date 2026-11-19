// Browser tests for <docx-editor> in the example app. Each case runs twice: without zone.js
// (the default) and with zone.js (`?zone`).
import { expect, test, type Page } from '@playwright/test';
import { unzipSync, strFromU8 } from 'fflate';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const fixture = (path: string) => fileURLToPath(new URL(path, import.meta.url));
const SAMPLE = fixture('../../../packages/core/src/editor/__tests__/fixtures/sample.docx');
const OTHER = fixture('../../../e2e/fixtures/styled-sample.docx');
const HEADING = '/word/document.xml#0.0.0';
const SECOND = '/word/document.xml#0.0.1';
const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';

/** The painted text of one paragraph, read from the engine's text runs. */
function paragraphText(page: Page, id: string): Promise<string> {
  return page.evaluate(
    (paragraphId) =>
      [...document.querySelectorAll(`[data-paragraph-id="${paragraphId}"][data-start]`)]
        .map((node) => node.textContent ?? '')
        .join(''),
    id
  );
}

/** The number of `change` outputs the app received, from its status line. */
async function changeCount(page: Page): Promise<number> {
  const status = (await page.getByTestId('status').textContent()) ?? '';
  return Number(/(\d+) alterações/.exec(status)?.[1] ?? Number.NaN);
}

async function open(page: Page, file: string) {
  const before = await changeCount(page).catch(() => 0);
  await page.setInputFiles('input[type=file]', file);
  await page.locator('.docx-page').first().waitFor();
  await expect(page.getByTestId('status')).toContainText('última: load');
  return before;
}

/** Puts the caret at the start of a paragraph by clicking its first text run. */
async function caretAtStart(page: Page, id: string) {
  await page.locator(`[data-paragraph-id="${id}"][data-start="0"]`).first().click();
  await page.keyboard.press('Home');
}

for (const mode of [
  { name: 'without zone.js', url: '/', status: 'sem zone.js' },
  { name: 'with zone.js', url: '/?zone', status: 'com zone.js' },
]) {
  test.describe(mode.name, () => {
    const errors: string[] = [];

    test.beforeEach(async ({ page }) => {
      errors.length = 0;
      page.on('pageerror', (error) => errors.push(error.message));
      page.on('console', (message) => {
        if (message.type() === 'error') errors.push(message.text());
      });
      await page.goto(mode.url);
      await expect(page.getByTestId('status')).toContainText(mode.status);
    });

    test.afterEach(() => {
      expect(errors).toEqual([]);
    });

    test('loads a document and stacks its pages in a scrolling viewport', async ({ page }) => {
      await open(page, SAMPLE);
      expect(await paragraphText(page, HEADING)).toBe('DOCX-EDITOR.DEV');
      const layout = await page.evaluate(() => {
        const pages = [...document.querySelectorAll('.docx-page')].map((node) =>
          node.getBoundingClientRect()
        );
        const scroller = document.querySelector('docx-editor .docx-editor__scroll-container')!;
        return {
          count: pages.length,
          secondBelowFirst: pages[1]!.top >= pages[0]!.bottom,
          sameLeft: Math.abs(pages[1]!.left - pages[0]!.left) < 1,
          scrolls: scroller.scrollHeight > scroller.clientHeight,
        };
      });
      expect(layout).toEqual({
        count: layout.count,
        secondBelowFirst: true,
        sameLeft: true,
        scrolls: true,
      });
      expect(layout.count).toBeGreaterThan(1);

      const scroller = page.locator('docx-editor .docx-editor__scroll-container');
      await scroller.evaluate((node) => node.scrollTo({ top: 3000 }));
      expect(await scroller.evaluate((node) => node.scrollTop)).toBeGreaterThan(0);
    });

    test('types, deletes, splits, selects, undoes, and redoes', async ({ page }) => {
      await open(page, SAMPLE);
      const loaded = await changeCount(page);

      await caretAtStart(page, HEADING);
      await page.keyboard.type('Abc ');
      await expect.poll(() => paragraphText(page, HEADING)).toBe('Abc DOCX-EDITOR.DEV');
      expect(await changeCount(page)).toBeGreaterThan(loaded);

      await page.keyboard.press('Backspace');
      await expect.poll(() => paragraphText(page, HEADING)).toBe('AbcDOCX-EDITOR.DEV');

      // Arrow keys move the caret; Delete removes the character after it.
      await page.keyboard.press('ArrowLeft');
      await page.keyboard.press('Delete');
      await expect.poll(() => paragraphText(page, HEADING)).toBe('AbDOCX-EDITOR.DEV');
      await page.keyboard.press('ArrowRight');
      await page.keyboard.type('!');
      await expect.poll(() => paragraphText(page, HEADING)).toBe('AbD!OCX-EDITOR.DEV');

      // Keyboard selection: Shift+Home selects back to the paragraph start.
      await page.keyboard.press('Shift+Home');
      await page.keyboard.type('Z');
      await expect.poll(() => paragraphText(page, HEADING)).toBe('ZOCX-EDITOR.DEV');

      // Enter splits the paragraph: the text after the caret moves down one line.
      const headingTop = () =>
        page.locator(`[data-paragraph-id="${HEADING}"][data-start]`).first().boundingBox();
      const topBefore = (await headingTop())!.y;
      await page.keyboard.press('Home');
      await page.keyboard.press('Enter');
      await expect
        .poll(async () =>
          page.evaluate(
            () => document.querySelectorAll('[data-paragraph-id][data-start="0"]').length
          )
        )
        .toBeGreaterThan(0);
      await expect
        .poll(async () => (await page.getByText('ZOCX-EDITOR.DEV').boundingBox())!.y)
        .toBeGreaterThan(topBefore);

      // Undo walks every edit back to the loaded text; redo applies the last undone edit again.
      for (let step = 0; step < 20; step += 1) {
        if ((await paragraphText(page, HEADING)) === 'DOCX-EDITOR.DEV') break;
        await page.keyboard.press(`${modifier}+z`);
      }
      await expect.poll(() => paragraphText(page, HEADING)).toBe('DOCX-EDITOR.DEV');
      await page.keyboard.press(`${modifier}+y`);
      await expect.poll(() => paragraphText(page, HEADING)).not.toBe('DOCX-EDITOR.DEV');
      await page.keyboard.press(`${modifier}+z`);
      await expect.poll(() => paragraphText(page, HEADING)).toBe('DOCX-EDITOR.DEV');

      // Mouse selection: a double click selects a word, and typing replaces it.
      await page
        .locator(`[data-paragraph-id="${SECOND}"][data-start="0"]`)
        .first()
        .dblclick({
          position: { x: 20, y: 10 },
        });
      await page.keyboard.type('Q');
      await expect.poll(() => paragraphText(page, SECOND)).toMatch(/^Q ?TEST DOCUMENT$/);
    });

    test('view mode refuses edits until edit mode returns', async ({ page }) => {
      await open(page, SAMPLE);
      await page.getByTestId('packaged-mode').selectOption('view');
      await caretAtStart(page, HEADING);
      const before = await changeCount(page);
      await page.keyboard.type('X');
      await page.keyboard.press('Enter');
      await page.waitForTimeout(300);
      expect(await paragraphText(page, HEADING)).toBe('DOCX-EDITOR.DEV');
      expect(await changeCount(page)).toBe(before);

      await page.getByTestId('packaged-mode').selectOption('edit');
      await caretAtStart(page, HEADING);
      await page.keyboard.type('X');
      await expect.poll(() => paragraphText(page, HEADING)).toBe('XDOCX-EDITOR.DEV');
    });

    test('a new file loads into the same component', async ({ page }) => {
      await open(page, SAMPLE);
      await page.evaluate(() => {
        const host = document.querySelector('docx-editor') as HTMLElement & { marker?: number };
        host.marker = 42;
        (window as unknown as { scroller: Element }).scroller = host.querySelector(
          '.docx-editor__scroll-container'
        )!;
      });
      await open(page, OTHER);
      expect(await paragraphText(page, HEADING)).not.toBe('DOCX-EDITOR.DEV');
      const same = await page.evaluate(() => {
        const host = document.querySelector('docx-editor') as HTMLElement & { marker?: number };
        const scroller = (window as unknown as { scroller: Element }).scroller;
        return {
          marker: host.marker,
          scrollerKept: scroller.isConnected && host.contains(scroller),
          hosts: document.querySelectorAll('docx-editor').length,
          surfaces: document.querySelectorAll('.docx-pages').length,
        };
      });
      expect(same).toEqual({ marker: 42, scrollerKept: true, hosts: 1, surfaces: 1 });
    });

    test('destroying the component releases its listeners and DOM', async ({ page }) => {
      const cdp = await page.context().newCDPSession(page);
      const listeners = async () => {
        const count = async (expression: string) => {
          const { result } = await cdp.send('Runtime.evaluate', { expression });
          const { listeners } = await cdp.send('DOMDebugger.getEventListeners', {
            objectId: result.objectId!,
          });
          return listeners.length;
        };
        return { window: await count('window'), document: await count('document') };
      };
      const before = await listeners();

      await open(page, SAMPLE);
      await caretAtStart(page, HEADING);
      await page.keyboard.type('X');
      expect((await listeners()).document).toBeGreaterThan(before.document);

      await page.getByRole('button', { name: 'Fechar' }).click();
      await expect(page.locator('docx-editor')).toHaveCount(0);
      await expect(page.locator('.docx-page')).toHaveCount(0);
      expect(await listeners()).toEqual(before);
    });

    test('save() returns DOCX bytes with the edits, and they reopen', async ({ page }) => {
      await open(page, SAMPLE);
      await caretAtStart(page, HEADING);
      await page.keyboard.type('SALVO123 ');
      await expect.poll(() => paragraphText(page, HEADING)).toBe('SALVO123 DOCX-EDITOR.DEV');

      const download = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Salvar' }).click();
      const file = await download;
      expect(file.suggestedFilename()).toBe('sample.docx');
      const path = await file.path();
      const parts = unzipSync(new Uint8Array(readFileSync(path)));
      expect(strFromU8(parts['word/document.xml']!)).toContain('SALVO123');

      await open(page, path);
      expect(await paragraphText(page, HEADING)).toBe('SALVO123 DOCX-EDITOR.DEV');
    });

    test('a zoom change repaints the pages without a reload', async ({ page }) => {
      await open(page, SAMPLE);
      await caretAtStart(page, HEADING);
      await page.keyboard.type('Z ');
      await expect.poll(() => paragraphText(page, HEADING)).toBe('Z DOCX-EDITOR.DEV');
      const pageWidth = () =>
        page.evaluate(() => document.querySelector('.docx-page')!.getBoundingClientRect().width);
      const before = await pageWidth();
      await page.evaluate(() => {
        (window as unknown as { surface: Element }).surface =
          document.querySelector('.docx-paginated-surface')!;
      });

      await page.getByRole('button', { name: 'Aumentar zoom' }).click();
      await page.getByRole('button', { name: 'Aumentar zoom' }).click();
      await expect(page.getByTestId('packaged-zoom')).toHaveText('150%');
      await expect.poll(pageWidth).toBeCloseTo(before * 1.5, 0);

      // The same surface and the edit stay: the engine was not created again.
      expect(
        await page.evaluate(() => (window as unknown as { surface: Element }).surface.isConnected)
      ).toBe(true);
      expect(await paragraphText(page, HEADING)).toBe('Z DOCX-EDITOR.DEV');
    });

    test('suggesting mode saves typing as a tracked insertion by the author', async ({ page }) => {
      await open(page, SAMPLE);
      await page.getByLabel('Autor').fill('Ada');
      await page.getByLabel('Autor').press('Tab');
      await page.getByTestId('packaged-mode').selectOption('suggesting');
      await caretAtStart(page, HEADING);
      await page.keyboard.type('NOVO ');
      await expect.poll(() => paragraphText(page, HEADING)).toContain('NOVO');

      const download = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Salvar' }).click();
      const parts = unzipSync(new Uint8Array(readFileSync(await (await download).path())));
      const xml = strFromU8(parts['word/document.xml']!);
      expect(xml).toMatch(/<w:ins [^>]*w:author="Ada"/);
      expect(xml).toContain('NOVO');
    });
  });
}
