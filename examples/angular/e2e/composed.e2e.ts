// Browser tests for the composition API variant of the example app (`?composed`). Each case
// runs twice: without zone.js (the default) and with zone.js.
import { expect, test, type Page } from '@playwright/test';
import { unzipSync, strFromU8 } from 'fflate';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const SAMPLE = fileURLToPath(
  new URL('../../../packages/core/src/editor/__tests__/fixtures/sample.docx', import.meta.url)
);
const HEADING = '/word/document.xml#0.0.0';

function paragraphText(page: Page, id: string): Promise<string> {
  return page.evaluate(
    (paragraphId) =>
      [...document.querySelectorAll(`[data-paragraph-id="${paragraphId}"][data-start]`)]
        .map((node) => node.textContent ?? '')
        .join(''),
    id
  );
}

async function open(page: Page) {
  await page.setInputFiles('input[type=file]', SAMPLE);
  await page.locator('.docx-page').first().waitFor();
  await expect(page.getByTestId('status')).toContainText('pronto');
}

/** Select the whole heading with a triple click on its first text run. */
async function selectHeading(page: Page) {
  await page.locator(`[data-paragraph-id="${HEADING}"][data-start="0"]`).first().click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+End');
}

const slot = (page: Page, id: string) => page.locator(`app-format-bar [data-slot="${id}"]`);

for (const mode of [
  { name: 'without zone.js', url: '/?composed' },
  { name: 'with zone.js', url: '/?composed&zone' },
]) {
  test.describe(`composition API, ${mode.name}`, () => {
    const errors: string[] = [];

    test.beforeEach(async ({ page }) => {
      errors.length = 0;
      page.on('pageerror', (error) => errors.push(error.message));
      page.on('console', (message) => {
        if (message.type() === 'error') errors.push(message.text());
      });
      await page.goto(mode.url);
      await open(page);
    });

    test.afterEach(() => {
      expect(errors).toEqual([]);
    });

    test('the format bar follows the selection and runs engine commands', async ({ page }) => {
      await expect(page.getByTestId('composed-page')).toHaveText(/^Página 1 de \d+$/);
      const bold = slot(page, 'text.bold');
      await selectHeading(page);
      const before = await bold.getAttribute('aria-pressed');
      await bold.click();
      await expect(bold).toHaveAttribute('aria-pressed', before === 'true' ? 'false' : 'true');
      await slot(page, 'history.undo').click();
      await expect(bold).toHaveAttribute('aria-pressed', before ?? 'false');

      await selectHeading(page);
      await page.getByLabel(/tamanhos da fonte|font sizes/i).selectOption('36');
      await expect(page.getByLabel(/tamanhos da fonte|font sizes/i)).toHaveValue('36');
    });

    test('view mode disables the commands, and edit mode enables them', async ({ page }) => {
      await selectHeading(page);
      await page.getByTestId('composed-mode').selectOption('view');
      await expect(slot(page, 'text.bold')).toBeDisabled();
      await page.getByTestId('composed-mode').selectOption('edit');
      await selectHeading(page);
      await expect(slot(page, 'text.bold')).toBeEnabled();
    });

    test('theme and language apply to the same engine', async ({ page }) => {
      const viewport = page.locator('docx-editor-viewport');
      await page.getByTestId('composed-theme').selectOption('dark');
      await expect(viewport).toHaveClass(/\bdark\b/);
      await expect(page.locator('app-format-bar')).toHaveClass(/\bdark\b/);

      const bold = slot(page, 'text.bold');
      const portuguese = await bold.getAttribute('aria-label');
      await page.getByTestId('composed-language').selectOption('en');
      await expect(bold).not.toHaveAttribute('aria-label', portuguese ?? '');
      await page.getByTestId('composed-direction').selectOption('rtl');
      await expect(page.locator('app-format-bar')).toHaveAttribute('dir', 'rtl');
      await expect(page.locator('app-editor-status')).toHaveAttribute('dir', 'rtl');
    });

    test('zoom buttons change the zoom level', async ({ page }) => {
      await expect(page.locator('app-editor-status output')).toHaveText('100%');
      await page.locator('app-editor-status button').last().click();
      await expect(page.locator('app-editor-status output')).toHaveText('125%');
    });

    test('the pages move to a pop-out window and back with their edits', async ({ page }) => {
      await page.locator(`[data-paragraph-id="${HEADING}"][data-start="0"]`).first().click();
      await page.keyboard.press('Home');
      await page.keyboard.type('Antes ');
      await expect.poll(() => paragraphText(page, HEADING)).toBe('Antes DOCX-EDITOR.DEV');

      const popup = page.waitForEvent('popup');
      await page.getByLabel('Páginas em outra janela').check();
      const other = await popup;
      await other.locator('.docx-page').first().waitFor();
      await expect(page.locator('app-editor-status')).toContainText('outra janela');
      await expect(page.locator('.composed .docx-page')).toHaveCount(0);
      expect(await paragraphText(other, HEADING)).toBe('Antes DOCX-EDITOR.DEV');

      await page.getByLabel('Páginas em outra janela').uncheck();
      await page.locator('.composed .docx-page').first().waitFor();
      await expect(page.locator('app-editor-status')).toContainText('esta janela');
      await slot(page, 'history.undo').click();
      await expect.poll(() => paragraphText(page, HEADING)).toBe('DOCX-EDITOR.DEV');
    });

    test('Salvar writes the edits of the composed editor', async ({ page }) => {
      await page.locator(`[data-paragraph-id="${HEADING}"][data-start="0"]`).first().click();
      await page.keyboard.press('Home');
      await page.keyboard.type('COMPOSTO ');
      await expect.poll(() => paragraphText(page, HEADING)).toBe('COMPOSTO DOCX-EDITOR.DEV');
      const download = page.waitForEvent('download');
      await page.getByRole('button', { name: 'Salvar' }).click();
      const parts = unzipSync(new Uint8Array(readFileSync(await (await download).path())));
      expect(strFromU8(parts['word/document.xml']!)).toContain('COMPOSTO');
    });
  });
}
