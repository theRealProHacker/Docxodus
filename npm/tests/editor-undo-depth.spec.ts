import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const fixture = Array.from(fs.readFileSync(
  path.join(__dirname, '../../TestFiles/HC001-5DayTourPlanTemplate.docx'),
));

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

/** Commit `edits` text edits to one block, then undo until nothing is left; count the undos. */
async function editThenUndoAll(page: Page, options: Record<string, unknown>, edits: number) {
  return page.evaluate(({ bytes, options, edits }) => {
    const D = (window as any).Docxodus;
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = D.DocxEditor.open(container, new Uint8Array(bytes), D, options);
    const target = () => Array.from(container.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'))
      .find((el) => (el.textContent ?? '').trim().length > 3)!;
    const original = target().textContent;
    for (let i = 0; i < edits; i++) {
      const block = target();
      block.focus();
      block.textContent = `edit ${i}`;
      block.dispatchEvent(new Event('blur'));
    }
    let undone = 0;
    for (let v = editor.version; ; v = editor.version) {
      editor.undo();
      if (editor.version === v) break;
      undone++;
    }
    const restored = target().textContent === original;
    editor.close();
    container.remove();
    return { undone, restored };
  }, { bytes: fixture, options, edits });
}

test.describe('DocxEditor: undoDepth option', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('undoDepth raises how many edits undo can take back', async ({ page }) => {
    const res = await editThenUndoAll(page, { undoDepth: 30 }, 25);
    expect(res.undone).toBe(25);
    expect(res.restored).toBe(true);
  });

  test('without it, the session default still applies', async ({ page }) => {
    const res = await editThenUndoAll(page, {}, 25);
    expect(res.undone).toBe(20);
    expect(res.restored).toBe(false);
  });
});
