import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const fixture = Array.from(fs.readFileSync(
  path.join(__dirname, '../../TestFiles/HC001-5DayTourPlanTemplate.docx'),
));
const CDN_ORIGIN = 'http://localhost:8083';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

test.describe('DocxEditor.commitPendingInput and grep', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('commits focused typing in place, keeping the node, focus and caret; grep sees it', async ({ page }) => {
    await page.evaluate((bytes) => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.id = 'ed';
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, new Uint8Array(bytes), D, {});
      (window as any).__editor = editor;
      const block = Array.from(container.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'))
        .find((el) => (el.textContent ?? '').trim().length > 3)!;
      (window as any).__block = block;
      block.focus();
      const range = document.createRange();
      range.selectNodeContents(block);
      range.collapse(false);
      getSelection()!.removeAllRanges();
      getSelection()!.addRange(range);
    }, fixture);
    await page.keyboard.type(' QZXW');
    const res = await page.evaluate(() => {
      const editor = (window as any).__editor;
      const block = (window as any).__block as HTMLElement;
      const before = editor.version;
      const before2 = editor.grep('QZXW').length;
      const ok = editor.commitPendingInput();
      const sel = getSelection()!;
      return {
        before, ok, after: editor.version, grepBefore: before2,
        sameNode: block.isConnected, focused: document.activeElement === block,
        caretInBlock: !!sel.anchorNode && block.contains(sel.anchorNode),
        grep: editor.grep('QZXW').length,
      };
    });
    // grep itself commits first, so even the first call saw the typing.
    expect(res.grepBefore).toBe(1);
    expect(res.ok).toBe(true);
    expect(res.after).toBe(res.before + 1);
    expect(res.sameNode).toBe(true);
    expect(res.focused).toBe(true);
    expect(res.caretInBlock).toBe(true);
    expect(res.grep).toBe(1);
  });
});

test('prewarmEditor resolves, and a later createEditor still opens and edits', async ({ page }) => {
  await page.goto('/');
  const res = await page.evaluate(async ({ cdn, bytes }) => {
    const mod = await import(/* @vite-ignore */ `${cdn}/embed.bundle.js`);
    await mod.prewarmEditor();
    await mod.prewarmEditor();
    const div = document.createElement('div');
    document.body.appendChild(div);
    const editor = await mod.createEditor(div, new Uint8Array(bytes), { windowed: true });
    const block = div.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]')!;
    block.focus();
    block.textContent = 'AFTER_PREWARM';
    block.dispatchEvent(new Event('blur'));
    const version = editor.version;
    editor.close();
    return { version, leftovers: document.querySelectorAll('[data-docxodus-embed-root]').length };
  }, { cdn: CDN_ORIGIN, bytes: fixture });
  expect(res.version).toBe(1);
  expect(res.leftovers).toBe(1);
});
