import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import * as zlib from 'zlib';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const fixture = Array.from(fs.readFileSync(
  path.join(__dirname, '../../TestFiles/HC001-5DayTourPlanTemplate.docx'),
));

/** See editor-headerfooter.spec.ts — read one entry out of a DOCX without a zip dependency. */
function readZipEntry(zip: Buffer, name: string): string {
  const eocd = zip.lastIndexOf(Buffer.from('PK\x05\x06', 'latin1'));
  const count = zip.readUInt16LE(eocd + 10);
  let p = zip.readUInt32LE(eocd + 16);
  for (let i = 0; i < count; i++) {
    const nameLen = zip.readUInt16LE(p + 28);
    const extraLen = zip.readUInt16LE(p + 30);
    const commentLen = zip.readUInt16LE(p + 32);
    if (zip.toString('latin1', p + 46, p + 46 + nameLen) === name) {
      const method = zip.readUInt16LE(p + 10);
      const compSize = zip.readUInt32LE(p + 20);
      const localOff = zip.readUInt32LE(p + 42);
      const start = localOff + 30 + zip.readUInt16LE(localOff + 26) + zip.readUInt16LE(localOff + 28);
      const data = zip.subarray(start, start + compSize);
      return (method === 0 ? data : zlib.inflateRawSync(data)).toString('utf8');
    }
    p += 46 + nameLen + extraLen + commentLen;
  }
  throw new Error(`entry not found: ${name}`);
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

/** Open the fixture, put the caret at the end of the first text block, and return its anchor. */
async function openAndPlaceCaret(page: Page, paginated: boolean): Promise<string> {
  return page.evaluate(({ bytes, paginated }) => {
    const D = (window as any).Docxodus;
    const container = document.createElement('div');
    container.id = 'ed';
    document.body.appendChild(container);
    const editor = D.DocxEditor.open(container, new Uint8Array(bytes), D, { paginated });
    (window as any).__editor = editor;
    const block = Array.from(container.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'))
      .find((el) => (el.textContent ?? '').trim().length > 3)!;
    block.focus();
    const range = document.createRange();
    range.selectNodeContents(block);
    range.collapse(false);
    getSelection()!.removeAllRanges();
    getSelection()!.addRange(range);
    return block.getAttribute('data-anchor')!;
  }, { bytes: fixture, paginated });
}

const savedXml = async (page: Page) => readZipEntry(Buffer.from(await page.evaluate(
  () => Array.from((window as any).__editor.save() as Uint8Array))), 'word/document.xml');
const blockText = (page: Page, unid: string) => page.evaluate(
  (u) => document.querySelector(`#ed [data-anchor="${u}"]`)?.textContent ?? null, unid);

// Typing is committed to the session on blur. While the caret is still in the block, undo,
// redo and save used to act on a session that had not seen it.
test.describe('DocxEditor: typing not yet committed (caret still in the block)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('save() includes it', async ({ page }) => {
    await openAndPlaceCaret(page, false);
    await page.keyboard.type(' QZXW');
    expect(await savedXml(page)).toContain('QZXW');
  });

  for (const paginated of [false, true]) {
    test(`Ctrl+Z takes it back, and redo() restores it (${paginated ? 'paginated' : 'flow'})`, async ({ page }) => {
      const unid = await openAndPlaceCaret(page, paginated);
      const original = await blockText(page, unid);
      await page.keyboard.type(' QZXW');

      await page.keyboard.press('Control+z');
      expect(await blockText(page, unid)).toBe(original);
      expect(await savedXml(page)).not.toContain('QZXW');

      // The repaint after an undo leaves no block focused, so redo through the API, not a key.
      await page.evaluate(() => (window as any).__editor.redo());
      expect(await blockText(page, unid)).toContain('QZXW');
      expect(await savedXml(page)).toContain('QZXW');
    });
  }

  test('undo() takes back only that typing, not the edit before it', async ({ page }) => {
    const unid = await openAndPlaceCaret(page, false);
    await page.keyboard.type(' FIRSTEDIT');
    // A blur commits the first edit; refocus and type a second one.
    await page.evaluate((u) => {
      const block = document.querySelector<HTMLElement>(`#ed [data-anchor="${u}"]`)!;
      block.blur();
      const again = document.querySelector<HTMLElement>(`#ed [data-anchor="${u}"]`)!;
      again.focus();
      const range = document.createRange();
      range.selectNodeContents(again);
      range.collapse(false);
      getSelection()!.removeAllRanges();
      getSelection()!.addRange(range);
    }, unid);
    await page.keyboard.type(' SECONDEDIT');

    await page.evaluate(() => (window as any).__editor.undo());

    const xml = await savedXml(page);
    expect(xml).toContain('FIRSTEDIT');
    expect(xml).not.toContain('SECONDEDIT');
    expect(await page.locator('#ed').textContent()).not.toContain('SECONDEDIT');
  });
});
