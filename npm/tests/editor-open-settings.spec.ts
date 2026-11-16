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

/**
 * The session a DocxEditor opens must not capture an initial projection. That capture (on by
 * default for sessions) projects the whole document to markdown and copies its bytes at open, for
 * GetDiff, semantic changes and delivery evidence — nothing the editor calls — and it was a large
 * share of the open's single synchronous main-thread call.
 */
test.describe('DocxEditor: session settings at open', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  for (const how of ['open', 'openAsync'] as const) {
    test(`${how}() opens without an initial projection and still edits, undoes and saves`, async ({ page }) => {
      const res = await page.evaluate(async ({ bytes, how }) => {
        const D = (window as any).Docxodus;
        const real = D.getWasmExports ? D.getWasmExports() : D;
        const settings: Record<string, unknown>[] = [];
        const bridge = {
          ...real.DocxSessionBridge,
          OpenSession: (b: Uint8Array, json: string) => {
            settings.push(JSON.parse(json));
            return real.DocxSessionBridge.OpenSession(b, json);
          },
        };
        const exports = { ...real, DocxSessionBridge: bridge };
        const container = document.createElement('div');
        document.body.appendChild(container);
        const editor = how === 'open'
          ? D.DocxEditor.open(container, new Uint8Array(bytes), exports, {})
          : await D.DocxEditor.openAsync(container, new Uint8Array(bytes), exports, {});
        const block = Array.from(container.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'))
          .find((el) => (el.textContent ?? '').trim().length > 3)!;
        block.focus();
        block.textContent = 'settings probe';
        block.dispatchEvent(new Event('blur'));
        const edited = editor.version;
        const savedEdited = new TextDecoder().decode(editor.save()).length > 0;
        editor.undo();
        const undone = editor.version;
        editor.close();
        container.remove();
        return { settings, edited, undone, savedEdited };
      }, { bytes: fixture, how });

      expect(res.settings).toHaveLength(1);
      expect(res.settings[0].captureInitialProjection).toBe(false);
      expect(res.edited).toBe(1);
      expect(res.undone).toBe(2);
      expect(res.savedEdited).toBe(true);
    });
  }
});
