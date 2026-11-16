import { test, expect } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
// 117 body units across four sections, bordered headings, tables, a footnote and an endnote.
const fixture = Array.from(fs.readFileSync(
  path.join(__dirname, '../../TestFiles/HC031-Complicated-Document.docx'),
));

// The embed bundle and its WASM come from the CORS origin, as in cdn-embed.spec.ts.
const CDN_ORIGIN = 'http://localhost:8083';

/**
 * `createEditor(..., { windowed: true })` mounts through `DocxEditor.openAsync`: the page keeps
 * running timers while the document lands, the result is the document `createEditor` renders
 * without the option, and the converter stylesheet is scoped like every other embed render.
 */
test.describe('createEditor windowed mount', () => {
  for (const paginated of [false, true]) {
    test(`yields while mounting and renders what the synchronous mount renders (${paginated ? 'paginated' : 'flow'})`, async ({ page }) => {
      await page.goto('/');
      const result = await page.evaluate(async ({ cdn, docBytes, paginated }) => {
        const hostStyle = document.createElement('style');
        hostStyle.textContent = 'body { margin: 7px; font-family: monospace; }';
        document.head.appendChild(hostStyle);
        const mod = await import(/* @vite-ignore */ `${cdn}/embed.bundle.js`);
        const mount = async (options: Record<string, unknown>) => {
          const div = document.createElement('div');
          document.body.appendChild(div);
          let ticks = 0;
          const ticker = setInterval(() => ticks++, 0);
          const editor = await mod.createEditor(div, new Uint8Array(docBytes), { paginated, ...options });
          clearInterval(ticker);
          const text = Array.from(div.querySelectorAll('[data-anchor]'), (el) => el.textContent).join('\n');
          return { div, editor, ticks, text };
        };
        const sync = await mount({});
        sync.editor.close();
        const progress: number[] = [];
        const windowed = await mount({ windowed: true, windowSize: 8, onProgress: (n: number) => progress.push(n) });

        const block = windowed.div.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]')!;
        block.focus();
        block.textContent = 'WINDOWED_EDIT';
        block.dispatchEvent(new Event('blur'));
        const saved = new TextDecoder('latin1').decode(windowed.editor.save());
        const styles = Array.from(windowed.div.querySelectorAll<HTMLStyleElement>('style'));
        const out = {
          sameText: windowed.text === sync.text && sync.text.length > 0,
          windowedTicks: windowed.ticks,
          progressCalls: progress.length,
          version: windowed.editor.version,
          savedNonEmpty: saved.length > 1000,
          hostMargin: getComputedStyle(document.body).margin,
          hostFont: getComputedStyle(document.body).fontFamily,
          styleCount: styles.length,
          stylesScoped: styles.every((style) => style.hasAttribute('data-docxodus-scoped-style')),
        };
        windowed.editor.close();
        return out;
      }, { cdn: CDN_ORIGIN, docBytes: fixture, paginated });

      expect(result.sameText).toBe(true);
      // 117 units in windows of 8: at least one timer turn per window.
      expect(result.progressCalls).toBeGreaterThanOrEqual(14);
      expect(result.windowedTicks).toBeGreaterThanOrEqual(result.progressCalls);
      expect(result.version).toBe(1);
      expect(result.savedNonEmpty).toBe(true);
      expect(result.hostMargin).toBe('7px');
      expect(result.hostFont).toContain('monospace');
      expect(result.styleCount).toBeGreaterThan(0);
      expect(result.stylesScoped).toBe(true);
    });
  }
});
