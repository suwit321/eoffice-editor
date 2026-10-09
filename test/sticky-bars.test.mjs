import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { spawn } from 'node:child_process';
import http from 'node:http';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

describe('Sticky Toolbar & Menubar (การตรึงและปรับให้แถบคำสั่งและแถบเครื่องมือติดกัน)', () => {
  const html = readFileSync(join(root, 'eoffice-editor.html'), 'utf8');
  const distSingle = readFileSync(join(root, 'dist/eoffice-editor-single.html'), 'utf8');

  test('CSS มีการกำหนดตัวแปร --topbar-h และตรึง .topbar ไว้ด้านบนสุด (sticky top: 0)', () => {
    assert.ok(html.includes('--topbar-h:36px'), 'ต้องมีตัวแปร --topbar-h:36px ใน :root');
    assert.match(html, /\.topbar\{[^}]*position:sticky;top:0/, '.topbar ต้องมี position:sticky;top:0');
    assert.match(html, /\.topbar\{[^}]*height:var\(--topbar-h\)/, '.topbar ต้องใช้ความสูงตามตัวแปร --topbar-h');
  });

  test('CSS ปรับ .desk และ .sheet ให้เชื่อมต่อกับ .topbar โดยไม่มีช่องว่าง (เพิ่มพื้นที่ทำงาน)', () => {
    assert.match(html, /\.desk\{[^}]*padding:0 16px 24px/, '.desk ต้องมี padding-top เป็น 0');
    assert.match(html, /\.desk\{[^}]*overflow:visible/, '.desk ต้องมี overflow:visible เพื่อให้ sticky ทำงานได้สมบูรณ์');
    assert.match(html, /\.sheet\{[^}]*border-top:0/, '.sheet ต้องไม่มีเส้นขอบบนซ้ำซ้อนกับ .topbar');
  });

  test('CSS ตรึง .head (toolbar และ ruler) ต่อเนื่องใต้ .topbar (sticky top: var(--topbar-h))', () => {
    assert.match(html, /\.head\{[^}]*position:sticky;top:var\(--topbar-h\)/, '.head ต้อง sticky ต่อใต้ .topbar พอดี');
    assert.match(html, /\.head\{[^}]*z-index:20/, '.head ต้องมี z-index สูงพอเพื่อบังเนื้อหาที่เลื่อนขึ้นมา');
  });

  test('ไฟล์ dist/eoffice-editor-single.html ได้รับการอัปเดตสไตล์การตรึงแถบเช่นกัน', () => {
    assert.ok(distSingle.includes('--topbar-h:36px'));
    assert.match(distSingle, /\.topbar\{[^}]*position:sticky;top:0/);
    assert.match(distSingle, /\.head\{[^}]*position:sticky;top:var\(--topbar-h\)/);
  });

  test('ย้ายปุ่ม "เพิ่มเติม ▾" และ "🎨" ไปไว้ในแถบคำสั่งบนสุด (.topbar .actions) เพื่อลดการตัดแถวในแถบเครื่องมือ', () => {
    const topbarActions = html.match(/<div class="actions">([\s\S]*?)<\/div>\s*<\/header>/)?.[1] || '';
    assert.ok(topbarActions.includes('id="btnMore"'), 'ต้องมี #btnMore ใน .topbar .actions');
    assert.ok(topbarActions.includes('id="morePop"'), 'ต้องมี #morePop ใน .topbar .actions');
    assert.ok(topbarActions.includes('id="btnTbColor"'), 'ต้องมี #btnTbColor ใน .topbar .actions');
    assert.ok(topbarActions.includes('id="tbcPop"'), 'ต้องมี #tbcPop ใน .topbar .actions');

    const toolbar = html.match(/<div class="toolbar"[\s\S]*?<\/div>\s*<\/div>\s*<div class="rulerwrap">/)?.[0] || '';
    assert.ok(!toolbar.includes('id="btnMore"'), 'ต้องไม่มี #btnMore ใน .toolbar');
    assert.ok(!toolbar.includes('id="btnTbColor"'), 'ต้องไม่มี #btnTbColor ใน .toolbar');
    assert.ok(!toolbar.includes('tb-more'), 'ต้องไม่มีคลาส .tb-more ใน .toolbar');

    assert.ok(!distSingle.includes('tb-more'), 'dist/eoffice-editor-single.html ต้องไม่มี tb-more');
  });

  test('ทดสอบการ Scroll ในเบราว์เซอร์: ทั้ง 2 แถบถูกตรึงไว้ด้านบนและติดกันพอดี (gap = 0)', async (t) => {
    const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
    if (!existsSync(edgePath)) {
      t.skip('ไม่พบ msedge.exe ข้ามการทดสอบเบราว์เซอร์จริง');
      return;
    }

    const port = 9240 + Math.floor(Math.random() * 50);
    const edge = spawn(edgePath, [
      '--headless=new',
      '--window-size=1280,900',
      `--remote-debugging-port=${port}`,
      `file:///${join(root, 'eoffice-editor.html').replace(/\\/g, '/')}`
    ]);

    try {
      let tabs = [];
      const t0 = Date.now();
      while (Date.now() - t0 < 5000) {
        try {
          tabs = await new Promise((resolve, reject) => {
            const req = http.get(`http://127.0.0.1:${port}/json`, res => {
              let d = ''; res.on('data', c => d += c);
              res.on('end', () => {
                try { resolve(JSON.parse(d)); } catch { resolve([]); }
              });
            });
            req.on('error', reject);
            req.setTimeout(800, () => req.destroy());
          });
          if (tabs.some(tab => tab.type === 'page' && tab.url.includes('eoffice-editor.html'))) break;
        } catch {}
        await new Promise(r => setTimeout(r, 100));
      }

      const page = tabs.find(tab => tab.type === 'page' && tab.url.includes('eoffice-editor.html'));
      assert.ok(page, 'ต้องพบหน้า eoffice-editor ใน DevTools');

      const ws = new WebSocket(page.webSocketDebuggerUrl);
      await new Promise(r => ws.onopen = r);

      let id = 1;
      const send = (method, params = {}) => new Promise(resolve => {
        const curId = id++;
        const handler = msg => {
          const res = JSON.parse(msg.data);
          if (res.id === curId) {
            ws.removeEventListener('message', handler);
            resolve(res.result);
          }
        };
        ws.addEventListener('message', handler);
        ws.send(JSON.stringify({ id: curId, method, params }));
      });

      // จำลองการใส่ข้อความตัวอย่างแล้วเลื่อนหน้าจอลง 300px
      await send('Runtime.evaluate', { expression: `document.querySelector('#btnSample').click()` });
      await new Promise(r => setTimeout(r, 300));
      await send('Runtime.evaluate', { expression: 'window.scrollTo(0, 200)' });

      const evalJs = `(() => {
        const tb = document.querySelector('.topbar').getBoundingClientRect();
        const hd = document.querySelector('.head').getBoundingClientRect();
        return {
          scrollY: window.scrollY,
          topbarTop: tb.top,
          topbarBottom: tb.bottom,
          headTop: hd.top,
          gap: hd.top - tb.bottom
        };
      })()`;

      const res = await send('Runtime.evaluate', { expression: evalJs, returnByValue: true });
      const val = res.result.value;

      assert.equal(val.scrollY, 200, 'ต้องเลื่อนหน้าจอลง 200px');
      assert.equal(val.topbarTop, 0, 'แถบบนสุด (.topbar) ต้องตรึงอยู่ที่ top: 0');
      assert.equal(val.headTop, val.topbarBottom, 'แถบเครื่องมือ (.head) ต้องอยู่ติดกับขอบล่างของแถบบนสุดพอดี (ไม่มีช่องว่าง)');
      assert.equal(val.gap, 0, 'ระยะห่างระหว่างทั้งสองแถบต้องเป็น 0');

      // ตรวจสอบการเปิด-ปิดเมนู เพิ่มเติม ▾ และ 🎨 บน .topbar รวมถึงความสูงแถบเครื่องมือ (.toolbar) ให้เหลือแถวเดียว
      const testPopups = await send('Runtime.evaluate', {
        expression: `(() => {
          const btnMore = document.querySelector('.topbar #btnMore');
          const morePop = document.querySelector('.topbar #morePop');
          const btnTbColor = document.querySelector('.topbar #btnTbColor');
          const tbcPop = document.querySelector('.topbar #tbcPop');
          const toolbar = document.querySelector('.toolbar');
          
          const initialMoreHidden = morePop.hidden;
          btnMore.click();
          const openMoreHidden = morePop.hidden;
          
          const initialTbcHidden = tbcPop.hidden;
          btnTbColor.click();
          const openTbcHidden = tbcPop.hidden;
          
          document.body.click();
          
          return {
            toolbarHeight: toolbar.getBoundingClientRect().height,
            initialMoreHidden,
            openMoreHidden,
            initialTbcHidden,
            openTbcHidden,
            afterCloseMore: morePop.hidden,
            afterCloseTbc: tbcPop.hidden
          };
        })()`,
        returnByValue: true
      });
      const popVal = testPopups.result.value;
      assert.equal(popVal.initialMoreHidden, true);
      assert.equal(popVal.openMoreHidden, false);
      assert.equal(popVal.initialTbcHidden, true);
      assert.equal(popVal.openTbcHidden, false);
      assert.equal(popVal.afterCloseMore, true);
      assert.equal(popVal.afterCloseTbc, true);
      assert.ok(popVal.toolbarHeight <= 35, `ความสูงแถบเครื่องมือ (${popVal.toolbarHeight}px) ต้องไม่เกิน 35px (แถวเดียว)`);

      ws.close();
    } finally {
      edge.kill();
    }
  });
});
