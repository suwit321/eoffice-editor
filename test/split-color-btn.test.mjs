import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { spawn } from 'node:child_process';
import { tmpdir } from 'node:os';
import http from 'node:http';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

describe('Word-style Split Button for Font Color & Highlight (การปรับแต่งปุ่มสีตัวอักษรและสีไฮไลต์แบบกดเปลี่ยนสีทันที)', () => {
  const html = readFileSync(join(root, 'eoffice-editor.html'), 'utf8');
  const distSingle = readFileSync(join(root, 'dist/eoffice-editor-single.html'), 'utf8');

  test('โครงสร้าง HTML มี Split Button (.tb-split) แยกปุ่มคลิกเปลี่ยนสีทันที และปุ่มลูกศรเลือกสี', () => {
    // ปุ่มสีตัวอักษร
    assert.ok(html.includes('id="splitColor"'), 'ต้องมีคอนเทนเนอร์ #splitColor');
    assert.ok(html.includes('class="tb tb-split-main" id="btnColor" data-cmd="color"'), 'ปุ่มหลัก #btnColor ต้องมี data-cmd="color" สำหรับคลิกเปลี่ยนสีทันที');
    assert.ok(html.includes('class="tb tb-split-arrow" id="btnColorArrow" data-menu="m-color"'), 'ปุ่มลูกศร #btnColorArrow ต้องมี data-menu="m-color" สำหรับเปิดพาเล็ตสี');

    // ปุ่มสีเน้นข้อความ (ไฮไลต์)
    assert.ok(html.includes('id="splitBg"'), 'ต้องมีคอนเทนเนอร์ #splitBg');
    assert.ok(html.includes('class="tb tb-split-main" id="btnBg" data-cmd="background"'), 'ปุ่มหลัก #btnBg ต้องมี data-cmd="background" สำหรับคลิกไฮไลต์ทันที');
    assert.ok(html.includes('class="tb tb-split-arrow" id="btnBgArrow" data-menu="m-bg"'), 'ปุ่มลูกศร #btnBgArrow ต้องมี data-menu="m-bg" สำหรับเปิดพาเล็ตสี');
  });

  test('CSS มีการกำหนดสไตล์สำหรับ .tb-split, .tb-split-main, .tb-split-arrow อย่างถูกต้อง', () => {
    assert.ok(html.includes('.tb-split{'), 'ต้องมีคลาส .tb-split');
    assert.ok(html.includes('.tb-split .tb-split-main{'), 'ต้องมีสไตล์ .tb-split .tb-split-main');
    assert.ok(html.includes('.tb-split .tb-split-arrow{'), 'ต้องมีสไตล์ .tb-split .tb-split-arrow');
    assert.ok(html.includes('.tb-split-arrow::after{'), 'ต้องมีสัญลักษณ์ลูกศรสามเหลี่ยม .tb-split-arrow::after');
  });

  test('ไฟล์ dist/eoffice-editor-single.html มีฟังก์ชันและโครงสร้าง Split Button พร้อมใช้งาน', () => {
    assert.ok(distSingle.includes('id="splitColor"'), 'distSingle ต้องมี #splitColor');
    assert.ok(distSingle.includes('id="btnColorArrow"'), 'distSingle ต้องมี #btnColorArrow');
    assert.ok(distSingle.includes('id="splitBg"'), 'distSingle ต้องมี #splitBg');
    assert.ok(distSingle.includes('id="btnBgArrow"'), 'distSingle ต้องมี #btnBgArrow');
  });

  test('ทดสอบในเบราว์เซอร์จริง (Edge Headless): คลิก #btnColor เปลี่ยนสีทันที 1 คลิก และสลับสีผ่านลูกศรพาเล็ต', async (t) => {
    const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
    if (!existsSync(edgePath)) {
      t.skip('ไม่พบ msedge.exe ข้ามการทดสอบเบราว์เซอร์จริง');
      return;
    }

    const port = 9290 + Math.floor(Math.random() * 30);
    const userData = join(tmpdir(), 'edge-color-' + Date.now() + '-' + Math.random().toString(36).slice(2));
    const edge = spawn(edgePath, [
      '--headless=new',
      '--window-size=1280,900',
      `--user-data-dir=${userData}`,
      `--remote-debugging-port=${port}`,
      `file:///${join(root, 'dist/eoffice-editor-single.html').replace(/\\/g, '/')}`
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
          if (tabs.some(tab => tab.type === 'page' && tab.url.includes('eoffice-editor-single.html'))) break;
        } catch {}
        await new Promise(r => setTimeout(r, 100));
      }

      const page = tabs.find(tab => tab.type === 'page' && tab.url.includes('eoffice-editor-single.html'));
      assert.ok(page, 'ต้องพบหน้า eoffice-editor ใน DevTools');

      const ws = new WebSocket(page.webSocketDebuggerUrl);
      await new Promise(r => ws.onopen = r);

      let id = 1;
      const send = (method, params = {}) => new Promise((resolve, reject) => {
        const curId = id++;
        const handler = msg => {
          const res = JSON.parse(msg.data);
          if (res.id === curId) {
            ws.removeEventListener('message', handler);
            if (res.error) reject(new Error(JSON.stringify(res.error)));
            else if (res.result?.exceptionDetails) reject(new Error(JSON.stringify(res.result.exceptionDetails)));
            else resolve(res.result?.result || res.result);
          }
        };
        ws.addEventListener('message', handler);
        ws.send(JSON.stringify({ id: curId, method, params }));
      });

      // รอให้ eofficeEditor โหลดและพร้อมทำงาน
      await send('Runtime.evaluate', {
        expression: `new Promise((resolve) => {
          if (window.eofficeEditor) return resolve();
          document.addEventListener('eoffice:ready', () => resolve(), { once: true });
          const t0 = Date.now();
          const iv = setInterval(() => {
            if (window.eofficeEditor || Date.now() - t0 > 5000) {
              clearInterval(iv);
              resolve();
            }
          }, 50);
        })`,
        awaitPromise: true
      });

      // 1. ตรวจสอบค่าสีเริ่มต้นของ swatch บน #btnColor (เริ่มต้นคือสีแดง #e03e2d)
      const initSwatch = await send('Runtime.evaluate', {
        expression: `(() => {
          const btn = document.querySelector('#btnColor');
          return btn.style.getPropertyValue('--swatch').trim();
        })()`,
        returnByValue: true
      });
      assert.strictEqual(initSwatch.value, '#e03e2d', 'แถบสีเริ่มต้นใต้ปุ่ม A ต้องเป็นสีแดง #e03e2d');

      // 2. ใส่ข้อความตัวอย่างและเลือกข้อความ "บันทึกข้อความ"
      await send('Runtime.evaluate', {
        expression: `(() => {
          window.eofficeEditor.quill.setText('บันทึกข้อความ ทดสอบระบบสีตัวอักษรทันที\\n');
          window.eofficeEditor.quill.setSelection(0, 12);
        })()`
      });

      // 3. คลิกที่ปุ่ม #btnColor โดยตรง (1 คลิกเปลี่ยนสีทันที)
      await send('Runtime.evaluate', {
        expression: `document.querySelector('#btnColor').click()`
      });

      // รอ microtask & schedule
      await new Promise(r => setTimeout(r, 200));

      // ตรวจสอบว่าข้อความที่เลือกเปลี่ยนเป็นสีแดงทันที และเมนู #m-color ยังคงปิดอยู่ (hidden)
      const directApplyRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const f = window.eofficeEditor.quill.getFormat(0, 12);
          const menu = document.querySelector('#m-color');
          const isMenuHidden = menu ? menu.hidden : true;
          return {
            color: f.color,
            isMenuHidden
          };
        })()`,
        returnByValue: true
      });
      assert.strictEqual(directApplyRes.value.color, '#e03e2d', 'ข้อความต้องเปลี่ยนเป็นสีแดง #e03e2d ทันที');
      assert.strictEqual(directApplyRes.value.isMenuHidden, true, 'เมนูพาเล็ตต้องไม่เด้งขึ้นมาเมื่อคลิกปุ่มหลัก');

      // 4. คลิก #btnColor ซ้ำอีกครั้ง -> ต้อง toggle คืนเป็นสีปกติ (false)
      await send('Runtime.evaluate', {
        expression: `document.querySelector('#btnColor').click()`
      });
      await new Promise(r => setTimeout(r, 200));
      const toggleRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const f = window.eofficeEditor.quill.getFormat(0, 12);
          return { color: f.color || 'none' };
        })()`,
        returnByValue: true
      });
      assert.strictEqual(toggleRes.value.color, 'none', 'คลิกซ้ำต้องยกเลิกสี (toggle กลับเป็นค่าปกติ)');

      // 5. คลิกที่ลูกศร #btnColorArrow -> เมนูพาเล็ตสีต้องเปิดออก
      await send('Runtime.evaluate', {
        expression: `document.querySelector('#btnColorArrow').click()`
      });
      await new Promise(r => setTimeout(r, 200));

      const menuOpenRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const menu = document.querySelector('#m-color');
          const arrow = document.querySelector('#btnColorArrow');
          return {
            isOpen: menu && !menu.hidden,
            ariaExpanded: arrow ? arrow.getAttribute('aria-expanded') : null
          };
        })()`,
        returnByValue: true
      });
      assert.strictEqual(menuOpenRes.value.isOpen, true, 'คลิกลูกศรต้องเปิดเมนู #m-color');
      assert.strictEqual(menuOpenRes.value.ariaExpanded, 'true', '#btnColorArrow ต้องมี aria-expanded="true"');

      // 6. คลิกเลือกสีน้ำเงิน (#0000ff) ในพาเล็ต
      await send('Runtime.evaluate', {
        expression: `(() => {
          const blueSw = document.querySelector('#m-color [data-val="#0000ff"]');
          if (blueSw) blueSw.click();
        })()`
      });
      await new Promise(r => setTimeout(r, 200));

      // ตรวจสอบว่าข้อความเปลี่ยนเป็นสีน้ำเงิน เมนูถูกปิด และแถบสีใต้ปุ่ม A อัปเดตเป็น #0000ff
      const bluePickRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const f = window.eofficeEditor.quill.getFormat(0, 12);
          const btn = document.querySelector('#btnColor');
          const menu = document.querySelector('#m-color');
          return {
            color: f.color,
            swatch: btn.style.getPropertyValue('--swatch').trim(),
            menuClosed: menu.hidden
          };
        })()`,
        returnByValue: true
      });
      assert.strictEqual(bluePickRes.value.color, '#0000ff', 'ข้อความต้องเปลี่ยนเป็นสีน้ำเงินตามที่เลือก');
      assert.strictEqual(bluePickRes.value.swatch, '#0000ff', 'แถบสีใต้ปุ่ม A ต้องอัปเดตเป็นสีน้ำเงิน #0000ff');
      assert.strictEqual(bluePickRes.value.menuClosed, true, 'เมนูต้องปิดลงหลังจากเลือกสี');

      // 7. เลือกข้อความช่วงที่สอง "ทดสอบระบบสี" (index 13 ถึง 25) แล้วคลิก #btnColor ทันที 1 คลิก
      await send('Runtime.evaluate', {
        expression: `(() => {
          window.eofficeEditor.quill.setSelection(13, 12);
          document.querySelector('#btnColor').click();
        })()`
      });
      await new Promise(r => setTimeout(r, 200));

      const secondApplyRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const f = window.eofficeEditor.quill.getFormat(13, 12);
          return { color: f.color };
        })()`,
        returnByValue: true
      });
      assert.strictEqual(secondApplyRes.value.color, '#0000ff', 'คลิก 1 ครั้งต้องใส่สีน้ำเงินที่จำไว้ทันที');

      // 8. ทดสอบปุ่มสีเน้นข้อความ (Highlight #btnBg) คลิกใส่สีทันที
      await send('Runtime.evaluate', {
        expression: `(() => {
          window.eofficeEditor.quill.setSelection(0, 12);
          document.querySelector('#btnBg').click();
        })()`
      });
      await new Promise(r => setTimeout(r, 200));

      const bgApplyRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const f = window.eofficeEditor.quill.getFormat(0, 12);
          return { bg: f.background };
        })()`,
        returnByValue: true
      });
      assert.strictEqual(bgApplyRes.value.bg, '#fff2cc', 'คลิก #btnBg 1 ครั้งต้องใส่สีไฮไลต์ทันที');

      // ถ่ายภาพหน้าจอเก็บเป็น Artifact
      const ssRes = await send('Page.captureScreenshot', { format: 'png' });
      if (ssRes && ssRes.data) {
        const artifactDir = 'C:\\Users\\suwit\\.gemini\\antigravity\\brain\\3e316c91-ec18-4b0a-9a1e-3224837de510';
        writeFileSync(join(artifactDir, 'split-color-btn.png'), Buffer.from(ssRes.data, 'base64'));
      }

      ws.close();
    } finally {
      edge.kill('SIGKILL');
    }
  });
});
