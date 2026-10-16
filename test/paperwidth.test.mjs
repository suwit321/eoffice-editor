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

describe('Paper Width & Edge Dragging (การขยายหน้ากระดาษด้วยการลากขอบ)', () => {
  const html = readFileSync(join(root, 'eoffice-editor.html'), 'utf8');
  const tools = readFileSync(join(root, 'eoffice-tools.js'), 'utf8');
  const distSingle = readFileSync(join(root, 'dist/eoffice-editor-single.html'), 'utf8');

  test('มี Handle สำหรับลากขยายขอบกระดาษทั้งบนไม้บรรทัดและขอบกระดาษ (ทั้งด้านซ้ายและด้านขวา)', () => {
    assert.ok(html.includes('id="rulerEdgeHandle"'), 'ต้องมี #rulerEdgeHandle บนไม้บรรทัด');
    assert.ok(html.includes('id="paperEdgeHandle"'), 'ต้องมี #paperEdgeHandle ที่ขอบกระดาษด้านขวา');
    assert.ok(html.includes('id="rulerEdgeHandleLeft"'), 'ต้องมี #rulerEdgeHandleLeft ที่ขอบซ้ายไม้บรรทัด');
    assert.ok(html.includes('id="paperEdgeHandleLeft"'), 'ต้องมี #paperEdgeHandleLeft ที่ขอบกระดาษด้านซ้าย');
    assert.ok(html.includes('.ruler-edge-handle'), 'ต้องมี class .ruler-edge-handle ใน CSS');
    assert.ok(html.includes('.paper-edge-handle'), 'ต้องมี class .paper-edge-handle ใน CSS');
    assert.ok(html.includes('.ruler-edge-handle-left'), 'ต้องมี class .ruler-edge-handle-left ใน CSS');
    assert.ok(html.includes('.paper-edge-handle-left'), 'ต้องมี class .paper-edge-handle-left ใน CSS');
  });

  test('มีปุ่มและสไตล์สำหรับโหมด "ขยายเต็มหน้าจอ (ซ่อนขอบสีเทาซ้าย-ขวา)"', () => {
    assert.ok(html.includes('id="btnFullWidth"'), 'ต้องมีปุ่ม #btnFullWidth ใน topbar');
    assert.ok(html.includes('body.full-width'), 'ต้องมีสไตล์ body.full-width ใน CSS');
    assert.ok(tools.includes('id="vwFullWidth"'), 'ต้องมี checkbox #vwFullWidth ในแผงเครื่องมือ');
    assert.ok(tools.includes('id="vwMarginPreset"'), 'ต้องมีตัวเลือกระยะขอบ #vwMarginPreset ในแผงเครื่องมือ');
  });

  test('CSS รองรับตัวแปร --text-w และ --paper-w แบบ dynamic', () => {
    assert.ok(html.includes('--text-w:16cm'), 'ต้องมีตัวแปรเริ่มต้น --text-w:16cm');
    assert.ok(html.includes('width:var(--text-w)'), 'สเกลไม้บรรทัดต้องผูกกับ var(--text-w)');
    assert.ok(html.includes('--paper-w:calc(var(--text-w)'), 'ความกว้างกระดาษต้องผูกกับ var(--text-w)');
  });

  test('มีฟังก์ชัน setTextWidth, resetPaperWidth และ Event eoffice:paperwidth', () => {
    assert.ok(html.includes('function setTextWidth'), 'ต้องมีฟังก์ชัน setTextWidth');
    assert.ok(html.includes('function resetPaperWidth'), 'ต้องมีฟังก์ชัน resetPaperWidth');
    assert.ok(html.includes('eoffice:paperwidth'), 'ต้องยิง custom event eoffice:paperwidth');
    assert.ok(html.includes('textWidth: () => rulerCm'), 'eofficeEditor ต้อง export textWidth');
    assert.ok(html.includes('setTextWidth: (w) => setTextWidth'), 'eofficeEditor ต้อง export setTextWidth');
  });

  test('แผงเครื่องมือเสริม (eoffice-tools.js) มีตัวเลือกความกว้างหน้ากระดาษและปุ่มคืนค่า', () => {
    assert.ok(tools.includes('id="vwPaperWidth"'), 'ต้องมี dropdown #vwPaperWidth ในแท็บมุมมอง');
    assert.ok(tools.includes('id="btnResetWidth"'), 'ต้องมีปุ่ม #btnResetWidth คืนค่า 16 ซม.');
    assert.ok(tools.includes('eoffice:paperwidth'), 'eoffice-tools ต้องดักฟัง event eoffice:paperwidth');
  });

  test('ไฟล์ผลลัพธ์ dist/eoffice-editor-single.html มีฟังก์ชันขยายหน้ากระดาษฝังพร้อมใช้งาน', () => {
    assert.ok(distSingle.includes('id="rulerEdgeHandle"'), 'dist/eoffice-editor-single.html ต้องมี #rulerEdgeHandle');
    assert.ok(distSingle.includes('id="paperEdgeHandle"'), 'dist/eoffice-editor-single.html ต้องมี #paperEdgeHandle');
    assert.ok(distSingle.includes('id="vwPaperWidth"'), 'dist/eoffice-editor-single.html ต้องมี #vwPaperWidth');
  });

  test('updatePaperWidthCss อัปเดตตัวแปร --text-w และ --paper-w บน document.body เพื่อให้พื้นที่ทำงานขยับตาม', () => {
    assert.ok(html.includes("document.body.style.setProperty('--text-w'"), 'updatePaperWidthCss ต้องอัปเดต --text-w บน body');
    assert.ok(html.includes("document.body.style.setProperty('--paper-w'"), 'updatePaperWidthCss ต้องอัปเดต --paper-w บน body');
    assert.ok(distSingle.includes("document.body.style.setProperty('--text-w'"), 'distSingle ต้องอัปเดต --text-w บน body เช่นกัน');
  });

  test('การลากปุ่มกั้นขวา (mk-right) ออกไปทางขวาเกิน rulerCm สามารถขยายความกว้างหน้ากระดาษได้', () => {
    assert.ok(html.includes("drag.kind === 'right' && rawX > rulerCm"), 'ต้องรองรับการลาก mk-right ขยายความกว้างหน้ากระดาษ');
    assert.ok(distSingle.includes("drag.kind === 'right' && rawX > rulerCm"), 'distSingle ต้องรองรับการลาก mk-right ขยายความกว้างหน้ากระดาษ');
  });

  test('ทดสอบในเบราว์เซอร์จริง (Edge Headless): พื้นที่ทำงานขยายตามกั้นขวาที่ 17 ซม. อย่างสมบูรณ์', async (t) => {
    const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
    if (!existsSync(edgePath)) {
      t.skip('ไม่พบ msedge.exe ข้ามการทดสอบเบราว์เซอร์จริง');
      return;
    }

    const port = 9260 + Math.floor(Math.random() * 30);
    const userData = join(tmpdir(), 'edge-paperw-' + Date.now() + '-' + Math.random().toString(36).slice(2));
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

      // ตรวจสอบขนาดเริ่มต้นที่ 16 ซม.
      const m16 = await send('Runtime.evaluate', {
        expression: `(() => {
          const box = document.querySelector('.editor-box');
          const marks = document.querySelector('#rulerMarks');
          const handle = document.querySelector('#rulerEdgeHandle');
          const mkRight = document.querySelector('.mk-right');
          const bodyStyle = getComputedStyle(document.body);
          return {
            textW: window.eofficeEditor.textWidth(),
            bodyTextW: bodyStyle.getPropertyValue('--text-w').trim(),
            boxWidth: box.getBoundingClientRect().width,
            marksWidth: marks.getBoundingClientRect().width,
            marksLeft: marks.getBoundingClientRect().left,
            handleLeft: handle.getBoundingClientRect().left,
            mkRightLeft: mkRight.getBoundingClientRect().left
          };
        })()`,
        returnByValue: true
      });

      assert.equal(m16.value.textW, 16, 'ความกว้างเริ่มต้นต้องเป็น 16 ซม.');

      // ขยายกั้นขวา / พื้นที่พิมพ์เป็น 17 ซม.
      const m17 = await send('Runtime.evaluate', {
        expression: `(() => {
          window.eofficeEditor.setTextWidth(17);
          const box = document.querySelector('.editor-box');
          const marks = document.querySelector('#rulerMarks');
          const handle = document.querySelector('#rulerEdgeHandle');
          const mkRight = document.querySelector('.mk-right');
          const bodyStyle = getComputedStyle(document.body);
          return {
            textW: window.eofficeEditor.textWidth(),
            bodyTextW: bodyStyle.getPropertyValue('--text-w').trim(),
            boxWidth: box.getBoundingClientRect().width,
            marksWidth: marks.getBoundingClientRect().width,
            marksLeft: marks.getBoundingClientRect().left,
            handleLeft: handle.getBoundingClientRect().left,
            mkRightLeft: mkRight.getBoundingClientRect().left
          };
        })()`,
        returnByValue: true
      });

      assert.equal(m17.value.textW, 17, 'textWidth ต้องเป็น 17 ซม.');
      assert.equal(m17.value.bodyTextW, '17cm', 'body --text-w ต้องเป็น 17cm');

      // ตรวจสอบว่า .editor-box ขยายตามกั้นขวาพอดี (+1 ซม. หรือประมาณ 37.8px)
      const diffBox = m17.value.boxWidth - m16.value.boxWidth;
      assert.ok(Math.abs(diffBox - 37.795) < 1.5, `editor-box ต้องขยายกว้างขึ้นประมาณ 37.8px (ได้ ${diffBox}px)`);

      // ตรวจสอบว่าไม้บรรทัด #rulerMarks ขยายตาม (+1 ซม.)
      const diffMarks = m17.value.marksWidth - m16.value.marksWidth;
      assert.ok(Math.abs(diffMarks - 37.795) < 1.5, `rulerMarks ต้องขยายกว้างขึ้นประมาณ 37.8px (ได้ ${diffMarks}px)`);

      // ตรวจสอบว่ามือจับขอบกระดาษขยับตามสเกลไม้บรรทัด (+1 ซม. จากจุด 0 ซม.)
      const distHandle16 = m16.value.handleLeft - m16.value.marksLeft;
      const distHandle17 = m17.value.handleLeft - m17.value.marksLeft;
      const diffHandle = distHandle17 - distHandle16;
      assert.ok(Math.abs(diffHandle - 37.795) < 1.5, `rulerEdgeHandle ต้องขยับไปทางขวาประมาณ 37.8px (ได้ ${diffHandle}px)`);

      // ตรวจสอบว่ากั้นขวา .mk-right อยู่ที่ 17 ซม. (+1 ซม. จากจุด 0 ซม.)
      const distRight16 = m16.value.mkRightLeft - m16.value.marksLeft;
      const distRight17 = m17.value.mkRightLeft - m17.value.marksLeft;
      const diffRight = distRight17 - distRight16;
      assert.ok(Math.abs(diffRight - 37.795) < 1.5, `mk-right ต้องขยับไปทางขวาประมาณ 37.8px (ได้ ${diffRight}px)`);

      // บันทึกภาพหน้าจอขณะขยายกั้นขวาเป็น 17 ซม.
      const ssRes = await send('Page.captureScreenshot', { format: 'png' });
      if (ssRes && ssRes.data) {
        const artifactDir = 'C:\\Users\\suwit\\.gemini\\antigravity\\brain\\3e316c91-ec18-4b0a-9a1e-3224837de510';
        writeFileSync(join(artifactDir, 'workspace-expanded-17cm.png'), Buffer.from(ssRes.data, 'base64'));
      }

      // 5. ทดสอบการมีอยู่ของมือจับขอบซ้าย #paperEdgeHandleLeft และ #rulerEdgeHandleLeft
      const leftHandleRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const hl = document.querySelector('#paperEdgeHandleLeft');
          const hr = document.querySelector('#rulerEdgeHandleLeft');
          return {
            hasPaperLeft: !!hl,
            hasRulerLeft: !!hr
          };
        })()`,
        returnByValue: true
      });
      assert.strictEqual(leftHandleRes.value.hasPaperLeft, true, 'ต้องมี #paperEdgeHandleLeft');
      assert.strictEqual(leftHandleRes.value.hasRulerLeft, true, 'ต้องมี #rulerEdgeHandleLeft');

      // 6. ทดสอบโหมดขยายเต็มหน้าจอ (ซ่อนขอบสีเทาซ้าย-ขวา)
      await send('Runtime.evaluate', {
        expression: `window.eofficeEditor.setFullWidth(true)`
      });
      const fullWidthState = await send('Runtime.evaluate', {
        expression: `(() => {
          const bodyHasClass = document.body.classList.contains('full-width');
          const sheet = document.querySelector('.sheet');
          const desk = document.querySelector('.desk');
          const btn = document.querySelector('#btnFullWidth');
          const sheetW = sheet.getBoundingClientRect().width;
          const deskW = desk.getBoundingClientRect().width;
          return {
            bodyHasClass,
            btnPressed: btn?.getAttribute('aria-pressed'),
            diffWidth: Math.abs(sheetW - deskW)
          };
        })()`,
        returnByValue: true
      });
      assert.strictEqual(fullWidthState.value.bodyHasClass, true, 'body ต้องมี class full-width');
      assert.strictEqual(fullWidthState.value.btnPressed, 'true', '#btnFullWidth ต้องมี aria-pressed="true"');
      assert.ok(fullWidthState.value.diffWidth < 2, 'ในโหมดเต็มจอ .sheet ต้องกว้างเท่ากับ .desk (ไร้ขอบสีเทาซ้าย-ขวา)');

      // บันทึกภาพหน้าจอขณะเปิดโหมดขยายเต็มหน้าจอ ไร้ขอบสีเทา
      const ssFull = await send('Page.captureScreenshot', { format: 'png' });
      if (ssFull && ssFull.data) {
        const artifactDir = 'C:\\Users\\suwit\\.gemini\\antigravity\\brain\\3e316c91-ec18-4b0a-9a1e-3224837de510';
        writeFileSync(join(artifactDir, 'workspace-full-width-no-gray.png'), Buffer.from(ssFull.data, 'base64'));
      }

      // ปิดโหมดเต็มจอ
      await send('Runtime.evaluate', {
        expression: `window.eofficeEditor.setFullWidth(false)`
      });

      // 7. ทดสอบการปรับระยะขอบกระดาษ (Margin Preset)
      await send('Runtime.evaluate', {
        expression: `window.eofficeEditor.setMarginPreset('narrow')`
      });
      const marginState = await send('Runtime.evaluate', {
        expression: `(() => {
          const pl = document.body.style.getPropertyValue('--pl');
          const pr = document.body.style.getPropertyValue('--pr');
          return { pl, pr };
        })()`,
        returnByValue: true
      });
      assert.strictEqual(marginState.value.pl, '1.25cm', 'ขอบแคบต้องกำหนด --pl เป็น 1.25cm');
      assert.strictEqual(marginState.value.pr, '1.25cm', 'ขอบแคบต้องกำหนด --pr เป็น 1.25cm');

      ws.close();
    } finally {
      edge.kill('SIGTERM');
    }
  });
});
