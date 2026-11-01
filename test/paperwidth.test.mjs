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

describe('Paper Width & Right Margin Dragging (การขยายหน้ากระดาษด้วยการลากกั้นขวา ไร้มือจับขอบ)', () => {
  const html = readFileSync(join(root, 'eoffice-editor.html'), 'utf8');
  const tools = readFileSync(join(root, 'eoffice-tools.js'), 'utf8');
  const distSingle = readFileSync(join(root, 'dist/eoffice-editor-single.html'), 'utf8');

  test('ไม่มี Handle สำหรับลากขยายขอบกระดาษทั้งบนไม้บรรทัดและขอบกระดาษ (ซ้ายและขวา)', () => {
    assert.ok(!html.includes('id="rulerEdgeHandle"'), 'ต้องไม่มี #rulerEdgeHandle บนไม้บรรทัด');
    assert.ok(!html.includes('id="paperEdgeHandle"'), 'ต้องไม่มี #paperEdgeHandle ที่ขอบกระดาษด้านขวา');
    assert.ok(!html.includes('id="rulerEdgeHandleLeft"'), 'ต้องไม่มี #rulerEdgeHandleLeft ที่ขอบซ้ายไม้บรรทัด');
    assert.ok(!html.includes('id="paperEdgeHandleLeft"'), 'ต้องไม่มี #paperEdgeHandleLeft ที่ขอบกระดาษด้านซ้าย');
    assert.ok(!html.includes('.ruler-edge-handle'), 'ต้องไม่มี class .ruler-edge-handle ใน CSS');
    assert.ok(!html.includes('.paper-edge-handle'), 'ต้องไม่มี class .paper-edge-handle ใน CSS');
    assert.ok(!html.includes('.ruler-edge-handle-left'), 'ต้องไม่มี class .ruler-edge-handle-left ใน CSS');
    assert.ok(!html.includes('.paper-edge-handle-left'), 'ต้องไม่มี class .paper-edge-handle-left ใน CSS');
  });

  test('ไม่มีปุ่มและสไตล์สำหรับโหมดขยายเต็มหน้าจอ (#btnFullWidth, body.full-width)', () => {
    assert.ok(!html.includes('id="btnFullWidth"'), 'ต้องไม่มีปุ่ม #btnFullWidth ใน topbar');
    assert.ok(!html.includes('body.full-width'), 'ต้องไม่มีสไตล์ body.full-width ใน CSS');
    assert.ok(!tools.includes('id="vwFullWidth"'), 'ต้องไม่มี checkbox #vwFullWidth ในแผงเครื่องมือ');
  });

  test('CSS กำหนดขนาดหน้ากระดาษเริ่มต้นเป็น A4 มาตรฐานราชการ (พื้นที่พิมพ์ 16 ซม., ขอบซ้าย 3 ซม., ขอบขวา 2 ซม.)', () => {
    assert.ok(html.includes('--text-w:16cm'), 'ต้องมีตัวแปรเริ่มต้น --text-w:16cm');
    assert.ok(html.includes('body.a4{--pl:3cm;--pr:2cm;--pt:2.5cm;--pb:2cm;--min-h:25.2cm}'), 'body.a4 ต้องมีระยะขอบมาตรฐาน 3/2/2.5/2 ซม.');
    assert.ok(html.includes('--paper-w:calc(var(--text-w) + var(--pl) + var(--pr) + 2px)'), 'ความกว้างกระดาษต้องผูกกับ var(--text-w) + ขอบซ้ายขวา');
  });

  test('การลากปุ่มกั้นขวา (mk-right) ออกไปทางขวาเกิน rulerCm สามารถขยายความกว้างหน้ากระดาษได้', () => {
    assert.ok(html.includes("drag.kind === 'right' && rawX > rulerCm"), 'ต้องมีโค้ดลาก mk-right ขยายกระดาษ');
    assert.ok(distSingle.includes("drag.kind === 'right' && rawX > rulerCm"), 'distSingle ต้องมีโค้ดลาก mk-right ขยายกระดาษ');
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

  test('ไฟล์ผลลัพธ์ dist/eoffice-editor-single.html ไม่มี Handle ลากขอบและไม่มีปุ่มเต็มจอ', () => {
    assert.ok(!distSingle.includes('id="rulerEdgeHandle"'), 'distSingle ต้องไม่มี #rulerEdgeHandle');
    assert.ok(!distSingle.includes('id="paperEdgeHandle"'), 'distSingle ต้องไม่มี #paperEdgeHandle');
    assert.ok(!distSingle.includes('id="rulerEdgeHandleLeft"'), 'distSingle ต้องไม่มี #rulerEdgeHandleLeft');
    assert.ok(!distSingle.includes('id="paperEdgeHandleLeft"'), 'distSingle ต้องไม่มี #paperEdgeHandleLeft');
    assert.ok(!distSingle.includes('id="btnFullWidth"'), 'distSingle ต้องไม่มี #btnFullWidth');
  });

  test('ทดสอบในเบราว์เซอร์จริง (Edge Headless): ลากกั้นขวาขยายหน้ากระดาษได้ และไม่มี Handle ลากขอบ', async (t) => {
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
      for (let i = 0; i < 40; i++) {
        try {
          const ready = await send('Runtime.evaluate', {
            expression: `Boolean(window.eofficeEditor && window.eofficeEditor.quill)`,
            returnByValue: true
          });
          if (ready?.value) break;
        } catch {}
        await new Promise(r => setTimeout(r, 100));
      }

      // ตรวจสอบขนาดมาตรฐานเริ่มต้นที่ 16 ซม. และไม่มี Edge Handles
      const m16 = await send('Runtime.evaluate', {
        expression: `(() => {
          const box = document.querySelector('.editor-box');
          const marks = document.querySelector('#rulerMarks');
          const handleRight = document.querySelector('#rulerEdgeHandle');
          const handleLeft = document.querySelector('#rulerEdgeHandleLeft');
          const paperHandleRight = document.querySelector('#paperEdgeHandle');
          const paperHandleLeft = document.querySelector('#paperEdgeHandleLeft');
          const btnFull = document.querySelector('#btnFullWidth');
          const mkRight = document.querySelector('.mk-right');
          const bodyStyle = getComputedStyle(document.body);
          return {
            textW: window.eofficeEditor.textWidth(),
            bodyTextW: bodyStyle.getPropertyValue('--text-w').trim(),
            boxWidth: box.getBoundingClientRect().width,
            marksWidth: marks.getBoundingClientRect().width,
            marksLeft: marks.getBoundingClientRect().left,
            hasHandleRight: !!handleRight,
            hasHandleLeft: !!handleLeft,
            hasPaperHandleRight: !!paperHandleRight,
            hasPaperHandleLeft: !!paperHandleLeft,
            hasBtnFull: !!btnFull,
            mkRightLeft: mkRight.getBoundingClientRect().left
          };
        })()`,
        returnByValue: true
      });

      assert.equal(m16.value.textW, 16, 'ความกว้างหน้ากระดาษเริ่มต้นต้องเป็น 16 ซม. A4');
      assert.equal(m16.value.bodyTextW, '16cm', 'body --text-w ต้องเป็น 16cm');
      assert.equal(m16.value.hasHandleRight, false, 'ต้องไม่มี #rulerEdgeHandle');
      assert.equal(m16.value.hasHandleLeft, false, 'ต้องไม่มี #rulerEdgeHandleLeft');
      assert.equal(m16.value.hasPaperHandleRight, false, 'ต้องไม่มี #paperEdgeHandle');
      assert.equal(m16.value.hasPaperHandleLeft, false, 'ต้องไม่มี #paperEdgeHandleLeft');
      assert.equal(m16.value.hasBtnFull, false, 'ต้องไม่มี #btnFullWidth');

      // ขยายกั้นขวา / พื้นที่พิมพ์เป็น 17 ซม.
      const m17 = await send('Runtime.evaluate', {
        expression: `(() => {
          window.eofficeEditor.setTextWidth(17);
          const box = document.querySelector('.editor-box');
          const marks = document.querySelector('#rulerMarks');
          const mkRight = document.querySelector('.mk-right');
          const bodyStyle = getComputedStyle(document.body);
          return {
            textW: window.eofficeEditor.textWidth(),
            bodyTextW: bodyStyle.getPropertyValue('--text-w').trim(),
            boxWidth: box.getBoundingClientRect().width,
            marksWidth: marks.getBoundingClientRect().width,
            marksLeft: marks.getBoundingClientRect().left,
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

      // คืนค่าความกว้างเป็น 16 ซม.
      await send('Runtime.evaluate', {
        expression: `window.eofficeEditor.resetTextWidth()`
      });
      const resetW = await send('Runtime.evaluate', {
        expression: `window.eofficeEditor.textWidth()`,
        returnByValue: true
      });
      assert.equal(resetW.value, 16, 'คืนค่าความกว้างต้องกลับมาเป็น 16 ซม.');

      // ทดสอบการปรับระยะขอบกระดาษ (Margin Preset)
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

      // คืนค่าขอบมาตรฐาน
      await send('Runtime.evaluate', {
        expression: `window.eofficeEditor.setMarginPreset('standard')`
      });

      // ทดสอบเมื่อเปิดเอกสารใหม่ ค่าเริ่มต้นของย่อหน้าบรรทัดแรกต้องเป็น 0.00 ซม. (0cm)
      const indentState = await send('Runtime.evaluate', {
        expression: `(() => {
          document.querySelector('#btnNew').click();
          const s = window.eofficeEditor.settings();
          const mkFirst = document.querySelector('.mk[data-kind="first"]');
          const mkHang = document.querySelector('.mk[data-kind="hang"]');
          const paraInfo = document.querySelector('#paraInfo')?.textContent || '';
          return {
            defaultIndent: s.defaults.indent,
            mkFirstVal: mkFirst?.getAttribute('aria-valuenow'),
            mkFirstLeft: mkFirst?.style.left,
            mkHangVal: mkHang?.getAttribute('aria-valuenow'),
            paraInfo
          };
        })()`,
        returnByValue: true
      });
      assert.strictEqual(indentState.value.defaultIndent, '0cm', 'ค่าเริ่มต้นของย่อหน้าเมื่อเปิดเอกสารใหม่ต้องเป็น 0cm (0.00)');
      assert.strictEqual(indentState.value.mkFirstVal, '0.00', 'ตัวเลื่อนย่อหน้าบรรทัดแรก (.mk-first) ต้องมีค่าเป็น 0.00 ซม.');
      assert.strictEqual(indentState.value.mkFirstLeft, '0cm', 'ตัวเลื่อนย่อหน้าบรรทัดแรก (.mk-first) ต้องอยู่ที่ 0cm');
      assert.strictEqual(indentState.value.mkHangVal, '0.00', 'ตัวเลื่อนกั้นซ้าย (.mk-hang) ต้องมีค่าเป็น 0.00 ซม.');
      assert.ok(indentState.value.paraInfo.includes('ย่อ 0'), 'แถบสถานะต้องแสดงระยะย่อ 0 ซม.');

      ws.close();
    } finally {
      edge.kill('SIGTERM');
    }
  });
});
