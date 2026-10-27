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

describe('Hidden Red Dot Visitor Counter & PIN 4321 Popup (ตัวนับการเข้าชมเว็บไซต์แบบจุดแดงซ่อนขวาล่าง)', () => {
  const html = readFileSync(join(root, 'eoffice-editor.html'), 'utf8');
  const distSingle = readFileSync(join(root, 'dist/eoffice-editor-single.html'), 'utf8');

  test('มีจุดแดงแบบซ่อน (#visitDot.st-visit-dot) ในแถบแสดงสถานะขวาล่าง (.statusbar)', () => {
    assert.ok(html.includes('id="visitDot" class="st-visit-dot"'), 'ต้องมีปุ่มจุดแดง #visitDot ในแถบสถานะ');
    const statusFooterMatch = /<footer class="statusbar">[\s\S]*?id="visitDot"[\s\S]*?<\/footer>/.test(html);
    assert.ok(statusFooterMatch, '#visitDot ต้องอยู่ใน <footer class="statusbar">');
  });

  test('มีหน้าต่าง Popup (#visitModal) พร้อมส่วนกรอกรหัสผ่าน 4321 และหน้ารายละเอียดการเข้าชม', () => {
    assert.ok(html.includes('id="visitModal"'), 'ต้องมี #visitModal');
    assert.ok(html.includes('id="visitAuthView"'), 'ต้องมีหน้าจอกรอกรหัสผ่าน #visitAuthView');
    assert.ok(html.includes('id="visitPinInput"'), 'ต้องมีช่องกรอกรหัส #visitPinInput');
    assert.ok(html.includes('id="visitPinSubmit"'), 'ต้องมีปุ่มตกลง #visitPinSubmit');
    assert.ok(html.includes('id="visitStatsView"'), 'ต้องมีหน้ารายละเอียดสถิติ #visitStatsView');
    assert.ok(html.includes("const VISIT_PIN = '4321'"), 'ต้องตรวจสอบรหัสผ่าน 4321');
    assert.ok(distSingle.includes('id="visitDot"'), 'distSingle ต้องมี #visitDot');
    assert.ok(distSingle.includes('id="visitModal"'), 'distSingle ต้องมี #visitModal');
  });

  test('ทดสอบในเบราว์เซอร์จริง (Edge Headless): นับจำนวนเข้าชมอัตโนมัติ กดจุดแดง ตรวจสอบรหัสผิด/ถูก (4321) และแสดงรายละเอียด', async (t) => {
    const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
    if (!existsSync(edgePath)) {
      t.skip('ไม่พบ msedge.exe ข้ามการทดสอบเบราว์เซอร์จริง');
      return;
    }

    const port = 9330 + Math.floor(Math.random() * 30);
    const userData = join(tmpdir(), 'edge-visit-' + Date.now() + '-' + Math.random().toString(36).slice(2));
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

      // รอให้ eofficeEditor พร้อมทำงาน
      for (let i = 0; i < 40; i++) {
        try {
          const ready = await send('Runtime.evaluate', {
            expression: `Boolean(window.eofficeEditor && window.eofficeEditor.visitorStats)`,
            returnByValue: true
          });
          if (ready?.value) break;
        } catch {}
        await new Promise(r => setTimeout(r, 100));
      }

      // 1. ตรวจสอบว่าบันทึกการเข้าชมอัตโนมัติแล้ว และจำลองการเข้าชมเพิ่มอีก 2 ครั้งเพื่อดูสถิติ
      const initStats = await send('Runtime.evaluate', {
        expression: `(() => {
          window.eofficeEditor.visitorStats.recordVisit();
          window.eofficeEditor.visitorStats.recordVisit();
          const st = window.eofficeEditor.visitorStats.getStats();
          const dot = document.querySelector('#visitDot');
          const r = dot.getBoundingClientRect();
          return {
            totalVisits: st.totalVisits,
            historyLen: st.history.length,
            dotRight: r.right,
            dotBottom: r.bottom,
            winW: window.innerWidth,
            winH: window.innerHeight
          };
        })()`,
        returnByValue: true
      });

      assert.strictEqual(initStats.value.totalVisits, 3, 'ต้องนับจำนวนการเข้าชมได้ถูกต้อง (เริ่มต้น 1 + จำลองเพิ่ม 2 = 3)');
      assert.ok(initStats.value.dotRight > initStats.value.winW - 40, 'จุดแดงต้องอยู่มุมขวาล่างในแถบสถานะ');

      // 2. คลิกที่จุดแดง #visitDot -> ต้องเปิดหน้าต่างถามรหัสผ่าน และยังไม่แสดงสถิติ
      await send('Runtime.evaluate', {
        expression: `document.querySelector('#visitDot').click()`
      });
      await new Promise(r => setTimeout(r, 150));

      const authOpenState = await send('Runtime.evaluate', {
        expression: `(() => {
          return {
            modalHidden: document.querySelector('#visitModal').hidden,
            authHidden: document.querySelector('#visitAuthView').hidden,
            statsHidden: document.querySelector('#visitStatsView').hidden
          };
        })()`,
        returnByValue: true
      });
      assert.strictEqual(authOpenState.value.modalHidden, false, 'Popup ต้องเปิดขึ้นเมื่อคลิกจุดแดง');
      assert.strictEqual(authOpenState.value.authHidden, false, 'ต้องแสดงหน้าจอกรอกรหัสผ่านก่อน');
      assert.strictEqual(authOpenState.value.statsHidden, true, 'หน้ารายละเอียดสถิติต้องถูกซ่อนไว้ก่อนกรอกรหัส');

      // 3. ทดสอบกรอกรหัสผิด (เช่น 9999) -> ต้องไม่เปิดหน้ารายละเอียด และแจ้งเตือนรหัสผิด
      const wrongPinState = await send('Runtime.evaluate', {
        expression: `(() => {
          const input = document.querySelector('#visitPinInput');
          input.value = '9999';
          document.querySelector('#visitPinSubmit').click();
          return {
            errHidden: document.querySelector('#visitPinErr').hidden,
            statsHidden: document.querySelector('#visitStatsView').hidden
          };
        })()`,
        returnByValue: true
      });
      assert.strictEqual(wrongPinState.value.errHidden, false, 'เมื่อกรอกรหัสผิดต้องแสดงข้อความแจ้งเตือน');
      assert.strictEqual(wrongPinState.value.statsHidden, true, 'เมื่อกรอกรหัสผิดต้องยังไม่แสดงหน้ารายละเอียด');

      // 4. กรอกรหัสที่ถูกต้อง 4321 -> ต้องเปิดแสดง Popup รายละเอียดการเข้าชมทันที
      const unlockState = await send('Runtime.evaluate', {
        expression: `(() => {
          const input = document.querySelector('#visitPinInput');
          input.value = '4321';
          document.querySelector('#visitPinSubmit').click();
          return {
            authHidden: document.querySelector('#visitAuthView').hidden,
            statsHidden: document.querySelector('#visitStatsView').hidden,
            totalText: document.querySelector('#vTotalCount').textContent.trim(),
            todayText: document.querySelector('#vTodayCount').textContent.trim(),
            chartBars: document.querySelectorAll('#vChart7Days .visit-bar-col').length,
            logRows: document.querySelectorAll('#vHistoryBody tr').length
          };
        })()`,
        returnByValue: true
      });

      assert.strictEqual(unlockState.value.authHidden, true, 'เมื่อกรอกรหัส 4321 ถูกต้อง หน้ากรอกรหัสต้องซ่อนลง');
      assert.strictEqual(unlockState.value.statsHidden, false, 'เมื่อกรอกรหัส 4321 ถูกต้อง ต้องแสดงหน้ารายละเอียดการเข้าชม');
      assert.strictEqual(unlockState.value.totalText, '3', 'ยอดเข้าชมทั้งหมดต้องแสดงเป็น 3');
      assert.strictEqual(unlockState.value.todayText, '3', 'ยอดเข้าชมวันนี้ต้องแสดงเป็น 3');
      assert.strictEqual(unlockState.value.chartBars, 7, 'ต้องแสดงกราฟย้อนหลัง 7 วันครบ 7 แท่ง');
      assert.strictEqual(unlockState.value.logRows, 3, 'ตารางประวัติการเข้าชมต้องแสดงรายการเข้าชมครบ 3 รายการ');

      // ถ่ายภาพหน้าจอ Popup รายละเอียดการเข้าชมเก็บไว้ใน Artifact
      const ssRes = await send('Page.captureScreenshot', { format: 'png' });
      if (ssRes && ssRes.data) {
        const artifactDir = 'C:\\Users\\suwit\\.gemini\\antigravity\\brain\\3e316c91-ec18-4b0a-9a1e-3224837de510';
        writeFileSync(join(artifactDir, 'visitor-counter-popup.png'), Buffer.from(ssRes.data, 'base64'));
      }

      ws.close();
    } finally {
      edge.kill('SIGKILL');
    }
  });
});
