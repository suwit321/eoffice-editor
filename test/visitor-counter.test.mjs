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

  test('เชื่อมต่อ Global Counter API (แนวทางที่ 1) เพื่อนับยอดเข้าชมรวมจากผู้ใช้และอุปกรณ์ทั้งหมด', () => {
    assert.ok(html.includes('https://abacus.jasoncameron.dev'), 'eoffice-editor.html ต้องมีโดเมน Counter API');
    assert.ok(distSingle.includes('https://abacus.jasoncameron.dev'), 'distSingle ต้องมีโดเมน Counter API');
    assert.ok(html.includes('syncGlobalCounter'), 'ต้องมีฟังก์ชัน syncGlobalCounter');
    assert.ok(html.includes('id="vSyncBar"'), 'ต้องมีแถบแสดงสถานะซิงค์คลาวด์ #vSyncBar');
    assert.ok(html.includes('id="vRefreshBtn"'), 'ต้องมีปุ่มรีเฟรชข้อมูล #vRefreshBtn');

    // ตรวจสอบ CSP ใน dist/index.html และ dist/_headers ต้องอนุญาต https://abacus.jasoncameron.dev
    const distIndex = readFileSync(join(root, 'dist/index.html'), 'utf8');
    const distHeaders = readFileSync(join(root, 'dist/_headers'), 'utf8');
    assert.ok(distIndex.includes('https://abacus.jasoncameron.dev'), 'dist/index.html CSP ต้องอนุญาต Counter API');
    assert.ok(distHeaders.includes('https://abacus.jasoncameron.dev'), 'dist/_headers CSP ต้องอนุญาต Counter API');
  });

  test('ระบบป้องกันการนับซ้ำ: ตรวจสอบ IP เดียวกันไม่นับซ้ำในชั่วโมงเดียวกัน', () => {
    assert.ok(html.includes('localHourKey'), 'ต้องมีฟังก์ชัน localHourKey');
    assert.ok(html.includes('api64.ipify.org'), 'ต้องมีปลายทางดึง IP api64.ipify.org');
    assert.ok(html.includes('maskIp'), 'ต้องมีฟังก์ชัน maskIp สำหรับซ่อน IP เพื่อความเป็นส่วนตัว');
    assert.ok(distSingle.includes('localHourKey'), 'distSingle ต้องมี localHourKey');
    assert.ok(distSingle.includes('isDuplicateHour'), 'distSingle ต้องมีตัวแปรตรวจ isDuplicateHour');

    // ตรวจสอบ CSP อนุญาต api64.ipify.org และ api.ipify.org
    const distIndex = readFileSync(join(root, 'dist/index.html'), 'utf8');
    const distHeaders = readFileSync(join(root, 'dist/_headers'), 'utf8');
    assert.ok(distIndex.includes('https://api64.ipify.org'), 'dist/index.html CSP ต้องอนุญาต api64.ipify.org');
    assert.ok(distHeaders.includes('https://api64.ipify.org'), 'dist/_headers CSP ต้องอนุญาต api64.ipify.org');
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

      // 1. ตรวจสอบว่าบันทึกการเข้าชมอัตโนมัติแล้ว และทดสอบว่าเข้าซ้ำในชั่วโมงเดียวกันไม่นับเพิ่ม แต่เข้าชั่วโมงใหม่นับเพิ่ม
      const initStats = await send('Runtime.evaluate', {
        expression: `(() => {
          // เข้าชมซ้ำในชั่วโมงเดียวกัน -> ต้องไม่เพิ่ม totalVisits
          window.eofficeEditor.visitorStats.recordVisit();
          const stDup = window.eofficeEditor.visitorStats.getStats();

          // เข้าชมในชั่วโมงใหม่ (forceNewHour) -> ต้องเพิ่ม totalVisits
          window.eofficeEditor.visitorStats.recordVisit({ forceNewHour: true });
          const stNew = window.eofficeEditor.visitorStats.getStats();

          const dot = document.querySelector('#visitDot');
          const r = dot.getBoundingClientRect();
          return {
            totalAfterDup: stDup.totalVisits,
            totalFinal: stNew.totalVisits,
            historyLen: stNew.history.length,
            dotRight: r.right,
            winW: window.innerWidth
          };
        })()`,
        returnByValue: true
      });

      assert.strictEqual(initStats.value.totalAfterDup, 1, 'เข้าชมซ้ำในชั่วโมงเดียวกันต้องไม่นับซ้ำ (คงที่ 1)');
      assert.strictEqual(initStats.value.totalFinal, 2, 'เข้าชมในชั่วโมงใหม่ต้องนับเพิ่มเป็น 2');
      assert.strictEqual(initStats.value.historyLen, 3, 'ประวัติการเข้าชมต้องบันทึกทั้ง 3 ครั้ง');
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
            logRows: document.querySelectorAll('#vHistoryBody tr').length,
            syncBarText: document.querySelector('#vSyncText')?.textContent || '',
            hasRefreshBtn: Boolean(document.querySelector('#vRefreshBtn'))
          };
        })()`,
        returnByValue: true
      });

      assert.strictEqual(unlockState.value.authHidden, true, 'เมื่อกรอกรหัส 4321 ถูกต้อง หน้ากรอกรหัสต้องซ่อนลง');
      assert.strictEqual(unlockState.value.statsHidden, false, 'เมื่อกรอกรหัส 4321 ถูกต้อง ต้องแสดงหน้ารายละเอียดการเข้าชม');
      assert.ok(Number(unlockState.value.totalText.replace(/,/g, '')) >= 2, 'ยอดเข้าชมทั้งหมดต้องเป็นตัวเลข >= 2');
      assert.ok(Number(unlockState.value.todayText.replace(/,/g, '')) >= 2, 'ยอดเข้าชมวันนี้ต้องเป็นตัวเลข >= 2');
      assert.ok(unlockState.value.syncBarText.length > 0, 'ต้องมีข้อความสถานะซิงค์คลาวด์');
      assert.strictEqual(unlockState.value.hasRefreshBtn, true, 'ต้องมีปุ่มรีเฟรชยอดเข้าชม');
      assert.strictEqual(unlockState.value.chartBars, 7, 'ต้องแสดงกราฟย้อนหลัง 7 วันครบ 7 แท่ง');
      assert.strictEqual(unlockState.value.logRows, 3, 'ตารางประวัติการเข้าชมต้องแสดงรายการเข้าชมครบ 3 รายการ');

      // 5. ทดสอบการบันทึกประวัติสูงสุด 100 รายการ และการระบุเครื่องเดียวกัน vs เครื่องอื่น
      const deviceTest = await send('Runtime.evaluate', {
        expression: `(() => {
          const vs = window.eofficeEditor.visitorStats;
          const myDevId = vs.getDeviceId();

          // จำลองการเข้าชมจากเครื่องอื่น (other device) 2 เครื่อง
          vs.recordVisit({ forceNewHour: true, deviceId: 'DEV-REMOTE1' });
          vs.recordVisit({ forceNewHour: true, deviceId: 'DEV-REMOTE2' });

          // ตรวจสอบความจุประวัติสูงสุด 100 รายการ โดยจำลองเพิ่มข้อมูลจนเกิน 100
          for (let i = 0; i < 110; i++) {
            vs.recordVisit({ forceNewHour: true, deviceId: (i % 2 === 0) ? myDevId : 'DEV-REMOTE' + (i % 5) });
          }
          const stCapped = vs.getStats();

          // รีเฟรชหน้าจอแสดงผล
          vs.unlock('4321');
          const doc = document;
          const myBadgeText = doc.querySelector('#vMyDevBadge')?.textContent || '';
          const totalInTable = doc.querySelectorAll('#vHistoryBody tr').length;
          const thisRowsCount = doc.querySelectorAll('#vHistoryBody tr.row-this-device').length;
          const otherRowsCount = doc.querySelectorAll('#vHistoryBody tr.row-other-device').length;

          // ทดสอบตัวกรอง: กดเฉพาะเครื่องนี้ (#vFilterThis)
          doc.querySelector('#vFilterThis')?.click();
          const filteredThisCount = doc.querySelectorAll('#vHistoryBody tr').length;
          const allAreThis = Array.from(doc.querySelectorAll('#vHistoryBody tr')).every(tr => tr.classList.contains('row-this-device'));

          // ทดสอบตัวกรอง: กดเฉพาะเครื่องอื่น (#vFilterOther)
          doc.querySelector('#vFilterOther')?.click();
          const filteredOtherCount = doc.querySelectorAll('#vHistoryBody tr').length;
          const allAreOther = Array.from(doc.querySelectorAll('#vHistoryBody tr')).every(tr => tr.classList.contains('row-other-device'));

          // คืนค่าตัวกรองทั้งหมด (#vFilterAll)
          doc.querySelector('#vFilterAll')?.click();

          return {
            myDevId,
            historyLength: stCapped.history.length,
            myBadgeText,
            totalInTable,
            thisRowsCount,
            otherRowsCount,
            filteredThisCount,
            allAreThis,
            filteredOtherCount,
            allAreOther,
          };
        })()`,
        returnByValue: true
      });

      assert.ok(deviceTest.value.myDevId.startsWith('DEV-'), 'รหัสเครื่องต้องขึ้นต้นด้วย DEV-');
      assert.strictEqual(deviceTest.value.historyLength, 100, 'ประวัติการเข้าชมต้องบันทึกได้สูงสุด 100 รายการ');
      assert.strictEqual(deviceTest.value.totalInTable, 100, 'ตารางต้องแสดงรายการสูงสุด 100 รายการ');
      assert.ok(deviceTest.value.thisRowsCount > 0, 'ต้องพบรายการที่เป็นเครื่องนี้');
      assert.ok(deviceTest.value.otherRowsCount > 0, 'ต้องพบรายการที่เป็นเครื่องอื่น');
      assert.strictEqual(deviceTest.value.allAreThis, true, 'เมื่อเลือกกรองเครื่องนี้ ทุกแถวต้องเป็นคลาส row-this-device');
      assert.strictEqual(deviceTest.value.allAreOther, true, 'เมื่อเลือกกรองเครื่องอื่น ทุกแถวต้องเป็นคลาส row-other-device');
      assert.ok(deviceTest.value.myBadgeText.includes(deviceTest.value.myDevId), 'ป้ายด้านบนต้องแสดงรหัสเครื่องนี้');

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
