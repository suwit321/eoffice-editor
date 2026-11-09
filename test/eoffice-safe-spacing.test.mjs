import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { spawn } from 'node:child_process';
import { tmpdir } from 'node:os';
import http from 'node:http';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

const edgePaths = [
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
];
const edgePath = edgePaths.find(p => existsSync(p));

describe('E-Office Safe Spacing & Word Protection (วิจัยและพัฒนาระบบช่องไฟและล็อกคำสำหรับ e-Office)', () => {
  const editorHtml = readFileSync(join(root, 'eoffice-editor.html'), 'utf8');
  const toolsJs = readFileSync(join(root, 'eoffice-tools.js'), 'utf8');
  const distSingle = readFileSync(join(root, 'dist/eoffice-editor-single.html'), 'utf8');

  test('โครงสร้างโค้ดมีระบบรองรับทั้ง 3 ข้อครบถ้วน', () => {
    // ข้อ 1: โหมดเซฟ e-Office (บีบฟอนต์ 15.8 - 15.9 pt)
    assert.ok(editorHtml.includes('data-fs-safe="15.9"'), 'ต้องมีปุ่มบีบฟอนต์ 15.9 pt');
    assert.ok(editorHtml.includes('data-fs-safe="15.85"'), 'ต้องมีปุ่มบีบฟอนต์ 15.85 pt');
    assert.ok(editorHtml.includes('data-fs-safe="15.8"'), 'ต้องมีปุ่มบีบฟอนต์ 15.8 pt');
    assert.ok(editorHtml.includes('data-fs-safe="16"'), 'ต้องมีปุ่มคืนฟอนต์ 16 pt ปกติ');
    assert.ok(editorHtml.includes('updateSafeFsButtons'), 'ต้องมีฟังก์ชัน updateSafeFsButtons');

    // ข้อ 2: MSO Letter Spacing Redundancy
    assert.ok(editorHtml.includes('mso-letter-spacing'), 'exportHtml ต้องใส่ mso-letter-spacing');

    // ข้อ 3: Auto Word Joiner สำหรับคำราชการสำคัญ
    assert.ok(editorHtml.includes('AUTO_JOIN_WORDS'), 'ต้องมี AUTO_JOIN_WORDS');
    assert.ok(editorHtml.includes('toJoinedWord'), 'ต้องมีฟังก์ชัน toJoinedWord');
    assert.ok(editorHtml.includes('id="lsLockWords"'), 'ต้องมีปุ่มล็อกคำราชการในแผงช่องไฟ');
    assert.ok(toolsJs.includes('ลงวันที่'), 'NOBREAK_DEFAULT ใน eoffice-tools.js ต้องมีคำว่า ลงวันที่');
    assert.ok(toolsJs.includes('ข้อหารือ'), 'NOBREAK_DEFAULT ใน eoffice-tools.js ต้องมีคำว่า ข้อหารือ');
    assert.ok(toolsJs.includes('คำปรึกษา'), 'NOBREAK_DEFAULT ใน eoffice-tools.js ต้องมีคำว่า คำปรึกษา');
  });

  test('ทดสอบในเบราว์เซอร์จริง (Edge Headless): การทำงานของ Safe Font-Size, MSO Spacing และ Auto Word Joiner', async (t) => {
    if (!edgePath) {
      t.skip('ไม่พบ msedge.exe ข้ามการทดสอบเบราว์เซอร์');
      return;
    }

    const port = 9380 + Math.floor(Math.random() * 20);
    const userData = join(tmpdir(), 'edge-safe-ls-' + Date.now() + '-' + Math.random().toString(36).slice(2));
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

      // รอ editor พร้อม
      await new Promise(r => setTimeout(r, 600));

      // Test Case A: ทดสอบการฝัง Word Joiner อัตโนมัติใน exportHtml
      const testA = await send('Runtime.evaluate', {
        expression: `(() => {
          const q = window.eofficeEditor.quill;
          q.setText('เอกสารนี้ลงวันที่ ๗ กุมภาพันธ์ เพื่อตอบข้อหารือ และให้คำปรึกษาแก่หน่วยรับตรวจ\\n');
          const exported = window.eofficeEditor.cleanHtml();
          const hasJoinedDate = exported.includes('ลง\u2060วัน\u2060ที่');
          const hasJoinedHarue = exported.includes('ข้อหา\u2060รือ');
          const hasJoinedPrueksa = exported.includes('คำ\u2060ปรึกษา');
          const hasJoinedNหน่วย = exported.includes('หน่วย\u2060รับ\u2060ตรวจ');
          return JSON.stringify({ exported, hasJoinedDate, hasJoinedHarue, hasJoinedPrueksa, hasJoinedNหน่วย });
        })()`,
        returnByValue: true
      });
      const resA = JSON.parse(testA.value);
      assert.ok(resA.hasJoinedDate, 'exportHtml ต้องฝัง Word Joiner ในคำว่า "ลงวันที่"');
      assert.ok(resA.hasJoinedHarue, 'exportHtml ต้องฝัง Word Joiner ในคำว่า "ข้อหารือ"');
      assert.ok(resA.hasJoinedPrueksa, 'exportHtml ต้องฝัง Word Joiner ในคำว่า "คำปรึกษา"');
      assert.ok(resA.hasJoinedNหน่วย, 'exportHtml ต้องฝัง Word Joiner ในคำว่า "หน่วยรับตรวจ"');

      // Test Case B: ทดสอบ mso-letter-spacing และ paragraph-level redundancy
      const testB = await send('Runtime.evaluate', {
        expression: `(() => {
          const q = window.eofficeEditor.quill;
          q.setText('ข้อความทดสอบช่องไฟแบบแคบพิเศษ\\n');
          q.formatText(0, 31, 'letterspacing', '-0.02em');
          const exported = window.eofficeEditor.cleanHtml();
          const hasMso = exported.includes('mso-letter-spacing');
          const div = document.createElement('div');
          div.innerHTML = exported;
          const p = div.querySelector('p');
          const pHasLs = p && p.style.letterSpacing === '-0.02em';
          const pHasMso = p && Boolean(p.getAttribute('style')?.includes('mso-letter-spacing'));
          return JSON.stringify({ exported, hasMso, pHasLs, pHasMso });
        })()`,
        returnByValue: true
      });
      const resB = JSON.parse(testB.value);
      assert.ok(resB.hasMso, 'exportHtml ต้องมี mso-letter-spacing');
      assert.ok(resB.pHasLs, 'แท็ก <p> ต้องได้รับ letter-spacing ร่วมด้วยเพื่อป้องกัน tag unwrap');
      assert.ok(resB.pHasMso, 'แท็ก <p> ต้องได้รับ mso-letter-spacing ร่วมด้วย');

      // Test Case C: ทดสอบปุ่ม Safe Font-Size ในแผงช่องไฟ
      const testC = await send('Runtime.evaluate', {
        expression: `(() => {
          const q = window.eofficeEditor.quill;
          q.setText('ทดสอบการกดปุ่มโหมดเซฟ e-Office บีบฟอนต์ 15.85 pt\\n');
          q.setSelection(0, 48);

          // เปิดแผงช่องไฟ
          const btnLs = document.querySelector('[data-menu=\"m-charspace\"]');
          if (btnLs) btnLs.click();

          // กดปุ่ม 15.85 pt
          const btn1585 = document.querySelector('[data-fs-safe=\"15.85\"]');
          if (btn1585) btn1585.click();

          const fmt = q.getFormat(0, 10);
          const activeSafeBtn = document.querySelector('[data-fs-safe=\"15.85\"][aria-checked=\"true\"]');
          const exported = window.eofficeEditor.cleanHtml();
          const has1585 = exported.includes('15.85pt');

          return JSON.stringify({
            fontSize: fmt.size,
            hasActiveAttr: Boolean(activeSafeBtn),
            has1585Exported: has1585
          });
        })()`,
        returnByValue: true
      });
      const resC = JSON.parse(testC.value);
      assert.equal(resC.fontSize, '15.85pt', 'Quill ต้องถูกฟอร์แมตขนาดฟอนต์เป็น 15.85pt');
      assert.ok(resC.hasActiveAttr, 'ปุ่ม 15.85 pt ต้องมี aria-checked="true"');
      assert.ok(resC.has1585Exported, 'exportHtml ต้องคงค่า font-size 15.85pt ไว้');

      // Test Case D: ทดสอบปุ่มล็อกคำราชการในแผงช่องไฟ
      const testD = await send('Runtime.evaluate', {
        expression: `(() => {
          const q = window.eofficeEditor.quill;
          q.setText('ออกคำสั่งลงวันที่ ๗ กุมภาพันธ์\\n');
          q.setSelection(0, 30);

          const btnLock = document.querySelector('#lsLockWords');
          if (btnLock) btnLock.click();

          const textAfter = q.getText();
          const hasWJ = textAfter.includes('\\u2060');
          return JSON.stringify({ hasWJ, textAfter });
        })()`,
        returnByValue: true
      });
      const resD = JSON.parse(testD.value);
      assert.ok(resD.hasWJ, 'การกดปุ่ม #lsLockWords ต้องฝัง Word Joiner ลงในข้อความจริงของ Quill');

      ws.close();
    } finally {
      edge.kill('SIGKILL');
    }
  });
});
