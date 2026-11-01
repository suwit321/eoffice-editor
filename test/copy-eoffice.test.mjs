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

describe('Universal e-Office Copy & Cut (การตั้งค่าการคัดลอกทุกประเภทให้เป็นรูปแบบ e-Office เสมอ)', () => {
  const editorHtml = readFileSync(join(root, 'eoffice-editor.html'), 'utf8');
  const distSingle = readFileSync(join(root, 'dist/eoffice-editor-single.html'), 'utf8');

  test('โครงสร้างโค้ดมี Event Listener สำหรับ copy และ cut ครอบคลุมทุกการคัดลอก', () => {
    assert.ok(editorHtml.includes("document.addEventListener('copy'"), 'ต้องมี document.addEventListener copy');
    assert.ok(editorHtml.includes("document.addEventListener('cut'"), 'ต้องมี document.addEventListener cut');
    assert.ok(editorHtml.includes('getEofficeExport'), 'ต้องมีฟังก์ชัน getEofficeExport');
    assert.ok(editorHtml.includes('cutToEoffice'), 'ต้องมีฟังก์ชัน cutToEoffice');
    assert.ok(editorHtml.includes("clip(kind)"), 'ต้องมีฟังก์ชัน clip(kind)');

    // ตรวจสอบใน dist/eoffice-editor-single.html
    assert.ok(distSingle.includes("document.addEventListener('copy'"), 'distSingle ต้องมี copy listener');
    assert.ok(distSingle.includes("document.addEventListener('cut'"), 'distSingle ต้องมี cut listener');
    assert.ok(distSingle.includes('getEofficeExport'), 'distSingle ต้องมี getEofficeExport');
  });

  test('ฟังก์ชัน clip และ commands เชื่อมโยงกับระบบ e-Office อย่างถูกต้อง', () => {
    assert.ok(editorHtml.includes("if (kind === 'cut') return cutToEoffice()"), 'clip(cut) ต้องเรียก cutToEoffice');
    assert.ok(editorHtml.includes("if (kind === 'copy') return copyToEoffice('auto')"), 'clip(copy) ต้องเรียก copyToEoffice');
  });

  test('ทดสอบการคัดลอกและตัดในเบราว์เซอร์จริง (Edge Headless)', async (t) => {
    if (!edgePath) {
      t.skip('ไม่พบ msedge.exe ข้ามการทดสอบเบราว์เซอร์');
      return;
    }

    const port = 9330 + Math.floor(Math.random() * 30);
    const userData = join(tmpdir(), 'edge-copy-' + Date.now() + '-' + Math.random().toString(36).slice(2));
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

      // รอให้ Quill พร้อม
      await send('Runtime.evaluate', {
        expression: `new Promise((resolve) => {
          let count = 0;
          const check = () => {
            if (window.eofficeEditor?.quill) return resolve();
            if (++count > 50) return resolve();
            setTimeout(check, 100);
          };
          check();
        })`,
        awaitPromise: true
      });

      // 1. ทดสอบการจำลองเหตุการณ์ Copy (Ctrl+C / คลิกขวาคัดลอก) เมื่อมีการเลือกข้อความ
      const copyWithSelRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const quill = window.eofficeEditor.quill;
          quill.setText('บันทึกข้อความราชการ เรื่อง ทดสอบการคัดลอกรูปแบบ e-Office\\n');
          quill.formatText(0, 15, { bold: true });
          quill.setSelection(0, 15);

          const r = { index: 0, length: 15 };
          const exp = window.eofficeEditor.getEofficeExport('auto');

          const dt = new DataTransfer();
          const evt = new ClipboardEvent('copy', { clipboardData: dt, bubbles: true, cancelable: true });
          const qlEditor = document.querySelector('#editor .ql-editor');
          qlEditor.dispatchEvent(evt);

          return {
            defaultPrevented: evt.defaultPrevented,
            html: dt.getData('text/html'),
            text: dt.getData('text/plain'),
            expHtml: exp?.html
          };
        })()`,
        returnByValue: true
      });

      const copyWithSel = copyWithSelRes.value;
      assert.equal(copyWithSel.defaultPrevented, true, 'เหตุการณ์ copy ต้องถูก preventDefault');
      assert.ok(copyWithSel.html.includes('TH Sarabun PSK'), 'HTML ที่คัดลอกต้องมีฟอนต์ TH Sarabun PSK');
      assert.ok(copyWithSel.html.includes('บันทึกข้อความรา'), 'HTML ที่คัดลอกต้องมีข้อความที่เลือก');
      assert.ok(copyWithSel.text.includes('บันทึกข้อความรา'), 'Text ที่คัดลอกต้องมีข้อความที่เลือก');

      // 2. ทดสอบการจำลองเหตุการณ์ Cut (Ctrl+X / คลิกขวาตัด)
      const cutRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const quill = window.eofficeEditor.quill;
          quill.setText('ข้อความที่จะถูกตัดออก เหลือข้อความนี้ไว้\\n');
          quill.setSelection(0, 20); // เลือก 'ข้อความที่จะถูกตัดออก'

          const dt = new DataTransfer();
          const evt = new ClipboardEvent('cut', { clipboardData: dt, bubbles: true, cancelable: true });
          const qlEditor = document.querySelector('#editor .ql-editor');
          qlEditor.dispatchEvent(evt);

          return {
            defaultPrevented: evt.defaultPrevented,
            html: dt.getData('text/html'),
            text: dt.getData('text/plain'),
            remainingText: quill.getText()
          };
        })()`,
        returnByValue: true
      });

      const cut = cutRes.value;
      assert.equal(cut.defaultPrevented, true, 'เหตุการณ์ cut ต้องถูก preventDefault');
      assert.ok(cut.html.includes('TH Sarabun PSK'), 'HTML ที่ตัดต้องมีฟอนต์ TH Sarabun PSK');
      assert.ok(cut.text.includes('ข้อความที่จะถูกตัด'), 'Text ที่ตัดต้องมีข้อความที่เลือก');
      assert.ok(!cut.remainingText.includes('ข้อความที่จะถูกตัดออก'), 'ข้อความที่เลือกต้องถูกลบออกจาก Quill หลังตัด');
      assert.ok(cut.remainingText.includes('เหลือข้อความนี้ไว้'), 'ข้อความส่วนที่เหลือต้องยังอยู่');

      // 3. ทดสอบการจำลองเหตุการณ์ Copy เมื่อไม่มีการเลือกข้อความ (กด Ctrl+C เพื่อคัดลอกทั้งหน้า)
      const copyPageRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const quill = window.eofficeEditor.quill;
          quill.setText('เนื้อหาเอกสารทั้งหน้าสำหรับคัดลอกทั้งหมด\\n');
          quill.setSelection(0, 0); // ไม่มีการเลือกข้อความ (ความยาว 0)

          const dt = new DataTransfer();
          const evt = new ClipboardEvent('copy', { clipboardData: dt, bubbles: true, cancelable: true });
          const qlEditor = document.querySelector('#editor .ql-editor');
          qlEditor.dispatchEvent(evt);

          return {
            defaultPrevented: evt.defaultPrevented,
            html: dt.getData('text/html'),
            text: dt.getData('text/plain')
          };
        })()`,
        returnByValue: true
      });

      const copyPage = copyPageRes.value;
      assert.equal(copyPage.defaultPrevented, true, 'เหตุการณ์ copy ทั้งหน้าต้องถูก preventDefault');
      assert.ok(copyPage.html.includes('TH Sarabun PSK'), 'HTML ต้องคงรูปแบบ TH Sarabun PSK');
      assert.ok(copyPage.html.includes('เนื้อหาเอกสารทั้งหน้าสำหรับคัดลอกทั้งหมด'), 'HTML ต้องมีเนื้อหาทั้งหน้า');
      assert.ok(copyPage.text.includes('เนื้อหาเอกสารทั้งหน้าสำหรับคัดลอกทั้งหมด'), 'Text ต้องมีเนื้อหาทั้งหน้า');

      // 4. ทดสอบการจำลองเหตุการณ์ Copy บนตารางในเอดิเตอร์
      const copyTableRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const quill = window.eofficeEditor.quill;
          window.eofficeEditor.insertTable(2, 2);
          const tbl = window.eofficeEditor.getActiveTable();
          if (tbl) {
            tbl.querySelectorAll('td')[0].textContent = 'หัวข้อ 1';
            tbl.querySelectorAll('td')[1].textContent = 'หัวข้อ 2';
          }
          // คัดลอกทั้งหน้าที่มีตาราง
          quill.setSelection(0, 0);

          const dt = new DataTransfer();
          const evt = new ClipboardEvent('copy', { clipboardData: dt, bubbles: true, cancelable: true });
          const qlEditor = document.querySelector('#editor .ql-editor');
          qlEditor.dispatchEvent(evt);

          return {
            defaultPrevented: evt.defaultPrevented,
            html: dt.getData('text/html'),
            text: dt.getData('text/plain')
          };
        })()`,
        returnByValue: true
      });

      const copyTable = copyTableRes.value;
      assert.equal(copyTable.defaultPrevented, true, 'เหตุการณ์ copy ตารางต้องถูก preventDefault');
      assert.ok(copyTable.html.includes('<table'), 'HTML ที่คัดลอกต้องมีแท็ก <table');
      assert.ok(copyTable.html.includes('border-collapse'), 'HTML ที่คัดลอกต้องมี border-collapse');
      assert.ok(copyTable.html.includes('TH Sarabun PSK'), 'ตารางต้องใช้ฟอนต์ TH Sarabun PSK');
      assert.ok(copyTable.html.includes('หัวข้อ 1'), 'ตารางต้องมีข้อมูลในเซลล์');

      // 5. ทดสอบปุ่มคัดลอกบน Toolbar ([data-cmd="copy"]) และปุ่มตัด ([data-cmd="cut"])
      const toolbarButtonsRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const btnCopy = document.querySelector('[data-cmd="copy"]');
          const btnCut = document.querySelector('[data-cmd="cut"]');
          return {
            hasBtnCopy: !!btnCopy,
            hasBtnCut: !!btnCut
          };
        })()`,
        returnByValue: true
      });
      assert.equal(toolbarButtonsRes.value.hasBtnCopy, true, 'ต้องมีปุ่ม copy บน Toolbar');
      assert.equal(toolbarButtonsRes.value.hasBtnCut, true, 'ต้องมีปุ่ม cut บน Toolbar');

      // 6. ทดสอบว่าการ Copy ใน input ภายนอก (เช่น #docName) จะไม่ถูกดักจับ
      const copyInputRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const input = document.querySelector('#docName');
          input.value = 'ชื่อเอกสารทดสอบ';
          input.focus();
          input.select();

          const dt = new DataTransfer();
          const evt = new ClipboardEvent('copy', { clipboardData: dt, bubbles: true, cancelable: true });
          input.dispatchEvent(evt);

          return {
            defaultPrevented: evt.defaultPrevented
          };
        })()`,
        returnByValue: true
      });

      assert.equal(copyInputRes.value.defaultPrevented, false, 'การ Copy ใน input ภายนอกต้องไม่ถูกดักจับ (ปล่อยให้เบราว์เซอร์จัดการตามปกติ)');

      ws.close();
    } finally {
      edge.kill('SIGTERM');
    }
  });
});
