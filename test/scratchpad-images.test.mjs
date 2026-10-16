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

describe('Scratchpad Images (กระดาษทดสามารถใส่รูปภาพได้)', () => {
  const tools = readFileSync(join(root, 'eoffice-tools.js'), 'utf8');
  const distSingle = readFileSync(join(root, 'dist/eoffice-editor-single.html'), 'utf8');

  test('โค้ดมีโครงสร้างรองรับการใส่รูปภาพในกระดาษทดครบถ้วน', () => {
    // ป้อนไฟล์ภาพ
    assert.ok(tools.includes('id="scratchFileInput"'), 'ต้องมี input file ซ่อนสำหรับเลือกรูปภาพ');
    assert.ok(tools.includes('data-note-img'), 'ต้องมีปุ่ม data-note-img ใน note-head');
    assert.ok(tools.includes('🖼️ รูป'), 'ต้องมีปุ่ม 🖼️ รูป ใน note-head');

    // พรีวิวรูปภาพในแผ่นกระดาษทด
    assert.ok(tools.includes('.note-images'), 'ต้องมี class .note-images');
    assert.ok(tools.includes('.note-img-wrap'), 'ต้องมี class .note-img-wrap');
    assert.ok(tools.includes('btn-img-ins'), 'ต้องมีปุ่มแทรกรูปเดี่ยวลงเอกสาร');
    assert.ok(tools.includes('btn-img-del'), 'ต้องมีปุ่มลบรูปภาพ');

    // การย่อและปรับขนาดภาพอัตโนมัติ
    assert.ok(tools.includes('readAndOptimizeImage'), 'ต้องมีฟังก์ชันบีบอัดปรับขนาดภาพ');

    // การวางภาพและลากวาง
    assert.ok(tools.includes("scratch.addEventListener('paste'"), 'ต้องรองรับการ Paste (Ctrl+V) ภาพ');
    assert.ok(tools.includes("scratch.addEventListener('drop'"), 'ต้องรองรับ Drag & Drop ภาพ');

    // การแทรกลงเอกสาร
    assert.ok(tools.includes('insertSingleImageAtCursor'), 'ต้องมีฟังก์ชันแทรกรูปเดี่ยวลงในเอกสาร');
    assert.ok(tools.includes('insertNoteAtCursor'), 'ต้องมีฟังก์ชันแทรกทั้งข้อความและรูปภาพทั้งหมดลงในเอกสาร');

    // ไฟล์คอมไพล์ dist/eoffice-editor-single.html มีฟังก์ชันครบถ้วน
    assert.ok(distSingle.includes('id="scratchFileInput"'), 'distSingle ต้องมี scratchFileInput');
    assert.ok(distSingle.includes('data-note-img'), 'distSingle ต้องมี data-note-img');
    assert.ok(distSingle.includes('.note-images'), 'distSingle ต้องมี .note-images');
  });

  test('ฟังก์ชัน normalizeNote รองรับทั้งข้อความเดิม (string) และโครงสร้างใหม่ (object พร้อมรูปภาพ)', () => {
    assert.ok(tools.includes('normalizeNote'), 'ต้องมีฟังก์ชัน normalizeNote');
    assert.ok(tools.includes("typeof n === 'string'"), 'ต้องแปลงข้อมูลโน้ตแบบ string เดิมให้เป็น object');
  });

  test('ทดสอบในเบราว์เซอร์จริง (Edge Headless): แสดงรูปภาพในกระดาษทดและแทรกลงใน Quill ได้อย่างถูกต้อง', async (t) => {
    const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
    if (!existsSync(edgePath)) {
      t.skip('ไม่พบ msedge.exe ข้ามการทดสอบเบราว์เซอร์จริง');
      return;
    }

    const port = 9310 + Math.floor(Math.random() * 40);
    const userData = join(tmpdir(), 'edge-scratch-' + Date.now() + '-' + Math.random().toString(36).slice(2));
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

      // รอให้ eofficeEditor และ scratch พร้อมทำงาน
      await send('Runtime.evaluate', {
        expression: `new Promise((resolve) => {
          const check = () => {
            if (window.eofficeEditor?.scratch) return resolve();
            setTimeout(check, 100);
          };
          check();
        })`,
        awaitPromise: true
      });

      // 1. เปิดกระดาษทด และใส่โน้ตที่มีทั้งข้อความและรูปภาพ 1 รูป
      const samplePng = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAALElEQVRYR+3QQREAAAzCMPunf68T5GBHAuLWrX1wAgQIECBAgAABAgQIECBAgMAr4D8H3XyqRksAAAAASUVORK5CYII=';
      const addRes = await send('Runtime.evaluate', {
        expression: `(() => {
          window.eofficeEditor.scratch.open();
          window.eofficeEditor.scratch.clear();
          window.eofficeEditor.scratch.addNote('ทดสอบข้อความในกระดาษทด', ['${samplePng}']);
          const noteEl = document.querySelector('.note');
          const imgEl = noteEl?.querySelector('.note-images img');
          const btnIns = noteEl?.querySelector('[data-note-ins]');
          const btnImg = noteEl?.querySelector('[data-note-img]');
          const btnInsImg = noteEl?.querySelector('.btn-img-ins');
          const isScratchOpen = document.body.classList.contains('scratch-open');
          return {
            isScratchOpen,
            hasNote: !!noteEl,
            hasImg: !!imgEl,
            hasBtnIns: !!btnIns,
            hasBtnImg: !!btnImg,
            hasBtnInsImg: !!btnInsImg,
            imgSrc: imgEl?.src?.slice(0, 30)
          };
        })()`,
        returnByValue: true
      });

      assert.equal(addRes.value.isScratchOpen, true, 'กระดาษทดต้องเปิดอยู่');
      assert.equal(addRes.value.hasNote, true, 'ต้องมีแผ่นโน้ตในกระดาษทด');
      assert.equal(addRes.value.hasImg, true, 'ต้องแสดงรูปภาพในแผ่นโน้ต');
      assert.equal(addRes.value.hasBtnImg, true, 'ต้องมีปุ่มแทรกรูปภาพ 🖼️ รูป');
      assert.equal(addRes.value.hasBtnInsImg, true, 'ต้องมีปุ่มแทรกเฉพาะรูปนี้');

      // 2. ทดสอบแทรกเฉพาะรูปนี้ลงในเอกสาร
      const insertImgRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const btnInsImg = document.querySelector('.btn-img-ins');
          btnInsImg.click();
          const qlImg = document.querySelector('#editor .ql-editor img');
          return {
            hasQlImg: !!qlImg,
            qlImgSrc: qlImg?.src?.slice(0, 30)
          };
        })()`,
        returnByValue: true
      });

      assert.equal(insertImgRes.value.hasQlImg, true, 'ต้องมีรูปภาพแทรกใน Quill editor');

      // 3. ทดสอบแทรกทั้งแผ่น (ข้อความ + รูปภาพ)
      const insertAllRes = await send('Runtime.evaluate', {
        expression: `(() => {
          // ล้างเอกสาร
          window.eofficeEditor.quill.setText('');
          // กดแทรกในเอกสาร
          const btnIns = document.querySelector('[data-note-ins="0"]');
          btnIns.click();
          const text = window.eofficeEditor.quill.getText();
          const qlImg = document.querySelector('#editor .ql-editor img');
          return {
            text: text.trim(),
            hasQlImg: !!qlImg
          };
        })()`,
        returnByValue: true
      });

      assert.ok(insertAllRes.value.text.includes('ทดสอบข้อความในกระดาษทด'), 'ข้อความจากกระดาษทดต้องถูกแทรกลงใน Quill');
      assert.equal(insertAllRes.value.hasQlImg, true, 'รูปภาพจากกระดาษทดต้องถูกแทรกลงใน Quill พร้อมข้อความ');

      // 4. บันทึกภาพหน้าจอกระดาษทดที่มีรูปภาพ
      const ssRes = await send('Page.captureScreenshot', { format: 'png' });
      if (ssRes && ssRes.data) {
        const artifactDir = 'C:\\Users\\suwit\\.gemini\\antigravity\\brain\\3e316c91-ec18-4b0a-9a1e-3224837de510';
        writeFileSync(join(artifactDir, 'scratchpad-with-image.png'), Buffer.from(ssRes.data, 'base64'));
      }

      ws.close();
    } finally {
      edge.kill('SIGTERM');
    }
  });
});
