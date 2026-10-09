import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { spawn } from 'node:child_process';
import http from 'node:http';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

describe('Table Enhancement (เครื่องมือตาราง วาดตาราง ย้าย/คัดลอก และฟังก์ชันเหมือน Word)', () => {
  const html = readFileSync(join(root, 'eoffice-editor.html'), 'utf8');
  const distSingle = readFileSync(join(root, 'dist/eoffice-editor-single.html'), 'utf8');

  test('โครงสร้าง HTML มี Overlay สำหรับย้าย/ปรับขนาด กล่องแสดงการวาดตาราง และเส้นกะระยะลากปรับขนาด', () => {
    assert.ok(html.includes('id="tblOverlay"'), 'ต้องมี #tblOverlay สำหรับจัดการตาราง');
    assert.ok(html.includes('id="tblHandleMove"'), 'ต้องมีปุ่มมือจับย้ายตาราง #tblHandleMove (✥)');
    assert.ok(html.includes('id="tblHandleResize"'), 'ต้องมีมือจับปรับขนาดตาราง #tblHandleResize');
    assert.ok(html.includes('id="tblDrawBox"'), 'ต้องมี #tblDrawBox สำหรับลากเส้นวาดตาราง');
    assert.ok(html.includes('id="tblLineGuide"'), 'ต้องมี #tblLineGuide สำหรับแสดงเส้นบอกตำแหน่งขณะลากปรับเส้นตาราง');
  });

  test('เมนูตาราง (#m-table) มีคำสั่งครบถ้วนตามแบบ Microsoft Word', () => {
    // ปุ่มวาดตาราง
    assert.ok(html.includes('id="btnDrawTable"'), 'ต้องมีปุ่มเปิดโหมดวาดตาราง');
    assert.ok(html.includes('data-act="table-draw"'), 'ต้องมี act table-draw');

    // การเลือก คัดลอก ตัด วาง ย้าย
    assert.ok(html.includes('data-act="table-select-all"'), 'ต้องมี act table-select-all (เลือกทั้งตาราง)');
    assert.ok(html.includes('data-act="table-copy"'), 'ต้องมี act table-copy (คัดลอกเฉพาะตาราง)');
    assert.ok(html.includes('data-act="table-cut"'), 'ต้องมี act table-cut (ตัดเฉพาะตาราง)');
    assert.ok(html.includes('data-act="table-paste"'), 'ต้องมี act table-paste (วางตาราง)');
    assert.ok(html.includes('data-act="table-move-up"'), 'ต้องมี act table-move-up (ย้ายตารางขึ้น)');
    assert.ok(html.includes('data-act="table-move-down"'), 'มี act table-move-down (ย้ายตารางลง)');

    // เส้นขอบตาราง (Borders)
    assert.ok(html.includes('data-act="table-border" data-val="all"'), 'มีตัวเลือกขอบทุกด้าน');
    assert.ok(html.includes('data-act="table-border" data-val="none"'), 'มีตัวเลือกไม่มีเส้นขอบ (สำหรับลงนาม)');
    assert.ok(html.includes('data-act="table-border" data-val="outside"'), 'มีตัวเลือกเส้นขอบนอก');
    assert.ok(html.includes('data-act="table-border" data-val="topbottom"'), 'มีตัวเลือกเส้นบน-ล่าง');

    // สีพื้นหลังเซลล์ และการกระจายคอลัมน์
    assert.ok(html.includes('data-act="table-cell-bg"'), 'มีตัวเลือกสีพื้นหลังเซลล์');
    assert.ok(html.includes('id="tCellColor"'), 'มี color picker สำหรับสีเซลล์แบบกำหนดเอง');
    assert.ok(html.includes('data-act="table-distribute-cols"'), 'มีคำสั่งกระจายคอลัมน์ให้เท่ากัน');
  });

  test('CSS มีสไตล์รองรับการจัดรูปแบบตาราง เส้นขอบ Overlay และเส้นบอกตำแหน่งลากปรับขนาด', () => {
    assert.ok(html.includes('#editor table[data-border="none"] td'), 'CSS ต้องมีกฎสำหรับ data-border="none"');
    assert.ok(html.includes('#editor table[data-border="outside"]'), 'CSS ต้องมีกฎสำหรับ data-border="outside"');
    assert.ok(html.includes('#editor table[data-border="topbottom"]'), 'CSS ต้องมีกฎสำหรับ data-border="topbottom"');
    assert.ok(html.includes('.tbl-overlay'), 'CSS ต้องมี .tbl-overlay');
    assert.ok(html.includes('.tbl-handle-move'), 'CSS ต้องมี .tbl-handle-move');
    assert.ok(html.includes('.tbl-handle-resize'), 'CSS ต้องมี .tbl-handle-resize');
    assert.ok(html.includes('.tbl-draw-box'), 'CSS ต้องมี .tbl-draw-box');
    assert.ok(html.includes('.tbl-line-guide'), 'CSS ต้องมี .tbl-line-guide');
  });

  test('ฟังก์ชัน cleanHtml ส่งออกรูปแบบตาราง เส้นขอบ และสีพื้นหลังอย่างถูกต้อง', () => {
    assert.ok(html.includes('data-border'), 'cleanHtml ต้องตรวจสอบ attribute data-border');
    assert.ok(html.includes('border:none'), 'cleanHtml ต้องรองรับเส้นขอบ border:none');
    assert.ok(html.includes('backgroundColor'), 'cleanHtml ต้องคงค่า backgroundColor ของเซลล์ไว้');
  });

  test('ไฟล์ dist/eoffice-editor-single.html ได้รับการอัปเดตฟังก์ชันตารางครบถ้วน', () => {
    assert.ok(distSingle.includes('id="tblOverlay"'));
    assert.ok(distSingle.includes('id="btnDrawTable"'));
    assert.ok(distSingle.includes('tbl-handle-move'));
    assert.ok(distSingle.includes('tbl-draw-box'));
    assert.ok(distSingle.includes('tblLineGuide'));
    assert.ok(distSingle.includes('resizeTableColumn'));
    assert.ok(distSingle.includes('btnPasteTable'));
  });

  test('ทดสอบการทำงานของตารางในเบราว์เซอร์จริง (Edge Headless)', async (t) => {
    const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
    if (!existsSync(edgePath)) {
      t.skip('ไม่พบ msedge.exe ข้ามการทดสอบเบราว์เซอร์จริง');
      return;
    }

    const port = 9250 + Math.floor(Math.random() * 40);
    const edge = spawn(edgePath, [
      '--headless=new',
      '--allow-file-access-from-files',
      `--remote-debugging-port=${port}`,
      `file:///${join(root, 'dist/eoffice-editor-single.html').replace(/\\/g, '/')}`
    ]);

    try {
      await new Promise(r => setTimeout(r, 1200));
      const tabs = await new Promise((resolve, reject) => {
        http.get(`http://127.0.0.1:${port}/json`, res => {
          let d = ''; res.on('data', c => d += c);
          res.on('end', () => resolve(JSON.parse(d)));
        }).on('error', reject);
      });

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
            if (res.error) {
              reject(new Error(JSON.stringify(res.error)));
            } else if (res.result && res.result.exceptionDetails) {
              reject(new Error(JSON.stringify(res.result.exceptionDetails)));
            } else {
              resolve(res.result?.result || res.result);
            }
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

      // 1. ทดสอบการแทรกตารางผ่าน window.eofficeEditor.insertTable(3, 3)
      const insertTblRes = await send('Runtime.evaluate', {
        expression: `(() => {
          window.eofficeEditor.insertTable(3, 3);
          const tbl = document.querySelector('#editor table');
          return {
            hasTable: !!tbl,
            rows: tbl ? tbl.querySelectorAll('tr').length : 0,
            cols: tbl ? tbl.querySelectorAll('tr:first-child td').length : 0
          };
        })()`,
        returnByValue: true
      });
      assert.equal(insertTblRes.value.hasTable, true, 'ต้องมีตารางแทรกใน Quill');
      assert.equal(insertTblRes.value.rows, 3, 'ต้องมี 3 แถว');
      assert.equal(insertTblRes.value.cols, 3, 'ต้องมี 3 คอลัมน์');

      // 2. ทดสอบการเลือกทั้งตาราง (selectEntireTable)
      const selectRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const firstCell = document.querySelector('#editor td');
          firstCell.focus();
          window.eofficeEditor.selectEntireTable();
          const range = window.eofficeEditor.quill.getSelection();
          return { hasSelection: !!range && range.length > 0 };
        })()`,
        returnByValue: true
      });
      assert.equal(selectRes.value.hasSelection, true, 'ต้องเลือกช่วงทั้งตารางได้');

      // 3. ทดสอบการกำหนดเส้นขอบเป็นแบบ "ไม่มีเส้นขอบ" (none) สำหรับบล็อกลงนาม
      const borderRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const firstCell = document.querySelector('#editor td');
          firstCell.focus();
          window.eofficeEditor.setTableBorder('none');
          const tbl = document.querySelector('#editor table');
          return {
            borderAttr: tbl.getAttribute('data-border'),
            firstCellBorder: window.getComputedStyle(firstCell).borderTopStyle
          };
        })()`,
        returnByValue: true
      });
      assert.equal(borderRes.value.borderAttr, 'none', 'ตารางต้องมี data-border="none"');
      assert.equal(borderRes.value.firstCellBorder, 'dashed', 'ในโหมดแก้ไขต้องแสดงเส้นประจางช่วยพิมพ์ (dashed gridline เหมือน Word)');

      // 4. ทดสอบการกำหนดสีพื้นหลังเซลล์
      const bgRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const firstCell = document.querySelector('#editor td');
          firstCell.focus();
          window.eofficeEditor.setTableCellBg('#e8f0fe');
          return {
            cellBg: firstCell.style.backgroundColor
          };
        })()`,
        returnByValue: true
      });
      assert.ok(bgRes.value.cellBg.includes('232, 240, 254') || bgRes.value.cellBg.includes('#e8f0fe'), 'เซลล์ต้องได้รับสีพื้นหลังที่กำหนด');

      // 5. ทดสอบการกระจายคอลัมน์ให้เท่ากัน (distributeColumns)
      const distColRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const firstCell = document.querySelector('#editor td');
          firstCell.focus();
          window.eofficeEditor.distributeColumns();
          const cells = Array.from(document.querySelectorAll('#editor tr:first-child td'));
          return cells.map(c => c.style.width);
        })()`,
        returnByValue: true
      });
      const widths = distColRes.value;
      assert.equal(widths.length, 3);
      assert.equal(widths[0], widths[1]);
      assert.equal(widths[1], widths[2]);

      // 6. ทดสอบ cleanHtml ว่ารักษา data-border="none" และ background color ไว้
      const cleanRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const html = window.eofficeEditor.cleanHtml();
          return {
            hasTable: html.includes('<table'),
            hasBorderNone: /border(?:-style)?\\s*:\\s*none/i.test(html) || html.includes('border: 0'),
            hasBg: html.includes('background'),
            htmlSnippet: html
          };
        })()`,
        returnByValue: true
      });
      assert.equal(cleanRes.value.hasTable, true);
      assert.equal(cleanRes.value.hasBorderNone, true, 'cleanRes HTML: ' + cleanRes.value.htmlSnippet);
      assert.equal(cleanRes.value.hasBg, true);

      // 7. ทดสอบโหมดวาดตาราง (startDrawTableMode / stopDrawTableMode)
      const drawModeRes = await send('Runtime.evaluate', {
        expression: `(() => {
          window.eofficeEditor.startDrawTableMode();
          const isDrawingActive = document.body.classList.contains('table-drawing-mode');
          window.eofficeEditor.stopDrawTableMode();
          const isDrawingStopped = !document.body.classList.contains('table-drawing-mode');
          return { isDrawingActive, isDrawingStopped };
        })()`,
        returnByValue: true
      });
      assert.equal(drawModeRes.value.isDrawingActive, true);
      assert.equal(drawModeRes.value.isDrawingStopped, true);

      // 8. ทดสอบการคัดลอกและวางตาราง (copyTableOnly / pasteTableOnly)
      const pasteRes = await send('Runtime.evaluate', {
        expression: `(async () => {
          const firstCell = document.querySelector('#editor td');
          firstCell.focus();
          await window.eofficeEditor.copyTableOnly();
          // ลบตารางออก
          window.eofficeEditor.deleteTableOnly();
          const countAfterDelete = document.querySelectorAll('#editor table').length;
          // สั่งวางตารางกลับคืน
          await window.eofficeEditor.pasteTableOnly();
          await new Promise(r => setTimeout(r, 100));
          const tblPasted = document.querySelector('#editor table');
          return {
            countAfterDelete,
            hasPastedTable: !!tblPasted,
            pastedRows: tblPasted ? tblPasted.querySelectorAll('tr').length : 0,
            pastedCols: tblPasted ? tblPasted.querySelectorAll('tr:first-child td').length : 0
          };
        })()`,
        awaitPromise: true,
        returnByValue: true
      });
      assert.equal(pasteRes.value.countAfterDelete, 0, 'ตารางต้องถูกลบออกก่อนวาง');
      assert.equal(pasteRes.value.hasPastedTable, true, 'ต้องมีตารางที่ถูกวางใหม่');
      assert.equal(pasteRes.value.pastedRows, 3, 'ตารางที่วางต้องมี 3 แถว');
      assert.equal(pasteRes.value.pastedCols, 3, 'ตารางที่วางต้องมี 3 คอลัมน์');

      // 9. ทดสอบการลากปรับเส้นตารางอิสระ (resizeTableColumn, resizeTableRow และ border detection)
      const lineDragRes = await send('Runtime.evaluate', {
        expression: `(() => {
          const tbl = document.querySelector('#editor table');
          const row0 = tbl.querySelectorAll('tr')[0];
          const c0 = row0.children[0];
          const c1 = row0.children[1];

          const beforeW0 = c0.offsetWidth;
          const beforeW1 = c1.offsetWidth;
          const beforeH0 = row0.offsetHeight;

          // ปรับความกว้างคอลัมน์ภายใน (+40px ให้คอลัมน์ 0, คอลัมน์ 1 ลดลง 40px)
          window.eofficeEditor.resizeTableColumn(tbl, 0, 40);
          const afterW0 = c0.offsetWidth;
          const afterW1 = c1.offsetWidth;

          // ปรับความสูงแถว (+25px ให้แถว 0)
          window.eofficeEditor.resizeTableRow(tbl, 0, 25);
          const afterH0 = row0.offsetHeight;

          // ตรวจจับเส้นตารางเมื่อเลื่อนเมาส์ชี้ใกล้ขอบ (border hover detection)
          const c0Rect = c0.getBoundingClientRect();
          const evtCol = new PointerEvent('pointermove', {
            clientX: c0Rect.right - 2,
            clientY: c0Rect.top + 10,
            bubbles: true
          });
          c0.dispatchEvent(evtCol);
          const cursorOnColBorder = document.body.style.cursor;

          const evtRow = new PointerEvent('pointermove', {
            clientX: c0Rect.left + 10,
            clientY: c0Rect.bottom - 2,
            bubbles: true
          });
          c0.dispatchEvent(evtRow);
          const cursorOnRowBorder = document.body.style.cursor;

          // เลื่อนเมาส์ออกจากขอบ
          const evtAway = new PointerEvent('pointermove', {
            clientX: c0Rect.left + 25,
            clientY: c0Rect.top + 15,
            bubbles: true
          });
          c0.dispatchEvent(evtAway);
          const cursorAway = document.body.style.cursor;

          return {
            diffW0: afterW0 - beforeW0,
            diffW1: afterW1 - beforeW1,
            diffH0: afterH0 - beforeH0,
            cursorOnColBorder,
            cursorOnRowBorder,
            cursorAway
          };
        })()`,
        returnByValue: true
      });

      assert.ok(Math.abs(lineDragRes.value.diffW0 - 40) <= 2, `ความกว้างคอลัมน์ 0 ต้องเพิ่มขึ้น 40px (ได้ diff: ${lineDragRes.value.diffW0})`);
      assert.ok(Math.abs(lineDragRes.value.diffW1 - (-40)) <= 2, `ความกว้างคอลัมน์ 1 ต้องลดลง 40px (ได้ diff: ${lineDragRes.value.diffW1})`);
      assert.ok(Math.abs(lineDragRes.value.diffH0 - 25) <= 2, `ความสูงแถว 0 ต้องเพิ่มขึ้น 25px (ได้ diff: ${lineDragRes.value.diffH0})`);
      assert.equal(lineDragRes.value.cursorOnColBorder, 'col-resize', 'เมื่อชี้ใกล้เส้นคอลัมน์ เคอร์เซอร์ต้องเป็น col-resize');
      assert.equal(lineDragRes.value.cursorOnRowBorder, 'row-resize', 'เมื่อชี้ใกล้เส้นแถว เคอร์เซอร์ต้องเป็น row-resize');
      assert.equal(lineDragRes.value.cursorAway, '', 'เมื่อออกจากเส้นตาราง เคอร์เซอร์ต้องกลับเป็นค่าปกติ');

      ws.close();
    } finally {
      edge.kill('SIGKILL');
    }
  });
});
