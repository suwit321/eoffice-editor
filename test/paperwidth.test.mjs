import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

describe('Paper Width & Edge Dragging (การขยายหน้ากระดาษด้วยการลากขอบ)', () => {
  const html = readFileSync(join(root, 'eoffice-editor.html'), 'utf8');
  const tools = readFileSync(join(root, 'eoffice-tools.js'), 'utf8');
  const distSingle = readFileSync(join(root, 'dist/eoffice-editor-single.html'), 'utf8');

  test('มี Handle สำหรับลากขยายขอบกระดาษทั้งบนไม้บรรทัดและขอบกระดาษ', () => {
    assert.ok(html.includes('id="rulerEdgeHandle"'), 'ต้องมี #rulerEdgeHandle บนไม้บรรทัด');
    assert.ok(html.includes('id="paperEdgeHandle"'), 'ต้องมี #paperEdgeHandle ที่ขอบกระดาษ');
    assert.ok(html.includes('.ruler-edge-handle'), 'ต้องมี class .ruler-edge-handle ใน CSS');
    assert.ok(html.includes('.paper-edge-handle'), 'ต้องมี class .paper-edge-handle ใน CSS');
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
});
