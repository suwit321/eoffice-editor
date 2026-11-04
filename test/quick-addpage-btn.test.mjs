import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

describe('ปุ่มเพิ่มหน้าใหม่ใกล้ปุ่มบันทึก (Quick Add Page Button in .topcenter)', () => {
  const html = readFileSync(join(root, 'eoffice-editor.html'), 'utf8');
  const distSingle = readFileSync(join(root, 'dist/eoffice-editor-single.html'), 'utf8');
  const distIndex = readFileSync(join(root, 'dist/index.html'), 'utf8');

  test('มีปุ่ม #btnQuickAddPage อยู่ใน .topcenter ในแนวเดียวกับปุ่มบันทึก', () => {
    // ต้องอยู่ในบล็อก .topcenter เดียวกับ btnQuickSave
    const topcenterMatch = html.match(/<div class="topcenter"[^>]*>([\s\S]*?)<\/div>/);
    assert.ok(topcenterMatch, 'ต้องมีคอนเทนเนอร์ .topcenter');
    const content = topcenterMatch[1];

    assert.ok(content.includes('id="btnQuickSave"'), 'ต้องมีปุ่มบันทึก #btnQuickSave');
    assert.ok(content.includes('id="btnQuickAddPage"'), 'ต้องมีปุ่มเพิ่มหน้าใหม่ #btnQuickAddPage ใน .topcenter');
    
    // ตรวจสอบลำดับ: btnQuickSave ต้องมาก่อน btnQuickAddPage
    const saveIdx = content.indexOf('id="btnQuickSave"');
    const addPageIdx = content.indexOf('id="btnQuickAddPage"');
    assert.ok(addPageIdx > saveIdx, 'ปุ่มเพิ่มหน้าใหม่ต้องอยู่ถัดจากปุ่มบันทึก');
  });

  test('ปุ่ม #btnQuickAddPage มีการเว้นระยะห่างที่พอดี (margin-left) ไม่ชิดปุ่มบันทึกเกินไป', () => {
    assert.ok(html.includes('.tc-addpage{margin-left:'), 'CSS ต้องมี margin-left สำหรับ .tc-addpage');
    assert.ok(html.includes('class="tb tc-addpage"') && html.includes('id="btnQuickAddPage"'), 'ปุ่มต้องมีคลาส tb, tc-addpage และ id btnQuickAddPage');
  });

  test('มี SVG icon #i-addpage สำหรับปุ่มเพิ่มหน้าใหม่', () => {
    assert.ok(html.includes('id="i-addpage"'), 'ต้องมี symbol id="i-addpage"');
    assert.ok(html.includes('href="#i-addpage"'), 'ปุ่มต้องเรียกใช้ href="#i-addpage"');
  });

  test('มี Event Listener เชื่อมโยงกับฟังก์ชัน addPage()', () => {
    assert.ok(
      html.includes("$('#btnQuickAddPage')?.addEventListener('click', () => addPage())"),
      'ต้องมี event listener เรียก addPage เมื่อคลิกปุ่ม #btnQuickAddPage'
    );
  });

  test('ไฟล์ dist ทั้งสองไฟล์ (index.html และ eoffice-editor-single.html) มีปุ่มและสไตล์ครบถ้วน', () => {
    for (const [name, content] of [['dist/index.html', distIndex], ['dist/eoffice-editor-single.html', distSingle]]) {
      assert.ok(content.includes('id="btnQuickAddPage"'), `${name} ต้องมี #btnQuickAddPage`);
      assert.ok(content.includes('.tc-addpage{margin-left:'), `${name} ต้องมี CSS .tc-addpage`);
      assert.ok(content.includes('id="i-addpage"'), `${name} ต้องมีไอคอน i-addpage`);
    }
  });
});
