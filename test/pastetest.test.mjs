import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

function loadPasteTest() {
  const code = readFileSync(join(root, 'eoffice-pastetest.js'), 'utf8');
  const mockWindow = {};
  const fn = new Function('window', code);
  fn(mockWindow);
  return mockWindow.eofficePasteTest;
}

describe('eoffice-pastetest.js', () => {
  const pt = loadPasteTest();

  test('ส่งออก API ครบถ้วน', () => {
    assert.ok(pt, 'eofficePasteTest ต้องถูก export');
    assert.ok(Array.isArray(pt.PROBES), 'PROBES ต้องเป็น array');
    assert.equal(typeof pt.buildHtml, 'function');
    assert.equal(typeof pt.buildText, 'function');
    assert.equal(typeof pt.analyze, 'function');
    assert.equal(typeof pt.simulate, 'function');
    assert.equal(typeof pt.profileOf, 'function');
    assert.equal(typeof pt.adviceFor, 'function');
  });

  test('มี probe ครบ 25 รายการ (T01 ถึง T25) และรหัสไม่ซ้ำกัน', () => {
    assert.equal(pt.PROBES.length, 25, 'ต้องมี 25 probes');
    const ids = pt.PROBES.map(p => p.id);
    const uniqueIds = new Set(ids);
    assert.equal(uniqueIds.size, 25, 'รหัส probe ต้องไม่ซ้ำกัน');
    
    for (let i = 1; i <= 25; i++) {
      const id = `T${String(i).padStart(2, '0')}`;
      assert.ok(uniqueIds.has(id), `ต้องมี probe ${id}`);
    }
  });

  test('ทุก probe มี label, short และ check function', () => {
    for (const p of pt.PROBES) {
      assert.ok(p.label && typeof p.label === 'string', `${p.id} ต้องมี label`);
      assert.ok(p.short && typeof p.short === 'string', `${p.id} ต้องมี short label`);
      assert.equal(typeof p.check, 'function', `${p.id} ต้องมีฟังก์ชัน check`);
    }
  });

  test('profileOf คำนวณ stripWJ และ tabsAsSpaces ถูกต้อง', () => {
    // กรณีทุกอย่าง ok
    const resOk = [
      { id: 'T15', state: 'ok' },
      { id: 'T17', state: 'ok' },
    ];
    const profOk = pt.profileOf(resOk);
    assert.equal(profOk.stripWJ, false);
    assert.equal(profOk.tabsAsSpaces, false);

    // กรณี T15 (Word Joiner) หาย -> stripWJ ต้องเป็น true
    const resNoWJ = [
      { id: 'T15', state: 'lost' },
      { id: 'T17', state: 'ok' },
    ];
    const profNoWJ = pt.profileOf(resNoWJ);
    assert.equal(profNoWJ.stripWJ, true);
    assert.equal(profNoWJ.tabsAsSpaces, false);

    // กรณี T17 (Tab) หาย -> tabsAsSpaces ต้องเป็น true
    const resNoTab = [
      { id: 'T15', state: 'ok' },
      { id: 'T17', state: 'changed' },
    ];
    const profNoTab = pt.profileOf(resNoTab);
    assert.equal(profNoTab.stripWJ, false);
    assert.equal(profNoTab.tabsAsSpaces, true);
  });

  test('adviceFor เตือนเฉพาะฟีเจอร์ที่มีปัญหาและเอกสารมีการใช้งานจริง', () => {
    const results = [
      { id: 'T01', state: 'lost', short: 'ย่อหน้าบรรทัดแรก' },
      { id: 'T02', state: 'ok', short: 'กั้นซ้าย' },
      { id: 'T15', state: 'lost', short: 'Word Joiner' }, // wj ไม่ต้องเตือนเพราะมี fallback
      { id: 'T17', state: 'lost', short: 'แท็บ' },        // tab ไม่ต้องเตือนเพราะมี fallback
    ];
    const html = '<p style="text-indent:2.5cm">ข้อความ</p>';
    const advice = pt.adviceFor(results, html);
    assert.ok(advice.includes('ย่อหน้าบรรทัดแรก'), 'ต้องเตือนเรื่องย่อหน้าบรรทัดแรก');
    assert.ok(!advice.includes('Word Joiner'), 'ต้องไม่เตือน Word Joiner');
    assert.ok(!advice.includes('แท็บ'), 'ต้องไม่เตือน แท็บ');
  });
});
