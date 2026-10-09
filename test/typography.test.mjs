import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

function extractToolsHelpers() {
  const code = readFileSync(join(root, 'eoffice-tools.js'), 'utf8');
  // ดึงฟังก์ชันตัวเลข, บาท, คำสะกดผิด
  const numChunk = code.match(/\/\/ ================= ตัวเลข \/ เงิน \/ วันที่ =================([\s\S]*?)\/\/ ================= ตรวจคำ/)[1];
  const misspellChunk = code.match(/const MISSPELL = \{([\s\S]*?)\};/)[0];

  const fn = new Function(`
    ${numChunk}
    ${misspellChunk}
    return { toThai, toArabic, bahtText, readInt, MISSPELL };
  `);
  return fn();
}

function extractEditorHelpers() {
  const code = readFileSync(join(root, 'eoffice-editor.html'), 'utf8');
  const toCmMatch = code.match(/function toCm\(v\) \{([\s\S]*?)\n  \}/)[0];
  const fn = new Function(`
    const DEFAULTS = { size: '16pt' };
    const CM = 96 / 2.54;
    ${toCmMatch}
    return { toCm };
  `);
  return fn();
}

describe('Typography & Number Helpers', () => {
  const { toThai, toArabic, bahtText, readInt, MISSPELL } = extractToolsHelpers();
  const { toCm } = extractEditorHelpers();

  describe('toThai & toArabic', () => {
    test('แปลงเลขอารบิกเป็นเลขไทยอย่างถูกต้อง', () => {
      assert.equal(toThai('0123456789'), '๐๑๒๓๔๕๖๗๘๙');
      assert.equal(toThai('วันที่ 25 มกราคม 2569'), 'วันที่ ๒๕ มกราคม ๒๕๖๙');
    });

    test('แปลงเลขไทยเป็นเลขอารบิกอย่างถูกต้อง', () => {
      assert.equal(toArabic('๐๑๒๓๔๕๖๗๘๙'), '0123456789');
      assert.equal(toArabic('ระเบียบข้อที่ ๑๒ วรรค ๓'), 'ระเบียบข้อที่ 12 วรรค 3');
    });
  });

  describe('bahtText', () => {
    test('แปลงจำนวนเต็มหลักต่าง ๆ', () => {
      assert.equal(bahtText('0'), 'ศูนย์บาทถ้วน');
      assert.equal(bahtText('1'), 'หนึ่งบาทถ้วน');
      assert.equal(bahtText('11'), 'สิบเอ็ดบาทถ้วน');
      assert.equal(bahtText('21'), 'ยี่สิบเอ็ดบาทถ้วน');
      assert.equal(bahtText('100'), 'หนึ่งร้อยบาทถ้วน');
      assert.equal(bahtText('101'), 'หนึ่งร้อยเอ็ดบาทถ้วน');
      assert.equal(bahtText('1000'), 'หนึ่งพันบาทถ้วน');
      assert.equal(bahtText('10000'), 'หนึ่งหมื่นบาทถ้วน');
      assert.equal(bahtText('100000'), 'หนึ่งแสนบาทถ้วน');
      assert.equal(bahtText('1000000'), 'หนึ่งล้านบาทถ้วน');
      assert.equal(bahtText('10000000'), 'สิบล้านบาทถ้วน');
    });

    test('แปลงจำนวนเงินที่มีเศษสตางค์และเครื่องหมายจุลภาค', () => {
      assert.equal(bahtText('1,250.50'), 'หนึ่งพันสองร้อยห้าสิบบาทห้าสิบสตางค์');
      assert.equal(bahtText('12,500.25'), 'หนึ่งหมื่นสองพันห้าร้อยบาทยี่สิบห้าสตางค์');
      assert.equal(bahtText('0.75'), 'เจ็ดสิบห้าสตางค์');
      assert.equal(bahtText('0.05'), 'ห้าสตางค์');
    });

    test('รองรับเลขอารบิกและเลขไทยปนกันหรือมีสัญลักษณ์ ฿ / บาท', () => {
      assert.equal(bahtText('๑,๒๕๐.๕๐ บาท'), 'หนึ่งพันสองร้อยห้าสิบบาทห้าสิบสตางค์');
      assert.equal(bahtText('฿ 500'), 'ห้าร้อยบาทถ้วน');
    });

    test('จัดการค่าติดลบและค่าที่ไม่ถูกต้อง', () => {
      assert.equal(bahtText('-500'), 'ลบห้าร้อยบาทถ้วน');
      assert.equal(bahtText('abc'), null);
      assert.equal(bahtText('12.34.56'), null);
    });
  });

  describe('toCm (Editor unit converter)', () => {
    test('แปลงหน่วย cm, mm, in, pt, px ถูกต้อง', () => {
      assert.equal(toCm('2.5cm'), 2.5);
      assert.equal(toCm('25mm'), 2.5);
      assert.equal(toCm('1in'), 2.54);
      assert.ok(Math.abs(toCm('72pt') - 2.54) < 0.001);
      assert.equal(toCm(null), null);
      assert.equal(toCm('invalid'), null);
    });
  });

  describe('MISSPELL dictionary integrity', () => {
    test('ตารางคำผิดไม่ว่างและไม่มีการแมปเข้าหาตัวเองหรือค่าว่าง', () => {
      const keys = Object.keys(MISSPELL);
      assert.ok(keys.length > 40, 'ต้องมีคำสะกดผิดอย่างน้อย 40 คำ');
      for (const [wrong, right] of Object.entries(MISSPELL)) {
        assert.ok(wrong && wrong.trim(), 'คำผิดต้องไม่ว่าง');
        assert.ok(right && right.trim(), 'คำถูกต้องไม่ว่าง');
        assert.notEqual(wrong, right, `คำผิดกับคำถูกต้องไม่เหมือนกัน (${wrong})`);
      }
    });

    test('คำสำคัญในงานราชการสะกดถูกต้อง', () => {
      assert.equal(MISSPELL['อนุญาติ'], 'อนุญาต');
      assert.equal(MISSPELL['กฏหมาย'], 'กฎหมาย');
      assert.equal(MISSPELL['ปรากฎ'], 'ปรากฏ');
      assert.equal(MISSPELL['งบประมาน'], 'งบประมาณ');
      assert.equal(MISSPELL['ประสิทธิภาพ'], undefined); // คำที่ถูกอยู่แล้วต้องไม่มีในคีย์
    });
  });
});
