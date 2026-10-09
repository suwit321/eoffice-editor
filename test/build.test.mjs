import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { execSync } from 'node:child_process';
import { createHash } from 'node:crypto';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

describe('build-single.mjs and dist artifacts', () => {
  test('รัน node build-single.mjs สำเร็จโดยไม่มี error', () => {
    const stdout = execSync('node build-single.mjs', { cwd: root, encoding: 'utf8' });
    assert.ok(stdout.includes('dist/index.html'), 'stdout ต้องแสดง dist/index.html');
    assert.ok(stdout.includes('dist/eoffice-editor-single.html'), 'stdout ต้องแสดง dist/eoffice-editor-single.html');
  });

  test('ไฟล์ผลลัพธ์ใน dist ถูกสร้างครบถ้วน', () => {
    const requiredFiles = [
      'dist/index.html',
      'dist/eoffice-editor-single.html',
      'dist/_headers',
      'dist/robots.txt',
      'dist/THIRD-PARTY-NOTICES.md',
    ];
    for (const f of requiredFiles) {
      assert.ok(existsSync(join(root, f)), `ไฟล์ ${f} ต้องมีอยู่จริง`);
    }
  });

  test('dist/index.html มี CSP ตรงกับแฮชของสคริปต์ inline ทั้ง 4 ก้อน', () => {
    const html = readFileSync(join(root, 'dist/index.html'), 'utf8');
    assert.ok(html.includes('<meta http-equiv="Content-Security-Policy"'), 'ต้องมี meta CSP');

    const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)];
    assert.equal(scripts.length, 4, 'ต้องมี script inline 4 ก้อน');

    for (const match of scripts) {
      const hash = `sha256-${createHash('sha256').update(match[1], 'utf8').digest('base64')}`;
      assert.ok(html.includes(hash), `CSP ต้องมีแฮช '${hash}' ของสคริปต์`);
    }
  });

  test('dist/eoffice-editor-single.html ฝัง SDK ในตัวและไม่มี CSP meta', () => {
    const single = readFileSync(join(root, 'dist/eoffice-editor-single.html'), 'utf8');
    assert.ok(single.includes('id="sdk-embedded"'), 'ต้องมีแท็ก id="sdk-embedded"');
    assert.ok(!single.includes('http-equiv="Content-Security-Policy"'), 'ไฟล์เดี่ยวสำหรับเปิด file:// ต้องไม่มี CSP meta ขวาง');
  });

  test('dist/_headers มี Header ความปลอดภัยที่ถูกต้อง', () => {
    const headers = readFileSync(join(root, 'dist/_headers'), 'utf8');
    assert.ok(headers.includes('Content-Security-Policy:'), 'ต้องมี Content-Security-Policy');
    assert.ok(headers.includes('X-Content-Type-Options: nosniff'), 'ต้องมี nosniff');
    assert.ok(headers.includes('X-Frame-Options: DENY'), 'ต้องมี DENY');
    assert.ok(headers.includes('Referrer-Policy: no-referrer'), 'ต้องมี no-referrer');
  });

  test('dist/robots.txt ป้องกัน Web Crawler ทำดัชนี', () => {
    const robots = readFileSync(join(root, 'dist/robots.txt'), 'utf8');
    assert.ok(robots.includes('Disallow: /'), 'ต้องมี Disallow: /');
  });
});
