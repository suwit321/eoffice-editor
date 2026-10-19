import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, writeFileSync } from 'node:fs';
import { resolve, join } from 'node:path';
import { spawn } from 'node:child_process';
import http from 'node:http';
import { tmpdir } from 'node:os';

const root = resolve('.');

describe('Tools Panel Organization & Color-coded Sub-tabs (การจัดกลุ่มและสีแท็บย่อยในแผงเครื่องมือ)', () => {
  const toolsCode = readFileSync(join(root, 'eoffice-tools.js'), 'utf8');
  const distSingle = readFileSync(join(root, 'dist/eoffice-editor-single.html'), 'utf8');

  test('มีโครงสร้างหมวดหมู่หลัก 4 กลุ่ม (.tools-groups, .tgroup-btn) ในแผงเครื่องมือ', () => {
    assert.ok(toolsCode.includes('class="tools-groups"'), 'ต้องมี .tools-groups');
    assert.ok(toolsCode.includes('data-group="text"'), 'ต้องมีกลุ่มข้อความ');
    assert.ok(toolsCode.includes('data-group="format"'), 'ต้องมีกลุ่มจัดหน้า');
    assert.ok(toolsCode.includes('data-group="audit"'), 'ต้องมีกลุ่มตรวจ & AI');
    assert.ok(toolsCode.includes('data-group="file"'), 'ต้องมีกลุ่มไฟล์ & เสริม');
  });

  test('แท็บย่อยมีสีประจำแท็บที่แตกต่างกัน (Distinct Color Palette for each sub-tab)', () => {
    const tabs = ['num', 'find', 'phrase', 'nobreak', 'layout', 'style', 'theme', 'auto', 'check', 'ai', 'file', 'cloud', 'paste'];
    for (const t of tabs) {
      assert.ok(toolsCode.includes(`.tsub-btn[data-tab="${t}"]`), `ต้องมีการกำหนดสไตล์เฉพาะสำหรับแท็บ ${t}`);
      assert.ok(toolsCode.includes(`--tab-color`), `ต้องมีการใช้ตัวแปร --tab-color ในแท็บย่อย`);
    }
  });

  test('มีส่วนพับเก็บย่อส่วน (.tools-fold) สำหรับเนื้อหาที่ไม่จำเป็นต้องแสดงตลอดเวลา', () => {
    assert.ok(toolsCode.includes('class="tools-fold"'), 'ต้องมี class .tools-fold');
    assert.ok(toolsCode.includes('id="aiConfigFold"'), 'ต้องมี fold สำหรับตั้งค่า AI');
    assert.ok(toolsCode.includes('ตราครุฑสำหรับแม่แบบ'), 'ต้องมีย่อส่วนตราครุฑ');
    assert.ok(toolsCode.includes('สัญลักษณ์พิเศษ'), 'ต้องมีย่อส่วนสัญลักษณ์พิเศษ');
    assert.ok(toolsCode.includes('โทนสีกระดาษและถนอมสายตา'), 'ต้องมีย่อส่วนโทนสีกระดาษ');
  });

  test('ไฟล์ผลลัพธ์ dist/eoffice-editor-single.html มีการอัปเดตโครงสร้างกลุ่มและแท็บสีครบถ้วน', () => {
    assert.ok(distSingle.includes('tools-groups'), 'dist ต้องมี tools-groups');
    assert.ok(distSingle.includes('tsub-btn'), 'dist ต้องมี tsub-btn');
    assert.ok(distSingle.includes('tools-fold'), 'dist ต้องมี tools-fold');
    assert.ok(distSingle.includes('data-active-group'), 'dist ต้องมี data-active-group');
  });

  test('ทดสอบในเบราว์เซอร์จริง (Edge Headless): การเปลี่ยนหมวดหมู่หลักและการเลือกแท็บย่อยสีต่างๆ', async () => {
    const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
    const port = 9345;
    const userData = join(tmpdir(), 'edge-tools-nav-' + Date.now());
    const fileUrl = 'file:///' + resolve(root, 'dist/eoffice-editor-single.html').replace(/\\/g, '/');

    const edge = spawn(edgePath, [
      '--headless=new',
      `--remote-debugging-port=${port}`,
      `--user-data-dir=${userData}`,
      '--no-first-run',
      '--no-default-browser-check',
      '--window-size=1280,900',
      fileUrl
    ]);

    try {
      let tabs = [];
      const t0 = Date.now();
      while (Date.now() - t0 < 6000) {
        try {
          tabs = await new Promise((res, rej) => {
            const req = http.get(`http://127.0.0.1:${port}/json`, r => {
              let d = ''; r.on('data', c => d += c);
              r.on('end', () => { try { res(JSON.parse(d)); } catch { res([]); } });
            });
            req.on('error', rej);
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

      // รอให้หน้าพร้อม
      await send('Runtime.evaluate', {
        expression: `new Promise(resolve => {
          if (window.eofficeTools) return resolve();
          document.addEventListener('eoffice:ready', () => resolve(), { once: true });
          setTimeout(resolve, 2000);
        })`,
        awaitPromise: true
      });

      // เปิดแผงเครื่องมือ
      const openResult = await send('Runtime.evaluate', {
        expression: `(() => {
          const btn = document.getElementById('btnTools');
          if (btn) btn.click();
          const panel = document.getElementById('toolsPanel');
          return {
            panelOpen: panel && !panel.hidden,
            groupsCount: document.querySelectorAll('.tgroup-btn').length,
            subtabsCount: document.querySelectorAll('.tsub-btn').length,
            activeGroup: document.querySelector('.tools-nav')?.dataset.activeGroup,
            activeTab: document.querySelector('.tsub-btn[aria-selected="true"]')?.dataset.tab
          };
        })()`,
        returnByValue: true
      });

      assert.ok(openResult.value.panelOpen, 'แผงเครื่องมือต้องเปิดขึ้นมา');
      assert.equal(openResult.value.groupsCount, 4, 'ต้องมี 4 หมวดหมู่หลัก');
      assert.equal(openResult.value.subtabsCount, 13, 'ต้องมี 13 แท็บย่อยทั้งหมด (รวมแท็บ Cloud ☁️)');
      assert.equal(openResult.value.activeGroup, 'text', 'หมวดเริ่มต้นต้องเป็น text');

      // คลิกเปลี่ยนกลุ่มไปที่ "จัดหน้า" (format)
      const switchGroup = await send('Runtime.evaluate', {
        expression: `(() => {
          const formatBtn = document.querySelector('.tgroup-btn[data-group="format"]');
          if (formatBtn) formatBtn.click();
          return {
            activeGroup: document.querySelector('.tools-nav')?.dataset.activeGroup,
            activeTab: document.querySelector('.tsub-btn[aria-selected="true"]')?.dataset.tab,
            layoutPanelVisible: !document.querySelector('section[data-panel="layout"]')?.hidden
          };
        })()`,
        returnByValue: true
      });

      assert.equal(switchGroup.value.activeGroup, 'format', 'ต้องสลับกลุ่มเป็น format');
      assert.ok(switchGroup.value.layoutPanelVisible, 'แท็บย่อยแรกของ format (จัดหน้า) ต้องเปิดขึ้นมา');

      // คลิกไปที่แท็บย่อย "สไตล์" (style)
      const switchSubtab = await send('Runtime.evaluate', {
        expression: `(() => {
          const styleTab = document.querySelector('.tsub-btn[data-tab="style"]');
          if (styleTab) styleTab.click();
          const styleBtnColor = getComputedStyle(styleTab).getPropertyValue('--tab-color').trim();
          return {
            activeTab: document.querySelector('.tsub-btn[aria-selected="true"]')?.dataset.tab,
            stylePanelVisible: !document.querySelector('section[data-panel="style"]')?.hidden,
            tabColor: styleBtnColor
          };
        })()`,
        returnByValue: true
      });

      assert.equal(switchSubtab.value.activeTab, 'style', 'ต้องเลือกแท็บย่อย style');
      assert.ok(switchSubtab.value.stylePanelVisible, 'หน้าจอของแท็บ style ต้องแสดงผล');
      assert.ok(switchSubtab.value.tabColor.length > 0, 'แท็บย่อยต้องมีตัวแปรสีประจำตัว');

      // ถ่ายภาพ Screenshot บันทึกไว้เป็นหลักฐานยืนยัน
      const shot = await send('Page.captureScreenshot', { format: 'png' });
      if (shot?.data) {
        const artifactDir = 'C:\\Users\\suwit\\.gemini\\antigravity\\brain\\3e316c91-ec18-4b0a-9a1e-3224837de510';
        writeFileSync(join(artifactDir, 'tools-grouped-navigation.png'), Buffer.from(shot.data, 'base64'));
      }

      ws.close();
    } finally {
      edge.kill('SIGKILL');
    }
  });
});
