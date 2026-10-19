import { test, describe } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { resolve, join } from 'node:path';
import { spawn } from 'node:child_process';
import http from 'node:http';
import { tmpdir } from 'node:os';

const root = resolve('.');

describe('Google Drive Cloud Storage Integration (eoffice-cloud.js)', () => {
  const cloudCode = readFileSync(join(root, 'eoffice-cloud.js'), 'utf8');
  const toolsCode = readFileSync(join(root, 'eoffice-tools.js'), 'utf8');
  const editorHtml = readFileSync(join(root, 'eoffice-editor.html'), 'utf8');
  const distSingle = readFileSync(join(root, 'dist/eoffice-editor-single.html'), 'utf8');
  const buildMjs = readFileSync(join(root, 'build-single.mjs'), 'utf8');

  test('ไฟล์ eoffice-cloud.js มีฟังก์ชัน Google Drive และโครงสร้างโฟลเดอร์เฉพาะครบถ้วน', () => {
    assert.ok(cloudCode.includes('window.eofficeCloud'), 'ต้องส่งออก window.eofficeCloud');
    assert.ok(cloudCode.includes('E-Office Documents'), 'ต้องมีชื่อโฟลเดอร์เฉพาะ E-Office Documents');
    assert.ok(cloudCode.includes('ensureFolder'), 'ต้องมีฟังก์ชัน ensureFolder');
    assert.ok(cloudCode.includes('getFolderInfo'), 'ต้องมีฟังก์ชัน getFolderInfo');
    assert.ok(cloudCode.includes('init,'), 'ต้องมีฟังก์ชัน init');
    assert.ok(cloudCode.includes('signIn,'), 'ต้องมีฟังก์ชัน signIn');
    assert.ok(cloudCode.includes('signOut,'), 'ต้องมีฟังก์ชัน signOut');
    assert.ok(cloudCode.includes('saveToCloud,'), 'ต้องมีฟังก์ชัน saveToCloud');
    assert.ok(cloudCode.includes('loadFromCloud,'), 'ต้องมีฟังก์ชัน loadFromCloud');
    assert.ok(cloudCode.includes('listMyDocs,'), 'ต้องมีฟังก์ชัน listMyDocs');
    assert.ok(cloudCode.includes('deleteDoc,'), 'ต้องมีฟังก์ชัน deleteDoc');
    assert.ok(cloudCode.includes('serializeDoc,'), 'ต้องมีฟังก์ชัน serializeDoc');
    assert.ok(cloudCode.includes('deserializeDoc,'), 'ต้องมีฟังก์ชัน deserializeDoc');
    assert.ok(cloudCode.includes('scheduleCloudSave'), 'ต้องมีฟังก์ชัน scheduleCloudSave');
    assert.ok(cloudCode.includes('setupAutoSave'), 'ต้องมีฟังก์ชัน setupAutoSave');
  });

  test('eoffice-tools.js มีแท็บและหน้าจอสำหรับ Google Drive ในแผงเครื่องมือ', () => {
    assert.ok(toolsCode.includes('data-tab="cloud"'), 'ต้องมีปุ่มแท็บ cloud');
    assert.ok(toolsCode.includes('Google Drive 📁'), 'ป้ายแท็บต้องแสดง Google Drive 📁');
    assert.ok(toolsCode.includes('data-panel="cloud"'), 'ต้องมี panel cloud');
    assert.ok(toolsCode.includes('cloudFolderBox'), 'ต้องมีกล่องแสดงโฟลเดอร์เฉพาะ Google Drive');
    assert.ok(toolsCode.includes('btnCloudSignIn'), 'ต้องมีปุ่มเข้าสู่ระบบด้วย Google');
    assert.ok(toolsCode.includes('cloudAutoSave'), 'ต้องมีตัวเลือก Auto-save');
    assert.ok(toolsCode.includes('btnCloudSaveDoc'), 'ต้องมีปุ่มบันทึกลง Drive');
    assert.ok(toolsCode.includes('cloudDocList'), 'ต้องมีพื้นที่แสดงรายการเอกสาร Drive');
    assert.ok(toolsCode.includes('cloudConfigFold'), 'ต้องมีย่อส่วนตั้งค่า Client ID');
    assert.ok(toolsCode.includes('btnSaveGdriveClient'), 'ต้องมีปุ่มบันทึก Client ID');
  });

  test('eoffice-editor.html มีปุ่ม Google Drive ในเมนูไฟล์และโหลด eoffice-cloud.js', () => {
    assert.ok(editorHtml.includes('id="btnCloudSave"'), 'เมนูไฟล์ต้องมีปุ่มบันทึกลง Google Drive');
    assert.ok(editorHtml.includes('id="btnCloudOpen"'), 'เมนูไฟล์ต้องมีปุ่มเปิดจาก Google Drive');
    assert.ok(editorHtml.includes('src="./eoffice-cloud.js"'), 'ต้องโหลดสคริปต์ eoffice-cloud.js');
  });

  test('build-single.mjs ฝัง eoffice-cloud.js และเพิ่ม CSP โดเมนของ Google Identity & Drive', () => {
    assert.ok(buildMjs.includes('eoffice-cloud.js'), 'build-single.mjs ต้องระบุ eoffice-cloud.js');
    assert.ok(buildMjs.includes('https://accounts.google.com'), 'CSP ต้องอนุญาต accounts.google.com');
    assert.ok(buildMjs.includes('https://apis.google.com'), 'CSP ต้องอนุญาต apis.google.com');
    assert.ok(buildMjs.includes('https://*.googleusercontent.com'), 'CSP ต้องอนุญาต googleusercontent.com');
  });

  test('dist/eoffice-editor-single.html ได้รับการฝัง eoffice-cloud.js สำเร็จ', () => {
    assert.ok(distSingle.includes('window.eofficeCloud'), 'dist ต้องมี window.eofficeCloud');
    assert.ok(distSingle.includes('data-tab="cloud"'), 'dist ต้องมีแท็บ cloud');
    assert.ok(distSingle.includes('E-Office Documents'), 'dist ต้องมีโฟลเดอร์ E-Office Documents');
  });

  test('ทดสอบในเบราว์เซอร์จริง (Edge Headless): ตรวจสอบ Google Drive integration, folder info และ UI', async () => {
    const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
    const port = 9347;
    const userData = join(tmpdir(), 'edge-gdrive-' + Date.now());
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
          if (window.eofficeEditor && window.eofficeCloud) return resolve();
          document.addEventListener('eoffice:ready', () => resolve(), { once: true });
          setTimeout(resolve, 2000);
        })`,
        awaitPromise: true
      });

      // ตรวจสอบโมดูล window.eofficeCloud
      const evalCloud = await send('Runtime.evaluate', {
        expression: `(() => {
          return {
            hasCloud: typeof window.eofficeCloud === 'object' && window.eofficeCloud !== null,
            provider: window.eofficeCloud?.getProvider(),
            folderName: window.eofficeCloud?.getFolderInfo()?.name,
            hasSerialize: typeof window.eofficeCloud?.serializeDoc === 'function',
            hasDeserialize: typeof window.eofficeCloud?.deserializeDoc === 'function',
            hasSignIn: typeof window.eofficeCloud?.signIn === 'function',
            hasSaveToCloud: typeof window.eofficeCloud?.saveToCloud === 'function',
            hasGetFolderInfo: typeof window.eofficeCloud?.getFolderInfo === 'function'
          };
        })()`,
        returnByValue: true
      });

      assert.ok(evalCloud.value.hasCloud, 'window.eofficeCloud ต้องถูกประกาศ');
      assert.equal(evalCloud.value.provider, 'gdrive', 'Default provider ต้องเป็น gdrive');
      assert.equal(evalCloud.value.folderName, 'E-Office Documents', 'ชื่อโฟลเดอร์ต้องเป็น E-Office Documents');
      assert.ok(evalCloud.value.hasSerialize, 'serializeDoc ต้องเป็นฟังก์ชัน');
      assert.ok(evalCloud.value.hasDeserialize, 'deserializeDoc ต้องเป็นฟังก์ชัน');
      assert.ok(evalCloud.value.hasSignIn, 'signIn ต้องเป็นฟังก์ชัน');
      assert.ok(evalCloud.value.hasSaveToCloud, 'saveToCloud ต้องเป็นฟังก์ชัน');

      // เปิดแผงเครื่องมือและสลับไปที่แท็บ Cloud (Google Drive)
      const evalUI = await send('Runtime.evaluate', {
        expression: `(() => {
          const btn = document.getElementById('btnTools');
          if (btn) btn.click();
          const cloudTab = document.querySelector('.tsub-btn[data-tab="cloud"]');
          if (cloudTab) cloudTab.click();
          const panel = document.querySelector('section[data-panel="cloud"]');
          const loginBtn = document.getElementById('btnCloudSignIn');
          const autoSaveCheckbox = document.getElementById('cloudAutoSave');
          const folderBox = document.getElementById('cloudFolderBox');
          const folderNameEl = document.getElementById('cloudFolderName');
          return {
            cloudTabSelected: cloudTab?.getAttribute('aria-selected') === 'true',
            cloudPanelVisible: panel && !panel.hidden,
            hasLoginBtn: !!loginBtn,
            hasAutoSave: !!autoSaveCheckbox,
            hasFolderBox: !!folderBox,
            folderNameText: folderNameEl?.textContent
          };
        })()`,
        returnByValue: true
      });

      assert.ok(evalUI.value.cloudTabSelected, 'แท็บ Cloud ต้องถูกเลือก (aria-selected="true")');
      assert.ok(evalUI.value.cloudPanelVisible, 'หน้าจอแผง Cloud ต้องแสดงขึ้นมา');
      assert.ok(evalUI.value.hasLoginBtn, 'ต้องมีปุ่ม Login Google Drive ในแผง Cloud');
      assert.ok(evalUI.value.hasAutoSave, 'ต้องมี Checkbox Auto-save ในแผง Cloud');
      assert.ok(evalUI.value.hasFolderBox, 'ต้องมีกล่องโฟลเดอร์เฉพาะในแผง Cloud');
      assert.equal(evalUI.value.folderNameText, 'E-Office Documents', 'ต้องแสดงชื่อโฟลเดอร์ E-Office Documents');

      // ถ่ายภาพ Screenshot บันทึกไว้เป็นหลักฐานยืนยัน
      const shot = await send('Page.captureScreenshot', { format: 'png' });
      if (shot?.data) {
        const artifactDir = 'C:\\Users\\suwit\\.gemini\\antigravity\\brain\\3e316c91-ec18-4b0a-9a1e-3224837de510';
        writeFileSync(join(artifactDir, 'gdrive-folder-panel.png'), Buffer.from(shot.data, 'base64'));
      }

      // ทดสอบการ Serialize / Deserialize เอกสาร
      const testSerialization = await send('Runtime.evaluate', {
        expression: `(() => {
          const json = window.eofficeCloud.serializeDoc();
          const docData = JSON.parse(json);
          return {
            hasVersion: docData.v === 2,
            hasPages: Array.isArray(docData.pages) && docData.pages.length > 0,
            hasSettings: typeof docData.settings === 'object'
          };
        })()`,
        returnByValue: true
      });

      assert.ok(testSerialization.value.hasVersion, 'เอกสารที่ serialize ต้องเป็น version 2');
      assert.ok(testSerialization.value.hasPages, 'เอกสารที่ serialize ต้องมี pages array');
      assert.ok(testSerialization.value.hasSettings, 'เอกสารที่ serialize ต้องมี settings');

      ws.close();
    } finally {
      edge.kill('SIGKILL');
    }
  });
});
