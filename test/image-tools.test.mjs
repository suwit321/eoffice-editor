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

describe('Image Tools & Formatting (การจัดการและจัดรูปแบบรูปภาพอิสระ)', () => {
  const html = readFileSync(join(root, 'eoffice-editor.html'), 'utf8');
  const distSingle = readFileSync(join(root, 'dist/eoffice-editor-single.html'), 'utf8');

  test('โครงสร้าง HTML มี Image Overlay, มือจับ 8 ทิศ, แถบเครื่องมือลอย และหน้าต่างครอบตัด (Crop)', () => {
    assert.ok(html.includes('id="imgOverlay"'), 'ต้องมี #imgOverlay');
    assert.ok(html.includes('id="imgSizeBadge"'), 'ต้องมี #imgSizeBadge แสดงขนาด px และ ซม.');
    assert.ok(html.includes('id="imgToolbar"'), 'ต้องมี #imgToolbar แถบคำสั่งลอย');
    assert.ok(html.includes('data-dir="nw"'), 'ต้องมี handle ทิศ nw');
    assert.ok(html.includes('data-dir="se"'), 'ต้องมี handle ทิศ se');
    assert.ok(html.includes('data-dir="e"'), 'ต้องมี handle ทิศ e');
    assert.ok(html.includes('data-dir="s"'), 'ต้องมี handle ทิศ s');
    assert.ok(html.includes('id="imgCropOverlay"'), 'ต้องมี #imgCropOverlay สำหรับครอบตัด');
    assert.ok(html.includes('id="imgCropBox"'), 'ต้องมี #imgCropBox');
    assert.ok(html.includes('id="btnCropApply"'), 'ต้องมีปุ่มยืนยันครอบตัด #btnCropApply');
    assert.ok(html.includes('id="btnCropCancel"'), 'ต้องมีปุ่มยกเลิกครอบตัด #btnCropCancel');
  });

  test('แถบเครื่องมือลอย (#imgToolbar) มีปุ่มจัดตำแหน่ง, ล้อมรอบข้อความ, ครอบตัด, หมุน, ขนาดสำเร็จรูป และลบ', () => {
    assert.ok(html.includes('data-act="align" data-val="left"'), 'มีปุ่มจัดชิดซ้าย');
    assert.ok(html.includes('data-act="align" data-val="center"'), 'มีปุ่มจัดกึ่งกลาง');
    assert.ok(html.includes('data-act="align" data-val="right"'), 'มีปุ่มจัดชิดขวา');
    assert.ok(html.includes('data-act="wrap" data-val="inline"'), 'มีปุ่มในแนวบรรทัด');
    assert.ok(html.includes('data-act="wrap" data-val="left"'), 'มีปุ่มล้อมรอบขวา (ตัดข้อความ)');
    assert.ok(html.includes('data-act="wrap" data-val="right"'), 'มีปุ่มล้อมรอบซ้าย');
    assert.ok(html.includes('data-act="crop"'), 'มีปุ่มครอบตัด');
    assert.ok(html.includes('data-act="rotate"'), 'มีปุ่มหมุน 90°');
    assert.ok(html.includes('id="imgSizePreset"'), 'มี dropdown ขนาดภาพสำเร็จรูป');
    assert.ok(html.includes('data-act="delete"'), 'มีปุ่มลบรูปภาพ');
  });

  test('CSS มีสไตล์รองรับการจัดวางภาพ, ขอบไฮไลต์, เส้นประครอบตัด และการพิมพ์/ส่งออก', () => {
    assert.ok(html.includes('#editor img[data-align="center"]'), 'มี CSS จัดกึ่งกลางบนจอ');
    assert.ok(html.includes('#printArea img[data-align="center"]'), 'มี CSS จัดกึ่งกลางสำหรับพิมพ์');
    assert.ok(html.includes('#editor img[data-wrap="left"]'), 'มี CSS float left');
    assert.ok(html.includes('#printArea img[data-wrap="left"]'), 'มี CSS float left สำหรับพิมพ์');
    assert.ok(html.includes('.img-overlay'), 'มี CSS .img-overlay');
    assert.ok(html.includes('.img-handle'), 'มี CSS .img-handle');
    assert.ok(html.includes('.img-size-badge'), 'มี CSS .img-size-badge');
    assert.ok(html.includes('.img-crop-box'), 'มี CSS .img-crop-box');
  });

  test('ฟังก์ชัน cleanHtml/exportHtml บันทึกคุณสมบัติภาพและรูปแบบ inline ครบถ้วน', () => {
    assert.ok(html.includes('KEEP_ATTRS = new Set'), 'มี KEEP_ATTRS');
    assert.ok(html.includes("'data-align'"), 'KEEP_ATTRS ต้องอนุญาต data-align');
    assert.ok(html.includes("'data-wrap'"), 'KEEP_ATTRS ต้องอนุญาต data-wrap');
    assert.ok(html.includes('srcImgs.forEach'), 'exportHtml ต้องคัดลอกแอตทริบิวต์และสไตล์ของรูปภาพ');
  });

  test('ไฟล์ผลลัพธ์ dist/eoffice-editor-single.html มีฟังก์ชันจัดการรูปภาพอิสระครบถ้วน', () => {
    assert.ok(distSingle.includes('id="imgOverlay"'), 'distSingle ต้องมี #imgOverlay');
    assert.ok(distSingle.includes('id="imgToolbar"'), 'distSingle ต้องมี #imgToolbar');
    assert.ok(distSingle.includes('id="imgCropOverlay"'), 'distSingle ต้องมี #imgCropOverlay');
    assert.ok(distSingle.includes('imageTools:'), 'distSingle ต้อง export imageTools');
  });

  test('ทดสอบในเบราว์เซอร์จริง (Edge Headless): จัดการรูปภาพ ย่อ/ขยาย จัดตำแหน่ง ล้อมรอบ หมุน ครอบตัด และลบ', async (t) => {
    const edgePath = 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe';
    if (!existsSync(edgePath)) {
      t.skip('ไม่พบ msedge.exe ข้ามการทดสอบเบราว์เซอร์จริง');
      return;
    }

    const port = 9310 + Math.floor(Math.random() * 30);
    const userData = join(tmpdir(), 'edge-imgtools-' + Date.now() + '-' + Math.random().toString(36).slice(2));
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

      const evalCode = async (expr) => {
        const res = await send('Runtime.evaluate', { expression: expr, returnByValue: true, awaitPromise: true });
        return res?.value;
      };

      // 1. รอให้เอดิเตอร์พร้อมใช้งาน
      await evalCode(`new Promise(resolve => {
        if (window.eofficeEditor?.quill) return resolve(true);
        document.addEventListener('eoffice:ready', () => resolve(true), { once: true });
      })`);

      // 2. แทรกรูปภาพจำลองลงใน Quill
      const sampleImg = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAYAAABw4pVUAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAA8SURBVHhe7cExAQAAAMKg9U9tCy8gAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA4FsDXYAAAc8k60sAAAAASUVORK5CYII=';

      await evalCode(`(() => {
        const q = window.eofficeEditor.quill;
        q.setContents([
          { insert: 'ข้อความก่อนรูปภาพ\\n' },
          { insert: { image: '${sampleImg}' } },
          { insert: '\\nข้อความหลังรูปภาพ\\n' }
        ], 'user');
      })()`);

      // 3. ตรวจสอบการเลือกรูปภาพและแสดง Overlay
      const selectResult = await evalCode(`(() => {
        const img = document.querySelector('#editor img');
        if (!img) return { ok: false, error: 'no img' };
        window.eofficeEditor.imageTools.selectImage(img);
        const overlay = document.querySelector('#imgOverlay');
        const badge = document.querySelector('#imgSizeBadge');
        return {
          ok: true,
          overlayVisible: !overlay.hidden,
          badgeText: badge.textContent,
          hasImgSelected: img.classList.contains('img-selected')
        };
      })()`);

      assert.strictEqual(selectResult.ok, true, 'ต้องพบรูปภาพ');
      assert.strictEqual(selectResult.overlayVisible, true, 'Overlay ต้องแสดงเมื่อเลือกรูป');
      assert.strictEqual(selectResult.hasImgSelected, true, 'รูปภาพต้องมี class img-selected');
      assert.ok(selectResult.badgeText.includes('px'), 'ป้ายขนาดต้องระบุหน่วย px');
      assert.ok(selectResult.badgeText.includes('ซม.'), 'ป้ายขนาดต้องระบุหน่วย ซม.');

      // 4. ทดสอบการจัดตำแหน่ง (กึ่งกลาง, ชิดขวา)
      await evalCode(`window.eofficeEditor.imageTools.setImageAlign('center')`);
      const alignCenter = await evalCode(`(() => {
        const img = document.querySelector('#editor img');
        const clean = window.eofficeEditor.cleanHtml();
        return {
          datasetAlign: img.dataset.align,
          inCleanHtml: clean.includes('data-align="center"')
        };
      })()`);
      assert.strictEqual(alignCenter.datasetAlign, 'center', 'ต้องมี data-align="center" บนรูปภาพ');
      assert.strictEqual(alignCenter.inCleanHtml, true, 'cleanHtml ต้องรักษา data-align="center"');

      // 5. ทดสอบการตัดข้อความล้อมรูปภาพ (Wrap / Float)
      await evalCode(`window.eofficeEditor.imageTools.setImageWrap('left')`);
      const wrapLeft = await evalCode(`(() => {
        const img = document.querySelector('#editor img');
        const clean = window.eofficeEditor.cleanHtml();
        return {
          datasetWrap: img.dataset.wrap,
          noAlign: !img.dataset.align,
          inCleanHtml: clean.includes('data-wrap="left"')
        };
      })()`);
      assert.strictEqual(wrapLeft.datasetWrap, 'left', 'ต้องมี data-wrap="left" บนรูปภาพ');
      assert.strictEqual(wrapLeft.noAlign, true, 'การล้อมข้อความต้องเคลียร์ data-align');
      assert.strictEqual(wrapLeft.inCleanHtml, true, 'cleanHtml ต้องรักษา data-wrap="left"');

      // 6. ทดสอบการปรับขนาดสำเร็จรูป (Preset 50%)
      await evalCode(`window.eofficeEditor.imageTools.setImageSizePreset('50')`);
      const presetSize = await evalCode(`(() => {
        const img = document.querySelector('#editor img');
        return {
          widthStyle: img.style.width,
          widthAttr: img.getAttribute('width')
        };
      })()`);
      assert.ok(parseInt(presetSize.widthStyle, 10) > 0, 'ต้องกำหนด widthStyle เมื่อเลือก Preset');

      // 7. ทดสอบโหมดครอบตัดภาพ (Crop Mode)
      await evalCode(`window.eofficeEditor.imageTools.startCropMode()`);
      const cropState = await evalCode(`(() => {
        const cropOverlay = document.querySelector('#imgCropOverlay');
        const cropBox = document.querySelector('#imgCropBox');
        return {
          cropVisible: !cropOverlay.hidden,
          boxW: cropBox.offsetWidth,
          boxH: cropBox.offsetHeight
        };
      })()`);
      assert.strictEqual(cropState.cropVisible, true, 'Crop Overlay ต้องแสดงเมื่อเข้าสู่โหมดครอบตัด');
      assert.ok(cropState.boxW > 0, 'Crop Box ต้องมีขนาด');

      // กดตกลงครอบตัด
      await evalCode(`window.eofficeEditor.imageTools.applyCrop()`);
      const afterCrop = await evalCode(`(() => {
        const cropOverlay = document.querySelector('#imgCropOverlay');
        const img = document.querySelector('#editor img');
        return {
          cropOverlayHidden: cropOverlay.hidden,
          imgSrc: img.src.slice(0, 30)
        };
      })()`);
      assert.strictEqual(afterCrop.cropOverlayHidden, true, 'เมื่อครอบตัดแล้ว Overlay ครอบตัดต้องถูกซ่อน');
      assert.ok(afterCrop.imgSrc.startsWith('data:image/png;base64,'), 'รูปภาพหลังครอบตัดต้องเป็น data URL สมบูรณ์');

      // 8. ทดสอบการหมุนภาพ 90° (Rotate)
      await evalCode(`window.eofficeEditor.imageTools.rotateActiveImage()`);
      const afterRotate = await evalCode(`(() => {
        const img = document.querySelector('#editor img');
        return {
          imgConnected: img.isConnected,
          hasSrc: !!img.src
        };
      })()`);
      assert.strictEqual(afterRotate.imgConnected, true, 'รูปภาพยังคงอยู่ในเอกสารหลังหมุน');
      assert.strictEqual(afterRotate.hasSrc, true, 'รูปภาพมี src สมบูรณ์หลังหมุน');

      // 9. ทดสอบการลบรูปภาพ (Delete)
      await evalCode(`window.eofficeEditor.imageTools.deleteActiveImage()`);
      const afterDelete = await evalCode(`(() => {
        const img = document.querySelector('#editor img');
        const overlay = document.querySelector('#imgOverlay');
        return {
          imgExists: !!img,
          overlayHidden: overlay.hidden
        };
      })()`);
      assert.strictEqual(afterDelete.imgExists, false, 'รูปภาพต้องถูกลบออกจากเอกสาร');
      assert.strictEqual(afterDelete.overlayHidden, true, 'Overlay ต้องถูกซ่อนเมื่อลบรูป');

      ws.close();
    } finally {
      edge.kill('SIGKILL');
    }
  });
});
