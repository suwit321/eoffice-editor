// ชุดทดสอบการวางใน e-Office — ไม่ผูกกับหน้าจอ: สร้าง HTML ทดสอบ, วิเคราะห์ HTML ที่คัดลอกกลับมา, จำลองการถูกล้าง
// ใช้งานผ่าน window.eofficePasteTest (แท็บ “ทดสอบวาง” ใน eoffice-tools.js เป็นผู้เรียก)
(() => {
  'use strict';

  const WJ = '⁠', NBSP = ' ';
  const PX = { cm: 96 / 2.54, mm: 96 / 25.4, in: 96, pt: 96 / 72, px: 1, em: 16, rem: 16 };
  const toPx = v => {
    const m = /^\s*(-?\d*\.?\d+)\s*(cm|mm|in|pt|px|em|rem)?\s*$/i.exec(String(v || ''));
    return m ? parseFloat(m[1]) * (PX[(m[2] || 'px').toLowerCase()] || 1) : NaN;
  };
  const near = (a, b, tol) => Number.isFinite(a) && Math.abs(a - b) <= tol;

  // อ่านค่า style จาก attribute ตรง ๆ (ไม่พึ่ง CSSOM เพราะ CSSOM ทิ้งคุณสมบัติที่เบราว์เซอร์ไม่รู้จัก)
  const styleIn = (el, prop) => {
    const m = new RegExp('(?:^|;)\\s*' + prop + '\\s*:\\s*([^;]+)', 'i').exec(el?.getAttribute?.('style') || '');
    return m ? m[1].trim() : '';
  };
  // ค่าของ element เอง → ลูกหลาน → บรรพบุรุษ (เผื่อ e-Office ย้ายสไตล์ไปอยู่ span หรือ div ครอบ)
  const styleOf = (el, prop) => {
    let v = styleIn(el, prop);
    if (v) return v;
    for (const d of el.querySelectorAll('[style]')) { v = styleIn(d, prop); if (v) return v; }
    for (let a = el.parentElement; a && a.tagName !== 'HTML'; a = a.parentElement) { v = styleIn(a, prop); if (v) return v; }
    return '';
  };
  const blockStyle = (el, prop) => {
    let v = styleIn(el, prop);
    for (let a = el.parentElement; !v && a && a.tagName !== 'HTML' && a.tagName !== 'BODY'; a = a.parentElement) v = styleIn(a, prop);
    return v;
  };

  const SAMPLE = 'ข้อความทดสอบสำหรับดูผลของรูปแบบนี้ ใช้ความยาวพอสมควรเพื่อให้ขึ้นหลายบรรทัดและเห็นการจัดแนวได้ชัดเจน';
  const IMG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
  const ok = detail => ({ state: 'ok', detail });
  const lost = detail => ({ state: 'lost', detail });
  const changed = detail => ({ state: 'changed', detail });

  // แต่ละข้อ: id (รหัสในเอกสาร), label (ชื่อ), short (ชื่อสั้นสำหรับคำเตือน), html (ย่อหน้าทดสอบ), check(block, doc) → {state, detail},
  // impact (ผลถ้าหาย), props (คุณสมบัติ CSS ที่ใช้ตอนจำลองการถูกล้าง), used (HTML ที่ส่งออกใช้ฟีเจอร์นี้หรือไม่)
  const PROBES = [
    {
      id: 'T01', key: 'textindent', label: 'ย่อหน้าบรรทัดแรก 2.5 ซม.', short: 'ย่อหน้าบรรทัดแรก', props: ['text-indent'], used: () => true,
      impact: 'ทุกย่อหน้าเนื้อความจะชิดซ้ายสุด ไม่มีการเยื้องบรรทัดแรก',
      html: `<p style="text-indent:2.5cm">[T01] ${SAMPLE}</p>`,
      check: b => { const v = blockStyle(b, 'text-indent'); if (!v) return lost('ไม่มี text-indent'); return near(toPx(v), 2.5 * PX.cm, 4) ? ok(v) : changed(`ได้ ${v} (ควร 2.5cm)`); },
    },
    {
      id: 'T02', key: 'marginleft', label: 'กั้นซ้าย 2 ซม.', short: 'กั้นซ้าย', props: ['margin-left'], used: h => /margin-left/i.test(h),
      impact: 'ส่วนลงชื่อและย่อหน้าที่กั้นซ้ายจะชิดขอบซ้าย',
      html: `<p style="margin-left:2cm">[T02] ${SAMPLE}</p>`,
      check: b => { const v = blockStyle(b, 'margin-left') || blockStyle(b, 'padding-left'); if (!v) return lost('ไม่มี margin-left'); return near(toPx(v), 2 * PX.cm, 4) ? ok(v) : changed(`ได้ ${v} (ควร 2cm)`); },
    },
    {
      id: 'T03', key: 'marginright', label: 'กั้นขวา 2 ซม.', short: 'กั้นขวา', props: ['margin-right'], used: h => /margin-right/i.test(h),
      impact: 'ย่อหน้าที่กั้นขวาจะยาวถึงขอบขวาเต็ม',
      html: `<p style="margin-right:2cm">[T03] ${SAMPLE}</p>`,
      check: b => { const v = blockStyle(b, 'margin-right') || blockStyle(b, 'padding-right'); if (!v) return lost('ไม่มี margin-right'); return near(toPx(v), 2 * PX.cm, 4) ? ok(v) : changed(`ได้ ${v} (ควร 2cm)`); },
    },
    {
      id: 'T04', key: 'lineheight', label: 'ระยะบรรทัด 200%', short: 'ระยะบรรทัด', props: ['line-height'], used: () => true,
      impact: 'ระยะบรรทัดกลับเป็นค่าเริ่มต้นของ e-Office (อาจไม่ตรงกับ 100%)',
      html: `<p style="line-height:200%">[T04] ${SAMPLE}</p>`,
      check: b => {
        const v = styleOf(b, 'line-height');
        if (!v) return lost('ไม่มี line-height');
        if (/%$/.test(v)) return near(parseFloat(v), 200, 5) ? ok(v) : changed(`ได้ ${v} (ควร 200%)`);
        if (/^[\d.]+$/.test(v)) return near(parseFloat(v), 2, 0.05) ? ok(v) : changed(`ได้ ${v} (ควร 2)`);
        const fs = toPx(styleOf(b, 'font-size')) || 16 * 96 / 72;
        return near(toPx(v) / fs, 2, 0.15) ? ok(v) : changed(`ได้ ${v}`);
      },
    },
    {
      id: 'T05', key: 'center', label: 'จัดกึ่งกลาง', short: 'จัดกึ่งกลาง', props: ['text-align'], used: h => /text-align:\s*center/i.test(h),
      impact: 'หัวหนังสือ "บันทึกข้อความ" และส่วนลงชื่อที่จัดกึ่งกลางจะชิดซ้าย',
      html: `<p style="text-align:center">[T05] ข้อความกึ่งกลาง</p>`,
      check: b => (/center/i.test(blockStyle(b, 'text-align')) || b.getAttribute('align') === 'center' ? ok('center') : lost('ไม่จัดกึ่งกลาง')),
    },
    {
      id: 'T06', key: 'justify', label: 'เสมอหน้าหลัง (justify)', short: 'เสมอหน้าหลัง', props: ['text-align'], used: h => /text-align:\s*justify/i.test(h),
      impact: 'ขอบขวาของเนื้อความไม่เรียบ',
      html: `<p style="text-align:justify;text-justify:inter-character">[T06] ${SAMPLE} ${SAMPLE}</p>`,
      check: b => (/justify/i.test(blockStyle(b, 'text-align')) || b.getAttribute('align') === 'justify' ? ok('justify') : lost('ไม่จัดเสมอหน้าหลัง')),
    },
    {
      id: 'T07', key: 'intercharacter', label: 'กระจายตัวอักษรแบบไทย (text-justify)', short: 'กระจายตัวอักษร', props: ['text-justify'], used: h => /text-justify/i.test(h),
      impact: 'ภาษาไทยที่ไม่มีเว้นวรรคจะไม่ถูกยืดให้ขอบขวาเรียบ แม้จัดเสมอหน้าหลังแล้ว',
      html: '',      // อยู่ในย่อหน้า T06 (ตรวจจากย่อหน้าเดียวกัน)
      check: (b, doc) => { const t = find(doc, 'T06'); if (!t) return lost('ไม่พบย่อหน้า T06'); const v = blockStyle(t, 'text-justify'); return /inter-character|distribute/i.test(v) ? ok(v) : lost('ไม่มี text-justify'); },
    },
    {
      id: 'T08', key: 'alignlast', label: 'กระจายถึงบรรทัดสุดท้าย (text-align-last)', short: 'กระจายบรรทัดสุดท้าย', props: ['text-align-last'], used: h => /text-align-last/i.test(h),
      impact: 'ปุ่ม "กระจายแบบไทย" จะไม่ยืดบรรทัดสุดท้าย',
      html: `<p style="text-align:justify;text-align-last:justify">[T08] ข้อความสั้น</p>`,
      check: b => (/justify/i.test(blockStyle(b, 'text-align-last')) ? ok('justify') : lost('ไม่มี text-align-last')),
    },
    {
      id: 'T09', key: 'fontfamily', label: 'แบบอักษร TH SarabunPSK', short: 'แบบอักษร', props: ['font-family'], used: () => true,
      impact: 'ข้อความใช้ฟอนต์เริ่มต้นของ e-Office แทน TH Sarabun',
      html: `<p style="font-family:'TH SarabunPSK','TH Sarabun PSK',sans-serif">[T09] ${SAMPLE}</p>`,
      check: b => { const v = styleOf(b, 'font-family'); if (!v) return lost('ไม่มี font-family'); return /sarabun/i.test(v) ? ok(v) : changed(`ได้ ${v}`); },
    },
    {
      id: 'T10', key: 'fontsize', label: 'ขนาดอักษร 20 pt', short: 'ขนาดอักษร', props: ['font-size'], used: () => true,
      impact: 'ขนาดตัวอักษรกลับเป็นค่าเริ่มต้นของ e-Office',
      html: `<p style="font-size:20pt">[T10] ขนาดยี่สิบพอยต์</p>`,
      check: b => { const v = styleOf(b, 'font-size'); if (!v) return lost('ไม่มี font-size'); return near(toPx(v), 20 * PX.pt, 1.5) ? ok(v) : changed(`ได้ ${v} (ควร 20pt)`); },
    },
    {
      id: 'T11', key: 'bius', label: 'ตัวหนา / เอียง / ขีดเส้นใต้', short: 'ตัวหนา/เอียง/ขีดเส้นใต้', unwrap: ['strong', 'b', 'em', 'i', 'u'], used: h => /<(strong|em|u|b|i)[ >]/i.test(h),
      impact: 'การเน้นข้อความด้วยตัวหนา ตัวเอียง หรือขีดเส้นใต้หายไป',
      html: `<p>[T11] <strong>หนา</strong> <em>เอียง</em> <u>ขีดเส้นใต้</u></p>`,
      check: b => {
        const has = {
          'ตัวหนา': !!b.querySelector('strong,b') || /bold|[6-9]00/.test(styleOf(b, 'font-weight')),
          'ตัวเอียง': !!b.querySelector('em,i') || /italic/.test(styleOf(b, 'font-style')),
          'ขีดเส้นใต้': !!b.querySelector('u') || /underline/.test(styleOf(b, 'text-decoration')),
        };
        const missing = Object.keys(has).filter(k => !has[k]);
        return !missing.length ? ok('ครบ') : missing.length === 3 ? lost('หายทั้งหมด') : changed('หาย: ' + missing.join(', '));
      },
    },
    {
      id: 'T12', key: 'color', label: 'สีตัวอักษร / สีเน้น', short: 'สีตัวอักษร/สีเน้น', props: ['color', 'background-color', 'background'], used: h => /(?:^|[;"\s])(color|background-color):/i.test(h),
      impact: 'ข้อความที่ใช้สีหรือไฮไลต์จะเป็นสีปกติ',
      html: `<p>[T12] <span style="color:#c00000">ตัวอักษรสีแดง</span> <span style="background-color:#ffff00">เน้นสีเหลือง</span></p>`,
      check: b => {
        const c = /#c00000|rgb\(\s*192\s*,\s*0\s*,\s*0\s*\)|red/i.test(styleOf(b, 'color'));
        const g = /#ffff00|rgb\(\s*255\s*,\s*255\s*,\s*0\s*\)|yellow/i.test(styleOf(b, 'background-color') || styleOf(b, 'background'));
        return c && g ? ok('ครบ') : !c && !g ? lost('หายทั้งสองสี') : changed(c ? 'สีเน้นหาย' : 'สีตัวอักษรหาย');
      },
    },
    {
      id: 'T13', key: 'letterspacing', label: 'ช่องไฟตัวอักษร 0.1 em', short: 'ช่องไฟตัวอักษร', props: ['letter-spacing'], used: h => /letter-spacing/i.test(h),
      impact: 'การปรับช่องไฟเพื่อดึงคำโดดหรือลดการยืดจะไม่มีผล',
      html: `<p><span style="letter-spacing:0.1em">[T13] ช่องไฟกว้างหนึ่งในสิบ</span></p>`,
      check: b => { const v = styleOf(b, 'letter-spacing'); if (!v || /^(normal|0(px|em)?)$/i.test(v)) return lost('ไม่มี letter-spacing'); return ok(v); },
    },
    {
      id: 'T14', key: 'supsub', label: 'ตัวยก / ตัวห้อย', short: 'ตัวยก/ตัวห้อย', unwrap: ['sup', 'sub'], used: h => /<(sup|sub)[ >]/i.test(h),
      impact: 'ตัวยก ตัวห้อยกลายเป็นตัวปกติ (เช่น ๑๒ ม.๓ จะเพี้ยน)',
      html: `<p>[T14] ตัวยก <sup>2</sup> และตัวห้อย <sub>2</sub></p>`,
      check: b => {
        const sup = !!b.querySelector('sup') || /super/.test(styleOf(b, 'vertical-align'));
        const sub = !!b.querySelector('sub');
        return sup && sub ? ok('ครบ') : !sup && !sub ? lost('หายทั้งสอง') : changed(sup ? 'ตัวห้อยหาย' : 'ตัวยกหาย');
      },
    },
    {
      id: 'T15', key: 'wj', label: 'Word Joiner (ผูกคำห้ามตัด)', short: 'Word Joiner', wj: true, used: h => h.includes(WJ),
      impact: 'คำที่ผูกไว้อาจถูกตัดกลางคำตอนขึ้นบรรทัดใหม่ (แอปจะถอดตัวผูกออกตอนส่งออกให้เอง)',
      html: `<p>[T15] ผู้${WJ}อำนวย${WJ}การ${WJ}สำนัก${WJ}งาน${WJ}คณะ${WJ}กรรมการ</p>`,
      check: b => { const t = b.textContent; if (t.includes(WJ)) return ok('ยังอยู่'); return /ผู้อำนวยการสำนักงานคณะกรรมการ/.test(t) ? lost('ถูกลบ (ข้อความยังครบ)') : changed('ข้อความเปลี่ยน: ' + t.slice(0, 50)); },
    },
    {
      id: 'T16', key: 'nbsp', label: 'เว้นวรรคห้ามตัดบรรทัด (nbsp) และเว้นวรรคซ้อน', short: 'เว้นวรรคห้ามตัด', nbsp: true, used: h => /&nbsp;| /.test(h),
      impact: 'เว้นวรรคซ้อนยุบเหลือตัวเดียว และ “๑๒ บาท” อาจถูกตัดแยกบรรทัด',
      html: `<p>[T16] ก${NBSP}${NBSP}${NBSP}ข ๑๒${NBSP}บาท</p>`,
      check: b => {
        const t = b.textContent.replace(/\[T16\]\s*/, '');
        if (t.includes(NBSP)) return /ก\s{3}ข/.test(t) ? ok('คง nbsp') : changed('nbsp ยังอยู่แต่จำนวนเปลี่ยน');
        return /ก\s{2,}ข/.test(t) ? changed('เป็นช่องว่างปกติหลายตัว (อาจยุบตอนแสดงผล)') : lost('ยุบเหลือช่องว่างเดียว');
      },
    },
    {
      id: 'T17', key: 'tab', label: 'แท็บ (Tab) และจุดแท็บ', short: 'แท็บ', tab: true, used: h => /white-space:\s*pre/i.test(h),
      impact: 'ช่อง “วันที่” ในบันทึกข้อความจะไม่ตรงตำแหน่ง (แอปจะใช้ช่องว่างห้ามตัด 8 ตัวแทนให้เอง)',
      html: `<p style="tab-size:4cm">[T17] ก<span style="white-space:pre">\t</span>ข<span style="white-space:pre">\t</span>ค</p>`,
      check: b => {
        const t = b.textContent;
        if (!t.includes('\t')) return lost(/ก\s{2,}ข/.test(t) ? 'Tab กลายเป็นช่องว่าง' : 'Tab หาย');
        const pre = !!b.querySelector('[style*="pre"]') || /pre/.test(blockStyle(b, 'white-space'));
        return pre ? ok('Tab + white-space:pre ยังอยู่') : changed('มี Tab แต่ไม่มี white-space:pre (จะแสดงเป็นช่องว่างเดียว)');
      },
    },
    {
      id: 'T18', key: 'spacing', label: 'ระยะก่อน/หลังย่อหน้า 12 pt', short: 'ระยะก่อน/หลังย่อหน้า', props: ['margin-top', 'margin-bottom'], used: h => /margin-(top|bottom):\s*(?!0)/i.test(h),
      impact: 'ย่อหน้าติดกันโดยไม่มีระยะเว้น',
      html: `<p style="margin-top:12pt;margin-bottom:12pt">[T18] ย่อหน้าที่มีระยะเว้นบนและล่าง</p>`,
      check: b => { const t = blockStyle(b, 'margin-top'), bt = blockStyle(b, 'margin-bottom'); const px = v => (v ? toPx(v) : 0); return near(px(t), 16, 3) || near(px(bt), 16, 3) ? ok(`${t || '-'} / ${bt || '-'}`) : (t || bt) ? changed(`${t || '-'} / ${bt || '-'}`) : lost('ไม่มี margin'); },
    },
    {
      id: 'T19', key: 'olthai', label: 'รายการลำดับเลขไทย (๑. ๒. ๓.)', short: 'เลขรายการไทย', props: ['list-style-type'], used: h => /list-style-type:\s*thai/i.test(h),
      impact: 'เลขรายการเป็น 1. 2. 3. แทนเลขไทย',
      html: `<ol style="list-style-type:thai"><li>[T19] ข้อที่หนึ่ง</li><li>ข้อที่สอง</li></ol>`,
      check: b => {
        const l = b.closest('li'), ol = l?.parentElement;
        if (!l || !/^(OL|UL)$/.test(ol?.tagName || '')) return changed('ไม่เป็นรายการ (เป็นย่อหน้าธรรมดา)');
        return /thai/i.test(styleOf(ol, 'list-style-type')) ? ok('thai') : changed('เป็นรายการแล้ว แต่ไม่ใช่เลขไทย');
      },
    },
    {
      id: 'T20', key: 'ul', label: 'รายการสัญลักษณ์ (•)', short: 'รายการสัญลักษณ์', used: h => /<ul[ >]/i.test(h),
      impact: 'หัวข้อย่อยแบบจุดกลายเป็นย่อหน้าธรรมดา',
      html: `<ul><li>[T20] หัวข้อที่หนึ่ง</li><li>หัวข้อที่สอง</li></ul>`,
      check: b => { const l = b.closest('li'); return l && l.parentElement?.tagName === 'UL' ? ok('UL') : l ? changed('เป็น ' + l.parentElement?.tagName) : lost('ไม่เป็นรายการ'); },
    },
    {
      id: 'T21', key: 'table', label: 'ตารางมีเส้นขอบ', short: 'ตาราง', used: h => /<table[ >]/i.test(h),
      impact: 'ตารางหายหรือไม่มีเส้นขอบ',
      html: `<table style="border-collapse:collapse;width:100%"><tbody><tr><td style="border:1px solid #000;padding:2px 6px">[T21] ช่อง ๑</td><td style="border:1px solid #000;padding:2px 6px">ช่อง ๒</td></tr></tbody></table>`,
      check: b => {
        const td = b.closest('td,th');
        if (!td) return lost('ไม่เป็นตาราง');
        const border = styleOf(td, 'border') || styleOf(td, 'border-left') || styleOf(td.closest('table'), 'border') || td.closest('table').getAttribute('border');
        return border && !/^(0|none)/i.test(border) ? ok(border) : changed('มีตารางแต่ไม่มีเส้นขอบ');
      },
    },
    {
      id: 'T22', key: 'link', label: 'ลิงก์', short: 'ลิงก์', used: h => /<a[ >]/i.test(h),
      impact: 'ลิงก์กลายเป็นข้อความธรรมดา',
      html: `<p>[T22] <a href="https://example.go.th/path">ลิงก์ทดสอบ</a></p>`,
      check: b => { const a = b.querySelector('a'); return a?.getAttribute('href')?.startsWith('http') ? ok(a.getAttribute('href')) : a ? changed('มี <a> แต่ไม่มี href') : lost('ไม่มีลิงก์'); },
    },
    {
      id: 'T23', key: 'image', label: 'รูปภาพที่ฝังในข้อความ (data URI)', short: 'รูปภาพ', used: h => /<img[ >]/i.test(h),
      impact: 'รูปครุฑและรูปในเอกสารไม่แสดง — ต้องอัปโหลดรูปแยกใน e-Office',
      html: `<p>[T23] <img src="${IMG}" width="24" height="24" alt="รูปทดสอบ"></p>`,
      check: b => { const i = b.querySelector('img'); if (!i) return lost('ไม่มีรูป'); const s = i.getAttribute('src') || ''; return /^data:image\//.test(s) ? ok('data URI ยังอยู่') : changed('src เปลี่ยนเป็น ' + s.slice(0, 40)); },
    },
    {
      id: 'T24', key: 'thaitext', label: 'ข้อความไทยครบถ้วน (สระ วรรณยุกต์ เลขไทย เครื่องหมาย)', short: 'ข้อความไทย', used: () => true,
      impact: 'ข้อความเพี้ยน — ห้ามใช้วิธีวางนี้จนกว่าจะแก้',
      html: `<p>[T24] ที่ ผู้ว่าราชการจังหวัด ๑๒,๕๐๐ บาท “ข้อความ” ฯลฯ ต่าง ๆ พ.ศ. ๒๕๖๙ ศรี สุวรรณ ก๋วยเตี๋ยว นายก ฏ ฐ ฤๅ</p>`,
      check: b => {
        const want = 'ที่ ผู้ว่าราชการจังหวัด ๑๒,๕๐๐ บาท “ข้อความ” ฯลฯ ต่าง ๆ พ.ศ. ๒๕๖๙ ศรี สุวรรณ ก๋วยเตี๋ยว นายก ฏ ฐ ฤๅ';
        const got = b.textContent.replace(/\[T24\]\s*/, '').replace(/[ ​⁠]/g, m => (m === NBSP ? ' ' : '')).replace(/\s+/g, ' ').trim();
        return got === want ? ok('ตรงทุกตัวอักษร') : changed('ได้: ' + got.slice(0, 60));
      },
    },
    {
      id: 'T25', key: 'blank', label: 'บรรทัดว่างระหว่างย่อหน้า', short: 'บรรทัดว่าง', used: h => /<br>/i.test(h),
      impact: 'บรรทัดว่างที่ใช้เว้นระยะ (เช่น ก่อนลงชื่อ) หายไป',
      html: `<p>[T25a] ย่อหน้าก่อนบรรทัดว่าง</p><p><br></p><p>[T25b] ย่อหน้าหลังบรรทัดว่าง</p>`,
      check: (b, doc) => {
        const a = find(doc, 'T25a'), z = find(doc, 'T25b');
        if (!a || !z) return lost('ไม่พบย่อหน้า T25a/T25b');
        let between = 0;
        for (let n = a.nextSibling; n && n !== z; n = n.nextSibling) if (n.nodeType === 1) between++;
        if (a.parentElement !== z.parentElement) between = a.parentElement && z.parentElement && (a.parentElement.querySelector('br') || a.parentElement.nextElementSibling) ? 1 : 0;
        return between > 0 ? ok('มีบรรทัดว่างคั่น') : lost('ย่อหน้าติดกัน');
      },
    },
  ];

  // หา block ที่มีรหัส [Txx]
  function find(doc, id) {
    const w = doc.createTreeWalker(doc.body || doc, NodeFilter.SHOW_TEXT);
    const tag = `[${id}]`;
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      if (n.data.includes(tag)) return n.parentElement.closest('p,li,td,th,div,h1,h2,h3,h4,h5,h6') || n.parentElement;
    }
    return null;
  }

  // HTML ของชุดทดสอบ (ใส่หัวคำอธิบาย + ทุกข้อ)
  function buildHtml() {
    const head = `<p style="font-weight:bold">ชุดทดสอบการวางใน e-Office — ห้ามแก้รหัส [Txx] หน้าแต่ละข้อ</p>`
      + `<p>ทดสอบเมื่อ ${new Date().toLocaleString('th-TH')}</p>`;
    return head + PROBES.filter(p => p.html).map(p => p.html).join('\n');
  }
  const buildText = () => new DOMParser().parseFromString(buildHtml(), 'text/html').body.innerText || '';

  // วิเคราะห์ HTML ที่คัดลอกกลับมา: คืน { results: [{id,key,label,state,detail,impact}], found, total }
  function analyze(html) {
    const doc = new DOMParser().parseFromString(String(html || ''), 'text/html');
    let found = 0;
    const results = PROBES.map(p => {
      const marker = p.id === 'T07' ? 'T06' : p.id === 'T25' ? 'T25a' : p.id;
      const block = find(doc, marker);
      let r;
      if (!block) r = { state: 'missing', detail: 'ไม่พบรหัส [' + marker + '] — วางไม่ครบหรือถูกลบ' };
      else { found++; try { r = p.check(block, doc); } catch (e) { r = { state: 'changed', detail: 'วิเคราะห์ไม่ได้: ' + e.message }; } }
      return { id: p.id, key: p.key || p.id, label: p.label, short: p.short, impact: p.impact, ...r };
    });
    return { results, found, total: PROBES.length };
  }

  // จำลองการถูกล้างตามผลทดสอบ (ใช้ในตัวอย่างหลังส่งออก) — โดยประมาณ
  function simulate(html, results) {
    const state = Object.fromEntries((results || []).map(r => [r.id, r.state]));
    const bad = id => state[id] && state[id] !== 'ok';
    const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
    const strip = new Set();
    for (const p of PROBES) if (bad(p.id) && p.props) p.props.forEach(x => strip.add(x));
    for (const el of doc.body.querySelectorAll('[style]')) {
      el.setAttribute('style', el.getAttribute('style').split(';').filter(d => d.trim() && !strip.has(d.split(':')[0].trim().toLowerCase())).join(';'));
      if (!el.getAttribute('style')) el.removeAttribute('style');
    }
    for (const p of PROBES) if (bad(p.id) && p.unwrap) {
      for (const el of [...doc.body.querySelectorAll(p.unwrap.join(','))]) el.replaceWith(...el.childNodes);
    }
    if (bad('T21')) doc.body.querySelectorAll('[style]').forEach(e => { e.style.removeProperty('border'); });
    if (bad('T22')) doc.body.querySelectorAll('a').forEach(a => a.removeAttribute('href'));
    if (bad('T23')) doc.body.querySelectorAll('img').forEach(i => i.remove());
    const w = doc.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      if (bad('T15')) n.data = n.data.replaceAll(WJ, '');
      if (bad('T16')) n.data = n.data.replaceAll(NBSP, ' ');
      if (bad('T17')) n.data = n.data.replaceAll('\t', ' ');
    }
    return doc.body.innerHTML;
  }

  // โปรไฟล์การส่งออกจากผลทดสอบ (ใช้ตั้ง ed.setExportProfile)
  const profileOf = results => {
    const st = Object.fromEntries((results || []).map(r => [r.id, r.state]));
    return { stripWJ: !!st.T15 && st.T15 !== 'ok', tabsAsSpaces: !!st.T17 && st.T17 !== 'ok' };
  };
  // ฟีเจอร์ที่เอกสาร (HTML ที่ส่งออก) ใช้อยู่ แต่ e-Office จะไม่รองรับ — ใช้เป็นคำเตือนตอนคัดลอก
  // (Word Joiner กับแท็บมีวิธีสำรองในตัวส่งออกอยู่แล้ว จึงไม่ต้องเตือน)
  const adviceFor = (results, html) => (results || [])
    .filter(r => r.state !== 'ok' && r.state !== 'missing')
    .map(r => ({ r, p: PROBES.find(p => p.id === r.id) }))
    .filter(({ p }) => p && !p.wj && !p.tab && (p.used ? p.used(html) : true))
    .map(({ r }) => r.short || r.label);

  window.eofficePasteTest = { PROBES, buildHtml, buildText, analyze, simulate, profileOf, adviceFor };
})();
