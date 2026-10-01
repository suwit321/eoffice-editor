// แผงเครื่องมือเสริมของ eoffice-editor.html — แยกจากส่วนเอดิเตอร์ (Quill)
// ใช้งานเอดิเตอร์ผ่าน window.eofficeEditor ที่ไฟล์หลักเปิดไว้ และเริ่มทำงานเมื่อได้รับเหตุการณ์ eoffice:ready
(() => {
  'use strict';

  const start = () => { if (window.eofficeEditor && !start.done) { start.done = true; initTools(window.eofficeEditor); } };
  document.addEventListener('eoffice:ready', start);
  start();

  function initTools(ed) {
    const { quill, Delta, say } = ed;
    const $ = (s, root = document) => root.querySelector(s);
    const $$ = (s, root = document) => [...root.querySelectorAll(s)];
    const store = {
      get(k) { try { return localStorage.getItem(k); } catch { return null; } },
      set(k, v) { try { localStorage.setItem(k, v); } catch { /* ไม่มี storage */ } },
      del(k) { try { localStorage.removeItem(k); } catch { /* ไม่มี storage */ } },
    };
    const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    const KEY = {
      theme: 'eoffice-editor:theme', warmth: 'eoffice-editor:warmth', panel: 'eoffice-editor:tools', apiKey: 'eoffice-editor:apikey',
      aiCfg: 'eoffice-editor:ai', presets: 'eoffice-editor:presets', phrases: 'eoffice-editor:phrases', tab: 'eoffice-editor:tools-tab',
    };

    // ================= ข้อความของเอกสาร =================
    // ข้อความของ delta โดยแทนรูป/ตารางฝังด้วย U+FFFC เพื่อให้ตำแหน่งตรงกับ index ของ Quill
    const OBJ = '￼';
    const docText = delta => delta.ops.map(op => (typeof op.insert === 'string' ? op.insert : OBJ)).join('');
    const attrsAt = (delta, i) => delta.slice(i, i + 1).ops[0]?.attributes;

    // หาข้อความที่ตรง regex แล้วแทนที่ในหน้าที่เลือก (scope: 'selection' | 'page' | 'all')
    // คืนจำนวนจุดที่แทน — หน้าปัจจุบันแก้ผ่าน Quill (กดย้อนกลับได้) หน้าอื่นแก้ที่ delta โดยตรง
    function replaceInDoc(scope, regex, replacer) {
      let total = 0;
      const edit = (delta, from, to) => {
        const text = docText(delta);
        const hits = [];
        regex.lastIndex = 0;
        for (let m = regex.exec(text); m; m = regex.exec(text)) {
          if (!m[0].length) { regex.lastIndex++; continue; }
          if (m.index < from || m.index + m[0].length > to || m[0].includes(OBJ)) continue;
          const rep = typeof replacer === 'function' ? replacer(...m) : replacer;
          if (rep !== m[0]) hits.push({ index: m.index, length: m[0].length, rep });
        }
        let change = new Delta(), pos = 0;
        for (const h of hits) {
          change = change.retain(h.index - pos).delete(h.length);
          if (h.rep) change = change.insert(h.rep, attrsAt(delta, h.index));
          pos = h.index + h.length;
        }
        total += hits.length;
        return hits.length ? change : null;
      };
      const r = ed.range();
      if (scope === 'selection' && !r.length) scope = 'page';
      const cur = quill.getContents();
      const change = scope === 'selection' ? edit(cur, r.index, r.index + r.length) : edit(cur, 0, cur.length());
      if (change) quill.updateContents(change, 'user');
      if (scope === 'all') {
        ed.pages().forEach((p, i) => {
          if (i === ed.currentPage()) return;
          const c = edit(p.delta, 0, p.delta.length());
          if (c) ed.setPageDelta(i, p.delta.compose(c));
        });
      }
      return total;
    }
    const scopeLabel = s => ({ selection: 'ส่วนที่เลือก', page: 'หน้านี้', all: 'ทุกหน้า' }[s]);

    // ================= ตัวเลข / เงิน / วันที่ =================
    const ARABIC = '0123456789', THAI = '๐๑๒๓๔๕๖๗๘๙';
    const toThai = s => s.replace(/[0-9]/g, d => THAI[d]);
    const toArabic = s => s.replace(/[๐-๙]/g, d => ARABIC[THAI.indexOf(d)]);

    const DIGIT = ['ศูนย์', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
    const PLACE = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน'];
    function readUnderMillion(n, hasHigher) {
      const s = String(n);
      let out = '';
      for (let i = 0; i < s.length; i++) {
        const d = +s[i], p = s.length - 1 - i;
        if (!d) continue;
        if (p === 1 && d === 1) out += 'สิบ';
        else if (p === 1 && d === 2) out += 'ยี่สิบ';
        else if (p === 0 && d === 1 && (s.length > 1 || hasHigher)) out += 'เอ็ด';
        else out += DIGIT[d] + PLACE[p];
      }
      return out;
    }
    function readInt(n) {
      if (n === 0) return 'ศูนย์';
      const high = Math.floor(n / 1e6), low = n % 1e6;
      return (high ? readInt(high) + 'ล้าน' : '') + (low ? readUnderMillion(low, high > 0) : '');
    }
    // จำนวนเงินเป็นตัวอักษร เช่น 1,250.50 → หนึ่งพันสองร้อยห้าสิบบาทห้าสิบสตางค์
    function bahtText(input) {
      const raw = toArabic(String(input)).replace(/[,\s฿]|บาท/g, '');
      if (!/^-?\d+(\.\d+)?$/.test(raw)) return null;
      const n = Number(raw);
      if (Math.abs(n) >= 1e15) return null;
      const total = Math.round(Math.abs(n) * 100);
      const baht = Math.floor(total / 100), satang = total % 100;
      let s = baht ? readInt(baht) + 'บาท' : '';
      s += satang ? readUnderMillion(satang, false) + 'สตางค์' : (baht ? 'ถ้วน' : 'ศูนย์บาทถ้วน');
      return (n < 0 ? 'ลบ' : '') + s;
    }
    function thaiDate(d, thaiDigits) {
      return new Intl.DateTimeFormat(thaiDigits ? 'th-TH-u-nu-thai' : 'th-TH', { day: 'numeric', month: 'long', year: 'numeric' }).format(d);
    }
    function insertAtCursor(text) {
      const r = ed.useRange(ed.range());
      quill.updateContents(new Delta().retain(r.index).delete(r.length).insert(text, quill.getFormat(r.index, 0)), 'user');
      quill.setSelection(r.index + text.length, 0, 'user');
    }

    // ================= ตรวจคำ (ออฟไลน์) =================
    // คำที่มักเขียนผิด → คำที่ถูก (อ้างอิงพจนานุกรมฉบับราชบัณฑิตยสถาน และคำทับศัพท์ตามหลักเกณฑ์ราชบัณฑิตยสภา)
    // ใส่เป็นคำเต็ม เพื่อไม่ให้ไปจับส่วนหนึ่งของคำที่ถูก (เช่น "กฏ" อยู่ใน "ปรากฏ")
    const MISSPELL = {
      'อนุญาติ': 'อนุญาต', 'กฏหมาย': 'กฎหมาย', 'กฏเกณฑ์': 'กฎเกณฑ์', 'กฏระเบียบ': 'กฎระเบียบ', 'กฏกระทรวง': 'กฎกระทรวง',
      'ปรากฎ': 'ปรากฏ', 'ปฎิ': 'ปฏิ', 'โอกาศ': 'โอกาส', 'ประสพการณ์': 'ประสบการณ์', 'ผลลัพท์': 'ผลลัพธ์', 'สังเกตุ': 'สังเกต',
      'เกณท์': 'เกณฑ์', 'คำนวน': 'คำนวณ', 'จำนวณ': 'จำนวน', 'ประมาน': 'ประมาณ', 'ทรมาร': 'ทรมาน', 'รสชาด': 'รสชาติ',
      'จราจล': 'จลาจล', 'ศรีษะ': 'ศีรษะ', 'ผลัดวันประกันพรุ่ง': 'ผัดวันประกันพรุ่ง', 'ลายเซ็นต์': 'ลายเซ็น', 'เซ็นต์ชื่อ': 'เซ็นชื่อ',
      'เซ็นต์สัญญา': 'เซ็นสัญญา', 'อาไหล่': 'อะไหล่', 'ทะแยง': 'ทแยง', 'เครื่องสำอางค์': 'เครื่องสำอาง', 'อานิสงค์': 'อานิสงส์',
      'ผาสุข': 'ผาสุก', 'โล่ห์': 'โล่', 'กิติมศักดิ์': 'กิตติมศักดิ์', 'ประณีประนอม': 'ประนีประนอม', 'ประนีต': 'ประณีต',
      'ทะเลสาป': 'ทะเลสาบ', 'บิณฑบาตร': 'บิณฑบาต', 'ปิกนิค': 'ปิกนิก', 'กงศุล': 'กงสุล', 'ภาพยนต์': 'ภาพยนตร์', 'รถยนตร์': 'รถยนต์',
      'อัธยาศรัย': 'อัธยาศัย', 'อนุเสาวรีย์': 'อนุสาวรีย์', 'ผูกพันธ์': 'ผูกพัน', 'ประสิทธิ์ภาพ': 'ประสิทธิภาพ', 'ไวยกรณ์': 'ไวยากรณ์',
      'ยานพาหะนะ': 'ยานพาหนะ', 'งบประมาน': 'งบประมาณ', 'ขมักเขม้น': 'ขะมักเขม้น', 'ลำใย': 'ลำไย', 'เลือกสรรค์': 'เลือกสรร',
      'มัคคุเทศน์': 'มัคคุเทศก์', 'สังสรรณ์': 'สังสรรค์', 'อุปสรรค์': 'อุปสรรค', 'ทรัพยากรณ์': 'ทรัพยากร', 'ลดหลั่ง': 'ลดหลั่น',
      'สมดุลย์': 'สมดุล', 'งบดุลย์': 'งบดุล', 'กระทันหัน': 'กะทันหัน', 'ประกาศนียบัติ': 'ประกาศนียบัตร', 'เกียรติ์': 'เกียรติ',
      'ผลประโยชณ์': 'ผลประโยชน์', 'จัดสรรค์': 'จัดสรร', 'อิสสระ': 'อิสระ', 'ลักษนะ': 'ลักษณะ', 'เกล็ดความรู้': 'เกร็ดความรู้',
      'กระเพรา': 'กะเพรา', 'ไอศครีม': 'ไอศกรีม', 'คลีนิก': 'คลินิก', 'คลินิค': 'คลินิก', 'คริสตมาส': 'คริสต์มาส', 'โควต้า': 'โควตา',
      'ซีรี่ย์': 'ซีรีส์', 'ซีรี่ส์': 'ซีรีส์', 'อีเมล์': 'อีเมล', 'เว็ปไซต์': 'เว็บไซต์', 'เว็บไซท์': 'เว็บไซต์', 'เวปไซต์': 'เว็บไซต์',
      'ลิ้งค์': 'ลิงก์', 'ลิงค์': 'ลิงก์', 'คอมเม้นท์': 'คอมเมนต์', 'คอมเมนท์': 'คอมเมนต์', 'ออฟฟิส': 'ออฟฟิศ', 'กราฟฟิก': 'กราฟิก',
      'เบรค': 'เบรก', 'โน๊ตบุ๊ค': 'โน้ตบุ๊ก', 'โน๊ต': 'โน้ต', 'แอพพลิเคชั่น': 'แอปพลิเคชัน', 'แอปพลิเคชั่น': 'แอปพลิเคชัน',
      'ดาวโหลด': 'ดาวน์โหลด', 'อัพเดท': 'อัปเดต', 'อัพเดต': 'อัปเดต', 'อัปเดท': 'อัปเดต', 'อัพโหลด': 'อัปโหลด', 'เซิฟเวอร์': 'เซิร์ฟเวอร์',
      'ซอฟท์แวร์': 'ซอฟต์แวร์', 'โพสท์': 'โพสต์', 'สเปรดชีท': 'สเปรดชีต', 'พรีเซนท์': 'พรีเซนต์', 'เปอร์เซนต์': 'เปอร์เซ็นต์',
      'อิเลคทรอนิกส์': 'อิเล็กทรอนิกส์', 'อิเล็คทรอนิกส์': 'อิเล็กทรอนิกส์', 'อิเลกทรอนิกส์': 'อิเล็กทรอนิกส์', 'ช็อคโกแลต': 'ช็อกโกแลต',
      'นะค่ะ': 'นะคะ', 'ไหมค่ะ': 'ไหมคะ', 'หรือเปล่าค่ะ': 'หรือเปล่าคะ', 'บรรยากาศ์': 'บรรยากาศ', 'สัมนา': 'สัมมนา', 'ผู้บริหาญ': 'ผู้บริหาร',
    };
    const reEsc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const MISSPELL_RE = new RegExp(Object.keys(MISSPELL).sort((a, b) => b.length - a.length).map(reEsc).join('|'), 'g');
    // กฎการเขียน: [ชื่อ, regex, ฟังก์ชันคำแนะนำ, คำอธิบาย]
    const RULES = [
      ['ไม้ยมก', /([^\s\n])ๆ/g, (m, a) => `${a} ๆ`, 'เว้นวรรคหน้าไม้ยมก (ๆ) ตามหลักราชบัณฑิตยสถาน'],
      ['ไม้ยมก', /ๆ([^\s\n.,)”"'])/g, (m, a) => `ๆ ${a}`, 'เว้นวรรคหลังไม้ยมก (ๆ)'],
      ['ฯลฯ', /([^\s\n])ฯลฯ/g, (m, a) => `${a} ฯลฯ`, 'เว้นวรรคหน้า ฯลฯ'],
      ['สระแอ', /เเ/g, () => 'แ', 'พิมพ์ "เ" สองตัวแทนสระ "แ" — ค้นหาและจัดเรียงข้อความจะผิด'],
      ['สระอำ', /ํา/g, () => 'ำ', 'นิคหิต + สระอา (มักมาจากการคัดลอก PDF) แทนสระอำ'],
      ['วรรณยุกต์ซ้ำ', /([่-๋])\1+/g, (m, a) => a, 'วรรณยุกต์ซ้ำกัน'],
      ['สระซ้ำ', /([ัิ-ฺ็])\1+/g, (m, a) => a, 'สระบน/ล่างซ้ำกัน'],
      ['เว้นวรรคซ้อน', / {2,}/g, () => ' ', 'เว้นวรรคติดกันหลายตัว'],
      ['ช่องว่างท้ายบรรทัด', / +(?=\n)/g, () => '', 'มีช่องว่างเกินที่ท้ายย่อหน้า'],
      ['ช่องว่างหน้าเครื่องหมาย', / +([,)\]])/g, (m, a) => a, 'ไม่ต้องเว้นวรรคหน้าเครื่องหมายนี้'],
      ['ห้ามตัดบรรทัด', /([0-9๐-๙]) (บาท|สตางค์|คน|ราย|วัน|เดือน|ปี|ชั่วโมง|นาที|แห่ง|ฉบับ|หน้า|ครั้ง|ชุด|เครื่อง|หลัง|กิโลเมตร|เมตร|ล้าน|เปอร์เซ็นต์|ข้อ|อัตรา|โครงการ)/g,
        (m, a, b) => `${a}\u00A0${b}`, 'ตัวเลขกับหน่วยอาจถูกตัดแยกคนละบรรทัด — ใช้เว้นวรรคห้ามตัดบรรทัด'],
    ];

    // ---------- ตรวจรูปแบบหนังสือราชการ ----------
    const MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
    const MONTHS_ABBR = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
    const MONTH_RE = MONTHS.join('|');
    const ABBR_RE = MONTHS_ABBR.map(reEsc).join('|');
    // เขียนตัวเลขตามแบบที่ใช้ในข้อความต้นฉบับ (เลขไทยหรืออารบิก)
    const sameDigits = (sample, n) => (/[๐-๙]/.test(sample) ? toThai(String(n)) : String(n));
    function thaiDateText(d, m, y, sample) {
      let year = +toArabic(String(y));
      if (year < 100) year += 2500;                       // ปี 2 หลักถือเป็น พ.ศ. เช่น 69 → 2569
      else if (year < 2400) year += 543;                  // ค.ศ. → พ.ศ.
      return `${sameDigits(sample, +toArabic(String(d)))} ${MONTHS[m - 1]} ${sameDigits(sample, year)}`;
    }
    const DATE_RULES = [
      [/(?<![0-9๐-๙/])([0-9๐-๙]{1,2})\/([0-9๐-๙]{1,2})\/([0-9๐-๙]{2,4})(?![0-9๐-๙/])/g, (m, d, mo, y) => {
        const mm = +toArabic(mo);
        return mm >= 1 && mm <= 12 && +toArabic(d) >= 1 && +toArabic(d) <= 31 ? thaiDateText(d, mm, y, m) : null;
      }, 'หนังสือราชการใช้วันที่แบบ “วัน เดือนเต็ม ปี พ.ศ.” เช่น ๒๙ กันยายน ๒๕๖๙'],
      [new RegExp(`(?<![0-9๐-๙])([0๐])([1-9๑-๙]) (${MONTH_RE})`, 'g'), (m, z, d, mo) => `${d} ${mo}`, 'วันที่ไม่ต้องมีเลขศูนย์นำหน้า'],
      [new RegExp(`(${MONTH_RE}) ((?:19|20|๑๙|๒๐)[0-9๐-๙]{2})(?![0-9๐-๙])`, 'g'), (m, mo, y) => `${mo} ${sameDigits(y, +toArabic(y) + 543)}`, 'ปีเป็น ค.ศ. — หนังสือราชการใช้ พ.ศ.'],
      [new RegExp(`([0-9๐-๙]{1,2}) (${ABBR_RE}) ?([0-9๐-๙]{2,4})`, 'g'), (m, d, ab, y) => thaiDateText(d, MONTHS_ABBR.indexOf(ab) + 1, y, m), 'ในหนังสือราชการควรเขียนชื่อเดือนเต็ม'],
    ];
    const mode = map => [...map.entries()].sort((a, b) => b[1].n - a[1].n)[0]?.[0];
    function scanFormat(text) {
      const found = [];
      const D = ed.settings().defaults;
      // วันที่
      for (const [re, fix, note] of DATE_RULES) {
        re.lastIndex = 0;
        for (let m = re.exec(text); m; m = re.exec(text)) {
          const suggest = fix(...m);
          if (suggest && suggest !== m[0]) found.push({ index: m.index, wrong: m[0], suggest, kind: 'วันที่', note });
        }
      }
      // ย่อหน้าเนื้อความ: ย่อหน้าบรรทัดแรกและระยะบรรทัดควรเท่ากันทั้งฉบับ
      const indents = new Map(), heights = new Map();
      const add = (map, key, line) => { const e = map.get(key) || { n: 0, lines: [] }; e.n++; e.lines.push(line); map.set(key, e); };
      for (const line of quill.getLines()) {
        if (line.statics.blotName !== 'block') continue;
        const len = line.length() - 1;
        if (len < 40) continue;                           // ข้ามบรรทัดสั้น เช่น หัวข้อ เรื่อง เรียน ลงชื่อ
        const f = line.formats();
        if ((f.align || D.align) === 'center' || f.marginleft) continue;
        const ind = f.textindent ?? D.indent;
        if (parseFloat(ind) > 0) add(indents, ind, line);
        add(heights, f.lineheight ?? D.lineHeight, line);
      }
      const uneven = (map, kind, attr, label) => {
        if (map.size < 2) return;
        const major = mode(map);
        for (const [val, e] of map) {
          if (val === major) continue;
          for (const line of e.lines) {
            const index = quill.getIndex(line);
            found.push({
              index, kind, select: { index, length: line.length() - 1 },
              note: `${label} ${val} ต่างจากย่อหน้าส่วนใหญ่ (${major}) — “${quill.getText(index, 24).trim()}…”`,
              apply: () => quill.formatLine(quill.getIndex(line), 1, attr, major, 'user'),
            });
          }
        }
      };
      uneven(indents, 'ย่อหน้าไม่เท่ากัน', 'textindent', 'ย่อหน้าบรรทัดแรก');
      uneven(heights, 'ระยะบรรทัดไม่เท่ากัน', 'lineheight', 'ระยะบรรทัด');
      // ขนาดและแบบอักษรของเนื้อความ (ไม่นับตัวหนาซึ่งมักเป็นหัวข้อ)
      let pos = 0;
      const odd = [];
      for (const op of quill.getContents().ops) {
        const len = typeof op.insert === 'string' ? op.insert.length : 1;
        const a = op.attributes || {};
        if (typeof op.insert === 'string' && op.insert.trim() && !a.bold && !a.table) {
          if (a.size && a.size !== D.size) odd.push({ index: pos, length: len, kind: 'ขนาดอักษร', note: `ขนาด ${a.size} ต่างจากเนื้อความ ${D.size}`, fmt: { size: false } });
          if (a.font && a.font !== D.font) odd.push({ index: pos, length: len, kind: 'แบบอักษร', note: `แบบอักษร ${a.font} ปนกับ ${D.font}`, fmt: { font: false } });
        }
        pos += len;
      }
      for (const o of odd) {
        found.push({
          index: o.index, kind: o.kind, select: { index: o.index, length: o.length },
          note: `${o.note} — “${quill.getText(o.index, Math.min(o.length, 24)).trim()}”`,
          apply: () => quill.formatText(o.index, o.length, o.fmt, 'user'),
        });
      }
      // คำขึ้นต้น–คำลงท้าย
      const open = /^(กราบเรียน|เรียน)[ \u00A0]/m.exec(text), close = /^\s*(ขอแสดงความนับถืออย่างยิ่ง|ขอแสดงความนับถือ)\s*$/m.exec(text);
      if (open && close) {
        const want = open[1] === 'กราบเรียน' ? 'ขอแสดงความนับถืออย่างยิ่ง' : 'ขอแสดงความนับถือ';
        if (close[1] !== want) {
          const at = close.index + close[0].indexOf(close[1]);
          found.push({ index: at, wrong: close[1], suggest: want, kind: 'คำขึ้นต้น–ลงท้าย', note: `ขึ้นต้นด้วย “${open[1]}” ควรลงท้ายด้วย “${want}”` });
        }
      }
      return found;
    }

    function scanOffline() {
      const text = docText(quill.getContents());
      const found = [];
      for (let m = (MISSPELL_RE.lastIndex = 0, MISSPELL_RE.exec(text)); m; m = MISSPELL_RE.exec(text)) {
        found.push({ index: m.index, wrong: m[0], suggest: MISSPELL[m[0]], kind: 'สะกด', note: `ที่ถูกต้องคือ “${MISSPELL[m[0]]}”` });
      }
      const allowBig = $('#allowBigSpace')?.checked;
      for (const [kind, re, fix, note] of RULES) {
        re.lastIndex = 0;
        for (let m = re.exec(text); m; m = re.exec(text)) {
          if (kind === 'เว้นวรรคซ้อน' && allowBig && m[0].length === 2) continue;
          found.push({ index: m.index, wrong: m[0], suggest: fix(...m), kind, note });
        }
      }
      found.push(...scanFormat(text));
      const thai = (text.match(/[๐-๙]/g) || []).length, arabic = (text.match(/[0-9]/g) || []).length;
      found.sort((a, b) => a.index - b.index);
      return { found, mixedDigits: thai && arabic ? { thai, arabic } : null };
    }
    // หาตำแหน่งปัจจุบันของรายการอีกครั้ง (ข้อความอาจถูกแก้ไปแล้ว) โดยเลือกตำแหน่งที่ใกล้ของเดิมที่สุด
    function locate(item) {
      const text = docText(quill.getContents());
      let best = -1;
      for (let i = text.indexOf(item.wrong); i !== -1; i = text.indexOf(item.wrong, i + 1)) {
        if (best === -1 || Math.abs(i - item.index) < Math.abs(best - item.index)) best = i;
      }
      return best;
    }
    function applyFix(item) {
      if (item.apply) { item.apply(); return true; }          // รายการที่แก้รูปแบบ (ไม่ได้แก้ข้อความ)
      const i = locate(item);
      if (i < 0) return false;
      const contents = quill.getContents();
      let change = new Delta().retain(i).delete(item.wrong.length);
      if (item.suggest) change = change.insert(item.suggest, attrsAt(contents, i));
      quill.updateContents(change, 'user');
      return true;
    }
    function selectItem(item) {
      if (item.select) { ed.useRange(item.select); return; }
      const i = locate(item);
      if (i < 0) { say('ไม่พบข้อความนี้แล้ว (อาจถูกแก้ไปแล้ว)'); return; }
      ed.useRange({ index: i, length: item.wrong.length });
      quill.scrollSelectionIntoView?.();
    }

    // ================= ผู้ช่วย AI (Claude) =================
    const AI_MODEL = 'claude-opus-5-5';
    const AI_MAX_CHARS = 30000;
    let apiKey = store.get(KEY.apiKey) || '';
    let Anthropic = null;
    // SDK ของ Anthropic: ใช้ไฟล์ที่ล็อกเวอร์ชันไว้ในโปรเจกต์ (vendor/) ไม่โหลดจาก CDN เพราะโค้ดที่โหลดมาจะรันในหน้าที่มีคีย์ AI ของผู้ใช้
    // ชื่อไฟล์มีเลขเวอร์ชัน (เปลี่ยนเวอร์ชันต้องแก้ที่นี่ และให้ build-single.mjs ตรวจว่าไฟล์ตรงกัน)
    const SDK_FILE = './vendor/anthropic-sdk-0.131.0.mjs';
    async function loadSdk() {
      if (Anthropic) return Anthropic;
      // ไฟล์เดี่ยวสำหรับส่งทางอีเมล: SDK ฝังมาในหน้า (เปิดแบบ file:// import ไฟล์ข้างเคียงไม่ได้) → โหลดผ่าน Blob
      const embedded = document.getElementById('sdk-embedded')?.textContent;
      if (embedded) {
        const url = URL.createObjectURL(new Blob([embedded], { type: 'text/javascript' }));
        // ถ้าหน้านี้ถูกเสิร์ฟจากเว็บที่มี CSP (ห้าม blob:) การโหลดจะล้มเหลว → ใช้ไฟล์ vendor/ ข้างเคียงแทน
        try { Anthropic = (await import(/* @vite-ignore */ url)).default; } catch { /* ลองไฟล์ข้างเคียงด้านล่าง */ } finally { URL.revokeObjectURL(url); }
      }
      if (!Anthropic) Anthropic = (await import(/* @vite-ignore */ SDK_FILE)).default;
      return Anthropic;
    }
    const SYSTEM = 'คุณเป็นผู้เชี่ยวชาญภาษาไทยและงานสารบรรณของหน่วยงานราชการไทย ยึดการสะกดตามพจนานุกรมฉบับราชบัณฑิตยสถาน '
      + 'หลักการเขียนคำทับศัพท์ และระเบียบสำนักนายกรัฐมนตรีว่าด้วยงานสารบรรณ ตอบเป็นภาษาไทย กระชับ และไม่เปลี่ยนสาระของต้นฉบับ';
    const TASKS = {
      check: {
        label: 'ตรวจคำผิด ไวยากรณ์ และการเว้นวรรค',
        prompt: 'ตรวจข้อความต่อไปนี้ แล้วระบุเฉพาะจุดที่ผิดจริง ได้แก่ คำสะกดผิด การใช้คำผิดความหมาย ไวยากรณ์ การเว้นวรรค และคำที่ไม่เหมาะกับหนังสือราชการ '
          + 'ในช่อง wrong ให้คัดลอกข้อความผิดจากต้นฉบับตรงทุกตัวอักษร (สั้นที่สุดที่ยังระบุตำแหน่งได้) และ suggest คือข้อความที่ใช้แทน ถ้าไม่มีจุดผิดให้ส่ง items ว่าง',
        schema: {
          type: 'object', additionalProperties: false, required: ['items'],
          properties: {
            items: {
              type: 'array',
              items: {
                type: 'object', additionalProperties: false, required: ['wrong', 'suggest', 'kind', 'reason'],
                properties: {
                  wrong: { type: 'string' },
                  suggest: { type: 'string' },
                  kind: { type: 'string', enum: ['สะกด', 'ใช้คำ', 'ไวยากรณ์', 'เว้นวรรค', 'ภาษาราชการ'] },
                  reason: { type: 'string' },
                },
              },
            },
          },
        },
      },
      formal: {
        label: 'เรียบเรียงเป็นภาษาราชการ',
        prompt: 'เรียบเรียงข้อความต่อไปนี้ใหม่เป็นภาษาหนังสือราชการที่สุภาพ กระชับ ถูกต้อง คงสาระ ตัวเลข ชื่อเฉพาะ และการแบ่งย่อหน้าเดิม (คั่นย่อหน้าด้วยบรรทัดใหม่) และอธิบายสั้น ๆ ว่าปรับอะไรบ้าง',
        schema: {
          type: 'object', additionalProperties: false, required: ['rewritten', 'notes'],
          properties: { rewritten: { type: 'string' }, notes: { type: 'string' } },
        },
      },
      layout: {
        label: 'วางแผนจัดหน้าให้สอดคล้องกัน',
        prompt: 'ต่อไปนี้คือรายการย่อหน้าของหนังสือราชการในรูป JSON ที่วัดจากหน้าจอจริง '
          + '(page = หน้า, n = ลำดับย่อหน้า, role = บทบาทที่ระบบเดา, text = ต้นข้อความ, formats = ค่าที่ตั้งไว้ ถ้าไม่มีใช้ standard, ls = ช่องไฟ em ปัจจุบัน). '
          + 'metrics: lines = จำนวนบรรทัด, stretch = ระยะที่ตัวอักษรถูกยืดเพิ่มต่อตัว (em) ในบรรทัดที่ยืดมากที่สุด จากการจัดเสมอหน้าหลังแบบกระจายระหว่างตัวอักษร (เกิน 0.08 = ห่างผิดสังเกต), '
          + 'stretchLine = บรรทัดที่ยืดมากที่สุด, stretchFix = ช่องไฟ (em) ที่ระบบทดลองแล้วว่าลดการยืดได้โดยไม่เพิ่มบรรทัด, '
          + 'lastFill = ความยาวบรรทัดสุดท้ายเทียบความกว้าง (ต่ำกว่า 0.2 = มีคำโดด), orphanFix = ช่องไฟ (em) ที่ระบบทดลองแล้วว่าดึงคำโดดขึ้นบรรทัดบนได้, '
          + 'badBreaks = จุดตัดบรรทัดที่ผิดหลัก (at = ท้ายบรรทัด|ต้นบรรทัดถัดไป, phrase = วลีที่ผูกด้วยเว้นวรรคห้ามตัดได้, join = คำที่ถูกตัดกลางคำ). '
          + 'nobreakDictionary = คำห้ามตัดที่ผู้ใช้กำหนด (ระบบผูกด้วย Word Joiner ให้ได้).\n'
          + 'ให้วิเคราะห์โดยให้น้ำหนักมากที่สุดตามลำดับนี้:\n'
          + '1) การตัดคำภาษาไทยท้ายบรรทัด: ห้ามขึ้นบรรทัดใหม่ด้วย ๆ ฯ สระหรือวรรณยุกต์ ห้ามแยกตัวเลขกับหน่วย คำนำหน้ากับชื่อ ชื่อกับนามสกุล เลขที่หนังสือ วันเดือนปี '
          + 'และห้ามตัดกลางคำประสม ชื่อหน่วยงาน ชื่อคน (เช่น ผู้อำนวย⏎การ) — แก้ด้วย bind (วลีที่มีเว้นวรรค ผูกด้วยเว้นวรรคห้ามตัด) '
          + 'หรือ nobreak (คำไม่มีเว้นวรรคที่ต้องอยู่บรรทัดเดียวกัน ระบบจะผูกด้วย Word Joiner) ทั้งสองอย่างต้องปรากฏในข้อความตรงทุกตัวอักษร.\n'
          + '2) การกระจายตัวอักษรภาษาไทยให้สวยงาม: การจัดเสมอหน้าหลังกระจายช่องว่างระหว่างตัวอักษร (text-justify: inter-character) ย่อหน้าที่ stretch สูงหรือมีคำโดดบรรทัดสุดท้าย '
          + 'ให้ปรับ letterspacing เล็กน้อย (ระหว่าง -0.03em ถึง +0.03em ใช้ stretchFix/orphanFix เป็นหลักถ้ามี) หรือพิจารณา distribute/justify/left ให้เหมาะกับบทบาท '
          + 'และให้ย่อหน้าเนื้อความใช้ช่องไฟใกล้เคียงกันทั้งฉบับ.\n'
          + '3) ระยะบรรทัด: เนื้อความทุกย่อหน้าใช้ระยะบรรทัดเดียวกันตาม standard ส่วนหัว รายการ และลงชื่อใช้ค่าที่สอดคล้องกัน และใช้ spacebefore/spaceafter แยกส่วนแทนบรรทัดว่างเมื่อเหมาะสม.\n'
          + 'เรื่องอื่น (ย่อหน้าบรรทัดแรก กั้นซ้าย การจัดแนวตามบทบาท) ปรับให้ถูกระเบียบงานสารบรรณเป็นลำดับรอง. '
          + 'เขียน plan สั้น ๆ เป็นข้อ ๆ ตามสามหัวข้อนี้ และใส่ใน items เฉพาะย่อหน้าที่ควรเปลี่ยน ใช้ null กับค่าที่ไม่เปลี่ยน. '
          + 'ความยาวใช้ cm หรือ pt (เช่น "2.5cm", "6pt"), ระยะบรรทัดใช้ % (เช่น "100%"), letterspacing ใช้ em (เช่น "-0.01em" หรือ "0em" = ปกติ), '
          + 'align เลือกจาก left center right justify distribute. ใน reason ให้บอกว่าแก้ปัญหาข้อไหน (ตัดคำ/กระจายตัวอักษร/ระยะบรรทัด). ห้ามแก้ถ้อยคำ',
        schema: {
          type: 'object', additionalProperties: false, required: ['plan', 'items'],
          properties: {
            plan: { type: 'string' },
            items: {
              type: 'array',
              items: {
                type: 'object', additionalProperties: false,
                required: ['page', 'n', 'role', 'align', 'textindent', 'marginleft', 'lineheight', 'spacebefore', 'spaceafter', 'letterspacing', 'bind', 'nobreak', 'reason'],
                properties: {
                  page: { type: 'integer' }, n: { type: 'integer' }, role: { type: 'string' },
                  align: { anyOf: [{ type: 'string', enum: ['left', 'center', 'right', 'justify', 'distribute'] }, { type: 'null' }] },
                  textindent: { anyOf: [{ type: 'string' }, { type: 'null' }] },
                  marginleft: { anyOf: [{ type: 'string' }, { type: 'null' }] },
                  lineheight: { anyOf: [{ type: 'string' }, { type: 'null' }] },
                  spacebefore: { anyOf: [{ type: 'string' }, { type: 'null' }] },
                  spaceafter: { anyOf: [{ type: 'string' }, { type: 'null' }] },
                  letterspacing: { anyOf: [{ type: 'string' }, { type: 'null' }] },
                  bind: { type: 'array', items: { type: 'string' } },
                  nobreak: { type: 'array', items: { type: 'string' } },
                  reason: { type: 'string' },
                },
              },
            },
          },
        },
      },
      rewrite: {
        label: 'เรียบเรียงภาษาราชการรายย่อหน้า',
        prompt: 'ต่อไปนี้คือย่อหน้าของหนังสือราชการในรูป JSON (n = ลำดับย่อหน้า, text = ข้อความทั้งย่อหน้า). '
          + 'เรียบเรียงเฉพาะย่อหน้าที่ถ้อยคำไม่เหมาะสมหรือไม่เป็นภาษาหนังสือราชการ ให้เป็นภาษาราชการที่สุภาพ กระชับ ถูกต้องตามหลักภาษา '
          + 'โดยต้องคงสาระ ข้อเท็จจริง วันที่ จำนวนเงิน ชื่อบุคคล ชื่อหน่วยงาน และ**ตัวเลขทุกตัว**ไว้เหมือนเดิมทุกตัวอักษร ห้ามเพิ่มข้อมูลใหม่ ห้ามรวมหรือแยกย่อหน้า. '
          + 'ย่อหน้าที่ดีอยู่แล้วไม่ต้องใส่ใน items. text ที่ส่งกลับต้องเป็นย่อหน้าเดียว ไม่มีการขึ้นบรรทัดใหม่ และ reason อธิบายสั้น ๆ ว่าปรับอะไร',
        schema: {
          type: 'object', additionalProperties: false, required: ['items'],
          properties: {
            items: {
              type: 'array',
              items: {
                type: 'object', additionalProperties: false, required: ['n', 'text', 'reason'],
                properties: { n: { type: 'integer' }, text: { type: 'string' }, reason: { type: 'string' } },
              },
            },
          },
        },
      },
      audit: {
        label: 'ตรวจโครงสร้างและความสอดคล้อง',
        prompt: 'ตรวจหนังสือราชการต่อไปนี้ (แต่ละหน้าคั่นด้วยบรรทัด ===== หน้า N =====) เพื่อหาข้อบกพร่องที่ผู้ตรวจต้องรู้ ได้แก่ '
          + '1) ความสอดคล้องของข้อมูล เช่น วันที่ในส่วนหัวกับเนื้อความไม่ตรงกัน จำนวนเงินตัวเลขกับตัวอักษรไม่ตรงกัน ชื่อตำแหน่งหรือชื่อหน่วยงานสะกดไม่เหมือนกันในฉบับ ตัวเลขข้อไม่ต่อเนื่อง '
          + '2) โครงสร้างหนังสือราชการ เช่น ขาดเรื่อง เรียน อ้างถึง ลงชื่อ คำขึ้นต้นกับคำลงท้ายไม่เข้าคู่ '
          + '3) ถ้อยคำที่ไม่เหมาะกับหนังสือราชการ. รายงานเฉพาะที่พบจริง ไม่ต้องแก้ไขข้อความ ถ้าไม่พบให้ notes ว่าง. summary คือสรุปคุณภาพโดยรวม 1–2 ประโยค. where ระบุหน้าและตำแหน่งโดยย่อ',
        schema: {
          type: 'object', additionalProperties: false, required: ['summary', 'notes'],
          properties: {
            summary: { type: 'string' },
            notes: {
              type: 'array',
              items: {
                type: 'object', additionalProperties: false, required: ['severity', 'where', 'issue', 'suggestion'],
                properties: { severity: { type: 'string', enum: ['สูง', 'กลาง', 'ต่ำ'] }, where: { type: 'string' }, issue: { type: 'string' }, suggestion: { type: 'string' } },
              },
            },
          },
        },
      },
      summary: {
        label: 'สรุปใจความสำคัญ',
        prompt: 'สรุปใจความสำคัญของข้อความต่อไปนี้เป็นข้อ ๆ ไม่เกิน 5 ข้อ แต่ละข้อสั้น กระชับ',
        schema: {
          type: 'object', additionalProperties: false, required: ['points'],
          properties: { points: { type: 'array', items: { type: 'string' } } },
        },
      },
      spacing: {
        label: 'วิเคราะห์ช่องไฟตัวอักษร',
        prompt: 'ต่อไปนี้คือข้อมูลย่อหน้าของหนังสือราชการภาษาไทยในรูป JSON ที่วัดจากหน้าจอจริง (n = ลำดับย่อหน้า, ls = ช่องไฟปัจจุบันเป็น em หรือ null ถ้าปนกัน, '
          + 'metrics.stretch = ระยะที่ตัวอักษรถูกยืดเพิ่มต่อตัวเมื่อจัดเสมอหน้าหลังในบรรทัดที่ยืดมากที่สุด, stretchFix = ช่องไฟที่ระบบทดลองแล้วว่าลดการยืดได้โดยไม่เพิ่มจำนวนบรรทัด, '
          + 'lastFill = ความยาวบรรทัดสุดท้ายเทียบความกว้าง, orphanFix = ช่องไฟที่ทดลองแล้วว่าดึงคำโดดกลับขึ้นบรรทัดบน) และ focus = รายการ n ที่ผู้ใช้ต้องการให้วิเคราะห์ '
          + '(ถ้า focus ว่าง = ทุกย่อหน้า). ให้วิเคราะห์ช่องไฟตัวอักษรเฉพาะย่อหน้าใน focus ให้สวยงามตามหลักการจัดหน้าภาษาไทย: ลดช่องว่างที่ถูกยืดจนตัวอักษรห่างผิดปกติ, '
          + 'ไม่ให้สระ/วรรณยุกต์ที่ซ้อนกันชิดหรือห่างเกินไป, ให้ย่อหน้าเนื้อความที่ใช้ฟอนต์เดียวกันมีช่องไฟใกล้เคียงกัน, และไม่ให้เกิดคำโดดท้ายย่อหน้า '
          + 'ค่าที่เสนอต้องอยู่ระหว่าง -0.03em ถึง +0.05em (ใช้ stretchFix/orphanFix เป็นหลักถ้ามี) และ "0em" = ปกติ ใส่เฉพาะย่อหน้าที่ควรเปลี่ยนจริง '
          + 'summary = สรุปการวิเคราะห์สั้น ๆ 1-3 ประโยคเป็นภาษาไทย reason = เหตุผลสั้น ๆ ต่อย่อหน้า',
        schema: {
          type: 'object', additionalProperties: false, required: ['summary', 'items'],
          properties: {
            summary: { type: 'string' },
            items: { type: 'array', items: { type: 'object', additionalProperties: false, required: ['n', 'letterspacing', 'reason'], properties: { n: { type: 'integer' }, letterspacing: { type: 'string' }, reason: { type: 'string' } } } },
          },
        },
      },
    };
    // ผู้ให้บริการ AI: Claude ใช้ SDK ทางการของ Anthropic ส่วนเจ้าอื่นเรียก REST API ของแต่ละเจ้า
    // ชื่อรุ่นของ OpenAI/Gemini/เซิร์ฟเวอร์อื่นแก้ได้ในแผง เพราะแต่ละเจ้าเปลี่ยนรุ่นบ่อย
    const PROVIDERS = {
      claude: { label: 'Claude (Anthropic)', keyHint: 'Anthropic API key (sk-ant-…)', model: AI_MODEL, fixedModel: true },
      openai: { label: 'ChatGPT (OpenAI)', keyHint: 'OpenAI API key (sk-…)', model: 'gpt-4.1-mini' },
      gemini: { label: 'Gemini (Google)', keyHint: 'Google AI Studio API key', model: 'gemini-2.5-flash' },
      custom: { label: 'เซิร์ฟเวอร์แบบ OpenAI-compatible (Ollama, Typhoon ฯลฯ)', keyHint: 'API key (ถ้ามี)', model: 'llama3.1', baseUrl: 'http://localhost:11434/v1', keyOptional: true },
    };
    const aiCfg = (() => { try { return JSON.parse(store.get(KEY.aiCfg) || '{}') || {}; } catch { return {}; } })();
    aiCfg.models ||= {};
    let provider = PROVIDERS[aiCfg.provider] ? aiCfg.provider : 'claude';
    const modelOf = p => (PROVIDERS[p].fixedModel ? PROVIDERS[p].model : aiCfg.models?.[p] || PROVIDERS[p].model);
    const baseUrlOf = () => (aiCfg.baseUrl || PROVIDERS.custom.baseUrl).replace(/\/+$/, '');
    const keyStoreKey = p => `${KEY.apiKey}:${p}`;
    const keys = {};
    for (const p of Object.keys(PROVIDERS)) keys[p] = store.get(keyStoreKey(p)) || '';
    if (!keys.claude && apiKey) keys.claude = apiKey;           // ย้ายคีย์ Claude ที่เคยจำไว้แบบเก่า
    const isLocalUrl = u => { try { return ['localhost', '127.0.0.1', '[::1]'].includes(new URL(u).hostname); } catch { return false; } };
    const staysLocal = () => provider === 'custom' && isLocalUrl(baseUrlOf());

    class HttpError extends Error {
      constructor(status, body) { super(`HTTP ${status}`); this.status = status; this.body = body; }
    }
    // อ่าน JSON จากคำตอบแบบผ่อนปรน (บางรุ่นใส่ ```json หรือข้อความนำหน้ามา)
    function parseJsonLoose(text) {
      const t = String(text || '').replace(/```(?:json)?/gi, '').trim();
      try { return JSON.parse(t); } catch { /* ลองตัดเฉพาะวัตถุ JSON */ }
      const a = t.indexOf('{'), b = t.lastIndexOf('}');
      if (a >= 0 && b > a) return JSON.parse(t.slice(a, b + 1));
      throw new SyntaxError('no json');
    }
    const jsonInstruction = schema => `\n\nตอบเป็น JSON อย่างเดียว ตามโครงสร้าง JSON Schema นี้:\n${JSON.stringify(schema)}`;
    async function postJson(url, headers, body) {
      const res = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json', ...headers }, body: JSON.stringify(body) });
      if (!res.ok) throw new HttpError(res.status, await res.text().catch(() => ''));
      return res.json();
    }
    async function askOpenAiLike(url, key, model, task, userText, jsonMode) {
      const data = await postJson(url, key ? { Authorization: `Bearer ${key}` } : {}, {
        model,
        messages: [
          { role: 'system', content: SYSTEM + jsonInstruction(task.schema) },
          { role: 'user', content: userText },
        ],
        ...(jsonMode ? { response_format: { type: 'json_object' } } : {}),
      });
      if (data.choices?.[0]?.finish_reason === 'length') throw new Error('ข้อความยาวเกินกว่าที่ AI ตอบได้ครบ — ลองเลือกทีละส่วน');
      return parseJsonLoose(data.choices?.[0]?.message?.content);
    }
    async function askGemini(key, model, task, userText) {
      const data = await postJson(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`, { 'x-goog-api-key': key }, {
        systemInstruction: { parts: [{ text: SYSTEM + jsonInstruction(task.schema) }] },
        contents: [{ role: 'user', parts: [{ text: userText }] }],
        generationConfig: { responseMimeType: 'application/json' },
      });
      const cand = data.candidates?.[0];
      if (!cand) throw new Error('Gemini ไม่ตอบกลับ' + (data.promptFeedback?.blockReason ? ` (ถูกบล็อก: ${data.promptFeedback.blockReason})` : ''));
      if (cand.finishReason === 'MAX_TOKENS') throw new Error('ข้อความยาวเกินกว่าที่ AI ตอบได้ครบ — ลองเลือกทีละส่วน');
      return parseJsonLoose((cand.content?.parts || []).map(p => p.text || '').join(''));
    }
    async function askAi(taskKey, text) {
      const task = TASKS[taskKey];
      const userText = `${task.prompt}\n\n<ข้อความ>\n${text}\n</ข้อความ>`;
      if (provider === 'openai') return askOpenAiLike('https://api.openai.com/v1/chat/completions', keys.openai, modelOf('openai'), task, userText, true);
      if (provider === 'gemini') return askGemini(keys.gemini, modelOf('gemini'), task, userText);
      if (provider === 'custom') return askOpenAiLike(`${baseUrlOf()}/chat/completions`, keys.custom, modelOf('custom'), task, userText, false);
      return askClaude(task, text);
    }
    async function askClaude(task, text) {
      const Sdk = await loadSdk();
      const client = new Sdk({ apiKey: keys.claude, dangerouslyAllowBrowser: true });
      const response = await client.beta.messages.create({
        model: AI_MODEL,
        max_tokens: 16000,
        betas: ['server-side-fallback-2026-07-01'],
        fallbacks: 'default',
        system: SYSTEM,
        output_config: { effort: 'medium', format: { type: 'json_schema', schema: task.schema } },
        messages: [{ role: 'user', content: `${task.prompt}\n\n<ข้อความ>\n${text}\n</ข้อความ>` }],
      });
      if (response.stop_reason === 'refusal') throw new Error('AI ปฏิเสธคำขอนี้' + (response.stop_details?.explanation ? `: ${response.stop_details.explanation}` : ''));
      if (response.stop_reason === 'max_tokens') throw new Error('ข้อความยาวเกินกว่าที่ AI ตอบได้ครบ — ลองเลือกทีละส่วน');
      const out = response.content.filter(b => b.type === 'text').map(b => b.text).join('');
      return JSON.parse(out);
    }
    function aiErrorText(err) {
      const A = Anthropic;
      if (A && err instanceof A.AuthenticationError) return 'API key ไม่ถูกต้อง';
      if (A && err instanceof A.PermissionDeniedError) return 'API key นี้ไม่มีสิทธิ์ใช้โมเดลนี้';
      if (A && err instanceof A.RateLimitError) return 'ใช้งานถี่เกินไป — รอสักครู่แล้วลองใหม่';
      if (A && err instanceof A.APIConnectionError) return 'เชื่อมต่อ Anthropic ไม่ได้ — ตรวจอินเทอร์เน็ต';
      if (A && err instanceof A.APIError) return `ระบบ AI ตอบกลับผิดพลาด (${err.status ?? '-'}) ${err.message}`;
      if (err instanceof HttpError) {
        if (err.status === 401 || err.status === 403) return `API key ไม่ถูกต้องหรือไม่มีสิทธิ์ (${err.status})`;
        if (err.status === 404) return `ไม่พบรุ่น “${modelOf(provider)}” หรือที่อยู่เซิร์ฟเวอร์ไม่ถูกต้อง — ตรวจชื่อรุ่น`;
        if (err.status === 429) return 'ใช้งานถี่เกินไปหรือโควตาหมด — รอสักครู่แล้วลองใหม่';
        return `ระบบ AI ตอบกลับผิดพลาด (${err.status}) ${String(err.body).slice(0, 200)}`;
      }
      if (err instanceof SyntaxError) return 'อ่านคำตอบของ AI ไม่ได้ — ลองใหม่ หรือเปลี่ยนรุ่นที่ตอบเป็น JSON ได้';
      if (err instanceof TypeError && provider === 'custom') return 'เชื่อมต่อเซิร์ฟเวอร์ไม่ได้ — ตรวจว่าเปิดเซิร์ฟเวอร์อยู่ และอนุญาต CORS แล้ว (Ollama: ตั้งค่า OLLAMA_ORIGINS=*)';
      if (err instanceof TypeError) return 'เชื่อมต่อไม่ได้ — ตรวจอินเทอร์เน็ต';
      return err?.message || String(err);
    }

    // ================= ธีม =================
    const THEMES = {
      white: { label: 'ขาว (ค่าเริ่มต้น)', paper: '#ffffff', ink: '#000000', page: '#eef0f3' },
      cream: { label: 'ครีมถนอมสายตา', paper: '#fbf4e2', ink: '#2b2620', page: '#e8e0cc' },
      green: { label: 'เขียวใบชา', paper: '#edf5e8', ink: '#1e2a1c', page: '#dbe6d4' },
      blue: { label: 'ฟ้านุ่ม', paper: '#eef3f9', ink: '#1b2430', page: '#dce4ee' },
      gray: { label: 'เทาลดแสง', paper: '#eceae6', ink: '#222222', page: '#d7d5d0' },
      night: { label: 'กลางคืน', paper: '#23272e', ink: '#e3e3e3', page: '#16191d' },
    };
    function applyTheme(key, warmth) {
      const t = THEMES[key] || THEMES.white;
      const root = document.documentElement.style;
      root.setProperty('--paper', t.paper);
      root.setProperty('--paper-ink', t.ink);
      root.setProperty('--page', t.page);
      $('.editor-box').style.filter = warmth > 0 ? `sepia(${warmth}%)` : '';
      for (const b of $$('.tt-theme')) b.setAttribute('aria-checked', String(b.dataset.theme === key));
    }

    // ================= แผงเครื่องมือ =================
    const css = document.createElement('style');
    css.textContent = `
.tools{position:fixed;top:0;right:0;bottom:0;z-index:30;width:360px;max-width:100vw;display:flex;flex-direction:column;background:#fff;border-left:1px solid var(--line);box-shadow:-4px 0 18px rgba(20,30,50,.12);font-size:13px}
.tools[hidden]{display:none}
body.tools-open .topbar,body.tools-open .desk,body.tools-open .statusbar{margin-right:360px}
@media (max-width:1100px){body.tools-open .topbar,body.tools-open .desk,body.tools-open .statusbar{margin-right:0}}
.tools-head{display:flex;flex-wrap:wrap;gap:6px;align-items:center;justify-content:space-between;padding:10px 12px;border-bottom:1px solid var(--line)}
.tools-head h2{margin:0;font-size:15px}
.tools-x{border:0;background:none;font-size:20px;line-height:1;cursor:pointer;color:var(--muted);padding:2px 6px}
.tools-tabs{display:flex;flex-wrap:wrap;gap:2px;padding:6px 8px 0;background:#f4f6f9;border-bottom:1px solid var(--line)}
.tools-tabs button{padding:6px 10px;border:1px solid transparent;border-bottom:0;border-radius:5px 5px 0 0;background:none;color:var(--muted);cursor:pointer;font:inherit}
.tools-tabs button[aria-selected="true"]{background:#fff;border-color:var(--line);color:var(--ink);font-weight:600;margin-bottom:-1px}
.tools-body{flex:1;overflow:auto;padding:12px}
.tools-body section[hidden]{display:none}
.tools h3{margin:14px 0 6px;font-size:13px}
.tools h3:first-child{margin-top:0}
.tools p.hint{margin:4px 0 8px;color:var(--muted);font-size:12px;line-height:1.5}
.tools .row{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin:6px 0}
.tools .tbtn{height:30px;padding:0 10px;border:1px solid var(--line-strong);border-radius:6px;background:#fff;cursor:pointer;font:inherit}
.tools .tbtn:hover{background:var(--hover)}
.tools .tbtn.primary{background:var(--accent);border-color:var(--accent);color:#fff}
.tools .tbtn:disabled{opacity:.5;cursor:default}
.tools input[type=text],.tools input[type=password],.tools select{height:30px;border:1px solid var(--line-strong);border-radius:6px;padding:0 8px;font:inherit;background:#fff;min-width:0}
.tools input.grow{flex:1}
.tools label.inline{display:inline-flex;align-items:center;gap:5px}
.tools .out{margin-top:6px;padding:8px 10px;border:1px solid var(--line);border-radius:6px;background:#f8f9fb;line-height:1.6;white-space:pre-wrap;word-break:break-word}
.tt-theme{display:flex;align-items:center;gap:8px;width:100%;padding:6px 8px;margin:3px 0;border:1px solid var(--line);border-radius:6px;background:#fff;cursor:pointer;font:inherit;text-align:left}
.tt-theme[aria-checked="true"]{border-color:var(--accent);box-shadow:0 0 0 1px var(--accent)}
.tt-theme i{width:34px;height:22px;border-radius:4px;border:1px solid rgba(0,0,0,.15);flex:none}
.find-list{list-style:none;margin:8px 0 0;padding:0}
.find-list li{padding:7px 8px;border:1px solid var(--line);border-radius:6px;margin-bottom:6px}
.find-list .k{display:inline-block;font-size:11px;padding:0 6px;border-radius:9px;background:#eef2f7;color:#34506f;margin-right:4px}
.find-list .w{color:#b42318;text-decoration:line-through;white-space:pre}
.find-list .s{color:#1e6b2d;font-weight:600;white-space:pre}
.find-list .n{display:block;color:var(--muted);font-size:12px;margin-top:2px}
.find-list .acts{display:flex;gap:6px;margin-top:5px}
.find-list .acts button{height:24px;padding:0 8px;border:1px solid var(--line-strong);border-radius:5px;background:#fff;cursor:pointer;font:inherit;font-size:12px}
.warnbox{padding:8px 10px;border:1px solid #ead18e;border-radius:6px;background:#fff7df;color:#5f4a12;font-size:12px;line-height:1.55}
.symbols{display:grid;grid-template-columns:repeat(8,1fr);gap:4px}
.symbols button{height:32px;border:1px solid var(--line);border-radius:5px;background:#fff;cursor:pointer;font:16px "TH SarabunPSK","TH Sarabun New",Tahoma,sans-serif}
.symbols button:hover{background:var(--hover)}
.au-opts{display:grid;gap:5px;margin:6px 0 10px}
.au-opts label{display:block;line-height:1.45;font-size:12.5px;cursor:pointer}
.au-opts label.sub{margin-left:22px;color:var(--muted);font-size:12px}
.au-opts input{margin-right:4px}
#auShield{position:fixed;inset:0;z-index:80;display:flex;align-items:flex-end;justify-content:center;padding-bottom:64px;background:rgba(15,23,42,.18);cursor:progress}
#auShield[hidden]{display:none}
.au-shield-card{display:flex;align-items:center;gap:12px;max-width:min(560px,90vw);padding:10px 14px;border-radius:10px;background:#fff;box-shadow:0 8px 28px rgba(15,23,42,.3);font-size:13px}
.au-shield-card .tbtn{height:30px;padding:0 12px;border:1px solid var(--line-strong);border-radius:6px;background:#fff;cursor:pointer;font:inherit}
.au-tag{margin-left:6px;padding:1px 8px;border-radius:9px;background:#fff7ed;color:#c2410c;font-size:11px;font-weight:600}
.au-step{display:flex;gap:6px;align-items:baseline}
.au-cnt{flex:none;min-width:28px;padding:0 7px;border-radius:9px;background:#ecfdf5;color:#047857;font-weight:700;text-align:center}
.au-cnt.zero{background:#f1f5f9;color:#94a3b8}
.au-ex{display:block;margin:2px 0 0 34px;color:var(--muted);font-size:11.5px;line-height:1.5;white-space:pre-wrap;word-break:break-word}
.steps{margin:4px 0 8px;padding-left:20px;color:var(--muted);font-size:12px;line-height:1.7}
.paste-box{min-height:44px;margin:6px 0;padding:8px 10px;border:2px dashed var(--line-strong);border-radius:8px;background:#fafbfc;color:var(--ink);overflow:hidden;max-height:60px;outline:none;font-size:12px}
.paste-box:empty::before{content:attr(data-placeholder);color:#9aa3ad}
.paste-box:focus{border-color:var(--accent);background:#fff}
.paste-frame{display:block;width:100%;height:320px;margin-top:8px;border:1px solid var(--line-strong);border-radius:6px;background:#fff}
.pt-sum{display:flex;flex-wrap:wrap;gap:6px;margin:6px 0}
.pt-pill{padding:2px 10px;border-radius:10px;font-weight:600;font-size:12px}
.pt-ok{background:#ecfdf5;color:#047857}.pt-changed{background:#fffbeb;color:#b45309}.pt-lost{background:#fef2f2;color:#b91c1c}.pt-missing{background:#f1f5f9;color:#64748b}
.pt-row{display:grid;grid-template-columns:22px 1fr;gap:2px 6px}
.pt-icon{font-weight:700}
.pt-row .n{grid-column:2}
.garuda-prev{height:34px;width:auto;border:1px solid var(--line);border-radius:4px;padding:2px;background:#fff}
.head-acts{display:flex;align-items:center;gap:4px}
.head-acts .tbtn{height:28px;padding:0 8px;font-size:12px}
.preset-list,.phrase-list{list-style:none;margin:0;padding:0}
.preset-list li,.phrase-list li{display:flex;gap:4px;margin-bottom:5px}
.preset,.phrase{flex:1;min-width:0;padding:6px 9px;border:1px solid var(--line);border-radius:6px;background:#fff;text-align:left;cursor:pointer;font:inherit}
.preset:hover,.phrase:hover{background:var(--hover);border-color:var(--line-strong)}
.preset strong{display:block;font-size:13px}
.preset small{display:block;color:var(--muted);font-size:11.5px;margin-top:1px}
.phrase{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.preset-list .del,.phrase-list .del{width:28px;flex:none;border:1px solid var(--line);border-radius:6px;background:#fff;color:var(--muted);cursor:pointer;font-size:15px}
.preset-list .del:hover,.phrase-list .del:hover{color:#b42318;border-color:#e3b1ab}
.calc{position:fixed;top:90px;right:380px;z-index:40;width:250px;padding:0 10px 10px;background:#fff;border:1px solid #c5cad3;border-radius:10px;box-shadow:0 10px 28px rgba(20,30,50,.22);font-size:13px}
.calc[hidden]{display:none}
.calc-head{display:flex;align-items:center;justify-content:space-between;margin:0 -10px 8px;padding:6px 6px 6px 10px;border-bottom:1px solid var(--line);cursor:move;user-select:none;font-weight:600;touch-action:none}
.calc-in{width:100%;height:34px;border:1px solid var(--line-strong);border-radius:6px;padding:0 8px;font:16px/1 Consolas,"Segoe UI",monospace;text-align:right}
.calc-out{min-height:24px;margin:4px 2px 8px;text-align:right;font:600 18px/1.3 Consolas,"Segoe UI",monospace;color:#1e4f8a;word-break:break-all}
.calc-out.err{font:12px/1.4 inherit;color:#b42318}
.calc-keys{display:grid;grid-template-columns:repeat(4,1fr);gap:5px}
.calc-keys button{height:34px;border:1px solid var(--line);border-radius:6px;background:#f7f8fa;cursor:pointer;font:15px "Segoe UI",sans-serif}
.calc-keys button:hover{background:var(--hover)}
.calc-keys .op{background:#eef3fb;color:#1e4f8a}
.calc-keys .fn{background:#fdf1ef;color:#9b2c1f}
.calc-keys .eq{grid-column:1/-1;background:var(--accent);border-color:var(--accent);color:#fff;font-weight:600}
.calc .row{margin-top:8px}
.calc .tbtn{height:28px;padding:0 8px;border:1px solid var(--line-strong);border-radius:6px;background:#fff;cursor:pointer;font:inherit;font-size:12px}
.calc .inline{display:inline-flex;align-items:center;gap:4px;font-size:12px}
.calc-hist{list-style:none;margin:6px 0 0;padding:0;max-height:110px;overflow:auto;font:12px Consolas,monospace;color:var(--muted)}
.calc-hist li{padding:3px 4px;border-radius:4px;cursor:pointer}
.calc-hist li:hover{background:var(--hover);color:var(--ink)}
.scratch{position:fixed;top:0;left:0;bottom:0;z-index:30;width:300px;max-width:100vw;display:flex;flex-direction:column;background:#fffdf5;border-right:1px solid var(--line);box-shadow:4px 0 18px rgba(20,30,50,.12);font-size:13px}
.scratch[hidden]{display:none}
.scratch .tools-head{background:#fff8dc}
.scratch h2{margin:0;font-size:15px}
.scratch .hint.pad{margin:6px 12px;color:var(--muted);font-size:12px}
.scratch .tbtn{height:28px;padding:0 8px;border:1px solid var(--line-strong);border-radius:6px;background:#fff;cursor:pointer;font:inherit;font-size:12px}
.scratch-body{flex:1;overflow:auto;padding:0 12px 12px}
.note{margin-bottom:10px;border:1px solid #ecdca7;border-radius:6px;background:#fff}
.note-head{display:flex;align-items:center;justify-content:space-between;padding:4px 6px 4px 8px;border-bottom:1px solid #f2e7c3;font-size:12px;color:#7a6320}
.note-head button{margin-left:4px;height:22px;padding:0 6px;border:1px solid #e6d59f;border-radius:4px;background:#fff;cursor:pointer;font:inherit;font-size:11.5px}
.note textarea{display:block;width:100%;min-height:150px;padding:8px;border:0;border-radius:0 0 6px 6px;resize:vertical;background:repeating-linear-gradient(#fff 0 23px,#f1ead2 23px 24px);font:15px/24px "TH SarabunPSK","TH Sarabun New",Tahoma,sans-serif;outline:none}
body.scratch-open .topbar,body.scratch-open .desk,body.scratch-open .statusbar{margin-left:300px}
@media (max-width:1100px){body.scratch-open .topbar,body.scratch-open .desk,body.scratch-open .statusbar{margin-left:0}}
@media print{.tools,.calc,.scratch{display:none!important}}
`;
    document.head.append(css);

    const panel = document.createElement('aside');
    panel.className = 'tools';
    panel.id = 'toolsPanel';
    panel.hidden = true;
    panel.setAttribute('aria-label', 'เครื่องมือเสริม');
    panel.innerHTML = `
<div class="tools-head"><h2>เครื่องมือเสริม</h2>
  <span class="head-acts"><button type="button" class="tbtn" id="btnCalc" title="เครื่องคิดเลข">🧮 เครื่องคิดเลข</button><button type="button" class="tbtn" id="btnScratch" title="กระดาษทดด้านซ้าย (หายเมื่อปิดหน้า)">📝 กระดาษทด</button><button type="button" class="tools-x" id="toolsClose" aria-label="ปิดแผงเครื่องมือ" title="ปิด">×</button></span></div>
<div class="tools-tabs" role="tablist">
  <button type="button" role="tab" data-tab="auto" aria-selected="false">จัดอัตโนมัติ</button>
  <button type="button" role="tab" data-tab="num" aria-selected="true">ตัวเลข</button>
  <button type="button" role="tab" data-tab="style" aria-selected="false">สไตล์</button>
  <button type="button" role="tab" data-tab="phrase" aria-selected="false">คำบ่อย</button>
  <button type="button" role="tab" data-tab="check" aria-selected="false">ตรวจคำ</button>
  <button type="button" role="tab" data-tab="ai" aria-selected="false">ผู้ช่วย AI</button>
  <button type="button" role="tab" data-tab="find" aria-selected="false">ค้นหา</button>
  <button type="button" role="tab" data-tab="file" aria-selected="false">ไฟล์</button>
  <button type="button" role="tab" data-tab="layout" aria-selected="false">จัดหน้า</button>
  <button type="button" role="tab" data-tab="nobreak" aria-selected="false">คำห้ามตัด</button>
  <button type="button" role="tab" data-tab="paste" aria-selected="false">ทดสอบวาง</button>
  <button type="button" role="tab" data-tab="theme" aria-selected="false">มุมมอง</button>
</div>
<div class="tools-body">
  <section data-panel="style" hidden>
    <h3>มาตรฐานเอกสาร</h3>
    <div class="row"><select id="setProfile" class="grow" aria-label="มาตรฐานเอกสาร"></select></div>
    <p class="hint">ค่าเริ่มต้นของย่อหน้าที่ไม่ได้ตั้งค่าเอง ใช้ทั้งบนจอ ตอนพิมพ์ และตอนคัดลอกไป e-Office</p>
    <div class="row"><label class="inline">เลขรายการลำดับ <select id="setListDigits"><option value="arabic">1. 2. 3.</option><option value="thai">๑. ๒. ๓.</option></select></label></div>
    <div class="row"><label class="inline"><input type="checkbox" id="setPageNum"> พิมพ์เลขหน้า “- ๒ -” กลางด้านบน (เริ่มหน้า 2)</label><select id="setPageDigits" aria-label="ตัวเลขของเลขหน้า"><option value="thai">เลขไทย</option><option value="arabic">อารบิก</option></select></div>
    <h3>จุดแท็บของย่อหน้า</h3>
    <div class="chips" id="tabChips">
      <button type="button" class="chip" data-tabsize="">ปกติ</button><button type="button" class="chip" data-tabsize="2.5cm">2.5 ซม.</button><button type="button" class="chip" data-tabsize="5cm">5 ซม.</button><button type="button" class="chip" data-tabsize="8cm">8 ซม.</button><button type="button" class="chip" data-tabsize="10cm">10 ซม.</button>
    </div>
    <p class="hint">กด Tab ในย่อหน้านั้นแล้วข้อความถัดไปจะเริ่มที่ระยะนี้พอดี (นับจากต้นบรรทัด) เช่น ตั้ง 8 ซม. ให้ “วันที่” ในบันทึกข้อความ</p>
    <h3>แม่แบบเอกสาร</h3>
    <div class="row"><button type="button" class="tbtn primary" data-act="tplMemo">บันทึกข้อความ</button><button type="button" class="tbtn primary" data-act="tplLetter">หนังสือภายนอก</button></div>
    <p class="hint">วางโครงตามระเบียบงานสารบรรณ แล้วพิมพ์ทับข้อความในวงเล็บ — ถ้าหน้านี้มีข้อความอยู่แล้ว จะถามว่าจะสร้างในหน้าใหม่หรือไม่</p>
    <div class="row"><img id="garudaPrev" class="garuda-prev" alt="รูปครุฑที่เลือกไว้" hidden><button type="button" class="tbtn" data-act="garudaPick">เลือกรูปครุฑ…</button><button type="button" class="tbtn" data-act="garudaClear" id="garudaClear">ลบรูป</button></div>
    <p class="hint">รูปครุฑใช้กับแม่แบบ (บันทึกข้อความ สูง 1.5 ซม. / หนังสือภายนอก สูง 3 ซม.) เก็บไว้ในเบราว์เซอร์นี้ — ถ้ายังไม่เลือก แม่แบบจะใส่ “(ครุฑ)” ไว้แทน</p>
    <input type="file" id="garudaFile" accept="image/png,image/jpeg,image/gif,image/webp" hidden>
    <h3>สไตล์ย่อหน้า (preset)</h3>
    <p class="hint">คลิกในย่อหน้า (หรือเลือกหลายย่อหน้า) แล้วกดชื่อสไตล์ เพื่อใช้รูปแบบนั้นทั้งย่อหน้า — กดย้อนกลับได้</p>
    <ul class="preset-list" id="presetList"></ul>
    <h3>บันทึกสไตล์ของฉัน</h3>
    <p class="hint">จัดย่อหน้าให้เรียบร้อย (ไม้บรรทัด ระยะบรรทัด ฟอนต์ ตัวหนา ฯลฯ) แล้วตั้งชื่อ ระบบจะเก็บรูปแบบของย่อหน้าที่เคอร์เซอร์อยู่</p>
    <div class="row"><input type="text" id="presetName" class="grow" placeholder="ชื่อสไตล์ เช่น ย่อหน้ารายงาน" maxlength="40"><button type="button" class="tbtn primary" data-act="presetSave">บันทึก</button></div>
  </section>

  <section data-panel="phrase" hidden>
    <h3>คำ/สำนวนที่ใช้บ่อย</h3>
    <p class="hint">คลิกเพื่อแทรกที่เคอร์เซอร์</p>
    <ul class="phrase-list" id="phraseList"></ul>
    <div class="row"><input type="text" id="phraseIn" class="grow" placeholder="พิมพ์คำหรือสำนวนใหม่" maxlength="300"><button type="button" class="tbtn primary" data-act="phraseAdd">เพิ่ม</button></div>
    <div class="row"><button type="button" class="tbtn" data-act="phraseFromSel">เพิ่มจากข้อความที่เลือก</button><button type="button" class="tbtn" data-act="phraseReset">คืนค่ารายการตัวอย่าง</button></div>
  </section>

  <section data-panel="file" hidden>
    <h3>ชื่อเอกสาร</h3>
    <div class="row"><input type="text" id="fileName" class="grow" maxlength="80" placeholder="ว่าง = ใช้วันเวลาที่บันทึกครั้งแรก"></div>
    <p class="hint">ชื่อนี้ใช้เป็นชื่อไฟล์ตอนบันทึก ถ้าไม่ตั้ง ระบบจะตั้งเป็น “ร่าง-ปีเดือนวัน-เวลา” ของการบันทึกครั้งแรกให้เอง</p>
    <div class="row"><button type="button" class="tbtn primary" data-act="fileDownload">บันทึกเป็นไฟล์ (ดาวน์โหลด)</button><button type="button" class="tbtn" data-act="fileLink" id="fileLinkBtn">บันทึกลงไฟล์อัตโนมัติ</button></div>
    <div class="row"><button type="button" class="tbtn" data-act="fileOpen">เปิดไฟล์…</button></div>
    <h3>ไฟล์ที่เคยบันทึก/เปิด</h3>
    <ul class="find-list" id="recentList"></ul>
    <p class="hint" id="recentHint"></p>
  </section>

  <section data-panel="num">
    <h3>แปลงตัวเลข</h3>
    <div class="row"><label class="inline">ขอบเขต <select id="numScope"><option value="selection">ส่วนที่เลือก (ถ้าไม่เลือก = หน้านี้)</option><option value="page">ทั้งหน้านี้</option><option value="all">ทุกหน้า</option></select></label></div>
    <div class="row"><button type="button" class="tbtn" data-act="toThai">123 → ๑๒๓ (เป็นเลขไทย)</button><button type="button" class="tbtn" data-act="toArabic">๑๒๓ → 123 (เป็นเลขอารบิก)</button></div>
    <p class="hint">เปลี่ยนเฉพาะตัวอักษรตัวเลข รูปแบบตัวอักษรเดิมยังอยู่ หน้าที่เปิดอยู่กดย้อนกลับได้</p>
    <h3>จำนวนเงินเป็นตัวอักษร</h3>
    <div class="row"><input type="text" id="bahtIn" class="grow" inputmode="decimal" placeholder="เช่น 12,500.50"><button type="button" class="tbtn" data-act="bahtFromSel" title="ใช้ตัวเลขที่เลือกไว้ในเอกสาร">จากที่เลือก</button></div>
    <div class="out" id="bahtOut">—</div>
    <div class="row"><button type="button" class="tbtn" data-act="bahtInsert">แทรก (คำอ่าน)</button><button type="button" class="tbtn" data-act="bahtInsertBoth">แทรก 12,500.50 บาท (คำอ่าน)</button></div>
    <h3>วันที่</h3>
    <div class="row"><label class="inline"><input type="checkbox" id="dateThaiDigits"> ใช้เลขไทย</label></div>
    <div class="row"><button type="button" class="tbtn" data-act="dateToday" id="dateBtn">แทรกวันที่วันนี้</button></div>
    <h3>เว้นวรรคห้ามตัดบรรทัด</h3>
    <div class="row"><button type="button" class="tbtn" data-act="nbsp">แทรกเว้นวรรคห้ามตัด</button><span class="hint">หรือกด Ctrl+Shift+Space</span></div>
    <p class="hint">ใช้ระหว่างคำที่ไม่ควรถูกแยกคนละบรรทัด เช่น ตัวเลขกับหน่วย “๑๒ บาท” หรือเลขที่หนังสือ</p>
    <h3>สัญลักษณ์</h3>
    <div class="symbols" id="symbols"></div>
  </section>

  <section data-panel="check" hidden>
    <h3>ตรวจคำ การเขียน และรูปแบบหนังสือราชการ (ออฟไลน์)</h3>
    <p class="hint">ตรวจหน้าที่เปิดอยู่: คำที่มักเขียนผิด การเว้นวรรค ไม้ยมก/ฯลฯ ตัวเลขกับหน่วยที่อาจถูกตัดบรรทัด รูปแบบวันที่ และความสม่ำเสมอของย่อหน้า ระยะบรรทัด ขนาดและแบบอักษร รวมถึงคำขึ้นต้น–ลงท้าย — ไม่ได้ตรวจทุกคำในพจนานุกรม ถ้าต้องการตรวจละเอียดใช้แท็บผู้ช่วย AI</p>
    <div class="row"><label class="inline"><input type="checkbox" id="allowBigSpace"> ยอมรับเว้นวรรคใหญ่ (2 เคาะ)</label></div>
    <div class="row"><button type="button" class="tbtn primary" data-act="scan">ตรวจหน้านี้</button><button type="button" class="tbtn" data-act="fixAll" id="fixAllBtn" disabled>แก้ทั้งหมด</button></div>
    <div id="scanSummary" class="hint"></div>
    <ul class="find-list" id="scanList"></ul>
  </section>

  <section data-panel="ai" hidden>
    <h3>ผู้ช่วยภาษาไทยด้วย AI</h3>
    <div class="row"><label class="inline grow">ผู้ให้บริการ <select id="aiProvider" class="grow"></select></label></div>
    <div class="row" id="aiModelRow"><label class="inline grow">รุ่น <input type="text" id="aiModel" class="grow" spellcheck="false"></label></div>
    <div class="row" id="aiUrlRow"><label class="inline grow">Base URL <input type="text" id="aiUrl" class="grow" spellcheck="false" placeholder="http://localhost:11434/v1"></label></div>
    <div class="warnbox" id="aiWarn"></div>
    <div class="row"><input type="password" id="aiKey" class="grow" autocomplete="off"></div>
    <div class="row"><label class="inline"><input type="checkbox" id="aiRemember"> จำคีย์ไว้ในเบราว์เซอร์นี้</label></div>
    <div class="row"><label class="inline"><input type="checkbox" id="aiConsent"> ยืนยันว่าข้อความนี้ส่งออกนอกเครื่องได้</label></div>
    <div class="row">
      <button type="button" class="tbtn primary" data-act="ai" data-task="check">ตรวจคำผิด/ไวยากรณ์</button>
      <button type="button" class="tbtn" data-act="ai" data-task="formal">เรียบเรียงภาษาราชการ</button>
      <button type="button" class="tbtn" data-act="ai" data-task="summary">สรุปใจความ</button>
    </div>
    <div id="aiStatus" class="hint"></div>
    <div id="aiResult"></div>
  </section>

  <section data-panel="find" hidden>
    <h3>ค้นหาและแทนที่</h3>
    <div class="row"><input type="text" id="findText" class="grow" placeholder="ค้นหา"></div>
    <div class="row"><input type="text" id="replText" class="grow" placeholder="แทนที่ด้วย (ว่าง = ลบ)"></div>
    <div class="row"><label class="inline">ขอบเขต <select id="findScope"><option value="page">หน้านี้</option><option value="selection">ส่วนที่เลือก</option><option value="all">ทุกหน้า</option></select></label><label class="inline"><input type="checkbox" id="findCase"> ตรงตัวพิมพ์ใหญ่-เล็ก</label></div>
    <div class="row"><button type="button" class="tbtn" data-act="findNext">ค้นหาถัดไป</button><button type="button" class="tbtn primary" data-act="replaceAll">แทนที่ทั้งหมด</button></div>
    <div id="findOut" class="hint"></div>
  </section>

  <section data-panel="auto" hidden>
    <h3>จัดรูปแบบอัตโนมัติ (ทีเดียวจบ)</h3>
    <p class="hint">รวมการล้างข้อความ จัดย่อหน้าให้สอดคล้อง ตัดคำ และเลขหน้า — กด “ดูตัวอย่าง” เพื่อดูก่อน-หลังโดยยังไม่เปลี่ยนเอกสาร แล้วค่อยกด “จัดทั้งหมด” ย้อนกลับทั้งเอกสารได้ทุกหน้า</p>
    <div class="row"><label class="inline">ขอบเขต <select id="auScope"><option value="page">หน้านี้</option><option value="all">ทุกหน้า</option></select></label>
      <label class="inline">ตัวเลข <select id="auDigits"><option value="keep">ไม่เปลี่ยน</option><option value="thai">เป็นเลขไทย</option><option value="arabic">เป็นเลขอารบิก</option></select></label></div>
    <div class="au-opts">
      <label><input type="checkbox" data-au="clean"> <b>ล้างข้อความ</b> เว้นวรรคซ้อน/ท้ายย่อหน้า อักขระล่องหน เเ→แ สระอำ วรรณยุกต์ซ้ำ ไม้ยมก ฯลฯ</label>
      <label class="sub"><input type="checkbox" data-au="bigspace"> คงเว้นวรรคใหญ่ (2 เคาะ) ไว้ — ยุบเฉพาะที่เกิน 2 เคาะ</label>
      <label><input type="checkbox" data-au="blank"> <b>ลบบรรทัดว่างซ้อนกัน</b> เหลือบรรทัดเดียว</label>
      <label><input type="checkbox" data-au="misspell"> <b>แก้คำสะกดที่ผิดบ่อย</b> จากรายการในแท็บตรวจคำ</label>
      <label><input type="checkbox" data-au="dates"> <b>วันที่แบบราชการ</b> เดือนเต็ม ปี พ.ศ. (แปลงปี ค.ศ.→พ.ศ. ด้วย)</label>
      <label><input type="checkbox" data-au="layout"> <b>จัดย่อหน้าให้สอดคล้อง</b> ย่อหน้าบรรทัดแรก ระยะบรรทัด จัดแนว ส่วนลงชื่อ</label>
      <label><input type="checkbox" data-au="bind"> <b>ตัดคำ</b> ผูกตัวเลขกับหน่วย และคำในพจนานุกรมคำห้ามตัด</label>
      <label><input type="checkbox" data-au="spread"> <b>กระจายตัวอักษร</b> ปรับช่องไฟดึงคำโดด/ลดการยืด (ไม่เกิน ±0.03 em)</label>
      <label><input type="checkbox" data-au="pagenum"> <b>เลขหน้า</b> “- ๒ -” กลางด้านบนตอนพิมพ์</label>
    </div>
    <h3>ให้ AI ช่วย (เลือกได้) <span class="au-tag">ส่งข้อความออกนอกเครื่อง</span></h3>
    <div class="au-opts">
      <label><input type="checkbox" data-au="aiLayout"> <b>AI วางแผนจัดหน้า</b> ตัดคำไทย กระจายตัวอักษร ระยะบรรทัด — แทนขั้น “จัดย่อหน้า/กระจายตัวอักษร” แบบในเครื่อง</label>
      <label><input type="checkbox" data-au="aiCheck"> <b>AI ตรวจคำผิด/ไวยากรณ์/การใช้คำ</b> แล้วแก้ให้</label>
      <label><input type="checkbox" data-au="aiRewrite"> <b>AI เรียบเรียงภาษาราชการ</b> แก้ถ้อยคำทั้งย่อหน้า (ตัวเลขต้องตรงต้นฉบับ ย่อหน้าที่มีตัวหนา/เอียงผสมจะข้าม) — ควรดูตัวอย่างก่อน</label>
      <label><input type="checkbox" data-au="aiAudit"> <b>AI ตรวจโครงสร้างและความสอดคล้อง</b> วันที่/จำนวนเงิน/ชื่อตำแหน่ง/คำขึ้นต้น–ลงท้าย (รายงานอย่างเดียว ไม่แก้)</label>
    </div>
    <div id="auAiInfo" class="hint"></div>
    <div class="row"><button type="button" class="tbtn" data-act="auPreview">ดูตัวอย่าง</button><button type="button" class="tbtn primary" data-act="auApply">จัดทั้งหมด</button><button type="button" class="tbtn" data-act="auUndo" id="auUndoBtn" disabled>ย้อนกลับทั้งหมด</button></div>
    <div id="auStatus" class="hint"></div>
    <ul class="find-list" id="auLog"></ul>
    <div id="auCompare" hidden>
      <div class="row"><button type="button" class="tbtn" data-act="auShow" data-which="before">ก่อน</button><button type="button" class="tbtn" data-act="auShow" data-which="after">หลัง</button><span class="hint" id="auShowing"></span></div>
      <iframe id="auFrame" class="paste-frame" sandbox="" title="ตัวอย่างก่อน-หลังการจัดอัตโนมัติ"></iframe>
    </div>
  </section>

  <section data-panel="paste" hidden>
    <h3>ทดสอบการวางใน e-Office</h3>
    <p class="hint">ตรวจว่ารูปแบบแต่ละอย่างรอดหรือหายเมื่อวางใน e-Office จริง แล้วปรับการส่งออกและเตือนให้ตามผล</p>
    <ol class="steps">
      <li>กด <b>คัดลอกชุดทดสอบ</b></li>
      <li>วางใน e-Office แล้วบันทึก/เปิดดูตามปกติ</li>
      <li>เลือกเนื้อหาที่ได้ใน e-Office ทั้งหมด แล้วคัดลอก (Ctrl+A, Ctrl+C)</li>
      <li>กลับมากด <b>อ่านจากคลิปบอร์ด</b> หรือวางในกล่องด้านล่าง (Ctrl+V)</li>
    </ol>
    <div class="row"><button type="button" class="tbtn primary" data-act="pasteCopy">1) คัดลอกชุดทดสอบ</button><button type="button" class="tbtn" data-act="pasteRead">4) อ่านจากคลิปบอร์ด</button></div>
    <div id="pasteBox" class="paste-box" contenteditable="true" role="textbox" aria-label="วางผลที่คัดลอกกลับมาจาก e-Office" data-placeholder="หรือคลิกที่นี่แล้วกด Ctrl+V…"></div>
    <div id="pasteStatus" class="hint"></div>
    <div id="pasteSummary"></div>
    <ul class="find-list" id="pasteList"></ul>
    <h3>ตัวอย่างหลังส่งออก</h3>
    <p class="hint">ดู HTML ที่จะคัดลอกไป e-Office ในกรอบแยก (ไม่มี CSS ของแอป) — เปิด “จำลอง” เพื่อลบสิ่งที่ผลทดสอบบอกว่า e-Office ล้างทิ้ง</p>
    <div class="row"><button type="button" class="tbtn" data-act="pastePreview">ดูตัวอย่างหน้านี้</button><label class="inline"><input type="checkbox" id="pasteSim"> จำลองตามผลทดสอบ</label><button type="button" class="tbtn" data-act="pasteClear">ลบผลทดสอบ</button></div>
    <iframe id="pasteFrame" class="paste-frame" sandbox="" title="ตัวอย่าง HTML ที่จะคัดลอกไป e-Office" hidden></iframe>
  </section>

  <section data-panel="nobreak" hidden>
    <h3>คำห้ามตัด (ไม่ให้ตัดกลางคำตอนขึ้นบรรทัดใหม่)</h3>
    <p class="hint">เบราว์เซอร์ตัดคำไทยด้วยพจนานุกรม จึงมักตัดคำประสม ชื่อหน่วยงาน และชื่อคนกลางคำ เช่น “ผู้อำนวย⏎การ” — ระบบใช้ Intl.Segmenter หาจุดที่อาจถูกตัดภายในคำเหล่านี้
      แล้วใส่ตัวผูกคำ (Word Joiner U+2060 มองไม่เห็น ไม่เปลี่ยนตัวอักษร) เฉพาะจุดนั้น — ดูตำแหน่งตัวผูกได้ด้วยปุ่ม ¶ (เครื่องหมาย ⁀ สีเขียว)</p>
    <div class="row"><input type="text" id="nbIn" class="grow" placeholder="เพิ่มคำ เช่น ชื่อหน่วยงาน ชื่อคน (คั่นหลายคำด้วยจุลภาค)" maxlength="400"><button type="button" class="tbtn primary" data-act="nbAdd">เพิ่ม</button></div>
    <div class="row"><button type="button" class="tbtn" data-act="nbFromSel">เพิ่มจากข้อความที่เลือก</button><button type="button" class="tbtn" data-act="nbReset">คืนค่ารายการตั้งต้น</button></div>
    <div class="row"><label class="inline">ขอบเขต <select id="nbScope"><option value="page">หน้านี้</option><option value="all">ทุกหน้า</option></select></label>
      <button type="button" class="tbtn primary" data-act="nbApply">ผูกคำในเอกสาร</button><button type="button" class="tbtn" data-act="nbRemove">ถอดตัวผูกทั้งหมด</button></div>
    <div id="nbStatus" class="hint"></div>
    <ul class="phrase-list" id="nbList"></ul>
  </section>

  <section data-panel="layout" hidden>
    <h3>ปรับให้สอดคล้องกันตามหนังสือราชการ</h3>
    <p class="hint">วิเคราะห์ทุกย่อหน้า แล้วเสนอการปรับทีละย่อหน้า (ย่อหน้าบรรทัดแรก ระยะบรรทัด การจัดแนว ระยะเว้น และวลีที่ควรห้ามตัดบรรทัด) ให้เลือกใช้หรือข้าม — ย่อหน้าในหน้าที่เปิดอยู่จะวัดช่องห่างจากการจัดเสมอหน้าหลังด้วย</p>
    <div class="row"><label class="inline">ขอบเขต <select id="lyScope"><option value="page">หน้านี้</option><option value="all">ทุกหน้า</option></select></label></div>
    <div class="row"><button type="button" class="tbtn primary" data-act="lyLocal">วิเคราะห์ (ในเครื่อง)</button><button type="button" class="tbtn" data-act="lyAi" id="lyAiBtn">ให้ AI วางแผนและวิเคราะห์</button></div>
    <p class="hint">AI ใช้ผู้ให้บริการ คีย์ และการยืนยันจากแท็บผู้ช่วย AI และส่งข้อความต้นย่อหน้า (ไม่เกิน 160 ตัวอักษร) พร้อมค่าการจัดหน้า</p>
    <div id="lyStatus" class="hint"></div>
    <div id="lyPlan"></div>
    <div class="row"><button type="button" class="tbtn" data-act="lyApplyAll" id="lyApplyAll" disabled>ใช้ทุกข้อที่เหลือ</button></div>
    <ul class="find-list" id="lyList"></ul>
  </section>

  <section data-panel="theme" hidden>
    <h3>หน้ากระดาษ</h3>
    <div class="row"><label class="inline"><input type="checkbox" id="vwPaper"> แสดงเป็นกระดาษ A4 แนวตั้ง (ขอบกระดาษ + กรอบขอบพิมพ์สีจาง)</label></div>
    <div class="row"><label class="inline"><input type="checkbox" id="vwBreaks"> เส้นแบ่งหน้าโดยประมาณ</label></div>
    <div class="row"><label class="inline"><input type="checkbox" id="vwGrid"> กระดาษกราฟ ตาราง 1 ซม.</label></div>
    <h3>เส้นบรรทัด (สลับสีตามย่อหน้า)</h3>
    <div class="row"><label class="inline"><input type="checkbox" id="vwLines"> แสดงเส้นประใต้ทุกบรรทัด</label></div>
    <div class="row" id="vwLineColors"></div>
    <p class="hint">แต่ละย่อหน้าใช้สีวนตามลำดับ พร้อมแถบสีด้านซ้าย ทำให้เห็นว่าบรรทัดไหนอยู่ย่อหน้าเดียวกัน</p>
    <h3>เลขที่ย่อหน้า / เลขที่บรรทัด</h3>
    <div class="row"><label class="inline"><input type="checkbox" id="vwParaNum"> เลขที่ย่อหน้า (ตรงกับ “ย่อหน้า N” ในแท็บจัดหน้า)</label></div>
    <div class="row"><label class="inline"><input type="checkbox" id="vwLineNum"> เลขที่บรรทัด (นับทุกบรรทัดในหน้านี้)</label></div>
    <h3>เครื่องหมายซ่อน (¶)</h3>
    <div class="row"><label class="inline"><input type="checkbox" id="vwMarks"> แสดงเครื่องหมาย (Ctrl+Shift+8)</label></div>
    <div id="vwMarkList"></div>
    <p class="hint">วาดทับบนหน้าจอเท่านั้น ไม่ถูกพิมพ์ ไม่ถูกคัดลอก และไม่ทำให้บรรทัดเลื่อน</p>
    <h3>สีพื้นกระดาษ</h3>
    <p class="hint">เปลี่ยนเฉพาะการแสดงผลบนจอ ไม่มีผลกับไฟล์ที่บันทึก การพิมพ์ หรือข้อความที่คัดลอกไป e-Office</p>
    <div id="themeList"></div>
    <h3>ความอุ่นของสี (ลดแสงสีฟ้า)</h3>
    <div class="row"><input type="range" id="warmth" min="0" max="60" step="5" class="grow" aria-label="ความอุ่นของสี"><span id="warmthVal"></span></div>
  </section>
</div>`;
    document.body.append(panel);

    // ปุ่มเปิดแผงในแถบด้านบน
    const openBtn = document.createElement('button');
    openBtn.type = 'button';
    openBtn.className = 'btn';
    openBtn.id = 'btnTools';
    openBtn.textContent = 'เครื่องมือเสริม';
    openBtn.setAttribute('aria-controls', 'toolsPanel');
    $('#btnCopy').before(openBtn);
    function setOpen(open) {
      panel.hidden = !open;
      document.body.classList.toggle('tools-open', open);
      openBtn.setAttribute('aria-expanded', String(open));
      store.set(KEY.panel, open ? '1' : '0');
      window.dispatchEvent(new Event('resize'));          // ให้ซูมอัตโนมัติคำนวณพื้นที่ใหม่
    }
    openBtn.addEventListener('click', () => setOpen(panel.hidden));
    $('#toolsClose').addEventListener('click', () => setOpen(false));

    const onTab = {};
    function showTab(name) {
      if (!$(`section[data-panel="${name}"]`, panel)) name = 'num';
      for (const b of $$('.tools-tabs button', panel)) b.setAttribute('aria-selected', String(b.dataset.tab === name));
      for (const s of $$('section[data-panel]', panel)) s.hidden = s.dataset.panel !== name;
      store.set(KEY.tab, name);
      onTab[name]?.();
    }
    $('.tools-tabs', panel).addEventListener('click', e => { const b = e.target.closest('button[data-tab]'); if (b) showTab(b.dataset.tab); });

    // ปุ่มในแผงไม่ดึงโฟกัสจากกล่องพิมพ์ ส่วนที่เลือกไว้จึงยังอยู่
    panel.addEventListener('mousedown', e => { if (e.target.closest('.tbtn, .find-list button, .symbols button, .tt-theme')) e.preventDefault(); });

    // ----- ตัวเลข -----
    const bahtIn = $('#bahtIn'), bahtOut = $('#bahtOut');
    const updateBaht = () => {
      const t = bahtText(bahtIn.value);
      bahtOut.textContent = bahtIn.value.trim() ? (t ? `(${t})` : 'ตัวเลขไม่ถูกต้อง') : '—';
      return t;
    };
    bahtIn.addEventListener('input', updateBaht);
    const fmtMoney = v => {
      const n = Number(toArabic(v).replace(/[,\s฿]|บาท/g, ''));
      return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    };
    $('#symbols').innerHTML = ['ฯ', 'ๆ', 'ฯลฯ', '฿', '๏', '๚', '๛', '•', '–', '—', '“', '”', '‘', '’', '…', '✓', '✗', '☐', '☑', '№', '§', '¶', '©', '→']
      .map(s => `<button type="button" data-sym="${esc(s)}" title="แทรก ${esc(s)}">${esc(s)}</button>`).join('');
    $('#symbols').addEventListener('click', e => { const b = e.target.closest('[data-sym]'); if (b) insertAtCursor(b.dataset.sym); });
    const refreshDateBtn = () => { $('#dateBtn').textContent = `แทรก “${thaiDate(new Date(), $('#dateThaiDigits').checked)}”`; };
    $('#dateThaiDigits').addEventListener('change', refreshDateBtn);
    refreshDateBtn();

    // ----- ตรวจคำ -----
    let scanItems = [];
    function renderScan() {
      const { found, mixedDigits } = scanOffline();
      scanItems = found;
      $('#fixAllBtn').disabled = !found.length;
      $('#scanSummary').textContent = (found.length ? `พบ ${found.length} จุดในหน้า ${ed.currentPage() + 1}` : `ไม่พบจุดผิดจากรายการที่ตรวจในหน้า ${ed.currentPage() + 1}`)
        + (mixedDigits ? ` · มีเลขไทย ${mixedDigits.thai} ตัว ปนกับเลขอารบิก ${mixedDigits.arabic} ตัว (แปลงให้เป็นแบบเดียวกันได้ในแท็บตัวเลข)` : '');
      $('#scanList').innerHTML = found.map((f, i) => `<li>
        <span class="k">${esc(f.kind)}</span>${f.wrong != null ? `<span class="w">${esc(show(f.wrong))}</span> → <span class="s">${esc(show(f.suggest)) || '(ลบ)'}</span>` : ''}
        <span class="n">${esc(f.note)}</span>
        <span class="acts"><button type="button" data-i="${i}" data-do="show">ดูในเอกสาร</button><button type="button" data-i="${i}" data-do="fix">แก้</button></span></li>`).join('');
    }
    // แสดงช่องว่างให้เห็น: ␣ = เว้นวรรคปกติ, ⍽ = เว้นวรรคห้ามตัดบรรทัด
    const show = s => s.replace(/ /g, '␣').replace(/\u00A0/g, '⍽');
    $('#scanList').addEventListener('click', e => {
      const b = e.target.closest('button[data-do]');
      if (!b) return;
      const item = scanItems[+b.dataset.i];
      if (b.dataset.do === 'show') selectItem(item);
      else { if (!applyFix(item)) say('ไม่พบข้อความนี้แล้ว'); renderScan(); }
    });

    // ----- AI -----
    const aiKey = $('#aiKey'), aiRemember = $('#aiRemember'), aiConsent = $('#aiConsent');
    const aiProvider = $('#aiProvider'), aiModel = $('#aiModel'), aiUrl = $('#aiUrl');
    aiProvider.innerHTML = Object.entries(PROVIDERS).map(([k, p]) => `<option value="${k}">${esc(p.label)}</option>`).join('');
    store.del(KEY.apiKey);                                         // คีย์แบบเก่าถูกย้ายไปเก็บแยกตามผู้ให้บริการแล้ว
    if (keys.claude && apiKey) store.set(keyStoreKey('claude'), keys.claude);
    const saveAiCfg = () => store.set(KEY.aiCfg, JSON.stringify({ ...aiCfg, provider }));
    function renderAiProvider() {
      const p = PROVIDERS[provider];
      aiProvider.value = provider;
      aiModel.value = modelOf(provider);
      aiModel.disabled = !!p.fixedModel;
      aiModel.title = p.fixedModel ? 'Claude ใช้รุ่นล่าสุดที่แนะนำ (Claude Opus 5.5)' : 'ชื่อรุ่นตามที่ผู้ให้บริการกำหนด แก้ได้';
      $('#aiUrlRow').hidden = provider !== 'custom';
      aiUrl.value = baseUrlOf();
      aiKey.placeholder = p.keyHint;
      aiKey.value = keys[provider];
      aiRemember.checked = !!store.get(keyStoreKey(provider));
      const dest = { claude: 'Anthropic', openai: 'OpenAI', gemini: 'Google' }[provider];
      $('#aiWarn').innerHTML = staysLocal()
        ? `เซิร์ฟเวอร์อยู่ในเครื่องนี้ (${esc(new URL(baseUrlOf()).host)}) ข้อความ<strong>ไม่ออกนอกเครื่อง</strong> — ต้องเปิดเซิร์ฟเวอร์และอนุญาต CORS ก่อน (Ollama: ตั้ง OLLAMA_ORIGINS=*)`
        : `ข้อความที่เลือก (หรือทั้งหน้าถ้าไม่ได้เลือก) จะถูก<strong>ส่งออกนอกเครื่องไปยัง ${esc(dest || new URL(baseUrlOf(), location.href).host)}</strong> ห้ามใช้กับเอกสารลับหรือข้อมูลส่วนบุคคล และมีค่าใช้จ่ายตาม API key ของคุณ`;
      aiConsent.closest('.row').hidden = staysLocal();
    }
    aiProvider.addEventListener('change', () => { provider = aiProvider.value; saveAiCfg(); renderAiProvider(); });
    aiModel.addEventListener('change', () => { aiCfg.models[provider] = aiModel.value.trim() || PROVIDERS[provider].model; saveAiCfg(); renderAiProvider(); });
    aiUrl.addEventListener('change', () => {
      const v = aiUrl.value.trim();
      if (v && !/^https?:\/\//i.test(v)) { $('#aiStatus').textContent = 'Base URL ต้องขึ้นต้นด้วย http:// หรือ https://'; return; }
      aiCfg.baseUrl = v || PROVIDERS.custom.baseUrl;
      saveAiCfg();
      renderAiProvider();
    });
    aiKey.addEventListener('input', () => { keys[provider] = aiKey.value.trim(); if (aiRemember.checked) store.set(keyStoreKey(provider), keys[provider]); });
    aiRemember.addEventListener('change', () => { if (aiRemember.checked && keys[provider]) store.set(keyStoreKey(provider), keys[provider]); else store.del(keyStoreKey(provider)); });
    renderAiProvider();
    let aiItems = [], aiBusy = false, aiRange = null;
    async function runAi(taskKey) {
      if (aiBusy) return;
      if (!keys[provider] && !PROVIDERS[provider].keyOptional) { $('#aiStatus').textContent = 'กรอก API key ก่อน'; aiKey.focus(); return; }
      if (!staysLocal() && !aiConsent.checked) { $('#aiStatus').textContent = 'ติ๊กยืนยันก่อนว่าข้อความนี้ส่งออกนอกเครื่องได้'; return; }
      const r = ed.range();
      aiRange = r.length ? { ...r } : null;
      const text = (r.length ? quill.getText(r.index, r.length) : quill.getText()).trim();
      if (!text) { $('#aiStatus').textContent = 'ยังไม่มีข้อความ'; return; }
      if (text.length > AI_MAX_CHARS) { $('#aiStatus').textContent = `ข้อความยาว ${text.length.toLocaleString()} ตัวอักษร เกิน ${AI_MAX_CHARS.toLocaleString()} — เลือกทีละส่วนแล้วลองใหม่`; return; }
      aiBusy = true;
      $$('[data-act=ai]', panel).forEach(b => { b.disabled = true; });
      $('#aiStatus').textContent = `กำลังส่ง${aiRange ? 'ส่วนที่เลือก' : `หน้า ${ed.currentPage() + 1}`} (${text.length.toLocaleString()} ตัวอักษร) ให้ ${PROVIDERS[provider].label} (${modelOf(provider)})… อาจใช้เวลาสักครู่`;
      $('#aiResult').innerHTML = '';
      try {
        const data = await askAi(taskKey, text);
        renderAi(taskKey, data);
        $('#aiStatus').textContent = `${TASKS[taskKey].label} — เสร็จแล้ว (ควรตรวจทานคำแนะนำของ AI ก่อนใช้)`;
      } catch (err) {
        $('#aiStatus').textContent = 'ไม่สำเร็จ: ' + aiErrorText(err);
      } finally {
        aiBusy = false;
        $$('[data-act=ai]', panel).forEach(b => { b.disabled = false; });
      }
    }
    function renderAi(taskKey, data) {
      const box = $('#aiResult');
      if (taskKey === 'check') {
        aiItems = (data.items || []).filter(it => it.wrong && it.wrong !== it.suggest).map(it => ({ ...it, index: aiRange?.index ?? 0, note: it.reason }));
        box.innerHTML = aiItems.length
          ? `<ul class="find-list">${aiItems.map((f, i) => `<li><span class="k">${esc(f.kind)}</span><span class="w">${esc(f.wrong)}</span> → <span class="s">${esc(f.suggest) || '(ลบ)'}</span><span class="n">${esc(f.reason)}</span><span class="acts"><button type="button" data-ai="${i}" data-do="show">ดูในเอกสาร</button><button type="button" data-ai="${i}" data-do="fix">แก้</button></span></li>`).join('')}</ul>`
          : '<div class="out">ไม่พบจุดที่ต้องแก้</div>';
      } else if (taskKey === 'formal') {
        box.innerHTML = `<div class="out" id="aiRewrite">${esc(data.rewritten)}</div><p class="hint">${esc(data.notes)}</p>
          <div class="row"><button type="button" class="tbtn primary" data-act="aiReplace">${aiRange ? 'แทนที่ส่วนที่เลือก' : 'แทนที่ทั้งหน้า'}</button><button type="button" class="tbtn" data-act="aiInsert">แทรกที่เคอร์เซอร์</button></div>`;
        box.dataset.text = data.rewritten;
      } else {
        box.innerHTML = `<div class="out">${(data.points || []).map(p => '• ' + esc(p)).join('\n')}</div>
          <div class="row"><button type="button" class="tbtn" data-act="aiInsertSummary">แทรกที่เคอร์เซอร์</button></div>`;
        box.dataset.text = (data.points || []).map(p => '• ' + p).join('\n');
      }
    }
    $('#aiResult').addEventListener('click', e => {
      const b = e.target.closest('button[data-ai]');
      if (!b) return;
      const item = aiItems[+b.dataset.ai];
      if (b.dataset.do === 'show') selectItem(item);
      else if (applyFix(item)) { b.closest('li').style.opacity = '.45'; b.disabled = true; }
      else say('ไม่พบข้อความนี้ในเอกสาร (อาจถูกแก้ไปแล้ว)');
    });
    function replaceRangeWithText(r, text) {
      // แทนที่ทั้งช่วง โดยคงรูปแบบตัวอักษรของตัวแรก (ย่อหน้าใหม่ใช้รูปแบบย่อหน้าเดิม)
      const attrs = attrsAt(quill.getContents(), r.index);
      quill.updateContents(new Delta().retain(r.index).delete(r.length).insert(text, attrs), 'user');
      quill.setSelection(r.index + text.length, 0, 'user');
    }

    // ----- ค้นหา -----
    const findRegex = flags => {
      const q = $('#findText').value;
      if (!q) return null;
      return new RegExp(reEsc(q), ($('#findCase').checked ? '' : 'i') + flags);
    };
    function findNext() {
      const re = findRegex('g');
      if (!re) { $('#findOut').textContent = 'พิมพ์คำที่ต้องการค้นหา'; return; }
      const text = docText(quill.getContents());
      const r = ed.range();
      re.lastIndex = r.index + r.length;
      let m = re.exec(text);
      if (!m) { re.lastIndex = 0; m = re.exec(text); }
      if (!m) { $('#findOut').textContent = `ไม่พบในหน้า ${ed.currentPage() + 1}`; return; }
      ed.useRange({ index: m.index, length: m[0].length });
      const count = (text.match(findRegex('g')) || []).length;
      $('#findOut').textContent = `พบ ${count} แห่งในหน้านี้`;
    }

    // ----- ธีม -----
    $('#themeList').innerHTML = Object.entries(THEMES).map(([k, t]) => `<button type="button" class="tt-theme" role="radio" data-theme="${k}" aria-checked="false"><i style="background:${t.paper};box-shadow:inset 0 0 0 5px ${t.page}"></i>${esc(t.label)}</button>`).join('');
    let themeKey = store.get(KEY.theme) || 'white';
    let warmth = +(store.get(KEY.warmth) || 0);
    const warmthIn = $('#warmth');
    warmthIn.value = warmth;
    $('#warmthVal').textContent = `${warmth}%`;
    applyTheme(themeKey, warmth);
    $('#themeList').addEventListener('click', e => {
      const b = e.target.closest('[data-theme]');
      if (!b) return;
      themeKey = b.dataset.theme;
      store.set(KEY.theme, themeKey);
      applyTheme(themeKey, warmth);
    });
    warmthIn.addEventListener('input', () => {
      warmth = +warmthIn.value;
      $('#warmthVal').textContent = `${warmth}%`;
      store.set(KEY.warmth, String(warmth));
      applyTheme(themeKey, warmth);
    });

    // ----- ปุ่มทั้งหมดในแผง -----
    const actions = {
      toThai() {
        const scope = $('#numScope').value;
        const label = scopeLabel(scope === 'selection' && !ed.range().length ? 'page' : scope);
        const n = replaceInDoc(scope, /[0-9]+/g, toThai);
        say(n ? `แปลงเป็นเลขไทยแล้ว ${n} จุด (${label})` : `ไม่พบเลขอารบิกใน${label}`);
      },
      toArabic() {
        const scope = $('#numScope').value;
        const label = scopeLabel(scope === 'selection' && !ed.range().length ? 'page' : scope);
        const n = replaceInDoc(scope, /[๐-๙]+/g, toArabic);
        say(n ? `แปลงเป็นเลขอารบิกแล้ว ${n} จุด (${label})` : `ไม่พบเลขไทยใน${label}`);
      },
      bahtFromSel() {
        const r = ed.range();
        if (!r.length) { say('เลือกตัวเลขจำนวนเงินในเอกสารก่อน'); return; }
        bahtIn.value = quill.getText(r.index, r.length).trim();
        updateBaht();
      },
      bahtInsert() {
        const t = updateBaht();
        if (t) insertAtCursor(`(${t})`); else say('กรอกจำนวนเงินให้ถูกต้องก่อน');
      },
      bahtInsertBoth() {
        const t = updateBaht();
        if (t) insertAtCursor(`${fmtMoney(bahtIn.value)} บาท (${t})`); else say('กรอกจำนวนเงินให้ถูกต้องก่อน');
      },
      dateToday() { insertAtCursor(thaiDate(new Date(), $('#dateThaiDigits').checked)); },
      scan() { renderScan(); },
      fixAll() {
        // แก้จากท้ายไปต้น ตำแหน่งของจุดที่เหลือจะได้ไม่เลื่อน
        let n = 0;
        for (const item of [...scanItems].sort((a, b) => b.index - a.index)) if (applyFix(item)) n++;
        say(`แก้แล้ว ${n} จุด (กดย้อนกลับได้)`);
        renderScan();
      },
      ai(btn) { runAi(btn.dataset.task); },
      aiReplace() {
        const text = $('#aiResult').dataset.text;
        if (!text) return;
        const r = aiRange || { index: 0, length: Math.max(0, quill.getLength() - 1) };
        if (!aiRange && !confirm('แทนที่ข้อความทั้งหน้าด้วยฉบับที่ AI เรียบเรียง? (กดย้อนกลับได้ รูปแบบพิเศษในหน้าอาจหาย)')) return;
        replaceRangeWithText(r, text);
        say('แทนที่ด้วยข้อความที่เรียบเรียงแล้ว (กดย้อนกลับได้)');
      },
      aiInsert() { const t = $('#aiResult').dataset.text; if (t) insertAtCursor(t); },
      aiInsertSummary() { const t = $('#aiResult').dataset.text; if (t) insertAtCursor(t); },
      findNext,
      replaceAll() {
        const re = findRegex('g');
        if (!re) { $('#findOut').textContent = 'พิมพ์คำที่ต้องการค้นหา'; return; }
        const scope = $('#findScope').value;
        if (scope === 'all' && !confirm('แทนที่ในทุกหน้า? หน้าที่ไม่ได้เปิดอยู่จะกดย้อนกลับไม่ได้')) return;
        const n = replaceInDoc(scope, re, $('#replText').value);
        $('#findOut').textContent = n ? `แทนที่แล้ว ${n} แห่ง` : 'ไม่พบคำที่ค้นหา';
      },
    };
    panel.addEventListener('click', e => {
      const b = e.target.closest('[data-act]');
      if (b && !b.disabled && actions[b.dataset.act]) actions[b.dataset.act](b);
    });
    $('#findText').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); findNext(); } });

    // ================= สไตล์ (preset) =================
    const BLOCK_KEYS = ['align', 'textalignlast', 'textindent', 'marginleft', 'marginright', 'lineheight', 'spacebefore', 'spaceafter'];
    const INLINE_KEYS = ['font', 'size', 'bold', 'italic', 'underline', 'letterspacing'];
    const fullBlock = o => Object.fromEntries(BLOCK_KEYS.map(k => [k, o[k] ?? false]));
    const SAMPLE_PRESETS = [
      { name: 'เนื้อความหนังสือราชการ', block: fullBlock({ align: 'justify', textindent: '2.5cm', lineheight: '100%' }), inline: { font: false, size: false, bold: false } },
      { name: 'เนื้อความ (ค่าเริ่มต้นเดิม)', block: fullBlock({ align: 'justify', textindent: '45px', lineheight: '150%' }), inline: { font: false, size: false, bold: false } },
      { name: 'ชื่อหนังสือ “บันทึกข้อความ”', block: fullBlock({ align: 'center', textindent: '0cm', lineheight: '100%' }), inline: { size: '29pt', bold: true } },
      { name: 'หัวข้อ (ตัวหนา ชิดซ้าย)', block: fullBlock({ align: 'left', textindent: '0cm', lineheight: '100%', spacebefore: '6pt' }), inline: { size: false, bold: true } },
      { name: 'เรื่อง / เรียน / อ้างถึง', block: fullBlock({ align: 'left', textindent: '0cm', lineheight: '100%' }), inline: { bold: false } },
      { name: 'ข้อย่อยแบบแขวน', block: fullBlock({ align: 'justify', marginleft: '3.25cm', textindent: '-0.75cm', lineheight: '100%' }), inline: {} },
      { name: 'คำลงท้าย / ลงชื่อ', block: fullBlock({ align: 'center', marginleft: '7cm', textindent: '0cm', lineheight: '100%' }), inline: { bold: false } },
    ];
    const ALIGN_TH = { left: 'ชิดซ้าย', center: 'กึ่งกลาง', right: 'ชิดขวา', justify: 'เสมอหน้าหลัง' };
    function describePreset(p) {
      const b = p.block, i = p.inline, bits = [];
      if (b.align) bits.push(b.textalignlast ? 'กระจายแบบไทย' : ALIGN_TH[b.align]);
      if (b.marginleft) bits.push(`กั้นซ้าย ${b.marginleft}`);
      if (b.textindent) bits.push(`ย่อหน้า ${b.textindent}`);
      if (b.marginright) bits.push(`กั้นขวา ${b.marginright}`);
      if (b.lineheight) bits.push(`บรรทัด ${b.lineheight}`);
      if (b.spacebefore) bits.push(`เว้นก่อน ${b.spacebefore}`);
      if (b.spaceafter) bits.push(`เว้นหลัง ${b.spaceafter}`);
      if (i.font) bits.push(i.font);
      if (i.size) bits.push(i.size);
      if (i.bold) bits.push('ตัวหนา');
      if (i.italic) bits.push('ตัวเอียง');
      if (i.underline) bits.push('ขีดเส้นใต้');
      if (i.letterspacing) bits.push(`ช่องไฟ ${i.letterspacing}`);
      return bits.join(' · ') || 'ค่าเริ่มต้น';
    }
    let myPresets = (() => { try { return JSON.parse(store.get(KEY.presets) || '[]'); } catch { return []; } })();
    const allPresets = () => [...SAMPLE_PRESETS, ...myPresets.map(p => ({ ...p, custom: true }))];
    function renderPresets() {
      $('#presetList').innerHTML = allPresets().map((p, i) => `<li><button type="button" class="preset" data-preset="${i}"><strong>${esc(p.name)}</strong><small>${esc(describePreset(p))}</small></button>${p.custom ? `<button type="button" class="del" data-preset-del="${i - SAMPLE_PRESETS.length}" title="ลบสไตล์นี้" aria-label="ลบสไตล์ ${esc(p.name)}">×</button>` : ''}</li>`).join('');
    }
    function applyPreset(p) {
      const r = ed.useRange(ed.range());
      quill.formatLine(r.index, r.length, p.block, 'user');
      if (Object.keys(p.inline).length) {
        const lines = r.length ? quill.getLines(r.index, r.length) : [quill.getLine(r.index)[0]];
        for (const line of lines.filter(Boolean)) {
          const len = line.length() - 1;
          if (len > 0) quill.formatText(quill.getIndex(line), len, p.inline, 'user');
        }
      }
      ed.useRange(r);
      say(`ใช้สไตล์ “${p.name}” แล้ว (กดย้อนกลับได้)`);
    }
    function capturePreset(name) {
      const [line] = quill.getLine(ed.range().index);
      if (!line) return null;
      const block = fullBlock(line.formats());
      const fi = quill.getFormat(quill.getIndex(line), Math.max(1, line.length() - 1));
      const inline = {};
      for (const k of INLINE_KEYS) if (!Array.isArray(fi[k])) inline[k] = fi[k] ?? false;
      return { name, block, inline };
    }
    $('#presetList').addEventListener('click', e => {
      const del = e.target.closest('[data-preset-del]');
      if (del) {
        const i = +del.dataset.presetDel;
        if (!confirm(`ลบสไตล์ “${myPresets[i].name}”?`)) return;
        myPresets.splice(i, 1);
        store.set(KEY.presets, JSON.stringify(myPresets));
        renderPresets();
        return;
      }
      const b = e.target.closest('[data-preset]');
      if (b) applyPreset(allPresets()[+b.dataset.preset]);
    });
    renderPresets();

    // ================= คำที่ใช้บ่อย =================
    const SAMPLE_PHRASES = [
      'ส่วนราชการ', 'ที่', 'วันที่', 'เรื่อง', 'เรียน', 'อ้างถึง', 'สิ่งที่ส่งมาด้วย', 'ตามที่', 'ด้วย', 'ในการนี้', 'ทั้งนี้',
      'ทั้งนี้ ตั้งแต่บัดนี้เป็นต้นไป', 'จึงเรียนมาเพื่อโปรดทราบ', 'จึงเรียนมาเพื่อโปรดพิจารณา',
      'จึงเรียนมาเพื่อโปรดพิจารณา หากเห็นชอบโปรดลงนามในหนังสือที่แนบมาพร้อมนี้', 'จึงเรียนมาเพื่อโปรดทราบและพิจารณาดำเนินการต่อไป',
      'จึงเรียนมาเพื่อโปรดทราบ และแจ้งผู้เกี่ยวข้องทราบต่อไปด้วย', 'ขอแสดงความนับถือ', 'ขอแสดงความนับถืออย่างยิ่ง',
      '(ลงชื่อ) ....................................................',
    ];
    let phrases = (() => { try { return JSON.parse(store.get(KEY.phrases) || 'null') || [...SAMPLE_PHRASES]; } catch { return [...SAMPLE_PHRASES]; } })();
    const savePhrases = () => store.set(KEY.phrases, JSON.stringify(phrases));
    function renderPhrases() {
      $('#phraseList').innerHTML = phrases.map((p, i) => `<li><button type="button" class="phrase" data-phrase="${i}" title="แทรก: ${esc(p)}">${esc(p)}</button><button type="button" class="del" data-phrase-del="${i}" title="ลบ" aria-label="ลบ ${esc(p)}">×</button></li>`).join('')
        || '<li class="hint">ยังไม่มีรายการ</li>';
    }
    function addPhrase(text) {
      const t = String(text || '').replace(/\s*\n\s*/g, ' ').trim();
      if (!t) { say('ยังไม่มีข้อความให้เพิ่ม'); return; }
      if (phrases.includes(t)) { say('มีคำนี้อยู่แล้ว'); return; }
      phrases.unshift(t);
      savePhrases();
      renderPhrases();
      say('เพิ่มคำที่ใช้บ่อยแล้ว');
    }
    $('#phraseList').addEventListener('click', e => {
      const del = e.target.closest('[data-phrase-del]');
      if (del) { phrases.splice(+del.dataset.phraseDel, 1); savePhrases(); renderPhrases(); return; }
      const b = e.target.closest('[data-phrase]');
      if (b) insertAtCursor(phrases[+b.dataset.phrase]);
    });
    $('#phraseIn').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); addPhrase(e.target.value); e.target.value = ''; } });
    renderPhrases();

    // ================= ไฟล์ =================
    let recent = [];
    async function renderRecent() {
      $('#fileName').value = ed.docName();
      $('#fileLinkBtn').hidden = !ed.canUseFiles;
      if (!ed.canUseFiles) {
        $('#recentList').innerHTML = '';
        $('#recentHint').textContent = 'เบราว์เซอร์นี้จำไฟล์ล่าสุดไม่ได้ — ใช้ Chrome หรือ Edge หรือกด “เปิดไฟล์…” แล้วเลือกไฟล์ที่ดาวน์โหลดไว้';
        return;
      }
      recent = await ed.recentFiles();
      $('#recentList').innerHTML = recent.map((r, i) => `<li><strong>${esc(r.name)}</strong><span class="n">ใช้ล่าสุด ${esc(new Date(r.time).toLocaleString('th-TH'))}</span>
        <span class="acts"><button type="button" data-recent="${i}" data-do="open">เปิด</button><button type="button" data-recent="${i}" data-do="del">ลบจากรายการ</button></span></li>`).join('');
      $('#recentHint').textContent = recent.length
        ? 'เปิดไฟล์ร่างของแอปนี้แล้ว การแก้ไขจะบันทึกทับไฟล์นั้นอัตโนมัติ (เบราว์เซอร์อาจถามสิทธิ์ก่อน)'
        : 'ยังไม่มี — ไฟล์ที่ผูกด้วย “บันทึกลงไฟล์อัตโนมัติ” หรือเปิดด้วย “เปิดไฟล์…” จะแสดงที่นี่ (ไฟล์ที่ดาวน์โหลดให้เปิดด้วย “เปิดไฟล์…”)';
    }
    onTab.file = renderRecent;
    document.addEventListener('eoffice:recent', () => { if (!$('section[data-panel="file"]', panel).hidden) renderRecent(); });
    $('#recentList').addEventListener('click', async e => {
      const b = e.target.closest('[data-recent]');
      if (!b) return;
      const r = recent[+b.dataset.recent];
      if (b.dataset.do === 'open') { if (await ed.openHandle(r.handle)) renderRecent(); }
      else { await ed.removeRecent(+b.dataset.recent); renderRecent(); }
    });
    $('#fileName').addEventListener('change', e => ed.setDocName(e.target.value));
    $('#docName').addEventListener('change', () => { $('#fileName').value = ed.docName(); });

    // ================= แม่แบบบันทึกข้อความ =================
    // ตามระเบียบสำนักนายกรัฐมนตรีว่าด้วยงานสารบรรณ: หัว "บันทึกข้อความ" 29 pt ตัวหนา, หัวข้อ ส่วนราชการ/ที่/วันที่/เรื่อง 20 pt ตัวหนา,
    // เนื้อความ 16 pt ย่อหน้า 2.5 ซม. บรรทัดเดี่ยว, ลงชื่อกึ่งกลางครึ่งขวาของหน้า
    const GARUDA_KEY = 'eoffice-editor:garuda';
    const CM_PX = 96 / 2.54;
    // ใช้ตัวเลขแบบเดียวกับเลขรายการที่ตั้งไว้ (เลขไทย/อารบิก)
    const today = () => thaiDate(new Date(), ed.settings().listDigits === 'thai');
    const garudaEmbed = heightCm => {
      const src = store.get(GARUDA_KEY);
      return src ? [{ insert: { image: src }, attributes: { height: String(Math.round(heightCm * CM_PX)) } }] : [{ insert: '(ครุฑ)' }];
    };
    function memoTemplate() {
      const head = { align: 'left', textindent: '0cm', lineheight: '100%' };
      const body = { align: 'justify', textindent: '2.5cm', lineheight: '100%' };
      const sign = { align: 'center', marginleft: '7cm', textindent: '0cm', lineheight: '100%' };
      const label = { bold: true, size: '20pt' };
      // ครุฑ 1.5 ซม. ชิดซ้าย แล้วใช้จุดแท็บ 5.9 ซม. ให้ “บันทึกข้อความ” (29 pt กว้างราว 4.1 ซม.) อยู่กึ่งกลางบรรทัด 16 ซม.
      return new Delta([...garudaEmbed(1.5), { insert: '\t' }])
        .insert('บันทึกข้อความ', { bold: true, size: '29pt' }).insert('\n', { ...head, tabsize: '5.9cm' })
        .insert('ส่วนราชการ', label).insert('  (ชื่อส่วนราชการเจ้าของเรื่อง โทร. ...)').insert('\n', head)
        .insert('ที่', label).insert('  (รหัส)/\t').insert('วันที่', label).insert(`  ${today()}`).insert('\n', { ...head, tabsize: '8cm' })
        .insert('เรื่อง', label).insert('  (สรุปเรื่องสั้น ๆ ที่สุด)').insert('\n', head)
        .insert('\n', head)
        .insert('เรียน  (ตำแหน่งของผู้รับ)').insert('\n', head)
        .insert('\n', head)
        .insert('(ส่วนเหตุ) ตามที่/ด้วย ...').insert('\n', body)
        .insert('(ส่วนรายละเอียด) ...').insert('\n', body)
        .insert('(ส่วนความประสงค์) จึงเรียนมาเพื่อโปรดพิจารณา').insert('\n', body)
        .insert('\n', head)
        .insert('\n', head)
        .insert('(....................................................)').insert('\n', sign)
        .insert('(ตำแหน่ง)').insert('\n', sign);
    }
    // หนังสือภายนอก: ครุฑ 3 ซม. กึ่งกลาง, ที่ ชิดซ้าย + ส่วนราชการผู้ส่งเริ่มที่ 10 ซม., วันที่เริ่มกลางหน้า,
    // ขอแสดงความนับถือ/ลงชื่อ/ตำแหน่ง ที่ครึ่งขวา, ท้ายหนังสือเป็นส่วนราชการเจ้าของเรื่องและช่องทางติดต่อ
    function letterTemplate() {
      const head = { align: 'left', textindent: '0cm', lineheight: '100%' };
      const addr = { ...head, marginleft: '10cm' };
      const body = { align: 'justify', textindent: '2.5cm', lineheight: '100%' };
      const sign = { align: 'center', marginleft: '7cm', textindent: '0cm', lineheight: '100%' };
      return new Delta(garudaEmbed(3)).insert('\n', { ...head, align: 'center' })
        .insert('ที่ (รหัส)/\t(ส่วนราชการเจ้าของหนังสือ)').insert('\n', { ...head, tabsize: '10cm' })
        .insert('(ที่ตั้ง)').insert('\n', addr)
        .insert('(รหัสไปรษณีย์)').insert('\n', addr)
        .insert('\n', head)
        .insert(today()).insert('\n', { ...head, marginleft: '8cm' })
        .insert('\n', head)
        .insert('เรื่อง  (สรุปเรื่องสั้น ๆ ที่สุด)').insert('\n', head)
        .insert('\n', head)
        .insert('เรียน  (คำขึ้นต้นและตำแหน่งผู้รับ)').insert('\n', head)
        .insert('\n', head)
        .insert('อ้างถึง  (ถ้ามี)').insert('\n', head)
        .insert('\n', head)
        .insert('สิ่งที่ส่งมาด้วย  (ถ้ามี)').insert('\n', head)
        .insert('\n', head)
        .insert('(ส่วนเหตุ) ตามที่/ด้วย ...').insert('\n', body)
        .insert('(ส่วนรายละเอียด) ...').insert('\n', body)
        .insert('(ส่วนความประสงค์) จึงเรียนมาเพื่อโปรดพิจารณา').insert('\n', body)
        .insert('\n', head)
        .insert('ขอแสดงความนับถือ').insert('\n', sign)
        .insert('\n', head)
        .insert('\n', head)
        .insert('\n', head)
        .insert('(ชื่อเต็มผู้ลงนาม)').insert('\n', sign)
        .insert('(ตำแหน่ง)').insert('\n', sign)
        .insert('\n', head)
        .insert('\n', head)
        .insert('(ส่วนราชการเจ้าของเรื่อง)').insert('\n', head)
        .insert('โทร. ...').insert('\n', head)
        .insert('ไปรษณีย์อิเล็กทรอนิกส์ ...').insert('\n', head);
    }
    function insertTemplate(delta, name, firstPlaceholder) {
      if (quill.getLength() > 1) {
        if (!confirm(`หน้านี้มีข้อความอยู่แล้ว — สร้าง${name}ในหน้าใหม่ต่อจากหน้านี้หรือไม่?`)) return;
        ed.addPage();
      }
      quill.setContents(delta, 'user');
      // เลือกข้อความในวงเล็บแรกไว้ให้พิมพ์ทับได้ทันที
      const text = docText(quill.getContents()), i = text.indexOf(firstPlaceholder);
      if (i >= 0) quill.setSelection(i, text.indexOf(')', i) - i + 1, 'user');
      say(`สร้างแม่แบบ${name}แล้ว — พิมพ์ทับข้อความในวงเล็บได้เลย`);
    }

    // ----- การตั้งค่าเอกสาร / จุดแท็บ / รูปครุฑ -----
    const setProfile = $('#setProfile');
    setProfile.innerHTML = Object.entries(ed.profiles).map(([k, p]) => `<option value="${k}">${esc(p.label)}</option>`).join('');
    function renderSettings() {
      const s = ed.settings();
      setProfile.value = s.profile;
      $('#setListDigits').value = s.listDigits;
      $('#setPageNum').checked = s.pageNumbers;
      $('#setPageDigits').value = s.pageDigits;
      $('#setPageDigits').disabled = !s.pageNumbers;
      const g = store.get(GARUDA_KEY);
      $('#garudaPrev').hidden = !g;
      if (g) $('#garudaPrev').src = g;
      $('#garudaClear').disabled = !g;
    }
    setProfile.addEventListener('change', () => { ed.setSetting('profile', setProfile.value); say(`มาตรฐานเอกสาร: ${ed.profiles[setProfile.value].label}`); });
    $('#setListDigits').addEventListener('change', e => ed.setSetting('listDigits', e.target.value));
    $('#setPageNum').addEventListener('change', e => { ed.setSetting('pageNumbers', e.target.checked); renderSettings(); });
    $('#setPageDigits').addEventListener('change', e => ed.setSetting('pageDigits', e.target.value));
    $('#tabChips').addEventListener('mousedown', e => e.preventDefault());
    $('#tabChips').addEventListener('click', e => {
      const b = e.target.closest('[data-tabsize]');
      if (!b) return;
      const r = ed.useRange(ed.range());
      quill.formatLine(r.index, r.length, 'tabsize', b.dataset.tabsize || false, 'user');
      say(b.dataset.tabsize ? `ตั้งจุดแท็บ ${b.textContent} ให้ย่อหน้านี้แล้ว` : 'ใช้จุดแท็บปกติ');
      markTabChip();
    });
    function markTabChip() {
      const [line] = quill.getLine(ed.range().index);
      const cur = line?.formats().tabsize || '';
      for (const c of $$('#tabChips [data-tabsize]')) c.setAttribute('aria-checked', String(c.dataset.tabsize === cur));
    }
    quill.on('selection-change', r => { if (r && !$('section[data-panel="style"]', panel).hidden) markTabChip(); });
    onTab.style = () => { renderSettings(); markTabChip(); };
    $('#garudaFile').addEventListener('change', e => {
      const f = e.target.files[0];
      e.target.value = '';
      if (!f) return;
      if (f.size > 400 * 1024) { say('รูปครุฑใหญ่เกิน 400 KB — ลดขนาดรูปก่อน'); return; }
      const rd = new FileReader();
      rd.onload = () => {
        store.set(GARUDA_KEY, rd.result);
        if (store.get(GARUDA_KEY) !== rd.result) { say('บันทึกรูปครุฑในเบราว์เซอร์ไม่ได้ (พื้นที่เต็ม)'); return; }
        renderSettings();
        say('บันทึกรูปครุฑแล้ว — ใช้กับแม่แบบครั้งถัดไป');
      };
      rd.readAsDataURL(f);
    });

    Object.assign(actions, {
      tplMemo: () => insertTemplate(memoTemplate(), 'บันทึกข้อความ', '(ชื่อส่วนราชการ'),
      tplLetter: () => insertTemplate(letterTemplate(), 'หนังสือภายนอก', '(รหัส)'),
      garudaPick: () => $('#garudaFile').click(),
      garudaClear() { store.del(GARUDA_KEY); renderSettings(); say('ลบรูปครุฑแล้ว'); },
      nbsp: () => insertAtCursor('\u00A0'),
      presetSave() {
        const name = $('#presetName').value.trim();
        if (!name) { say('ตั้งชื่อสไตล์ก่อน'); $('#presetName').focus(); return; }
        const p = capturePreset(name);
        if (!p) return;
        const i = myPresets.findIndex(x => x.name === name);
        if (i >= 0) { if (!confirm(`มีสไตล์ชื่อ “${name}” แล้ว บันทึกทับหรือไม่?`)) return; myPresets[i] = p; }
        else myPresets.push(p);
        store.set(KEY.presets, JSON.stringify(myPresets));
        $('#presetName').value = '';
        renderPresets();
        say(`บันทึกสไตล์ “${name}” แล้ว — ${describePreset(p)}`);
      },
      phraseAdd() { addPhrase($('#phraseIn').value); $('#phraseIn').value = ''; },
      phraseFromSel() {
        const r = ed.range();
        if (!r.length) { say('เลือกข้อความในเอกสารก่อน'); return; }
        addPhrase(quill.getText(r.index, r.length));
      },
      phraseReset() {
        if (!confirm('คืนค่ารายการคำที่ใช้บ่อยเป็นชุดตัวอย่าง? คำที่เพิ่มเองจะหายไป')) return;
        phrases = [...SAMPLE_PHRASES];
        savePhrases();
        renderPhrases();
      },
      fileDownload() {
        if ($('#fileName').value.trim() !== ed.docName()) ed.setDocName($('#fileName').value);
        ed.saveDraft();
        $('#fileName').value = ed.docName();
      },
      fileLink() {
        if ($('#fileName').value.trim() !== ed.docName()) ed.setDocName($('#fileName').value);
        ed.fileAutosave();
      },
      fileOpen() { ed.openPicker(); },
    });

    // ================= เครื่องคิดเลข =================
    // ตัวแยกนิพจน์เอง (ไม่ใช้ eval): + − × ÷ วงเล็บ ทศนิยม และ % (หาร 100) รับเลขไทยและเครื่องหมายคั่นหลักพันได้
    function calculate(input) {
      const s = toArabic(String(input)).replace(/[,\s]/g, '').replace(/×|x|X/g, '*').replace(/÷/g, '/').replace(/−|–/g, '-');
      if (!s) return null;
      let i = 0;
      const fail = msg => { throw new Error(msg); };
      const number = () => {
        const m = /^(\d+\.?\d*|\.\d+)/.exec(s.slice(i));
        if (!m) fail('รูปแบบไม่ถูกต้อง');
        i += m[0].length;
        return parseFloat(m[0]);
      };
      const factor = () => {
        if (s[i] === '+') { i++; return factor(); }
        if (s[i] === '-') { i++; return -factor(); }
        let v;
        if (s[i] === '(') { i++; v = sum(); if (s[i++] !== ')') fail('วงเล็บไม่ครบ'); }
        else v = number();
        while (s[i] === '%') { i++; v /= 100; }
        return v;
      };
      const product = () => {
        let v = factor();
        while (s[i] === '*' || s[i] === '/') {
          const op = s[i++], r = factor();
          if (op === '/' && r === 0) fail('หารด้วยศูนย์ไม่ได้');
          v = op === '*' ? v * r : v / r;
        }
        return v;
      };
      const sum = () => {
        let v = product();
        while (s[i] === '+' || s[i] === '-') { const op = s[i++], r = product(); v = op === '+' ? v + r : v - r; }
        return v;
      };
      const v = sum();
      if (i < s.length) fail('รูปแบบไม่ถูกต้อง');
      if (!Number.isFinite(v)) fail('ผลลัพธ์ใหญ่เกินไป');
      return +v.toFixed(10);
    }
    const calcBox = document.createElement('div');
    calcBox.className = 'calc';
    calcBox.id = 'calc';
    calcBox.hidden = true;
    calcBox.setAttribute('role', 'dialog');
    calcBox.setAttribute('aria-label', 'เครื่องคิดเลข');
    const KEYS = ['C', '(', ')', '⌫', '7', '8', '9', '÷', '4', '5', '6', '×', '1', '2', '3', '−', '0', '.', '%', '+'];
    calcBox.innerHTML = `
<div class="calc-head" id="calcHead"><span>🧮 เครื่องคิดเลข</span><button type="button" class="tools-x" id="calcClose" aria-label="ปิดเครื่องคิดเลข">×</button></div>
<input type="text" id="calcIn" class="calc-in" inputmode="decimal" autocomplete="off" spellcheck="false" placeholder="เช่น 1,250×12+7%" aria-label="นิพจน์">
<div class="calc-out" id="calcOut" aria-live="polite">0</div>
<div class="calc-keys">${KEYS.map(k => `<button type="button" data-key="${k}"${/[÷×−+%()]/.test(k) ? ' class="op"' : k === 'C' || k === '⌫' ? ' class="fn"' : ''}>${k}</button>`).join('')}<button type="button" data-key="=" class="eq">=</button></div>
<div class="row"><button type="button" class="tbtn" data-calc="insert">แทรกผลลัพธ์</button><button type="button" class="tbtn" data-calc="copy">คัดลอก</button><label class="inline"><input type="checkbox" id="calcComma" checked> คั่นหลักพัน</label></div>
<ul class="calc-hist" id="calcHist"></ul>`;
    document.body.append(calcBox);
    const calcIn = $('#calcIn'), calcOut = $('#calcOut');
    let calcResult = null;
    const fmtNum = v => ($('#calcComma').checked ? v.toLocaleString('en-US', { maximumFractionDigits: 10 }) : String(v));
    function calcPreview() {
      try {
        calcResult = calculate(calcIn.value);
        calcOut.textContent = calcResult === null ? '0' : `= ${fmtNum(calcResult)}`;
        calcOut.classList.remove('err');
      } catch (err) {
        calcResult = null;
        calcOut.textContent = err.message;
        calcOut.classList.add('err');
      }
    }
    function calcEquals() {
      calcPreview();
      if (calcResult === null) return;
      const li = document.createElement('li');
      li.textContent = `${calcIn.value} = ${fmtNum(calcResult)}`;
      li.dataset.value = String(calcResult);
      $('#calcHist').prepend(li);
      while ($('#calcHist').children.length > 6) $('#calcHist').lastChild.remove();
      calcIn.value = String(calcResult);
    }
    calcIn.addEventListener('input', calcPreview);
    calcIn.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); calcEquals(); }
      if (e.key === 'Escape') calcBox.hidden = true;
    });
    $('#calcComma').addEventListener('change', calcPreview);
    calcBox.addEventListener('mousedown', e => { if (e.target.closest('.calc-keys button, [data-calc], .calc-hist li')) e.preventDefault(); });
    calcBox.addEventListener('click', e => {
      const k = e.target.closest('[data-key]')?.dataset.key;
      if (k) {
        if (k === 'C') calcIn.value = '';
        else if (k === '⌫') calcIn.value = calcIn.value.slice(0, -1);
        else if (k === '=') { calcEquals(); return; }
        else calcIn.value += k;
        calcPreview();
        return;
      }
      const act = e.target.closest('[data-calc]')?.dataset.calc;
      if (act && calcResult !== null) {
        const text = fmtNum(calcResult);
        if (act === 'insert') insertAtCursor(text);
        else navigator.clipboard?.writeText(text).then(() => say(`คัดลอก ${text} แล้ว`), () => say('คัดลอกไม่ได้'));
      }
      const h = e.target.closest('.calc-hist li');
      if (h) { calcIn.value = h.dataset.value; calcPreview(); }
    });
    $('#calcClose').addEventListener('click', () => { calcBox.hidden = true; });
    $('#btnCalc').addEventListener('click', () => {
      calcBox.hidden = !calcBox.hidden;
      if (!calcBox.hidden) calcIn.focus();
    });
    // ลากย้ายหน้าต่างเครื่องคิดเลขด้วยแถบหัว
    $('#calcHead').addEventListener('pointerdown', e => {
      if (e.target.closest('button')) return;
      const rect = calcBox.getBoundingClientRect(), dx = e.clientX - rect.left, dy = e.clientY - rect.top;
      const move = ev => {
        calcBox.style.left = `${Math.min(Math.max(0, ev.clientX - dx), innerWidth - rect.width)}px`;
        calcBox.style.top = `${Math.min(Math.max(0, ev.clientY - dy), innerHeight - 40)}px`;
        calcBox.style.right = 'auto';
      };
      const up = () => { removeEventListener('pointermove', move); removeEventListener('pointerup', up); };
      addEventListener('pointermove', move);
      addEventListener('pointerup', up);
    });

    // ================= กระดาษทด (ด้านซ้าย) =================
    // เก็บใน sessionStorage: รีเฟรชหน้าแล้วยังอยู่ แต่หายเมื่อปิดแท็บ/หน้าต่าง (ไม่เก็บถาวร)
    const scratch = document.createElement('aside');
    scratch.className = 'scratch';
    scratch.id = 'scratch';
    scratch.hidden = true;
    scratch.setAttribute('aria-label', 'กระดาษทด');
    scratch.innerHTML = `
<div class="tools-head"><h2>📝 กระดาษทด</h2><span class="head-acts"><button type="button" class="tbtn" id="scratchAdd">＋ เพิ่มแผ่น</button><button type="button" class="tools-x" id="scratchClose" aria-label="ปิดกระดาษทด" title="ปิด">×</button></span></div>
<p class="hint pad">จดชั่วคราว — รีเฟรชหน้าแล้วยังอยู่ แต่จะหายเมื่อปิดแท็บนี้</p>
<div class="scratch-body" id="scratchBody"></div>`;
    document.body.append(scratch);
    const SCRATCH_KEY = 'eoffice-editor:scratch';
    const session = {
      get() { try { return JSON.parse(sessionStorage.getItem(SCRATCH_KEY) || 'null'); } catch { return null; } },
      set(v) { try { sessionStorage.setItem(SCRATCH_KEY, JSON.stringify(v)); } catch { /* ไม่มี storage */ } },
    };
    const savedScratch = session.get();
    const notes = Array.isArray(savedScratch?.notes) ? savedScratch.notes.map(String) : [];
    const saveNotes = () => session.set({ open: !scratch.hidden, notes });
    function renderNotes(focusIndex) {
      $('#scratchBody').innerHTML = notes.map((n, i) => `<div class="note">
        <div class="note-head"><span>แผ่นที่ ${i + 1}</span><span><button type="button" data-note-ins="${i}" title="แทรกข้อความแผ่นนี้ลงในเอกสารที่เคอร์เซอร์">แทรกในเอกสาร</button><button type="button" data-note-del="${i}" title="ลบแผ่นนี้" aria-label="ลบแผ่นที่ ${i + 1}">×</button></span></div>
        <textarea data-note="${i}" aria-label="กระดาษทดแผ่นที่ ${i + 1}" placeholder="จดอะไรก็ได้…">${esc(n)}</textarea></div>`).join('');
      if (focusIndex != null) $(`[data-note="${focusIndex}"]`)?.focus();
    }
    function addNote() { notes.push(''); renderNotes(notes.length - 1); saveNotes(); }
    $('#scratchBody').addEventListener('input', e => { const t = e.target.closest('[data-note]'); if (t) { notes[+t.dataset.note] = t.value; saveNotes(); } });
    $('#scratchBody').addEventListener('mousedown', e => { if (e.target.closest('[data-note-ins]')) e.preventDefault(); });
    $('#scratchBody').addEventListener('click', e => {
      const ins = e.target.closest('[data-note-ins]');
      if (ins) { const t = notes[+ins.dataset.noteIns].trim(); if (t) insertAtCursor(t); else say('แผ่นนี้ยังว่าง'); return; }
      const del = e.target.closest('[data-note-del]');
      if (del) {
        const i = +del.dataset.noteDel;
        if (notes[i].trim() && !confirm(`ลบกระดาษทดแผ่นที่ ${i + 1}?`)) return;
        notes.splice(i, 1);
        if (!notes.length) notes.push('');
        renderNotes();
        saveNotes();
      }
    });
    $('#scratchAdd').addEventListener('click', addNote);
    function setScratch(open) {
      scratch.hidden = !open;
      saveNotes();
      document.body.classList.toggle('scratch-open', open);
      if (open && !notes.length) addNote();
      window.dispatchEvent(new Event('resize'));
    }
    $('#btnScratch').addEventListener('click', () => setScratch(scratch.hidden));
    $('#scratchClose').addEventListener('click', () => setScratch(false));
    renderNotes();
    if (savedScratch?.open) setScratch(true);

    // ================= มุมมอง: หน้ากระดาษ / เส้นบรรทัด / เครื่องหมายซ่อน =================
    const MARK_TYPES = [
      ['space', 'เว้นวรรค', '·'], ['nbsp', 'เว้นวรรคห้ามตัดบรรทัด', '°'], ['tab', 'แท็บ', '→'], ['para', 'ท้ายย่อหน้า', '¶'],
      ['dup', 'ระบายเว้นวรรคซ้อน / ท้ายย่อหน้า', '▭'], ['wj', 'ตัวผูกคำห้ามตัด (Word Joiner)', '⁀'], ['zw', 'อักขระล่องหน (zero-width)', '⌀'],
    ];
    function renderView() {
      const s = ed.settings();
      $('#vwPaper').checked = s.paper === 'a4';
      $('#vwBreaks').checked = s.pageBreaks;
      $('#vwBreaks').disabled = s.paper !== 'a4';
      $('#vwGrid').checked = s.grid;
      $('#vwLines').checked = s.lineGuides;
      $('#vwMarks').checked = s.showMarks;
      $('#vwParaNum').checked = s.paraNumbers;
      $('#vwLineNum').checked = s.lineNumbers;
      $('#vwLineColors').innerHTML = s.lineColors.map((c, i) => `<label class="inline">สี ${i + 1} <input type="color" data-linecolor="${i}" value="${c}"></label>`).join('');
      $('#vwMarkList').innerHTML = MARK_TYPES.map(([k, label, sym]) => `<div class="row"><label class="inline grow"><input type="checkbox" data-mark="${k}"${s.marks[k] ? ' checked' : ''}> <b class="mk-sym" style="color:${s.markColors[k]}">${sym}</b> ${label}</label><input type="color" data-markcolor="${k}" value="${s.markColors[k]}" aria-label="สี${label}"></div>`).join('');
    }
    onTab.theme = renderView;
    $('#vwPaper').addEventListener('change', e => { ed.setSetting('paper', e.target.checked ? 'a4' : 'compact'); renderView(); });
    $('#vwBreaks').addEventListener('change', e => ed.setSetting('pageBreaks', e.target.checked));
    $('#vwGrid').addEventListener('change', e => ed.setSetting('grid', e.target.checked));
    $('#vwLines').addEventListener('change', e => ed.setSetting('lineGuides', e.target.checked));
    $('#vwMarks').addEventListener('change', e => ed.setSetting('showMarks', e.target.checked));
    $('#vwParaNum').addEventListener('change', e => ed.setSetting('paraNumbers', e.target.checked));
    $('#vwLineNum').addEventListener('change', e => ed.setSetting('lineNumbers', e.target.checked));
    $('#vwLineColors').addEventListener('input', e => {
      const i = e.target.dataset.linecolor;
      if (i == null) return;
      const colors = [...ed.settings().lineColors];
      colors[+i] = e.target.value;
      ed.setSetting('lineColors', colors);
    });
    $('#vwMarkList').addEventListener('change', e => {
      const k = e.target.dataset.mark;
      if (k) ed.setSetting('marks', { ...ed.settings().marks, [k]: e.target.checked });
    });
    $('#vwMarkList').addEventListener('input', e => {
      const k = e.target.dataset.markcolor;
      if (!k) return;
      ed.setSetting('markColors', { ...ed.settings().markColors, [k]: e.target.value });
      const sym = e.target.closest('.row').querySelector('.mk-sym');
      if (sym) sym.style.color = e.target.value;
    });

    // ================= คำห้ามตัด =================
    const NOBREAK_KEY = 'eoffice-editor:nobreak';
    const NOBREAK_DEFAULT = [
      'ผู้อำนวยการ', 'รองผู้อำนวยการ', 'ผู้ว่าราชการจังหวัด', 'รองผู้ว่าราชการจังหวัด', 'ปลัดกระทรวง', 'รองปลัดกระทรวง', 'ปลัดจังหวัด',
      'คณะกรรมการ', 'คณะอนุกรรมการ', 'คณะทำงาน', 'คณะรัฐมนตรี', 'นายกรัฐมนตรี', 'รัฐมนตรีว่าการกระทรวง', 'ผู้บังคับบัญชา', 'ผู้ตรวจราชการ',
      'ผู้รับผิดชอบ', 'ผู้เกี่ยวข้อง', 'ผู้ประสานงาน', 'ส่วนราชการ', 'หน่วยงาน', 'สำนักงาน', 'องค์กรปกครองส่วนท้องถิ่น', 'องค์การบริหารส่วนจังหวัด',
      'องค์การบริหารส่วนตำบล', 'เทศบาลนคร', 'เทศบาลเมือง', 'เทศบาลตำบล', 'ปีงบประมาณ', 'งบประมาณ', 'พระราชบัญญัติ', 'พระราชกฤษฎีกา',
      'ระเบียบสำนักนายกรัฐมนตรี', 'ข้าราชการ', 'พนักงานราชการ', 'ลูกจ้างประจำ', 'ประชาสัมพันธ์', 'แผนปฏิบัติราชการ', 'ยุทธศาสตร์',
      'จัดซื้อจัดจ้าง', 'ประสิทธิภาพ', 'ประสิทธิผล', 'ท้องถิ่น', 'ส่งเสริม', 'ราชอาณาจักร', 'พุทธศักราช', 'การเลือกตั้ง',
    ];
    let noBreak = (() => { try { return JSON.parse(store.get(NOBREAK_KEY) || 'null') || [...NOBREAK_DEFAULT]; } catch { return [...NOBREAK_DEFAULT]; } })();
    const syncNoBreak = () => { store.set(NOBREAK_KEY, JSON.stringify(noBreak)); ed.setNoBreakWords(noBreak); renderNoBreak(); };
    function renderNoBreak() {
      $('#nbList').innerHTML = noBreak.map((w, i) => {
        const parts = ed.segmentWords(w);
        return `<li><span class="phrase" title="${parts.length > 1 ? 'เบราว์เซอร์อาจตัดตรง |' : 'Segmenter มองเป็นคำเดียวอยู่แล้ว'}">${esc(w)} <small style="color:${parts.length > 1 ? '#b91c1c' : 'var(--muted)'}">${esc(parts.join('|'))}</small></span><button type="button" class="del" data-nb-del="${i}" title="ลบ" aria-label="ลบ ${esc(w)}">×</button></li>`;
      }).join('') || '<li class="hint">ยังไม่มีรายการ</li>';
    }
    // เพิ่มคำ (ตัดเว้นวรรคออก — วลีที่มีเว้นวรรคให้ใช้ “เว้นวรรคห้ามตัดบรรทัด” แทน)
    function addNoBreak(words, quiet) {
      const add = words.flatMap(w => String(w).split(/[,\n]/)).map(w => w.replace(/[\s\u2060]/g, '')).filter(w => w.length > 1 && !noBreak.includes(w));
      if (!add.length) { if (!quiet) say('ไม่มีคำใหม่ให้เพิ่ม (คำต้องยาวอย่างน้อย 2 ตัวอักษรและยังไม่อยู่ในรายการ)'); return 0; }
      noBreak = [...add, ...noBreak];
      syncNoBreak();
      if (!quiet) say(`เพิ่มคำห้ามตัด ${add.length} คำ — กด “ผูกคำในเอกสาร” เพื่อใช้กับข้อความที่มีอยู่`);
      return add.length;
    }
    // ชื่อคนหลังคำนำหน้า ที่ Segmenter แยกเป็นหลายชิ้นและยังไม่อยู่ในรายการ
    const TITLE_RE = /(?:นางสาว|นาย|นาง|ดร\.|ผศ\.|รศ\.|ศ\.)([ก-๙\u2060]{2,})(?:[ \u00A0]+([ก-๙\u2060]{2,}))?/g;
    function nameCandidates(text) {
      const out = new Set();
      for (const m of text.matchAll(TITLE_RE)) {
        for (const w of [m[1], m[2]]) {
          if (!w || w.includes('\u2060') || w.length > 20) continue;      // too long = other words attached, not a name
          if (!noBreak.includes(w) && ed.segmentWords(w).length > 1) out.add(w);
        }
      }
      return [...out];
    }
    $('#nbList').addEventListener('click', e => {
      const b = e.target.closest('[data-nb-del]');
      if (!b) return;
      noBreak.splice(+b.dataset.nbDel, 1);
      syncNoBreak();
    });
    $('#nbIn').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); addNoBreak([e.target.value]); e.target.value = ''; } });
    Object.assign(actions, {
      nbAdd() { addNoBreak([$('#nbIn').value]); $('#nbIn').value = ''; },
      nbFromSel() {
        const r = ed.range();
        if (!r.length) { say('เลือกคำในเอกสารก่อน'); return; }
        addNoBreak([quill.getText(r.index, r.length)]);
      },
      nbReset() {
        if (!confirm('คืนค่ารายการคำห้ามตัดเป็นชุดตั้งต้น? คำที่เพิ่มเองจะหายไป')) return;
        noBreak = [...NOBREAK_DEFAULT];
        syncNoBreak();
      },
      nbApply() {
        const scope = $('#nbScope').value;
        if (scope === 'all' && !confirm('ผูกคำในทุกหน้า? หน้าที่ไม่ได้เปิดอยู่จะกดย้อนกลับไม่ได้')) return;
        const n = ed.bindWords(scope, noBreak);
        $('#nbStatus').textContent = n ? `ใส่ตัวผูกคำ ${n} จุดแล้ว — กดปุ่ม ¶ เพื่อดูตำแหน่ง (⁀)` : 'ไม่มีจุดที่ต้องผูกเพิ่ม';
      },
      nbRemove() {
        const scope = $('#nbScope').value;
        if (scope === 'all' && !confirm('ถอดตัวผูกคำในทุกหน้า? หน้าที่ไม่ได้เปิดอยู่จะกดย้อนกลับไม่ได้')) return;
        const n = replaceInDoc(scope, /\u2060/g, '');
        $('#nbStatus').textContent = n ? `ถอดตัวผูกคำ ${n} จุดแล้ว` : 'ไม่มีตัวผูกคำในเอกสาร';
      },
    });
    ed.setNoBreakWords(noBreak);
    renderNoBreak();

    // ================= จัดหน้า: ปรับให้สอดคล้องกันทีละย่อหน้า =================
    const ROLE_TH = { title: 'หัวหนังสือ', head: 'ส่วนหัว', body: 'เนื้อความ', list: 'รายการ', sign: 'ลงชื่อ', foot: 'ท้ายหนังสือ', blank: 'บรรทัดว่าง', table: 'ตาราง', other: 'อื่น ๆ' };
    const LEN_RE = /^-?\d+(\.\d+)?(cm|mm|pt|px)$/, LH_RE = /^\d+(\.\d+)?%$/;
    const cmOf = v => { const m = /^(-?\d*\.?\d+)(cm|mm|pt|px)?$/.exec(String(v ?? '').trim()); if (!m) return 0; const n = +m[1]; return { cm: n, mm: n / 10, pt: n * 2.54 / 72 }[m[2]] ?? n / CM_PX; };
    const FMT_TH = { align: 'จัดแนว', textalignlast: 'บรรทัดสุดท้าย', textindent: 'ย่อหน้า', marginleft: 'กั้นซ้าย', lineheight: 'บรรทัด', spacebefore: 'เว้นก่อน', spaceafter: 'เว้นหลัง', letterspacing: 'ช่องไฟ' };
    // หมวดของข้อเสนอ เรียงตามความสำคัญที่เน้น: ตัดคำ > กระจายตัวอักษร > ระยะบรรทัด > รูปแบบอื่น
    const CATS = [
      { key: 'break', label: 'ตัดคำ', test: it => it.bind.length > 0 || (it.join || []).length > 0, color: '#b91c1c', bg: '#fef2f2' },
      { key: 'spread', label: 'กระจายตัวอักษร', test: it => ['letterspacing', 'textalignlast', 'align'].some(k => k in it.changes), color: '#6d28d9', bg: '#f5f3ff' },
      { key: 'line', label: 'ระยะบรรทัด', test: it => ['lineheight', 'spacebefore', 'spaceafter'].some(k => k in it.changes), color: '#047857', bg: '#ecfdf5' },
      { key: 'form', label: 'รูปแบบ', test: it => ['textindent', 'marginleft'].some(k => k in it.changes), color: '#475569', bg: '#f1f5f9' },
    ];
    const catsOf = it => CATS.filter(c => c.test(it));
    const priority = it => CATS.reduce((sum, c, i) => sum + (c.test(it) ? 2 ** (CATS.length - i) : 0), 0);
    const signedEm = v => `${v > 0 ? '+' : v < 0 ? '−' : ''}${Math.abs(v).toFixed(3)} em`;
    const ALIGN_TXT = { left: 'ชิดซ้าย', center: 'กึ่งกลาง', right: 'ชิดขวา', justify: 'เสมอหน้าหลัง' };
    const UNIT_RE = /[0-9๐-๙][0-9๐-๙,.]* (?:บาท|สตางค์|คน|ราย|วัน|เดือน|ปี|ชั่วโมง|นาที|แห่ง|ฉบับ|หน้า|ครั้ง|ชุด|เครื่อง|หลัง|กิโลเมตร|เมตร|ล้าน|เปอร์เซ็นต์|ข้อ|อัตรา|โครงการ)/g;
    function guessRole(p) {
      const t = p.text.replace(/\u00A0/g, ' ').trim();
      if (p.kind === 'table') return 'table';
      if (!t) return 'blank';
      if (p.kind === 'list') return 'list';
      if (/^บันทึกข้อความ$/.test(t) || (p.formats.align === 'center' && p.bold && t.length < 40 && !p.formats.marginleft)) return 'title';
      // ส่วนหัวเป็นบรรทัดสั้น (ที่ใหญ่สุดคือบรรทัดส่วนราชการ/เรื่อง) — ถ้ายาวกว่านี้ถือเป็นเนื้อความที่ขึ้นต้นด้วยคำเดียวกัน
      if (t.length <= 90 && /^(ส่วนราชการ|ที่|วันที่|เรื่อง|เรียน|กราบเรียน|อ้างถึง|สิ่งที่ส่งมาด้วย)(\s|$)/.test(t)) return 'head';
      if (/^(โทร\.|โทรสาร|ไปรษณีย์อิเล็กทรอนิกส์)/.test(t)) return 'foot';
      if (/^ขอแสดงความนับถือ/.test(t) || cmOf(p.formats.marginleft) >= 4 || /^\(.*\)$/.test(t)) return 'sign';
      return t.length >= 40 ? 'body' : 'other';
    }
    // ค่าที่ใช้จริงของย่อหน้า (ถ้าไม่ได้ตั้ง = ค่าเริ่มต้นของมาตรฐานเอกสาร; รายการ/ตารางไม่มีย่อหน้าบรรทัดแรกเริ่มต้น)
    function effective(p, k, D) {
      if (p.formats[k] != null) return p.formats[k];
      if (k === 'align') return p.kind === 'block' ? D.align : 'left';
      if (k === 'textindent') return p.kind === 'block' ? D.indent : '0cm';
      if (k === 'lineheight') return D.lineHeight;
      if (k === 'marginleft') return '0cm';
      if (k === 'letterspacing') return p.ls;              // ตัวเลข em หรือ null = หลายค่าปนกัน
      return null;
    }
    const sameVal = (k, a, b) => (['textindent', 'marginleft', 'spacebefore', 'spaceafter'].includes(k) ? Math.abs(cmOf(a) - cmOf(b)) < 0.01 : String(a ?? '') === String(b ?? ''));
    // คืนเฉพาะค่าที่ต่างจากปัจจุบัน; ถ้าค่าเป้าหมายเท่ากับค่าเริ่มต้น จะลบค่าที่ตั้งไว้แทน (false) เพื่อให้ HTML สะอาด
    function diffChanges(p, target, D) {
      const out = {};
      for (const [k, v] of Object.entries(target)) {
        if (k === 'textalignlast') { if ((p.formats.textalignlast || null) !== (v || null)) out[k] = v || false; continue; }
        if (k === 'letterspacing') { if (p.ls == null || Math.abs(p.ls - v) >= 0.0005) out[k] = v; continue; }
        if (sameVal(k, effective(p, k, D), v)) continue;
        const def = effective({ ...p, formats: {} }, k, D);
        out[k] = def != null && sameVal(k, def, v) ? false : v;
      }
      return out;
    }
    function localPlan(paras) {
      const D = ed.settings().defaults;
      // ระยะกั้นซ้ายของส่วนลงชื่อ: ใช้ค่าที่พบมากที่สุด ให้ทุกบรรทัดลงชื่อตรงแนวเดียวกัน
      const signLefts = new Map();
      for (const p of paras) if (guessRole(p) === 'sign' && p.formats.marginleft) signLefts.set(p.formats.marginleft, (signLefts.get(p.formats.marginleft) || 0) + 1);
      const signLeft = [...signLefts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
      const items = [];
      for (const p of paras) {
        const role = guessRole(p);
        let target = {};
        const notes = [];
        if (role === 'body') {
          target = { textindent: D.indent, marginleft: '0cm', lineheight: D.lineHeight };
          if (!['justify'].includes(effective(p, 'align', D))) target.align = 'justify';
        } else if (role === 'head' || role === 'foot') target = { align: 'left', textindent: '0cm', lineheight: D.lineHeight };
        else if (role === 'title' || role === 'list' || role === 'blank' || role === 'other') target = { lineheight: D.lineHeight };
        else if (role === 'sign') target = { lineheight: D.lineHeight, ...(signLeft ? { marginleft: signLeft } : {}) };
        else continue;
        const m = p.metrics || {};
        // 1) ตัดคำ: จุดตัดบรรทัดที่ผิดหลักซึ่งผูกได้ + ตัวเลขกับหน่วยทุกจุด
        const bind = new Set(role === 'body' || role === 'list' ? (p.text.match(UNIT_RE) || []) : []);
        const join = new Set();
        for (const b of m.badBreaks || []) {
          if (b.join) { join.add(b.join); notes.push(`ตัดคำ: ${b.reason} (“${b.at.replace('|', '⏎')}”) — ผูกคำด้วย Word Joiner`); }
          else if (b.phrase) { bind.add(b.phrase); notes.push(`ตัดคำ: ${b.reason} (“${b.at.replace('|', '⏎')}”)`); }
          else notes.push(`ตัดคำ: ${b.reason} (“${b.at.replace('|', '⏎')}”) — ไม่มีเว้นวรรคให้ผูก ลองปรับช่องไฟหรือเรียบเรียงใหม่`);
        }
        // ชื่อคนหลังคำนำหน้าที่ Segmenter แยกเป็นหลายชิ้น (ไม่อยู่ในพจนานุกรม) เสี่ยงถูกตัดกลางชื่อ → เสนอเป็นคำห้ามตัด
        for (const name of nameCandidates(p.text)) { join.add(name); notes.push(`ตัดคำ: ชื่อ “${name}” อาจถูกตัดกลางชื่อ (${ed.segmentWords(name).join('|')}) — เพิ่มเป็นคำห้ามตัด`); }
        // 2) กระจายตัวอักษร: คำโดดบรรทัดสุดท้าย (ใช้ค่าที่ทดลองกับสำเนาแล้วว่าได้ผล) และช่องว่างที่ถูกยืดมาก
        if (m.orphanFix != null && role === 'body') {
          target.letterspacing = m.orphanFix;
          notes.push(`กระจายตัวอักษร: บรรทัดสุดท้ายมีคำโดด (ยาว ${Math.round(m.lastFill * 100)}%) — ช่องไฟ ${signedEm(m.orphanFix)} ดึงขึ้นไปรวมบรรทัดบนได้ (ทดลองแล้ว)`);
        } else if (m.lastFill != null && m.lastFill < 0.2 && m.lines > 1 && role === 'body') {
          notes.push(`กระจายตัวอักษร: บรรทัดสุดท้ายมีคำโดด (ยาว ${Math.round(m.lastFill * 100)}%) แต่ช่องไฟถึง −0.03 em ยังดึงขึ้นไม่ได้ — พิจารณาเรียบเรียงใหม่`);
        }
        if (m.stretch > 0.08 && m.lines > 1) {
          if (m.stretchFix != null && role === 'body' && target.letterspacing == null) {
            target.letterspacing = m.stretchFix;
            notes.push(`กระจายตัวอักษร: บรรทัด ${m.stretchLine} ตัวอักษรถูกยืดห่าง +${m.stretch.toFixed(3)} em/ตัว — ช่องไฟ ${signedEm(m.stretchFix)} ลดการยืดได้โดยไม่เพิ่มบรรทัด (ทดลองแล้ว)`);
          } else notes.push(`กระจายตัวอักษร: บรรทัด ${m.stretchLine} ตัวอักษรถูกยืดห่าง +${m.stretch.toFixed(3)} em/ตัว — มักเกิดจากคำยาวที่ขึ้นบรรทัดใหม่ ลองเรียบเรียงใหม่หรือเลิกผูกคำนั้น`);
        }
        if (/ {2,}/.test(p.text) && role === 'body') notes.push('มีเว้นวรรคซ้อน (ตรวจได้ในแท็บตรวจคำ)');
        const changes = diffChanges(p, target, D);
        // 3) ระยะบรรทัด (อยู่ใน target ของทุกบทบาทแล้ว)
        if ('lineheight' in changes) notes.push(`ระยะบรรทัด: ให้${ROLE_TH[role]}ใช้ ${D.lineHeight} เท่ากันทั้งฉบับ`);
        if (!Object.keys(changes).length && !bind.size && !join.size) continue;
        if (!notes.length) notes.push(`ให้${ROLE_TH[role]}ใช้ค่าเดียวกันทั้งฉบับตามมาตรฐานที่เลือก`);
        items.push({ page: p.page, n: p.n, role, text: p.text, changes, bind: [...bind], join: [...join], reason: notes.join(' · '), p });
      }
      return items.sort((a, b) => priority(b) - priority(a) || a.page - b.page || a.n - b.n);
    }
    const ALLOWED_AI_ROLES = new Set(Object.keys(ROLE_TH));
    function fromAi(data, paras) {
      const D = ed.settings().defaults;
      const byKey = new Map(paras.map(p => [`${p.page}:${p.n}`, p]));
      const items = [];
      for (const it of data.items || []) {
        const p = byKey.get(`${it.page}:${it.n}`);
        if (!p) continue;
        const target = {};
        if (it.align) Object.assign(target, it.align === 'distribute' ? { align: 'justify', textalignlast: 'justify' } : { align: it.align, textalignlast: false });
        for (const k of ['textindent', 'marginleft', 'spacebefore', 'spaceafter']) if (typeof it[k] === 'string' && LEN_RE.test(it[k].trim())) target[k] = it[k].trim();
        if (typeof it.lineheight === 'string' && LH_RE.test(it.lineheight.trim())) target.lineheight = it.lineheight.trim();
        const lsm = /^(-?\d*\.?\d+)\s*(em)?$/.exec(String(it.letterspacing ?? '').trim());
        if (lsm && Math.abs(+lsm[1]) <= 0.1) target.letterspacing = +lsm[1];        // ช่องไฟจาก AI จำกัดไว้ไม่เกิน ±0.1 em
        const changes = diffChanges(p, target, D);
        const text = p.text.replace(/\u00A0/g, ' ');
        const bind = (it.bind || []).filter(b => typeof b === 'string' && b.includes(' ') && text.includes(b));
        const join = (it.nobreak || []).filter(w => typeof w === 'string' && w.length > 1 && !/\s/.test(w) && text.includes(w) && ed.segmentWords(w).length > 1);
        if (!Object.keys(changes).length && !bind.length && !join.length) continue;
        const roleKey = Object.keys(ROLE_TH).find(k => ROLE_TH[k] === it.role) || (ALLOWED_AI_ROLES.has(it.role) ? it.role : null);
        items.push({ page: p.page, n: p.n, role: roleKey || 'other', roleText: roleKey ? null : it.role, text: p.text, changes, bind, join, reason: it.reason || '', p });
      }
      return items.sort((a, b) => priority(b) - priority(a) || a.page - b.page || a.n - b.n);
    }
    let lyItems = [];
    const fmtVal = (k, v) => (k === 'letterspacing' ? (v == null ? 'หลายค่า' : signedEm(+v)) : v === false || v == null ? 'ค่าเริ่มต้น' : k === 'align' ? ALIGN_TXT[v] || v : k === 'textalignlast' ? 'กระจาย' : v);
    function describeChanges(it) {
      const D = ed.settings().defaults;
      const parts = Object.entries(it.changes).filter(([k]) => k !== 'textalignlast' || it.changes.textalignlast).map(([k, v]) => {
        const before = effective(it.p, k, D);
        return `${FMT_TH[k] || k} ${fmtVal(k, before)} → ${fmtVal(k, v === false ? effective({ ...it.p, formats: {} }, k, D) : v)}`;
      });
      if (it.bind.length) parts.push(`ห้ามตัดบรรทัด: ${it.bind.map(b => `“${b}”`).join(', ')}`);
      if ((it.join || []).length) parts.push(`ห้ามตัดกลางคำ: ${it.join.map(w => `“${w}”`).join(', ')}`);
      return parts.join(' · ');
    }
    function renderLayout() {
      const left = lyItems.filter(x => !x.status).length;
      $('#lyApplyAll').disabled = !left;
      $('#lyList').innerHTML = lyItems.map((it, i) => `<li${it.status ? ' style="opacity:.5"' : ''}>
        <span class="k">หน้า ${it.page + 1} · ย่อหน้า ${it.n + 1} · ${esc(it.roleText || ROLE_TH[it.role])}</span>${catsOf(it).map(c => `<span class="k" style="color:${c.color};background:${c.bg}">${c.label}</span>`).join('')}
        <span class="n">“${esc(it.text.replace(/\u00A0/g, ' ').trim().slice(0, 50)) || '(บรรทัดว่าง)'}${it.text.length > 50 ? '…' : ''}”</span>
        <span class="s" style="white-space:normal;display:block;margin-top:3px">${esc(describeChanges(it))}</span>
        <span class="n">${esc(it.reason)}</span>
        <span class="acts"><button type="button" data-ly="${i}" data-do="show">ดู</button>${it.status ? `<span class="n" style="margin:0">${it.status === 'done' ? 'ใช้แล้ว' : 'ข้ามแล้ว'}</span>` : `<button type="button" data-ly="${i}" data-do="apply">ใช้</button><button type="button" data-ly="${i}" data-do="skip">ข้าม</button>`}</span></li>`).join('');
    }
    function applyLy(it) {
      if (it.status) return;
      const { letterspacing, ...block } = it.changes;
      // คำที่ผูกจะถูกเพิ่มเข้าพจนานุกรมคำห้ามตัดด้วย เพื่อใช้กับย่อหน้าอื่นและการตรวจครั้งต่อไป
      if ((it.join || []).length) addNoBreak(it.join, true);
      it.status = ed.applyParagraph(it.page, it.n, block, it.bind, letterspacing != null ? { letterspacing } : {}, it.join || []) ? 'done' : 'skip';
    }
    $('#lyList').addEventListener('click', e => {
      const b = e.target.closest('button[data-ly]');
      if (!b) return;
      const it = lyItems[+b.dataset.ly];
      if (b.dataset.do === 'show') { ed.gotoLine(it.page, it.n); return; }
      if (b.dataset.do === 'apply') { applyLy(it); say(`ปรับหน้า ${it.page + 1} ย่อหน้า ${it.n + 1} แล้ว`); }
      else it.status = 'skip';
      renderLayout();
    });
    function showLayoutResult(items, plan) {
      lyItems = items;
      $('#lyPlan').innerHTML = plan ? `<div class="out">${esc(plan)}</div>` : '';
      $('#lyStatus').textContent = items.length ? `เสนอการปรับ ${items.length} ย่อหน้า — ตรวจทีละข้อแล้วกด “ใช้” หรือ “ข้าม”` : 'ทุกย่อหน้าสอดคล้องกันแล้ว ไม่มีข้อเสนอ';
      renderLayout();
    }
    // ข้อมูลที่ส่งให้ AI วางแผนจัดหน้า (ใช้ทั้งแท็บจัดหน้าและตัวจัดอัตโนมัติ)
    function layoutPayload(paras) {
      const D = ed.settings().defaults;
      return {
        nobreakDictionary: noBreak.slice(0, 200),
        standard: { font: D.font, size: D.size, align: D.align, indent: D.indent, lineHeight: D.lineHeight, page: 'A4 ขอบ บน 2.5 ซ้าย 3 ขวา 2 ล่าง 2 ซม. พื้นที่ข้อความกว้าง 16 ซม.' },
        paragraphs: paras.map(p => ({ page: p.page, n: p.n, role: ROLE_TH[guessRole(p)], text: p.text.replace(/\u00A0/g, ' ').slice(0, 160), formats: p.formats, ls: p.ls, ...(p.metrics ? { metrics: p.metrics } : {}) })),
      };
    }
    // ตรวจความพร้อมของ AI ก่อนส่ง (ใช้ค่าจากแท็บผู้ช่วย AI)
    function aiPrecheck() {
      if (!keys[provider] && !PROVIDERS[provider].keyOptional) return 'กรอก API key ในแท็บผู้ช่วย AI ก่อน';
      if (!staysLocal() && !aiConsent.checked) return 'ติ๊กยืนยันในแท็บผู้ช่วย AI ก่อนว่าข้อความส่งออกนอกเครื่องได้';
      return null;
    }
    // ผู้ช่วย AI วิเคราะห์ช่องไฟ: เรียกจากแผงช่องไฟตัวอักษร (scope = sel | para | page, r = ช่วงที่เลือก)
    ed.aiSpacing = async (scope, r) => {
      const err = aiPrecheck();
      if (err) return { error: err };
      const paras = ed.paragraphs('page').filter(p => p.text.trim());
      if (!paras.length) return { error: 'ยังไม่มีข้อความในหน้านี้' };
      if (paras.length > 150) return { error: `มี ${paras.length} ย่อหน้า มากเกินไปสำหรับ AI ครั้งเดียว` };
      const all = quill.getLines();
      let focus = [];
      if (scope !== 'page' && r) {
        const ls = r.length ? quill.getLines(r.index, r.length) : [quill.getLine(r.index)[0]].filter(Boolean);
        focus = ls.map(l => all.indexOf(l)).filter(n => n >= 0);
      }
      const payload = layoutPayload(paras);
      payload.focus = focus;
      delete payload.nobreakDictionary;
      try {
        const data = await askAiCached('spacing', JSON.stringify(payload));
        const okN = new Set(paras.map(p => p.n));
        const items = (Array.isArray(data.items) ? data.items : []).map(it => {
          const v = parseFloat(String(it.letterspacing).replace('\u2212', '-'));
          return { n: it.n, ls: Number.isFinite(v) ? Math.max(-0.03, Math.min(0.05, Math.round(v * 1000) / 1000)) : null, reason: String(it.reason || '') };
        }).filter(it => it.ls != null && okN.has(it.n) && (!focus.length || focus.includes(it.n)));
        return { summary: String(data.summary || ''), items, page: ed.currentPage(), model: PROVIDERS[provider].label };
      } catch (e) { return { error: aiErrorText(e) }; }
    };
    Object.assign(actions, {
      lyLocal() {
        const paras = ed.paragraphs($('#lyScope').value);
        $('#lyPlan').innerHTML = '';
        showLayoutResult(localPlan(paras), '');
      },
      async lyAi() {
        const err = aiPrecheck();
        if (err) { $('#lyStatus').textContent = err; return; }
        const paras = ed.paragraphs($('#lyScope').value);
        if (paras.length > 400) { $('#lyStatus').textContent = `มี ${paras.length} ย่อหน้า มากเกินไปสำหรับ AI ครั้งเดียว — เลือกขอบเขต “หน้านี้”`; return; }
        const payload = layoutPayload(paras);
        const btn = $('#lyAiBtn');
        btn.disabled = true;
        $('#lyStatus').textContent = `กำลังให้ ${PROVIDERS[provider].label} วางแผนจัดหน้า ${paras.length} ย่อหน้า… อาจใช้เวลาสักครู่`;
        try {
          const data = await askAi('layout', JSON.stringify(payload));
          showLayoutResult(fromAi(data, paras), data.plan || '');
        } catch (e) {
          $('#lyStatus').textContent = 'ไม่สำเร็จ: ' + aiErrorText(e);
        } finally { btn.disabled = false; }
      },
      lyApplyAll() {
        let n = 0;
        for (const it of lyItems) if (!it.status) { applyLy(it); if (it.status === 'done') n++; }
        renderLayout();
        say(`ปรับแล้ว ${n} ย่อหน้า (หน้าที่เปิดอยู่กดย้อนกลับได้)`);
      },
    });

    // ================= ทดสอบการวางใน e-Office =================
    // ผลทดสอบเก็บในเบราว์เซอร์นี้ และใช้ปรับการส่งออก (ถอด Word Joiner / ใช้ช่องว่างแทนแท็บ) กับเตือนตอนคัดลอก
    const PT = window.eofficePasteTest;
    const PASTE_KEY = 'eoffice-editor:pastetest';
    let pasteSaved = (() => { try { return JSON.parse(store.get(PASTE_KEY) || 'null'); } catch { return null; } })();
    const ICON = { ok: ['✓', 'pt-ok', 'รอด'], changed: ['≈', 'pt-changed', 'เปลี่ยน'], lost: ['✗', 'pt-lost', 'หาย'], missing: ['?', 'pt-missing', 'ไม่พบ'] };
    function applyPasteProfile() {
      const results = pasteSaved?.results;
      ed.setExportProfile(results ? PT.profileOf(results) : {});
      ed.setCopyAdvisor(results ? html => PT.adviceFor(results, html) : null);
    }
    function renderPasteResults() {
      const sum = $('#pasteSummary'), list = $('#pasteList');
      if (!pasteSaved) { sum.innerHTML = ''; list.innerHTML = ''; return; }
      const rs = pasteSaved.results, count = s => rs.filter(r => r.state === s).length;
      sum.innerHTML = `<div class="hint">ผลทดสอบเมื่อ ${esc(new Date(pasteSaved.at).toLocaleString('th-TH'))} (เก็บในเบราว์เซอร์นี้)</div>
        <div class="pt-sum">${['ok', 'changed', 'lost', 'missing'].filter(s => count(s)).map(s => `<span class="pt-pill ${ICON[s][1]}">${ICON[s][0]} ${ICON[s][2]} ${count(s)}</span>`).join('')}<span class="pt-pill pt-missing">จาก ${rs.length} ข้อ</span></div>`;
      const profile = PT.profileOf(rs), notes = [];
      if (profile.stripWJ) notes.push('การส่งออกจะถอด Word Joiner ออก');
      if (profile.tabsAsSpaces) notes.push('การส่งออกจะใช้ช่องว่างห้ามตัด 8 ตัวแทนแท็บ');
      if (notes.length) sum.innerHTML += `<div class="hint">ปรับการส่งออกอัตโนมัติ: ${notes.join(' · ')}</div>`;
      // ข้อที่ไม่ผ่านขึ้นก่อน เรียงตามความรุนแรง
      const order = { lost: 0, changed: 1, missing: 2, ok: 3 };
      list.innerHTML = [...rs].sort((a, b) => order[a.state] - order[b.state]).map(r => {
        const [ic, cls] = ICON[r.state];
        return `<li><div class="pt-row"><span class="pt-icon ${cls}" style="background:none">${ic}</span><span><b>${esc(r.id)}</b> ${esc(r.label)} <span class="n" style="display:inline">— ${esc(r.detail || '')}</span></span>${r.state !== 'ok' ? `<span class="n">ผลกระทบ: ${esc(r.impact || '')}</span>` : ''}</div></li>`;
      }).join('');
    }
    function showPasteAnalysis(html) {
      const a = PT.analyze(html);
      if (a.found === 0) { $('#pasteStatus').textContent = 'ไม่พบรหัสทดสอบ [Txx] ในสิ่งที่วาง — ตรวจว่าคัดลอกจากหน้า e-Office ที่วางชุดทดสอบไว้'; return; }
      pasteSaved = { at: Date.now(), results: a.results };
      store.set(PASTE_KEY, JSON.stringify(pasteSaved));
      applyPasteProfile();
      renderPasteResults();
      $('#pasteStatus').textContent = a.found < a.total ? `พบรหัสทดสอบ ${a.found} จาก ${a.total} ข้อ — ข้อที่ไม่พบอาจวางไม่ครบ` : `วิเคราะห์ครบ ${a.total} ข้อแล้ว`;
      say('บันทึกผลทดสอบการวางแล้ว — การส่งออกจะปรับตามผลนี้');
    }
    $('#pasteBox').addEventListener('paste', e => {
      e.preventDefault();
      const html = e.clipboardData.getData('text/html'), text = e.clipboardData.getData('text/plain');
      if (!html) { $('#pasteStatus').textContent = 'คลิปบอร์ดมีแต่ข้อความธรรมดา (ไม่มี HTML) จึงวิเคราะห์รูปแบบไม่ได้ — คัดลอกจากหน้า e-Office ที่แสดงผลเอกสาร ไม่ใช่จากช่องแก้ไขข้อความดิบ'; return; }
      showPasteAnalysis(html);
      $('#pasteBox').textContent = `วางแล้ว (${html.length.toLocaleString()} ตัวอักษร HTML, ${text.length.toLocaleString()} ตัวอักษรข้อความ)`;
    });
    Object.assign(actions, {
      async pasteCopy() {
        const ok = await ed.writeClipboard(PT.buildHtml(), PT.buildText());
        $('#pasteStatus').textContent = ok ? 'คัดลอกชุดทดสอบแล้ว — วางใน e-Office (ขั้นที่ 2)' : 'คัดลอกอัตโนมัติไม่ได้ — เบราว์เซอร์ไม่อนุญาต';
        say($('#pasteStatus').textContent);
      },
      async pasteRead() {
        try {
          const items = await navigator.clipboard.read();
          for (const it of items) {
            if (it.types.includes('text/html')) { showPasteAnalysis(await (await it.getType('text/html')).text()); return; }
          }
          $('#pasteStatus').textContent = 'คลิปบอร์ดไม่มี HTML — คัดลอกจาก e-Office แล้วลองใหม่ หรือวางในกล่องด้านบน';
        } catch (err) {
          $('#pasteStatus').textContent = 'อ่านคลิปบอร์ดไม่ได้ (' + (err.name === 'NotAllowedError' ? 'ยังไม่ได้อนุญาต' : err.message) + ') — วางในกล่องด้านบนด้วย Ctrl+V แทน';
        }
      },
      pastePreview() {
        let html = ed.exportHtml();
        if ($('#pasteSim').checked) {
          if (!pasteSaved) { $('#pasteStatus').textContent = 'ยังไม่มีผลทดสอบให้จำลอง — ทำขั้นตอนทดสอบก่อน'; return; }
          html = PT.simulate(html, pasteSaved.results);
        }
        const f = $('#pasteFrame');
        f.hidden = false;
        f.srcdoc = `<!doctype html><meta charset="utf-8"><body style="margin:12px;font-family:Tahoma,sans-serif">${html}</body>`;
      },
      pasteClear() {
        if (!confirm('ลบผลทดสอบการวาง? การส่งออกจะกลับไปใช้ค่าปกติ')) return;
        pasteSaved = null;
        store.del(PASTE_KEY);
        applyPasteProfile();
        renderPasteResults();
        $('#pasteStatus').textContent = 'ลบผลทดสอบแล้ว';
      },
    });
    applyPasteProfile();
    renderPasteResults();

    // ================= จัดรูปแบบอัตโนมัติ (ทีเดียวจบ) =================
    // รันขั้นตอนที่เลือกต่อกัน: ถ่ายภาพทั้งเอกสารไว้ก่อน → ทำทีละขั้นพร้อมนับจุดที่เปลี่ยน → (ตัวอย่าง) คืนภาพเดิม / (จัดจริง) เก็บภาพไว้ให้ย้อนกลับทุกหน้า
    const AU_KEY = 'eoffice-editor:auto';
    const AU_DEFAULT = { scope: 'page', digits: 'keep', clean: true, bigspace: true, blank: true, misspell: false, dates: false, layout: true, bind: true, spread: true, pagenum: true, aiLayout: false, aiCheck: false, aiRewrite: false, aiAudit: false };
    const auOpts = (() => { try { return { ...AU_DEFAULT, ...JSON.parse(store.get(AU_KEY) || '{}') }; } catch { return { ...AU_DEFAULT }; } })();
    const CLEAN_KINDS = new Set(['ไม้ยมก', 'ฯลฯ', 'สระแอ', 'สระอำ', 'วรรณยุกต์ซ้ำ', 'สระซ้ำ', 'เว้นวรรคซ้อน', 'ช่องว่างท้ายบรรทัด', 'ช่องว่างหน้าเครื่องหมาย']);
    let auUndo = null, auFrames = null;
    const auVis = s => String(s).replace(/ /g, '␣').replace(/ /g, '⍽').replace(/⁠/g, '⁀').replace(/\t/g, '→').replace(/[​‌‍﻿]/g, '⌀').replace(/\n/g, '⏎');
    const isBlankLine = l => l.statics.blotName === 'block' && l.length() === 1;

    // ----- ขั้นตอน AI: เรียกผ่านผู้ให้บริการ/คีย์/ยินยอมที่ตั้งไว้ในแท็บผู้ช่วย AI -----
    const AI_FUNCS = ['aiLayout', 'aiCheck', 'aiRewrite', 'aiAudit'];
    const aiCache = new Map();         // เก็บคำตอบของ AI ตามข้อมูลที่ส่ง: ดูตัวอย่างแล้วจัดจริงทันทีจะไม่เรียก AI (และไม่เสียค่าใช้จ่าย) ซ้ำ
    const hashStr = s => { let h = 5381; for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0; return h; };
    async function askAiCached(taskKey, payloadText) {
      const key = `${provider}|${modelOf(provider)}|${taskKey}|${payloadText.length}|${hashStr(payloadText)}`;
      if (aiCache.has(key)) return aiCache.get(key);
      const data = await askAi(taskKey, payloadText);
      aiCache.set(key, data);
      if (aiCache.size > 60) aiCache.delete(aiCache.keys().next().value);
      return data;
    }
    const digitsOf = s => toArabic(String(s)).replace(/[^0-9]/g, '').split('').sort().join('');
    const pageTextOf = p => (p === ed.currentPage() ? quill.getText() : docText(ed.pages()[p].delta).replaceAll(OBJ, ''));
    // แก้ข้อความตามรายการ {wrong, suggest} ของ AI: แทนที่เฉพาะตำแหน่งแรกที่พบ และข้ามรายการที่ผิดปกติ
    function applyTextEdits(q, items, samples, page) {
      let n = 0;
      for (const it of items) {
        const wrong = String(it.wrong ?? ''), suggest = String(it.suggest ?? '');
        if (!wrong || wrong === suggest || wrong.length > 120 || suggest.length > 200 || /\n/.test(wrong + suggest)) continue;
        const delta = q.getContents(), i = docText(delta).indexOf(wrong);
        if (i < 0 || wrong.includes(OBJ)) continue;
        q.updateContents(new Delta().retain(i).delete(wrong.length).insert(suggest, attrsAt(delta, i)), 'user');
        n++;
        if (samples.length < 12) samples.push(`หน้า ${page + 1}: “${auVis(wrong)}” → “${auVis(suggest)}”${it.reason ? ` (${it.reason})` : ''}`);
      }
      return n;
    }
    // เขียนย่อหน้าใหม่ตามที่ AI เรียบเรียง พร้อมเกราะป้องกัน: ความยาวไม่เปลี่ยนมาก ตัวเลขทุกตัวต้องตรง ข้อความต้องไม่ถูกแก้ระหว่างทาง
    function applyRewrites(q, items, chunk, samples, page) {
      let n = 0;
      for (const it of items) {
        const orig = chunk.find(c => c.n === it.n);
        if (!orig) continue;
        const text = String(it.text ?? '').replace(/\s*\n+\s*/g, ' ').trim();
        if (!text || text === orig.text) continue;
        const skip = why => { if (samples.length < 12) samples.push(`ข้ามย่อหน้า ${it.n + 1} (หน้า ${page + 1}): ${why}`); };
        const ratio = text.length / orig.text.length;
        if (ratio < 0.5 || ratio > 1.8) { skip('ความยาวเปลี่ยนมากเกินไป'); continue; }
        if (digitsOf(text) !== digitsOf(orig.text)) { skip('ตัวเลขไม่ตรงกับต้นฉบับ'); continue; }
        const line = q.getLines()[it.n];
        if (!line) continue;
        const at = q.getIndex(line), len = line.length() - 1;
        const ops = q.getContents(at, len).ops;
        if (!ops.every(op => typeof op.insert === 'string')) { skip('มีรูปหรือวัตถุฝังอยู่'); continue; }
        if (ops.map(op => op.insert).join('').replace(/\u00A0/g, ' ').replace(/\u2060/g, '') !== orig.text) { skip('ข้อความถูกแก้ไประหว่างทาง'); continue; }
        q.updateContents(new Delta().retain(at).delete(len).insert(text, ops[0]?.attributes), 'user');
        n++;
        if (samples.length < 12) samples.push(`หน้า ${page + 1} ย่อหน้า ${it.n + 1}: “${orig.text.slice(0, 40)}…” → “${text.slice(0, 40)}…”${it.reason ? ` (${it.reason})` : ''}`);
      }
      return n;
    }
    let auBusy = false;

    async function auRun(opts, { restore }) {
      const scope = opts.scope, log = [];
      const snap = ed.snapshotAll();
      const s0 = ed.settings();
      const settings0 = { pageNumbers: s0.pageNumbers, pageDigits: s0.pageDigits, listDigits: s0.listDigits };
      const undoInfo = { snap, settings0 };
      const before = ed.exportHtml();
      const prof = pasteSaved?.results ? PT.profileOf(pasteSaved.results) : {};
      const progress = msg => { $('#auStatus').textContent = msg; $('#auShieldMsg').textContent = msg; if (auCancel) throw new Error('ยกเลิกโดยผู้ใช้'); };
      const chk = () => { if (auCancel) throw new Error('ยกเลิกโดยผู้ใช้'); };
      const step = (key, label, fn) => {
        const samples = [];
        const count = fn(samples) || 0;
        log.push({ key, label, count, samples: samples.slice(0, 3) });
      };
      // แทนที่ตาม regex ทั้งขอบเขต พร้อมเก็บตัวอย่างก่อน→หลัง
      const rep = (re, f, samples) => replaceInDoc(scope, re, (...m) => {
        const r = f(...m);
        if (r !== m[0] && samples.length < 12) samples.push(`“${auVis(m[0])}” → “${auVis(r)}”`);
        return r;
      });
      const pageList = () => (scope === 'all' ? ed.pages().map((_, i) => i) : [ed.currentPage()]);
      // ขั้นตอน AI: ล้มเหลวแล้วข้ามขั้นนั้นพร้อมแจ้งเหตุผล ขั้นอื่นทำต่อ
      const aiStep = async (key, label, fn, { max = 3, report = false } = {}) => {
        const samples = [];
        let count = 0;
        try { count = (await fn(samples)) || 0; chk(); } catch (err) { if (auCancel) throw err; samples.unshift('ไม่สำเร็จ: ' + aiErrorText(err)); }
        log.push({ key, label, count, report, samples: samples.slice(0, max) });
      };

      try {
        if (opts.blank) step('blank', 'ลบบรรทัดว่างซ้อนกัน', samples => {
          let n = 0;
          for (const p of pageList()) ed.withPage(p, q => {
            const lines = q.getLines();
            let i = lines.length - 1;
            while (i > 0) {
              if (!isBlankLine(lines[i])) { i--; continue; }
              let j = i - 1;
              while (j >= 0 && isBlankLine(lines[j])) { q.deleteText(q.getIndex(lines[j]), 1, 'user'); n++; j--; }
              if (i - j > 1 && samples.length < 3) samples.push(`หน้า ${p + 1}: ลบ ${i - j - 1} บรรทัดว่างที่ซ้อนกัน`);
              i = j;
            }
          });
          return n;
        });
        if (opts.clean) step('clean', 'ล้างข้อความ', samples => {
          let n = rep(/[\u200B\u200C\u200D\uFEFF]/g, () => '', samples);
          for (const [kind, re, fix] of RULES) {
            if (!CLEAN_KINDS.has(kind)) continue;
            n += kind === 'เว้นวรรคซ้อน' ? rep(/ {2,}/g, m => (opts.bigspace ? (m.length > 2 ? '  ' : m) : ' '), samples) : rep(re, fix, samples);
          }
          return n;
        });
        if (opts.misspell) step('misspell', 'แก้คำสะกดที่ผิดบ่อย', samples => rep(MISSPELL_RE, m => MISSPELL[m] || m, samples));
        if (opts.dates) step('dates', 'วันที่แบบราชการ', samples => DATE_RULES.reduce((n, [re, fix]) => n + rep(re, (...m) => fix(...m) || m[0], samples), 0));
        if (opts.digits === 'thai') step('digits', 'เปลี่ยนเป็นเลขไทย', samples => rep(/[0-9]+/g, toThai, samples));
        if (opts.digits === 'arabic') step('digits', 'เปลี่ยนเป็นเลขอารบิก', samples => rep(/[๐-๙]+/g, toArabic, samples));

        // ---- AI ที่แก้ข้อความ (ทำก่อนจัดหน้า เพราะการวัดบรรทัดต้องวัดหลังข้อความนิ่งแล้ว) ----
        let aiOk = false;
        if (AI_FUNCS.some(k => opts[k])) {
          const pre = aiPrecheck();
          if (pre) log.push({ key: 'ai', label: 'AI (ข้ามทั้งหมด)', count: 0, samples: [pre] });
          else aiOk = true;
        }
        if (aiOk && opts.aiCheck) await aiStep('aiCheck', 'AI ตรวจคำผิด/ไวยากรณ์/การใช้คำ', async samples => {
          let n = 0;
          for (const p of pageList()) {
            progress(`AI ตรวจคำผิด หน้า ${p + 1}…`);
            const text = pageTextOf(p).replace(/\n+$/, '');
            if (text.trim().length < 3) continue;
            if (text.length > AI_MAX_CHARS) { samples.push(`หน้า ${p + 1} ยาวเกิน ${AI_MAX_CHARS.toLocaleString()} ตัวอักษร จึงข้าม`); continue; }
            const data = await askAiCached('check', text);
            n += ed.withPage(p, q => applyTextEdits(q, data.items || [], samples, p));
          }
          return n;
        });
        if (aiOk && opts.aiRewrite) await aiStep('aiRewrite', 'AI เรียบเรียงภาษาราชการ', async samples => {
          let n = 0;
          const all = ed.paragraphs(scope);
          for (const p of pageList()) {
            // เฉพาะย่อหน้าเนื้อความ/รายการ ที่ยาวพอ และมีรูปแบบอักษรชุดเดียว (ไม่เสียตัวหนา/เอียงที่ตั้งใจใส่)
            const cand = all.filter(x => x.page === p && ['body', 'list'].includes(guessRole(x)) && x.variety <= 1 && x.text.trim().length >= 30 && x.text.length <= 1500)
              .map(x => ({ n: x.n, text: x.text.replace(/\u00A0/g, ' ').replace(/\u2060/g, '') }));
            const chunks = [];
            for (const c of cand) {
              const last = chunks[chunks.length - 1];
              if (last && last.chars + c.text.length <= 12000) { last.items.push(c); last.chars += c.text.length; } else chunks.push({ items: [c], chars: c.text.length });
            }
            for (const [i, ch] of chunks.entries()) {
              progress(`AI เรียบเรียงภาษาราชการ หน้า ${p + 1}${chunks.length > 1 ? ` (ชุด ${i + 1}/${chunks.length})` : ''}…`);
              const data = await askAiCached('rewrite', JSON.stringify({ paragraphs: ch.items }));
              n += ed.withPage(p, q => applyRewrites(q, data.items || [], ch.items, samples, p));
            }
          }
          return n;
        });

        // ---- จัดหน้า: แบบในเครื่อง (วัดจากหน้าจอ) หรือให้ AI วางแผน ----
        // ถ้าเลือก AI วางแผนจัดหน้า ขั้น “จัดย่อหน้า/กระจายตัวอักษร” ในเครื่องจะถูกแทน เหลือเฉพาะการผูกตัวเลขกับหน่วย
        const L = opts.aiLayout ? { layout: false, spread: false, bind: opts.bind } : opts;
        if (L.layout || L.bind || L.spread) {
          const items = localPlan(ed.paragraphs(scope));
          const cnt = { layout: 0, bind: 0, spread: 0 }, smp = { layout: [], bind: [], spread: [] };
          const where = it => `หน้า ${it.page + 1} ย่อหน้า ${it.n + 1}`;
          for (const it of items) {
            const { letterspacing, ...block } = it.changes;
            const lay = L.layout ? block : {};
            const ls = L.spread && letterspacing != null && Math.abs(letterspacing) <= 0.03 ? letterspacing : null;
            const bind = L.bind ? it.bind : [], join = L.bind && !prof.stripWJ ? (it.join || []) : [];
            if (!Object.keys(lay).length && ls == null && !bind.length && !join.length) continue;
            if (join.length) addNoBreak(join, true);
            ed.applyParagraph(it.page, it.n, lay, bind, ls != null ? { letterspacing: ls } : {}, join);
            if (Object.keys(lay).length) { cnt.layout++; smp.layout.push(`${where(it)}: ${describeChanges({ ...it, changes: lay, bind: [], join: [] })}`); }
            if (bind.length || join.length) { cnt.bind++; smp.bind.push(`${where(it)}: ${[...bind, ...join].map(w => `“${auVis(w)}”`).join(', ')}`); }
            if (ls != null) { cnt.spread++; smp.spread.push(`${where(it)}: ช่องไฟ ${signedEm(ls)}`); }
          }
          const lab = { layout: 'จัดย่อหน้าให้สอดคล้อง', bind: 'ตัดคำ: ผูกตัวเลขกับหน่วย/ชื่อ', spread: 'กระจายตัวอักษร: ปรับช่องไฟ' };
          for (const k of ['layout', 'bind', 'spread']) if (L[k]) log.push({ key: k, label: lab[k], count: cnt[k], samples: smp[k].slice(0, 3) });
        }
        if (aiOk && opts.aiLayout) await aiStep('aiLayout', 'AI วางแผนจัดหน้า (ตัดคำ/กระจายตัวอักษร/ระยะบรรทัด)', async samples => {
          let n = 0;
          const all = ed.paragraphs(scope);
          for (const p of pageList()) {
            const paras = all.filter(x => x.page === p);
            if (!paras.length) continue;
            progress(`AI วางแผนจัดหน้า หน้า ${p + 1}…`);
            const data = await askAiCached('layout', JSON.stringify(layoutPayload(paras)));
            for (const it of fromAi(data, paras)) {
              const { letterspacing, ...block } = it.changes;
              const ls = letterspacing != null && Math.abs(letterspacing) <= 0.05 ? letterspacing : null;
              const join = prof.stripWJ ? [] : (it.join || []);
              if (!Object.keys(block).length && ls == null && !it.bind.length && !join.length) continue;
              if (join.length) addNoBreak(join, true);
              ed.applyParagraph(it.page, it.n, block, it.bind, ls != null ? { letterspacing: ls } : {}, join);
              n++;
              if (samples.length < 12) samples.push(`หน้า ${it.page + 1} ย่อหน้า ${it.n + 1}: ${describeChanges({ ...it, changes: { ...block, ...(ls != null ? { letterspacing: ls } : {}) }, join })}${it.reason ? ` — ${it.reason}` : ''}`);
            }
          }
          return n;
        });
        if (opts.bind) {
          if (prof.stripWJ) log.push({ key: 'wj', label: 'ผูกคำในพจนานุกรม (ข้าม)', count: 0, samples: ['ผลทดสอบวางบอกว่า e-Office ล้าง Word Joiner จึงไม่ผูกคำ'] });
          else step('wj', 'ผูกคำในพจนานุกรมคำห้ามตัด', () => ed.bindWords(scope, noBreak));
        }
        if (opts.pagenum || opts.digits !== 'keep') step('pagenum', 'เลขหน้า / เลขรายการ', samples => {
          let n = 0;
          const set = (k, v, text) => { if (ed.settings()[k] !== v) { ed.setSetting(k, v); n++; samples.push(text); } };
          if (opts.pagenum) set('pageNumbers', true, 'เปิดเลขหน้า “- ๒ -” ตอนพิมพ์');
          if (opts.digits !== 'keep') { set('pageDigits', opts.digits, `เลขหน้าเป็น${opts.digits === 'thai' ? 'เลขไทย' : 'เลขอารบิก'}`); set('listDigits', opts.digits, `เลขรายการเป็น${opts.digits === 'thai' ? 'เลขไทย' : 'เลขอารบิก'}`); }
          return n;
        });
        // ---- AI ตรวจรายงาน (ทำท้ายสุด ตรวจข้อความฉบับสุดท้าย ไม่แก้เอกสาร) ----
        if (aiOk && opts.aiAudit) await aiStep('aiAudit', 'AI ตรวจโครงสร้างและความสอดคล้อง (รายงาน)', async samples => {
          progress('AI ตรวจโครงสร้างหนังสือ…');
          const text = pageList().map(p => `===== หน้า ${p + 1} =====\n${pageTextOf(p).replace(/\n+$/, '').replace(/\u2060/g, '')}`).join('\n');
          if (text.length > AI_MAX_CHARS) { samples.push(`เอกสารยาวเกิน ${AI_MAX_CHARS.toLocaleString()} ตัวอักษร — เลือกขอบเขต “หน้านี้”`); return 0; }
          const data = await askAiCached('audit', text);
          if (data.summary) samples.push('สรุป: ' + data.summary);
          for (const nt of data.notes || []) samples.push(`[${nt.severity}] ${nt.where}: ${nt.issue} → ${nt.suggestion}`);
          return (data.notes || []).length;
        }, { max: 14, report: true });

        const after = ed.exportHtml();
        if (restore) auRestore(undoInfo);
        return { log, before, after, undoInfo };
      } catch (err) {
        auRestore(undoInfo);          // เกิดข้อผิดพลาดกลางทาง: คืนเอกสารทั้งหมดก่อนแจ้ง ไม่ปล่อยให้ค้างครึ่งทาง
        throw err;
      }
    }
    function auRestore(info) {
      ed.restoreAll(info.snap);
      for (const [k, v] of Object.entries(info.settings0)) if (ed.settings()[k] !== v) ed.setSetting(k, v);
    }
    function auReadOpts() {
      for (const c of $$('[data-au]', panel)) auOpts[c.dataset.au] = c.checked;
      auOpts.scope = $('#auScope').value;
      auOpts.digits = $('#auDigits').value;
      store.set(AU_KEY, JSON.stringify(auOpts));
      return { ...auOpts };
    }
    // ประเมินปริมาณที่จะส่งให้ AI (ตัวอักษรของขอบเขตที่เลือก × จำนวนฟังก์ชัน AI) และแสดงผู้ให้บริการที่ใช้
    function auAiInfo() {
      const on = AI_FUNCS.filter(k => $(`[data-au=${k}]`, panel).checked);
      const info = $('#auAiInfo');
      for (const k of ['layout', 'spread']) $(`[data-au=${k}]`, panel).disabled = $('[data-au=aiLayout]', panel).checked;
      if (!on.length) { info.textContent = 'ไม่ได้เลือกฟังก์ชัน AI — ทำงานในเครื่องทั้งหมด ไม่ส่งข้อมูลออกนอกเครื่อง'; return; }
      const scope = $('#auScope').value;
      const idx = scope === 'all' ? ed.pages().map((_, i) => i) : [ed.currentPage()];
      const chars = idx.reduce((n, p) => n + pageTextOf(p).length, 0);
      const dest = staysLocal() ? 'เซิร์ฟเวอร์ในเครื่อง (ไม่ออกนอกเครื่อง)' : `${PROVIDERS[provider].label} (${modelOf(provider)})`;
      const pre = aiPrecheck();
      info.innerHTML = `${esc(dest)} · ส่งประมาณ ${chars.toLocaleString()} ตัวอักษร × ${on.length} ฟังก์ชัน · ดูตัวอย่างจะเรียก AI จริง (มีค่าใช้จ่าย) แต่ผลที่ได้ถูกเก็บไว้ ถ้ากด “จัดทั้งหมด” ต่อทันทีจะไม่เรียกซ้ำ${pre ? ` · <b style="color:#b91c1c">⚠ ${esc(pre)}</b>` : ''}`;
    }
    function renderAuto() {
      for (const c of $$('[data-au]', panel)) c.checked = !!auOpts[c.dataset.au];
      $('#auScope').value = auOpts.scope;
      $('#auDigits').value = auOpts.digits;
      $('#auUndoBtn').disabled = !auUndo;
      auAiInfo();
    }
    panel.addEventListener('change', e => { if (e.target.closest('[data-panel="auto"]')) auAiInfo(); });
    function auShow(which) {
      if (!auFrames) return;
      $('#auCompare').hidden = false;
      $('#auShowing').textContent = which === 'before' ? 'ก่อนจัด (หน้าที่เปิดอยู่)' : 'หลังจัด (หน้าที่เปิดอยู่)';
      $('#auFrame').srcdoc = `<!doctype html><meta charset="utf-8"><body style="margin:12px;font-family:Tahoma,sans-serif">${auFrames[which]}</body>`;
    }
    function auRender(res, applied) {
      const total = res.log.filter(s => !s.report).reduce((n, s) => n + s.count, 0);
      const notes = res.log.filter(s => s.report).reduce((n, s) => n + s.count, 0);
      $('#auLog').innerHTML = res.log.map(s => `<li><div class="au-step"><span class="au-cnt${s.count ? '' : ' zero'}">${s.count}</span><span>${esc(s.label)}</span></div>${s.samples.map(x => `<span class="au-ex">${esc(x)}</span>`).join('')}</li>`).join('')
        || '<li class="hint">ไม่ได้เลือกขั้นตอนใดเลย</li>';
      $('#auStatus').textContent = applied
        ? (total ? `จัดแล้ว ${total} จุด — ย้อนกลับทั้งเอกสารได้ด้วยปุ่ม “ย้อนกลับทั้งหมด”` : 'ไม่มีจุดที่ต้องแก้ เอกสารเรียบร้อยอยู่แล้ว')
        : (total ? `ตัวอย่าง: จะเปลี่ยน ${total} จุด (ยังไม่เปลี่ยนเอกสาร) — กด “จัดทั้งหมด” เพื่อใช้จริง` : 'ตัวอย่าง: ไม่มีจุดที่ต้องแก้');
      if (notes) $('#auStatus').textContent += ` · ข้อสังเกตจาก AI ${notes} ข้อ (รายงานอย่างเดียว)`;
      auFrames = { before: res.before, after: res.after };
      auShow('after');
    }
    Object.assign(actions, {
      auPreview() { return auGo(false); },
      auApply() { return auGo(true); },
      auUndo() {
        if (!auUndo) return;
        auRestore(auUndo);
        auUndo = null;
        $('#auUndoBtn').disabled = true;
        $('#auStatus').textContent = 'ย้อนกลับการจัดอัตโนมัติแล้ว (ทุกหน้า)';
        say('ย้อนกลับการจัดอัตโนมัติทั้งหมดแล้ว');
      },
      auShow(btn) { auShow(btn.dataset.which); },
    });
    // รันตัวจัดอัตโนมัติ (ตัวอย่างหรือจริง) พร้อมล็อกหน้าจอระหว่างรอ AI เพื่อไม่ให้แก้เอกสารพร้อมกัน
    // ล็อกหน้าจอระหว่างรัน: ใช้แผ่นกั้นเต็มจอ + บล็อกคีย์บอร์ด (ห้ามตั้ง contenteditable=false เพราะ Quill จะทิ้งการแก้ไขแบบ user ทั้งหมดโดยไม่แจ้ง)
    const shield = document.createElement('div');
    shield.id = 'auShield';
    shield.hidden = true;
    shield.innerHTML = '<div class="au-shield-card"><div id="auShieldMsg" role="status">กำลังจัดอัตโนมัติ…</div><button type="button" class="tbtn" id="auCancelBtn">ยกเลิก</button></div>';
    document.body.append(shield);
    let auCancel = false;
    $('#auCancelBtn').addEventListener('click', () => { auCancel = true; $('#auShieldMsg').textContent = 'กำลังยกเลิกและคืนเอกสารเดิม…'; });
    document.addEventListener('keydown', e => { if (auBusy) { e.preventDefault(); e.stopPropagation(); } }, true);
    function auLock(on) {
      auBusy = on;
      auCancel = false;
      shield.hidden = !on;
      if (on) { document.activeElement?.blur?.(); $('#auShieldMsg').textContent = 'กำลังจัดอัตโนมัติ…'; }
      for (const b of $$('[data-act^=au]', panel)) if (b.dataset.act !== 'auShow') b.disabled = on || (b.dataset.act === 'auUndo' && !auUndo);
    }
    async function auGo(apply) {
      if (auBusy) return;
      const o = auReadOpts();
      if (apply && o.scope === 'all' && !confirm(`จัดรูปแบบอัตโนมัติทุกหน้า (${ed.pages().length} หน้า)? ย้อนกลับทั้งหมดได้ด้วยปุ่ม “ย้อนกลับทั้งหมด”`)) return;
      auLock(true);
      try {
        const res = await auRun(o, { restore: !apply });
        if (apply) auUndo = res.undoInfo;
        auRender(res, apply);
        say(apply ? 'จัดรูปแบบอัตโนมัติแล้ว' : 'ดูตัวอย่างการจัดอัตโนมัติแล้ว — เอกสารยังไม่เปลี่ยน');
      } catch (err) {
        $('#auStatus').textContent = (auCancel ? 'ยกเลิกแล้ว (คืนเอกสารเดิม)' : 'ไม่สำเร็จ (คืนเอกสารเดิมแล้ว): ' + (err?.message || err));
      } finally { auLock(false); }
    }
    onTab.auto = renderAuto;
    renderAuto();

    showTab(store.get(KEY.tab) || 'num');
    if (store.get(KEY.panel) === '1') setOpen(true);
    window.eofficeTools = { bahtText, thaiDate, scanOffline, replaceInDoc, toThai, toArabic };
  }
})();
