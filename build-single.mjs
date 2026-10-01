// สร้างชุดไฟล์พร้อม deploy ลงโฟลเดอร์ dist/  —  ใช้งาน: node build-single.mjs   (ต้องมี Node.js 18+) รันซ้ำทุกครั้งที่แก้โค้ด
//
//   dist/index.html                    หน้าสำหรับ deploy ขึ้นเว็บ: รวมทุกอย่างในไฟล์เดียว + นโยบายความปลอดภัย (CSP) ที่คำนวณแฮชจากโค้ดจริง
//   dist/vendor/anthropic-sdk-*.mjs    SDK ของ Anthropic ที่ล็อกเวอร์ชัน (โหลดเมื่อใช้ AI ของ Claude เท่านั้น)
//   dist/_headers                      header ความปลอดภัย (Cloudflare Pages / Netlify อ่านไฟล์นี้อัตโนมัติ)
//   dist/robots.txt                    ห้ามเครื่องมือค้นหาจัดทำดัชนี (เป็นเครื่องมือภายใน)
//   dist/eoffice-editor-single.html    ไฟล์เดียวสำหรับส่งอีเมล/แชร์ไดรฟ์ ฝัง SDK ไว้ในไฟล์ ไม่มี CSP เปิดจาก file:// ได้
import { readFileSync, writeFileSync, mkdirSync, existsSync, copyFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';

const root = dirname(fileURLToPath(import.meta.url));
const read = f => readFileSync(join(root, f), 'utf8');
const must = (cond, what) => { if (!cond) { console.error('ไม่พบ/ไม่ถูกต้อง: ' + what); process.exit(1); } };

// ป้องกันข้อความในโค้ดไปปิดแท็ก <script> / เปิดคอมเมนต์ HTML ก่อนเวลา (ค่าในสตริงยังเหมือนเดิม) และตัด source map
const safeJs = js => js.replaceAll('</script', '<\\/script').replaceAll('<!--', '<\\!--').replace(/\n\/\/# sourceMappingURL=.*$/m, '');
const safeCss = css => css.replaceAll('</style', '<\\/style').replace(/\/\*# sourceMappingURL=.*?\*\//g, '');

// ---- SDK ของ Anthropic: ชื่อไฟล์ต้องตรงกับที่ eoffice-tools.js อ้างถึง ----
const sdkRef = /const SDK_FILE = '\.\/vendor\/([^']+\.mjs)'/.exec(read('eoffice-tools.js'));
must(sdkRef, 'ค่า SDK_FILE ใน eoffice-tools.js');
const sdkName = sdkRef[1];
must(existsSync(join(root, 'vendor', sdkName)), `vendor/${sdkName} (ไฟล์ SDK ที่ eoffice-tools.js อ้างถึง)`);
const sdk = read(`vendor/${sdkName}`);

// ---- ประกอบหน้าเดียว (ไม่รวม SDK / CSP) ----
let html = read('eoffice-editor.html');
const linkRe = /<link rel="stylesheet" href="\.\/quill\.snow\.css"[^>]*>/;
must(linkRe.test(html), 'แท็ก <link> ของ quill.snow.css');
html = html.replace(linkRe, () => `<style>${safeCss(read('quill.snow.css'))}</style>`);

const tagRe = /<script src="\.\/eoffice-pastetest\.js"[^>]*><\/script>\s*<script src="\.\/eoffice-tools\.js"[^>]*><\/script>/;
must(tagRe.test(html), 'แท็ก <script> ของ eoffice-pastetest.js / eoffice-tools.js');
html = html.replace(tagRe, '');

// Quill + pastetest + tools ฝังก่อนสคริปต์หลัก (สคริปต์หลักข้ามการโหลด Quill เมื่อ window.Quill มีแล้ว)
const mainTag = '\n<script>\n';
const at = html.indexOf(mainTag);
must(at > 0 && html.indexOf(mainTag, at + 1) === -1, 'แท็ก <script> หลักตัวเดียว');
const inlined = ['quill.js', 'eoffice-pastetest.js', 'eoffice-tools.js'].map(f => `<script>/* ${f} */\n${safeJs(read(f))}\n</script>`).join('\n');
const base = html.slice(0, at) + '\n' + inlined + html.slice(at);
must(!/\son(?:error|load|click)=/i.test(base.replace(/<script>[\s\S]*?<\/script>/g, '')), 'ไม่ควรมี event handler แบบ inline (HTML attribute) เหลือ — จะขัดกับ CSP');

// ---- CSP: อนุญาตสคริปต์เฉพาะที่มีแฮชตรงกับโค้ดจริง ----
const hashes = [...base.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => `'sha256-${createHash('sha256').update(m[1], 'utf8').digest('base64')}'`);
must(hashes.length === 4, `สคริปต์ inline 4 ก้อน (พบ ${hashes.length})`);
const csp = [
  "default-src 'none'",
  `script-src 'self' ${hashes.join(' ')}`,
  "style-src 'self' 'unsafe-inline'",                 // Quill และแอปตั้ง style แบบ inline ทั่วไป (CSS ไม่รันโค้ด)
  "img-src 'self' data: blob:",
  "font-src 'self' data:",
  // ปลายทาง AI ที่รองรับ + เซิร์ฟเวอร์ในเครื่อง (Ollama ฯลฯ) — เซิร์ฟเวอร์ OpenAI-compatible ของหน่วยงานให้เพิ่มชื่อโดเมนที่นี่
  "connect-src 'self' https://api.anthropic.com https://api.openai.com https://generativelanguage.googleapis.com http://localhost:* http://127.0.0.1:*",
  "base-uri 'none'", "form-action 'none'", "object-src 'none'", "manifest-src 'none'",
];
const cspMeta = csp.join('; ');
const cspHeader = [...csp, "frame-ancestors 'none'", 'upgrade-insecure-requests'].join('; ');   // 2 ข้อนี้ใช้ได้เฉพาะแบบ header

const hosted = base.replace('<meta charset="utf-8">',
  `<meta charset="utf-8">\n<meta http-equiv="Content-Security-Policy" content="${cspMeta}">\n<meta name="referrer" content="no-referrer">\n<meta name="robots" content="noindex,nofollow">`);
must(hosted !== base, 'แท็ก <meta charset="utf-8">');

// ---- ไฟล์เดี่ยวสำหรับส่งอีเมล: ฝัง SDK (โหลดผ่าน Blob) ไม่มี CSP ----
const firstInline = '\n<script>/* quill.js */';
must(base.includes(firstInline), 'จุดแทรก SDK (สคริปต์ quill.js)');
const single = base.replace(firstInline, () => `\n<script type="text/plain" id="sdk-embedded">${safeJs(sdk)}</script>${firstInline}`);

// ---- เขียนไฟล์ ----
const dist = join(root, 'dist');
mkdirSync(join(dist, 'vendor'), { recursive: true });
writeFileSync(join(dist, 'index.html'), hosted);
writeFileSync(join(dist, 'eoffice-editor-single.html'), single);
copyFileSync(join(root, 'vendor', sdkName), join(dist, 'vendor', sdkName));
// ประกาศลิขสิทธิ์ของ Quill และ SDK (เงื่อนไขของใบอนุญาตให้แจกจ่ายพร้อมกัน)
must(existsSync(join(root, 'THIRD-PARTY-NOTICES.md')), 'THIRD-PARTY-NOTICES.md');
copyFileSync(join(root, 'THIRD-PARTY-NOTICES.md'), join(dist, 'THIRD-PARTY-NOTICES.md'));
writeFileSync(join(dist, 'robots.txt'), 'User-agent: *\nDisallow: /\n');
writeFileSync(join(dist, '_headers'), `# Header ความปลอดภัย — Cloudflare Pages และ Netlify อ่านไฟล์นี้อัตโนมัติ (สร้างโดย build-single.mjs ห้ามแก้ตรง ๆ)
/*
  Content-Security-Policy: ${cspHeader}
  X-Content-Type-Options: nosniff
  Referrer-Policy: no-referrer
  Cross-Origin-Opener-Policy: same-origin
  Cross-Origin-Resource-Policy: same-origin
  X-Frame-Options: DENY
  Permissions-Policy: camera=(), microphone=(), geolocation=(), payment=(), usb=(), serial=(), interest-cohort=()

# หน้าหลักต้องตรวจเวอร์ชันใหม่ทุกครั้ง เพื่อให้การอัปเดตถึงผู้ใช้ทันที (ไม่ต้องกด Ctrl+F5)
/
  Cache-Control: no-cache
/index.html
  Cache-Control: no-cache

# ไฟล์ SDK มีเลขเวอร์ชันในชื่อ เก็บแคชได้นาน
/vendor/*
  Cache-Control: public, max-age=31536000, immutable
`);

const kb = s => `${(Buffer.byteLength(s) / 1024).toFixed(0)} KB`;
console.log(`dist/index.html                    ${kb(hosted)}  (CSP: ${hashes.length} แฮช)`);
console.log(`dist/eoffice-editor-single.html    ${kb(single)}  (ฝัง SDK)`);
console.log(`dist/vendor/${sdkName}  ${kb(sdk)}`);
console.log('dist/_headers, dist/robots.txt');
