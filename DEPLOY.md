# คู่มือ deploy

แอปนี้เป็นไฟล์สถิต ไม่มีเซิร์ฟเวอร์และฐานข้อมูล สิ่งที่ deploy คือโฟลเดอร์ **`dist/`** ทั้งโฟลเดอร์

| ไฟล์ใน `dist/` | หน้าที่ |
|---|---|
| `index.html` | แอปทั้งหมดในไฟล์เดียว + นโยบายความปลอดภัย (CSP) ที่ผูกกับโค้ดด้วยแฮช |
| `vendor/anthropic-sdk-0.131.0.mjs` | SDK ของ Anthropic (ล็อกเวอร์ชัน โหลดเมื่อใช้ AI ของ Claude เท่านั้น) |
| `_headers` | header ความปลอดภัย + การแคช (Cloudflare Pages / Netlify อ่านอัตโนมัติ) |
| `robots.txt` | ห้ามเครื่องมือค้นหาจัดทำดัชนี |
| `eoffice-editor-single.html` | ไฟล์เดียวสำหรับส่งอีเมล (ฝัง SDK ไว้ในไฟล์ ไม่มี CSP) |

## ก่อน deploy ทุกครั้ง

```
node build-single.mjs
```

แล้วตรวจว่าไม่มีข้อความผิดพลาด **ห้ามแก้ไฟล์ใน `dist/` ตรง ๆ** (จะถูกเขียนทับ และแฮช CSP จะไม่ตรงกับโค้ด ทำให้แอปไม่ทำงาน) ให้แก้ที่ไฟล์ต้นฉบับแล้ว build ใหม่

ลองทดสอบในเครื่องก่อนขึ้นจริง: `python -m http.server 8080 --directory dist` แล้วเปิด http://localhost:8080

---

## วิธีที่แนะนำ: Cloudflare (ฟรี อัปเดตด้วย `git push`)

1. สร้าง repository บน GitHub (เลือก **Private** ได้ ฟรี) แล้ว push โฟลเดอร์โปรเจกต์นี้ขึ้นไป
   (`.gitignore` ตัดไฟล์ในเครื่องออกให้แล้ว)
2. เข้า Cloudflare Dashboard → **Workers & Pages** → **Create application** → **Connect to Git** → เลือก repository
   (ตอนนี้ Cloudflare รวม Pages เข้ากับ Workers แล้ว หน้าจอที่เห็นจะเป็นแบบ Workers ซึ่งใช้ได้เหมือนกัน)
3. ตั้งค่าตามนี้ — **ค่าอื่นปล่อยตามค่าเริ่มต้น**:

   | ช่อง | ค่า |
   |---|---|
   | Build command | **เว้นว่าง** (โฟลเดอร์ `dist` ถูก build และ commit ไว้แล้ว) หรือ `node build-single.mjs` |
   | Deploy command | `npx wrangler deploy` (ค่าเริ่มต้น ไม่ต้องแก้) |
   | Root directory | `/` |

   คำสั่ง deploy อ่านค่าจากไฟล์ **`wrangler.jsonc`** ในโปรเจกต์ (บอกว่าเสิร์ฟโฟลเดอร์ `dist`) ถ้าไม่มีไฟล์นี้ ขั้น Deploying จะล้มเหลว
   และ **ชื่อ** ใน `wrangler.jsonc` (`"name"`) ต้องตรงกับชื่อโปรเจกต์ที่ตั้งใน Cloudflare
4. กด **Deploy** จะได้ที่อยู่แบบ `ชื่อโปรเจกต์.บัญชี.workers.dev` (หรือ `.pages.dev` ถ้าสร้างเป็น Pages)

> ถ้าสร้างเป็น **Pages** (ทางเลือกอยู่ที่ลิงก์เล็ก ๆ ใต้ปุ่มเลือกประเภทตอน Create) ให้ตั้ง Framework preset `None`,
> Build command เว้นว่าง, Build output directory `dist` — ไม่ต้องใช้ `wrangler.jsonc`
5. **อัปเดตภายหลัง:** แก้โค้ด → `node build-single.mjs` → `git commit` → `git push` เว็บจะอัปเดตเองใน 1–2 นาที
   (หน้าหลักตั้งให้ตรวจเวอร์ชันใหม่ทุกครั้ง ผู้ใช้ไม่ต้องกด Ctrl+F5)

### จำกัดว่าใครเข้าได้ (ฟรีไม่เกิน 50 ผู้ใช้)

Cloudflare **Zero Trust** → **Access** → **Applications** → **Add an application** → **Self-hosted**
ใส่โดเมนของโปรเจกต์ แล้วสร้าง Policy แบบ *Allow* เฉพาะอีเมลที่ลงท้ายด้วยโดเมนหน่วยงาน (เช่น `@agency.go.th`)
ผู้ใช้จะได้รับรหัสยืนยันทางอีเมลก่อนเข้าแอป

### โดเมนของตัวเอง

ตัดสินใจ **ก่อนแจกจ่ายให้ผู้ใช้** เพราะร่างที่จำไว้ในเบราว์เซอร์ผูกกับโดเมน ถ้าเปลี่ยนโดเมนภายหลัง ร่างและการตั้งค่าที่ผู้ใช้เก็บไว้จะไม่ตามไป

---

## ตัวเลือกอื่น

**Netlify** — ลากโฟลเดอร์ `dist` ไปวางที่หน้า Deploys หรือเชื่อม Git แบบเดียวกับ Cloudflare (Build command `node build-single.mjs`, Publish directory `dist`) อ่าน `_headers` ได้เหมือนกัน
แต่การใส่รหัสผ่านหน้าเว็บเป็นฟีเจอร์เสียเงิน

**GitHub Pages** — ใช้ได้ แต่: (1) repository ต้อง public เว้นแต่มีแผนองค์กร (2) ไม่อ่าน `_headers` จึงเหลือเฉพาะ CSP ใน `<meta>` ของ `index.html` (ไม่มี `X-Frame-Options`, `Referrer-Policy` ฯลฯ) (3) แคช 10 นาที ผู้ใช้อาจเห็นของเก่าชั่วคราวหลังอัปเดต
ตั้ง Pages ให้เสิร์ฟจากโฟลเดอร์ `dist` (ผ่าน GitHub Actions หรือ branch `gh-pages`)

**เซิร์ฟเวอร์ภายในหน่วยงาน (ไม่ออกอินเทอร์เน็ตเลย)** — คัดลอกโฟลเดอร์ `dist` ไปวางบนเว็บเซิร์ฟเวอร์ใด ๆ ที่เสิร์ฟไฟล์สถิตได้
และใส่ header จากไฟล์ `dist/_headers` ในการตั้งค่าเซิร์ฟเวอร์ (ตัวอย่าง nginx: `add_header Content-Security-Policy "…" always;` ทีละบรรทัด)
ทุกครั้งที่ build ใหม่ แฮชใน CSP จะเปลี่ยน ต้องนำค่าใหม่จาก `_headers` ไปใส่ด้วย ไม่เช่นนั้นแอปจะถูกบล็อกสคริปต์

---

## นโยบายความปลอดภัย (CSP) ที่ตั้งไว้

- สคริปต์รันได้เฉพาะโค้ดที่มีแฮชตรงกับ `index.html` และไฟล์จากโดเมนเดียวกัน — สคริปต์แปลกปลอมที่แทรกเข้ามาจะไม่ทำงาน
- ติดต่อออกนอกเครื่องได้เฉพาะ `api.anthropic.com`, `api.openai.com`, `generativelanguage.googleapis.com` และเซิร์ฟเวอร์ในเครื่อง (`localhost`)
- ไม่ให้ฝังหน้านี้ในเว็บอื่น ไม่ส่ง referrer ไม่ให้เครื่องมือค้นหาจัดทำดัชนี

**ถ้าหน่วยงานมีเซิร์ฟเวอร์ AI ของตัวเอง** (แบบ OpenAI-compatible) ต้องเพิ่มโดเมนในรายการ `connect-src` ที่ไฟล์ `build-single.mjs` แล้ว build ใหม่
ไม่เช่นนั้นเบราว์เซอร์จะบล็อกการเชื่อมต่อ (ใช้ได้เฉพาะ `localhost` เช่น Ollama)

## ปรับปรุง SDK ของ Anthropic (ถ้าต้องการเวอร์ชันใหม่)

SDK ถูก build เป็นไฟล์เดียวไว้ใน `vendor/` เพื่อไม่ให้โค้ดจากภายนอกมารันในหน้าที่มีคีย์ AI ของผู้ใช้ ขั้นตอนเปลี่ยนเวอร์ชัน (ในโฟลเดอร์ชั่วคราว ไม่ใช่ในโปรเจกต์):

```
npm install --save-exact @anthropic-ai/sdk@<เวอร์ชัน> esbuild
npx esbuild node_modules/@anthropic-ai/sdk/index.mjs --bundle --format=esm --platform=browser --target=es2022 --minify --legal-comments=none --outfile=anthropic-sdk-<เวอร์ชัน>.mjs
```

แล้วคัดลอกไฟล์ไปไว้ใน `vendor/`, แก้ชื่อไฟล์ในค่า `SDK_FILE` ของ `eoffice-tools.js`, ลบไฟล์เวอร์ชันเก่า และรัน `node build-single.mjs`
(ตัวสร้างจะหยุดพร้อมแจ้ง ถ้าชื่อไฟล์ที่อ้างถึงไม่มีอยู่จริง) ควรทดสอบฟังก์ชัน AI หลังเปลี่ยนเวอร์ชันทุกครั้ง

## เช็กลิสต์ก่อนเปิดให้ผู้ใช้

- [ ] `node build-single.mjs` ผ่านโดยไม่มีข้อผิดพลาด
- [ ] เปิดที่อยู่ที่ deploy แล้วเอดิเตอร์และปุ่ม “เครื่องมือเสริม” ทำงาน (กด F12 → Console ต้องไม่มีข้อความสีแดงเรื่อง Content Security Policy)
- [ ] repository ไม่มีรหัสผ่านหรือ API key (แอปไม่เก็บคีย์ไว้ในโค้ด ผู้ใช้กรอกเอง)
- [ ] ตั้งการจำกัดผู้เข้าใช้ (Cloudflare Access) ถ้าเป็นเครื่องมือภายใน
- [ ] ทดสอบ “ทดสอบวาง” กับ e-Office จริง 1 ครั้ง (แท็บ “ทดสอบวาง” ในเครื่องมือเสริม)
- [ ] บอกผู้ใช้ว่าต้องติดตั้งฟอนต์ TH Sarabun PSK ในเครื่องเอง (ไม่ได้รวมไว้ในแอป)
