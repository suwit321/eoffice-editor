// eoffice-cloud.js — ระบบ Cloud Storage สำหรับ E-Office Editor
// รองรับ Google Drive (บันทึกลงโฟลเดอร์เฉพาะ "E-Office Documents") และ Firebase Firestore
(() => {
  'use strict';

  // ================= Storage Keys & Constants =================
  const FOLDER_NAME = 'E-Office Documents';
  const KEY_PROVIDER = 'eoffice-editor:cloud-provider';        // 'gdrive' (default) | 'firebase'
  const KEY_GDRIVE_CLIENT_ID = 'eoffice-editor:gdrive-client-id';
  const KEY_FIREBASE_CFG = 'eoffice-editor:firebase';
  const KEY_CURRENT_DOC = 'eoffice-editor:cloud-doc';
  const KEY_AUTOSAVE = 'eoffice-editor:cloud-autosave';
  const KEY_SESSION_TOKEN = 'eoffice:gdrive:token';
  const KEY_SESSION_EXPIRY = 'eoffice:gdrive:expiry';
  const AUTOSAVE_DELAY = 3000;    // บันทึกอัตโนมัติ 3 วินาทีหลังพิมพ์เสร็จ

  const GIS_SRC = 'https://accounts.google.com/gsi/client';
  const DRIVE_API = 'https://www.googleapis.com/drive/v3';
  const DRIVE_UPLOAD = 'https://www.googleapis.com/upload/drive/v3';
  const USERINFO_API = 'https://www.googleapis.com/oauth2/v3/userinfo';
  const GDRIVE_SCOPES = [
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/userinfo.profile',
    'https://www.googleapis.com/auth/userinfo.email',
  ].join(' ');

  // Firebase Constants (เป็นทางเลือกรอง)
  const FIREBASE_VERSION = '10.14.1';
  const FIREBASE_CDN = `https://www.gstatic.com/firebasejs/${FIREBASE_VERSION}`;
  const FIREBASE_COLLECTION = 'documents';

  // ================= State =================
  let provider = 'gdrive';
  try {
    const savedProv = localStorage.getItem(KEY_PROVIDER);
    if (savedProv === 'firebase') provider = 'firebase';
  } catch {}

  let ed = null;                      // window.eofficeEditor
  let user = null;                    // { uid, name, email, photo }
  let currentDocId = null;            // Google Drive File ID หรือ Firestore Doc ID
  let currentDocLink = null;          // webViewLink สำหรับเปิดใน Google Drive
  let folderId = null;                // Google Drive Folder ID
  let folderLink = null;              // ลิงก์ไปยังโฟลเดอร์ใน Google Drive
  let autoSaveEnabled = false;
  let autoSaveTimer = null;
  let isSaving = false;
  let lastSavedJson = '';

  // Google Drive State
  let gisLoaded = false;
  let gisLoading = null;
  let tokenClient = null;
  let accessToken = null;
  let tokenExpiry = 0;

  // Firebase State
  let firebase = null;
  let auth = null;
  let db = null;
  let fbInitialized = false;

  // ================= Utilities =================
  function emit(name, detail = {}) {
    document.dispatchEvent(new CustomEvent(name, { detail }));
  }

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      if (document.querySelector(`script[src="${src}"]`)) return resolve();
      const s = document.createElement('script');
      s.src = src;
      s.onload = resolve;
      s.onerror = () => reject(new Error(`โหลดสคริปต์ไม่สำเร็จ: ${src}`));
      document.head.appendChild(s);
    });
  }

  // ================= Serialization =================
  function serializeDoc() {
    if (!ed) return null;
    const snap = ed.snapshotAll();
    const data = {
      v: 2,
      name: ed.docName(),
      cur: snap.cur,
      pages: snap.pages.map(p => p.delta.ops),
    };
    try {
      const s = ed.settings();
      if (s) data.settings = s;
    } catch {}
    return JSON.stringify(data);
  }

  function deserializeDoc(docData) {
    if (!ed || !docData || !Array.isArray(docData.pages)) return false;
    const { Delta } = ed;
    if (docData.name) {
      const cleanName = docData.name.replace(/\.eoffice$/i, '');
      ed.docName(cleanName);
    }
    if (docData.settings) {
      try { ed.settings(docData.settings); } catch {}
    }
    const pageDeltas = docData.pages.map(ops => new Delta(ops));
    ed.loadPages(pageDeltas, docData.cur || 0);
    return true;
  }

  // ================= Google Identity Services (GIS) =================
  function getGdriveClientId() {
    try { return (localStorage.getItem(KEY_GDRIVE_CLIENT_ID) || '').trim(); } catch { return ''; }
  }

  function setGdriveClientId(id) {
    try {
      if (id) localStorage.setItem(KEY_GDRIVE_CLIENT_ID, id.trim());
      else localStorage.removeItem(KEY_GDRIVE_CLIENT_ID);
    } catch {}
    tokenClient = null;
  }

  async function loadGIS() {
    if (gisLoaded && window.google?.accounts?.oauth2) return;
    if (gisLoading) return gisLoading;
    gisLoading = (async () => {
      await loadScript(GIS_SRC);
      gisLoaded = true;
    })();
    return gisLoading;
  }

  async function initGIS() {
    const clientId = getGdriveClientId();
    if (!clientId) return false;
    await loadGIS();

    if (!window.google?.accounts?.oauth2) {
      throw new Error('ไม่สามารถโหลด Google Identity Services ได้');
    }

    tokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: GDRIVE_SCOPES,
      callback: async (resp) => {
        if (resp.error) {
          console.error('[E-Office Cloud] Google Auth Error:', resp);
          emit('eoffice:cloud:error', { message: 'เข้าสู่ระบบไม่สำเร็จ: ' + (resp.error_description || resp.error) });
          return;
        }
        accessToken = resp.access_token;
        tokenExpiry = Date.now() + (Number(resp.expires_in) || 3599) * 1000;
        try {
          sessionStorage.setItem(KEY_SESSION_TOKEN, accessToken);
          sessionStorage.setItem(KEY_SESSION_EXPIRY, String(tokenExpiry));
        } catch {}

        await fetchUserProfile();
        await ensureFolder();
        await listMyDocs();
        emit('eoffice:cloud:ready', { user, folderLink });
      }
    });

    // ตรวจสอบว่ามี Token เดิมในเซสชันที่ยังไม่หมดอายุหรือไม่
    try {
      const savedToken = sessionStorage.getItem(KEY_SESSION_TOKEN);
      const savedExp = Number(sessionStorage.getItem(KEY_SESSION_EXPIRY) || 0);
      if (savedToken && savedExp > Date.now() + 60000) {
        accessToken = savedToken;
        tokenExpiry = savedExp;
        await fetchUserProfile();
        await ensureFolder();
      }
    } catch {}

    return true;
  }

  async function fetchUserProfile() {
    if (!accessToken) return null;
    try {
      const res = await fetch(USERINFO_API, {
        headers: { Authorization: `Bearer ${accessToken}` }
      });
      if (!res.ok) throw new Error('ไม่สามารถดึงข้อมูลโปรไฟล์ได้');
      const info = await res.json();
      user = {
        uid: info.sub,
        name: info.name || info.given_name || 'ผู้ใช้งาน Google',
        email: info.email || '',
        photo: info.picture || ''
      };
      return user;
    } catch (e) {
      console.warn('[E-Office Cloud] UserInfo error:', e);
      user = { uid: 'google-user', name: 'ผู้ใช้งาน Google', email: '', photo: '' };
      return user;
    }
  }

  // ================= Google Drive Folder & File CRUD =================
  async function ensureFolder() {
    if (folderId) return folderId;
    if (!accessToken) return null;

    try {
      // 1. ค้นหาว่ามีโฟลเดอร์ "E-Office Documents" อยู่แล้วหรือไม่
      const q = encodeURIComponent(`name = '${FOLDER_NAME}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false`);
      const searchRes = await fetch(`${DRIVE_API}/files?q=${q}&fields=files(id,name,webViewLink)`, {
        headers: { Authorization: `Bearer ${accessToken}` }
      });

      if (searchRes.ok) {
        const data = await searchRes.json();
        if (data.files && data.files.length > 0) {
          folderId = data.files[0].id;
          folderLink = data.files[0].webViewLink || `https://drive.google.com/drive/folders/${folderId}`;
          return folderId;
        }
      }

      // 2. ถ้ายังไม่มี ให้สร้างโฟลเดอร์ใหม่
      const createRes = await fetch(`${DRIVE_API}/files?fields=id,name,webViewLink`, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          name: FOLDER_NAME,
          mimeType: 'application/vnd.google-apps.folder',
          description: 'โฟลเดอร์จัดเก็บเอกสารราชการสำหรับโปรแกรม E-Office Editor'
        })
      });

      if (!createRes.ok) {
        const err = await createRes.text();
        throw new Error(`สร้างโฟลเดอร์ Google Drive ไม่สำเร็จ: ${err}`);
      }

      const folderData = await createRes.json();
      folderId = folderData.id;
      folderLink = folderData.webViewLink || `https://drive.google.com/drive/folders/${folderId}`;
      return folderId;
    } catch (err) {
      console.error('[E-Office Cloud] ensureFolder error:', err);
      throw err;
    }
  }

  async function saveToDrive(options = {}) {
    if (!accessToken || !user) {
      emit('eoffice:cloud:error', { message: 'กรุณาล็อกอิน Google Drive ก่อนบันทึก' });
      return null;
    }
    if (isSaving) return currentDocId;
    isSaving = true;
    emit('eoffice:cloud:saving', {});

    try {
      const json = serializeDoc();
      if (!json) throw new Error('ไม่สามารถอ่านเนื้อหาเอกสารได้');

      const fid = await ensureFolder();
      const docName = (ed?.docName() || 'เอกสารไม่มีชื่อ').trim();
      const fileName = docName.endsWith('.eoffice') ? docName : `${docName}.eoffice`;

      const boundary = '-------EOfficeMultipartBoundary' + Date.now();
      const delimiter = `\r\n--${boundary}\r\n`;
      const closeDelim = `\r\n--${boundary}--`;

      const metadata = {
        name: fileName,
        mimeType: 'application/json',
        description: 'เอกสาร E-Office สร้างโดย E-Office Editor'
      };

      let url;
      let method;
      let targetFileId = options.docId || currentDocId;

      if (targetFileId) {
        // อัปเดตไฟล์เดิม
        url = `${DRIVE_UPLOAD}/files/${targetFileId}?uploadType=multipart&fields=id,name,modifiedTime,webViewLink`;
        method = 'PATCH';
      } else {
        // สร้างไฟล์ใหม่ในโฟลเดอร์เฉพาะ
        metadata.parents = [fid];
        url = `${DRIVE_UPLOAD}/files?uploadType=multipart&fields=id,name,modifiedTime,webViewLink`;
        method = 'POST';
      }

      const multipartBody =
        delimiter +
        'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
        JSON.stringify(metadata) +
        delimiter +
        'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
        json +
        closeDelim;

      const res = await fetch(url, {
        method,
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': `multipart/related; boundary=${boundary}`
        },
        body: multipartBody
      });

      if (!res.ok) {
        const errText = await res.text();
        throw new Error(`Google Drive API error (${res.status}): ${errText}`);
      }

      const file = await res.json();
      currentDocId = file.id;
      currentDocLink = file.webViewLink || `https://drive.google.com/file/d/${file.id}/view`;
      lastSavedJson = json;
      try { localStorage.setItem(KEY_CURRENT_DOC, file.id); } catch {}

      emit('eoffice:cloud:saved', {
        docId: file.id,
        name: docName,
        webViewLink: currentDocLink
      });
      if (ed) ed.say(`📁 บันทึกลง Google Drive แล้ว — ${docName}`);
      return file.id;
    } catch (err) {
      emit('eoffice:cloud:error', { message: 'บันทึกไม่สำเร็จ: ' + err.message });
      if (ed) ed.say('❌ บันทึกลง Google Drive ไม่สำเร็จ: ' + err.message);
      return null;
    } finally {
      isSaving = false;
    }
  }

  async function loadFromDrive(fileId) {
    if (!accessToken) {
      emit('eoffice:cloud:error', { message: 'กรุณาล็อกอิน Google Drive ก่อนเปิดไฟล์' });
      return null;
    }
    emit('eoffice:cloud:loading', {});

    try {
      // 1. อ่านข้อมูลไฟล์
      const metaRes = await fetch(`${DRIVE_API}/files/${fileId}?fields=id,name,webViewLink`, {
        headers: { Authorization: `Bearer ${accessToken}` }
      });
      if (!metaRes.ok) throw new Error('ไม่พบไฟล์นี้ใน Google Drive');
      const meta = await metaRes.json();

      // 2. อ่านเนื้อหาไฟล์
      const contentRes = await fetch(`${DRIVE_API}/files/${fileId}?alt=media`, {
        headers: { Authorization: `Bearer ${accessToken}` }
      });
      if (!contentRes.ok) throw new Error('ไม่สามารถดาวน์โหลดเนื้อหาไฟล์ได้');
      const data = await contentRes.json();

      if (!data || !data.pages) throw new Error('รูปแบบเอกสาร E-Office ไม่ถูกต้อง');

      deserializeDoc(data);

      currentDocId = fileId;
      currentDocLink = meta.webViewLink || `https://drive.google.com/file/d/${fileId}/view`;
      lastSavedJson = JSON.stringify(data);
      try { localStorage.setItem(KEY_CURRENT_DOC, fileId); } catch {}

      if (autoSaveEnabled) setupAutoSave();

      emit('eoffice:cloud:loaded', {
        docId: fileId,
        name: meta.name,
        webViewLink: currentDocLink
      });
      if (ed) ed.say(`📂 เปิดเอกสารจาก Google Drive สำเร็จ — ${meta.name}`);
      return data;
    } catch (err) {
      emit('eoffice:cloud:error', { message: 'เปิดไฟล์ไม่สำเร็จ: ' + err.message });
      if (ed) ed.say('❌ เปิดจาก Google Drive ไม่สำเร็จ: ' + err.message);
      return null;
    }
  }

  async function listDriveDocs() {
    if (!accessToken) return [];
    try {
      const fid = await ensureFolder();
      const q = encodeURIComponent(`'${fid}' in parents and trashed = false`);
      const res = await fetch(`${DRIVE_API}/files?q=${q}&orderBy=modifiedTime desc&fields=files(id,name,modifiedTime,size,webViewLink)`, {
        headers: { Authorization: `Bearer ${accessToken}` }
      });
      if (!res.ok) throw new Error('ไม่สามารถอ่านรายการไฟล์ได้');
      const data = await res.json();
      const files = (data.files || []).map(f => ({
        id: f.id,
        name: (f.name || '').replace(/\.eoffice$/i, ''),
        rawName: f.name,
        updatedAt: f.modifiedTime ? new Date(f.modifiedTime) : new Date(),
        size: Number(f.size) || 0,
        webViewLink: f.webViewLink
      }));
      emit('eoffice:cloud:list', { docs: files, folderLink });
      return files;
    } catch (err) {
      emit('eoffice:cloud:error', { message: 'โหลดรายการไฟล์ไม่สำเร็จ: ' + err.message });
      return [];
    }
  }

  async function deleteDriveDoc(fileId) {
    if (!accessToken) return false;
    try {
      const res = await fetch(`${DRIVE_API}/files/${fileId}`, {
        method: 'DELETE',
        headers: { Authorization: `Bearer ${accessToken}` }
      });
      if (!res.ok && res.status !== 404) throw new Error('ลบไฟล์ไม่สำเร็จ');

      if (currentDocId === fileId) {
        currentDocId = null;
        currentDocLink = null;
        try { localStorage.removeItem(KEY_CURRENT_DOC); } catch {}
      }
      emit('eoffice:cloud:deleted', { docId: fileId });
      if (ed) ed.say('🗑️ ลบไฟล์ออกจาก Google Drive แล้ว');
      await listDriveDocs();
      return true;
    } catch (err) {
      emit('eoffice:cloud:error', { message: 'ลบไฟล์ไม่สำเร็จ: ' + err.message });
      return false;
    }
  }

  // ================= Firebase Provider Implementation =================
  function getFirebaseConfig() {
    try {
      const raw = localStorage.getItem(KEY_FIREBASE_CFG);
      return raw ? JSON.parse(raw) : null;
    } catch { return null; }
  }

  async function initFirebase() {
    if (fbInitialized && auth) return true;
    const cfg = getFirebaseConfig();
    if (!cfg || !cfg.apiKey || !cfg.projectId) return false;

    await loadScript(`${FIREBASE_CDN}/firebase-app-compat.js`);
    await loadScript(`${FIREBASE_CDN}/firebase-auth-compat.js`);
    await loadScript(`${FIREBASE_CDN}/firebase-firestore-compat.js`);
    firebase = window.firebase;

    if (!firebase.apps.length) firebase.initializeApp(cfg);
    auth = firebase.auth();
    db = firebase.firestore();

    auth.onAuthStateChanged(u => {
      user = u ? { uid: u.uid, name: u.displayName, email: u.email, photo: u.photoURL } : null;
      if (u) {
        try { currentDocId = localStorage.getItem(KEY_CURRENT_DOC) || null; } catch {}
        if (autoSaveEnabled && currentDocId) setupAutoSave();
      } else {
        currentDocId = null;
        teardownAutoSave();
      }
      emit('eoffice:cloud:ready', { user });
    });

    fbInitialized = true;
    return true;
  }

  // ================= Unified Public Operations =================
  async function init() {
    if (provider === 'gdrive') {
      return initGIS();
    } else {
      return initFirebase();
    }
  }

  async function signIn() {
    if (provider === 'gdrive') {
      const clientId = getGdriveClientId();
      if (!clientId) {
        emit('eoffice:cloud:need-config', { provider: 'gdrive' });
        if (ed) ed.say('⚙️ กรุณาใส่ Google Client ID ก่อนเข้าสู่ระบบ');
        return false;
      }
      if (!tokenClient) await initGIS();
      if (tokenClient) {
        tokenClient.requestAccessToken({ prompt: 'select_account' });
        return true;
      }
      return false;
    } else {
      if (!fbInitialized && !(await initFirebase())) {
        emit('eoffice:cloud:need-config', { provider: 'firebase' });
        if (ed) ed.say('⚙️ กรุณาตั้งค่า Firebase Config ก่อน');
        return false;
      }
      const p = new firebase.auth.GoogleAuthProvider();
      try {
        const cred = await auth.signInWithPopup(p);
        user = { uid: cred.user.uid, name: cred.user.displayName, email: cred.user.email, photo: cred.user.photoURL };
        emit('eoffice:cloud:ready', { user });
        return true;
      } catch (err) {
        emit('eoffice:cloud:error', { message: 'ล็อกอินไม่สำเร็จ: ' + err.message });
        return false;
      }
    }
  }

  async function signOut() {
    teardownAutoSave();
    currentDocId = null;
    currentDocLink = null;
    user = null;
    try { localStorage.removeItem(KEY_CURRENT_DOC); } catch {}

    if (provider === 'gdrive') {
      if (accessToken && window.google?.accounts?.oauth2) {
        try { window.google.accounts.oauth2.revoke(accessToken, () => {}); } catch {}
      }
      accessToken = null;
      tokenExpiry = 0;
      try {
        sessionStorage.removeItem(KEY_SESSION_TOKEN);
        sessionStorage.removeItem(KEY_SESSION_EXPIRY);
      } catch {}
    } else {
      if (auth) await auth.signOut();
    }

    emit('eoffice:cloud:signout', {});
    if (ed) ed.say('ออกจากระบบ Cloud เรียบร้อยแล้ว');
  }

  async function saveToCloud(options = {}) {
    if (provider === 'gdrive') {
      return saveToDrive(options);
    } else {
      if (!user) { emit('eoffice:cloud:error', { message: 'กรุณาล็อกอินก่อนบันทึก' }); return null; }
      const json = serializeDoc();
      if (!json) return null;
      const now = firebase.firestore.FieldValue.serverTimestamp();
      const name = ed.docName() || 'เอกสารไม่มีชื่อ';
      const docData = { uid: user.uid, ownerName: user.name, name, content: json, updatedAt: now };
      let docId = options.docId || currentDocId;
      if (docId) {
        await db.collection(FIREBASE_COLLECTION).doc(docId).set(docData, { merge: true });
      } else {
        docData.createdAt = now;
        const newDoc = await db.collection(FIREBASE_COLLECTION).add(docData);
        docId = newDoc.id;
      }
      currentDocId = docId;
      lastSavedJson = json;
      try { localStorage.setItem(KEY_CURRENT_DOC, docId); } catch {}
      emit('eoffice:cloud:saved', { docId, name });
      if (ed) ed.say(`☁️ บันทึกลง Firebase แล้ว — ${name}`);
      return docId;
    }
  }

  async function loadFromCloud(docId) {
    if (provider === 'gdrive') {
      return loadFromDrive(docId);
    } else {
      if (!fbInitialized && !(await initFirebase())) return null;
      emit('eoffice:cloud:loading', {});
      const snap = await db.collection(FIREBASE_COLLECTION).doc(docId).get();
      if (!snap.exists) { emit('eoffice:cloud:error', { message: 'ไม่พบเอกสารนี้' }); return null; }
      const data = snap.data();
      let doc;
      try { doc = JSON.parse(data.content); } catch { doc = null; }
      if (!doc) return null;
      deserializeDoc(doc);
      currentDocId = docId;
      lastSavedJson = data.content;
      try { localStorage.setItem(KEY_CURRENT_DOC, docId); } catch {}
      if (autoSaveEnabled) setupAutoSave();
      emit('eoffice:cloud:loaded', { docId, name: data.name });
      return doc;
    }
  }

  async function listMyDocs() {
    if (provider === 'gdrive') {
      return listDriveDocs();
    } else {
      if (!user || !db) return [];
      const snap = await db.collection(FIREBASE_COLLECTION)
        .where('uid', '==', user.uid)
        .orderBy('updatedAt', 'desc')
        .get();
      const docs = snap.docs.map(d => {
        const item = d.data();
        return {
          id: d.id,
          name: item.name,
          updatedAt: item.updatedAt ? item.updatedAt.toDate() : new Date(),
          size: (item.content || '').length
        };
      });
      emit('eoffice:cloud:list', { docs });
      return docs;
    }
  }

  async function deleteDoc(docId) {
    if (provider === 'gdrive') {
      return deleteDriveDoc(docId);
    } else {
      if (!db) return false;
      await db.collection(FIREBASE_COLLECTION).doc(docId).delete();
      if (currentDocId === docId) {
        currentDocId = null;
        try { localStorage.removeItem(KEY_CURRENT_DOC); } catch {}
      }
      emit('eoffice:cloud:deleted', { docId });
      await listMyDocs();
      return true;
    }
  }

  async function newCloudDoc() {
    currentDocId = null;
    currentDocLink = null;
    lastSavedJson = '';
    try { localStorage.removeItem(KEY_CURRENT_DOC); } catch {}
    emit('eoffice:cloud:new', {});
    if (ed) ed.say('📄 เริ่มเอกสารใหม่แล้ว (ยังไม่บันทึก)');
  }

  // ================= Auto-save (Debounce 3s) =================
  function scheduleCloudSave() {
    if (!autoSaveEnabled || !currentDocId || !user || isSaving) return;
    if (autoSaveTimer) clearTimeout(autoSaveTimer);
    autoSaveTimer = setTimeout(async () => {
      const json = serializeDoc();
      if (json && json !== lastSavedJson) await saveToCloud();
    }, AUTOSAVE_DELAY);
  }

  function setupAutoSave() {
    teardownAutoSave();
    if (!ed || !ed.quill) return;
    ed.quill.on('text-change', scheduleCloudSave);
  }

  function teardownAutoSave() {
    if (autoSaveTimer) { clearTimeout(autoSaveTimer); autoSaveTimer = null; }
    if (ed?.quill) ed.quill.off('text-change', scheduleCloudSave);
  }

  function setAutoSave(enabled) {
    autoSaveEnabled = !!enabled;
    try { localStorage.setItem(KEY_AUTOSAVE, enabled ? '1' : '0'); } catch {}
    if (enabled && currentDocId) setupAutoSave(); else teardownAutoSave();
    emit('eoffice:cloud:autosave', { enabled: autoSaveEnabled });
  }

  // ================= URL Parameter: ?driveDoc=ID or ?doc=ID =================
  function checkUrlForDoc() {
    const params = new URLSearchParams(location.search);
    const driveDoc = params.get('driveDoc');
    const docId = params.get('doc');
    const target = driveDoc || docId;
    if (!target) return;

    const doLoad = async () => {
      if (provider === 'gdrive' && accessToken) {
        await loadFromDrive(target);
      } else if (provider === 'firebase') {
        if (!fbInitialized) await initFirebase();
        if (db) await loadFromCloud(target);
      }
    };

    if (ed) doLoad();
    else document.addEventListener('eoffice:ready', () => { ed = window.eofficeEditor; doLoad(); }, { once: true });
  }

  // ================= Public API =================
  window.eofficeCloud = {
    // Provider
    getProvider: () => provider,
    setProvider(p) {
      provider = p === 'firebase' ? 'firebase' : 'gdrive';
      try { localStorage.setItem(KEY_PROVIDER, provider); } catch {}
      emit('eoffice:cloud:provider-changed', { provider });
    },

    // Init & Config
    init,
    isConfigured: () => {
      return provider === 'gdrive' ? !!getGdriveClientId() : !!getFirebaseConfig();
    },
    getGdriveClientId,
    setGdriveClientId,
    configure(cfg) {
      if (provider === 'gdrive') {
        if (typeof cfg === 'string') setGdriveClientId(cfg);
        else if (cfg?.clientId) setGdriveClientId(cfg.clientId);
      } else {
        try { localStorage.setItem(KEY_FIREBASE_CFG, JSON.stringify(cfg)); } catch {}
        fbInitialized = false;
      }
    },
    getConfig: () => {
      return provider === 'gdrive' ? { clientId: getGdriveClientId() } : getFirebaseConfig();
    },

    // Auth
    signIn,
    signOut,
    getUser: () => user,

    // Document CRUD
    serializeDoc,
    deserializeDoc,
    saveToCloud,
    loadFromCloud,
    listMyDocs,
    deleteDoc,
    newCloudDoc,

    // Folder info (Google Drive)
    getFolderInfo: () => ({
      name: FOLDER_NAME,
      id: folderId,
      link: folderLink || (folderId ? `https://drive.google.com/drive/folders/${folderId}` : null)
    }),
    getShareLink: () => currentDocLink || (currentDocId ? `https://drive.google.com/file/d/${currentDocId}/view` : null),

    // Auto-save
    scheduleCloudSave,
    setupAutoSave,
    setAutoSave,
    getAutoSave: () => autoSaveEnabled,

    // State
    getCloudDocId: () => currentDocId,
    getCurrentDocLink: () => currentDocLink,
    isSaving: () => isSaving,
  };

  // ================= Auto-init =================
  const onReady = () => {
    ed = window.eofficeEditor;
    try { autoSaveEnabled = localStorage.getItem(KEY_AUTOSAVE) === '1'; } catch {}
    init().then(() => checkUrlForDoc()).catch(() => checkUrlForDoc());
  };

  if (window.eofficeEditor) onReady();
  else document.addEventListener('eoffice:ready', onReady, { once: true });
})();
