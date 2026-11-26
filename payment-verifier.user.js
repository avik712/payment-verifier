// ==UserScript==
// @name         Payment Verifier
// @namespace    https://github.com/avik712/payment-verifier
// @version      2.0.0
// @description  Gmail add-on: paste a payment screenshot -> saved to Google Sheet. Admin uploads the bank statement (Excel/CSV/PDF, password supported) -> rows marked Verified (green) / Not Verified (red).
// @match        https://mail.google.com/*
// @run-at       document-idle
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @connect      cdnjs.cloudflare.com
// @require      https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js
// @updateURL    https://raw.githubusercontent.com/avik712/payment-verifier/main/payment-verifier.user.js
// @downloadURL  https://raw.githubusercontent.com/avik712/payment-verifier/main/payment-verifier.user.js
// ==/UserScript==

(function () {
  'use strict';
  if (window.top !== window.self) return;

  const SHOW_FLOATING_BUTTON = true;   // false = open only from the Tampermonkey menu
  const DATE_TOLERANCE_DAYS = 0;       // 0 = bank date must equal payment date exactly; 1 = allow +/- 1 day

  // ---- CORE-START (pure logic, no browser APIs) ----
  const p2 = (n) => String(n).padStart(2, '0');
  const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  const cleanId = (v) => String(v == null ? '' : v).replace(/[^A-Za-z0-9]/g, '').toUpperCase();
  const toAmount = (v) => { const n = parseFloat(String(v == null ? '' : v).replace(/[^\d.\-]/g, '')); return isNaN(n) ? null : n; };
  const isoDate = (y, m, d) => {
    y = +y; m = +m; d = +d; if (y < 100) y += 2000;
    if (m < 1 || m > 12 || d < 1 || d > 31 || y < 2000 || y > 2100) return '';
    return `${y}-${p2(m)}-${p2(d)}`;
  };
  const dayDiff = (a, b) => Math.abs(Date.parse(a + 'T00:00:00Z') - Date.parse(b + 'T00:00:00Z')) / 864e5;

  // returns {dates:[YYYY-MM-DD], rest: text with the dates removed}
  function extractDates(text) {
    const dates = []; let rest = String(text);
    const mon = (s) => MONTHS[s.slice(0, 3).toLowerCase()];
    const rules = [
      [/\b(\d{4})-(\d{1,2})-(\d{1,2})\b/g, (a) => isoDate(a[1], a[2], a[3])],
      [/\b(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4}|\d{2})\b/g, (a) => isoDate(a[3], a[2], a[1])],
      [/\b(\d{1,2})[\s\-\/]*([A-Za-z]{3,9})[\s\-\/,.]*(\d{4}|\d{2})\b/g, (a) => (mon(a[2]) ? isoDate(a[3], mon(a[2]), a[1]) : '')],
      [/\b([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{4})\b/g, (a) => (mon(a[1]) ? isoDate(a[3], mon(a[1]), a[2]) : '')],
    ];
    for (const [re, fn] of rules)
      rest = rest.replace(re, (...a) => { const r = fn(a); if (r) { dates.push(r); return ' '; } return a[0]; });
    return { dates, rest };
  }

  // group text lines into statement records (a record starts at a line that begins with a date)
  function buildRecords(lines) {
    const start = /^\s*(\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}|\d{4}-\d{2}-\d{2}|\d{1,2}[\s\-][A-Za-z]{3,9}[\s\-,]*\d{2,4})/;
    const recs = []; let cur = null;
    for (const l of lines) {
      if (start.test(l)) { if (cur) recs.push(cur); cur = l; } else if (cur) cur += ' ' + l;
    }
    if (cur) recs.push(cur);
    if (recs.length < 3) return lines.map((_, i) => lines.slice(Math.max(0, i - 1), i + 2).join(' '));
    return recs;
  }

  function prepareRecords(texts) {
    return texts.filter((t) => t && t.trim()).map((text) => {
      const upper = text.toUpperCase(), { dates, rest } = extractDates(text);
      const nums = (rest.replace(/,/g, '').replace(/\d{9,}/g, ' ').match(/\d+(?:\.\d{1,2})?/g) || []).map(parseFloat);
      return { upper, clean: upper.replace(/[^A-Z0-9]/g, ''), dates, nums };
    });
  }

  // one sheet row (utr, amount, date) vs. the statement -> {status, remark}
  function matchOne(sub, recs, tol) {
    const utr = cleanId(sub.utr), amt = Number(sub.amount), d = sub.date;
    if (utr.length < 9) return { status: 'Not Verified', remark: 'UTR missing/too short' };
    let best = null;
    for (const r of recs) {
      if (!(r.upper.includes(utr) || r.clean.includes(utr))) continue;
      const amtOk = r.nums.some((n) => Math.abs(n - amt) < 0.01);
      const dateOk = !!d && r.dates.some((x) => dayDiff(x, d) <= tol);
      if (amtOk && dateOk) return { status: 'Verified', remark: 'UTR, amount and date matched' };
      const score = (amtOk ? 1 : 0) + (dateOk ? 1 : 0);
      if (!best || score > best.score) best = { score, amtOk, dateOk, r };
    }
    if (!best) return { status: 'Not Verified', remark: 'UTR not found in bank statement' };
    const why = [];
    if (!best.amtOk) why.push(`amount does not match (sheet ${amt})`);
    if (!best.dateOk) why.push(`date does not match (sheet ${d}; bank ${best.r.dates.join(', ') || 'none found'})`);
    return { status: 'Not Verified', remark: 'UTR found but ' + why.join(' and ') };
  }

  // ---- OCR text -> fields (tuned for PhonePe / GPay / Paytm / bank-app screenshots) ----
  function findAmount(raw) {
    const flat = raw.replace(/,/g, '');
    let m = flat.match(/(?:[₹%]|Rs\.?|INR)\s?(\d{1,8}(?:\.\d{1,2})?)/);
    if (m) return m[1];
    // symbol glued to number but misread (e.g. "€499"): the amount usually appears more than once
    const loose = (re, minCount) => {
      const c = {};
      (flat.match(re) || []).forEach((tok) => { const n = tok.replace(/^\D+/, ''); c[n] = (c[n] || 0) + 1; });
      const top = Object.entries(c).sort((a, b) => b[1] - a[1])[0];
      return top && top[1] >= minCount ? top[0] : '';
    };
    const a1 = loose(/(?<![\w])[^\w\s]{1,2}\d{1,7}(?:\.\d{1,2})?(?![\w.])/g, 1); if (a1) return a1;
    const a2 = loose(/(?<![\w])[A-Za-z]{1,2}\d{1,7}(?:\.\d{1,2})?(?![\w.])/g, 2); if (a2) return a2;
    const lines = raw.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    const stop = lines.findIndex((l) => /completed|successful|success|paid|failed|\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}|[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4}/i.test(l));
    const head = lines.slice(0, stop > 0 ? stop : 8);
    for (const l of head) {
      if (/^to\b/i.test(l) || (l.match(/\d/g) || []).length >= 9) continue;
      const mm = l.replace(/,/g, '').match(/^[^\dA-Za-z]{0,3}[A-Za-z]?\s?(\d{1,8}(?:\.\d{1,2})?)\s*\W{0,2}$/);
      if (mm) return mm[1];
    }
    return '';
  }
  function to24(h, m, ap) {
    h = +h;
    if (ap) { ap = ap.toLowerCase(); if (ap === 'pm' && h < 12) h += 12; if (ap === 'am' && h === 12) h = 0; }
    return `${p2(h)}:${m}`;
  }
  function parseScreenshot(text) {
    const t = text.replace(/,/g, '');
    const out = { utr: '', amount: '', date: '', time: '', payer: '', last4: '', failedHint: false };
    let m = t.match(/(?:UPI\s*transaction\s*ID|UTR|UPI\s*Ref\w*\.?\s*(?:No|ID)?|Ref(?:erence)?\.?\s*(?:No|ID)?)\D{0,6}(\d{12})/i) || t.match(/\b(\d{12})\b/);
    if (m) out.utr = m[1];
    out.amount = findAmount(text);
    const mk = (d, mo, y) => (MONTHS[mo.slice(0, 3).toLowerCase()] ? `${y}-${p2(MONTHS[mo.slice(0, 3).toLowerCase()])}-${p2(d)}` : '');
    if ((m = t.match(/(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4}),?\s+(?:at\s+)?(\d{1,2}):(\d{2})\s*([AaPp][Mm])?/))) {
      out.date = mk(m[1], m[2], m[3]); out.time = to24(m[4], m[5], m[6]);
    } else if ((m = t.match(/([A-Za-z]{3,9})\s+(\d{1,2}),?\s+(\d{4}),?\s+(?:at\s+)?(\d{1,2}):(\d{2})\s*([AaPp][Mm])?/))) {
      out.date = mk(m[2], m[1], m[3]); out.time = to24(m[4], m[5], m[6]);
    } else if ((m = t.match(/(\d{1,2}):(\d{2})\s*([AaPp][Mm])?,?\s+(?:on\s+)?(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})/i))) {
      out.date = mk(m[4], m[5], m[6]); out.time = to24(m[1], m[2], m[3]);
    }
    if ((m = t.match(/(?:From|Paid by|Debited from|Sender)\s*:?\s*([A-Za-z][A-Za-z .'\-]{1,50}?)\s*(?:\(|\n|$)/i)) && !/^[Xx*•\s]+$/.test(m[1]))
      out.payer = m[1].replace(/\s+/g, ' ').trim();
    // PAYER's bank account last 4 = the masked number that comes after "Debited from" / "From" / "Paid from"
    m = t.match(/(?:Debited\s*from|Paid\s*from|Pay(?:ing)?\s*from|Sent\s*from|Source\s*account|From)[\s\S]{0,90}?[•·.*xX]{2,}\s?(\d{4})\b/i);
    if (m) out.last4 = m[1];
    out.failedHint = /\b(failed|declined|cancelled|pending)\b/i.test(t);
    return out;
  }
  // ---- CORE-END ----

  // ---- OFFCRYPTO-START (open password-protected .xlsx: Office 2010+ "Agile" encryption) ----
  async function decryptAgile(XLSX, bytes, password, subtle) {
    const cat = (...a) => { const o = new Uint8Array(a.reduce((s, x) => s + x.length, 0)); let p = 0; a.forEach((x) => { o.set(x, p); p += x.length; }); return o; };
    const u32 = (n) => new Uint8Array([n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255]);
    const b64 = (s) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));
    const hn = (a) => ({ SHA1: 'SHA-1', SHA256: 'SHA-256', SHA384: 'SHA-384', SHA512: 'SHA-512' })[a];
    const attr = (tag, n) => (tag.match(new RegExp('\\b' + n + '="([^"]*)"')) || [])[1];
    const unsupported = () => new Error('This Excel file uses an old/unsupported encryption. Remove the password in Excel (File > Info > Protect Workbook > Encrypt with Password, clear it) or upload the PDF/CSV version.');

    const cfb = XLSX.CFB.read(bytes, { type: 'array' });
    const info = XLSX.CFB.find(cfb, 'EncryptionInfo'), pkg = XLSX.CFB.find(cfb, 'EncryptedPackage');
    if (!info || !pkg) throw unsupported();
    const ib = Uint8Array.from(info.content);
    if (!((ib[0] | (ib[1] << 8)) === 4 && (ib[2] | (ib[3] << 8)) === 4)) throw unsupported();
    const xml = new TextDecoder().decode(ib.subarray(8));
    const kd = (xml.match(/<(?:\w+:)?keyData\b[^>]*>/) || [])[0], ek = (xml.match(/<(?:\w+:)?encryptedKey\b[^>]*>/) || [])[0];
    if (!kd || !ek) throw unsupported();
    const spin = +attr(ek, 'spinCount'), keyBits = +attr(ek, 'keyBits'), hashSize = +attr(ek, 'hashSize');
    const ekHash = hn(attr(ek, 'hashAlgorithm')), kdHash = hn(attr(kd, 'hashAlgorithm'));
    if (!ekHash || !kdHash || !/AES/i.test(attr(ek, 'cipherAlgorithm') || '')) throw unsupported();
    const salt = b64(attr(ek, 'saltValue')), kdSalt = b64(attr(kd, 'saltValue')), blockSize = +attr(kd, 'blockSize') || 16;

    const digest = async (name, d) => new Uint8Array(await subtle.digest(name, d));
    const aes = async (key, iv, data) => {   // AES-CBC decrypt WITHOUT padding (WebCrypto insists on padding, so add a valid pad block)
      const k = await subtle.importKey('raw', key, { name: 'AES-CBC' }, false, ['decrypt', 'encrypt']);
      const pad = new Uint8Array(await subtle.encrypt({ name: 'AES-CBC', iv: data.subarray(data.length - 16) }, k, new Uint8Array(0)));
      return new Uint8Array(await subtle.decrypt({ name: 'AES-CBC', iv }, k, cat(data, pad.subarray(0, 16))));
    };
    const pw = new Uint8Array(password.length * 2); for (let i = 0; i < password.length; i++) { pw[2 * i] = password.charCodeAt(i) & 255; pw[2 * i + 1] = password.charCodeAt(i) >> 8; }
    let H = await digest(ekHash, cat(salt, pw));
    for (let i = 0; i < spin; i++) H = await digest(ekHash, cat(u32(i), H));
    const derive = async (hex) => {
      const x = await digest(ekHash, cat(H, Uint8Array.from(hex.match(/../g).map((h) => parseInt(h, 16)))));
      const kb = keyBits / 8; return x.length >= kb ? x.subarray(0, kb) : cat(x, new Uint8Array(kb - x.length).fill(0x36));
    };
    let vIn, vHash;
    try {
      vIn = (await aes(await derive('fea7d2763b4b9e79'), salt, b64(attr(ek, 'encryptedVerifierHashInput')))).subarray(0, +attr(ek, 'saltSize'));
      vHash = (await aes(await derive('d7aa0f6d3061344e'), salt, b64(attr(ek, 'encryptedVerifierHashValue')))).subarray(0, hashSize);
    } catch (e) { return null; }
    const chk = await digest(ekHash, vIn);
    for (let i = 0; i < hashSize; i++) if (chk[i] !== vHash[i]) return null;   // wrong password
    const secret = (await aes(await derive('146e0be7abacd0d6'), salt, b64(attr(ek, 'encryptedKeyValue')))).subarray(0, keyBits / 8);

    const pb = Uint8Array.from(pkg.content);
    const size = (pb[0] | (pb[1] << 8) | (pb[2] << 16) | (pb[3] << 24)) >>> 0;
    const data = pb.subarray(8), parts = [];
    for (let i = 0, off = 0; off < data.length; i++, off += 4096) {
      const iv = (await digest(kdHash, cat(kdSalt, u32(i)))).subarray(0, blockSize);
      parts.push(await aes(secret, iv, data.subarray(off, Math.min(off + 4096, data.length))));
    }
    return cat(...parts).subarray(0, size);
  }
  // ---- OFFCRYPTO-END ----

  // ================= helpers =================
  const cfg = () => ({ url: GM_getValue('url', ''), key: GM_getValue('key', '') });
  function api(payload) {
    const c = cfg();
    if (!c.url || !c.key) return Promise.reject(new Error('Open Settings (gear icon) and enter the Web App URL and Access key first.'));
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method: 'POST', url: c.url, data: JSON.stringify({ ...payload, token: c.key }),
        headers: { 'Content-Type': 'text/plain;charset=utf-8' }, timeout: 90000,
        onload: (r) => { try { resolve(JSON.parse(r.responseText)); } catch (e) { reject(new Error('Bad reply from server. Check the Web App URL and that access is set to "Anyone".')); } },
        onerror: () => reject(new Error('Network error – check internet / Web App URL.')),
        ontimeout: () => reject(new Error('Server timed out.')),
      });
    });
  }
  const fetchText = (url) => new Promise((res, rej) => GM_xmlhttpRequest({ method: 'GET', url, onload: (r) => res(r.responseText), onerror: () => rej(new Error('Could not download ' + url)) }));
  const needPw = (wrong) => { const e = new Error(wrong ? 'Wrong password' : 'Password required'); e.needPassword = true; e.wrong = wrong; return e; };

  // ---- statement readers (return array of record texts) ----
  async function readStatement(file, pw) {
    const buf = await file.arrayBuffer();
    if (/\.pdf$/i.test(file.name) || file.type === 'application/pdf') return readPdf(buf, pw);
    const bytes = new Uint8Array(buf); let wb;
    try { wb = XLSX.read(bytes, { type: 'array' }); } catch (e) {
      if (!/password/i.test(String(e && e.message))) throw e;
      if (!pw) throw needPw(false);
      const plain = await decryptAgile(XLSX, bytes, pw, crypto.subtle);
      if (!plain) throw needPw(true);
      wb = XLSX.read(plain, { type: 'array' });
    }
    const texts = [];
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name]; if (!ws['!ref']) continue;
      const rg = XLSX.utils.decode_range(ws['!ref']);
      for (let r = rg.s.r; r <= rg.e.r; r++) {
        const cells = [];
        for (let c = rg.s.c; c <= rg.e.c; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })]; if (!cell || cell.v === '' || cell.v == null) continue;
          if (cell.t === 'n' && cell.z && XLSX.SSF.is_date(cell.z)) {
            const d = XLSX.SSF.parse_date_code(cell.v); cells.push(`${d.y}-${p2(d.m)}-${p2(d.d)}`);
          } else cells.push(String(cell.v));
        }
        if (cells.length) texts.push(cells.join(' | '));
      }
    }
    return texts;
  }
  let pdfReady = false;
  async function readPdf(buf, pw) {
    if (!pdfReady) {
      const code = await fetchText('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js');
      pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(new Blob([code], { type: 'text/javascript' }));
      pdfReady = true;
    }
    let pdf;
    try { pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buf.slice(0)), password: pw || undefined }).promise; } catch (e) {
      if (e && e.name === 'PasswordException') throw needPw(e.code === 2);
      throw e;
    }
    const lines = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      const items = (await (await pdf.getPage(p)).getTextContent()).items.filter((i) => i.str && i.str.trim());
      const byY = [];
      for (const it of items) {
        const y = it.transform[5]; let row = byY.find((r) => Math.abs(r.y - y) < 3);
        if (!row) byY.push((row = { y, parts: [] })); row.parts.push({ x: it.transform[4], s: it.str });
      }
      byY.sort((a, b) => b.y - a.y).forEach((r) => lines.push(r.parts.sort((a, b) => a.x - b.x).map((q) => q.s).join(' ')));
    }
    if (!lines.length) throw new Error('No text found in PDF (scanned PDF?). Please upload the Excel/CSV version of the statement.');
    return buildRecords(lines);
  }

  // ---- screenshot OCR ----
  async function prepImage(file) {
    const img = await createImageBitmap(file);
    const c = document.createElement('canvas'); c.width = img.width * 2; c.height = img.height * 2;
    const x = c.getContext('2d'); x.drawImage(img, 0, 0, c.width, c.height);
    const px = x.getImageData(0, 0, c.width, c.height);
    let sum = 0; for (let i = 0; i < px.data.length; i += 400) sum += px.data[i];
    if (sum / (px.data.length / 400) < 110) { // dark-mode screenshot -> invert
      for (let i = 0; i < px.data.length; i += 4) { px.data[i] = 255 - px.data[i]; px.data[i + 1] = 255 - px.data[i + 1]; px.data[i + 2] = 255 - px.data[i + 2]; }
      x.putImageData(px, 0, 0);
    }
    return c;
  }
  async function ocr(file, onStatus) {
    const worker = await Tesseract.createWorker('eng', 1, { logger: (m) => m.status && onStatus(m.status, m.progress || 0) });
    try {
      const canvas = await prepImage(file);
      let text = (await worker.recognize(canvas)).data.text, r = parseScreenshot(text);
      if (!r.amount || !r.date || !r.payer || !r.last4) {
        await worker.setParameters({ tessedit_pageseg_mode: '11' });
        text += '\n' + (await worker.recognize(canvas)).data.text;
        const r2 = parseScreenshot(text); for (const k of Object.keys(r2)) if (!r[k] && r2[k]) r[k] = r2[k];
      }
      return { text, r };
    } finally { await worker.terminate(); }
  }

  // ================= UI (Shadow DOM, styled) =================
  let policy = null;
  try { if (window.trustedTypes && trustedTypes.createPolicy) policy = trustedTypes.createPolicy('pv-ui', { createHTML: (s) => s }); } catch (e) { policy = null; }
  const html = (s) => (policy ? policy.createHTML(s) : s);
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  const CSS = `
:host{all:initial}
*{box-sizing:border-box;font-family:'Google Sans',Inter,system-ui,-apple-system,'Segoe UI',Roboto,Arial,sans-serif}
.fab{position:fixed;right:76px;bottom:24px;display:flex;align-items:center;gap:8px;padding:12px 20px;border:0;border-radius:999px;color:#fff;font-size:14px;font-weight:600;letter-spacing:.2px;cursor:pointer;
 background:linear-gradient(135deg,#4f46e5,#7c3aed 55%,#db2777);box-shadow:0 8px 24px rgba(79,70,229,.45);transition:transform .2s,box-shadow .2s}
.fab:hover{transform:translateY(-3px) scale(1.03);box-shadow:0 14px 32px rgba(124,58,237,.55)}
.ov{position:fixed;inset:0;display:none;background:rgba(15,23,42,.55);backdrop-filter:blur(6px);overflow:auto;animation:fade .2s ease}
.ov.on{display:block}
@keyframes fade{from{opacity:0}to{opacity:1}} @keyframes rise{from{opacity:0;transform:translateY(18px)}to{opacity:1;transform:none}}
.panel{max-width:820px;margin:28px auto;background:#f8fafc;border-radius:22px;overflow:hidden;color:#0f172a;font-size:14px;box-shadow:0 30px 80px rgba(0,0,0,.4);animation:rise .3s ease}
.hd{display:flex;align-items:center;justify-content:space-between;padding:20px 26px;color:#fff;background:linear-gradient(120deg,#4338ca,#7c3aed 55%,#db2777)}
.hd h1{margin:0;font-size:21px;font-weight:700}.hd p{margin:3px 0 0;font-size:12px;opacity:.85}
.ic{width:38px;height:38px;border-radius:12px;border:0;margin-left:8px;background:rgba(255,255,255,.18);color:#fff;font-size:17px;cursor:pointer;transition:background .2s}
.ic:hover{background:rgba(255,255,255,.32)}
.body{padding:8px 26px 26px}
.card{background:#fff;border:1px solid #e2e8f0;border-radius:16px;padding:20px;margin-top:18px;box-shadow:0 2px 10px rgba(15,23,42,.05)}
.card.admin{background:linear-gradient(180deg,#fff,#f5f3ff)}
.ttl{display:flex;align-items:center;gap:12px;margin:0 0 6px}
.num{width:30px;height:30px;border-radius:10px;display:grid;place-items:center;color:#fff;font-weight:700;background:linear-gradient(135deg,#4f46e5,#9333ea)}
.ttl h2{margin:0;font-size:17px}.sub{margin:0 0 4px 42px;color:#64748b;font-size:12px}
label{display:block;margin:14px 0 5px;font-size:12px;font-weight:700;color:#334155;text-transform:uppercase;letter-spacing:.5px}
label i{color:#e11d48;font-style:normal}
input[type=text],input[type=email],input[type=password],input[type=date],input[type=time]{width:100%;padding:11px 13px;font-size:14px;color:#0f172a;background:#f8fafc;border:1.5px solid #cbd5e1;border-radius:11px;outline:none;transition:border .15s,box-shadow .15s,background .15s}
input:focus{border-color:#7c3aed;background:#fff;box-shadow:0 0 0 4px rgba(124,58,237,.15)}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:0 16px}
@media(max-width:640px){.grid{grid-template-columns:1fr}}
.drop{display:flex;flex-direction:column;align-items:center;gap:4px;padding:22px;text-align:center;cursor:pointer;color:#475569;border:2px dashed #a5b4fc;border-radius:14px;background:linear-gradient(180deg,#eef2ff,#f5f3ff);transition:all .2s}
.drop:hover,.drop.over{border-color:#7c3aed;background:#ede9fe;transform:scale(1.01)}
.drop{margin:6px 0 0;text-transform:none;letter-spacing:0;font-weight:400;font-size:14px}
.drop b{font-size:15px;color:#4338ca}.drop span{font-size:12px}.big{font-size:30px}
.prev{display:none;max-width:100%;max-height:230px;margin-top:12px;border-radius:12px;border:1px solid #e2e8f0;box-shadow:0 4px 14px rgba(0,0,0,.12)}
.bar{display:none;height:6px;margin-top:12px;border-radius:99px;background:#e2e8f0;overflow:hidden}.bar i{display:block;height:100%;width:0;border-radius:99px;background:linear-gradient(90deg,#4f46e5,#db2777);transition:width .25s}
.btn{margin-top:18px;padding:12px 22px;border:0;border-radius:12px;color:#fff;font-size:14px;font-weight:700;cursor:pointer;background:linear-gradient(135deg,#4f46e5,#7c3aed);box-shadow:0 6px 16px rgba(79,70,229,.35);transition:transform .15s,box-shadow .15s,opacity .15s}
.btn:hover{transform:translateY(-2px);box-shadow:0 10px 22px rgba(79,70,229,.45)}.btn:disabled{opacity:.45;cursor:not-allowed;transform:none;box-shadow:none}
.btn.green{background:linear-gradient(135deg,#059669,#10b981);box-shadow:0 6px 16px rgba(5,150,105,.35)}
.btn.ghost{background:#e2e8f0;color:#334155;box-shadow:none}
.msg{margin-top:10px;font-size:13px;white-space:pre-line;color:#475569}.msg.err{color:#be123c}.msg.okm{color:#047857}
.lock{display:none;margin-top:14px;padding:14px;border-radius:12px;background:#fffbeb;border:1.5px solid #fcd34d}
.lock.on{display:block;animation:rise .25s ease}.lock b{color:#92400e}
.row2{display:flex;gap:10px;align-items:center}.row2 input{flex:1}.row2 .btn{margin-top:0;white-space:nowrap}
input[type=file]{font-size:13px;padding:9px;width:100%;background:#f8fafc;border:1.5px dashed #cbd5e1;border-radius:11px}
.tw{overflow-x:auto;margin-top:12px;border-radius:12px;border:1px solid #e2e8f0}
table{border-collapse:collapse;width:100%;font-size:12px;background:#fff}th{background:#eef2ff;color:#3730a3;text-align:left;padding:9px 10px}td{padding:8px 10px;border-top:1px solid #e2e8f0}
.chip{display:inline-block;padding:3px 11px;border-radius:99px;font-weight:700;font-size:11px}.chip.v{background:#b7e1cd;color:#065f46}.chip.n{background:#f4c7c3;color:#9f1239}
.hint{margin-top:6px;font-size:11px;color:#94a3b8}
.set{display:none}.set.on{display:block}`;

  const TPL = `
<button class="fab" id="fab"><span>💳</span> Payment Verifier</button>
<div class="ov" id="ov"><div class="panel">
 <div class="hd"><div><h1>💳 Payment Verifier</h1><p>Screenshot → Google Sheet → verified against the bank statement</p></div>
  <div><button class="ic" id="bSet" title="Settings">⚙</button><button class="ic" id="bClose" title="Close">✕</button></div></div>
 <div class="body">
  <div class="card set" id="settings"><div class="ttl"><div class="num">⚙</div><h2>Settings</h2></div><p class="sub">Saved only in this browser</p>
   <label>Web App URL (Google Apps Script)</label><input type="text" id="sUrl" placeholder="https://script.google.com/macros/s/.../exec" autocomplete="off">
   <label>Access key</label><input type="password" id="sKey" autocomplete="off">
   <button class="btn" id="bSave">Save settings</button></div>

  <div class="card"><div class="ttl"><div class="num">1</div><h2>Employee · Submit payment</h2></div><p class="sub">Paste (Ctrl/Cmd+V), drop or choose the payment screenshot</p>
   <label>Payer email <i>*</i></label><input type="email" id="fEmail" placeholder="Copy-paste or type the payer's email" autocomplete="off">
   <label>Payment screenshot <i>*</i></label>
   <label class="drop" id="drop"><input type="file" id="fFile" accept="image/*" hidden><span class="big">🖼️</span><b>Drop image here, or click to choose</b><span>You can also just press Ctrl/Cmd + V</span></label>
   <img class="prev" id="prev"><div class="bar" id="bar"><i id="barFill"></i></div><div class="msg" id="ocrMsg"></div>
   <div class="grid">
    <div><label>UTR / Transaction ID <i>*</i></label><input type="text" id="fUtr" autocomplete="off"></div>
    <div><label>Amount (₹) <i>*</i></label><input type="text" id="fAmt" autocomplete="off"></div>
    <div><label>Payment date <i>*</i></label><input type="date" id="fDate"></div>
    <div><label>Payment time</label><input type="time" id="fTime"></div>
    <div><label>Payer name</label><input type="text" id="fPayer" autocomplete="off"></div>
    <div><label>Payer bank account – last 4 digits <i>*</i></label><input type="text" id="fLast4" maxlength="4" inputmode="numeric" placeholder="e.g. 0229" autocomplete="off"></div>
   </div>
   <button class="btn green" id="bSubmit">💾 Save to Google Sheet</button><div class="msg" id="empMsg"></div></div>

  <div class="card admin"><div class="ttl"><div class="num">2</div><h2>Admin · Verify against bank statement</h2></div>
   <p class="sub">Only rows that are Pending or Not Verified are checked – Verified (green) rows are skipped</p>
   <label>Admin password</label><input type="password" id="aPass" autocomplete="off">
   <label>Bank statement (Excel / CSV / PDF)</label><input type="file" id="aFile" accept=".xlsx,.xls,.csv,.pdf">
   <div class="lock" id="lock"><b>🔒 This statement is password protected.</b>
    <div class="row2" id="lockRow"><input type="password" id="sPass" placeholder="Enter the statement password" autocomplete="off"><button class="btn" id="bUnlock">Unlock</button></div></div>
   <div class="msg" id="stMsg"></div>
   <button class="btn green" id="bVerify" disabled>✅ Verify all pending rows</button>
   <div class="msg" id="admMsg"></div><div id="admTable"></div></div>
 </div></div></div>`;

  const host = document.createElement('div'); host.style.cssText = 'all:initial;position:fixed;z-index:2147483647;';
  const root = host.attachShadow({ mode: 'open' });
  try { const sh = new CSSStyleSheet(); sh.replaceSync(CSS); root.adoptedStyleSheets = [sh]; }
  catch (e) { const st = document.createElement('style'); st.textContent = CSS; root.appendChild(st); }
  const wrap = document.createElement('div'); wrap.innerHTML = html(TPL); root.appendChild(wrap);
  document.documentElement.appendChild(host);
  const $ = (id) => root.getElementById(id);
  if (!SHOW_FLOATING_BUTTON) $('fab').style.display = 'none';

  const openUI = () => { $('ov').classList.add('on'); if (!cfg().url) showSettings(true); };
  const closeUI = () => $('ov').classList.remove('on');
  const showSettings = (on) => { $('settings').classList.toggle('on', on); if (on) { $('sUrl').value = cfg().url; $('sKey').value = cfg().key; } };
  const msg = (el, t, cls) => { el.textContent = t; el.className = 'msg ' + (cls || ''); };
  $('fab').onclick = openUI; $('bClose').onclick = closeUI;
  $('bSet').onclick = () => showSettings(!$('settings').classList.contains('on'));
  $('bSave').onclick = () => { GM_setValue('url', $('sUrl').value.trim()); GM_setValue('key', $('sKey').value.trim()); showSettings(false); };
  window.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeUI(); });
  GM_registerMenuCommand('Open Payment Verifier', openUI);

  // ---- employee ----
  async function handleImage(file) {
    if (!file || !file.type.startsWith('image/')) return;
    $('prev').src = URL.createObjectURL(file); $('prev').style.display = 'block';
    $('bar').style.display = 'block'; $('barFill').style.width = '5%';
    msg($('ocrMsg'), 'Reading screenshot… (first time it downloads OCR data, ~20s)');
    try {
      const { r } = await ocr(file, (s, p) => { $('barFill').style.width = Math.max(5, Math.round(p * 100)) + '%'; msg($('ocrMsg'), 'Reading screenshot: ' + s); });
      $('fUtr').value = r.utr; $('fAmt').value = r.amount; $('fDate').value = r.date; $('fTime').value = r.time;
      $('fPayer').value = r.payer; $('fLast4').value = r.last4;
      const miss = [['UTR', r.utr], ['Amount', r.amount], ['Date', r.date], ['Payer bank last 4', r.last4]].filter((x) => !x[1]).map((x) => x[0]);
      msg($('ocrMsg'), (miss.length ? 'Could not read: ' + miss.join(', ') + '. Please type them.\n' : '✓ Details read. ') +
        (r.failedHint ? '⚠ Screenshot may show a failed/pending payment.\n' : '') + 'Check them against the screenshot before saving.', miss.length ? 'err' : 'okm');
    } catch (e) { msg($('ocrMsg'), 'Could not read the screenshot (' + e.message + '). Please type the details manually.', 'err'); }
    $('bar').style.display = 'none';
  }
  $('fFile').onchange = (e) => handleImage(e.target.files[0]);
  $('drop').ondragover = (e) => { e.preventDefault(); $('drop').classList.add('over'); };
  $('drop').ondragleave = () => $('drop').classList.remove('over');
  $('drop').ondrop = (e) => { e.preventDefault(); $('drop').classList.remove('over'); handleImage(e.dataTransfer.files[0]); };
  window.addEventListener('paste', (e) => {
    if (!$('ov').classList.contains('on')) return;
    const it = [...((e.clipboardData && e.clipboardData.items) || [])].find((i) => i.type.startsWith('image/'));
    if (it) { e.preventDefault(); e.stopPropagation(); handleImage(it.getAsFile()); }
  }, true);

  $('bSubmit').onclick = async () => {
    const f = { email: $('fEmail').value.trim(), utr: cleanId($('fUtr').value), amount: toAmount($('fAmt').value), date: $('fDate').value,
      time: $('fTime').value, payer: $('fPayer').value.trim(), last4: $('fLast4').value.trim() };
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(f.email)) return msg($('empMsg'), 'Enter a valid payer email.', 'err');
    if (f.utr.length < 9 || !(f.amount > 0) || !f.date) return msg($('empMsg'), 'UTR, amount and date are required.', 'err');
    if (!/^\d{4}$/.test(f.last4)) return msg($('empMsg'), 'Payer bank account last 4 digits are mandatory (exactly 4 digits).', 'err');
    $('bSubmit').disabled = true; msg($('empMsg'), 'Saving…');
    try {
      const r = await api({ action: 'submit', ...f });
      if (!r.ok) msg($('empMsg'), r.error, 'err');
      else {
        msg($('empMsg'), `✓ Saved to row ${r.row}. Status: Pending (admin will verify).`, 'okm');
        ['fEmail', 'fUtr', 'fAmt', 'fDate', 'fTime', 'fPayer', 'fLast4'].forEach((i) => ($(i).value = ''));
        $('fFile').value = ''; $('prev').style.display = 'none'; msg($('ocrMsg'), '');
      }
    } catch (e) { msg($('empMsg'), e.message, 'err'); }
    $('bSubmit').disabled = false;
  };

  // ---- admin ----
  let stmt = null;   // prepared statement records once the file is read / unlocked
  async function loadStatement(pw) {
    const file = $('aFile').files[0]; if (!file) return;
    stmt = null; $('bVerify').disabled = true;
    msg($('stMsg'), pw ? 'Unlocking… (a few seconds)' : 'Reading statement…');
    try {
      const recs = prepareRecords(await readStatement(file, pw));
      if (!recs.length) throw new Error('Statement looks empty.');
      stmt = recs; $('lock').classList.remove('on'); $('bVerify').disabled = false;
      msg($('stMsg'), `✓ Statement ready (${recs.length} lines read).`, 'okm');
    } catch (e) {
      if (e.needPassword) {
        $('lock').classList.add('on'); $('sPass').focus();
        msg($('stMsg'), e.wrong ? '✗ Wrong password – try again.' : 'Enter the statement password to continue.', e.wrong ? 'err' : '');
      } else { $('lock').classList.remove('on'); msg($('stMsg'), e.message, 'err'); }
    }
  }
  $('aFile').onchange = () => { $('sPass').value = ''; $('admTable').textContent = ''; msg($('admMsg'), ''); loadStatement(''); };
  $('bUnlock').onclick = () => loadStatement($('sPass').value);
  $('sPass').addEventListener('keydown', (e) => { if (e.key === 'Enter') loadStatement($('sPass').value); });

  $('bVerify').onclick = async () => {
    const password = $('aPass').value;
    if (!password) return msg($('admMsg'), 'Enter the admin password.', 'err');
    if (!stmt) return msg($('admMsg'), 'Choose (and unlock) the bank statement first.', 'err');
    $('bVerify').disabled = true; $('admTable').textContent = '';
    try {
      msg($('admMsg'), 'Fetching Pending / Not Verified rows from the sheet…');
      const pend = await api({ action: 'pending', password });
      if (!pend.ok) throw new Error(pend.error);
      if (!pend.rows.length) { msg($('admMsg'), `Nothing to verify. ${pend.skipped || 0} verified row(s) skipped.`, 'okm'); $('bVerify').disabled = false; return; }
      const updates = [];
      for (let i = 0; i < pend.rows.length; i++) {            // one sheet row at a time
        const s = pend.rows[i]; msg($('admMsg'), `Checking row ${i + 1} of ${pend.rows.length}…`);
        updates.push({ row: s.row, utr: s.utr, ...matchOne(s, stmt, DATE_TOLERANCE_DAYS) });
        if (i % 50 === 0) await new Promise((r) => setTimeout(r));
      }
      for (let i = 0; i < updates.length; i += 150) {
        const r = await api({ action: 'update', password, updates: updates.slice(i, i + 150) });
        if (!r.ok) throw new Error(r.error);
      }
      const ok = updates.filter((u) => u.status === 'Verified').length;
      msg($('admMsg'), `✓ Done. ${ok} verified, ${updates.length - ok} not verified. ${pend.skipped || 0} already-verified row(s) skipped. The sheet is updated.`, 'okm');
      $('admTable').innerHTML = html('<div class="tw"><table><tr><th>Row</th><th>UTR</th><th>Status</th><th>Remark</th></tr>' + updates.map((u) =>
        `<tr><td>${u.row}</td><td>${esc(u.utr)}</td><td><span class="chip ${u.status === 'Verified' ? 'v' : 'n'}">${u.status}</span></td><td>${esc(u.remark)}</td></tr>`).join('') + '</table></div>');
    } catch (e) { msg($('admMsg'), e.message, 'err'); }
    $('bVerify').disabled = !stmt;
  };
})();
