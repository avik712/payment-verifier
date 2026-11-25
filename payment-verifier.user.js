// ==UserScript==
// @name         Payment Verifier
// @namespace    https://github.com/avik712/payment-verifier
// @version      1.0.0
// @description  Upload/paste a bank-transfer screenshot -> saved to Google Sheet. Admin uploads bank statement (Excel/CSV/PDF) -> rows marked Verified (green) / Not Verified (red).
// @match        *://*/*
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

  const SHOW_FLOATING_BUTTON = true;   // false = open only from Tampermonkey menu
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

  // ---- OCR text -> fields (tuned for GPay / PhonePe / Paytm / bank-app screenshots) ----
  function findAmount(raw) {
    let m = raw.replace(/,/g, '').match(/(?:[₹%]|Rs\.?|INR)\s?(\d{1,8}(?:\.\d{1,2})?)/);
    if (m) return m[1];
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
    } else if ((m = t.match(/(\d{1,2}):(\d{2})\s*([AaPp][Mm])?,?\s+(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{4})/))) {
      out.date = mk(m[4], m[5], m[6]); out.time = to24(m[1], m[2], m[3]);
    }
    if ((m = t.match(/(?:From|Paid by|Debited from|Sender)\s*:?\s*([A-Za-z][A-Za-z .'\-]{1,50}?)\s*(?:\(|\n|$)/i)))
      out.payer = m[1].replace(/\s+/g, ' ').trim();
    const toIdx = t.search(/\bTo:/i);
    const after = toIdx >= 0 ? t.slice(toIdx) : t;
    if ((m = after.match(/[•·.*xXeo°]{2,}\s?(\d{4})\b/)) || (m = t.match(/\bTo\s+(\d{4})\b/))) out.last4 = m[1];
    out.failedHint = /\b(failed|declined|cancelled|pending)\b/i.test(t);
    return out;
  }
  // ---- CORE-END ----

  // ================= helpers =================
  const cfg = () => ({ url: GM_getValue('url', ''), key: GM_getValue('key', '') });
  function api(payload) {
    const c = cfg();
    if (!c.url || !c.key) return Promise.reject(new Error('Open Settings and enter the Web App URL and Access key first.'));
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

  // ---- statement readers (return array of record texts) ----
  async function readStatement(file) {
    const buf = await file.arrayBuffer();
    if (/\.pdf$/i.test(file.name) || file.type === 'application/pdf') return readPdf(buf);
    const wb = XLSX.read(buf, { type: 'array' });
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
          } else if (cell.t === 'n') cells.push(String(cell.v));
          else cells.push(String(cell.v));
        }
        if (cells.length) texts.push(cells.join(' | '));
      }
    }
    return texts;
  }
  let pdfReady = false;
  async function readPdf(buf) {
    if (!pdfReady) {
      const code = await fetchText('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js');
      pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(new Blob([code], { type: 'text/javascript' }));
      pdfReady = true;
    }
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
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
    const worker = await Tesseract.createWorker('eng', 1, { logger: (m) => m.status && onStatus(m.status + (m.progress ? ' ' + Math.round(m.progress * 100) + '%' : '')) });
    try {
      const canvas = await prepImage(file);
      let text = (await worker.recognize(canvas)).data.text, r = parseScreenshot(text);
      if (!r.amount || !r.date || !r.payer) {
        await worker.setParameters({ tessedit_pageseg_mode: '11' });
        text += '\n' + (await worker.recognize(canvas)).data.text;
        const r2 = parseScreenshot(text); for (const k of Object.keys(r2)) if (!r[k] && r2[k]) r[k] = r2[k];
      }
      return { text, r };
    } finally { await worker.terminate(); }
  }

  // ================= UI (Shadow DOM) =================
  const host = document.createElement('div'); host.style.cssText = 'all:initial;position:fixed;z-index:2147483647;';
  const root = host.attachShadow({ mode: 'open' });
  root.innerHTML = `
<style>
 *{box-sizing:border-box;font-family:Arial,Helvetica,sans-serif}
 #fab{position:fixed;right:16px;bottom:16px;background:#2563eb;color:#fff;border:0;border-radius:24px;padding:10px 16px;font-size:14px;cursor:pointer;box-shadow:0 2px 8px #0005}
 #overlay{display:none;position:fixed;inset:0;background:#0007;overflow:auto}
 #panel{background:#fff;color:#222;max-width:760px;margin:20px auto;border-radius:12px;padding:18px 20px;font-size:14px}
 h2{margin:0;font-size:20px} h3{margin:0 0 8px;font-size:16px}
 .sec{border:1px solid #ddd;border-radius:10px;padding:14px;margin:14px 0} .admin{background:#f8fafc}
 label{display:block;margin-top:9px;font-weight:bold;font-size:13px}
 input[type=text],input[type=email],input[type=password],input[type=date],input[type=time]{width:100%;padding:7px;font-size:14px;border:1px solid #bbb;border-radius:6px}
 .row{display:flex;gap:10px}.row>div{flex:1}
 button{padding:8px 14px;border:0;border-radius:6px;background:#2563eb;color:#fff;font-size:14px;cursor:pointer;margin-top:10px}
 button.g{background:#16a34a} button.x{background:#6b7280} button:disabled{opacity:.5}
 #drop{border:2px dashed #94a3b8;border-radius:8px;padding:16px;text-align:center;color:#475569;margin-top:6px}
 #prev{max-width:100%;max-height:220px;margin-top:8px;display:none}
 .msg{white-space:pre-line;margin-top:8px;font-size:13px} .err{color:#b91c1c} .okm{color:#15803d}
 table{border-collapse:collapse;width:100%;margin-top:10px;font-size:12px} td,th{border:1px solid #ccc;padding:4px 6px;text-align:left}
 tr.v td.s{background:#b7e1cd;font-weight:bold} tr.n td.s{background:#f4c7c3;font-weight:bold}
 #settings{display:none}
 .top{display:flex;justify-content:space-between;align-items:center}
</style>
<button id="fab">💳 Payment Verifier</button>
<div id="overlay"><div id="panel">
 <div class="top"><h2>Payment Verifier</h2><span><button class="x" id="bSet">Settings</button> <button class="x" id="bClose">Close</button></span></div>

 <div id="settings" class="sec"><h3>Settings (saved in this browser)</h3>
  <label>Web App URL (Google Apps Script)</label><input type="text" id="sUrl" placeholder="https://script.google.com/macros/s/.../exec">
  <label>Access key</label><input type="password" id="sKey">
  <button id="bSave">Save</button></div>

 <div class="sec" id="emp"><h3>1. Employee – submit payment screenshot</h3>
  <label>Your email *</label><input type="email" id="fEmail" placeholder="you@company.com">
  <label>Payment screenshot * (choose file, drag &amp; drop, or just press Ctrl+V / Cmd+V to paste)</label>
  <div id="drop">Drop image here or <input type="file" id="fFile" accept="image/*"></div>
  <img id="prev">
  <div class="msg" id="ocrMsg"></div>
  <div class="row"><div><label>UTR / Transaction ID *</label><input type="text" id="fUtr"></div>
   <div><label>Amount (₹) *</label><input type="text" id="fAmt"></div></div>
  <div class="row"><div><label>Payment date *</label><input type="date" id="fDate"></div>
   <div><label>Payment time</label><input type="time" id="fTime"></div></div>
  <div class="row"><div><label>Payer name</label><input type="text" id="fPayer"></div>
   <div><label>Receiver account last 4</label><input type="text" id="fLast4" maxlength="4"></div></div>
  <button class="g" id="bSubmit">Save to Google Sheet</button>
  <div class="msg" id="empMsg"></div></div>

 <div class="sec admin" id="adm"><h3>2. Admin – verify against bank statement</h3>
  <label>Admin password</label><input type="password" id="aPass">
  <label>Bank statement (Excel / CSV / PDF)</label><input type="file" id="aFile" accept=".xlsx,.xls,.csv,.pdf">
  <button class="g" id="bVerify">Verify all unverified rows</button>
  <div class="msg" id="admMsg"></div><div id="admTable"></div></div>
</div></div>`;
  document.documentElement.appendChild(host);
  const $ = (id) => root.getElementById(id);
  if (!SHOW_FLOATING_BUTTON) $('fab').style.display = 'none';
  const openUI = () => { $('overlay').style.display = 'block'; if (!cfg().url) showSettings(true); };
  const closeUI = () => ($('overlay').style.display = 'none');
  const showSettings = (on) => { $('settings').style.display = on ? 'block' : 'none'; if (on) { $('sUrl').value = cfg().url; $('sKey').value = cfg().key; } };
  $('fab').onclick = openUI; $('bClose').onclick = closeUI;
  $('bSet').onclick = () => showSettings($('settings').style.display !== 'block');
  $('bSave').onclick = () => { GM_setValue('url', $('sUrl').value.trim()); GM_setValue('key', $('sKey').value.trim()); showSettings(false); };
  GM_registerMenuCommand('Open Payment Verifier', openUI);
  $('fEmail').value = GM_getValue('email', '');
  const msg = (el, t, cls) => { el.textContent = t; el.className = 'msg ' + (cls || ''); };

  // ---- employee ----
  async function handleImage(file) {
    if (!file || !file.type.startsWith('image/')) return;
    $('prev').src = URL.createObjectURL(file); $('prev').style.display = 'block';
    msg($('ocrMsg'), 'Reading screenshot (first time downloads OCR data, wait ~20s)...');
    try {
      const { r } = await ocr(file, (s) => msg($('ocrMsg'), 'Reading screenshot: ' + s));
      $('fUtr').value = r.utr; $('fAmt').value = r.amount; $('fDate').value = r.date; $('fTime').value = r.time;
      $('fPayer').value = r.payer; $('fLast4').value = r.last4;
      const miss = [['UTR', r.utr], ['Amount', r.amount], ['Date', r.date]].filter((x) => !x[1]).map((x) => x[0]);
      msg($('ocrMsg'), (miss.length ? 'Could not read: ' + miss.join(', ') + '. Please type them.\n' : 'Details read. ') +
        (r.failedHint ? 'WARNING: screenshot may show a failed/pending payment.\n' : '') + 'Check them against the screenshot before saving.', miss.length ? 'err' : 'okm');
    } catch (e) { msg($('ocrMsg'), 'OCR failed (' + e.message + '). Type the details manually. If this site blocks OCR, open the script on a different website.', 'err'); }
  }
  $('fFile').onchange = (e) => handleImage(e.target.files[0]);
  $('drop').ondragover = (e) => e.preventDefault();
  $('drop').ondrop = (e) => { e.preventDefault(); handleImage(e.dataTransfer.files[0]); };
  window.addEventListener('paste', (e) => {
    if ($('overlay').style.display !== 'block') return;
    const it = [...(e.clipboardData?.items || [])].find((i) => i.type.startsWith('image/'));
    if (it) { e.preventDefault(); handleImage(it.getAsFile()); }
  });
  $('bSubmit').onclick = async () => {
    const f = { email: $('fEmail').value.trim(), utr: cleanId($('fUtr').value), amount: toAmount($('fAmt').value), date: $('fDate').value,
      time: $('fTime').value, payer: $('fPayer').value.trim(), last4: $('fLast4').value.trim() };
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(f.email)) return msg($('empMsg'), 'Enter a valid email.', 'err');
    if (f.utr.length < 9 || !(f.amount > 0) || !f.date) return msg($('empMsg'), 'UTR, amount and date are required.', 'err');
    $('bSubmit').disabled = true; msg($('empMsg'), 'Saving...');
    try {
      GM_setValue('email', f.email);
      const r = await api({ action: 'submit', ...f });
      if (!r.ok) msg($('empMsg'), r.error, 'err');
      else { msg($('empMsg'), `Saved to row ${r.row}. Status: Pending (admin will verify).`, 'okm'); ['fUtr', 'fAmt', 'fDate', 'fTime', 'fPayer', 'fLast4'].forEach((i) => ($(i).value = '')); $('prev').style.display = 'none'; msg($('ocrMsg'), ''); }
    } catch (e) { msg($('empMsg'), e.message, 'err'); }
    $('bSubmit').disabled = false;
  };

  // ---- admin ----
  $('bVerify').onclick = async () => {
    const file = $('aFile').files[0], password = $('aPass').value;
    if (!password) return msg($('admMsg'), 'Enter the admin password.', 'err');
    if (!file) return msg($('admMsg'), 'Choose the bank statement file.', 'err');
    $('bVerify').disabled = true; $('admTable').innerHTML = '';
    try {
      msg($('admMsg'), 'Reading bank statement...');
      const recs = prepareRecords(await readStatement(file));
      if (!recs.length) throw new Error('Statement looks empty.');
      msg($('admMsg'), 'Fetching rows from the sheet...');
      const pend = await api({ action: 'pending', password });
      if (!pend.ok) throw new Error(pend.error);
      const updates = [];
      for (let i = 0; i < pend.rows.length; i++) {            // one sheet row at a time
        const s = pend.rows[i]; msg($('admMsg'), `Checking row ${i + 1} of ${pend.rows.length}...`);
        updates.push({ row: s.row, utr: s.utr, ...matchOne(s, recs, DATE_TOLERANCE_DAYS) });
        if (i % 50 === 0) await new Promise((r) => setTimeout(r));
      }
      for (let i = 0; i < updates.length; i += 150) {
        const r = await api({ action: 'update', password, updates: updates.slice(i, i + 150) });
        if (!r.ok) throw new Error(r.error);
      }
      const ok = updates.filter((u) => u.status === 'Verified').length;
      msg($('admMsg'), `Done. ${ok} verified, ${updates.length - ok} not verified. The sheet has been updated.`, 'okm');
      $('admTable').innerHTML = '<table><tr><th>Row</th><th>UTR</th><th>Remark</th><th>Status</th></tr>' + updates.map((u) =>
        `<tr class="${u.status === 'Verified' ? 'v' : 'n'}"><td>${u.row}</td><td>${u.utr}</td><td>${u.remark.replace(/</g, '&lt;')}</td><td class="s">${u.status}</td></tr>`).join('') + '</table>';
    } catch (e) { msg($('admMsg'), e.message, 'err'); }
    $('bVerify').disabled = false;
  };
})();
