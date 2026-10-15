// Ambil email laporan sales BSP terbaru dari Outlook (Microsoft Graph), lalu ganti data BSP di Firestore.
// Mode: run (default) | force (upload walau tanggal sama) | bootstrap (login Microsoft pertama kali)
// Env opsional: LOCAL_XLSX=<path> memakai file lokal (tanpa Graph), DRY=1 hanya cek tanpa menulis.
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const XLSX = require('xlsx');

const MODE = (process.argv[2] || 'run').toLowerCase();
const DRY = process.env.DRY === '1';
const LOCAL_XLSX = process.env.LOCAL_XLSX || '';

const SENDER = 'samks@binasanprima.com';
const SUBJECT_RE = /SALES DAN STOK BARANG/i;
const ATTACH_RE = /update\s*sales.*\.xlsx$/i;
const CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e'; // Microsoft Graph Command Line Tools (publik)
const AUTH = 'https://login.microsoftonline.com/organizations/oauth2/v2.0';
const SCOPE = 'https://graph.microsoft.com/Mail.Read offline_access';
const TOKEN_FILE = path.join(__dirname, '.token', 'rt.json');

const FS_BASE = 'https://firestore.googleapis.com/v1/projects/mds-sulawesi/databases/(default)/documents/insentifData/BSP';
const CHUNK = 1000;
const MONTHS = { JAN:1, FEB:2, MAR:3, APR:4, MEI:5, JUN:6, JUL:7, AGU:8, AGT:8, SEP:9, OKT:10, NOV:11, DES:12 };

const sleep = ms => new Promise(r => setTimeout(r, ms));
function log(...a){ console.log(...a); }
function summary(line){
  if(process.env.GITHUB_STEP_SUMMARY) fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, line + '\n');
}

async function form(url, body){
  const r = await fetch(url, { method:'POST', headers:{ 'Content-Type':'application/x-www-form-urlencoded' }, body:new URLSearchParams(body) });
  return { status: r.status, json: await r.json() };
}

// Repo ini publik: token Outlook selalu disimpan terenkripsi (AES-256-GCM) dengan kunci dari secret TOKEN_KEY.
function tokenKey(){
  const k = process.env.TOKEN_KEY;
  if(!k || k.length < 12) throw new Error('Secret TOKEN_KEY belum diisi (minimal 12 karakter). Tambahkan di Settings > Secrets and variables > Actions.');
  return crypto.scryptSync(k, 'insentif-bsp-auto', 32);
}

function saveToken(rt){
  fs.mkdirSync(path.dirname(TOKEN_FILE), { recursive:true });
  const iv = crypto.randomBytes(12);
  const c = crypto.createCipheriv('aes-256-gcm', tokenKey(), iv);
  const data = Buffer.concat([c.update(rt, 'utf8'), c.final()]);
  fs.writeFileSync(TOKEN_FILE, JSON.stringify({ v:1, iv: iv.toString('base64'), tag: c.getAuthTag().toString('base64'), data: data.toString('base64') }));
}

function loadToken(){
  const o = JSON.parse(fs.readFileSync(TOKEN_FILE, 'utf8'));
  const d = crypto.createDecipheriv('aes-256-gcm', tokenKey(), Buffer.from(o.iv, 'base64'));
  d.setAuthTag(Buffer.from(o.tag, 'base64'));
  return Buffer.concat([d.update(Buffer.from(o.data, 'base64')), d.final()]).toString('utf8');
}

async function bootstrapLogin(){
  const dc = await form(AUTH + '/devicecode', { client_id: CLIENT_ID, scope: SCOPE });
  if(!dc.json.user_code){ throw new Error('Gagal minta kode login: ' + JSON.stringify(dc.json)); }
  const msg = `LOGIN MICROSOFT: buka ${dc.json.verification_uri} lalu masukkan kode ${dc.json.user_code} (berlaku ${Math.round(dc.json.expires_in/60)} menit)`;
  log('\n' + '='.repeat(70) + '\n' + msg + '\n' + '='.repeat(70) + '\n');
  summary('## Login Microsoft diperlukan\n\n' + `1. Buka **${dc.json.verification_uri}**\n2. Masukkan kode **${dc.json.user_code}**\n3. Login dengan akun Outlook Nutrifood dan setujui izin **Baca email (Mail.Read)**\n\nKode berlaku ${Math.round(dc.json.expires_in/60)} menit.`);
  console.log(`::notice title=Login Microsoft::Buka ${dc.json.verification_uri} dan masukkan kode ${dc.json.user_code}`);
  const deadline = Date.now() + dc.json.expires_in * 1000;
  let interval = (dc.json.interval || 5) * 1000;
  while(Date.now() < deadline){
    await sleep(interval);
    const t = await form(AUTH + '/token', { grant_type:'urn:ietf:params:oauth:grant-type:device_code', client_id: CLIENT_ID, device_code: dc.json.device_code });
    if(t.json.access_token){
      saveToken(t.json.refresh_token);
      log('Login berhasil, token tersimpan.');
      return t.json.access_token;
    }
    const err = t.json.error;
    if(err === 'authorization_pending') continue;
    if(err === 'slow_down'){ interval += 5000; continue; }
    throw new Error('Login ditolak: ' + (t.json.error_description || err));
  }
  throw new Error('Kode login kedaluwarsa sebelum dipakai. Jalankan ulang mode bootstrap.');
}

async function getAccessToken(){
  if(fs.existsSync(TOKEN_FILE)){
    let refresh_token;
    try{ refresh_token = loadToken(); }
    catch(e){ throw new Error('Token tersimpan tidak bisa dibuka (' + e.message + '). Jalankan workflow mode "bootstrap" untuk login ulang.'); }
    const t = await form(AUTH + '/token', { grant_type:'refresh_token', client_id: CLIENT_ID, refresh_token, scope: SCOPE });
    if(t.json.access_token){
      if(t.json.refresh_token) saveToken(t.json.refresh_token);
      return t.json.access_token;
    }
    throw new Error('Token Outlook ditolak/kedaluwarsa: ' + (t.json.error_description || t.json.error) + ' -> jalankan workflow mode "bootstrap" untuk login ulang.');
  }
  if(MODE === 'bootstrap') return bootstrapLogin();
  throw new Error('Belum ada token Outlook. Jalankan workflow ini sekali dengan mode "bootstrap" lalu login.');
}

async function graph(token, url){
  const r = await fetch(url.startsWith('http') ? url : 'https://graph.microsoft.com/v1.0' + url, { headers:{ Authorization:'Bearer ' + token } });
  if(!r.ok) throw new Error('Graph ' + r.status + ': ' + (await r.text()).slice(0, 300));
  return r;
}

function subjectDate(subject){
  const m = String(subject).match(/(\d{1,2})\s*([A-Za-z]{3})[A-Za-z]*\s*(\d{4})/);
  if(!m) return null;
  const mon = MONTHS[m[2].toUpperCase()];
  if(!mon) return null;
  return `${m[3]}-${String(mon).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`;
}

async function cloudState(){
  const s = await (await fetch(FS_BASE)).json();
  if(!s.fields) return { rows:0, max:'' };
  const n = Number(s.fields.chunkCount.integerValue);
  let max = '', rows = 0;
  for(let i = 0; i < n; i++){
    const r = await (await fetch(`${FS_BASE}/chunks/${String(i).padStart(5,'0')}`)).json();
    for(const x of (r.fields.rows.arrayValue.values || [])){
      rows++;
      const f = x.mapValue.fields.Tgl || {};
      const t = f.stringValue || f.timestampValue || '';
      if(t > max) max = t;
    }
  }
  return { rows, max: String(max).slice(0, 10) };
}

function cleanName(raw){
  if(!raw) return raw;
  const m = String(raw).match(/\(([^)]+)\)\s*$/);
  return m ? m[1].trim() : String(raw).trim();
}

function parseSales(buffer){
  const wb = XLSX.read(buffer, { type:'buffer', cellDates:true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const all = XLSX.utils.sheet_to_json(ws, { defval:null, range:2 });
  const rows = all.filter(r => {
    const sg = String(r['Sub Group Barang'] || '').trim().toUpperCase();
    return sg === 'HILO TRADISIONAL' || sg === 'NS TRADITIONAL';
  }).map(r => ({
    'Tgl': r['Tgl Faktur'],
    'Nomor Faktur': r['Nomor Faktur'],
    'Nama Toko': cleanName(r['Outlet'] || r['Nama Customer']),
    'Nama Barang': r['Nama Barang'],
    'Sub Group Barang': r['Sub Group Barang'],
    'Salesman': r['Salesman'],
    'Netto': r['Netto']
  }));
  return { total: all.length, rows };
}

function enc(v){
  if(v === null || v === undefined) return { nullValue: null };
  if(v instanceof Date) return { timestampValue: v.toISOString() };
  if(typeof v === 'boolean') return { booleanValue: v };
  if(typeof v === 'number') return Number.isInteger(v) ? { integerValue: String(v) } : { doubleValue: v };
  return { stringValue: String(v) };
}

async function fsWrite(docPath, fields){
  const r = await fetch(`${FS_BASE}${docPath}`, { method:'PATCH', headers:{ 'Content-Type':'application/json' }, body: JSON.stringify({ fields }) });
  if(!r.ok) throw new Error(`Firestore ${r.status} di ${docPath}: ${(await r.text()).slice(0, 300)}`);
}

async function upload(rows){
  const chunks = [];
  for(let i = 0; i < rows.length; i += CHUNK) chunks.push(rows.slice(i, i + CHUNK));
  const now = new Date().toISOString();
  for(let idx = 0; idx < chunks.length; idx++){
    const values = chunks[idx].map(row => ({ mapValue:{ fields: Object.fromEntries(Object.entries(row).map(([k, v]) => [k, enc(v)])) } }));
    await fsWrite(`/chunks/${String(idx).padStart(5,'0')}`, { rows:{ arrayValue:{ values } }, updatedAt:{ timestampValue: now } });
  }
  const list = await (await fetch(`${FS_BASE}/chunks?pageSize=300`)).json();
  for(const d of (list.documents || [])){
    const idx = parseInt(d.name.split('/').pop(), 10);
    if(!isNaN(idx) && idx >= chunks.length){
      const r = await fetch(`https://firestore.googleapis.com/v1/${d.name}`, { method:'DELETE' });
      if(!r.ok) throw new Error('Gagal hapus chunk sisa ' + d.name);
    }
  }
  await fsWrite('', { chunkCount:{ integerValue: String(chunks.length) }, rowCount:{ integerValue: String(rows.length) }, updatedAt:{ timestampValue: now } });
  return chunks.length;
}

async function main(){
  log(`Mode: ${MODE}${DRY ? ' (DRY RUN)' : ''}`);
  const cloud = await cloudState();
  log(`Cloud BSP sekarang: ${cloud.rows} baris, tanggal terakhir ${cloud.max || '-'}`);

  let buffer, subject = '(file lokal)', reportDate = null;
  if(LOCAL_XLSX){
    buffer = fs.readFileSync(LOCAL_XLSX);
  } else {
    const token = await getAccessToken();
    if(MODE === 'bootstrap' && !process.env.GITHUB_ACTIONS) log('Bootstrap selesai.');
    const since = new Date(Date.now() - 4 * 86400000).toISOString();
    const q = `/me/messages?$filter=${encodeURIComponent(`receivedDateTime ge ${since} and from/emailAddress/address eq '${SENDER}'`)}&$orderby=receivedDateTime desc&$top=15&$select=id,subject,receivedDateTime,hasAttachments`;
    const list = (await (await graph(token, q)).json()).value || [];
    const mail = list.find(m => m.hasAttachments && SUBJECT_RE.test(m.subject));
    if(!mail){
      log('Belum ada email laporan sales dari BSP dalam 4 hari terakhir. Coba lagi nanti.');
      summary('Belum ada email laporan sales baru.');
      return;
    }
    subject = mail.subject;
    reportDate = subjectDate(mail.subject);
    log(`Email laporan terbaru diterima ${mail.receivedDateTime} (tanggal laporan: ${reportDate || 'tidak terbaca'})`);
    if(MODE !== 'force' && reportDate && cloud.max && cloud.max >= reportDate){
      log(`Data cloud (${cloud.max}) sudah mencakup laporan ${reportDate}. Tidak ada yang perlu diupload.`);
      summary(`Sudah terbaru: cloud ${cloud.max}, laporan ${reportDate}.`);
      return;
    }
    const atts = (await (await graph(token, `/me/messages/${mail.id}/attachments?$select=id,name,size`)).json()).value || [];
    const att = atts.find(a => ATTACH_RE.test(a.name));
    if(!att) throw new Error('Lampiran "update sales ... .xlsx" tidak ditemukan. Lampiran ada: ' + atts.map(a => a.name).join(', '));
    log(`Mengunduh lampiran: ${att.name} (${att.size} byte)`);
    buffer = Buffer.from(await (await graph(token, `/me/messages/${mail.id}/attachments/${att.id}/$value`)).arrayBuffer());
  }

  const { total, rows } = parseSales(buffer);
  const dates = rows.map(r => String(r['Tgl'] instanceof Date ? r['Tgl'].toISOString() : r['Tgl']).slice(0, 10));
  const bad = dates.filter(d => !/^\d{4}-\d{2}-\d{2}$/.test(d)).length;
  const sorted = dates.filter(d => /^\d{4}-\d{2}-\d{2}$/.test(d)).sort();
  log(`File: ${total} baris total, ${rows.length} baris Hilo/NS Tradisional, tanggal ${sorted[0]} s/d ${sorted[sorted.length - 1]}, tanggal tidak valid ${bad}`);

  if(!rows.length) throw new Error('Tidak ada baris tradisional di file. Data cloud TIDAK diganti.');
  if(bad) throw new Error(`${bad} baris punya tanggal tidak valid. Data cloud TIDAK diganti.`);
  const newMax = sorted[sorted.length - 1];
  if(MODE !== 'force' && cloud.max && newMax < cloud.max) throw new Error(`File lebih lama (${newMax}) daripada cloud (${cloud.max}). Data cloud TIDAK diganti.`);

  if(DRY){ log('DRY RUN: tidak menulis ke cloud.'); return; }
  const n = await upload(rows);
  const check = await cloudState();
  log(`Upload selesai: ${n} chunk, ${rows.length} baris. Cek ulang cloud: ${check.rows} baris, tanggal terakhir ${check.max}`);
  summary(`Terupload: **${rows.length}** baris sampai **${newMax}**.`);
  if(check.rows !== rows.length) throw new Error(`Jumlah baris di cloud (${check.rows}) tidak sama dengan yang diupload (${rows.length})`);
}

main().catch(e => { console.error('GAGAL: ' + e.message); summary('GAGAL: ' + e.message); process.exit(1); });
