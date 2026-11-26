/**
 * ONEMILL 견적·계약 서버 (구글 앱스 스크립트)
 * - 견적서 승인, 문서 열람 기록, 계약서 이메일 인증·전자서명, 서명 계약서 PDF 보관·발송
 * - 기록: 시트 "ONEMILL 견적 승인 기록"
 *     · 첫 번째 탭(시트1): 견적 승인 기록
 *     · "문서현황" 탭: 보낸 문서마다 한 줄 (관리 페이지가 발송할 때 등록, 열람·승인·서명 시 갱신)
 * - 알림: hello@dnalabs.kr 메일 + 텔레그램(스크립트 속성 TELEGRAM_TOKEN, TELEGRAM_CHAT_ID)
 */
const OWNER_EMAIL = 'hello@dnalabs.kr';
const OWNER_SIGN_URL = 'https://quote.dnalabs.kr/owner-sign.png';
const DOC_HEAD = ['문서ID', '종류', '업체', '발송 시각', '처음 열람', '마지막 열람', '열람 수', '완료 시각', '완료자', '이메일', 'PDF', '문서 지문', '받는 사람'];
// 열 번호(1부터)
const COL = { id: 1, type: 2, client: 3, sent: 4, first: 5, last: 6, views: 7, done: 8, by: 9, email: 10, pdf: 11, hash: 12, to: 13 };

function props_() { return PropertiesService.getScriptProperties(); }

function ss_() {
  const p = props_();
  const id = p.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('ONEMILL 견적 승인 기록');
  p.setProperty('SHEET_ID', ss.getId());
  ss.getSheets()[0].appendRow(['승인 시각', '견적번호', '업체', '패키지', '합계(부가세 포함)', '승인자', '연락처', '메모', '키', '견적서 링크']);
  return ss;
}
function sheet_() { return ss_().getSheets()[0]; }
function docsSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName('문서현황');
  if (!sh) {
    sh = ss.insertSheet('문서현황');
    sh.appendRow(DOC_HEAD);
    sh.setFrozenRows(1);
  }
  return sh;
}
function findDoc_(id) {
  const sh = docsSheet_();
  const v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i > 0; i--) if (String(v[i][0]) === id) return { sh: sh, row: i + 1, v: v[i] };
  return null;
}
function ts_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'); }
function iso_(d) {
  if (d && typeof d === 'string' && /^\d{4}-\d\d-\d\d \d\d:\d\d/.test(d)) return d.replace(' ', 'T') + '+09:00'; return d ? new Date(d).toISOString() : ''; }

function find_(key) {
  const values = sheet_().getDataRange().getValues();
  for (let i = values.length - 1; i > 0; i--) if (values[i][8] === key) return values[i];
  return null;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 시트에 넣을 글자: 길이 제한 + 수식으로 읽히지 않게
function clean_(v, max) {
  let s = String(v == null ? '' : v).slice(0, max || 500);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
function escHtml_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; });
}
function sha_(text) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

function telegram_(text) {
  const p = props_();
  const token = (p.getProperty('TELEGRAM_TOKEN') || '').trim(), chat = (p.getProperty('TELEGRAM_CHAT_ID') || '').trim();
  if (!token || !chat) return null;
  try {
    return UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ chat_id: chat, text: text, disable_web_page_preview: true })
    });
  } catch (err) { return null; }
}

// 텔레그램 연결 점검: …/exec?action=tgtest
function tgTest_() {
  const p = props_();
  const res = { hasToken: !!(p.getProperty('TELEGRAM_TOKEN') || '').trim(), hasChatId: !!(p.getProperty('TELEGRAM_CHAT_ID') || '').trim(), savedNames: p.getKeys() };
  if (!res.hasToken || !res.hasChatId) return res;
  const r = telegram_('[ONEMILL] 견적 승인 알림 텔레그램 연결 시험입니다.');
  if (r) {
    res.telegramStatus = r.getResponseCode();
    try { res.telegramMessage = JSON.parse(r.getContentText()).description || 'ok'; } catch (err) {}
  }
  return res;
}

// 문서 상태(고객 페이지가 열릴 때 + 열람 기록)
function view_(p) {
  const id = String(p.id || '').slice(0, 80);
  if (!id) return { ok: false, error: 'missing' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let d, first = false;
  try {
    d = findDoc_(id);
    const now = ts_();
    if (!d && p.count === '0') return { ok: true, registered: false, doneAt: '', doneBy: '', hash: '' };
    if (!d) {
      const sh = docsSheet_();
      sh.appendRow([id, clean_(p.type, 10), clean_(p.client), '', now, now, 1, '', '', '', '', '', '']);
      first = true;
      d = findDoc_(id);
    } else if (p.count !== '0') {
      const views = Number(d.v[COL.views - 1]) || 0;
      if (!d.v[COL.first - 1]) { d.sh.getRange(d.row, COL.first).setValue(now); first = true; }
      d.sh.getRange(d.row, COL.last).setValue(now);
      d.sh.getRange(d.row, COL.views).setValue(views + 1);
    }
  } finally { lock.releaseLock(); }
  if (first) telegram_('[' + (p.type === '계약서' ? '계약서' : '견적서') + ' 열람] ' + clean_(p.client || d.v[COL.client - 1]) + '\n고객이 처음으로 열어 봤습니다.');
  const v = d.v;
  return { ok: true, registered: !!v[COL.sent - 1], doneAt: iso_(v[COL.done - 1]), doneBy: v[COL.by - 1] || '', hash: v[COL.hash - 1] || '' };
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'tgtest') return out_(tgTest_());
  if (p.action === 'view') return out_(view_(p));
  if (p.action === 'setup') { docsSheet_(); return out_({ ok: true }); }
  if (p.action === 'status' && p.key) {
    const r = find_(String(p.key).slice(0, 64));
    return out_({ ok: true, approved: !!r, at: r ? new Date(r[0]).toISOString() : null, name: r ? r[5] : null });
  }
  return out_({ ok: true });
}

function approve_(d) {
  if (!d.key || !d.name) return { ok: false, error: 'missing' };
  const key = String(d.key).slice(0, 64);
  const now = new Date();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const found = find_(key);
    if (found) return { ok: true, already: true, at: new Date(found[0]).toISOString() };
    if (d.id) {
      const doc = findDoc_(String(d.id).slice(0, 80));
      if (doc && doc.v[COL.hash - 1] && d.hash && doc.v[COL.hash - 1] !== d.hash) return { ok: false, error: 'changed' };
      if (doc) {
        doc.sh.getRange(doc.row, COL.done).setValue(ts_());
        doc.sh.getRange(doc.row, COL.by).setValue(clean_(d.name, 60));
      }
    }
    sheet_().appendRow([now, clean_(d.no), clean_(d.client), clean_(d.pkg), clean_(d.total), clean_(d.name, 60),
      clean_(d.phone, 40), clean_(d.memo, 1000), key, clean_(d.url, 5000)]);
  } finally {
    lock.releaseLock();
  }
  const lines = [
    '[견적 승인] ' + clean_(d.client),
    '견적번호: ' + clean_(d.no),
    '패키지: ' + clean_(d.pkg),
    '합계: ' + clean_(d.total) + ' (부가세 포함)',
    '승인자: ' + clean_(d.name, 60) + (d.phone ? ' / ' + clean_(d.phone, 40) : ''),
    '메모: ' + (d.memo ? clean_(d.memo, 1000) : '-')
  ];
  try {
    MailApp.sendEmail(OWNER_EMAIL, '[ONEMILL] ' + clean_(d.client) + ' 견적 승인',
      lines.join('\n') + '\n\n승인한 견적서 보기:\n' + clean_(d.url, 5000), { name: 'ONEMILL 견적 알림' });
  } catch (err) {}
  telegram_(lines.join('\n') + '\n(견적서 링크는 메일과 승인 기록 시트에 있습니다)');
  return { ok: true, at: now.toISOString() };
}

// 계약서 서명 인증번호 보내기
function code_(d) {
  const id = String(d.id || '').slice(0, 80), email = String(d.email || '').trim().toLowerCase();
  if (!id || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'email' };
  const doc = findDoc_(id);
  if (!doc || !doc.v[COL.sent - 1]) return { ok: false, error: 'unregistered' };
  if (doc.v[COL.done - 1]) return { ok: false, error: 'signed' };
  const cache = CacheService.getScriptCache();
  if (cache.get('rl:' + id)) return { ok: false, error: 'wait' };
  const code = String(Math.floor(100000 + Math.random() * 900000));
  cache.put('code:' + id + ':' + email, code, 600);
  cache.put('rl:' + id, '1', 50);
  MailApp.sendEmail(email, '[디엔에이랩스] 계약서 서명 인증번호 ' + code,
    '디엔에이랩스 웹사이트 제작 용역 계약서 서명 인증번호입니다.\n\n인증번호: ' + code + '\n\n10분 안에 서명 화면에 입력해 주세요.\n본인이 요청하지 않았다면 이 메일은 무시하셔도 됩니다.',
    { name: '디엔에이랩스' });
  return { ok: true };
}

function ownerSignData_() {
  try {
    const b = UrlFetchApp.fetch(OWNER_SIGN_URL, { muteHttpExceptions: true });
    if (b.getResponseCode() === 200) return 'data:image/png;base64,' + Utilities.base64Encode(b.getBlob().getBytes());
  } catch (err) {}
  return '';
}

function folder_() {
  const p = props_();
  const id = p.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (err) {} }
  const f = DriveApp.createFolder('ONEMILL 서명 계약서');
  p.setProperty('FOLDER_ID', f.getId());
  return f;
}

const PDF_CSS = 'body{font-family:"Noto Sans KR","Malgun Gothic",sans-serif;font-size:10.5pt;line-height:1.55;color:#17181b}' +
  'h1{font-size:19pt;margin:0 0 6px}h2{font-size:13pt;margin:20px 0 8px}h3{font-size:11pt;margin:16px 0 4px}p{margin:4px 0}' +
  'table{width:100%;border-collapse:collapse;font-size:9.5pt;margin:4px 0}th,td{border:1px solid #c9cad0;padding:5px 7px;text-align:left;vertical-align:top}th{background:#f2f2f4;font-weight:normal}' +
  'ol,ul{margin:4px 0;padding-left:20px}li{margin:2px 0}.note{font-size:9pt;color:#555}.lead{margin-bottom:14px}' +
  '.signimg{height:46px;vertical-align:middle}.audit{margin-top:24px;font-size:8.5pt;color:#333}.audit td{border-color:#ddd}';

// 계약서 서명
function sign_(d) {
  const id = String(d.id || '').slice(0, 80), email = String(d.email || '').trim().toLowerCase();
  const name = clean_(d.name, 60), html = String(d.html || '');
  if (!id || !email || !name || !html || !d.sig) return { ok: false, error: 'missing' };
  if (html.length > 400000 || String(d.sig).length > 300000 || !/^data:image\/png;base64,[A-Za-z0-9+/=]+$/.test(String(d.sig))) return { ok: false, error: 'invalid' };
  const cache = CacheService.getScriptCache();
  const want = cache.get('code:' + id + ':' + email);
  if (!want || want !== String(d.code || '').trim()) return { ok: false, error: 'code' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let doc;
  try {
    doc = findDoc_(id);
    if (!doc || !doc.v[COL.sent - 1]) return { ok: false, error: 'unregistered' };
    if (doc.v[COL.done - 1]) return { ok: true, already: true, at: iso_(doc.v[COL.done - 1]) };
    const regHash = String(doc.v[COL.hash - 1] || '');
    const gotHash = sha_(html);
    if (regHash && regHash !== gotHash) return { ok: false, error: 'changed' };
    const now = new Date();
    const when = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy년 M월 d일 HH:mm');
    const dateOnly = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy년 M월 d일');
    const consents = d.consents || {};
    const mark = function (k, v) { return consents[k] === v ? '■' : '□'; };
    let body = html
      .replace(/<span class="cdate"><\/span>/g, escHtml_(dateOnly))
      .replace(/<span class="gsig"><\/span>/g, '<img class="signimg" src="' + d.sig + '">')
      .replace(/<span class="osig"><\/span>/g, '<img class="signimg" src="' + ownerSignData_() + '">')
      .replace(/<span class="cs" data-k="(\w+)" data-v="(동의|비동의)"><\/span>/g, function (m, k, v) { return mark(k, v) + ' '; });
    const audit = '<div class="audit"><h3>전자서명 기록</h3><table>' +
      '<tr><th style="width:140px">문서 번호</th><td>' + escHtml_(id) + '</td></tr>' +
      '<tr><th>서명자</th><td>' + escHtml_(name) + (d.phone ? ' / ' + escHtml_(clean_(d.phone, 40)) : '') + '</td></tr>' +
      '<tr><th>본인 확인</th><td>' + escHtml_(email) + ' (이메일 인증번호 확인)</td></tr>' +
      '<tr><th>서명 시각</th><td>' + escHtml_(when) + ' (한국 시간)</td></tr>' +
      '<tr><th>문서 지문(SHA-256)</th><td style="word-break:break-all">' + gotHash + '</td></tr></table>' +
      '<p>이 계약서는 서명자가 본인 이메일로 받은 인증번호를 입력한 뒤 전자서명한 문서입니다.</p></div>';
    const full = '<html><head><meta charset="utf-8"><style>' + PDF_CSS + '</style></head><body>' + body + audit + '</body></html>';
    const fileName = '계약서_' + clean_(d.client || doc.v[COL.client - 1], 40).replace(/[\\/:*?"<>|]/g, '') + '_서명완료_' + Utilities.formatDate(now, 'Asia/Seoul', 'yyyyMMdd') + '.pdf';
    const pdf = Utilities.newBlob(full, 'text/html', 'contract.html').getAs('application/pdf').setName(fileName);
    const file = folder_().createFile(pdf);
    doc.sh.getRange(doc.row, COL.done).setValue(ts_());
    doc.sh.getRange(doc.row, COL.by).setValue(name);
    doc.sh.getRange(doc.row, COL.email).setValue(email);
    doc.sh.getRange(doc.row, COL.pdf).setValue(file.getUrl());
    const consentText = '포트폴리오 공개: ' + (consents.portfolio || '표시 안 함') + ' / 제작사 표기: ' + (consents.credit || '표시 안 함');
    try {
      MailApp.sendEmail(email, '[디엔에이랩스] 서명 완료된 웹사이트 제작 용역 계약서',
        name + '님, 계약서 서명이 완료되었습니다.\n서명된 계약서를 첨부해 드립니다. 보관해 주세요.\n\n디엔에이랩스 최환준 드림\n010-4642-3477 / hello@dnalabs.kr',
        { name: '디엔에이랩스', attachments: [pdf], replyTo: OWNER_EMAIL });
      MailApp.sendEmail(OWNER_EMAIL, '[ONEMILL] ' + clean_(d.client) + ' 계약서 서명 완료',
        '업체: ' + clean_(d.client) + '\n서명자: ' + name + '\n이메일: ' + email + '\n서명 시각: ' + when + '\n' + consentText + '\n\n드라이브 보관: ' + file.getUrl(),
        { name: 'ONEMILL 계약 알림', attachments: [pdf] });
    } catch (err) {}
    telegram_('[계약서 서명 완료] ' + clean_(d.client) + '\n서명자: ' + name + ' (' + email + ')\n' + consentText);
    return { ok: true, at: now.toISOString() };
  } finally {
    lock.releaseLock();
  }
}

// PDF 변환 점검 (주인만 아는 SHEET_ID 필요)
function pdfTest_(p) {
  if (p.k !== props_().getProperty('SHEET_ID')) return { ok: false };
  const sig = ownerSignData_();
  const full = '<html><head><meta charset="utf-8"><style>' + PDF_CSS + '</style></head><body><h1>PDF 변환 시험</h1><p>한글 글자와 서명 그림이 보이면 정상입니다.</p><p>대표자 최환준 (인) <img class="signimg" src="' + sig + '"></p></body></html>';
  const pdf = Utilities.newBlob(full, 'text/html', 't.html').getAs('application/pdf').setName('PDF변환시험.pdf');
  const f = folder_().createFile(pdf);
  return { ok: true, bytes: pdf.getBytes().length, sigBytes: sig.length, url: f.getUrl() };
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'approve') return out_(approve_(d));
    if (d.action === 'code') return out_(code_(d));
    if (d.action === 'sign') return out_(sign_(d));
    if (d.action === 'pdftest') return out_(pdfTest_(d));
    return out_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}
