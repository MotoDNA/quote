/**
 * ONEMILL 견적 승인 알림 (구글 앱스 스크립트)
 * - 고객이 견적서 링크에서 [견적 승인하기]를 누르면 실행됩니다.
 * - 승인 기록을 구글 시트("ONEMILL 견적 승인 기록")에 남기고,
 *   hello@dnalabs.kr 로 메일을, 텔레그램이 설정돼 있으면 텔레그램으로도 알림을 보냅니다.
 * 텔레그램 알림(선택): 프로젝트 설정 > 스크립트 속성에
 *   TELEGRAM_TOKEN, TELEGRAM_CHAT_ID 를 추가하세요 (수정요청 접수 시스템과 같은 값).
 */
const OWNER_EMAIL = 'hello@dnalabs.kr';

function sheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id).getSheets()[0];
  const ss = SpreadsheetApp.create('ONEMILL 견적 승인 기록');
  props.setProperty('SHEET_ID', ss.getId());
  const sh = ss.getSheets()[0];
  sh.appendRow(['승인 시각', '견적번호', '업체', '패키지', '합계(부가세 포함)', '승인자', '연락처', '메모', '키', '견적서 링크']);
  return sh;
}

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

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'status' && p.key) {
    const r = find_(String(p.key).slice(0, 64));
    return out_({ ok: true, approved: !!r, at: r ? new Date(r[0]).toISOString() : null, name: r ? r[5] : null });
  }
  return out_({ ok: true });
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action !== 'approve') return out_({ ok: false, error: 'unknown action' });
    if (!d.key || !d.name) return out_({ ok: false, error: 'missing' });
    const key = String(d.key).slice(0, 64);
    const now = new Date();
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const found = find_(key);
      if (found) return out_({ ok: true, already: true, at: new Date(found[0]).toISOString() });
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
        lines.join('\n') + '\n\n승인한 견적서 보기:\n' + clean_(d.url, 5000));
    } catch (err) {}
    const props = PropertiesService.getScriptProperties();
    const token = props.getProperty('TELEGRAM_TOKEN'), chat = props.getProperty('TELEGRAM_CHAT_ID');
    if (token && chat) {
      try {
        UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
          method: 'post', contentType: 'application/json', muteHttpExceptions: true,
          payload: JSON.stringify({ chat_id: chat, text: lines.join('\n') + '\n(견적서 링크는 메일과 승인 기록 시트에 있습니다)' })
        });
      } catch (err) {}
    }
    return out_({ ok: true, at: now.toISOString() });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}
