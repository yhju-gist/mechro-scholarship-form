/**
 * 기계로봇공학과 박사재학생 장학금 연구실적 접수 (Google Apps Script 웹앱)
 * - 학생: 웹페이지에서 논문·특허·수상 실적과 증빙 PDF 제출
 * - 담당자: 구글 시트 한 곳에서 전체 제출을 보고, Claude가 검증·JCR 조회 결과를 채운 뒤 RAWDATA 엑셀 생성
 *
 * 설치: 1) 구글 드라이브에서 빈 스프레드시트 만들기 → [확장 프로그램 → Apps Script]
 *          → Code.gs 내용을 모두 지우고 이 파일 전체를 붙여넣기 (파일 하나로 끝)
 *       2) setup() 한 번 실행(권한 허용) → 시트 탭·드라이브 증빙 폴더 자동 생성
 *       3) 배포 → 새 배포 → 웹 앱 (실행: 나 / 액세스: 모든 사용자)
 *          학생용 주소 = 배포 주소, 외국인 학생용 = 배포 주소 끝에 ?lang=en (화면 오른쪽 위에서도 전환)
 */

const SHEETS = {
  설정: ['항목', '값', '설명'],
  학생: ['접수번호', '접수시각', '상태', '학번', '성명', '과정', '재학학기', '지도교수', '이메일', '제출계정',
         '논문수', '특허수', '수상수', '증빙폴더', '학생메모', '논문영문명'],
  논문: ['접수번호', '상태', '학번', '성명', '순번', '저널명(학생)', '논문제목', 'DOI', 'Accept일',
         '역할(학생)', '저자수(학생)', '증빙1', '증빙2',
         // ↓ Claude 검증 칸
         '역할(AI)', '저자수(AI)', '저자목록(AI)', 'AI근거', 'JCR정식명', 'JCR연도', '최상위카테고리',
         '순위', '전체', '상위%', 'IF(JCR)', '다른카테고리', 'JCR상태', '검증일', '담당자확정', '담당자메모', 'Accept일(AI)'],
  특허: ['접수번호', '상태', '학번', '성명', '순번', '구분', '명칭', '출원·등록번호', '일자', '기여도(%)', '증빙',
         'AI확인', '담당자확정', '담당자메모'],
  수상: ['접수번호', '상태', '학번', '성명', '순번', '구분', '학회·대회명', '수상명', '일자', '증빙',
         'AI확인', '담당자확정', '담당자메모'],
  JCR조회표: ['입력저널명', 'JCR정식명', 'ISSN', 'eISSN', 'JCR연도', '최상위카테고리', '순위', '전체', '상위%',
              'IF', '전체카테고리', '조회일', '비고'],
  이전인정논문: ['연도', '학번', '성명', '저널명', 'DOI', '논문제목', '비고']
};

const DEFAULT_SETTINGS = [
  ['평가명', '2026학년도 박사과정 재학생 연구우수장학금', '웹페이지 제목에 표시'],
  ['평가명(영문)', '2026 Ph.D. Student Research Excellence Scholarship', '영어 화면 제목'],
  ['접수마감', '2026-12-04 23:59', '이 시각 이후 제출 막힘 (yyyy-MM-dd HH:mm)'],
  ['실적인정 시작', '2025-09-01', 'Accept일이 이 날짜 이전이면 더블체크 표시'],
  ['실적인정 끝', '2026-08-31', 'Accept일이 이 날짜 이후면 더블체크 표시'],
  ['지도교수 목록', '고광희/Kwang Hee Ko, 김민경/Minkyung Kim, 김태성/Taeseong Kim, 김표진/Pyojin Kim, 배진현/Jinhyun Bae, 설재훈/Jae Hun Seol, 안효성/Hyo Sung Ahn, 양성/Sung Yang, 오현석/Hyunseok Oh, 이승현/Seunghyun Lee, 이재욱/Jaewook Lee, 정성호/Sungho Jeong, 지솔근/Solkeun Jee, 최성임/Seongim Choi, 허필원/Pilwon Hur, Arshad Afzal', '쉼표로 구분, "한글/영문" (기계로봇공학과 홈페이지 학과교수 2026-11-10 기준)'],
  ['담당자', '기계로봇공학과 류지나 (T.2383)', '하단 문의처'],
  ['담당자(영문)', 'Jina Ryu, Dept. of Mechanical and Robotics Engineering (ext. 2383)', '영어 화면 하단 문의처'],
  ['안내문', '논문 첫 페이지(저자·소속이 보이는 쪽)나 논문 PDF를 꼭 올려 주세요. 아직 게재 전이면 Accept 메일을 함께 올려 주세요.', '입력 화면 상단 안내'],
  ['안내문(영문)', 'Please upload the first page of each paper (showing authors and affiliations) or the full PDF. If it is not yet published, also upload the acceptance email.', '영어 화면 상단 안내'],
  ['파일최대MB', '15', '파일 1개 최대 크기'],
  ['허용 메일 도메인', 'gist.ac.kr, gm.gist.ac.kr', '본인 확인 메일을 받을 수 있는 주소 (쉼표로 구분)'],
  ['문의 메일', 'jnryu@gist.ac.kr', '인증 메일에 답장하면 이 주소로 감 (담당자 GIST 메일)']
];

/* ---------------- 초기 설정 ---------------- */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let ss = SpreadsheetApp.getActiveSpreadsheet();   // 시트에서 [확장 프로그램 → Apps Script]로 만든 경우
  if (ss) {
    props.setProperty('SHEET_ID', ss.getId());
  } else if (props.getProperty('SHEET_ID')) {
    ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  } else {
    ss = SpreadsheetApp.create('박사재학생 장학금 연구실적 접수');
    props.setProperty('SHEET_ID', ss.getId());
  }
  Object.keys(SHEETS).forEach((name, i) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = i === 0 ? ss.getSheets()[0].setName(name) : ss.insertSheet(name);
    const h = SHEETS[name];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold')
      .setBackground('#1F3864').setFontColor('#FFFFFF').setWrap(true).setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    if (name === '논문') {
      const s = h.indexOf('역할(AI)') + 1;
      sh.getRange(1, s, 1, h.length - s + 1).setBackground('#DC2715');
    }
  });
  const set = ss.getSheetByName('설정');
  const have = set.getDataRange().getValues().map(r => String(r[0]).trim());
  DEFAULT_SETTINGS.filter(r => have.indexOf(r[0]) < 0).forEach(r => set.appendRow(r));   // 빠진 항목만 추가
  set.setColumnWidth(1, 140).setColumnWidth(2, 420).setColumnWidth(3, 320);

  if (!props.getProperty('FOLDER_ID')) {
    const f = DriveApp.createFolder('박사재학생 장학금 증빙(학생 제출)');
    props.setProperty('FOLDER_ID', f.getId());
  }
  if (!props.getProperty('ADMIN_KEY')) makeAdminKey();
  tokKey_(true);
  Logger.log('시트: ' + ss.getUrl());
  Logger.log('증빙 폴더: ' + DriveApp.getFolderById(props.getProperty('FOLDER_ID')).getUrl());
}

/* 관리자 비밀번호 새로 만들기 — 실행하면 실행 로그에 새 비밀번호가 나옵니다(이전 비밀번호는 무효) */
function makeAdminKey() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
  let k = ''; for (let i = 0; i < 10; i++) k += chars[Math.floor(Math.random() * chars.length)];
  PropertiesService.getScriptProperties().setProperty('ADMIN_KEY', k);
  Logger.log('관리자 비밀번호: ' + k);
  return k;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('장학금 실적')
    .addItem('웹페이지 주소 보기', 'showUrl')
    .addItem('중복·누락 점검', 'checkDuplicates')
    .addToUi();
}

function showUrl() {
  SpreadsheetApp.getUi().alert('학생 입력 웹페이지 주소\n\n' + (ScriptApp.getService().getUrl() || '아직 배포 전입니다. 배포 → 새 배포 → 웹 앱'));
}

/* ---------------- 공통 ---------------- */
let SS_ = null;   // 한 번 실행하는 동안 시트를 한 번만 엶 (속도)
function ss_() { return SS_ || (SS_ = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'))); }

/* 다섯 탭을 한 번의 요청으로 읽음 (Google Sheets API 고급 서비스가 있으면 4~5배 빠름, 없으면 기존 방식) — 값은 화면에 보이는 글자 그대로 */
const TABS_ = ['설정', '학생', '논문', '특허', '수상'];
let ALL_ = null;
function readAll_() {
  if (ALL_) return ALL_;
  ALL_ = {};
  if (typeof Sheets !== 'undefined') {
    try {
      const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
      const r = Sheets.Spreadsheets.Values.batchGet(id, { ranges: TABS_.map(n => "'" + n + "'!A1:AZ"), valueRenderOption: 'FORMATTED_VALUE' });
      r.valueRanges.forEach((vr, i) => {
        const v = vr.values && vr.values.length ? vr.values : [[]], w = Math.max.apply(null, v.map(x => x.length));
        ALL_[TABS_[i]] = v.map(row => { const a = row.slice(); while (a.length < w) a.push(''); return a; });
      });
      return ALL_;
    } catch (e) { ALL_ = {}; }
  }
  TABS_.forEach(n => ALL_[n] = ss_().getSheetByName(n).getDataRange().getDisplayValues());
  return ALL_;
}
const colA1_ = c => { let s = ''; for (c++; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + (c - 1) % 26) + s; return s; };
/* 여러 칸을 한 번의 요청으로 기록 — data: [[탭, 행(1부터), 열(0부터), 2차원 값]] */
function writeAll_(data) {
  if (!data.length) return;
  if (typeof Sheets !== 'undefined') {
    try {
      Sheets.Spreadsheets.Values.batchUpdate({ valueInputOption: 'USER_ENTERED',
        data: data.map(([n, r, c, vals]) => ({ range: "'" + n + "'!" + colA1_(c) + r, values: vals })) },
        PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
      return;
    } catch (e) { /* 아래 기존 방식으로 */ }
  }
  data.forEach(([n, r, c, vals]) => ss_().getSheetByName(n).getRange(r, c + 1, vals.length, vals[0].length).setValues(vals));
}

function settings_() {
  const c = CacheService.getScriptCache(), hit = c.get('settings_v1');   // 설정은 30초 동안 기억 (시트를 매번 열지 않음)
  if (hit) { const o = JSON.parse(hit); if (o.__d) Object.keys(o.__d).forEach(k => o[k] = new Date(o.__d[k])); delete o.__d; return o; }
  const rows = readAll_()['설정'].slice(1);
  const o = {}, d = {}; rows.forEach(r => { if (r[0]) { const k = String(r[0]).trim(); o[k] = r[1]; if (r[1] instanceof Date) d[k] = r[1].getTime(); } });
  try { c.put('settings_v1', JSON.stringify(Object.assign({}, o, { __d: d })), 30); } catch (e) {}
  return o;
}
/* 임시 업로드 폴더 — ID를 기억해 두고 바로 엶 (속도) */
function tmpFolder_() {
  const P = PropertiesService.getScriptProperties(), id = P.getProperty('TMP_ID');
  if (id) try { return DriveApp.getFolderById(id); } catch (e) {}
  const f = getOrCreate_(DriveApp.getFolderById(P.getProperty('FOLDER_ID')), '_임시업로드');
  P.setProperty('TMP_ID', f.getId()); return f;
}

function parseDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  // '2026-12-04 23:59', '2026. 12. 4 오후 11:59:00', '2026.12.4. 23:59' 모두 읽음 (시각이 없으면 그날 23:59)
  const m = String(v).match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})(?:\D*?(오전|오후|AM|PM)?\s*(\d{1,2}):(\d{2}))?/i);
  if (!m) return null;
  let h = m[5] !== undefined ? +m[5] : 23; const mi = m[6] !== undefined ? +m[6] : 59;
  if (m[4] && /오후|PM/i.test(m[4]) && h < 12) h += 12;
  if (m[4] && /오전|AM/i.test(m[4]) && h === 12) h = 0;
  return new Date(+m[1], +m[2] - 1, +m[3], h, mi);
}

/* ---------------- 웹페이지 ---------------- */
function publicCfg_(e) {
  const s = settings_();
  const deadline = parseDate_(s['접수마감']);
  const p = (e && e.parameter) || {};
  return {
    title: s['평가명'] || '박사재학생 장학금 연구실적 제출',
    titleEn: s['평가명(영문)'] || '',
    lang: p.lang === 'en' ? 'en' : (p.lang === 'ko' ? 'ko' : ''),
    deadline: s['접수마감'] instanceof Date ? Utilities.formatDate(s['접수마감'], 'Asia/Seoul', 'yyyy-MM-dd HH:mm') : String(s['접수마감'] || ''),
    closed: deadline ? new Date() > deadline : false,
    from: s['실적인정 시작'] ? Utilities.formatDate(parseDate_(s['실적인정 시작']), 'Asia/Seoul', 'yyyy.M.d.') : '',
    to: s['실적인정 끝'] ? Utilities.formatDate(parseDate_(s['실적인정 끝']), 'Asia/Seoul', 'yyyy.M.d.') : '',
    profs: String(s['지도교수 목록'] || '').split(/[,，]/).map(x => x.trim()).filter(Boolean),
    contact: s['담당자'] || '',
    contactEmail: String(s['문의 메일'] || '').trim(),
    contactEn: s['담당자(영문)'] || '',
    notice: s['안내문'] || '',
    noticeEn: s['안내문(영문)'] || '',
    maxMb: Number(s['파일최대MB'] || 15)
  };
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* 깃허브 홈페이지(index.html)가 설정을 읽어 갈 때: …/exec?action=config
   주소만 열면: 이 웹앱 자체 화면(예비용) */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'config') return json_(publicCfg_(e));
  const t = HtmlService.createTemplate(INDEX_HTML);
  t.cfg = JSON.stringify(publicCfg_(e));
  return t.evaluate().setTitle('장학금 연구실적 제출 · Scholarship Submission | GIST')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/* 깃허브 홈페이지에서 보내는 업로드·제출 */
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const A = req.args || [];
    if (req.action === 'requestCode') return json_({ ok: true, data: requestCode(A[0], A[1]) });
    if (req.action === 'verifyCode') return json_({ ok: true, data: verifyCode(A[0], A[1]) });
    if (req.action === 'myData') return json_({ ok: true, data: myData(A[0]) });
    if (req.action === 'myFile') return json_({ ok: true, data: myFile(A[0], A[1]) });
    if (req.action === 'uploadFile') return json_({ ok: true, data: uploadFile(A[0], A[1], A[2]) });
    if (req.action === 'submit') return json_({ ok: true, data: submit(A[0], A[1]) });
    if (req.action === 'adminList') return json_({ ok: true, data: adminList(req.args[0]) });
    if (req.action === 'adminFile') return json_({ ok: true, data: adminFile(req.args[0], req.args[1]) });
    throw new Error('알 수 없는 요청 / Unknown request');
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------------- 관리자(담당자) 화면용 ---------------- */
function checkAdmin_(key) {
  const cache = CacheService.getScriptCache(), fk = 'admfail_' + Math.floor(Date.now() / 600000);
  if (Number(cache.get(fk) || 0) >= 20) throw new Error('비밀번호를 너무 많이 틀렸습니다. 10분 뒤 다시 시도해 주세요.');
  const k = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!k || String(key || '') !== k) {
    cache.put(fk, String(Number(cache.get(fk) || 0) + 1), 700);
    Utilities.sleep(1500); throw new Error('관리자 비밀번호가 맞지 않습니다.');
  }
}

/* ---------------- 학생 본인 확인 (GIST 메일 인증번호) ---------------- */
function emailOk_(email) {
  const doms = String(settings_()['허용 메일 도메인'] || 'gist.ac.kr, gm.gist.ac.kr').toLowerCase().split(/[,\s]+/).filter(Boolean);
  const e = String(email || '').trim().toLowerCase();
  if (!/^[a-z0-9._%+-]{1,64}@[a-z0-9.-]{3,}$/.test(e)) return '';
  return doms.indexOf(e.split('@')[1]) >= 0 ? e : '';
}

function rand6_() {
  const h = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  return String(parseInt(h.slice(0, 12), 16) % 1000000).padStart(6, '0');
}

function requestCode(email, lang) {
  const e = emailOk_(email);
  if (!e) throw new Error('GIST 메일 주소(@gist.ac.kr 또는 @gm.gist.ac.kr)를 입력해 주세요. / Please enter your GIST email address.');
  const c = CacheService.getScriptCache();
  if (c.get('otpwait_' + e)) throw new Error('인증번호는 1분에 한 번 받을 수 있습니다. / Please wait a minute before requesting again.');
  const hk = 'otph_' + e + '_' + Math.floor(Date.now() / 3600000), n = Number(c.get(hk) || 0) + 1;
  if (n > 5) throw new Error('요청이 너무 많습니다. 1시간 뒤 다시 시도해 주세요. / Too many requests. Try again in an hour.');
  if (MailApp.getRemainingDailyQuota() < 1) throw new Error('오늘 보낼 수 있는 인증 메일이 모두 소진되었습니다. 담당자에게 문의해 주세요. / Daily email limit reached. Please contact the office.');
  const code = rand6_();
  c.put('otp_' + e, JSON.stringify({ code: code, tries: 0 }), 600);
  c.put('otpwait_' + e, '1', 60); c.put(hk, String(n), 3700);
  const s = settings_();
  const rt = String(s['문의 메일'] || 'jnryu@gist.ac.kr').trim().replace(/[<>"'\s]/g, '');
  const html = '<div style="font-family:Apple SD Gothic Neo,Malgun Gothic,sans-serif;max-width:480px;margin:0 auto;padding:28px;border:1px solid #e5e5e5;border-radius:14px">' +
    '<div style="font-size:13px;color:#DB2615;font-weight:700">GIST 기계로봇공학과 · Dept. of Mechanical and Robotics Engineering</div>' +
    '<h2 style="margin:10px 0 6px;font-size:20px;color:#1B1F23">' + String(s['평가명'] || '장학금 연구실적 제출') + '</h2>' +
    '<p style="color:#555;font-size:14px;margin:0 0 18px">본인 확인 인증번호입니다. 10분 안에 입력해 주세요.<br>Your verification code (valid for 10 minutes):</p>' +
    '<div style="font-size:34px;letter-spacing:10px;font-weight:800;color:#1B1F23;background:#F5F6F4;border-radius:10px;padding:16px;text-align:center">' + code + '</div>' +
    '<p style="color:#888;font-size:12px;margin-top:18px">직접 요청하지 않았다면 이 메일을 무시하세요. If you did not request this, please ignore this email.<br>문의 Contact: ' + String(s['담당자'] || '') + (rt ? ' · <a href="mailto:' + rt + '" style="color:#DB2615">' + rt + '</a>' : '') + '<br>이 메일에 답장하면 담당자에게 전달됩니다. Replies to this email go to the office.</p></div>';
  MailApp.sendEmail({ to: e, name: 'GIST 기계로봇공학과 장학금 접수', subject: '[GIST 기계로봇공학과] 인증번호 ' + code + ' (Verification code)',
    replyTo: rt, htmlBody: html, body: '인증번호 / Verification code: ' + code + ' (10분 / 10 min)' + (rt ? '\n문의 Contact: ' + rt : '') });
  return { sent: true, email: e };
}

function verifyCode(email, code) {
  const e = emailOk_(email); if (!e) throw new Error('메일 주소를 확인해 주세요. / Check your email address.');
  const c = CacheService.getScriptCache(), raw = c.get('otp_' + e);
  if (!raw) throw new Error('인증번호가 만료되었습니다. 다시 받아 주세요. / The code has expired. Please request a new one.');
  const o = JSON.parse(raw); o.tries++;
  if (o.tries > 5) { c.remove('otp_' + e); throw new Error('5번 틀려 인증번호가 무효가 되었습니다. 다시 받아 주세요. / Too many attempts. Request a new code.'); }
  if (String(code || '').trim() !== o.code) { c.put('otp_' + e, JSON.stringify(o), 600); throw new Error('인증번호가 맞지 않습니다. / Incorrect code.'); }
  c.remove('otp_' + e);
  return { token: signTok_(e, Date.now() + 6 * 3600e3, true), email: e, hours: 6 };
}

/* 본인 확인 표 — 서버에 따로 저장하지 않고 서명으로 확인 (캐시가 지워져도 6시간 동안 유지) */
const hex_ = bytes => bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
function tokKey_(create) {
  const P = PropertiesService.getScriptProperties(); let k = P.getProperty('TOKEN_KEY');
  if (!k && create) {   // 서명 열쇠는 처음 한 번만 만듦 (인증번호 확인 때). 확인 단계에서는 절대 새로 만들지 않음
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    try { k = P.getProperty('TOKEN_KEY'); if (!k) { k = Utilities.getUuid() + Utilities.getUuid(); P.setProperty('TOKEN_KEY', k); } } finally { lock.releaseLock(); }
  }
  return k;
}
function signTok_(email, exp, create) {
  const key = tokKey_(create); if (!key) return '';
  const msg = email + '|' + Math.floor(exp);
  return hex_(Utilities.newBlob(email).getBytes()) + '.' + Math.floor(exp) + '.' + hex_(Utilities.computeHmacSha256Signature(msg, key));
}
function who_(token) {
  const fail = (why) => { throw new Error('LOGIN: 본인 확인이 만료되었습니다. 다시 인증해 주세요. / Your session has expired. Please verify again. [' + (why || '') + ']'); };
  const m = String(token || '').match(/^([0-9a-f]{2,400})\.(\d{13})\.([0-9a-f]{64})$/);
  if (!m) fail('f'); if (Number(m[2]) < Date.now()) fail('x');
  let email = '';
  try { email = Utilities.newBlob(m[1].match(/../g).map(h => { const v = parseInt(h, 16); return v > 127 ? v - 256 : v; })).getDataAsString(); } catch (e) { fail('d'); }
  if (!/^[a-z0-9._%+-]{1,64}@[a-z0-9.-]{3,}$/.test(email)) fail('e');
  let sig = signTok_(email, Number(m[2]));
  if (sig !== m[0]) { Utilities.sleep(800); sig = signTok_(email, Number(m[2])); }   // 배포 직후 설정값을 늦게 읽는 경우 한 번 더
  if (sig !== m[0]) fail(sig ? 's' : 'k');
  return email;
}

/* 이 메일로 낸 제출의 증빙 파일 ID 목록 */
function ownFileIds_(email) {
  const ids = {}, mine = {};
  sheetObjects_('학생').forEach(s => { if (String(s['이메일']).toLowerCase() === email) mine[s['접수번호']] = 1; });
  [['논문', ['증빙1', '증빙2']], ['특허', ['증빙']], ['수상', ['증빙']]].forEach(([n, cols]) =>
    sheetObjects_(n).forEach(r => { if (mine[r['접수번호']]) cols.forEach(k => { const m = String(r[k] || '').match(/[-\w]{25,}/); if (m) ids[m[0]] = r[k]; }); }));
  return ids;
}

/* 학생 본인 제출 내역 (AI·담당자 검증 칸은 보여 주지 않음) */
function myData(token) {
  const email = who_(token);
  const pick = (o, keys) => { const r = {}; keys.forEach(k => r[k] = o[k]); return r; };
  const by = {};
  sheetObjects_('학생').forEach(s => { if (String(s['이메일']).toLowerCase() === email)
    by[s['접수번호']] = Object.assign(pick(s, ['접수번호', '접수시각', '상태', '학번', '성명', '논문영문명', '과정', '재학학기', '지도교수', '학생메모']), { papers: [], patents: [], awards: [] }); });
  const P = ['순번', '저널명(학생)', '논문제목', 'DOI', 'Accept일', '역할(학생)', '저자수(학생)', '증빙1', '증빙2'],
        T = ['순번', '구분', '명칭', '출원·등록번호', '일자', '기여도(%)', '증빙'],
        W = ['순번', '구분', '학회·대회명', '수상명', '일자', '증빙'];
  [['논문', 'papers', P], ['특허', 'patents', T], ['수상', 'awards', W]].forEach(([n, k, keys]) =>
    sheetObjects_(n).forEach(r => { if (by[r['접수번호']]) by[r['접수번호']][k].push(pick(r, keys)); }));
  return { email: email, submissions: Object.keys(by).map(k => by[k]).sort((a, b) => String(b['접수시각']).localeCompare(String(a['접수시각']))) };
}

function myFile(token, url) {
  const email = who_(token);
  const m = String(url).match(/[-\w]{25,}/);
  if (!m || !ownFileIds_(email)[m[0]]) throw new Error('본인이 제출한 파일만 볼 수 있습니다. / You can only view your own files.');
  const f = DriveApp.getFileById(m[0]), bytes = f.getBlob().getBytes(), kind = sniff_(bytes);
  if (!kind) throw new Error('PDF·JPG·PNG 파일이 아닙니다.');
  return { name: f.getName(), mime: kind.mime, data: Utilities.base64Encode(bytes) };
}

function sheetObjects_(name) {
  const v = readAll_()[name];
  const h = v[0];
  return v.slice(1).filter(r => r.join('') !== '').map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
}

function adminList(key) {
  checkAdmin_(key);
  const by = {};
  sheetObjects_('학생').forEach(s => { by[s['접수번호']] = Object.assign({ papers: [], patents: [], awards: [] }, s); });
  [['논문', 'papers'], ['특허', 'patents'], ['수상', 'awards']].forEach(([n, k]) =>
    sheetObjects_(n).forEach(r => { if (by[r['접수번호']]) by[r['접수번호']][k].push(r); }));
  const props = PropertiesService.getScriptProperties();
  return {
    students: Object.keys(by).map(k => by[k]).sort((a, b) => String(b['접수시각']).localeCompare(String(a['접수시각']))),
    sheetUrl: ss_().getUrl(),
    folderUrl: DriveApp.getFolderById(props.getProperty('FOLDER_ID')).getUrl(),
    now: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm')
  };
}

/* 증빙 파일 내용을 관리자 화면에 보여 주기 (증빙 폴더 안의 파일만) */
function adminFile(key, url) {
  checkAdmin_(key);
  const m = String(url).match(/[-\w]{25,}/);
  if (!m) throw new Error('파일 주소가 올바르지 않습니다.');
  const f = DriveApp.getFileById(m[0]);
  const root = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
  let ok = false, level = [f.getParents()];
  for (let d = 0; d < 4 && !ok && level.length; d++) {
    const next = [];
    level.forEach(it => { while (it.hasNext()) { const p = it.next(); if (p.getId() === root) ok = true; next.push(p.getParents()); } });
    level = next;
  }
  if (!ok) throw new Error('증빙 폴더의 파일이 아닙니다.');
  const bytes = f.getBlob().getBytes(), kind = sniff_(bytes);
  if (!kind) throw new Error('PDF·JPG·PNG 파일이 아닙니다.');
  return { name: f.getName(), mime: kind.mime, data: Utilities.base64Encode(bytes) };
}

/* 파일 하나 업로드 (제출 전에 하나씩 올려 용량 한도 회피) → 파일 ID 반환 */
function uploadFile(meta, dataUrl, token) {
  who_(token);   // 본인 확인한 학생만 업로드
  const s = settings_();
  // 너무 잦은 업로드 막기 (1분에 60개 초과 시 거절)
  const cache = CacheService.getScriptCache(), ck = 'up_' + Math.floor(Date.now() / 60000);
  const cnt = Number(cache.get(ck) || 0) + 1; cache.put(ck, String(cnt), 120);
  if (cnt > 60) throw new Error('잠시 후 다시 시도해 주세요. / Too many uploads, please retry shortly.');
  const m = String(dataUrl).match(/^data:([^;]*);base64,(.*)$/);
  if (!m) throw new Error('파일을 읽지 못했습니다. / Could not read the file.');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > Number(s['파일최대MB'] || 15) * 1024 * 1024) throw new Error('파일이 너무 큽니다. / The file is too large.');
  const kind = sniff_(bytes);   // 파일 내용으로 종류 판별 (확장자·브라우저 정보는 믿지 않음)
  if (!kind) throw new Error('증빙은 PDF·JPG·PNG 파일만 올릴 수 있습니다. / Only PDF, JPG or PNG files are accepted.');
  const tmp = tmpFolder_();
  const base = String(meta && meta.name || 'file').replace(/\.[^.]*$/, '').replace(/[\\/:*?"<>|]/g, '_').slice(0, 80);
  const f = tmp.createFile(Utilities.newBlob(bytes, kind.mime, base + kind.ext));
  return f.getId();
}

/* 파일 앞부분(매직 넘버)으로 PDF·PNG·JPG만 인정 */
function sniff_(b) {
  const u = i => (b[i] + 256) % 256;
  if (b.length > 4 && u(0) === 0x25 && u(1) === 0x50 && u(2) === 0x44 && u(3) === 0x46) return { mime: 'application/pdf', ext: '.pdf' };
  if (b.length > 8 && u(0) === 0x89 && u(1) === 0x50 && u(2) === 0x4E && u(3) === 0x47) return { mime: 'image/png', ext: '.png' };
  if (b.length > 3 && u(0) === 0xFF && u(1) === 0xD8 && u(2) === 0xFF) return { mime: 'image/jpeg', ext: '.jpg' };
  return null;
}

/* 학생이 쓴 글자가 시트에서 수식(=, +, -, @)으로 실행되지 않게 막고 길이 제한 */
function t_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').slice(0, max || 300);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

/* 제출되지 않고 남은 임시 업로드 파일(2일 지난 것) 정리 */
function cleanTemp_() {
  const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
  const it = root.getFoldersByName('_임시업로드'); if (!it.hasNext()) return;
  const files = it.next().getFiles(), limit = Date.now() - 2 * 86400000; let n = 0;
  while (files.hasNext() && n < 50) { const f = files.next(); if (f.getDateCreated().getTime() < limit) { f.setTrashed(true); n++; } }
}

function getOrCreate_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/* 최종 제출 */
function submit(p, token) {
  const T0 = Date.now(), MS = {}, mark = k => MS[k] = Date.now() - T0;   // 단계별 걸린 시간 (속도 점검용)
  const email = who_(token);   // 인증된 GIST 메일
  const s = settings_();
  const deadline = parseDate_(s['접수마감']);
  if (deadline && new Date() > deadline) throw new Error('접수가 마감되었습니다. / Submissions are closed.');
  ['학번', '성명', '과정', '지도교수'].forEach(k => { if (!p[k] || !String(p[k]).trim()) throw new Error(k + '을(를) 입력해 주세요. / Required field missing.'); });
  if (!/[A-Za-z]/.test(String(p.영문명 || ''))) throw new Error('논문에 쓰는 영문 이름을 입력해 주세요. / Please enter the name used in your papers.');
  if (!/^\d{8}$/.test(String(p.학번).trim())) throw new Error('학번은 숫자 8자리로 입력해 주세요. / Student ID must be 8 digits.');

  // 동시에 여러 명이 내면 한 명씩 차례로 기록 (1건 약 3~4초). 마감 직전 몰림에 대비해 최대 100초까지 기다림
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(100000)) throw new Error('지금 제출이 몰려 있습니다. 1~2분 뒤 다시 제출해 주세요(입력한 내용은 화면에 그대로 있습니다). / Too many submissions right now. Please try again in a minute or two.');
  mark('lock');
  try {
    const id = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMdd-HHmmss') + '-' + String(p.학번).slice(-4);
    const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    const who = email;

    // 네 시트를 한 번씩만 읽음 (속도) — 이후 확인은 모두 이 값으로
    const NAMES = ['학생', '논문', '특허', '수상'], V = readAll_();
    mark('read');
    const col = (n, k) => V[n][0].indexOf(k);
    const sid = String(p.학번);

    // 학번과 메일을 1:1로 묶음 — 다른 사람이 남의 학번으로 제출하거나, 한 계정이 여러 학번으로 제출하는 것 막기
    const cE = col('학생', '이메일'), cN = col('학생', '학번'), cR = col('학생', '접수번호');
    const mine = {};
    V['학생'].slice(1).forEach(r => {
      const re = String(r[cE] || '').toLowerCase(), rid = String(r[cN]).replace(/^'/, '');
      if (!rid) return;
      if (rid === sid && re && re !== email) throw new Error('이 학번은 다른 GIST 계정으로 이미 제출되었습니다. 담당자에게 문의해 주세요. / This student ID was already submitted from another account. Please contact the office.');
      if (re === email && rid !== sid) throw new Error('이 계정은 이미 다른 학번(' + rid + ')으로 제출했습니다. 담당자에게 문의해 주세요. / This account already submitted under another student ID.');
      if (re === email) mine[r[cR]] = 1;
    });
    // 이 메일로 예전에 낸 증빙 (다시 올리지 않고 그대로 쓸 수 있음)
    const own = {};
    [['논문', ['증빙1', '증빙2']], ['특허', ['증빙']], ['수상', ['증빙']]].forEach(([n, cols]) => {
      const ci = col(n, '접수번호'), cc = cols.map(k => col(n, k));
      V[n].slice(1).forEach(r => { if (mine[r[ci]]) cc.forEach(c => { const m = String(r[c] || '').match(/[-\w]{25,}/); if (m) own[m[0]] = r[c]; }); });
    });

    // 먼저 증빙 파일이 모두 이번에 임시로 올린 파일인지 확인 (확인 전에는 아무것도 기록하지 않음)
    const tmpId = tmpFolder_().getId();
    const fids = [], FILES = {};
    (p.papers || []).forEach(x => { if (x.file1) fids.push(x.file1); if (x.file2) fids.push(x.file2); });
    (p.patents || []).forEach(x => { if (x.file) fids.push(x.file); });
    (p.awards || []).forEach(x => { if (x.file) fids.push(x.file); });
    fids.forEach(fid => {
      fid = String(fid); if (own[fid]) return;   // 이전에 본인이 낸 파일 그대로 쓰기
      let ok = false;
      try { const f = DriveApp.getFileById(fid), par = f.getParents(); ok = par.hasNext() && par.next().getId() === tmpId; if (ok) FILES[fid] = f; } catch (e) {}
      if (!ok) throw new Error('증빙 파일을 다시 올려 주세요. / Please upload the evidence file again.');
    });

    mark('check');
    // 증빙 파일 → 학번_성명 폴더로 이동·이름 정리 (위에서 확인한 파일만)
    const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
    const safe = String(p.성명).replace(/[\\/:*?"<>|]/g, '_').slice(0, 40);
    const folder = getOrCreate_(root, sid + '_' + safe);
    const place = (fid, label) => {
      if (!fid) return '';
      fid = String(fid); if (own[fid]) return own[fid];   // 기존 본인 파일은 그대로
      const f = FILES[fid]; if (!f) throw new Error('증빙 파일 오류 / Invalid file');
      const ext = (f.getName().match(/\.[^.]+$/) || ['.pdf'])[0];
      f.setName(sid + '_' + safe + '_' + label + '_' + id + ext);
      f.moveTo(folder);
      return f.getUrl();
    };

    const papers = (p.papers || []).filter(x => x.journal);
    const patents = (p.patents || []).filter(x => x.type);
    const awards = (p.awards || []).filter(x => x.type);

    // 줄을 모두 만든 뒤 (파일 이동 포함) 시트마다 한 번에 기록
    const ROWS = {
      학생: [[id, now, '유효', "'" + sid, t_(p.성명, 60), t_(p.과정, 20), t_(p.재학학기, 10), t_(p.지도교수, 40),
        email, who, papers.length, patents.length, awards.length, folder.getUrl(), t_(p.메모, 1000), t_(p.영문명, 80)]],
      논문: papers.map((x, i) => [id, '유효', "'" + sid, t_(p.성명, 60), i + 1,
        t_(x.journal, 200), t_(x.title, 400), t_(x.doi, 120), t_(x.accept, 20), t_(x.role, 10), Number(x.authors) || '',
        place(x.file1, '논문' + (i + 1)), place(x.file2, '논문' + (i + 1) + '_추가')]),
      특허: patents.map((x, i) => [id, '유효', "'" + sid, t_(p.성명, 60), i + 1,
        t_(x.type, 10), t_(x.name, 300), t_(x.no, 60), t_(x.date, 20), Number(x.share) || '', place(x.file, '특허' + (i + 1))]),
      수상: awards.map((x, i) => [id, '유효', "'" + sid, t_(p.성명, 60), i + 1,
        t_(x.type, 10), t_(x.event, 200), t_(x.prize, 200), t_(x.date, 20), place(x.file, '수상' + (i + 1))])
    };

    mark('files');
    // 같은 학번 이전 제출은 '대체됨' 처리 (최신 제출만 유효) → 새 줄 기록
    const W = [];
    NAMES.forEach(n => {
      const ci = col(n, '학번'), cs = col(n, '상태');
      for (let r = 1; r < V[n].length; r++) if (String(V[n][r][ci]).replace(/^'/, '') === sid && V[n][r][cs] === '유효') W.push([n, r + 1, cs, [['대체됨']]]);
      let last = V[n].length; while (last > 1 && V[n][last - 1].join('') === '') last--;   // 마지막 자료 줄 다음에 기록
      if (ROWS[n].length) W.push([n, last + 1, 0, ROWS[n]]);
    });
    writeAll_(W);

    try { const cc = CacheService.getScriptCache(); if (!cc.get('cleaned')) { cc.put('cleaned', '1', 3600); cleanTemp_(); } } catch (e) {}   // 1시간에 한 번만
    mark('write');
    return { id: id, papers: papers.length, patents: patents.length, awards: awards.length, time: now, ms: MS };
  } finally { lock.releaseLock(); }
}

/* ---------------- 담당자 점검 (시트 메뉴) ---------------- */
function checkDuplicates() {
  const ss = ss_();
  const v = ss.getSheetByName('논문').getDataRange().getValues();
  const h = v[0], cS = h.indexOf('상태'), cN = h.indexOf('학번'), cJ = h.indexOf('저널명(학생)'),
        cD = h.indexOf('DOI'), cT = h.indexOf('논문제목'), cM = h.indexOf('담당자메모'), c1 = h.indexOf('증빙1');
  const prev = ss.getSheetByName('이전인정논문').getDataRange().getValues().slice(1);
  const key = (doi, title) => (String(doi || '').toLowerCase().replace(/^https?:\/\/(dx\.)?doi\.org\//, '').trim()) ||
                              String(title || '').toLowerCase().replace(/[^a-z0-9가-힣]/g, '');
  const seen = {}; let n = 0;
  for (let r = 1; r < v.length; r++) {
    if (v[r][cS] !== '유효') continue;
    const notes = [];
    const k = key(v[r][cD], v[r][cT]);
    const sk = v[r][cN] + '|' + k;
    if (k && seen[sk]) notes.push('같은 학생이 같은 논문을 두 번 입력');
    seen[sk] = true;
    if (k && prev.some(p => String(p[1]) === String(v[r][cN]) && key(p[4], p[5]) === k)) notes.push('이전 포상에서 이미 인정된 논문');
    if (!v[r][c1]) notes.push('증빙 파일 없음');
    if (notes.length) { ss.getSheetByName('논문').getRange(r + 1, cM + 1).setValue(notes.join(' / ')); n++; }
  }
  SpreadsheetApp.getUi().alert('점검 완료: 확인할 논문 ' + n + '건 (논문 시트 담당자메모 칸에 표시)');
}


/* ---------------- 학생 입력 화면 (HTML) — 수정하지 않아도 됩니다 ---------------- */
const INDEX_HTML = `<!DOCTYPE html>
<html lang="ko">
<head>
<base target="_top">
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
<style>
:root{
  --red:#DC2715; --red-d:#B81F10; --red-s:#FDECEA; --ink:#1B1F23; --gray:#4F5B54; --mute:#7A8580;
  --line:#E4E7E5; --bg:#F5F6F4; --card:#FFFFFF; --ok:#1E7B4F; --ok-s:#E7F4EC; --warn:#A35C00; --warn-s:#FFF4E0;
  --r:16px;
}
*{box-sizing:border-box}
/* PC 화면에서는 글자·버튼을 전체적으로 키워 읽기 쉽게 (휴대폰은 그대로) */
@media (min-width:900px){html{zoom:1.12}}
@media (min-width:1500px){html{zoom:1.2}}
@media (min-width:2200px){html{zoom:1.4}}
body{margin:0;background:var(--bg);color:var(--ink);font-family:Pretendard,-apple-system,"Malgun Gothic",sans-serif;font-size:15px;line-height:1.55;-webkit-font-smoothing:antialiased;word-break:keep-all;overflow-wrap:anywhere}
.hero{padding-top:58px!important;background:linear-gradient(135deg,#1B1F23 0%,#2B3330 60%,#4F5B54 100%);color:#fff;padding:44px 16px 72px;position:relative;overflow:hidden}
.hero:after{content:"";position:absolute;right:-80px;top:-80px;width:320px;height:320px;border-radius:50%;background:radial-gradient(circle,rgba(220,39,21,.55),rgba(220,39,21,0) 70%)}
.wrap{max-width:788px;margin:0 auto;position:relative;z-index:1}
.brand{display:flex;align-items:center;gap:10px;font-weight:700;letter-spacing:.02em;font-size:13px;opacity:.9}
.brand i{display:inline-block;width:10px;height:10px;border-radius:2px;background:var(--red)}
.hero h1{font-size:28px;font-weight:800;margin:14px 0 8px;letter-spacing:-.02em;line-height:1.3}
.hero p{margin:0;opacity:.82;font-size:14.5px}
.chips{display:flex;flex-wrap:wrap;gap:8px;margin-top:18px}
.chip{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.18);border-radius:999px;padding:6px 12px;font-size:13px}
.chip b{font-weight:700}
main{max-width:820px;margin:-44px auto 0;padding:0 16px 120px;position:relative;z-index:2}
.card{background:var(--card);border-radius:var(--r);box-shadow:0 1px 2px rgba(0,0,0,.04),0 8px 24px rgba(27,31,35,.06);padding:24px;margin-bottom:16px}
.card h2{font-size:18px;margin:0 0 4px;display:flex;align-items:center;gap:10px;font-weight:800;letter-spacing:-.01em}
.num{width:26px;height:26px;border-radius:8px;background:var(--red);color:#fff;font-size:13px;display:inline-flex;align-items:center;justify-content:center;font-weight:800}
.desc{color:var(--mute);font-size:13.5px;margin:0 0 18px}
.notice{background:var(--warn-s);border-left:4px solid #F0A020;padding:12px 14px;border-radius:10px;font-size:13.5px;color:#5E3C00;margin-bottom:16px}
.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}
.grid .full{grid-column:1/-1}
.grid>*{min-width:0}
.item{min-width:0}
@media(max-width:640px){.grid{grid-template-columns:1fr}.hero h1{font-size:23px}.card{padding:18px}}
label{display:block;font-size:13px;font-weight:700;color:var(--gray);margin-bottom:6px}
label .req{color:var(--red);margin-left:2px}
label .opt{color:var(--mute);font-weight:500;margin-left:4px}
input,select,textarea{width:100%;font:inherit;font-size:15px;padding:11px 13px;border:1.5px solid var(--line);border-radius:10px;background:#fff;color:var(--ink);transition:.15s}
input:focus,select:focus,textarea:focus{outline:none;border-color:var(--red);box-shadow:0 0 0 4px rgba(220,39,21,.1)}
.hint{font-size:12px;color:var(--mute);margin-top:5px}
.seg{display:grid;grid-template-columns:repeat(4,1fr);gap:6px}
.seg.two{grid-template-columns:repeat(2,1fr)}
@media(max-width:640px){.seg{grid-template-columns:repeat(2,1fr)}}
.seg input{display:none}
.seg label{margin:0;border:1.5px solid var(--line);border-radius:10px;padding:10px 6px;text-align:center;cursor:pointer;font-weight:600;color:var(--ink);font-size:13.5px;line-height:1.3;background:#fff}
.seg label small{display:block;color:var(--mute);font-weight:500;font-size:11.5px}
.seg input:checked+label{border-color:var(--red);background:var(--red-s);color:var(--red-d)}
.item{border:1.5px solid var(--line);border-radius:14px;padding:18px;margin-bottom:12px;position:relative;background:#FCFCFB}
.item-h{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px}
.item-h b{font-size:14.5px}
.del{border:none;background:none;color:var(--mute);cursor:pointer;font:inherit;font-size:13px;padding:4px 8px;border-radius:8px}
.del:hover{background:#F0F0F0;color:var(--red)}
.add{width:100%;border:1.5px dashed #C9CFCB;background:#fff;border-radius:12px;padding:13px;font:inherit;font-weight:700;color:var(--gray);cursor:pointer}
.add:hover{border-color:var(--red);color:var(--red)}
.file{display:flex;align-items:center;gap:10px;border:1.5px dashed #C9CFCB;border-radius:10px;padding:10px 12px;background:#fff;cursor:pointer}
.file input{display:none}
.file .ico{width:34px;height:34px;border-radius:8px;background:var(--red-s);color:var(--red);display:flex;align-items:center;justify-content:center;font-weight:800;font-size:11px;flex:none}
.file .t{font-size:13.5px;color:var(--gray);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.file.has{border-style:solid;border-color:#9BC9AE;background:var(--ok-s)}
.file.has .ico{background:#fff;color:var(--ok)}
.empty{color:var(--mute);font-size:13.5px;text-align:center;padding:6px 0 14px}
.check{display:flex;gap:10px;align-items:flex-start;font-size:14px}
.check input{width:20px;height:20px;flex:none;margin-top:2px;accent-color:var(--red)}
.bar{position:fixed;left:0;right:0;bottom:0;background:rgba(255,255,255,.94);backdrop-filter:blur(10px);border-top:1px solid var(--line);z-index:5}
.bar .in{max-width:820px;margin:0 auto;padding:12px 16px;display:flex;align-items:center;gap:12px}
.sum{flex:1;font-size:13px;color:var(--gray)}
.sum b{color:var(--ink)}
.go{border:none;background:var(--red);color:#fff;font:inherit;font-weight:800;font-size:15.5px;padding:13px 26px;border-radius:12px;cursor:pointer;box-shadow:0 6px 16px rgba(220,39,21,.28)}
.go:hover{background:var(--red-d)}
.go:disabled{background:#C9CFCB;box-shadow:none;cursor:default}
.err{color:var(--red);font-size:12.5px;margin-top:5px;display:none}
.bad input,.bad select{border-color:var(--red)}
.bad .err{display:block}
.modal{position:fixed;inset:0;background:rgba(27,31,35,.55);display:none;align-items:center;justify-content:center;padding:16px;z-index:10}
.modal.on{display:flex}
.box{background:#fff;border-radius:20px;max-width:440px;width:100%;padding:28px;text-align:center}
.box .big{width:64px;height:64px;border-radius:50%;background:var(--ok-s);color:var(--ok);display:flex;align-items:center;justify-content:center;font-size:30px;margin:0 auto 14px}
.box h3{margin:0 0 6px;font-size:20px}
.box p{color:var(--gray);font-size:14px;margin:4px 0}
.prog{height:6px;background:#EEE;border-radius:9px;overflow:hidden;margin:16px 0 6px}
.prog i{display:block;height:100%;width:0;background:var(--red);transition:width .3s}
.closed{text-align:center;padding:40px 20px}
.role-help{font-size:12px;color:var(--mute);margin-top:6px}
footer{text-align:center;color:var(--mute);font-size:12px;margin-top:24px}
.lang{position:absolute;top:12px;right:12px;z-index:20;display:flex;background:rgba(255,255,255,.14);backdrop-filter:blur(8px);border:1px solid rgba(255,255,255,.25);border-radius:999px;padding:3px}
.lang button{border:none;background:none;color:#fff;font:inherit;font-size:12.5px;font-weight:700;padding:6px 12px;border-radius:999px;cursor:pointer;opacity:.75}
.lang button.on{background:#fff;color:#1B1F23;opacity:1}

/* ---------- 관리자 화면 ---------- */
.adminlink{display:block;text-align:center;margin:10px auto 0;color:var(--mute);font-size:12px;text-decoration:none;opacity:.7}
.adminlink:hover{opacity:1;color:var(--red)}
#admin{max-width:1180px;margin:-44px auto 0;padding:0 16px 80px;position:relative;z-index:2}
.a-login{max-width:420px;margin:0 auto}
.a-top{display:flex;flex-wrap:wrap;gap:10px;align-items:center;justify-content:space-between;margin-bottom:14px}
.a-stats{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:14px}
@media(max-width:760px){.a-stats{grid-template-columns:repeat(2,1fr)}}
.a-stat{background:#fff;border-radius:14px;padding:14px 16px;box-shadow:0 1px 2px rgba(0,0,0,.04),0 8px 24px rgba(27,31,35,.06)}
.a-stat b{display:block;font-size:24px;font-weight:800;letter-spacing:-.02em}
.a-stat span{font-size:12.5px;color:var(--mute)}
.a-tools{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
.a-tools input[type=search]{width:240px;padding:9px 12px}
.btn{border:1.5px solid var(--line);background:#fff;border-radius:10px;padding:9px 14px;font:inherit;font-size:13.5px;font-weight:700;color:var(--ink);cursor:pointer;text-decoration:none;display:inline-flex;align-items:center;gap:6px}
.btn:hover{border-color:var(--red);color:var(--red)}
.btn.red{background:var(--red);border-color:var(--red);color:#fff}
.btn.sm{padding:5px 10px;font-size:12.5px;border-radius:8px}
.a-list{display:flex;flex-direction:column;gap:10px}
.a-row{background:#fff;border-radius:14px;box-shadow:0 1px 2px rgba(0,0,0,.04),0 6px 18px rgba(27,31,35,.05);overflow:hidden}
.a-head{display:grid;grid-template-columns:110px 1.2fr 1fr 1fr 150px 160px 24px;gap:10px;align-items:center;padding:14px 18px;cursor:pointer}
.a-head:hover{background:#FAFAF9}
@media(max-width:860px){.a-head{grid-template-columns:1fr 1fr;row-gap:4px}.a-head .hide-m{display:none}}
.a-head .nm{font-weight:800}
.a-head .sub{font-size:12.5px;color:var(--mute)}
.pill{display:inline-block;border-radius:999px;padding:2px 9px;font-size:12px;font-weight:700;background:#F0F2F1;color:var(--gray);margin-right:4px}
.pill.old{background:#FFF4E0;color:var(--warn)}
.pill.ok{background:var(--ok-s);color:var(--ok)}
.a-body{display:none;border-top:1px solid var(--line);padding:14px 18px 18px;background:#FCFCFB}
.a-row.open .a-body{display:block}
.a-row.open .chev{transform:rotate(90deg)}
.chev{transition:.15s;color:var(--mute)}
.a-body h4{margin:14px 0 8px;font-size:13.5px;color:var(--gray)}
.a-body h4:first-child{margin-top:0}
.a-tbl{width:100%;border-collapse:collapse;font-size:13px;background:#fff;border-radius:10px;overflow:hidden}
.a-tbl th{background:#F2F4F3;text-align:left;padding:8px 10px;font-size:12px;color:var(--gray);white-space:nowrap}
.a-tbl td{border-top:1px solid var(--line);padding:8px 10px;vertical-align:top}
.a-wrap{overflow-x:auto}
.viewer{position:fixed;inset:0;background:rgba(27,31,35,.72);display:none;z-index:30;padding:20px}
.viewer.on{display:flex;flex-direction:column}
.viewer .vbar{display:flex;gap:8px;align-items:center;color:#fff;margin-bottom:10px}
.viewer .vbar b{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.viewer .vbody{flex:1;background:#fff;border-radius:12px;overflow:auto;display:flex;align-items:center;justify-content:center}
.viewer iframe{width:100%;height:100%;border:0}
.viewer img{max-width:100%;max-height:100%}
.empty-a{text-align:center;color:var(--mute);padding:40px 0}
[hidden]{display:none!important}
.mlogo{height:26px;width:auto;display:block}
@media(max-width:640px){.mlogo{height:21px}.hero{padding-top:70px!important}}
.hero .rule{width:44px;height:3px;background:var(--red);border-radius:3px;margin:22px 0 12px}
.gate{max-width:560px;margin-left:auto;margin-right:auto}
.gate .ico-l{width:46px;height:46px;border-radius:14px;background:var(--red-s);color:var(--red);display:flex;align-items:center;justify-content:center;font-size:22px;margin-bottom:12px}
.row2{display:flex;gap:8px}.row2 input{flex:1;min-width:0}.row2 .btn{white-space:nowrap}
.code-in{font-size:24px!important;letter-spacing:10px;text-align:center;font-weight:800}
.gnote{font-size:12.5px;color:var(--mute);margin-top:14px;line-height:1.6}
.sess{display:flex;align-items:center;justify-content:space-between;gap:10px;background:#fff;border-radius:14px;padding:12px 16px;margin-bottom:12px;box-shadow:0 1px 2px rgba(0,0,0,.04),0 6px 18px rgba(27,31,35,.05);font-size:13.5px}
.sess .who{display:flex;align-items:center;gap:8px;min-width:0}.sess .who b{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sess .dot{width:8px;height:8px;border-radius:50%;background:var(--ok);flex:none}
.linkbtn{border:none;background:none;color:var(--mute);font:inherit;font-size:13px;cursor:pointer;text-decoration:underline}
.tabs{display:grid;grid-template-columns:1fr 1fr;background:#E9EBEA;border-radius:12px;padding:4px;margin-bottom:16px}
.tabs button{border:none;background:none;font:inherit;font-weight:700;font-size:14px;color:var(--gray);padding:10px;border-radius:9px;cursor:pointer}
.tabs button.on{background:#fff;color:var(--ink);box-shadow:0 1px 3px rgba(0,0,0,.08)}
.sub2{font-size:12.5px;color:var(--mute)}
.mine-head{display:flex;justify-content:space-between;align-items:flex-start;gap:10px;flex-wrap:wrap;margin-bottom:6px}
.file.keep{border-style:solid;border-color:#B9C7E8;background:#EEF2FB}
.file.keep .ico{background:#fff;color:#3F61B2}
input[readonly]{background:#F5F6F4;color:var(--gray)}
</style>
</head>
<body>
<div class="lang"><button type="button" id="lk" onclick="setLang('ko')">한국어</button><button type="button" id="le" onclick="setLang('en')">English</button></div>
<div class="hero"><div class="wrap">
  <img class="mlogo" src="data:image/svg+xml;base64,PHN2ZyB3aWR0aD0nMjc4JyBoZWlnaHQ9JzI0JyB2aWV3Qm94PScwIDAgMjc4IDI0JyBmaWxsPSdub25lJyB4bWxucz0naHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmcnPjxnIGNsaXAtcGF0aD0ndXJsKCNjbGlwMF80MTI2Xzg1NjQpJz48cGF0aCBkPSdNNzEuNTMzMSA1LjYwNjZMNjIuNzAzMSA1LjU5NzU2TDY5LjAxMzEgMC4wMjcwOTk2SDg2Ljg4MzFWNS41Nzk0N0g3OC4yMjMxVjI0SDcxLjUzMzFWNS42MDY2WicgZmlsbD0nd2hpdGUnLz48cGF0aCBkPSdNNjEuMDk3OCAxMS45ODE5QzU5LjU5NzggMTAuMjkwOSA1Ny4yOTc4IDkuMjA1NzMgNTQuNzE3OCA5LjIwNTczSDQ5LjUzNzhDNDguNDM3OCA5LjIwNTczIDQ3LjUzNzggOC40MDA5IDQ3LjUzNzggNy4zOTcxNEM0Ny41Mzc4IDYuMzkzMzcgNDguNDI3OCA1LjU4ODU1IDQ5LjUzNzggNS41ODg1NUg1Ni40Mzc4TDYyLjc1NzggMEg0OC45Mzc4QzQ0LjQyNzggMCA0MC43NTc4IDMuMzA5NzIgNDAuNzU3OCA3LjM5NzE0QzQwLjc1NzggOS4xNDI0MyA0MS40Mjc4IDEwLjc1MjEgNDIuNTU3OCAxMi4wMTgxQzQ0LjA1NzggMTMuNzA5MSA0Ni4zNTc4IDE0Ljc5NDMgNDguOTM3OCAxNC43OTQzSDU0LjExNzhDNTUuMjE3OCAxNC43OTQzIDU2LjExNzggMTUuNTk5MSA1Ni4xMTc4IDE2LjYwMjlDNTYuMTE3OCAxNy42MDY2IDU1LjIxNzggMTguNDExNSA1NC4xMTc4IDE4LjQxMTVINDAuNzU3OFYyNEg1NC43MTc4QzU5LjIzNzggMjQgNjIuODk3OCAyMC42OTAzIDYyLjg5NzggMTYuNjAyOUM2Mi44OTc4IDE0Ljg1NzYgNjIuMjI3OCAxMy4yNDc5IDYxLjA5NzggMTEuOTgxOVonIGZpbGw9J3doaXRlJy8+PHBhdGggZD0nTTIwLjAzIDkuNDU4OUwxNC44NyAxNC4wMTY1SDE5LjUxVjE4LjQ0NzZIMTMuOTNDMTAgMTguNDQ3NiA2LjgyIDE1LjU2MjkgNi44MiAxMi4wMTgxQzYuODIgOC40NzMyMiAxMCA1LjU4ODUyIDEzLjkzIDUuNTg4NTJIMjQuNDFMMzAuNzEgMC4wMjcwOTk2SDEzLjI2QzUuOTQgMC4wMjcwOTk2IDAgNS4zODk1NyAwIDEyLjAwOUMwIDE4LjYyODUgNS45NCAyNCAxMy4yNiAyNEgyNi4yMVY5LjQ1ODlIMjAuMDRIMjAuMDNaJyBmaWxsPScjREIyNjE1Jy8+PHBhdGggZD0nTTMwLjE0MDYgMjRIMzYuODMwNlYwLjAwOTAzMzJMMzAuMTQwNiA1Ljk0MTIxVjI0WicgZmlsbD0nd2hpdGUnLz48cGF0aCBvcGFjaXR5PScwLjMnIGQ9J00xMDAuMzgzIDJWMjInIHN0cm9rZT0nd2hpdGUnLz48cGF0aCBkPSdNMjA2LjI3NyAxLjAwNDY0VjMuMzQ4ODJIMjA2LjU0M1Y0LjI3NjE5SDIwMS45OTJWMy4zNDg4MkgyMDMuMTI2VjIuNTkzMTlIMjA0LjE5OVYzLjM0ODgySDIwNS4yMDRWMS45MzIwMUgyMDIuMjU5VjEuMDA0NjRIMjA2LjI3N1pNMjA3LjI0OCA1LjA5MTk0QzIwNy41NjUgNS40MDk2NSAyMDcuNzIgNS43OTYwNSAyMDcuNzIgNi4yNDI1NkMyMDcuNzIgNi42ODkwNyAyMDcuNTY1IDcuMDU4MzEgMjA3LjI0OCA3LjM2NzQzQzIwNi45MyA3LjY4NTE0IDIwNi41NTIgNy44Mzk3IDIwNi4xMTQgNy44Mzk3QzIwNS42NzYgNy44Mzk3IDIwNS4yOSA3LjY4NTE0IDIwNC45NzIgNy4zNjc0M0MyMDQuNjYzIDcuMDU4MzEgMjA0LjUwOCA2LjY4MDQ5IDIwNC41MDggNi4yNDI1NkMyMDQuNTA4IDUuODA0NjQgMjA0LjY2MyA1LjQwOTY1IDIwNC45NzIgNS4wOTE5NEMyMDUuMjgxIDQuNzgyODEgMjA1LjY2OCA0LjYyODI1IDIwNi4xMTQgNC42MjgyNUMyMDYuNTYxIDQuNjI4MjUgMjA2LjkzIDQuNzgyODEgMjA3LjI0OCA1LjA5MTk0Wk0yMDUuNjkzIDUuODIxODFDMjA1LjU4MiA1LjkzMzQ0IDIwNS41MyA2LjA3MDgzIDIwNS41MyA2LjI0MjU2QzIwNS41MyA2LjM5NzEzIDIwNS41OSA2LjUzNDUxIDIwNS42OTMgNi42NDYxNEMyMDUuODE0IDYuNzY2MzYgMjA1Ljk1MSA2LjgxNzg4IDIwNi4xMTQgNi44MTc4OEMyMDYuMjc3IDYuODE3ODggMjA2LjQwNiA2Ljc1Nzc3IDIwNi41MTggNi42NDYxNEMyMDYuNjM4IDYuNTI1OTMgMjA2LjY4OSA2LjM5NzEzIDIwNi42ODkgNi4yNDI1NkMyMDYuNjg5IDYuMDc5NDEgMjA2LjYyOSA1LjkzMzQ0IDIwNi41MTggNS44MjE4MUMyMDYuNDA2IDUuNzEwMTggMjA2LjI2OSA1LjY1ODY2IDIwNi4xMTQgNS42NTg2NkMyMDUuOTUxIDUuNjU4NjYgMjA1LjgwNSA1LjcxMDE4IDIwNS42OTMgNS44MjE4MVpNMjA4LjY0NyAyLjEyOTVIMjA5LjQ4OVYzLjA2NTQ2SDIwOC42NDdWNC4yNzYxOUgyMDcuNDk3VjEuMDY0NzVIMjA4LjY0N1YyLjEzODA5VjIuMTI5NVonIGZpbGw9J3doaXRlJy8+PHBhdGggZD0nTTIxNy4yNzcgNC43Mzk5MlY1LjcxODgxSDIxNC4wNjVWNi44ODY2MUgyMTIuOTkyVjUuNzE4ODFIMjA5Ljc4MVY0LjczOTkySDIxNy4yNzdaTTIxNS41OTQgMS4wNTYxOVYyLjA2MDg1SDIxMy43MzlMMjE2LjA5MiAzLjUyOTE4TDIxNS4zMjggNC4zMjc3NUwyMTMuNTU5IDMuMTQyNzhMMjExLjkyNyA0LjMzNjM0TDIxMS4wNjkgMy41MjkxOEwyMTMuMzI3IDIuMDUyMjZIMjExLjU5MlYxLjA0NzYxSDIxNS42MDJMMjE1LjU5NCAxLjA1NjE5WicgZmlsbD0nd2hpdGUnLz48cGF0aCBkPSdNMjE5LjM3NiA0Ljg4NjAxVjMuMjgwMjhIMjIwLjQ0OVY0Ljg4NjAxSDIyMi44NjJWNS44MTMzOEgyMTguMzJWNC44ODYwMUgyMTkuMzc2Wk0yMjIuNTk2IDEuMDczNDlWNC4zMDIxMUgyMjEuNTIzVjIuMDAwODZIMjE4LjU3OFYxLjA3MzQ5SDIyMi41OTZaTTIyNC45NjYgMi44MDgwMUgyMjUuODA4VjMuNzM1MzhIMjI0Ljk2NlY2Ljk5ODM1SDIyMy44MjRWMS4wNzM0OUgyMjQuOTY2VjIuODA4MDFaJyBmaWxsPSd3aGl0ZScvPjxwYXRoIGQ9J00yMjYuNzAzIDIuMDYwOTRWMS4wNjQ4OEgyMjguMjE0VjAuMzI2NDE2SDIyOS4yODhWMS4wNjQ4OEgyMzAuNzlWMi4wNjA5NEgyMzAuMjE1QzIzMC4yNSAyLjE2Mzk4IDIzMC4yNzUgMi4yNTg0NCAyMzAuMzAxIDIuMzYxNDhDMjMwLjMyNyAyLjQ3MzExIDIzMC4zMzUgMi41ODQ3MyAyMzAuMzM1IDIuNjk2MzZDMjMwLjMzNSAzLjEyNTcgMjMwLjE4MSAzLjUwMzUyIDIyOS44NzIgMy44MTI2NEMyMjkuNTYzIDQuMTMwMzUgMjI5LjE4NSA0LjI5MzUgMjI4Ljc0NyA0LjI5MzVDMjI4LjMwOSA0LjI5MzUgMjI3LjkyMyA0LjEzMDM1IDIyNy42MTMgMy44MTI2NEMyMjcuMzA0IDMuNTAzNTIgMjI3LjE1OCAzLjEyNTcgMjI3LjE1OCAyLjY5NjM2QzIyNy4xNTggMi41NzYxNSAyMjcuMTY3IDIuNDY0NTIgMjI3LjE4NCAyLjM2MTQ4QzIyNy4yMSAyLjI1ODQ0IDIyNy4yNDQgMi4xNjM5OCAyMjcuMjg3IDIuMDYwOTRIMjI2LjcxMkgyMjYuNzAzWk0yMzMuMTQzIDQuODA4NzFWNi45NTU0SDIzMi4wN1Y1Ljc0NDY2SDIyNi45NDRWNC44MTcyOUgyMzMuMTQzVjQuODA4NzFaTTIyOC4zNDMgMi4zMDk5NkMyMjguMjMyIDIuNDEzIDIyOC4xNzIgMi41NTAzOSAyMjguMTcyIDIuNzA0OTVDMjI4LjE3MiAyLjg1OTUxIDIyOC4yMjMgMi45ODgzMSAyMjguMzQzIDMuMDk5OTRDMjI4LjQ1NSAzLjIyMDE1IDIyOC41OTIgMy4yNzE2OCAyMjguNzU1IDMuMjcxNjhDMjI4LjkwMSAzLjI3MTY4IDIyOS4wMzkgMy4yMTE1NyAyMjkuMTUgMy4wOTk5NEMyMjkuMjYyIDIuOTg4MzEgMjI5LjMyMiAyLjg1OTUxIDIyOS4zMjIgMi43MDQ5NUMyMjkuMzIyIDIuNTUwMzkgMjI5LjI2MiAyLjQyMTU5IDIyOS4xNSAyLjMwOTk2QzIyOS4wMzkgMi4xOTgzMyAyMjguOTEgMi4xNDY4MSAyMjguNzU1IDIuMTQ2ODFDMjI4LjU5MiAyLjE0NjgxIDIyOC40NTUgMi4xOTgzMyAyMjguMzQzIDIuMzA5OTZaTTIzMy4xNDMgMi4xMjk2NEgyMzMuOTg1VjMuMDY1NTlIMjMzLjE0M1Y0LjI3NjMzSDIzMi4wN1YxLjA1NjI5SDIzMy4xNDNWMi4xMjk2NFonIGZpbGw9J3doaXRlJy8+PHBhdGggZD0nTTIzOC40MzkgMS4wNzM0OVY0LjI4NDk0SDIzNy4zNjVWMi4wMDA4NkgyMzQuNjk1VjEuMDczNDlIMjM4LjQzOVpNMjM5LjcwOSAxLjA3MzQ5SDI0MC43ODNWNi45Mjk2NkgyMzkuNzA5VjEuMDczNDlaJyBmaWxsPSd3aGl0ZScvPjxwYXRoIGQ9J00yNDkuMzAxIDMuMjExNDVWNC4xNDc0MUgyNDYuMDlWNC42NTQwM0gyNDguNzc3VjYuNjk3NjdIMjQzLjQyOFY3LjAyMzk3SDI0OC44MDNWNy44OTk4MkgyNDIuMzU0VjUuODIxODJIMjQ3LjcwNFY1LjUyMTI5SDI0Mi4zNTRWNC42NDU0NEgyNDUuMDI1VjQuMTM4ODJIMjQxLjgxMlYzLjIwMjg2SDI0OS4zMUwyNDkuMzAxIDMuMjExNDVaTTI0Ni45MTMgMi43MDQ4M0wyNDUuNTU3IDEuMjcwODRMMjQ0LjIgMi43MDQ4M0wyNDMuMzQyIDIuMDUyMjRMMjQ1LjM0MiAwSDI0NS43NzlMMjQ3Ljc4MSAyLjAzNTA2TDI0Ni45MjIgMi43MDQ4M0gyNDYuOTEzWicgZmlsbD0nd2hpdGUnLz48cGF0aCBkPSdNMjU0LjY3NiAzLjM1NzQyVjQuMjc2MjFIMjUyLjkzM1Y1LjEwMDU0SDI1MS44NlY0LjI3NjIxSDI1MC4xMTdWMy4zNTc0MkgyNTQuNjY4SDI1NC42NzZaTTI1MS42NTQgNi4wNDUwOEgyNTcuMTkyVjYuOTgxMDRIMjUwLjU4VjUuMjAzNThIMjUxLjY1NFY2LjA0NTA4Wk0yNTMuNTM0IDAuNDcyMjcyQzI1My44NTIgMC43ODk5ODIgMjU0LjAxNSAxLjE2NzggMjU0LjAxNSAxLjYxNDMxQzI1NC4wMTUgMi4wNjA4MiAyNTMuODUyIDIuNDMwMDUgMjUzLjUzNCAyLjc0Nzc2QzI1My4yMjUgMy4wNjU0NyAyNTIuODQ3IDMuMjIwMDQgMjUyLjQwMSAzLjIyMDA0QzI1MS45NTQgMy4yMjAwNCAyNTEuNTc2IDMuMDY1NDcgMjUxLjI1OSAyLjc0Nzc2QzI1MC45NSAyLjQzODY0IDI1MC43OTUgMi4wNjA4MiAyNTAuNzk1IDEuNjE0MzFDMjUwLjc5NSAxLjE2NzggMjUwLjk1IDAuNzg5OTgyIDI1MS4yNTkgMC40NzIyNzJDMjUxLjU3NiAwLjE2MzE0OCAyNTEuOTU0IDAgMjUyLjQwMSAwQzI1Mi44NDcgMCAyNTMuMjE3IDAuMTU0NTYyIDI1My41MzQgMC40NzIyNzJaTTI1MS45OCAxLjE3NjM5QzI1MS44NjggMS4yOTY2IDI1MS44MTcgMS40MzM5OSAyNTEuODE3IDEuNTk3MTRDMjUxLjgxNyAxLjc2MDI5IDI1MS44NzcgMS44ODkwOSAyNTEuOTggMi4wMDA3MkMyNTIuMSAyLjEyMDkzIDI1Mi4yMzggMi4xODEwNCAyNTIuNDAxIDIuMTgxMDRDMjUyLjU2NCAyLjE4MTA0IDI1Mi43MDEgMi4xMjA5MyAyNTIuODEzIDIuMDAwNzJDMjUyLjkzMyAxLjg4OTA5IDI1Mi45OTMgMS43NTE3IDI1Mi45OTMgMS41OTcxNEMyNTIuOTkzIDEuNDQyNTggMjUyLjkzMyAxLjI5NjYgMjUyLjgxMyAxLjE3NjM5QzI1Mi43MDEgMS4wNjQ3NiAyNTIuNTY0IDEuMDEzMjQgMjUyLjQwMSAxLjAxMzI0QzI1Mi4yMzggMS4wMTMyNCAyNTIuMSAxLjA2NDc2IDI1MS45OCAxLjE3NjM5Wk0yNTcuMTg0IDEuMDU2MTdWNC4yNjc2MkgyNTYuMTFWMy4wNTY4OUgyNTQuOTQyVjIuMTI5NTJIMjU2LjExVjEuMDU2MTdIMjU3LjE4NFonIGZpbGw9J3doaXRlJy8+PHBhdGggZD0nTTE3Ny40OTIgMTIuMjI3NUgxNzEuNTE2VjIwLjIzOUgxNzcuNDkyVjEyLjIyNzVaJyBmaWxsPSd3aGl0ZScvPjxwYXRoIGQ9J00xOTIuMDEzIDkuMzE2NjVIMTcxLjUxNlYxNC40MzQ0SDE5Mi4wMTNWOS4zMTY2NVonIGZpbGw9J3doaXRlJy8+PHBhdGggZD0nTTE5Mi4wMTMgMTguODEzN0gxNzEuNTE2VjI0LjAwMDFIMTkyLjAxM1YxOC44MTM3WicgZmlsbD0nd2hpdGUnLz48cGF0aCBkPSdNMTkyLjAxMyAwLjEzNzQ1MUgxNzEuNTE2VjUuMzQ5NjJIMTkyLjAxM1YwLjEzNzQ1MVonIGZpbGw9J3doaXRlJy8+PHBhdGggZD0nTTEzMy4wMDIgMC4wOTQ0ODI0TDEyNS45MTggMTAuMjA5N0wxMTguODM0IDAuMDk0NDgyNEgxMTIuODgzVjI0SDExOC44NTlWOS45MDA1NkwxMjUuODA2IDE4LjgwNUwxMjUuOTE4IDE4Ljk1MUwxMjYuMDI5IDE4LjgwNUwxMzIuOTc2IDkuOTAwNTZWMjRIMTM4Ljk1MlYwLjA5NDQ4MjRIMTMzLjAwMlonIGZpbGw9J3doaXRlJy8+PHBhdGggZD0nTTE0Mi42NDEgMTcuMjE2NlYxOC45ODU0VjI0LjAwMDFIMTQ4LjU5MVYxOC45ODU0VjEyLjk3NDdWMTAuNjk5MkwxNDIuNjQxIDE1LjE5MDFWMTcuMjE2NlonIGZpbGw9JyNEQjI2MTUnLz48cGF0aCBkPSdNMTY3LjQxMiA0LjczMTM3QzE2Ni45NzQgMy41NjM1NyAxNjYuMjc5IDIuNjE5MDMgMTY1LjMzNSAxLjg5Nzc0QzE2NC41MzYgMS4yNDUxNCAxNjMuNjE2IDAuNzcyODcyIDE2Mi41OTQgMC40ODA5MjJDMTYxLjkxNiAwLjI5MjAxMyAxNjEuMTg3IDAuMTcxNzk4IDE2MC40MDYgMC4xMzc0NTFIMTQyLjMxMkwxNDguNTcyIDguNzQ5OTdWNS4zNDk2MkgxNTEuODA5VjUuMzY2NzlIMTU0LjM4NUgxNTkuNDg2QzE1OS42MDYgNS4zNjY3OSAxNTkuODA0IDUuMzgzOTYgMTYwLjAwMSA1LjQxODMxQzE2MC4wMzYgNS40MTgzMSAxNjAuMDYxIDUuNDM1NDggMTYwLjA5NiA1LjQ0NDA3QzE2MC4xMDQgNS40NDQwNyAxNjAuMTIxIDUuNDQ0MDcgMTYwLjEzIDUuNDQ0MDdDMTYwLjM4OCA1LjUyMTM1IDE2MC42MTIgNS42MzI5OCAxNjAuODE4IDUuNzYxNzhDMTYwLjgzNSA1Ljc3MDM3IDE2MC44NDQgNS43Nzg5NSAxNjAuODYxIDUuNzk2MTNDMTYwLjkzOCA1Ljg0NzY1IDE2MS4wMDcgNS45MDc3NiAxNjEuMDc2IDUuOTc2NDVDMTYxLjA5MyA1Ljk5MzYyIDE2MS4xMTcgNi4wMTkzOCAxNjEuMTM1IDYuMDM2NTZDMTYxLjM0OSA2LjI1OTgxIDE2MS41MDUgNi41MjYgMTYxLjYxNyA2Ljg1MjNDMTYxLjY1MSA2Ljk2MzkzIDE2MS42ODQgNy4wNzU1NSAxNjEuNzEgNy4xOTU3N0MxNjEuNzE5IDcuMjQ3MjkgMTYxLjcyNyA3LjI5ODgxIDE2MS43NDQgNy4zNTAzM0MxNjEuNzQ0IDcuMzkzMjcgMTYxLjc2MiA3LjQyNzYxIDE2MS43NjIgNy40NzA1NUMxNjEuNzc5IDcuNTk5MzUgMTYxLjc4NyA3LjcyODE1IDE2MS43ODcgNy44NjU1NEMxNjEuNzYyIDkuMjMwODMgMTYxLjIwMyAxMC4wNzIzIDE2MC4xMjEgMTAuMzk4NkMxNjAuMTEzIDEwLjM5ODYgMTYwLjEwNCAxMC4zOTg2IDE2MC4wODcgMTAuMzk4NkMxNjAuMDQ0IDEwLjQwNzIgMTYwLjAwMiAxMC40MjQ0IDE1OS45NTEgMTAuNDMzQzE1OS43NzkgMTAuNDU4NyAxNTkuNjA2IDEwLjQ2NzMgMTU5LjUwMyAxMC40NzU5SDE1MS44MDlWMTAuNDkzMUgxNTAuNzE5TDE1MS4wNTUgMTAuOTY1NEwxNTQuNDIxIDE1LjY5NjdMMTYwLjMwMiAyMy45NzQzSDE2Ny41NjZMMTYxLjYxNyAxNS41Njc5QzE2MS45NTEgMTUuNTE2MyAxNjIuMjc3IDE1LjQzOTEgMTYyLjU5NCAxNS4zNTMyQzE2My41NjUgMTUuMDc4NCAxNjQuNDIzIDE0LjY0MDUgMTY1LjE4OCAxNC4wMzk0QzE2Ni40MjQgMTMuMTU1IDE2Ny4yNTcgMTEuOTA5OSAxNjcuNjYxIDEwLjMxMjhDMTY3LjY2MSAxMC4yOTU2IDE2Ny42NjEgMTAuMjg3IDE2Ny42NjkgMTAuMjY5OEMxNjcuNzA0IDEwLjE1ODIgMTY3LjcyOSAxMC4wNDY2IDE2Ny43NTUgOS45MjYzNkMxNjcuNzg5IDkuNzQ2MDQgMTY3LjgyNCA5LjU2NTcyIDE2Ny44NSA5LjM4NTM5QzE2Ny44NSA5LjM1OTYzIDE2Ny44NSA5LjMzMzg3IDE2Ny44NTggOS4yOTk1M0MxNjcuODg0IDkuMTAyMDMgMTY3LjkxIDguOTA0NTQgMTY3LjkyNyA4LjY5ODQ1QzE2Ny45MjcgOC42MTI1OSAxNjcuOTM2IDguNTI2NzIgMTY3Ljk0NSA4LjQ0MDg1QzE2Ny45NDUgOC4zODA3NCAxNjcuOTQ1IDguMzI5MjIgMTY3Ljk0NSA4LjI2OTExQzE2Ny45NDUgOC4xNDAzMSAxNjcuOTQ1IDguMDExNTEgMTY3Ljk0NSA3Ljg4MjcxQzE2Ny45MTkgNi42ODkxNSAxNjcuNzMgNS42MzI5OCAxNjcuMzg3IDQuNzE0MkwxNjcuNDEyIDQuNzMxMzdaJyBmaWxsPScjREIyNjE1Jy8+PHBhdGggZD0nTTI3Mi4wMjUgMTQuMDUxN0gyNjguMTdWMTIuNjgyNEgyNzMuNjAxVjE2Ljg1OTRDMjczLjYwMSAxNy43NTcgMjczLjUyMSAxOC44NjE2IDI3My4zNiAxOS41NTJIMjcxLjgxOEMyNzEuOTU2IDE4Ljg3MzEgMjcyLjAyNSAxNy45MTgxIDI3Mi4wMjUgMTcuMTAxMVYxNC4wNTE3Wk0yNzAuNzk0IDE2LjUwMjdWMjEuMTE3SDI3My4wMDNDMjczLjg1NCAyMS4xMTcgMjc0LjQ3NiAyMC45Nzg5IDI3NC44MSAyMC42Nzk3VjEyLjQ5ODNIMjc2LjM4NlYxNi42ODY4SDI3Ny45NzRWMTguMDQ0NkgyNzYuMzg2VjIzLjA2MTdIMjc0LjgxVjIxLjk2ODVDMjc0LjQ2NCAyMi4zMTM3IDI3My42OTMgMjIuNDg2MyAyNzIuNjkyIDIyLjQ4NjNIMjY3Ljg1OVYyMS4xMTdIMjY5LjIxN1YxNi41MDI3SDI3MC43OTRaJyBmaWxsPSd3aGl0ZScvPjxwYXRoIGQ9J00yNTguODI5IDEyLjE4NzZIMjYwLjQwNlYxMy4xODg3SDI2Mi41ODFWMTQuNDc3NUgyNTYuNTYyVjEzLjE4ODdIMjU4LjgyOVYxMi4xODc2Wk0yNTYuODczIDE2LjkyODVDMjU2Ljg3MyAxNi4yNzI2IDI1Ny4xODQgMTUuNzQzMyAyNTcuNjY3IDE1LjM4NjZDMjU4LjE1IDE1LjA0MTMgMjU4LjgxOCAxNC44NTcyIDI1OS41ODkgMTQuODU3MkMyNjAuMzcxIDE0Ljg1NzIgMjYxLjAzOSAxNS4wNDEzIDI2MS41MSAxNS4zODY2QzI2Mi4wMDUgMTUuNzQzMyAyNjIuMzE2IDE2LjI3MjYgMjYyLjMxNiAxNi45Mjg1QzI2Mi4zMTYgMTcuNTg0NCAyNjIuMDA1IDE4LjExMzcgMjYxLjUxIDE4LjQ3MDRDMjYxLjAzOSAxOC44MTU2IDI2MC4zNzEgMTguOTk5NyAyNTkuNTg5IDE4Ljk5OTdDMjU4LjgxOCAxOC45OTk3IDI1OC4xNSAxOC44MTU2IDI1Ny42NjcgMTguNDcwNEMyNTcuMTg0IDE4LjExMzcgMjU2Ljg3MyAxNy41ODQ0IDI1Ni44NzMgMTYuOTI4NVpNMjU4LjM5MiAxNi45Mjg1QzI1OC4zOTIgMTcuMTQ3MSAyNTguNDg0IDE3LjMzMTIgMjU4LjY1NyAxNy40NjkzQzI1OC44NzUgMTcuNjUzNCAyNTkuMjA5IDE3Ljc1NyAyNTkuNTg5IDE3Ljc1N0MyNTkuOTggMTcuNzU3IDI2MC4zMDIgMTcuNjUzNCAyNjAuNTMyIDE3LjQ2OTNDMjYwLjcwNSAxNy4zMzEyIDI2MC43OTcgMTcuMTQ3MSAyNjAuNzk3IDE2LjkyODVDMjYwLjc5NyAxNi43MDk5IDI2MC43MDUgMTYuNTI1NyAyNjAuNTMyIDE2LjM4NzdDMjYwLjMwMiAxNi4yMDM1IDI1OS45OCAxNi4xIDI1OS41ODkgMTYuMUMyNTkuMjA5IDE2LjEgMjU4Ljg3NSAxNi4yMDM1IDI1OC42NTcgMTYuMzg3N0MyNTguNDg0IDE2LjUyNTcgMjU4LjM5MiAxNi43MDk5IDI1OC4zOTIgMTYuOTI4NVpNMjY1LjM0MiAxNS4wNzU5SDI2Ni44MDRWMTYuNDQ1MkgyNjUuMzQyVjE4Ljk5OTdIMjYzLjc2NlYxMi40OTgzSDI2NS4zNDJWMTUuMDc1OVpNMjY1LjM0MiAxOS40OTQ1VjIzLjIzNDNIMjYzLjc2NlYyMC44NTI0SDI1Ny4zNjhWMTkuNDk0NUgyNjUuMzQyWicgZmlsbD0nd2hpdGUnLz48cGF0aCBkPSdNMjUyLjkyNSAxMy45ODI3SDI0Ni4yNjJWMTIuNjI0OUgyNTQuNTAxVjE0LjY5NjFDMjU0LjUwMSAxNS4zMTc1IDI1NC40NTUgMTUuOTYxOSAyNTQuMzE3IDE2LjQxMDdIMjUyLjc3NUMyNTIuODkgMTUuOTYxOSAyNTIuOTI1IDE1LjM0MDUgMjUyLjkyNSAxNC43NjUyVjEzLjk4MjdaTTI0NS42NjQgMTYuOTRIMjQ4LjcyNVYxNC45MTQ4SDI1MC4zMDFWMTYuOTRIMjU1LjE2OVYxOC4zMDkzSDI0NS42NjRWMTYuOTRaTTI0Ni4xNDcgMjEuMTk3NkMyNDYuMTQ3IDIwLjQyNjYgMjQ2LjUzOSAxOS45MjAzIDI0Ny4xOTQgMTkuNTg2NkMyNDcuOTg4IDE5LjE4MzggMjQ5LjE3NCAxOS4wNDU4IDI1MC40MjggMTkuMDQ1OEMyNTEuNjgyIDE5LjA0NTggMjUyLjg2NyAxOS4xODM4IDI1My42NjEgMTkuNTg2NkMyNTQuMzE3IDE5LjkyMDMgMjU0LjcwOSAyMC40MjY2IDI1NC43MDkgMjEuMTk3NkMyNTQuNzA5IDIxLjk4IDI1NC4zMTcgMjIuNDg2MyAyNTMuNjYxIDIyLjgyQzI1Mi44NjcgMjMuMjIyOCAyNTEuNjgyIDIzLjM2MDkgMjUwLjQyOCAyMy4zNjA5QzI0OS4xNzQgMjMuMzYwOSAyNDcuOTg4IDIzLjIyMjggMjQ3LjE5NCAyMi44MkMyNDYuNTM5IDIyLjQ4NjMgMjQ2LjE0NyAyMS45OCAyNDYuMTQ3IDIxLjE5NzZaTTI0Ny43MjQgMjEuMTk3NkMyNDcuNzI0IDIxLjQyNzcgMjQ3Ljg1IDIxLjYwMDMgMjQ4LjA5MiAyMS43MTU0QzI0OC41MjkgMjEuOTU3IDI0OS4zNDYgMjIuMDI2MSAyNTAuNDI4IDIyLjAyNjFDMjUxLjUxIDIyLjAyNjEgMjUyLjMyNyAyMS45NTcgMjUyLjc2NCAyMS43MTU0QzI1My4wMDUgMjEuNjAwMyAyNTMuMTMyIDIxLjQyNzcgMjUzLjEzMiAyMS4xOTc2QzI1My4xMzIgMjAuOTc4OSAyNTMuMDA1IDIwLjgwNjMgMjUyLjc2NCAyMC42Nzk3QzI1Mi4zMjcgMjAuNDQ5NiAyNTEuNTEgMjAuMzgwNiAyNTAuNDI4IDIwLjM4MDZDMjQ5LjM0NiAyMC4zODA2IDI0OC41MjkgMjAuNDQ5NiAyNDguMDkyIDIwLjY3OTdDMjQ3Ljg1IDIwLjgwNjMgMjQ3LjcyNCAyMC45Nzg5IDI0Ny43MjQgMjEuMTk3NlonIGZpbGw9J3doaXRlJy8+PHBhdGggZD0nTTI0MS44NzggMTMuNjE0NVYxMi42MjQ5SDI0My40NTVWMTcuMDU1MUgyNDAuMTI5VjE3LjgyNkgyNDQuMTIyVjE5LjE3MjNIMjM0LjYxN1YxNy44MjZIMjM4LjU2NFYxNy4wNTUxSDIzNS4yODVWMTIuNjI0OUgyMzYuODYxVjEzLjYxNDVIMjQxLjg3OFpNMjQxLjg3OCAxNC45Mzc4SDIzNi44NjFWMTUuNzMxOEgyNDEuODc4VjE0LjkzNzhaTTIzOS4zNTggMjEuODk5NUMyMzguMzggMjIuNzYyNSAyMzYuNjE5IDIzLjE0MjIgMjM0LjkwNSAyMy4xNzY3VjIxLjg0MTlDMjM2Ljg1IDIxLjc3MjkgMjM4LjU2NCAyMS4yMDkxIDIzOC41NjQgMTkuODk3M1YxOS43MjQ3SDI0MC4xNDFWMTkuODk3M0MyNDAuMTQxIDIxLjE5NzYgMjQxLjg1NSAyMS43MjY5IDI0My44MTEgMjEuODQxOVYyMy4xNzY3QzI0MS45ODIgMjMuMTY1MiAyNDAuMzI1IDIyLjczOTUgMjM5LjM1OCAyMS44OTk1WicgZmlsbD0nd2hpdGUnLz48cGF0aCBkPSdNMjI1LjczMSAxOC4xMzY3SDIzMi42VjE5LjUwNjFIMjI5LjEyNVYyMS4yMzIxSDIzMy4wMTVWMjIuNjAxNEgyMjMuNjI1VjIxLjIzMjFIMjI3LjU2VjE5LjUwNjFIMjI0LjE1NFYxNS4zODY2SDIzMC45MDlWMTQuMDUxOEgyMjQuMTU0VjEyLjY4MjRIMjMyLjQ4NVYxNi43NTU5SDIyNS43MzFWMTguMTM2N1onIGZpbGw9J3doaXRlJy8+PHBhdGggZD0nTTIxNS4zMTIgMTYuMDg4NVYxNC4wNTE3SDIxMi44NjFWMTIuNjgyNEgyMTYuODU0VjE0Ljg0NTdIMjE3LjkwMVYxMi41NTU4SDIxOS40MlYyMi45MzUxSDIxNy45MDFWMTkuNzkzN0gyMTYuMTI5QzIxNS40OTYgMjEuMDk0IDIxNC40MDMgMjIuMDE0NSAyMTIuNzM0IDIyLjQ4NjNWMjAuOTU1OUMyMTQuMzQ1IDIwLjM5MjEgMjE1LjMxMiAxOC44MDQxIDIxNS4zMTIgMTYuMDg4NVpNMjE2Ljg1NCAxNi4yMTVDMjE2Ljg0MiAxNy4wMjA1IDIxNi43NjIgMTcuNzU3IDIxNi42MDEgMTguNDM1OUgyMTcuOTAxVjE2LjIxNUgyMTYuODU0Wk0yMjAuMDQxIDIzLjA2MTdWMTIuNDk4M0gyMjEuNTcyVjIzLjA2MTdIMjIwLjA0MVonIGZpbGw9J3doaXRlJy8+PHBhdGggZD0nTTIwNS41MDYgMTQuOTk1M1YxNC4wNTE3SDIwMS44N1YxMi42ODI0SDIwNy4wNzFWMTQuODExMkMyMDcuMDcxIDE4LjgyNzEgMjA1LjYyMSAyMS41MzEyIDIwMS44MTIgMjIuNDg2M1YyMC45OTA0QzIwNC40MzYgMjAuMjMxIDIwNS41MDYgMTguMTEzNyAyMDUuNTA2IDE0Ljk5NTNaTTIxMC4zNzQgMjMuMDYxN0gyMDguNzk3VjEyLjQ5ODNIMjEwLjM3NFYyMy4wNjE3WicgZmlsbD0nd2hpdGUnLz48L2c+PGRlZnM+PGNsaXBQYXRoIGlkPSdjbGlwMF80MTI2Xzg1NjQnPjxyZWN0IHdpZHRoPScyNzgnIGhlaWdodD0nMjQnIGZpbGw9J3doaXRlJy8+PC9jbGlwUGF0aD48L2RlZnM+PC9zdmc+" alt="GIST 기계로봇공학과">
  <div class="rule"></div>
  <h1 id="title"></h1>
  <p data-t="lead"></p>
  <div class="chips">
    <span class="chip"><span data-t="deadline"></span> <b id="dl">-</b></span>
    <span class="chip"><span data-t="period"></span> <b id="range">-</b> <span data-t="acceptBasis"></span></span>
  </div>
</div></div>

<main id="app">
  <div class="notice" id="notice"></div>

  <section class="card gate" id="gate" hidden>
    <div class="ico-l">✉</div>
    <h2 data-t="gateT"></h2>
    <p class="desc" data-t="gateD"></p>
    <div class="f" id="gEmailBox"><label data-t="gEmail"></label>
      <div class="row2"><input id="gEmail" type="email" autocomplete="email" placeholder="id@gm.gist.ac.kr"><button class="btn red" type="button" id="gSend" onclick="sendCode()" data-t="gSend"></button></div>
      <div class="err" id="gErr1"></div></div>
    <div class="f" id="gCodeBox" hidden style="margin-top:14px"><label data-t="gCode"></label>
      <div class="row2"><input id="gCode" class="code-in" inputmode="numeric" autocomplete="one-time-code" maxlength="6" placeholder="······"><button class="btn red" type="button" id="gVerify" onclick="checkCode()" data-t="gVerify"></button></div>
      <div class="hint" id="gHint"></div><div class="err" id="gErr2"></div></div>
    <p class="gnote" data-t="gNote"></p>
  </section>

  <div class="sess" id="sess" hidden><div class="who"><span class="dot"></span><b id="sEmail"></b><span class="sub2" data-t="verified"></span></div><button class="linkbtn" type="button" onclick="logout()" data-t="logout"></button></div>
  <div class="tabs" id="tabs" hidden><button type="button" data-tab="form" class="on" onclick="showTab('form')" data-t="tabForm"></button><button type="button" data-tab="mine" onclick="showTab('mine')" data-t="tabMine"></button></div>
  <div id="mineArea" hidden></div>
  <div id="formArea" hidden>

  <section class="card">
    <h2><span class="num">1</span><span data-t="s1"></span></h2>
    <p class="desc" data-t="s1d"></p>
    <div class="grid">
      <div class="f" data-k="학번"><label><span data-t="sid"></span><span class="req">*</span></label><input id="학번" inputmode="numeric" maxlength="8" data-ph="sidPh"><div class="err" data-t="sidErr"></div></div>
      <div class="f" data-k="성명"><label><span data-t="name"></span><span class="req">*</span></label><input id="성명" data-ph="namePh"><div class="err" data-t="nameErr"></div></div>
      <div class="f" data-k="영문명"><label><span data-t="ename"></span><span class="req">*</span></label><input id="영문명" autocomplete="off" data-ph="enamePh"><div class="hint" data-t="enameHint"></div><div class="err" data-t="enameErr"></div></div>
      <div class="f" data-k="과정"><label><span data-t="course"></span><span class="req">*</span></label>
        <div class="seg two"><input type="radio" name="과정" id="c1" value="박사"><label for="c1" data-t="phd"></label><input type="radio" name="과정" id="c2" value="석박통합"><label for="c2" data-t="integ"></label></div>
        <div class="err" data-t="courseErr"></div></div>
      <div class="f"><label><span data-t="sem"></span><span class="opt" data-t="semOpt"></span></label><input id="재학학기" inputmode="numeric" data-ph="semPh"></div>
      <div class="f" data-k="지도교수"><label><span data-t="adv"></span><span class="req">*</span></label><select id="지도교수Sel"></select><input id="지도교수" data-ph="advPh" style="margin-top:8px;display:none"><div class="err" data-t="advErr"></div></div>
      <div class="f full"><label><span data-t="email"></span> <span class="opt" data-t="emailVer"></span></label><input id="vEmail" readonly></div>
    </div>
  </section>

  <section class="card">
    <h2><span class="num">2</span><span data-t="s2"></span></h2>
    <p class="desc" data-t="s2d"></p>
    <div id="papers"></div>
    <button class="add" type="button" onclick="addPaper()" data-t="addPaper"></button>
  </section>

  <section class="card">
    <h2><span class="num">3</span><span data-t="s3"></span></h2>
    <p class="desc" data-t="s3d"></p>
    <div id="patents"></div>
    <button class="add" type="button" onclick="addPatent()" data-t="addPatent"></button>
  </section>

  <section class="card">
    <h2><span class="num">4</span><span data-t="s4"></span></h2>
    <p class="desc" data-t="s4d"></p>
    <div id="awards"></div>
    <button class="add" type="button" onclick="addAward()" data-t="addAward"></button>
  </section>

  <section class="card">
    <h2><span class="num">5</span><span data-t="s5"></span></h2>
    <div class="f" style="margin-bottom:14px"><label><span data-t="memo"></span><span class="opt" data-t="optional"></span></label><textarea id="메모" rows="2" data-ph="memoPh"></textarea></div>
    <label class="check" id="agreeBox"><input type="checkbox" id="agree"><span data-t="agree"></span></label>
  </section>
  </div>
  <footer id="contact"></footer>
  <a class="adminlink" href="#admin">관리자</a>
</main>

<section id="admin" hidden>
  <div class="card a-login" id="aLogin">
    <h2><span class="num">⚙</span>담당자 화면</h2>
    <p class="desc">학생 제출 목록과 증빙 PDF를 봅니다. 담당자 비밀번호를 넣어 주세요.</p>
    <div class="f" id="aKeyBox"><input id="aKey" type="password" placeholder="관리자 비밀번호" autocomplete="current-password"><div class="err" id="aErr"></div></div>
    <div style="display:flex;gap:8px;margin-top:14px"><button class="btn red" type="button" onclick="adminLogin()">들어가기</button><a class="btn" href="#">← 제출 화면</a></div>
  </div>
  <div id="aMain" hidden>
    <div class="a-stats" id="aStats"></div>
    <div class="card" style="padding:16px 18px">
      <div class="a-top">
        <div class="a-tools">
          <input type="search" id="aQ" placeholder="이름·학번·지도교수·저널 검색" oninput="renderAdmin()">
          <label class="check" style="font-size:13px"><input type="checkbox" id="aOld" onchange="renderAdmin()" style="width:16px;height:16px">이전 제출(대체됨)도 보기</label>
        </div>
        <div class="a-tools">
          <button class="btn" type="button" onclick="loadAdmin()">↻ 새로고침</button>
          <button class="btn" type="button" onclick="adminCsv()">⬇ 목록 CSV</button>
          <a class="btn" id="aSheet" target="_blank" rel="noopener">구글 시트</a>
          <button class="btn" type="button" onclick="adminLogout()">나가기</button>
        </div>
      </div>
      <div class="sub" id="aNow" style="font-size:12px;color:var(--mute)"></div>
    </div>
    <div class="a-list" id="aList"></div>
  </div>
</section>

<div class="viewer" id="viewer">
  <div class="vbar"><b id="vName"></b><a class="btn sm" id="vOpen" target="_blank" rel="noopener">새 창</a><a class="btn sm" id="vDown">내려받기</a><button class="btn sm" type="button" onclick="closeViewer()">닫기 ✕</button></div>
  <div class="vbody" id="vBody"></div>
</div>

<div class="bar"><div class="in">
  <div class="sum"><span data-t="cntP"></span> <b id="sP">0</b> · <span data-t="cntT"></span> <b id="sT">0</b> · <span data-t="cntA"></span> <b id="sA">0</b></div>
  <button class="go" id="go" onclick="send()" data-t="submit"></button>
</div></div>

<div class="modal" id="modal"><div class="box" id="box"></div></div>

<template id="tPaper">
  <div class="item">
    <div class="item-h"><b><span data-t="paper"></span> <span class="n"></span></b><button class="del" type="button" data-t="del"></button></div>
    <div class="grid">
      <div class="f full" data-k="journal"><label><span data-t="journal"></span><span class="req">*</span></label><input name="journal" data-ph="journalPh"><div class="hint" data-t="journalHint"></div><div class="err" data-t="journalErr"></div></div>
      <div class="f full"><label data-t="ptitle"></label><input name="title" data-ph="ptitlePh"></div>
      <div class="f" data-k="doi"><label>DOI<span class="opt" data-t="ifAny"></span></label><input name="doi" placeholder="10.xxxx/xxxxx"><div class="err" data-t="doiErr"></div></div>
      <div class="f" data-k="accept"><label><span data-t="accept"></span><span class="req">*</span> <span class="opt" data-t="acceptOpt"></span></label><input name="accept" type="date"><div class="err" data-t="acceptErr"></div></div>
      <div class="f full" data-k="role"><label><span data-t="role"></span><span class="req">*</span></label>
        <div class="seg">
          <input type="radio" value="FA"><label>FA<small data-t="rFA"></small></label>
          <input type="radio" value="co-FA"><label>co-FA<small data-t="rCoFA"></small></label>
          <input type="radio" value="CA"><label>CA<small data-t="rCA"></small></label>
          <input type="radio" value="ETC"><label>ETC<small data-t="rETC"></small></label>
        </div><div class="err" data-t="roleErr"></div></div>
      <div class="f" data-k="authors"><label><span data-t="authors"></span><span class="req">*</span></label><input name="authors" type="number" min="1" inputmode="numeric" data-ph="authorsPh"><div class="err" data-t="authorsErr"></div></div>
      <div class="f full" data-k="file1"><label><span data-t="file1"></span><span class="req">*</span></label>
        <label class="file"><input type="file" name="file1" accept=".pdf,.jpg,.jpeg,.png"><span class="ico">PDF</span><span class="t" data-t="file1Ph"></span></label><div class="err" data-t="fileErr"></div></div>
      <div class="f full"><label><span data-t="file2"></span><span class="opt" data-t="file2Opt"></span></label>
        <label class="file"><input type="file" name="file2" accept=".pdf,.jpg,.jpeg,.png"><span class="ico">＋</span><span class="t" data-t="choose"></span></label></div>
    </div>
  </div>
</template>

<template id="tPatent">
  <div class="item">
    <div class="item-h"><b><span data-t="patent"></span> <span class="n"></span></b><button class="del" type="button" data-t="del"></button></div>
    <div class="grid">
      <div class="f full" data-k="type"><label><span data-t="ptype"></span><span class="req">*</span></label>
        <div class="seg">
          <input type="radio" value="국내출원"><label data-t="kApp"></label>
          <input type="radio" value="국내등록"><label data-t="kReg"></label>
          <input type="radio" value="국제출원"><label data-t="iApp"></label>
          <input type="radio" value="국제등록"><label data-t="iReg"></label>
        </div><div class="err" data-t="typeErr"></div></div>
      <div class="f full"><label data-t="pname"></label><input name="name"></div>
      <div class="f"><label data-t="pno"></label><input name="no" placeholder="10-2025-0000000"></div>
      <div class="f"><label data-t="pdate"></label><input name="date" type="date"></div>
      <div class="f" data-k="share"><label><span data-t="share"></span><span class="req">*</span></label><input name="share" type="number" min="1" max="100" inputmode="numeric" data-ph="sharePh"><div class="err" data-t="shareErr"></div></div>
      <div class="f" data-k="file"><label><span data-t="pfile"></span><span class="req">*</span></label>
        <label class="file"><input type="file" name="file" accept=".pdf,.jpg,.jpeg,.png"><span class="ico">PDF</span><span class="t" data-t="choose"></span></label><div class="err" data-t="fileErr"></div></div>
    </div>
  </div>
</template>

<template id="tAward">
  <div class="item">
    <div class="item-h"><b><span data-t="award"></span> <span class="n"></span></b><button class="del" type="button" data-t="del"></button></div>
    <div class="grid">
      <div class="f full" data-k="type"><label><span data-t="ptype"></span><span class="req">*</span></label>
        <div class="seg two">
          <input type="radio" value="국내"><label data-t="aDom"></label>
          <input type="radio" value="해외"><label data-t="aInt"></label>
        </div><div class="err" data-t="typeErr"></div></div>
      <div class="f"><label data-t="aevent"></label><input name="event"></div>
      <div class="f"><label data-t="aprize"></label><input name="prize" data-ph="aprizePh"></div>
      <div class="f"><label data-t="adate"></label><input name="date" type="date"></div>
      <div class="f" data-k="file"><label><span data-t="afile"></span><span class="req">*</span></label>
        <label class="file"><input type="file" name="file" accept=".pdf,.jpg,.jpeg,.png"><span class="ico">PDF</span><span class="t" data-t="choose"></span></label><div class="err" data-t="fileErr"></div></div>
    </div>
  </div>
</template>

<script>
/* ▼ 담당자 설정: Apps Script 웹앱 주소(…/exec)를 따옴표 안에 넣으면 제출이 켜집니다 ▼ */
const API = '';
/* ▲ ------------------------------------------------------------------------- ▲ */
const CFG_TPL = <?!= cfg ?>;
let CFG = CFG_TPL || {title:'2026학년도 박사과정 재학생 연구우수장학금',titleEn:'2026 Ph.D. Student Research Excellence Scholarship',lang:'',deadline:'2026-12-04 23:59',closed:false,from:'',to:'',profs:[],
  contact:'기계로봇공학과 류지나 (T.2383)',contactEn:'Jina Ryu, Dept. of Mechanical and Robotics Engineering (ext. 2383)',
  notice:'논문 첫 페이지(저자·소속이 보이는 쪽)나 논문 PDF를 꼭 올려 주세요. 아직 게재 전이면 Accept 메일을 함께 올려 주세요.',
  noticeEn:'Please upload the first page of each paper (showing authors and affiliations) or the full PDF. If it is not yet published, also upload the acceptance email.',maxMb:15};
const PREVIEW = !CFG_TPL && !API;
const I18N = {
 ko:{brand:'GIST · 기계로봇공학과',lead:'논문·특허·수상 실적을 입력하고 증빙을 올려 주세요. 담당자가 증빙과 JCR로 확인합니다.',deadline:'접수마감',period:'실적인정',acceptBasis:'(Accept 기준)',
  s1:'기본 정보',s1d:'같은 학번으로 다시 제출하면 가장 마지막 제출만 반영됩니다.',sid:'학번',sidPh:'예) 20244040',sidErr:'숫자 8자리로 입력해 주세요.',name:'성명',namePh:'홍길동 (외국인은 여권 영문명)',nameErr:'성명을 입력해 주세요.',
  ename:'논문 저자 표기 이름',enamePh:'예) Gildong Hong',enameHint:'논문 저자 목록에 적힌 그대로 써 주세요. 증빙에서 본인을 찾을 때 씁니다.',enameErr:'논문에 쓰는 영문 이름을 입력해 주세요.',course:'과정',phd:'박사',integ:'석박통합',courseErr:'과정을 골라 주세요.',sem:'재학학기',semOpt:'(이번 학기 기준)',semPh:'예) 4',adv:'지도교수',advPh:'지도교수 성함',advErr:'지도교수를 입력해 주세요.',email:'이메일',optional:'(선택)',
  s2:'SCI 논문',s2d:'Accept 이후 논문만 넣어 주세요. 저널 순위·IF는 담당자가 JCR에서 찾아 입력합니다.',addPaper:'＋ 논문 추가',
  s3:'특허',s3d:'출원·등록 모두 가능합니다. 기여도는 발명자 지분(%)을 적어 주세요.',addPatent:'＋ 특허 추가',
  s4:'학회·경진대회 수상',s4d:'국내·해외(국제) 학회와 경진대회 수상 실적을 넣어 주세요.',addAward:'＋ 수상 추가',
  s5:'확인',memo:'담당자에게 남길 말',memoPh:'예) 논문 2는 아직 게재 전이라 Accept 메일을 함께 올렸습니다.',agree:'입력한 실적과 증빙이 사실이며, 증빙(논문 PDF 등)이 장학금 심사에 쓰이는 것에 동의합니다.',
  cntP:'논문',cntT:'특허',cntA:'수상',submit:'제출하기',paper:'논문',patent:'특허',award:'수상',del:'삭제',
  journal:'저널명 (전체 이름)',journalPh:'예) Advanced Functional Materials',journalHint:'약어 대신 전체 이름으로 적어 주세요.',journalErr:'저널명을 입력해 주세요.',ptitle:'논문 제목',ptitlePh:'논문 제목',ifAny:'(있으면)',
  accept:'Accept 날짜',acceptOpt:'(일자가 불명확한 경우 1일로)',acceptErr:'Accept 날짜를 입력해 주세요.',acceptFuture:'오늘 이후 날짜는 넣을 수 없습니다.',doiErr:'DOI 형식이 아닙니다. 예) 10.1016/j.ijmecsci.2026.110123',authorsMin2:'공동 제1저자·그 외 공저자는 저자가 2명 이상이어야 합니다.',dupPaper:'같은 논문을 두 번 넣었습니다.',journalAbbr:'약어 말고 저널 전체 이름으로 써 주세요. 예) Int. J. Mech. Sci. → International Journal of Mechanical Sciences',chkTitle:'제출 전에 한 번 더 확인해 주세요',chkOut:'논문 {i}: Accept 날짜({d})가 실적인정 기간({r}) 밖입니다.',chkGo:'이대로 제출',chkFix:'고치러 가기',upPar:'증빙 올리는 중 ({i}/{n})',shrink:'사진 크기 줄이는 중…',role:'나의 역할',rFA:'제1저자',rCoFA:'공동 제1저자',rCA:'교신저자',rETC:'그 외 공저자',roleErr:'역할을 골라 주세요.',
  authors:'총 저자 수 (본인 포함)',authorsPh:'예) 5',authorsErr:'저자 수를 입력해 주세요.',file1:'증빙 ① 논문 첫 페이지 또는 PDF',file1Ph:'저자·소속이 보이는 첫 페이지를 올려 주세요',fileErr:'증빙 파일을 올려 주세요.',
  file2:'증빙 ② Accept 메일 등',file2Opt:'(게재 전이면)',choose:'파일 선택',
  ptype:'구분',kApp:'국내 출원',kReg:'국내 등록',iApp:'국제 출원',iReg:'국제 등록',typeErr:'구분을 골라 주세요.',pname:'발명의 명칭',pno:'출원·등록번호',pdate:'출원·등록일',share:'나의 기여도(%)',sharePh:'예) 30',shareErr:'1~100 사이로 입력해 주세요.',pfile:'증빙 (출원서·등록증, 기여도 표시)',
  aDom:'국내 학회·대회',aInt:'해외(국제) 학회·대회',aevent:'학회·대회명',aprize:'수상명',aprizePh:'예) Best Paper Award',adate:'수상일',afile:'증빙 (상장 등)',
  contact:'문의',closed:'접수가 마감되었습니다',emptyP:'논문 실적이 없으면 비워 두세요.',emptyT:'특허 실적이 없으면 비워 두세요.',emptyA:'수상 실적이 없으면 비워 두세요.',
  selProf:'선택해 주세요',etcProf:'목록에 없음 (직접 입력)',badType:'PDF·JPG·PNG 파일만 올릴 수 있습니다.',tooBig:'파일이 {mb}MB를 넘습니다. 첫 페이지만 따로 저장해 올려 주세요.',needOne:'논문·특허·수상 중 하나 이상 입력해 주세요.',
  sending:'제출하는 중',uploading:'증빙 파일을 올리고 있습니다…',upN:'증빙 파일 {i} / {n} 올리는 중',saving:'접수 기록 중…',noClose:'창을 닫지 마세요.',
  done:'제출되었습니다',rcpt:'접수번호',doneCnt:'논문 {p}편 · 특허 {t}건 · 수상 {a}건',doneNote:'{time} 접수 · 고칠 것이 있으면 같은 학번으로 다시 제출하세요(마지막 제출만 반영).',fail:'제출하지 못했습니다',back:'돌아가기',previewBanner:'미리보기 화면입니다 — 담당자가 접수를 열면 제출할 수 있습니다.',previewErr:'아직 접수가 열리지 않았습니다.',
  gateT:'GIST 메일로 본인 확인',gateD:'제출과 조회는 본인만 할 수 있도록, 학교 메일로 받은 인증번호를 입력해 주세요.',gEmail:'GIST 메일 주소',gSend:'인증번호 받기',gCode:'메일로 받은 6자리 인증번호',gVerify:'확인',
  gNote:'@gist.ac.kr 또는 @gm.gist.ac.kr 주소만 사용할 수 있습니다. 메일이 안 오면 스팸함을 확인해 주세요. 인증은 6시간 동안 유지됩니다.',gSent:'{email} 로 인증번호를 보냈습니다. (10분 안에 입력)',gResend:'다시 보내기',gBadEmail:'GIST 메일 주소를 입력해 주세요.',
  verified:'인증됨',logout:'로그아웃',tabForm:'실적 제출',tabMine:'내 제출 내역',emailVer:'(본인 확인된 메일)',
  mineEmpty:'아직 제출한 내역이 없습니다.',mineLoad:'불러오는 중…',mineLatest:'현재 반영되는 제출',mineOld:'이전 제출 (대체됨)',mineEdit:'이 내용으로 수정하기',mineHint:'수정하려면 내용을 불러와 고친 뒤 다시 제출하세요. 마지막 제출만 반영됩니다.',
  keepFile:'이전에 낸 파일 사용 (바꾸려면 새 파일 선택)',view:'보기',none:'없음',loadedEdit:'이전 제출 내용을 불러왔습니다. 고친 뒤 다시 제출하세요.',sessionExp:'본인 확인이 만료되었습니다. 다시 인증해 주세요.'},
 en:{brand:'GIST · Dept. of Mechanical and Robotics Engineering',lead:'Enter your papers, patents and awards and upload the evidence. The office verifies them against the documents and JCR.',deadline:'Deadline',period:'Eligible period',acceptBasis:'(by acceptance date)',
  s1:'Basic information',s1d:'If you submit again with the same student ID, only your latest submission counts.',sid:'Student ID',sidPh:'e.g. 20244040',sidErr:'Enter your 8-digit student ID.',name:'Name',namePh:'Name as in your passport',nameErr:'Enter your name.',
  ename:'Name as it appears in your papers',enamePh:'e.g. Gildong Hong',enameHint:'Write it exactly as in the author list. Used to find you on the evidence.',enameErr:'Please enter the name used in your papers.',course:'Program',phd:'Ph.D.',integ:'Integrated M.S.–Ph.D.',courseErr:'Choose your program.',sem:'Semester',semOpt:'(current semester)',semPh:'e.g. 4',adv:'Advisor',advPh:"Advisor's name",advErr:'Enter your advisor.',email:'Email',optional:'(optional)',
  s2:'SCI papers',s2d:'Only papers accepted for publication. You do not need to enter the JCR rank or IF — the office looks them up.',addPaper:'＋ Add a paper',
  s3:'Patents',s3d:'Both applications and registrations count. Enter your inventor share (%).',addPatent:'＋ Add a patent',
  s4:'Conference / competition awards',s4d:'Awards from domestic or international conferences and competitions.',addAward:'＋ Add an award',
  s5:'Confirmation',memo:'Note to the office',memoPh:'e.g. Paper 2 is not yet published, so I attached the acceptance email.',agree:'I confirm that the information and evidence are true, and agree that the evidence (paper PDFs etc.) will be used for the scholarship review.',
  cntP:'Papers',cntT:'Patents',cntA:'Awards',submit:'Submit',paper:'Paper',patent:'Patent',award:'Award',del:'Remove',
  journal:'Journal (full title)',journalPh:'e.g. Advanced Functional Materials',journalHint:'Use the full journal title, not the abbreviation.',journalErr:'Enter the journal title.',ptitle:'Paper title',ptitlePh:'Paper title',ifAny:'(if any)',
  accept:'Acceptance date',acceptOpt:'(if the day is unknown, use the 1st)',acceptErr:'Enter the acceptance date.',acceptFuture:'The date cannot be in the future.',doiErr:'This is not a DOI. e.g. 10.1016/j.ijmecsci.2026.110123',authorsMin2:'A co-first author or other co-author means at least 2 authors.',dupPaper:'The same paper is entered twice.',journalAbbr:'Please write the full journal title, not the abbreviation. e.g. Int. J. Mech. Sci. → International Journal of Mechanical Sciences',chkTitle:'Please check before submitting',chkOut:'Paper {i}: the acceptance date ({d}) is outside the eligible period ({r}).',chkGo:'Submit anyway',chkFix:'Go back and fix',upPar:'Uploading evidence ({i}/{n})',shrink:'Reducing photo size…',role:'Your role',rFA:'First author',rCoFA:'Co-first author',rCA:'Corresponding',rETC:'Other co-author',roleErr:'Choose your role.',
  authors:'Total number of authors (incl. you)',authorsPh:'e.g. 5',authorsErr:'Enter the number of authors.',file1:'Evidence ① First page of the paper or full PDF',file1Ph:'Upload the first page showing authors and affiliations',fileErr:'Upload the evidence file.',
  file2:'Evidence ② Acceptance email etc.',file2Opt:'(if not yet published)',choose:'Choose file',
  ptype:'Type',kApp:'Korean application',kReg:'Korean registration',iApp:'International application',iReg:'International registration',typeErr:'Choose the type.',pname:'Title of invention',pno:'Application / registration no.',pdate:'Filing / registration date',share:'Your share (%)',sharePh:'e.g. 30',shareErr:'Enter a number from 1 to 100.',pfile:'Evidence (application / certificate showing shares)',
  aDom:'Domestic conference / competition',aInt:'International conference / competition',aevent:'Conference / competition',aprize:'Award title',aprizePh:'e.g. Best Paper Award',adate:'Award date',afile:'Evidence (certificate etc.)',
  contact:'Contact',closed:'Submissions are closed',emptyP:'Leave empty if you have no papers.',emptyT:'Leave empty if you have no patents.',emptyA:'Leave empty if you have no awards.',
  selProf:'Select',etcProf:'Not in the list (type it)',badType:'Only PDF, JPG or PNG files can be uploaded.',tooBig:'The file is larger than {mb}MB. Please save and upload only the first page.',needOne:'Enter at least one paper, patent or award.',
  sending:'Submitting',uploading:'Uploading evidence files…',upN:'Uploading file {i} of {n}',saving:'Saving your submission…',noClose:'Please do not close this window.',
  done:'Submitted',rcpt:'Receipt no.',doneCnt:'{p} paper(s) · {t} patent(s) · {a} award(s)',doneNote:'Received {time}. To make changes, submit again with the same student ID (only the latest counts).',fail:'Submission failed',back:'Back',previewBanner:'Preview — submissions open once the office connects the form.',previewErr:'Submissions are not open yet.',
  gateT:'Verify with your GIST email',gateD:'So that only you can submit and view your records, enter the code sent to your school email.',gEmail:'GIST email address',gSend:'Send code',gCode:'6-digit code from the email',gVerify:'Verify',
  gNote:'Only @gist.ac.kr or @gm.gist.ac.kr addresses are accepted. Check your spam folder if the email does not arrive. Verification lasts 6 hours.',gSent:'A code was sent to {email} (valid for 10 minutes).',gResend:'Resend',gBadEmail:'Please enter your GIST email address.',
  verified:'verified',logout:'Sign out',tabForm:'Submit',tabMine:'My submissions',emailVer:'(verified)',
  mineEmpty:'You have no submissions yet.',mineLoad:'Loading…',mineLatest:'Current submission',mineOld:'Earlier submissions (replaced)',mineEdit:'Edit this submission',mineHint:'To make changes, load your submission, edit it and submit again. Only the latest one counts.',
  keepFile:'Using your previously uploaded file (choose a new one to replace)',view:'View',none:'None',loadedEdit:'Your previous submission has been loaded. Edit and submit again.',sessionExp:'Your session has expired. Please verify again.'}
};
let LANG = CFG.lang || new URLSearchParams(location.search).get('lang') || ((navigator.language||'ko').toLowerCase().startsWith('ko') ? 'ko' : 'en');
if (LANG !== 'en') LANG = 'ko';
const T = (k, v={}) => String(I18N[LANG][k] ?? I18N.ko[k] ?? k).replace(/\\{(\\w+)\\}/g, (_,x)=>v[x] ?? '');
const $ = s => document.querySelector(s);
let uid = 0;

function translate(root){
  root.querySelectorAll('[data-t]').forEach(e => e.textContent = T(e.dataset.t));
  root.querySelectorAll('[data-ph]').forEach(e => e.placeholder = T(e.dataset.ph));
  root.querySelectorAll('.file').forEach(b => { if (b.classList.contains('has')) { const f = b.querySelector('input').files[0]; if (f) b.querySelector('.t').textContent = f.name; } });
  root.querySelectorAll('.file.keep').forEach(b => { b.querySelector('.t').textContent = T('keepFile'); });
  if (typeof MINE !== 'undefined' && MINE && TAB === 'mine' && root === document) renderMine();
}
function setLang(l){
  LANG = l; document.documentElement.lang = l;
  $('#lk').classList.toggle('on', l==='ko'); $('#le').classList.toggle('on', l==='en');
  translate(document);
  $('#title').textContent = (l==='en' && CFG.titleEn) || CFG.title;
  const n = PREVIEW ? T('previewBanner') : ((l==='en' && CFG.noticeEn) || CFG.notice); $('#notice').textContent = n; $('#notice').style.display = n ? '' : 'none';
  const c = (l==='en' && CFG.contactEn) || CFG.contact; $('#contact').textContent = c ? T('contact') + ': ' + c + (CFG.contactEmail ? ' · ' + CFG.contactEmail : '') : '';
  if (CFG._dlKo) $('#dl').textContent = l==='en' ? CFG._dlEn : CFG._dlKo;
  fillProfs(); refresh();
}
function fillProfs(){
  const sel = $('#지도교수Sel'); if (!CFG.profs.length){ sel.style.display='none'; $('#지도교수').style.display='block'; return; }
  const cur = sel.value;
  sel.innerHTML = \`<option value="">\${T('selProf')}</option>\` + CFG.profs.map(p=>{ const [ko,en] = String(p).split('/').map(s=>s.trim()); const lab = LANG==='en' && en ? \`\${en} (\${ko})\` : (en ? \`\${ko} · \${en}\` : ko); return \`<option value="\${esc(ko)}">\${esc(lab)}</option>\`; }).join('') + \`<option value="__etc">\${T('etcProf')}</option>\`;
  sel.value = cur;
}

function init(){
  let dl = CFG.deadline || '-';
  if (/^\\d{4}-\\d{2}-\\d{2}T/.test(dl)) { const d = new Date(new Date(dl).getTime() + 9*3600e3); dl = d.toISOString().slice(0,16).replace('T',' '); }
  const mm = String(dl).match(/^(\\d{4})-(\\d{2})-(\\d{2})(?:[ T](\\d{2}:\\d{2}))?/);
  if (mm){ const dd = new Date(+mm[1], +mm[2]-1, +mm[3]); const wk = ['일','월','화','수','목','금','토'][dd.getDay()], wkE = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'][dd.getDay()];
    CFG._dlKo = \`\${mm[1]}.\${+mm[2]}.\${+mm[3]}.(\${wk})\${mm[4] ? ' '+mm[4] : ''}\`;
    CFG._dlEn = \`\${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][+mm[2]-1]} \${+mm[3]}, \${mm[1]} (\${wkE})\${mm[4] ? ' '+mm[4]+' KST' : ''}\`; }
  $('#dl').textContent = dl;
  $('#range').textContent = (CFG.from && CFG.to) ? CFG.from + ' ~ ' + CFG.to : '-';
  const sel = $('#지도교수Sel');
  sel.onchange = () => { const etc = sel.value==='__etc'; $('#지도교수').style.display = etc?'block':'none'; $('#지도교수').value = etc ? '' : sel.value; };
  if (CFG.closed){
    $('#app').innerHTML = '<section class="card closed"><h2 style="justify-content:center" data-t="closed"></h2><p class="desc" id="contact"></p></section>';
    document.querySelector('.bar').style.display='none';
  } else addPaper();
  setLang(LANG);
  authInit();
}

function mount(tplId, boxId){
  const node = document.getElementById(tplId).content.firstElementChild.cloneNode(true);
  node.querySelectorAll('.seg').forEach(seg => {
    const nm = 'r' + (++uid);
    seg.querySelectorAll('input').forEach((inp,i) => { const id = nm+'_'+i; inp.name = nm; inp.id = id; inp.nextElementSibling.htmlFor = id; });
  });
  node.querySelectorAll('.file input').forEach(inp => inp.onchange = () => {
    const f = inp.files[0], box = inp.parentElement;
    if (f && f.size > CFG.maxMb*1024*1024){ alert(T('tooBig',{mb:CFG.maxMb})); inp.value=''; }
    else if (f && !/\\.(pdf|jpe?g|png)$/i.test(f.name)){ alert(T('badType')); inp.value=''; }
    const ok = !!inp.files[0];
    if (ok) { delete box.dataset.keep; box.classList.remove('keep'); }
    box.classList.toggle('has', ok); box.querySelector('.t').textContent = ok ? inp.files[0].name : T('choose'); box.querySelector('.ico').textContent = ok ? '✓' : 'PDF';
  });
  node.querySelector('.del').onclick = () => { node.remove(); refresh(); };
  node.addEventListener('input', () => node.querySelectorAll('.bad').forEach(b=>b.classList.remove('bad')));
  translate(node);
  document.getElementById(boxId).appendChild(node); refresh();
}
const addPaper = () => mount('tPaper','papers');
const addPatent = () => mount('tPatent','patents');
const addAward = () => mount('tAward','awards');

function refresh(){
  [['papers','sP','emptyP'],['patents','sT','emptyT'],['awards','sA','emptyA']].forEach(([b,s,msg])=>{
    const box = document.getElementById(b); if (!box) return;
    const items = box.querySelectorAll('.item');
    items.forEach((it,i)=>it.querySelector('.n').textContent = i+1);
    document.getElementById(s).textContent = items.length;
    let e = box.querySelector('.empty');
    if (!items.length){ if(!e){ e = document.createElement('div'); e.className='empty'; box.appendChild(e); } e.textContent = T(msg); }
    else if (e) e.remove();
  });
}

function val(item, name){ const el = item.querySelector(\`[name="\${name}"]\`); return el ? el.value.trim() : ''; }
function radio(item){ const c = item.querySelector('.seg input:checked'); return c ? c.value : ''; }
function fileOf(item, name){ const el = item.querySelector(\`input[type=file][name="\${name}"]\`); return el && el.files[0]; }
function keepOf(item, name){ const el = item.querySelector(\`input[type=file][name="\${name}"]\`); return el ? (el.closest('.file').dataset.keep || '') : ''; }
function bad(scope, key, msg){ const f = scope.querySelector(\`.f[data-k="\${key}"]\`); if (!f) return f; f.classList.add('bad');
  const e = f.querySelector('.err'); if (e){ if (!e.dataset.t0) e.dataset.t0 = e.dataset.t; e.dataset.t = msg || e.dataset.t0; e.textContent = T(e.dataset.t); } return f; }
const normDoi = d => String(d||'').trim().replace(/^https?:\\/\\/(dx\\.)?doi\\.org\\//i,'').replace(/^doi:\\s*/i,'');
const isAbbr = j => /\\b[A-Z][a-z]{0,5}\\.(\\s|$)/.test(j) && (j.match(/\\./g)||[]).length >= 2;
const parseK = t => { const m = String(t||'').match(/(\\d{4})\\D+(\\d{1,2})\\D+(\\d{1,2})/); return m ? \`\${m[1]}-\${m[2].padStart(2,'0')}-\${m[3].padStart(2,'0')}\` : ''; };

function collect(){
  document.querySelectorAll('.bad').forEach(b=>b.classList.remove('bad'));
  const errs = []; const P = {lang: LANG};
  P.학번 = $('#학번').value.trim(); if(!/^\\d{8}$/.test(P.학번)) errs.push(bad(document,'학번'));
  P.성명 = $('#성명').value.trim(); if(!P.성명) errs.push(bad(document,'성명'));
  P.영문명 = $('#영문명').value.trim(); if(!/[A-Za-z]/.test(P.영문명)) errs.push(bad(document,'영문명'));
  const c = document.querySelector('input[name="과정"]:checked'); P.과정 = c ? c.value : ''; if(!P.과정) errs.push(bad(document,'과정'));
  P.재학학기 = $('#재학학기').value.trim();
  P.지도교수 = $('#지도교수').value.trim(); if(!P.지도교수) errs.push(bad(document,'지도교수'));
  P.메모 = $('#메모').value.trim();
  const files = [], seenP = {};
  P.papers = [...document.querySelectorAll('#papers .item')].map(it => {
    const x = {journal:val(it,'journal'), title:val(it,'title'), doi:normDoi(val(it,'doi')), accept:val(it,'accept'), role:radio(it), authors:val(it,'authors')};
    if(!x.journal) errs.push(bad(it,'journal')); else if(isAbbr(x.journal)) errs.push(bad(it,'journal','journalAbbr'));
    if(!x.accept) errs.push(bad(it,'accept')); else if(x.accept > TODAY) errs.push(bad(it,'accept','acceptFuture'));
    if(!x.role) errs.push(bad(it,'role'));
    if(x.doi && !/^10\\.\\d{4,9}\\/\\S+$/.test(x.doi)) errs.push(bad(it,'doi'));
    if(!(+x.authors>=1 && Number.isInteger(+x.authors))) errs.push(bad(it,'authors'));
    else if((x.role==='co-FA'||x.role==='ETC') && +x.authors<2) errs.push(bad(it,'authors','authorsMin2'));
    const key = (x.doi || x.title).toLowerCase().replace(/[^a-z0-9가-힣]/g,'');
    if(key && seenP[key]) errs.push(bad(it,'journal','dupPaper')); seenP[key] = 1;
    const f1 = fileOf(it,'file1'), f2 = fileOf(it,'file2'), k1 = keepOf(it,'file1'), k2 = keepOf(it,'file2');
    if(!f1 && !k1) errs.push(bad(it,'file1'));
    if(f1) files.push([x,'file1',f1]); else if(k1) x.file1 = k1;
    if(f2) files.push([x,'file2',f2]); else if(k2) x.file2 = k2; return x;
  });
  P.patents = [...document.querySelectorAll('#patents .item')].map(it => {
    const x = {type:radio(it), name:val(it,'name'), no:val(it,'no'), date:val(it,'date'), share:val(it,'share')};
    if(!x.type) errs.push(bad(it,'type')); if(!(+x.share>0 && +x.share<=100)) errs.push(bad(it,'share'));
    const f = fileOf(it,'file'), k = keepOf(it,'file'); if(f) files.push([x,'file',f]); else if(k) x.file = k; else errs.push(bad(it,'file')); return x;
  });
  P.awards = [...document.querySelectorAll('#awards .item')].map(it => {
    const x = {type:radio(it), event:val(it,'event'), prize:val(it,'prize'), date:val(it,'date')};
    if(!x.type) errs.push(bad(it,'type'));
    const f = fileOf(it,'file'), k = keepOf(it,'file'); if(f) files.push([x,'file',f]); else if(k) x.file = k; else errs.push(bad(it,'file')); return x;
  });
  return {P, files, errs: errs.filter(Boolean)};
}

const TODAY = new Date(Date.now() + 9*3600e3).toISOString().slice(0,10);
/* 휴대폰 사진(JPG·PNG)이 크면 글자가 충분히 보이는 크기(긴 변 2200px)로 줄여 올림 — PDF는 그대로 */
async function shrink(f){
  if (!/^image\\/(jpeg|png)$/.test(f.type) || f.size < 1.5*1024*1024) return f;
  try {
    const bmp = await createImageBitmap(f); const k = Math.min(1, 2200 / Math.max(bmp.width, bmp.height));
    const c = document.createElement('canvas'); c.width = Math.round(bmp.width*k); c.height = Math.round(bmp.height*k);
    const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0,0,c.width,c.height); g.drawImage(bmp, 0, 0, c.width, c.height);
    const b = await new Promise(ok => c.toBlob(ok, 'image/jpeg', 0.88));
    return b && b.size < f.size ? new File([b], f.name.replace(/\\.\\w+$/,'') + '.jpg', {type:'image/jpeg'}) : f;
  } catch(e){ return f; }
}
/* 실적인정 기간 밖 Accept 날짜 등 — 막지는 않고 한 번 더 확인 */
function softChecks(P){
  const a = parseK(CFG.from), b = parseK(CFG.to), w = [];
  if (a && b) P.papers.forEach((x,i) => { if (x.accept && (x.accept < a || x.accept > b)) w.push(T('chkOut',{i:i+1,d:x.accept,r:CFG.from+' ~ '+CFG.to})); });
  return w;
}
function askWarn(w){
  return new Promise(ok => {
    const box = $('#box'); $('#modal').classList.add('on');
    box.innerHTML = \`<h3>\${T('chkTitle')}</h3>\${w.map(x=>\`<p style="text-align:left">• \${esc(x)}</p>\`).join('')}<p style="margin-top:16px;display:flex;gap:8px;justify-content:center"><button class="btn" id="wFix">\${T('chkFix')}</button><button class="go" id="wGo">\${T('chkGo')}</button></p>\`;
    $('#wFix').onclick = () => { $('#modal').classList.remove('on'); ok(false); };
    $('#wGo').onclick = () => ok(true);
  });
}
function readUrl(f){ return new Promise((ok,no)=>{ const r=new FileReader(); r.onload=()=>ok(r.result); r.onerror=no; r.readAsDataURL(f); }); }
function gs(fn, ...a){
  if (CFG_TPL && window.google && google.script) return new Promise((ok,no)=>{ const run = google.script.run.withSuccessHandler(ok).withFailureHandler(no); run[fn](...a); });
  if (!API) return Promise.reject(new Error(T('previewErr')));
  const once = () => fetch(API, {method:'POST', headers:{'Content-Type':'text/plain;charset=utf-8'}, body: JSON.stringify({action:fn, args:a})}).then(r => r.json());
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  // 일시적인 서버 오류(로그인 확인 실패·네트워크)는 한 번 더 시도 — 학생이 다시 인증하거나 처음부터 다시 내는 일을 줄임
  return once().catch(() => sleep(1500).then(once)).then(j => {
    if (!j.ok && String(j.error||'').startsWith('LOGIN:') && AUTH.exp > Date.now()) return sleep(1500).then(once);
    return j;
  }).then(j => { if (j.ok) return j.data; const m = String(j.error||''); if (m.startsWith('LOGIN:')) { onExpired(); throw new Error(T('sessionExp')); } throw new Error(m); });
}
function esc(s){ return String(s).replace(/[&<>"]/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

async function send(){
  const {P, files, errs} = collect();
  if (errs.length){ errs[0].scrollIntoView({behavior:'smooth', block:'center'}); return; }
  if (!P.papers.length && !P.patents.length && !P.awards.length){ alert(T('needOne')); return; }
  if (!$('#agree').checked){ $('#agreeBox').scrollIntoView({behavior:'smooth',block:'center'}); $('#agreeBox').style.color='var(--red)'; return; }
  const warn = softChecks(P);
  if (warn.length && !(await askWarn(warn))) return;
  $('#go').disabled = true;
  const box = $('#box'); $('#modal').classList.add('on');
  box.innerHTML = \`<h3>\${T('sending')}</h3><p id="pm">\${T('uploading')}</p><div class="prog"><i id="pb"></i></div><p style="font-size:12px">\${T('noClose')}</p>\`;
  try{
    // 증빙을 3개씩 동시에 올림 (한 개씩보다 약 2배 빠름)
    let done = 0, next = 0;
    const show = () => { $('#pm').textContent = T('upPar',{i:done,n:files.length}); $('#pb').style.width = Math.round(done/(files.length+1)*100)+'%'; };
    show();
    const worker = async () => { while (next < files.length){ const [obj,key,f0] = files[next++];
      const f = await shrink(f0);
      obj[key] = await gs('uploadFile', {name:f.name}, await readUrl(f), AUTH.token); done++; show(); } };
    await Promise.all([worker(), worker(), worker()]);
    $('#pm').textContent = T('saving');
    const r = await gs('submit', P, AUTH.token);
    MINE = null; $('#go').disabled = false; $('#agree').checked = false;
    $('#pb').style.width = '100%';
    box.innerHTML = \`<div class="big">✓</div><h3>\${T('done')}</h3><p>\${T('rcpt')} <b>\${esc(r.id)}</b></p><p>\${T('doneCnt',{p:r.papers,t:r.patents,a:r.awards})}</p><p style="margin-top:12px;font-size:13px;color:var(--mute)">\${T('doneNote',{time:esc(r.time)})}</p><p style="margin-top:16px"><button class="go" onclick="document.getElementById('modal').classList.remove('on');showTab('mine')">\${T('tabMine')}</button></p>\`;
  }catch(e){
    box.innerHTML = \`<h3>\${T('fail')}</h3><p>\${esc((e && e.message) || e)}</p><p style="margin-top:14px"><button class="go" onclick="document.getElementById('modal').classList.remove('on')">\${T('back')}</button></p>\`;
    $('#go').disabled = false;
  }
}

/* ================= 학생 본인 확인 · 내 제출 내역 ================= */
let AUTH = { token:'', email:'', exp:0 };
let MINE = null, TAB = 'form';
const ls_get = k => { try { return localStorage.getItem(k) || ''; } catch(e) { return ''; } };
const ls_set = (k,v) => { try { v ? localStorage.setItem(k,v) : localStorage.removeItem(k); } catch(e) {} };
function authInit(){
  try { const a = JSON.parse(ls_get('mreAuth') || '{}'); if (a.token && a.exp > Date.now()) AUTH = a; } catch(e) {}
  renderAuth();
}
function renderAuth(){
  const ok = !!AUTH.token && !PREVIEW && !CFG.closed;
  $('#gate').hidden = ok || PREVIEW || CFG.closed;
  $('#sess').hidden = !ok; $('#tabs').hidden = !ok;
  if (ok){ $('#sEmail').textContent = AUTH.email; $('#vEmail').value = AUTH.email; showTab(TAB); }
  else { $('#formArea').hidden = !(PREVIEW && !CFG.closed); $('#mineArea').hidden = true; }
  updateBar();
}
function updateBar(){
  const show = location.hash !== '#admin' && !CFG.closed && !$('#formArea').hidden;
  document.querySelector('.bar').style.display = show ? '' : 'none';
}
function showTab(t){
  TAB = t;
  document.querySelectorAll('#tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === t));
  $('#formArea').hidden = t !== 'form'; $('#mineArea').hidden = t !== 'mine';
  updateBar();
  if (t === 'mine') loadMine();
  else if (!MINE) loadMine(true);
}
let resendT = 0;
async function sendCode(){
  const email = $('#gEmail').value.trim().toLowerCase();
  $('#gErr1').textContent=''; $('#gEmailBox').classList.remove('bad');
  if (!/^[^@\\s]+@(gm\\.)?gist\\.ac\\.kr$/.test(email)){ $('#gErr1').textContent = T('gBadEmail'); $('#gEmailBox').classList.add('bad'); return; }
  const b = $('#gSend'); b.disabled = true;
  try{
    await gs('requestCode', email, LANG);
    $('#gCodeBox').hidden = false; $('#gHint').textContent = T('gSent',{email}); $('#gCode').focus();
    let s = 60; clearInterval(resendT);
    resendT = setInterval(()=>{ s--; b.textContent = s>0 ? \`\${T('gResend')} (\${s})\` : T('gResend'); if (s<=0){ clearInterval(resendT); b.disabled=false; } }, 1000);
  }catch(e){ $('#gErr1').textContent = (e && e.message) || e; $('#gEmailBox').classList.add('bad'); b.disabled = false; }
}
async function checkCode(){
  const email = $('#gEmail').value.trim().toLowerCase(), code = $('#gCode').value.replace(/\\D/g,'');
  $('#gErr2').textContent=''; $('#gCodeBox').classList.remove('bad');
  if (code.length !== 6){ $('#gCodeBox').classList.add('bad'); return; }
  const b = $('#gVerify'); b.disabled = true;
  try{
    const r = await gs('verifyCode', email, code);
    AUTH = { token:r.token, email:r.email, exp: Date.now() + (r.hours||6)*3600e3 - 60e3 };
    ls_set('mreAuth', JSON.stringify(AUTH)); MINE = null; TAB = 'form';
    renderAuth();
    window.scrollTo({top:0, behavior:'smooth'});
  }catch(e){ $('#gErr2').textContent = (e && e.message) || e; $('#gCodeBox').classList.add('bad'); }
  b.disabled = false;
}
function logout(){ AUTH = {token:'',email:'',exp:0}; MINE = null; ls_set('mreAuth',''); $('#gCode').value=''; $('#gCodeBox').hidden = true; renderAuth(); }
function onExpired(){ if (AUTH.token){ logout(); alert(T('sessionExp')); } }
document.addEventListener('keydown', e => { if (e.key==='Enter' && e.target.id==='gEmail') sendCode(); if (e.key==='Enter' && e.target.id==='gCode') checkCode(); });

async function loadMine(silent){
  if (!AUTH.token) return;
  const box = $('#mineArea');
  if (!silent) box.innerHTML = \`<section class="card empty-a">\${T('mineLoad')}</section>\`;
  try{ MINE = await gs('myData', AUTH.token); }catch(e){ if (!silent) box.innerHTML = \`<section class="card empty-a">\${esc((e&&e.message)||e)}</section>\`; return; }
  // 이미 낸 학번은 고정
  const cur = MINE.submissions.find(s => s['상태']==='유효') || MINE.submissions[0];
  if (cur){ const sid = $('#학번'); if (!sid.value) sid.value = cur['학번']; sid.readOnly = true; }
  if (!silent || TAB==='mine') renderMine();
}
function renderMine(){
  const box = $('#mineArea'); const subs = (MINE && MINE.submissions) || [];
  if (!subs.length){ box.innerHTML = \`<section class="card empty-a">\${T('mineEmpty')}</section>\`; return; }
  const ev = (u,label,s) => u ? \`<button class="btn sm evbtn" type="button" data-s="mine" data-u="\${esc(u)}" data-l="\${esc(s['학번']+' '+label)}">📄 \${T('view')}</button>\` : \`<span class="sub2">\${T('none')}</span>\`;
  const card = (s, latest) => \`<section class="card">
    <div class="mine-head"><div><div style="font-weight:800;font-size:17px">\${esc(s['성명'])} <span class="sub2">\${esc(s['학번'])}\${s['논문영문명'] ? ' · '+esc(s['논문영문명']) : ''}</span></div>
      <div class="sub2">\${T('rcpt')} \${esc(s['접수번호'])} · \${esc(s['접수시각'])}</div></div>
      <div>\${latest ? '<span class="pill ok">'+T('mineLatest')+'</span>' : '<span class="pill old">'+T('mineOld')+'</span>'}</div></div>
    <div class="sub2" style="margin-bottom:10px">\${esc(s['과정'])} · \${esc(s['지도교수'])}</div>
    <h4 style="margin:12px 0 6px">\${T('s2')} \${s.papers.length}</h4>
    \${s.papers.length ? \`<div class="a-wrap"><table class="a-tbl"><tr><th>#</th><th>\${T('journal')}</th><th>\${T('accept')}</th><th>\${T('role')}</th><th>\${T('authors')}</th><th></th></tr>\${s.papers.map(p=>\`<tr><td>\${esc(p['순번'])}</td><td><b>\${esc(p['저널명(학생)'])}</b><div class="sub2">\${esc(p['논문제목'])}</div></td><td>\${esc(p['Accept일'])}</td><td>\${esc(p['역할(학생)'])}</td><td>\${esc(p['저자수(학생)'])}</td><td>\${ev(p['증빙1'],'논문'+p['순번'],s)} \${p['증빙2']?ev(p['증빙2'],'논문'+p['순번']+'_추가',s):''}</td></tr>\`).join('')}</table></div>\` : \`<div class="sub2">\${T('none')}</div>\`}
    <h4 style="margin:12px 0 6px">\${T('s3')} \${s.patents.length}</h4>
    \${s.patents.length ? \`<div class="a-wrap"><table class="a-tbl"><tr><th>#</th><th>\${T('ptype')}</th><th>\${T('pname')}</th><th>\${T('share')}</th><th></th></tr>\${s.patents.map(p=>\`<tr><td>\${esc(p['순번'])}</td><td>\${esc(p['구분'])}</td><td>\${esc(p['명칭'])}<div class="sub2">\${esc(p['출원·등록번호'])} \${esc(p['일자'])}</div></td><td>\${esc(p['기여도(%)'])}%</td><td>\${ev(p['증빙'],'특허'+p['순번'],s)}</td></tr>\`).join('')}</table></div>\` : \`<div class="sub2">\${T('none')}</div>\`}
    <h4 style="margin:12px 0 6px">\${T('s4')} \${s.awards.length}</h4>
    \${s.awards.length ? \`<div class="a-wrap"><table class="a-tbl"><tr><th>#</th><th>\${T('ptype')}</th><th>\${T('aevent')}</th><th></th></tr>\${s.awards.map(p=>\`<tr><td>\${esc(p['순번'])}</td><td>\${esc(p['구분'])}</td><td>\${esc(p['학회·대회명'])}<div class="sub2">\${esc(p['수상명'])} \${esc(p['일자'])}</div></td><td>\${ev(p['증빙'],'수상'+p['순번'],s)}</td></tr>\`).join('')}</table></div>\` : \`<div class="sub2">\${T('none')}</div>\`}
    \${latest ? \`<div style="margin-top:16px;display:flex;gap:10px;align-items:center;flex-wrap:wrap"><button class="btn red" type="button" onclick="editMine()">\${T('mineEdit')}</button><span class="sub2">\${T('mineHint')}</span></div>\` : ''}
  </section>\`;
  const latest = subs.find(s => s['상태']==='유효'), olds = subs.filter(s => s !== latest);
  box.innerHTML = (latest ? card(latest, true) : '') + (olds.length ? \`<details class="card" style="padding:16px 20px"><summary style="cursor:pointer;font-weight:700">\${T('mineOld')} \${olds.length}</summary>\${olds.map(s=>card(s,false)).join('')}</details>\` : '');
}
const fid = u => (String(u||'').match(/[-\\w]{25,}/) || [''])[0];
function setKeep(item, name, url){
  const el = item.querySelector(\`input[type=file][name="\${name}"]\`); if (!el || !url) return;
  const box = el.closest('.file'); box.dataset.keep = fid(url); box.classList.add('keep');
  box.querySelector('.t').textContent = T('keepFile'); box.querySelector('.ico').textContent = '↺';
}
function setRadio(item, v){ const r = [...item.querySelectorAll('.seg input')].find(x => x.value === v); if (r) r.checked = true; }
function editMine(){
  const s = MINE && MINE.submissions.find(x => x['상태']==='유효'); if (!s) return;
  $('#학번').value = s['학번']; $('#성명').value = s['성명']; $('#영문명').value = s['논문영문명'] || ''; $('#재학학기').value = s['재학학기'] || '';
  const c = [...document.querySelectorAll('input[name="과정"]')].find(x => x.value === s['과정']); if (c) c.checked = true;
  const sel = $('#지도교수Sel'); const has = [...sel.options].some(o => o.value === s['지도교수']);
  if (has){ sel.value = s['지도교수']; $('#지도교수').style.display='none'; } else { sel.value = '__etc'; $('#지도교수').style.display='block'; }
  $('#지도교수').value = s['지도교수']; $('#메모').value = s['학생메모'] || '';
  ['papers','patents','awards'].forEach(b => document.getElementById(b).innerHTML = '');
  s.papers.forEach(p => { addPaper(); const it = [...document.querySelectorAll('#papers .item')].pop();
    it.querySelector('[name=journal]').value = p['저널명(학생)']; it.querySelector('[name=title]').value = p['논문제목'];
    it.querySelector('[name=doi]').value = p['DOI']; it.querySelector('[name=accept]').value = p['Accept일'];
    it.querySelector('[name=authors]').value = p['저자수(학생)']; setRadio(it, p['역할(학생)']); setKeep(it,'file1',p['증빙1']); setKeep(it,'file2',p['증빙2']); });
  s.patents.forEach(p => { addPatent(); const it = [...document.querySelectorAll('#patents .item')].pop();
    setRadio(it, p['구분']); it.querySelector('[name=name]').value = p['명칭']; it.querySelector('[name=no]').value = p['출원·등록번호'];
    it.querySelector('[name=date]').value = p['일자']; it.querySelector('[name=share]').value = p['기여도(%)']; setKeep(it,'file',p['증빙']); });
  s.awards.forEach(p => { addAward(); const it = [...document.querySelectorAll('#awards .item')].pop();
    setRadio(it, p['구분']); it.querySelector('[name=event]').value = p['학회·대회명']; it.querySelector('[name=prize]').value = p['수상명'];
    it.querySelector('[name=date]').value = p['일자']; setKeep(it,'file',p['증빙']); });
  refresh(); showTab('form');
  const n = $('#notice'); n.style.display=''; n.textContent = T('loadedEdit');
  window.scrollTo({top:0, behavior:'smooth'});
}

/* ================= 담당자(관리자) 화면 ================= */
let ADM = { key:'', data:null };
const ss_get = k => { try { return sessionStorage.getItem(k) || ''; } catch(e) { return ''; } };
const ss_set = (k,v) => { try { v ? sessionStorage.setItem(k,v) : sessionStorage.removeItem(k); } catch(e) {} };
function route(){
  const isA = location.hash === '#admin';
  $('#app').hidden = isA; updateBar();
  $('#admin').hidden = !isA; document.querySelector('.chips').style.display = isA ? 'none' : '';
  $('#title').textContent = isA ? '장학금 연구실적 — 담당자 화면' : ((LANG==='en' && CFG.titleEn) || CFG.title);
  document.querySelector('.lang').style.display = isA ? 'none' : '';
  document.querySelector('[data-t=lead]').textContent = isA ? '학생 제출 목록을 보고 증빙 PDF를 바로 열어 봅니다.' : T('lead');
  if (isA){ ADM.key = ADM.key || ss_get('admKey'); if (ADM.key && !ADM.data) loadAdmin(); else showAdmin(!!ADM.data); }
}
function showAdmin(inside){ $('#aLogin').hidden = inside; $('#aMain').hidden = !inside; }
async function adminLogin(){ ADM.key = $('#aKey').value.trim(); if (!ADM.key) return; await loadAdmin(); }
function adminLogout(){ ADM = {key:'', data:null}; ss_set('admKey',''); $('#aKey').value=''; showAdmin(false); }
async function loadAdmin(){
  $('#aErr').textContent=''; $('#aKeyBox').classList.remove('bad');
  $('#aNow').textContent = '불러오는 중…';
  try{
    ADM.data = await gs('adminList', ADM.key); ss_set('admKey', ADM.key);
    $('#aSheet').href = ADM.data.sheetUrl; showAdmin(true); renderAdmin();
  }catch(e){
    ADM.data = null; ss_set('admKey',''); showAdmin(false);
    $('#aErr').textContent = (e && e.message) || String(e); $('#aKeyBox').classList.add('bad');
  }
}
const nItems = s => s.papers.length + s.patents.length + s.awards.length;
function renderAdmin(){
  const D = ADM.data; if (!D) return;
  const all = D.students, valid = all.filter(s => s['상태']==='유효');
  const sum = k => valid.reduce((a,s)=>a+s[k].length,0);
  const last = valid[0] ? valid[0]['접수시각'] : '-';
  $('#aStats').innerHTML = [['제출 학생',valid.length+'명'],['논문',sum('papers')+'편'],['특허',sum('patents')+'건'],['수상',sum('awards')+'건'],['최근 접수',last.slice(5,16)||'-']]
    .map(([a,b])=>\`<div class="a-stat"><span>\${a}</span><b>\${esc(b)}</b></div>\`).join('');
  $('#aNow').textContent = \`\${D.now} 기준 · 같은 학번은 마지막 제출만 '유효' · 행을 누르면 실적과 증빙이 열립니다\`;
  const q = $('#aQ').value.trim().toLowerCase(), old = $('#aOld').checked;
  const list = all.filter(s => (old || s['상태']==='유효') && (!q || JSON.stringify(s).toLowerCase().includes(q)));
  if (!list.length){ $('#aList').innerHTML = '<div class="card empty-a">아직 제출이 없습니다.</div>'; return; }
  $('#aList').innerHTML = list.map((s,i) => {
    const pil = s['상태']==='유효' ? '<span class="pill ok">유효</span>' : '<span class="pill old">대체됨</span>';
    const nSub = all.filter(x => x['학번']===s['학번']).length;
    const rep = nSub > 1 ? \`<span class="pill old" title="같은 학번으로 \${nSub}번 제출됨 — '이전 제출도 보기'로 비교">제출 \${nSub}회</span>\` : '';
    const ev = (u,label) => u ? \`<button class="btn sm evbtn" type="button" data-u="\${esc(u)}" data-l="\${esc(s['학번']+' '+s['성명']+' '+label)}">📄 \${label}</button>\` : '<span class="sub">없음</span>';
    const papers = s.papers.map(p => \`<tr><td>\${esc(p['순번'])}</td><td><b>\${esc(p['저널명(학생)'])}</b><div class="sub">\${esc(p['논문제목'])}</div><div class="sub">\${esc(p['DOI'])}</div></td><td>\${esc(p['Accept일'])}</td><td>\${esc(p['역할(학생)'])}</td><td>\${esc(p['저자수(학생)'])}</td><td>\${ev(p['증빙1'],'증빙')} \${p['증빙2']?ev(p['증빙2'],'추가'):''}</td><td>\${p['IF(JCR)']?esc(p['순위']+'/'+p['전체']+' · IF '+p['IF(JCR)']):'<span class="sub">검증 전</span>'}</td></tr>\`).join('');
    const pats = s.patents.map(p => \`<tr><td>\${esc(p['순번'])}</td><td>\${esc(p['구분'])}</td><td>\${esc(p['명칭'])}<div class="sub">\${esc(p['출원·등록번호'])} \${esc(p['일자'])}</div></td><td>\${esc(p['기여도(%)'])}%</td><td>\${ev(p['증빙'],'증빙')}</td></tr>\`).join('');
    const aws = s.awards.map(p => \`<tr><td>\${esc(p['순번'])}</td><td>\${esc(p['구분'])}</td><td>\${esc(p['학회·대회명'])}<div class="sub">\${esc(p['수상명'])} \${esc(p['일자'])}</div></td><td>\${ev(p['증빙'],'증빙')}</td></tr>\`).join('');
    return \`<div class="a-row" id="ar\${i}">
      <div class="a-head" onclick="this.parentElement.classList.toggle('open')">
        <div><div class="nm">\${esc(s['성명'])}</div><div class="sub">\${esc(s['학번'])}\${s['논문영문명'] ? ' · '+esc(s['논문영문명']) : ''}</div><div class="sub" style="font-size:11.5px">\${esc(s['이메일']||'')}</div></div>
        <div>\${pil}\${rep}<span class="pill">\${esc(s['과정'])}</span><span class="sub">\${esc(s['재학학기'] ? s['재학학기']+'학기' : '')}</span></div>
        <div><span class="sub">지도교수</span><div>\${esc(s['지도교수'])}</div></div>
        <div class="hide-m"><span class="pill">논문 \${s.papers.length}</span><span class="pill">특허 \${s.patents.length}</span><span class="pill">수상 \${s.awards.length}</span></div>
        <div class="hide-m sub">\${esc(s['접수시각'])}</div>
        <div class="hide-m sub">\${esc(s['접수번호'])}</div>
        <div class="chev">›</div>
      </div>
      <div class="a-body">
        \${s['학생메모'] ? \`<div class="notice" style="margin-bottom:12px">학생 메모: \${esc(s['학생메모'])}</div>\` : ''}
        <h4>SCI 논문 \${s.papers.length}편</h4>
        \${s.papers.length ? \`<div class="a-wrap"><table class="a-tbl"><tr><th>#</th><th>저널 / 제목</th><th>Accept</th><th>역할</th><th>저자수</th><th>증빙</th><th>JCR</th></tr>\${papers}</table></div>\` : '<div class="sub">없음</div>'}
        <h4>특허 \${s.patents.length}건</h4>
        \${s.patents.length ? \`<div class="a-wrap"><table class="a-tbl"><tr><th>#</th><th>구분</th><th>명칭</th><th>기여도</th><th>증빙</th></tr>\${pats}</table></div>\` : '<div class="sub">없음</div>'}
        <h4>수상 \${s.awards.length}건</h4>
        \${s.awards.length ? \`<div class="a-wrap"><table class="a-tbl"><tr><th>#</th><th>구분</th><th>학회·대회 / 수상</th><th>증빙</th></tr>\${aws}</table></div>\` : '<div class="sub">없음</div>'}
      </div></div>\`;
  }).join('');
}
let VURL = '';
async function viewFile(url, label, src){
  const v = $('#viewer'); v.classList.add('on'); $('#vName').textContent = label + ' — 불러오는 중…'; $('#vBody').innerHTML = '';
  try{
    const f = src === 'mine' ? await gs('myFile', AUTH.token, url) : await gs('adminFile', ADM.key, url);
    const bin = atob(f.data), arr = new Uint8Array(bin.length); for (let i=0;i<bin.length;i++) arr[i]=bin.charCodeAt(i);
    // 파일 내용으로 종류를 다시 확인 — PDF·PNG·JPG만 화면에 띄움
    const kind = (arr[0]===0x25&&arr[1]===0x50&&arr[2]===0x44&&arr[3]===0x46) ? 'application/pdf'
               : (arr[0]===0x89&&arr[1]===0x50&&arr[2]===0x4E&&arr[3]===0x47) ? 'image/png'
               : (arr[0]===0xFF&&arr[1]===0xD8&&arr[2]===0xFF) ? 'image/jpeg' : '';
    if (VURL) URL.revokeObjectURL(VURL);
    VURL = URL.createObjectURL(new Blob([arr], {type: kind || 'application/octet-stream'}));
    $('#vName').textContent = f.name; $('#vDown').href = VURL; $('#vDown').download = f.name;
    $('#vOpen').style.display = kind ? '' : 'none'; $('#vOpen').href = kind ? VURL : '#';
    if (!kind) { $('#vBody').innerHTML = '<p style="padding:30px">PDF·JPG·PNG가 아닌 파일이라 화면에 띄우지 않습니다. 필요하면 내려받아 확인하세요.</p>'; return; }
    if (kind.startsWith('image/')) { $('#vBody').innerHTML = \`<img src="\${VURL}" alt="">\`; return; }
    try { await renderPdf(arr); } catch(err) { $('#vBody').innerHTML = \`<iframe src="\${VURL}" title="증빙" sandbox></iframe>\`; }
  }catch(e){ $('#vName').textContent = label; $('#vBody').innerHTML = \`<p style="padding:30px">\${esc((e && e.message) || e)}</p>\`; }
}
function loadPdfJs(){
  if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
  return new Promise((ok,no)=>{ const s=document.createElement('script'); s.src='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    s.onload=()=>{ pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js'; ok(pdfjsLib); }; s.onerror=no; document.head.appendChild(s); });
}
async function renderPdf(bytes){
  const lib = await loadPdfJs();
  const doc = await lib.getDocument({data: bytes}).promise;
  const body = $('#vBody'); body.innerHTML = ''; body.style.display = 'block'; body.style.background = '#E9EBEA'; body.style.padding = '14px';
  const w = Math.min(body.clientWidth - 28, 1000);
  for (let i = 1; i <= doc.numPages; i++){
    const page = await doc.getPage(i), vp0 = page.getViewport({scale:1}), sc = w / vp0.width, dpr = window.devicePixelRatio || 1;
    const vp = page.getViewport({scale: sc * dpr}), c = document.createElement('canvas');
    c.width = vp.width; c.height = vp.height; c.style.width = (vp.width/dpr) + 'px'; c.style.display = 'block'; c.style.margin = '0 auto 12px'; c.style.background = '#fff'; c.style.boxShadow = '0 2px 10px rgba(0,0,0,.12)';
    body.appendChild(c); await page.render({canvasContext: c.getContext('2d'), viewport: vp}).promise;
  }
}
function closeViewer(){ const b=$('#vBody'); b.style.display=''; b.style.background=''; b.style.padding=''; $('#viewer').classList.remove('on'); $('#vBody').innerHTML=''; }
document.addEventListener('keydown', e => { if (e.key==='Escape') closeViewer(); });
function adminCsv(){
  const D = ADM.data; if (!D) return;
  const rows = [['상태','접수번호','접수시각','학번','성명','과정','재학학기','지도교수','구분','순번','내용','세부','날짜','역할/구분','저자수/기여도','증빙']];
  D.students.forEach(s => {
    const base = [s['상태'],s['접수번호'],s['접수시각'],s['학번'],s['성명'],s['과정'],s['재학학기'],s['지도교수']];
    s.papers.forEach(p => rows.push(base.concat(['논문',p['순번'],p['저널명(학생)'],p['논문제목'],p['Accept일'],p['역할(학생)'],p['저자수(학생)'],p['증빙1']])));
    s.patents.forEach(p => rows.push(base.concat(['특허',p['순번'],p['명칭'],p['출원·등록번호'],p['일자'],p['구분'],p['기여도(%)'],p['증빙']])));
    s.awards.forEach(p => rows.push(base.concat(['수상',p['순번'],p['학회·대회명'],p['수상명'],p['일자'],p['구분'],'',p['증빙']])));
    if (!nItems(s)) rows.push(base.concat(['','','','','','','','']));
  });
  const csv = '\\ufeff' + rows.map(r => r.map(x => '"' + String(x ?? '').replace(/"/g,'""') + '"').join(',')).join('\\r\\n');
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([csv], {type:'text/csv'})); a.download = \`장학금_제출목록_\${(D.now||'').slice(0,10)}.csv\`; a.click();
}
window.addEventListener('hashchange', route);
document.addEventListener('click', e => { const b = e.target.closest('.evbtn'); if (b) viewFile(b.dataset.u, b.dataset.l, b.dataset.s); });

(async () => {
  if (!CFG_TPL && API){ try { const r = await fetch(API + '?action=config'); CFG = Object.assign({}, CFG, await r.json()); } catch(e) {} }
  init();
  if (PREVIEW){ const n = $('#notice'); n.style.display=''; n.textContent = T('previewBanner'); $('#go').disabled = true; }
  route();
})();
</script>
</body>
</html>
`;
