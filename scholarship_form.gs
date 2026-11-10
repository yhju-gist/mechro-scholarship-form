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
         '논문수', '특허수', '수상수', '증빙폴더', '학생메모'],
  논문: ['접수번호', '상태', '학번', '성명', '순번', '저널명(학생)', '논문제목', 'DOI', 'Accept일',
         '역할(학생)', '저자수(학생)', '증빙1', '증빙2',
         // ↓ Claude 검증 칸
         '역할(AI)', '저자수(AI)', '저자목록(AI)', 'AI근거', 'JCR정식명', 'JCR연도', '최상위카테고리',
         '순위', '전체', '상위%', 'IF(JCR)', '다른카테고리', 'JCR상태', '검증일', '담당자확정', '담당자메모'],
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
  ['파일최대MB', '15', '파일 1개 최대 크기']
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
function ss_() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')); }

function settings_() {
  const rows = ss_().getSheetByName('설정').getDataRange().getValues().slice(1);
  const o = {}; rows.forEach(r => { if (r[0]) o[String(r[0]).trim()] = r[1]; });
  return o;
}

function parseDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const m = String(v).match(/(\d{4})\D(\d{1,2})\D(\d{1,2})(?:\D+(\d{1,2}):(\d{2}))?/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 23), +(m[5] || 59)) : null;
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
    if (req.action === 'uploadFile') return json_({ ok: true, data: uploadFile(req.args[0], req.args[1]) });
    if (req.action === 'submit') return json_({ ok: true, data: submit(req.args[0]) });
    if (req.action === 'adminList') return json_({ ok: true, data: adminList(req.args[0]) });
    if (req.action === 'adminFile') return json_({ ok: true, data: adminFile(req.args[0], req.args[1]) });
    throw new Error('알 수 없는 요청 / Unknown request');
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------------- 관리자(담당자) 화면용 ---------------- */
function checkAdmin_(key) {
  const k = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!k || String(key || '') !== k) { Utilities.sleep(1500); throw new Error('관리자 비밀번호가 맞지 않습니다.'); }
}

function sheetObjects_(name) {
  const v = ss_().getSheetByName(name).getDataRange().getDisplayValues();
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
  const b = f.getBlob();
  return { name: f.getName(), mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
}

/* 파일 하나 업로드 (제출 전에 하나씩 올려 용량 한도 회피) → 파일 ID 반환 */
function uploadFile(meta, dataUrl) {
  const s = settings_();
  const m = String(dataUrl).match(/^data:([^;]*);base64,(.*)$/);
  if (!m) throw new Error('파일을 읽지 못했습니다. / Could not read the file.');
  const bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > Number(s['파일최대MB'] || 15) * 1024 * 1024) throw new Error('파일이 너무 큽니다. / The file is too large.');
  const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
  const tmp = getOrCreate_(root, '_임시업로드');
  const safe = String(meta.name || 'file.pdf').replace(/[\\/:*?"<>|]/g, '_');
  const f = tmp.createFile(Utilities.newBlob(bytes, m[1] || 'application/pdf', safe));
  return f.getId();
}

function getOrCreate_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/* 최종 제출 */
function submit(p) {
  const s = settings_();
  const deadline = parseDate_(s['접수마감']);
  if (deadline && new Date() > deadline) throw new Error('접수가 마감되었습니다. / Submissions are closed.');
  ['학번', '성명', '과정', '지도교수'].forEach(k => { if (!p[k] || !String(p[k]).trim()) throw new Error(k + '을(를) 입력해 주세요. / Required field missing.'); });
  if (!/^\d{8}$/.test(String(p.학번).trim())) throw new Error('학번은 숫자 8자리로 입력해 주세요. / Student ID must be 8 digits.');

  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss = ss_();
    const id = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMdd-HHmmss') + '-' + String(p.학번).slice(-4);
    const now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    const who = (Session.getActiveUser().getEmail() || '');

    // 같은 학번 이전 제출은 '대체됨' 처리 (최신 제출만 유효)
    ['학생', '논문', '특허', '수상'].forEach(name => {
      const sh = ss.getSheetByName(name); const v = sh.getDataRange().getValues();
      const ci = v[0].indexOf('학번'), cs = v[0].indexOf('상태');
      for (let r = 1; r < v.length; r++) if (String(v[r][ci]) === String(p.학번) && v[r][cs] === '유효') sh.getRange(r + 1, cs + 1).setValue('대체됨');
    });

    // 증빙 파일 → 학번_성명 폴더로 이동·이름 정리
    const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID'));
    const folder = getOrCreate_(root, p.학번 + '_' + p.성명);
    const place = (fid, label) => {
      if (!fid) return '';
      const f = DriveApp.getFileById(fid);
      const ext = (f.getName().match(/\.[^.]+$/) || ['.pdf'])[0];
      f.setName(p.학번 + '_' + p.성명 + '_' + label + '_' + id + ext);
      f.moveTo(folder);
      return f.getUrl();
    };

    const papers = (p.papers || []).filter(x => x.journal);
    const patents = (p.patents || []).filter(x => x.type);
    const awards = (p.awards || []).filter(x => x.type);

    ss.getSheetByName('학생').appendRow([id, now, '유효', "'" + p.학번, p.성명, p.과정, p.재학학기, p.지도교수,
      p.이메일 || '', who, papers.length, patents.length, awards.length, folder.getUrl(), p.메모 || '']);

    papers.forEach((x, i) => ss.getSheetByName('논문').appendRow([id, '유효', "'" + p.학번, p.성명, i + 1,
      x.journal, x.title || '', x.doi || '', x.accept || '', x.role, Number(x.authors) || '',
      place(x.file1, '논문' + (i + 1)), place(x.file2, '논문' + (i + 1) + '_추가')]));

    patents.forEach((x, i) => ss.getSheetByName('특허').appendRow([id, '유효', "'" + p.학번, p.성명, i + 1,
      x.type, x.name || '', x.no || '', x.date || '', Number(x.share) || '', place(x.file, '특허' + (i + 1))]));

    awards.forEach((x, i) => ss.getSheetByName('수상').appendRow([id, '유효', "'" + p.학번, p.성명, i + 1,
      x.type, x.event || '', x.prize || '', x.date || '', place(x.file, '수상' + (i + 1))]));

    return { id: id, papers: papers.length, patents: patents.length, awards: awards.length, time: now };
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
body{margin:0;background:var(--bg);color:var(--ink);font-family:Pretendard,-apple-system,"Malgun Gothic",sans-serif;font-size:15px;line-height:1.55;-webkit-font-smoothing:antialiased}
.hero{padding-top:58px!important;background:linear-gradient(135deg,#1B1F23 0%,#2B3330 60%,#4F5B54 100%);color:#fff;padding:44px 20px 72px;position:relative;overflow:hidden}
.hero:after{content:"";position:absolute;right:-80px;top:-80px;width:320px;height:320px;border-radius:50%;background:radial-gradient(circle,rgba(220,39,21,.55),rgba(220,39,21,0) 70%)}
.wrap{max-width:820px;margin:0 auto;position:relative;z-index:1}
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
</style>
</head>
<body>
<div class="lang"><button type="button" id="lk" onclick="setLang('ko')">한국어</button><button type="button" id="le" onclick="setLang('en')">English</button></div>
<div class="hero"><div class="wrap">
  <div class="brand"><i></i><span data-t="brand"></span></div>
  <h1 id="title"></h1>
  <p data-t="lead"></p>
  <div class="chips">
    <span class="chip"><span data-t="deadline"></span> <b id="dl">-</b></span>
    <span class="chip"><span data-t="period"></span> <b id="range">-</b> <span data-t="acceptBasis"></span></span>
  </div>
</div></div>

<main id="app">
  <div class="notice" id="notice"></div>

  <section class="card">
    <h2><span class="num">1</span><span data-t="s1"></span></h2>
    <p class="desc" data-t="s1d"></p>
    <div class="grid">
      <div class="f" data-k="학번"><label><span data-t="sid"></span><span class="req">*</span></label><input id="학번" inputmode="numeric" maxlength="8" data-ph="sidPh"><div class="err" data-t="sidErr"></div></div>
      <div class="f" data-k="성명"><label><span data-t="name"></span><span class="req">*</span></label><input id="성명" data-ph="namePh"><div class="err" data-t="nameErr"></div></div>
      <div class="f" data-k="과정"><label><span data-t="course"></span><span class="req">*</span></label>
        <div class="seg two"><input type="radio" name="과정" id="c1" value="박사"><label for="c1" data-t="phd"></label><input type="radio" name="과정" id="c2" value="석박통합"><label for="c2" data-t="integ"></label></div>
        <div class="err" data-t="courseErr"></div></div>
      <div class="f"><label><span data-t="sem"></span><span class="opt" data-t="semOpt"></span></label><input id="재학학기" inputmode="numeric" data-ph="semPh"></div>
      <div class="f" data-k="지도교수"><label><span data-t="adv"></span><span class="req">*</span></label><select id="지도교수Sel"></select><input id="지도교수" data-ph="advPh" style="margin-top:8px;display:none"><div class="err" data-t="advErr"></div></div>
      <div class="f"><label><span data-t="email"></span><span class="opt" data-t="optional"></span></label><input id="이메일" type="email" placeholder="you@gm.gist.ac.kr"></div>
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
      <div class="f"><label>DOI<span class="opt" data-t="ifAny"></span></label><input name="doi" placeholder="10.xxxx/xxxxx"></div>
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
        <label class="file"><input type="file" name="file1" accept=".pdf,image/*"><span class="ico">PDF</span><span class="t" data-t="file1Ph"></span></label><div class="err" data-t="fileErr"></div></div>
      <div class="f full"><label><span data-t="file2"></span><span class="opt" data-t="file2Opt"></span></label>
        <label class="file"><input type="file" name="file2" accept=".pdf,image/*"><span class="ico">＋</span><span class="t" data-t="choose"></span></label></div>
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
        <label class="file"><input type="file" name="file" accept=".pdf,image/*"><span class="ico">PDF</span><span class="t" data-t="choose"></span></label><div class="err" data-t="fileErr"></div></div>
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
        <label class="file"><input type="file" name="file" accept=".pdf,image/*"><span class="ico">PDF</span><span class="t" data-t="choose"></span></label><div class="err" data-t="fileErr"></div></div>
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
  course:'과정',phd:'박사',integ:'석박통합',courseErr:'과정을 골라 주세요.',sem:'재학학기',semOpt:'(이번 학기 기준)',semPh:'예) 4',adv:'지도교수',advPh:'지도교수 성함',advErr:'지도교수를 입력해 주세요.',email:'이메일',optional:'(선택)',
  s2:'SCI 논문',s2d:'Accept 이후 논문만 넣어 주세요. 저널 순위·IF는 담당자가 JCR에서 찾아 입력합니다.',addPaper:'＋ 논문 추가',
  s3:'특허',s3d:'출원·등록 모두 가능합니다. 기여도는 발명자 지분(%)을 적어 주세요.',addPatent:'＋ 특허 추가',
  s4:'학회·경진대회 수상',s4d:'국내·해외(국제) 학회와 경진대회 수상 실적을 넣어 주세요.',addAward:'＋ 수상 추가',
  s5:'확인',memo:'담당자에게 남길 말',memoPh:'예) 논문 2는 아직 게재 전이라 Accept 메일을 함께 올렸습니다.',agree:'입력한 실적과 증빙이 사실이며, 증빙(논문 PDF 등)이 장학금 심사에 쓰이는 것에 동의합니다.',
  cntP:'논문',cntT:'특허',cntA:'수상',submit:'제출하기',paper:'논문',patent:'특허',award:'수상',del:'삭제',
  journal:'저널명 (전체 이름)',journalPh:'예) Advanced Functional Materials',journalHint:'약어 대신 전체 이름으로 적어 주세요.',journalErr:'저널명을 입력해 주세요.',ptitle:'논문 제목',ptitlePh:'논문 제목',ifAny:'(있으면)',
  accept:'Accept 날짜',acceptOpt:'(일자가 불명확한 경우 1일로)',acceptErr:'Accept 날짜를 입력해 주세요.',role:'나의 역할',rFA:'제1저자',rCoFA:'공동 제1저자',rCA:'교신저자',rETC:'그 외 공저자',roleErr:'역할을 골라 주세요.',
  authors:'총 저자 수 (본인 포함)',authorsPh:'예) 5',authorsErr:'저자 수를 입력해 주세요.',file1:'증빙 ① 논문 첫 페이지 또는 PDF',file1Ph:'저자·소속이 보이는 첫 페이지를 올려 주세요',fileErr:'증빙 파일을 올려 주세요.',
  file2:'증빙 ② Accept 메일 등',file2Opt:'(게재 전이면)',choose:'파일 선택',
  ptype:'구분',kApp:'국내 출원',kReg:'국내 등록',iApp:'국제 출원',iReg:'국제 등록',typeErr:'구분을 골라 주세요.',pname:'발명의 명칭',pno:'출원·등록번호',pdate:'출원·등록일',share:'나의 기여도(%)',sharePh:'예) 30',shareErr:'1~100 사이로 입력해 주세요.',pfile:'증빙 (출원서·등록증, 기여도 표시)',
  aDom:'국내 학회·대회',aInt:'해외(국제) 학회·대회',aevent:'학회·대회명',aprize:'수상명',aprizePh:'예) Best Paper Award',adate:'수상일',afile:'증빙 (상장 등)',
  contact:'문의',closed:'접수가 마감되었습니다',emptyP:'논문 실적이 없으면 비워 두세요.',emptyT:'특허 실적이 없으면 비워 두세요.',emptyA:'수상 실적이 없으면 비워 두세요.',
  selProf:'선택해 주세요',etcProf:'목록에 없음 (직접 입력)',tooBig:'파일이 {mb}MB를 넘습니다. 첫 페이지만 따로 저장해 올려 주세요.',needOne:'논문·특허·수상 중 하나 이상 입력해 주세요.',
  sending:'제출하는 중',uploading:'증빙 파일을 올리고 있습니다…',upN:'증빙 파일 {i} / {n} 올리는 중',saving:'접수 기록 중…',noClose:'창을 닫지 마세요.',
  done:'제출되었습니다',rcpt:'접수번호',doneCnt:'논문 {p}편 · 특허 {t}건 · 수상 {a}건',doneNote:'{time} 접수 · 고칠 것이 있으면 같은 학번으로 다시 제출하세요(마지막 제출만 반영).',fail:'제출하지 못했습니다',back:'돌아가기',previewBanner:'미리보기 화면입니다 — 담당자가 접수를 열면 제출할 수 있습니다.',previewErr:'아직 접수가 열리지 않았습니다.'},
 en:{brand:'GIST · Dept. of Mechanical and Robotics Engineering',lead:'Enter your papers, patents and awards and upload the evidence. The office verifies them against the documents and JCR.',deadline:'Deadline',period:'Eligible period',acceptBasis:'(by acceptance date)',
  s1:'Basic information',s1d:'If you submit again with the same student ID, only your latest submission counts.',sid:'Student ID',sidPh:'e.g. 20244040',sidErr:'Enter your 8-digit student ID.',name:'Name',namePh:'Name as in your passport',nameErr:'Enter your name.',
  course:'Program',phd:'Ph.D.',integ:'Integrated M.S.–Ph.D.',courseErr:'Choose your program.',sem:'Semester',semOpt:'(current semester)',semPh:'e.g. 4',adv:'Advisor',advPh:"Advisor's name",advErr:'Enter your advisor.',email:'Email',optional:'(optional)',
  s2:'SCI papers',s2d:'Only papers accepted for publication. You do not need to enter the JCR rank or IF — the office looks them up.',addPaper:'＋ Add a paper',
  s3:'Patents',s3d:'Both applications and registrations count. Enter your inventor share (%).',addPatent:'＋ Add a patent',
  s4:'Conference / competition awards',s4d:'Awards from domestic or international conferences and competitions.',addAward:'＋ Add an award',
  s5:'Confirmation',memo:'Note to the office',memoPh:'e.g. Paper 2 is not yet published, so I attached the acceptance email.',agree:'I confirm that the information and evidence are true, and agree that the evidence (paper PDFs etc.) will be used for the scholarship review.',
  cntP:'Papers',cntT:'Patents',cntA:'Awards',submit:'Submit',paper:'Paper',patent:'Patent',award:'Award',del:'Remove',
  journal:'Journal (full title)',journalPh:'e.g. Advanced Functional Materials',journalHint:'Use the full journal title, not the abbreviation.',journalErr:'Enter the journal title.',ptitle:'Paper title',ptitlePh:'Paper title',ifAny:'(if any)',
  accept:'Acceptance date',acceptOpt:'(if the day is unknown, use the 1st)',acceptErr:'Enter the acceptance date.',role:'Your role',rFA:'First author',rCoFA:'Co-first author',rCA:'Corresponding',rETC:'Other co-author',roleErr:'Choose your role.',
  authors:'Total number of authors (incl. you)',authorsPh:'e.g. 5',authorsErr:'Enter the number of authors.',file1:'Evidence ① First page of the paper or full PDF',file1Ph:'Upload the first page showing authors and affiliations',fileErr:'Upload the evidence file.',
  file2:'Evidence ② Acceptance email etc.',file2Opt:'(if not yet published)',choose:'Choose file',
  ptype:'Type',kApp:'Korean application',kReg:'Korean registration',iApp:'International application',iReg:'International registration',typeErr:'Choose the type.',pname:'Title of invention',pno:'Application / registration no.',pdate:'Filing / registration date',share:'Your share (%)',sharePh:'e.g. 30',shareErr:'Enter a number from 1 to 100.',pfile:'Evidence (application / certificate showing shares)',
  aDom:'Domestic conference / competition',aInt:'International conference / competition',aevent:'Conference / competition',aprize:'Award title',aprizePh:'e.g. Best Paper Award',adate:'Award date',afile:'Evidence (certificate etc.)',
  contact:'Contact',closed:'Submissions are closed',emptyP:'Leave empty if you have no papers.',emptyT:'Leave empty if you have no patents.',emptyA:'Leave empty if you have no awards.',
  selProf:'Select',etcProf:'Not in the list (type it)',tooBig:'The file is larger than {mb}MB. Please save and upload only the first page.',needOne:'Enter at least one paper, patent or award.',
  sending:'Submitting',uploading:'Uploading evidence files…',upN:'Uploading file {i} of {n}',saving:'Saving your submission…',noClose:'Please do not close this window.',
  done:'Submitted',rcpt:'Receipt no.',doneCnt:'{p} paper(s) · {t} patent(s) · {a} award(s)',doneNote:'Received {time}. To make changes, submit again with the same student ID (only the latest counts).',fail:'Submission failed',back:'Back',previewBanner:'Preview — submissions open once the office connects the form.',previewErr:'Submissions are not open yet.'}
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
}
function setLang(l){
  LANG = l; document.documentElement.lang = l;
  $('#lk').classList.toggle('on', l==='ko'); $('#le').classList.toggle('on', l==='en');
  translate(document);
  $('#title').textContent = (l==='en' && CFG.titleEn) || CFG.title;
  const n = PREVIEW ? T('previewBanner') : ((l==='en' && CFG.noticeEn) || CFG.notice); $('#notice').textContent = n; $('#notice').style.display = n ? '' : 'none';
  const c = (l==='en' && CFG.contactEn) || CFG.contact; $('#contact').textContent = c ? T('contact') + ': ' + c : '';
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
  $('#dl').textContent = dl;
  $('#range').textContent = (CFG.from && CFG.to) ? CFG.from + ' ~ ' + CFG.to : '-';
  const sel = $('#지도교수Sel');
  sel.onchange = () => { const etc = sel.value==='__etc'; $('#지도교수').style.display = etc?'block':'none'; $('#지도교수').value = etc ? '' : sel.value; };
  if (CFG.closed){
    $('#app').innerHTML = '<section class="card closed"><h2 style="justify-content:center" data-t="closed"></h2><p class="desc" id="contact"></p></section>';
    document.querySelector('.bar').style.display='none';
  } else addPaper();
  setLang(LANG);
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
    const ok = !!inp.files[0];
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
function bad(scope, key){ const f = scope.querySelector(\`.f[data-k="\${key}"]\`); if (f) f.classList.add('bad'); return f; }

function collect(){
  document.querySelectorAll('.bad').forEach(b=>b.classList.remove('bad'));
  const errs = []; const P = {lang: LANG};
  P.학번 = $('#학번').value.trim(); if(!/^\\d{8}$/.test(P.학번)) errs.push(bad(document,'학번'));
  P.성명 = $('#성명').value.trim(); if(!P.성명) errs.push(bad(document,'성명'));
  const c = document.querySelector('input[name="과정"]:checked'); P.과정 = c ? c.value : ''; if(!P.과정) errs.push(bad(document,'과정'));
  P.재학학기 = $('#재학학기').value.trim();
  P.지도교수 = $('#지도교수').value.trim(); if(!P.지도교수) errs.push(bad(document,'지도교수'));
  P.이메일 = $('#이메일').value.trim(); P.메모 = $('#메모').value.trim();
  const files = [];
  P.papers = [...document.querySelectorAll('#papers .item')].map(it => {
    const x = {journal:val(it,'journal'), title:val(it,'title'), doi:val(it,'doi'), accept:val(it,'accept'), role:radio(it), authors:val(it,'authors')};
    if(!x.journal) errs.push(bad(it,'journal')); if(!x.accept) errs.push(bad(it,'accept')); if(!x.role) errs.push(bad(it,'role'));
    if(!(+x.authors>=1)) errs.push(bad(it,'authors'));
    const f1 = fileOf(it,'file1'), f2 = fileOf(it,'file2'); if(!f1) errs.push(bad(it,'file1'));
    if(f1) files.push([x,'file1',f1]); if(f2) files.push([x,'file2',f2]); return x;
  });
  P.patents = [...document.querySelectorAll('#patents .item')].map(it => {
    const x = {type:radio(it), name:val(it,'name'), no:val(it,'no'), date:val(it,'date'), share:val(it,'share')};
    if(!x.type) errs.push(bad(it,'type')); if(!(+x.share>0 && +x.share<=100)) errs.push(bad(it,'share'));
    const f = fileOf(it,'file'); if(!f) errs.push(bad(it,'file')); else files.push([x,'file',f]); return x;
  });
  P.awards = [...document.querySelectorAll('#awards .item')].map(it => {
    const x = {type:radio(it), event:val(it,'event'), prize:val(it,'prize'), date:val(it,'date')};
    if(!x.type) errs.push(bad(it,'type'));
    const f = fileOf(it,'file'); if(!f) errs.push(bad(it,'file')); else files.push([x,'file',f]); return x;
  });
  return {P, files, errs: errs.filter(Boolean)};
}

function readUrl(f){ return new Promise((ok,no)=>{ const r=new FileReader(); r.onload=()=>ok(r.result); r.onerror=no; r.readAsDataURL(f); }); }
function gs(fn, ...a){
  if (CFG_TPL && window.google && google.script) return new Promise((ok,no)=>{ const run = google.script.run.withSuccessHandler(ok).withFailureHandler(no); run[fn](...a); });
  if (!API) return Promise.reject(new Error(T('previewErr')));
  return fetch(API, {method:'POST', headers:{'Content-Type':'text/plain;charset=utf-8'}, body: JSON.stringify({action:fn, args:a})})
    .then(r => r.json()).then(j => { if (j.ok) return j.data; throw new Error(j.error); });
}
function esc(s){ return String(s).replace(/[&<>"]/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

async function send(){
  const {P, files, errs} = collect();
  if (errs.length){ errs[0].scrollIntoView({behavior:'smooth', block:'center'}); return; }
  if (!P.papers.length && !P.patents.length && !P.awards.length){ alert(T('needOne')); return; }
  if (!$('#agree').checked){ $('#agreeBox').scrollIntoView({behavior:'smooth',block:'center'}); $('#agreeBox').style.color='var(--red)'; return; }
  $('#go').disabled = true;
  const box = $('#box'); $('#modal').classList.add('on');
  box.innerHTML = \`<h3>\${T('sending')}</h3><p id="pm">\${T('uploading')}</p><div class="prog"><i id="pb"></i></div><p style="font-size:12px">\${T('noClose')}</p>\`;
  try{
    for (let i=0;i<files.length;i++){
      const [obj,key,f] = files[i];
      $('#pm').textContent = T('upN',{i:i+1,n:files.length}) + ' · ' + f.name;
      obj[key] = await gs('uploadFile', {name:f.name}, await readUrl(f));
      $('#pb').style.width = Math.round((i+1)/(files.length+1)*100)+'%';
    }
    $('#pm').textContent = T('saving');
    const r = await gs('submit', P);
    $('#pb').style.width = '100%';
    box.innerHTML = \`<div class="big">✓</div><h3>\${T('done')}</h3><p>\${T('rcpt')} <b>\${esc(r.id)}</b></p><p>\${T('doneCnt',{p:r.papers,t:r.patents,a:r.awards})}</p><p style="margin-top:12px;font-size:13px;color:var(--mute)">\${T('doneNote',{time:esc(r.time)})}</p>\`;
  }catch(e){
    box.innerHTML = \`<h3>\${T('fail')}</h3><p>\${esc((e && e.message) || e)}</p><p style="margin-top:14px"><button class="go" onclick="document.getElementById('modal').classList.remove('on')">\${T('back')}</button></p>\`;
    $('#go').disabled = false;
  }
}
/* ================= 담당자(관리자) 화면 ================= */
let ADM = { key:'', data:null };
const ss_get = k => { try { return sessionStorage.getItem(k) || ''; } catch(e) { return ''; } };
const ss_set = (k,v) => { try { v ? sessionStorage.setItem(k,v) : sessionStorage.removeItem(k); } catch(e) {} };
function route(){
  const isA = location.hash === '#admin';
  $('#app').hidden = isA; document.querySelector('.bar').style.display = isA ? 'none' : '';
  $('#admin').hidden = !isA; document.querySelector('.chips').style.display = isA ? 'none' : '';
  $('#title').textContent = isA ? '장학금 연구실적 — 담당자 화면' : ((LANG==='en' && CFG.titleEn) || CFG.title);
  document.querySelector('.lang').style.display = isA ? 'none' : '';
  document.querySelector('[data-t=lead]').textContent = isA ? '학생 제출 목록을 보고 증빙 PDF를 바로 열어 봅니다.' : T('lead');
  document.querySelector('[data-t=brand]').textContent = isA ? 'GIST · 기계로봇공학과' : T('brand');
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
    const ev = (u,label) => u ? \`<button class="btn sm evbtn" type="button" data-u="\${esc(u)}" data-l="\${esc(s['학번']+' '+s['성명']+' '+label)}">📄 \${label}</button>\` : '<span class="sub">없음</span>';
    const papers = s.papers.map(p => \`<tr><td>\${esc(p['순번'])}</td><td><b>\${esc(p['저널명(학생)'])}</b><div class="sub">\${esc(p['논문제목'])}</div><div class="sub">\${esc(p['DOI'])}</div></td><td>\${esc(p['Accept일'])}</td><td>\${esc(p['역할(학생)'])}</td><td>\${esc(p['저자수(학생)'])}</td><td>\${ev(p['증빙1'],'증빙')} \${p['증빙2']?ev(p['증빙2'],'추가'):''}</td><td>\${p['IF(JCR)']?esc(p['순위']+'/'+p['전체']+' · IF '+p['IF(JCR)']):'<span class="sub">검증 전</span>'}</td></tr>\`).join('');
    const pats = s.patents.map(p => \`<tr><td>\${esc(p['순번'])}</td><td>\${esc(p['구분'])}</td><td>\${esc(p['명칭'])}<div class="sub">\${esc(p['출원·등록번호'])} \${esc(p['일자'])}</div></td><td>\${esc(p['기여도(%)'])}%</td><td>\${ev(p['증빙'],'증빙')}</td></tr>\`).join('');
    const aws = s.awards.map(p => \`<tr><td>\${esc(p['순번'])}</td><td>\${esc(p['구분'])}</td><td>\${esc(p['학회·대회명'])}<div class="sub">\${esc(p['수상명'])} \${esc(p['일자'])}</div></td><td>\${ev(p['증빙'],'증빙')}</td></tr>\`).join('');
    return \`<div class="a-row" id="ar\${i}">
      <div class="a-head" onclick="this.parentElement.classList.toggle('open')">
        <div><div class="nm">\${esc(s['성명'])}</div><div class="sub">\${esc(s['학번'])}</div></div>
        <div>\${pil}<span class="pill">\${esc(s['과정'])}</span><span class="sub">\${esc(s['재학학기'] ? s['재학학기']+'학기' : '')}</span></div>
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
async function viewFile(url, label){
  const v = $('#viewer'); v.classList.add('on'); $('#vName').textContent = label + ' — 불러오는 중…'; $('#vBody').innerHTML = '';
  try{
    const f = await gs('adminFile', ADM.key, url);
    const bin = atob(f.data), arr = new Uint8Array(bin.length); for (let i=0;i<bin.length;i++) arr[i]=bin.charCodeAt(i);
    if (VURL) URL.revokeObjectURL(VURL);
    VURL = URL.createObjectURL(new Blob([arr], {type: f.mime}));
    $('#vName').textContent = f.name; $('#vOpen').href = VURL; $('#vDown').href = VURL; $('#vDown').download = f.name;
    if (f.mime.startsWith('image/')) { $('#vBody').innerHTML = \`<img src="\${VURL}" alt="">\`; return; }
    try { await renderPdf(arr); } catch(err) { $('#vBody').innerHTML = \`<iframe src="\${VURL}" title="증빙"></iframe>\`; }
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
document.addEventListener('click', e => { const b = e.target.closest('.evbtn'); if (b) viewFile(b.dataset.u, b.dataset.l); });

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
