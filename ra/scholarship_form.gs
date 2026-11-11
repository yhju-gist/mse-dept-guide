/**
 * 신소재공학과 RA 장학금 연구실적 접수 (Google Apps Script 웹앱) — 기계로봇공학과 시스템을 학과용으로 옮김(2026-11-11)
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
  ['평가명', '2026학년도 RA(GUP) 장학금 박사 연구실적 우수자', '웹페이지 제목에 표시'],
  ['평가명(영문)', '2026 RA (GUP) Scholarship — Ph.D. Research Excellence', '영어 화면 제목'],
  ['접수마감', '2026-12-04 23:59', '이 시각 이후 제출 막힘 (yyyy-MM-dd HH:mm)'],
  ['실적인정 시작', '2025-12-01', 'Accept일이 이 날짜 이전이면 더블체크 표시'],
  ['실적인정 끝', '2026-11-30', 'Accept일이 이 날짜 이후면 더블체크 표시'],
  ['지도교수 목록', '고흥조/Heung Cho Ko, 권인찬/Inchan Kwon, 김동유/Dong-Yu Kim, 김봉중/Bong Joong Kim, 김현호/Hyun Ho Kim, 김호범/Hobeom Kim, 박지웅/Ji-Woong Park, 엄광섭/Kwang Sup Eom, 연한울/Hanwool Yeon, 유승준/Seung Joon Yoo, 윤명한/Myung-Han Yoon, 이상한/Sanghan Lee, 이은지/Eunji Lee, 이재영/Jaeyoung Lee, 이주형/Joo-Hyoung Lee, 이한얼/Haneol Lee, 정건영/Gun-Young Jung, 조지영/Ji Young Jo, 태기융/Giyoong Tae, 하민정/Minjeong Ha', '쉼표로 구분, "한글/영문" (신소재공학과 홈페이지 전임교수 2026-11-11 기준)'],
  ['담당자', '신소재공학과 학과사무실 이혜인 (T.2303)', '하단 문의처'],
  ['담당자(영문)', 'Hyein Lee, Dept. of Materials Science and Engineering (ext. 2303)', '영어 화면 하단 문의처'],
  ['안내문', '연구실적 부문 대상: 박사 5학기 이상 · 석박통합 7학기 이상. 논문 첫 페이지나 PDF를 꼭 올려 주세요. 특허는 등록만 인정됩니다.', '입력 화면 상단 안내'],
  ['안내문(영문)', 'Research category: Ph.D. semester 5+ or integrated semester 7+. Upload the first page or PDF of each paper. Only registered patents count.', '영어 화면 상단 안내'],
  ['파일최대MB', '15', '파일 1개 최대 크기'],
  ['허용 메일 도메인', 'gist.ac.kr, gm.gist.ac.kr', '본인 확인 메일을 받을 수 있는 주소 (쉼표로 구분)'],
  ['문의 메일', 'toto9430@gist.ac.kr', '인증 메일에 답장하면 이 주소로 감 (담당자 GIST 메일)']
];

/* ---------------- 초기 설정 ---------------- */
/* 담당자 전용 함수 보호: 웹앱 화면(…/exec)에서 google.script.run 으로 불러도 실행되지 않게 함.
   담당자가 Apps Script 편집기·시트 메뉴에서 직접 실행할 때만 통과 (실행 계정 = 소유자 계정) */
function ownerOnly_() {
  let me = '', who = '';
  try { me = Session.getEffectiveUser().getEmail(); who = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!me || who !== me) throw new Error('담당자 계정에서 Apps Script 편집기로만 실행할 수 있습니다.');
}

function setup() {
  ownerOnly_();
  const props = PropertiesService.getScriptProperties();
  let ss = SpreadsheetApp.getActiveSpreadsheet();   // 시트에서 [확장 프로그램 → Apps Script]로 만든 경우
  if (ss) {
    props.setProperty('SHEET_ID', ss.getId());
  } else if (props.getProperty('SHEET_ID')) {
    ss = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  } else {
    ss = SpreadsheetApp.create('신소재공학과 RA장학금 실적 접수');
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
    const f = DriveApp.createFolder('신소재 RA장학금 증빙(학생 제출)');
    props.setProperty('FOLDER_ID', f.getId());
  }
  if (!props.getProperty('ADMIN_KEY')) makeAdminKey();
  tokKey_(true);
  Logger.log('시트: ' + ss.getUrl());
  Logger.log('증빙 폴더: ' + DriveApp.getFolderById(props.getProperty('FOLDER_ID')).getUrl());
}

/* 관리자 비밀번호 새로 만들기 — 실행하면 실행 로그에 새 비밀번호가 나옵니다(이전 비밀번호는 무효) */
function makeAdminKey() {
  ownerOnly_();
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789';
  let k = ''; for (let i = 0; i < 10; i++) k += chars[Math.floor(Math.random() * chars.length)];
  PropertiesService.getScriptProperties().setProperty('ADMIN_KEY', k);
  Logger.log('관리자 비밀번호: ' + k);   // 비밀번호는 실행 로그에만 남기고 돌려주지 않음
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('RA 장학금 실적')
    .addItem('웹페이지 주소 보기', 'showUrl')
    .addItem('중복·누락 점검', 'checkDuplicates')
    .addToUi();
}

function showUrl() {
  ownerOnly_();
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
   주소만 열면: 안내 글자만 보여 줌.
   ※ 보안: 이 주소에서 HTML 화면을 띄우면 그 화면의 개발자 도구에서 google.script.run 으로
     서버 함수를 직접 부를 수 있으므로, HTML 화면(예비용)은 쓰지 않는다 (2026-11-11). */
const HOME_URL = 'https://yhju-gist.github.io/mse-dept-guide/ra/';
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'config') return json_(publicCfg_(e));
  return ContentService.createTextOutput('장학금 연구실적 제출은 아래 홈페이지에서 해 주세요.\nPlease submit at:\n\n' + HOME_URL);
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
  const rt = String(s['문의 메일'] || 'toto9430@gist.ac.kr').trim().replace(/[<>"'\s]/g, '');
  const html = '<div style="font-family:Apple SD Gothic Neo,Malgun Gothic,sans-serif;max-width:480px;margin:0 auto;padding:28px;border:1px solid #e5e5e5;border-radius:14px">' +
    '<div style="font-size:13px;color:#5B3FD6;font-weight:700">GIST 신소재공학과 · Dept. of Materials Science and Engineering</div>' +
    '<h2 style="margin:10px 0 6px;font-size:20px;color:#1B1F23">' + String(s['평가명'] || '장학금 연구실적 제출') + '</h2>' +
    '<p style="color:#555;font-size:14px;margin:0 0 18px">본인 확인 인증번호입니다. 10분 안에 입력해 주세요.<br>Your verification code (valid for 10 minutes):</p>' +
    '<div style="font-size:34px;letter-spacing:10px;font-weight:800;color:#1B1F23;background:#F5F6F4;border-radius:10px;padding:16px;text-align:center">' + code + '</div>' +
    '<p style="color:#888;font-size:12px;margin-top:18px">직접 요청하지 않았다면 이 메일을 무시하세요. If you did not request this, please ignore this email.<br>문의 Contact: ' + String(s['담당자'] || '') + (rt ? ' · <a href="mailto:' + rt + '" style="color:#5B3FD6">' + rt + '</a>' : '') + '<br>이 메일에 답장하면 담당자에게 전달됩니다. Replies to this email go to the office.</p></div>';
  MailApp.sendEmail({ to: e, name: 'GIST 신소재공학과 RA장학금 접수', subject: '[GIST 신소재공학과] 인증번호 ' + code + ' (Verification code)',
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

/* 이전인정논문 탭 (없으면 빈 목록) — 관리자 RAWDATA 실시간에서 1회 인정 확인용 */
function prevPapers_() {
  const sh = ss_().getSheetByName('이전인정논문'); if (!sh) return [];
  const v = sh.getDataRange().getDisplayValues(), h = v[0];
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
    prev: prevPapers_(),   // RAWDATA 실시간: 이전 인정 논문 중복 확인용
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
  if (!kind) throw new Error('증빙은 PDF 또는 사진(JPG·PNG) 파일만 올릴 수 있습니다. / Only PDF or photo (JPG, PNG) files are accepted.');
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
  ownerOnly_();
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
