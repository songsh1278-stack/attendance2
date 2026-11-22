/**
 * 교육부 출석부 — Google Apps Script 웹앱 (서버 코드)
 *
 * 이 스크립트가 붙어 있는 구글 시트의 탭에 데이터를 저장합니다.
 *   명단 / 출석 / 달란트 / 반 설정  : 사람이 직접 봐도 되는 탭
 *   _설정                           : 관리자 비밀번호, 임원 번호, 담임 번호, 등반 기록 (숨김 탭)
 *
 * 권한 3단계
 *   총관리자     : _설정 B1 비밀번호. 모든 부서, 임원 번호 발급, 임원·공동관리자 이메일 등록
 *   공동관리자   : 총관리자가 이메일+개인 비밀번호로 등록. 총관리자와 같은 권한(이메일 등록·수정·삭제만 제외)
 *   팀별 관리자  : 임원 번호(부서 고유번호). 그 부서 전체, 담임 번호 발급, 연말 등반
 *   담임         : 담임 번호. 자기 반(학년) 학생만 출석·달란트·명단·주소록 관리
 * 탭이 없으면 처음 실행할 때 자동으로 만듭니다.
 */

/**
 * 앱 주소 설정
 *   PUBLIC_URL : 선생님들께 보낼 바로가기 페이지 주소. 앱의 [주소 복사]가 이 주소를 복사해요.
 *                _설정 탭 F1 칸에 주소를 적으면 그 주소를 쓰고, 둘 다 비어 있으면 이 앱의 웹앱 주소(…/exec)를 씁니다.
 *   ICON_URL   : 브라우저 탭 아이콘 그림 주소. _설정 탭 F2 칸에 적으면 그 주소를 씁니다.
 */
const PUBLIC_URL = 'https://songsh1278-stack.github.io/attendance2/';
const ICON_URL = 'https://songsh1278-stack.github.io/attendance2/icon-192.png';

function cfgCell_(a1) { try { const sh = cfgSheet_(); return sh ? String(sh.getRange(a1).getValue() || '').trim() : ''; } catch (e) { return ''; } }
/** 선생님들께 보낼 앱 주소 */
function publicUrl_() {
  const own = cfgCell_('F1');
  if (/^https?:\/\//.test(own)) return own;
  return PUBLIC_URL || ScriptApp.getService().getUrl() || '';
}
function iconUrl_() { const own = cfgCell_('F2'); return /^https?:\/\//.test(own) ? own : ICON_URL; }

const SHEET = { roster: '명단', att: '출석', tal: '달란트', cls: '반 설정', cfg: '_설정', book: '주소록', memo: '출결메모', events: '일정', hist: '이력', lib: '자료실' };
const BCOLS = ['ID', '이름', '부서', '반', '학생 연락처', '부모 연락처', '보호자', '주소'];
const RCOLS = ['ID', '부서', '반', '분반', '이름', '구분', '생일', '직분', '새친구', '재적', '메모', '사진', '기타'];
const WEEK_COL = 7; // G열부터 주일 칸
const TEAMS = {
  infant:  { label: '영아교육팀',   groups: [['0','0세'],['1','1세'],['2','2세'],['3','3세'],['4','4세']] },
  toddler: { label: '유아교육팀',   groups: [['5','5세'],['6','6세'],['7','7세']] },
  youth:   { label: '유소년교육팀', groups: [['1','1학년'],['2','2학년'],['3','3학년'],['4','4학년'],['5','5학년'],['6','6학년']] },
  teen:    { label: '중고등교육팀', groups: [['m1','중1'],['m2','중2'],['m3','중3'],['h1','고1'],['h2','고2'],['h3','고3']] },
};
const TEAM_ORDER = ['infant', 'toddler', 'youth', 'teen'];
/**
 * 영아교육팀은 나이와 관계없이 반을 편성한다.
 *  - 반 목록은 _설정 teamSettings[부서].classes = [[키, 반 이름], ...] (없으면 나이반 0세~4세 그대로)
 *  - 나이는 사람마다 따로(age: '0'~'4'). 아이는 늘 어느 한 반에 속한다(미배정 없음)
 * 유아·유소년·중고등은 지금처럼 반 = 나이·학년.
 */
const FLEX_TEAMS = ['infant'];
let TEAM_SETTINGS_CACHE_ = null;
function isFlex_(team) { return FLEX_TEAMS.indexOf(team) >= 0; }
function groupsOf_(team) {
  if (!TEAMS[team]) return [];
  if (!TEAM_SETTINGS_CACHE_) TEAM_SETTINGS_CACHE_ = cfgAll_().teamSettings || {};
  const c = (TEAM_SETTINGS_CACHE_[team] || {}).classes;
  if (isFlex_(team)) return Array.isArray(c) && c.length ? c : TEAMS[team].groups;
  // 유아·유소년·중고등: 나이·학년 순서는 고정, 반 이름만 바꿀 수 있다
  const names = {}; (Array.isArray(c) ? c : []).forEach(([k, l]) => { if (l) names[k] = l; });
  return TEAMS[team].groups.map(([k, l]) => [k, names[k] || l]);
}
/** 나이 키('0'~'4'). "3세", 3 도 받아 준다. 영아교육팀이 아니면 '' */
function ageKey_(team, v) {
  if (!isFlex_(team)) return '';
  const s = String(v == null ? '' : v).replace(/\s|세/g, '');
  return FLEX_TEAMS.some(t => TEAMS[t].groups.some(x => x[0] === s)) ? s : '';
}

/* ---------------- 웹앱 진입 ---------------- */
function doGet(e) {
  // 구글 주소(…/exec)로 들어오면 항상 바로가기 주소(PUBLIC_URL)로만 보낸다.
  // (예전처럼 구글 화면에서 직접 열리는 길은 막았다)
  const url = publicUrl_();
  if (!/^https?:\/\//.test(url) || /script\.google\.com|googleusercontent\.com/.test(url))
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px">출석부 주소가 아직 설정되지 않았어요. 관리자에게 문의해 주세요.</p>').setTitle('신당중앙교회 교육팀 출석부');
  return redirectPage_(url);
}
function redirectPage_(url) {
  const safe = String(url).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  const html = '<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<style>body{font-family:sans-serif;margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;background:#F1F5F2;color:#17302B;text-align:center}' +
    '.b{padding:24px}h2{margin:0 0 8px}p{margin:0 0 20px;color:#5F726D}' +
    'a{display:inline-block;background:#18856B;color:#fff;text-decoration:none;font-weight:700;border-radius:12px;padding:14px 26px;font-size:17px}</style></head>' +
    '<body><div class="b"><h2>신당중앙교회 교육팀 출석부</h2><p>출석부 주소로 이동하고 있어요.<br>움직이지 않으면 아래 버튼을 눌러 주세요.</p>' +
    '<a href="' + safe + '" target="_top" rel="noopener">출석부 열기</a></div>' +
    '<script>try{window.top.location.replace(' + JSON.stringify(String(url)).replace(/</g, '\\u003c') + ')}catch(e){}</script></body></html>';
  const out = HtmlService.createHtmlOutput(html).setTitle('신당중앙교회 교육팀 출석부')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  const icon = iconUrl_(); if (icon) { try { out.setFaviconUrl(icon); } catch (x) {} }
  return out;
}
function serveApp_() {
  // 시트 점검은 화면이 뜬 뒤 boot 에서 한다 (첫 화면을 빨리 띄우려고)
  // 화면 파일 이름을 조금 다르게 만들었어도 찾도록 여러 이름을 차례로 시도한다.
  const names = ['Index', 'index', 'Index.html', 'index.html', 'INDEX'];
  let out = null;
  for (let i = 0; i < names.length && !out; i++) {
    try { out = HtmlService.createHtmlOutputFromFile(names[i]); } catch (e) {}
  }
  if (!out) {
    out = HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;max-width:560px;margin:40px auto;padding:0 16px;line-height:1.6">' +
      '<h2>화면 파일(Index)을 찾지 못했어요</h2>' +
      '<p>Apps Script 편집기 왼쪽 <b>파일</b> 목록에 <b>Index.html</b> 파일이 있어야 해요.</p>' +
      '<ol><li>파일 옆 <b>+ → HTML</b>을 누르고 이름에 <b>Index</b> 만 적으세요 (.html은 자동으로 붙어요).</li>' +
      '<li>그 파일에 Index.html 내용을 전부 붙여 넣고 저장하세요.</li>' +
      '<li><b>배포 → 배포 관리 → 연필 → 버전: 새 버전 → 배포</b>를 누르세요.</li></ol></div>');
  }
  const icon = iconUrl_(); if (icon) { try { out.setFaviconUrl(icon); } catch (e) {} }
  return out.setTitle('신당중앙교회 교육팀 출석부')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * GitHub 바로가기 페이지(화면)가 fetch 로 부르는 창구.
 * 본문: {"op":"boot","auth":{...},"args":[...]}  → api() 와 같은 JSON 을 돌려준다.
 */
/** 화면 버전(CLIENT_V). 이 값보다 낮은(또는 없는) 예전 화면은 서버가 거절한다. 화면을 크게 바꿔 예전 것을 막고 싶을 때만 올린다. */
const MIN_CLIENT = 2;
function doPost(e) {
  let q = {};
  try { q = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  if ((Number(q.cv) || 0) < MIN_CLIENT)
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: '예전 화면이에요. 앱을 완전히 닫고 공식 출석부 주소(' + publicUrl_() + ')로 다시 열어 주세요. (받은 화면 버전: ' + (q.cv == null ? '없음' : q.cv) + ' / 필요한 버전: ' + MIN_CLIENT + ')' }))
      .setMimeType(ContentService.MimeType.JSON);
  const a = Array.isArray(q.args) ? q.args : [];
  return ContentService.createTextOutput(api(q.op, q.auth, a[0], a[1], a[2]))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 공동관리자2에게 막는 요청: 번호(임원·담임)와 이메일 등록 관리. 그 밖에는 공동관리자1과 같다 */
const CO2_BLOCK = { setCode: 1, changeTeamCode: 1, setGradeCode: 1, clearGradeCodes: 1, gradeCodes: 1, tcodeList: 1, resetTCode: 1, setTCode: 1, setSubCode: 1, activity: 1, tgKinds: 1, tgKindsSet: 1, askReply: 1, askReplyLog: 1 };
/** 화면에서 부르는 단일 창구. 결과는 JSON 문자열로 돌려준다. */
function api(op, auth, a, b, c) {
  try {
    if (op === 'ping') { ensureTgPoll_(); ensureBackupTrigger_(); return JSON.stringify({ ok: true, data: Array.isArray(a) ? verMap_(a) : ver_() }); }
    REQ_DEV_ = String((auth && auth.dv) || '').replace(/[^0-9A-Za-z가-힣· ]/g, '').trim().slice(0, 30);
    auth = cleanAuth_(auth);
    if (auth.co2 && CO2_BLOCK[op]) throw new Error('공동관리자2는 임원 번호·담임 번호·이메일 등록을 다룰 수 없어요');
    let data;
    switch (op) {
      case 'boot':       data = boot_(auth, a); break;
      case 'photos':     data = photos_(auth, a); break;
      case 'write':      data = withLock_(() => write_(auth, a, b, c), true); break;
      case 'promoExclude': data = withLock_(() => promoExclude_(auth, a, b), true); break;
      case 'batch':      data = withLock_(() => batch_(auth, a), true); break;
      case 'adminLogin': data = adminLogin_(a, b); break;
      case 'verifyEmail': data = verifyEmail_(a); break;
      case 'coLogin':    data = coLogin_(a, b); break;
      case 'qrMake':     data = qrMake_(auth, a); break;
      case 'qrUse':      data = qrUse_(a); break;
      case 'qrPeek':     data = qrPeek_(a); break;
      case 'qrDevs':     data = qrDevs_(auth); break;
      case 'qrStaff':    data = qrStaff_(auth); break;
      case 'qrCounts':   data = qrCounts_(auth); break;
      case 'devBind':    data = devBind_(auth); break;
      case 'staffPw':    data = withLock_(() => { touchVer_('x'); return staffPw_(auth, a, b); }, true); break;
      case 'checkMyPw':  data = checkMyPw_(auth, a); break;
      case 'login':      data = login_(a, b); break;
      case 'tchStatus':  data = tchStatus_(auth, a); break;
      case 'tchReset':   data = withLock_(() => { touchVer_('x'); return tchReset_(auth, a, b); }, true); break;
      case 'qrRevoke':   data = qrRevoke_(auth, a); break;
      case 'askAdmin':   data = askAdmin_(a); break;
      case 'askReply':   data = askReply_(auth, a); break;
      case 'askReplyLog': data = askReplyLog_(auth, a); break;
      case 'staffAdd':   data = withLock_(() => staffAdd_(auth, a)); break;
      case 'staffEdit':  data = withLock_(() => staffEdit_(auth, a, b)); break;
      case 'staffDel':   data = withLock_(() => staffDel_(auth, a)); break;
      case 'appUrl':     data = publicUrl_(); break;
      case 'contacts':   data = contacts_(auth, a); break;
      case 'saveContact': data = withLock_(() => { touchVer_('x'); return saveContact_(auth, a, b); }, true); break;
      case 'changeMyCode': data = withLock_(() => { touchVer_('x'); return changeMyCode_(auth, a, (b || {}).code, (b || {}).old); }, true); break;
      case 'setTeamSetting': data = withLock_(() => setTeamSetting_(auth, a, b, c)); break;
      case 'setClassCallFirst': data = withLock_(() => setClassCallFirst_(auth, a, b, c)); break;
      case 'setPersonCallFirst': data = withLock_(() => setPersonCallFirst_(auth, a, b)); break;
      case 'setClasses': data = withLock_(() => setClasses_(auth, a, b)); break;
      case 'importRoster': data = withLock_(() => importRoster_(auth, a, b), true); break;
      case 'history':    data = history_(auth, a, b); break;
      case 'activity':   data = activity_(auth, a, b, c); break;
      case 'tgKinds':    data = tgKinds_(auth); break;
      case 'tgKindsSet': data = withLock_(() => tgKindsSet_(auth, a)); break;
      case 'events':     data = listEvents_(auth, a); break;
      case 'saveEvent':  data = withLock_(() => { touchVer_('x'); return saveEvent_(auth, a, b); }); break;
      case 'deleteEvent': data = withLock_(() => { touchVer_('x'); return deleteEvent_(auth, a, b); }); break;
      case 'downloadFile': data = downloadFile_(auth, a, b); break;
      case 'libList':    data = libList_(auth, a); break;
      case 'libUpload':  data = withLock_(() => { touchVer_('x'); return libUpload_(auth, a, b); }); break;
      case 'libDelete':  data = withLock_(() => { touchVer_('x'); return libDelete_(auth, a, b); }); break;
      case 'libDoc':     data = withLock_(() => { touchVer_('x'); return libDocCreate_(auth, a, b); }); break;
      case 'libMail':    data = libMail_(auth, a, b); break;
      case 'libVis':     data = withLock_(() => { touchVer_('x'); return libVis_(auth, a, b, c); }); break;
      case 'rptGet':     data = rptGet_(auth, a); break;
      case 'rptSave':    data = withLock_(() => rptSave_(auth, a, b, c), true); break;
      case 'rptMeta':    data = withLock_(() => rptMeta_(auth, a, b), true); break;
      case 'rptDelete':  data = withLock_(() => rptDelete_(auth, a, b), true); break;
      case 'rptHwp':     data = rptHwp_(auth, a); break;
      case 'rptOrig':    data = rptOrig_(auth, a, b); break;
      case 'rptMail':    data = withLock_(() => rptMail_(auth, a, b), true); break;
      case 'rptFromDel': data = withLock_(() => rptFromDel_(auth, a, b), true); break;
      case 'rptToDel':   data = withLock_(() => rptToDel_(auth, a, b), true); break;
      case 'gameScore':  data = gameScore_(auth, a, b, c); break;
      case 'bibleHit':   data = bibleHit_(auth, a, b); break;
      case 'jubo':       data = jubo_(a); break;
      case 'gameRank':   data = gameRank_(auth, a, b); break;
      case 'gameRankAll': data = gameRankAll_(auth); break;
      default: throw new Error('알 수 없는 요청이에요: ' + op);
    }
    return JSON.stringify({ ok: true, data: data === undefined ? null : data, v: LAST_VER_ || undefined, pv: LAST_VER_ ? PREV_VER_ : undefined });
  } catch (e) {
    tgErr_(op, e);
    return JSON.stringify({ ok: false, error: String((e && e.message) || e) });
  } finally { tgFlushLog_(); }
}

/* ---------------- 성경게임 랭킹 ---------------- */
const GAME_SHEET = '게임랭킹';
const GAME_MAX = { quiz: 9999, race: 9999, type: 9999, map: 9999, cross: 9999, twenty: 9999, shoot: 9999, pair: 9999, diff: 9999, art: 9999, board: 99999 };   // 게임별로 나올 수 있는 최고 점수
/** 랭킹에는 이름만: 괄호, 부서, 직분을 뺀다 (송수현(총관리자), 유소년 김철수 집사 → 송수현, 김철수) */
function gameClean_(name) {
  const role = /(담임목사|부목사|목사|전도사|강도사|장로|권사|안수집사|서리집사|집사|팀장|부장|총무|간사|교사|선생님|선생|관리자)$/;
  const dept = /^(영아|유아|유년|소년|유소년|중고등|청년|교육)(부|팀|교육팀)?$/;
  const raw = String(name || '').replace(/[(（\[].*?[)）\]]/g, ' ').replace(/[(（\[].*$/, ' ');
  const parts = raw.split(/[\s,·/|-]+/).map(x => x.trim()).filter(x => x && !dept.test(x) && !role.test(x) || (x && role.test(x) && x.replace(role, '').length >= 2));
  const first = (parts[0] || '').replace(role, '');
  return first || String(name || '').trim();
}
/** 랭킹에 쓸 이름은 화면이 보낸 값이 아니라 로그인 정보에서 찾는다 */
function gameName_(auth) {
  if (auth.coName) return gameClean_(auth.coName);
  const e = normEmail_(auth.email || auth.staffDev || auth.staffPw || '');
  const s = e ? staffList_().find(x => normEmail_(x.email) === e) : null;
  if (s && s.name) return gameClean_(s.name);
  const m = Object.keys(auth.hrMe || {}).map(t => auth.hrMe[t]).find(x => x && x.name);
  return m ? gameClean_(m.name) : isRoot_(auth) ? '총관리자' : '';
}
function gameTop_(rows, game, lv, name) {
  const list = rows.filter(r => r[1] === game && Number(r[2]) === lv).map(r => ({ n: gameClean_(r[3]), s: Number(r[4]), t: new Date(r[0]).getTime() }))
    .sort((x, y) => y.s - x.s || x.t - y.t);
  list.forEach((x, i) => { x.r = i && list[i - 1].s === x.s ? list[i - 1].r : i + 1; });
  const me = list.find(x => x.n === name);
  return { name, best: me ? me.s : 0, rank: me ? me.r : 0, top: list.slice(0, 10).map(x => ({ r: x.r, n: x.n, s: x.s })) };
}
/** 게임을 하지 않고 랭킹만 보기 */
function gameRank_(auth, game, lv) {
  game = String(game); lv = Number(lv);
  if (!GAME_MAX[game] || [1, 2, 3].indexOf(lv) < 0) throw new Error('랭킹을 찾지 못했어요');
  const sh = ss_().getSheetByName(GAME_SHEET); const n = sh ? sh.getLastRow() - 1 : 0;
  return gameTop_(n > 0 ? sh.getRange(2, 1, n, 5).getValues() : [], game, lv, gameName_(auth));
}
/** 랭킹 보기 화면: 단계별로 사람마다 게임별 최고 점수를 더해 상위 10명과 내 순위를 돌려준다 (없어진 말씀타자는 뺀다) */
function gameRankAll_(auth) {
  const sh = ss_().getSheetByName(GAME_SHEET); const n = sh ? sh.getLastRow() - 1 : 0;
  const rows = n > 0 ? sh.getRange(2, 1, n, 5).getValues() : [], name = gameName_(auth), all = {};
  [1, 2, 3].forEach(l => {
    const sum = {};
    rows.forEach(r => { if (Number(r[2]) === l && GAME_MAX[r[1]] && r[1] !== 'type') { const k = gameClean_(r[3]); sum[k] = (sum[k] || 0) + Number(r[4]); } });
    all[l] = gameTop_(Object.keys(sum).map(k => [0, 'sum', l, k, sum[k]]), 'sum', l, name);
  });
  return { name, all };
}
/** 점수를 사람·게임·단계별 최고 기록으로 남기고 상위 10명과 내 순위를 돌려준다 (출석 데이터 버전은 건드리지 않는다) */
function gameScore_(auth, game, lv, score) {
  game = String(game); lv = Number(lv); score = Math.floor(Number(score));
  if (!GAME_MAX[game] || [1, 2, 3].indexOf(lv) < 0 || !(score >= 0 && score <= GAME_MAX[game])) throw new Error('게임 점수가 올바르지 않아요');
  const name = gameName_(auth);
  if (!name) throw new Error('로그인하면 랭킹에 이름이 올라가요');
  const lock = gameLock_();
  lock.waitLock(20000);
  try {
    const ss = ss_(); let sh = ss.getSheetByName(GAME_SHEET);
    if (!sh) { sh = ss.insertSheet(GAME_SHEET); sh.getRange(1, 1, 1, 5).setValues([['일시', '게임', '단계', '이름', '최고 점수']]).setFontWeight('bold'); sh.setFrozenRows(1); }
    const n = sh.getLastRow() - 1, rows = n > 0 ? sh.getRange(2, 1, n, 5).getValues() : [];
    const at = rows.findIndex(r => r[1] === game && Number(r[2]) === lv && gameClean_(r[3]) === name);
    const old = at >= 0 ? Number(rows[at][4]) : -1, newBest = score > old;
    if (newBest) {
      const rec = [new Date(), game, lv, name, score];
      if (at >= 0) { sh.getRange(at + 2, 1, 1, 5).setValues([rec]); rows[at] = rec; } else { sh.appendRow(rec); rows.push(rec); }
    }
    return Object.assign({ newBest }, gameTop_(rows, game, lv, name));
  } finally { SpreadsheetApp.flush(); lock.releaseLock(); }
}

/* ---------------- 부록-주보보기: 교회 홈페이지 주보 게시판을 읽어 온다 (30분 보관) ---------------- */
function juboGet_(path) {
  const r = UrlFetchApp.fetch('https://www.sdcc.or.kr/chnet2/home/' + path, { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('주보를 불러오지 못했어요');
  return r.getContentText('EUC-KR');
}
/** 최근 주보 8개: 제목이 '11/8 교회소식'처럼 날짜로 시작하거나 '주보'가 들어간 글.
 *  같은 날짜 글이 여럿이면 하나만 (제목에 '주보'가 있는 글을 먼저) */
function juboParseList_(html) {
  const out = [], at = {}, re = /threadKey=(\d+)[^>]*>([^<]*)<[\s\S]*?class="date">([\d.]+)</g; let m;
  while ((m = re.exec(html))) {
    const t = m[2].replace(/파일\s*$/, '').trim(), md = (t.match(/^(\d{1,2}\/\d{1,2})(?!\d)/) || [])[1];
    if (!md && !/주보/.test(t)) continue;
    const x = { k: m[1], t, d: m[3] };
    if (md && md in at) { if (!/주보/.test(out[at[md]].t) && /주보/.test(t)) out[at[md]] = x; continue; }
    if (out.length >= 8) break;
    if (md) at[md] = out.length;
    out.push(x);
  }
  return out;
}
function juboParseImg_(html) {
  const m = html.match(/\/upload_data\/bbs\/bd02\/[\w.-]+\.(?:jpe?g|png|gif)/gi) || [];
  return m.filter((x, i) => m.indexOf(x) === i).map(x => 'https://www.sdcc.or.kr' + x);
}
/** 게시판 한 쪽(20개)의 모든 글: 번호, 제목, 작성일(Date) */
function juboParseRows_(html) {
  const out = [], re = /threadKey=(\d+)[^>]*>([^<]*)<[\s\S]*?class="date">(\d{4})\.(\d{2})\.(\d{2})</g; let m;
  while ((m = re.exec(html))) out.push({ k: m[1], t: m[2].replace(/파일\s*$/, '').trim(), at: new Date(+m[3], m[4] - 1, +m[5]) });
  return out;
}
/** 'YYYY-MM-DD' 날짜의 주보 글 번호. 제목 앞의 '월/일'로 찾고, 작성일이 그 날짜 3주 전~3일 뒤인 글만 인정.
 *  같은 날짜에 '주보' 글과 '교회소식' 글이 있으면 '주보'를 고른다. 찾은 결과는 6시간 보관 */
function juboByDate_(ymd) {
  const c = CacheService.getScriptCache(), ck = 'jubo:d:' + ymd, hit = c.get(ck);
  if (hit) return JSON.parse(hit);
  const [y, mo, d] = ymd.split('-').map(Number), want = new Date(y, mo - 1, d), day = 864e5;
  if (isNaN(want) || want > new Date(Date.now() + 7 * day)) throw new Error('그 날짜의 주보는 아직 없어요');
  const lo = want - 21 * day, hi = +want + 3 * day, md = mo + '/' + d;
  // 한 주에 글이 2개쯤 올라오니 몇 쪽 뒤에 있을지 어림잡고, 앞뒤로 옮겨 가며 찾는다
  let start = Math.max(0, Math.floor((Date.now() - want) / (7 * day) * 1.9 / 20) * 20);
  const seen = {};
  for (let n = 0; n < 10 && !seen[start]; n++) {
    seen[start] = 1;
    const rows = juboParseRows_(juboGet_('?code=bd02&step1=12&sid=107&start=' + start));
    if (!rows.length) { if (start === 0) break; start = Math.max(0, start - 20); continue; }
    const hits = rows.filter(x => x.at >= lo && x.at <= hi && x.t.split(/\s/)[0] === md);
    if (hits.length) {
      const h = hits.find(x => /주보/.test(x.t)) || hits[0], r = { k: h.k, t: h.t };
      c.put(ck, JSON.stringify(r), 21600);
      return r;
    }
    if (rows[rows.length - 1].at > hi) start += 20;                  // 이 쪽은 모두 더 최근 글 → 뒤쪽으로
    else if (rows[0].at < lo && start > 0) start -= 20;              // 모두 더 지난 글 → 앞쪽으로
    else break;                                                      // 이 쪽 기간에 들어 있는데 없다
  }
  throw new Error(md + ' 주보를 찾지 못했어요');
}
/** key가 없으면 가장 최근 주보. key가 'YYYY-MM-DD'면 그 날짜 주보. 돌려주는 값: 최근 목록, 고른 글 번호, 그림 주소 */
function jubo_(key) {
  const c = CacheService.getScriptCache(), hit = n => { const s = c.get(n); return s ? JSON.parse(s) : null; };
  let list = hit('jubo:list2');
  if (!list) { list = juboParseList_(juboGet_('?code=bd02&step1=12&sid=107')); if (list.length) c.put('jubo:list2', JSON.stringify(list), 1800); }
  if (!list.length) throw new Error('주보를 찾지 못했어요');
  let k, found = null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(String(key))) { found = juboByDate_(String(key)); k = found.k; }
  else k = list.some(x => x.k === String(key)) ? String(key) : list[0].k;
  let img = hit('jubo:' + k);
  if (!img) { img = juboParseImg_(juboGet_('?type=sub&step1=12&sid=107&code=bd02&gfile=boardView&threadKey=' + k)); if (img.length) c.put('jubo:' + k, JSON.stringify(img), 1800); }
  return { list, k, img, found };
}

/* ---------------- 성경게임·성경지도 조회수 (총관리자만 본다) ---------------- */
function bibleViews_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('bibleViews') || '{}'); } catch (e) { return {}; } }
/** 메뉴를 열 때마다 1씩 세고, sec가 있으면 이용시간(초)만 더한다. k: 'game' | 'map' | 'jubo' */
function bibleHit_(auth, k, sec) {
  if (k !== 'game' && k !== 'map' && k !== 'jubo') return null;
  const lock = gameLock_(); lock.waitLock(10000);
  try {
    const v = bibleViews_(), s = Math.min(600, Math.max(0, Math.floor(Number(sec) || 0)));
    if (sec === undefined || sec === null) v[k] = (Number(v[k]) || 0) + 1; else v['t' + k] = (Number(v['t' + k]) || 0) + s;
    PropertiesService.getScriptProperties().setProperty('bibleViews', JSON.stringify(v));
    return isRoot_(auth) ? v : null;
  } finally { lock.releaseLock(); }
}

function gameLock_() { return LockService.getDocumentLock() || LockService.getUserLock(); }   // 출석 저장의 잠금과 따로 쓴다
/**
 * 데이터 버전: 쓰기가 끝날 때마다 바뀌는 표시. 화면은 30초마다 이 표시만 물어보고(ping), 달라졌을 때만 전체를 새로 받아요.
 * 버전은 나눠서 둔다 — 'g'(전체), '<부서>'(명단·반 설정), '<부서>_<반>'(그 반 출석·달란트·메모), 'x'(화면 데이터와 무관: 일정·자료실·주소록).
 * 화면은 자기가 보는 부서·반의 버전만 물어봐서, 다른 부서나 다른 반의 출석 체크 때문에 전체를 다시 받지 않는다.
 * (캐시가 비면 '0'이 되어 한 번 새로 받는 것으로 끝나요. 시트를 직접 고친 것은 10분 안에 반영돼요.)
 */
let LAST_VER_ = null, PREV_VER_ = null, TOUCHED_ = {};
function ver_() { try { return CacheService.getScriptCache().get('data-ver') || '0'; } catch (e) { return '0'; } }
function verMap_(keys) {
  keys = (keys || []).map(String).filter(k => k && k.length <= 60).slice(0, 120);
  let hit = {}; try { hit = CacheService.getScriptCache().getAll(keys.map(k => 'v:' + k)); } catch (e) {}
  const out = {}; keys.forEach(k => out[k] = hit['v:' + k] || '0'); return out;
}
/** 이번 쓰기로 바뀌는 버전 표시. 아무것도 표시하지 않은 쓰기는 'g'(모든 화면이 새로 받음) */
function touchVer_(key) { TOUCHED_[key] = 1; }
function touchTeam_(team) { if (team) touchVer_(team); }
function touchGroup_(team, group) { if (team) touchVer_(team + '_' + group); }
function bumpVer_() {
  const keys = Object.keys(TOUCHED_).filter(k => k !== 'x'); if (!keys.length && !TOUCHED_.x) keys.push('g');
  const v = Date.now() + '.' + Math.floor(Math.random() * 1000);
  PREV_VER_ = verMap_(keys); const next = {}; const put = { 'data-ver': v };
  keys.forEach(k => { next[k] = v; put['v:' + k] = v; });
  try { CacheService.getScriptCache().putAll(put, 21600); } catch (e) {}
  TOUCHED_ = {};
  return next;
}
/** hot=true: 출석·명단처럼 자주 하는 쓰기. 일정·자료실·설정 캐시는 건드리지 않는다(설정은 cfgSet_ 이 직접 비움). */
function withLock_(fn, hot) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  TOUCHED_ = {};
  // 쓰기 전후로 일정·자료실 캐시를 비운다 (쓰기 안에서 목록을 다시 읽어도 최신값이 나오게)
  if (!hot) clearListCache_();
  try { return fn(); } finally { if (!hot) clearListCache_(); SpreadsheetApp.flush(); LAST_VER_ = bumpVer_(); lock.releaseLock(); }
}
/** 일정·자료실 목록은 시트에서 매번 읽지 않고 10분간 캐시해 둔다. 쓰기(withLock_)가 있으면 바로 비운다. */
function cachedList_(key, fn) {
  const c = CacheService.getScriptCache();
  const hit = c.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const v = fn();
  try { c.put(key, JSON.stringify(v), 600); } catch (e) {} // 100KB를 넘으면 저장되지 않고 그냥 매번 읽는다
  return v;
}
function clearListCache_() { try { CacheService.getScriptCache().removeAll(['ev-all', 'lib-all', 'cfg-rows']); } catch (e) {} }
const cachedEvents_ = () => cachedList_('ev-all', () => readEvents_().list);
const cachedLib_ = () => cachedList_('lib-all', () => readLib_().list);

/* ---------------- 시트 준비 ---------------- */
/** 숫자 서식(텍스트/날짜) 지정. 표(Table)의 '유형'이 적용된 칸은 서식을 바꿀 수 없어서 오류가 나므로, 실패해도 그냥 넘어간다. */
function fmt_(range, f) { try { range.setNumberFormat(f); } catch (e) {} return range; }
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
/* ---------------- 보안: 번호 무작위 대입 막기 + 시트 수식 주입 막기 ---------------- */
/**
 * 틀린 번호(서로 다른 값)가 쌓이면 그 대상(총관리자, 또는 부서별 입장)을 잠시 잠가요. 잠겨 있는 동안은 맞는 번호도 받지 않아요.
 *   총관리자      : 15번 틀리면 15분부터 시작해 또 걸릴 때마다 2배(최대 2시간)
 *   부서 입장     : (임원 번호·담임 번호) 처음 10번 틀리면 1시간 정지, 그 뒤로는 5번 틀릴 때마다 1시간씩 다시 정지
 * 같은 틀린 번호를 되풀이하는 것(바뀐 예전 번호가 저장된 기기)은 한 번으로 세요.
 * 부서 입장은 맞는 번호로 들어와도 틀린 횟수를 지우지 않아요(캐시 최대 6시간 기억).
 * 잠금 풀기: 편집기에서 clearLoginLocks 실행.
 */
const GUARD_MAX = 15;
const GUARD_WINDOW = 1800;
const GUARD_BASE_SEC = 900;
const GUARD_MAX_SEC = 7200;
const ENTRY_FIRST = 10;        // 부서 입장: 처음 10번 틀리면 정지
const ENTRY_STEP = 5;          // 그 뒤로는 5번 틀릴 때마다 정지
const ENTRY_LOCK_SEC = 3600;   // 한 번에 1시간
const ENTRY_KEEP_SEC = 21600;  // 틀린 횟수를 기억하는 시간(캐시 최대 6시간)
const GUARD_MSG = '번호를 여러 번 틀려서 입력을 막았어요. 시간이 지난 뒤(최대 1시간) 다시 시도해 주세요.';
function isEntry_(kind) { return String(kind).indexOf('entry:') === 0; }
function guardDigest_(v) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'g|' + String(v))).slice(0, 10); }
function guardLocked_(kind) { try { return !!CacheService.getScriptCache().get('lock-' + kind); } catch (e) { return false; } }
function guardOk_(kind) { if (isEntry_(kind)) return; try { CacheService.getScriptCache().remove('fail-' + kind); } catch (e) {} }
/** 담임 번호 비교: 대소문자·앞뒤 공백 구분 없이 */
const gcNorm_ = v => String(v == null ? '' : v).trim().toLowerCase();
const gcEq_ = (a, b) => !!gcNorm_(a) && gcNorm_(a) === gcNorm_(b);
function isGradeCodeOf_(team, code) {
  const g = gradeCodes_();
  return !!gcNorm_(code) && groupsOf_(team).some(([k]) => gcEq_(g[team + '_' + k], code));
}
function logFail_(team, text) { try { withLock_(() => logRows_({}, [['', '', team || '', '보안', text + devTag_()]])); } catch (e) {} }
/** 막힌 동안 다시 시도했는지. 막혀 있으면 '로그인 실패'가 아니라 '막힌 동안 시도'로 (10분에 한 번만) 기록하고 true */
function guardStop_(kind, team, what) {
  if (!guardLocked_(kind)) return false;
  try { const c = CacheService.getScriptCache(), k = 'blk-' + kind; if (c.get(k)) return true; c.put(k, '1', 600); } catch (e) { return true; }
  logFail_(team, (what || kind) + ': 입력이 막힌 동안 다시 시도함 (번호는 확인하지 않고 막음, 틀린 횟수에 안 셈)');
  return true;
}
function guardFail_(kind, attempt, team, what) {
  attempt = String(attempt == null ? '' : attempt); if (!attempt) return;
  if (guardStop_(kind, team, what)) return;
  const c = CacheService.getScriptCache(); const key = 'fail-' + kind; const now = Math.floor(Date.now() / 1000);
  const entry = isEntry_(kind);
  let evt = '';
  const lk = LockService.getUserLock(); let held = false;
  try { lk.waitLock(5000); held = true; } catch (e) {}
  try {
    let st = null; try { st = JSON.parse(c.get(key) || 'null'); } catch (e) {}
    if (!st || (!entry && now - st.t > GUARD_WINDOW)) st = { t: now, a: [], n: 0, k: 0 };
    const d = guardDigest_(attempt);
    if (st.a.indexOf(d) >= 0) return;
    st.a.push(d); if (st.a.length > 40) st.a.shift();
    st.n = (st.n || 0) + 1;
    if (entry) {
      const limit = st.k ? ENTRY_STEP : ENTRY_FIRST;
      if (st.n >= limit) {
        st.k = (st.k || 0) + 1; st.n = 0;
        c.put('lock-' + kind, '1', ENTRY_LOCK_SEC);
        evt = (what || kind) + ': 틀린 번호가 ' + limit + '번 쌓여서 1시간 동안 입력을 막음 (' + st.k + '번째 정지)';
      } else evt = (what || kind) + ': 번호를 틀림 (' + st.n + '/' + limit + ')';
      c.put(key, JSON.stringify(st), ENTRY_KEEP_SEC);
    } else if (st.n >= GUARD_MAX) {
      const strikes = Number(c.get('strike-' + kind) || 0);
      const sec = Math.min(GUARD_MAX_SEC, GUARD_BASE_SEC * Math.pow(2, strikes));
      c.put('lock-' + kind, '1', sec); c.put('strike-' + kind, String(strikes + 1), 21600); c.remove(key);
      evt = (what || kind) + ': 틀린 번호를 ' + GUARD_MAX + '번 넣어서 ' + Math.round(sec / 60) + '분 동안 입력을 막음';
    } else {
      c.put(key, JSON.stringify(st), Math.max(1, GUARD_WINDOW - (now - st.t)));
      evt = (what || kind) + ': 번호를 틀림 (' + st.n + '/' + GUARD_MAX + ')';
    }
  } finally { if (held) lk.releaseLock(); }
  Utilities.sleep(500);
  logFail_(team, evt);
  // 잠금이 걸린 순간은 활동 기록 묶음과 따로, 알림 종류 설정과 상관없이 바로 알린다
  if (evt.indexOf('입력을 막음') >= 0) tgNotify_('🚨 로그인 잠금\n' + evt + devTag_() + '\n\n본인이 아니라면 번호를 바꾸세요. 바로 풀려면 편집기에서 clearLoginLocks 실행');
}

/* ---------------- 임원 이상 이메일 등록 (총관리자가 메인 화면 맨 아래에서 관리) ---------------- */
function staffList_() { const s = cfgAll_().staff; return Array.isArray(s) ? s : []; }
function normEmail_(v) { return String(v || '').trim().toLowerCase(); }
/** 실패 기록에 남길 이메일: 입력한 그대로(소문자), 너무 길면 자르고 비었으면 '없음' */
function mailTag_(v) { const e = normEmail_(v).replace(/\s+/g, ' '); return e ? (e.length > 60 ? e.slice(0, 60) + '…' : e) : '없음'; }
function emailOk_(email) { const e = normEmail_(email); return !!e && staffList_().some(s => normEmail_(s.email) === e); }
/** 임원은 등록할 때 지정한 부서로만 들어갈 수 있다. (총관리자·공동관리자는 모든 부서. 부서가 아직 지정되지 않은 예전 등록은 임시로 전체 허용) */
function staffTeamOk_(email, team) {
  const e = normEmail_(email); const hit = staffList_().find(s => normEmail_(s.email) === e);
  if (!hit) return false;
  if (staffRole_(e) !== 'team') return true;
  return !hit.team || hit.team === team;
}
function rootNameMatch_(name) { const n = String(name || '').replace(/\s+/g, ''); return n === '송수현' || n.startsWith('송수현('); }
function staffRole_(email) { const e = normEmail_(email); const hit = staffList_().find(s => normEmail_(s.email) === e); return hit ? (hit.role === 'admin' ? 'admin' : hit.role === 'co' ? 'co' : hit.role === 'co2' ? 'co2' : 'team') : ''; }
/** 관리자 비밀번호가 정확할 때만 송수현 명단의 임원 역할을 총관리자로 자동 복구합니다. */
function migrateRootStaffRole_(email) { const e = normEmail_(email); const list = staffList_(); let changed = false; list.forEach(s => { if (normEmail_(s.email) === e && rootNameMatch_(s.name) && s.role !== 'admin') { s.role = 'admin'; changed = true; } }); if (changed) cfgSet_('staff', list); }
function isRoot_(auth) { return isAdmin_(auth) && !auth.co; }   // 공동관리자가 아닌 총관리자 1인
function coPwMap_() { const m = cfgAll_().coPw; return (m && typeof m === 'object') ? m : {}; }
function coHash_(salt, pw) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + String(pw))); }
function coCheck_(email, pw) { const r = coPwMap_()[normEmail_(email)]; if (!r || !r.s) return false; const p = String(pw == null ? '' : pw); return r.h === coHash_(r.s, p.toLowerCase()) || r.h === coHash_(r.s, p); }
function coSet_(map, email, pw) { const s = Utilities.getUuid(); map[email] = { s, h: coHash_(s, String(pw).toLowerCase()) }; }
/** 선생님 비밀번호 검증 (새 해시 저장본과 기존 평문 저장본 모두 호환) */
function tcCheck_(r, pw) {
  if (!r) return false;
  if (r.s && r.h) return r.h === coHash_(r.s, gcNorm_(pw));
  return gcEq_(r.c, pw);
}
/** 개인 비밀번호로 로그인하는 역할: 공동관리자1 co, 공동관리자2 co2 */
const PW_ROLES = ['co', 'co2'];
const STAFF_ROLES = ['admin', 'co', 'co2'];
function roleName_(role, team) { return role === 'admin' ? '총관리자' : role === 'co' ? '공동관리자1' : role === 'co2' ? '공동관리자2' : '임원' + (team && TEAMS[team] ? '(' + TEAMS[team].label + ')' : ''); }
/** 모든 개인 비밀번호 규칙: 영문+숫자 포함 6자리 이상(30자까지), 공백 없이. 대소문자는 구분하지 않는다 */
const PW_RULE = /^(?=.*[A-Za-z])(?=.*[0-9])\S{6,30}$/;
const PW_RULE_MSG = '비밀번호는 영문과 숫자를 모두 넣은 6자리 이상이에요 (공백 없이, 대소문자 구분 없음)';
function validCoPw_(pw) { return PW_RULE.test(String(pw || '')); }
/** 직분: 목사·전도사·팀장은 자기 부서(등록된 담당 부서)의 임원을 관리(등록·수정·삭제·QR·기기 해제)할 수 있다 */
const STAFF_POS = ['목사', '전도사', '팀장'];
/** 관리자 등록을 다룰 수 있는 범위: 'full'(총관리자·공동관리자1: 모든 등록) / [부서](목사·전도사·팀장: 자기 부서) / null */
function staffMgr_(auth) {
  if (isAdmin_(auth) && !auth.co2) return 'full';
  const em = normEmail_(auth.staffDev || auth.staffPw || ''); const me = em ? staffList_().find(s => normEmail_(s.email) === em) : null;
  if (!me) return null;
  if (STAFF_POS.includes(me.pos) && TEAMS[me.team]) return [me.team];
  // 직분을 따로 안 정했어도, 부서 교사 명단의 직책이 목사·전도사·팀장이면 자기 부서 임원 관리 권한을 준다
  if (TEAMS[me.team] && /목사|전도사|팀장/.test(me.duty || staffDuty_(me, activeTeachers_()))) return [me.team];
  return null;
}
/** 직분으로 관리하는 사람이 이 등록(s)을 다룰 수 있는지: 자기 부서의 직분 없는 임원만 (직분이 있는 사람은 총관리자·공동관리자1만) */
function staffInScope_(sc, s) {
  if (sc === 'full') return true;
  if (!sc || !s || (STAFF_ROLES.includes(s.role) ? s.role : 'team') !== 'team') return false;
  if (s.pos && !(sc === 'all' && s.pos === '팀장')) return false;
  return sc === 'all' || sc.indexOf(s.team) >= 0;
}
/** 화면에 보여 줄 등록 명단 (직분 관리자는 범위 안 임원만) */
/** 부서 명단(교사 직분)에서 가져온 부서 직책: 팀장·서기·회계·총무 */
const DUTY_RE = /목사|전도사|팀장|서기|회계|총무/, STAFF_DUTY = ['목사', '전도사', '팀장', '서기', '총무', '회계'];
function staffDuty_(s, tch) {
  const ids = Object.keys(tch).filter(pid => sameName_(s.name, tch[pid].n) && (!s.team || tch[pid].t === s.team));
  const r = ids.map(pid => tch[pid].r).find(x => DUTY_RE.test(x || ''));
  return r ? String(r).slice(0, 20) : '';
}
function staffView_(auth) {
  const sc = staffMgr_(auth); if (!sc) return [];
  const pwm = coPwMap_(), devs = {}, props = qrProps_().getProperties();   // 개인 비밀번호 유무와 QR 기기 수
  Object.keys(props).forEach(k => { if (k.indexOf('qd:') !== 0) return; try { (JSON.parse(props[k]).ids || []).forEach(x => { if (x.e) devs[normEmail_(x.e)] = (devs[normEmail_(x.e)] || 0) + 1; }); } catch (e) {} });
  const tch = activeTeachers_();
  const all = staffList_().map(s => ({ name: s.name, email: s.email, role: STAFF_ROLES.includes(s.role) ? s.role : 'team', team: s.team || '', pos: s.pos || '', duty: s.duty || staffDuty_(s, tch), hasPw: !!pwm[normEmail_(s.email)], devs: devs[normEmail_(s.email)] || 0 }));
  if (sc === 'full') return all;
  if (!sc) return [];
  return all.filter(s => s.role === 'team' && (sc === 'all' || sc.indexOf(s.team) >= 0));
}
function staffFields_(f) {
  f = f || {};
  const name = String(f.name || '').trim().slice(0, 20), email = normEmail_(f.email), role = STAFF_ROLES.includes(f.role) ? f.role : 'team';
  if (!name) throw new Error('이름을 적어 주세요');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('올바른 이메일 주소가 아니에요');
  const team = role === 'team' ? String(f.team || '') : '';
  if (role === 'team' && !TEAMS[team]) throw new Error('임원은 담당 부서를 지정해 주세요');
  const pos = STAFF_POS.includes(f.pos) ? f.pos : '';
  const duty = role === 'team' && STAFF_DUTY.includes(f.duty) ? f.duty : '';   // 부서 직책 (표시만, 권한과 상관없어요)
  return { name, email, role, team, pos, duty, pw: String(f.pw || '').trim() };
}
function staffAdd_(auth, f) {
  const sc = staffMgr_(auth); if (!sc) throw new Error('관리자만 등록할 수 있어요');
  f = staffFields_(f);
  if (sc !== 'full') {   // 목사·전도사·팀장: 범위 안 부서의 임원만 (직분·비밀번호는 정할 수 없어요)
    if (f.role !== 'team') throw new Error('임원만 등록할 수 있어요');
    if (sc !== 'all' && sc.indexOf(f.team) < 0) throw new Error('자기 부서 임원만 등록할 수 있어요');
    f.pos = sc === 'all' && f.pos === '팀장' ? '팀장' : ''; f.pw = '';
  }
  if (auth.co && f.role === 'admin') throw new Error('공동관리자는 총관리자 계정을 등록할 수 없어요');
  const list = staffList_().map(s => ({ name: s.name, email: s.email, role: STAFF_ROLES.includes(s.role) ? s.role : 'team', team: s.team || '', pos: s.pos || '', duty: s.duty || '' }));
  if (list.some(s => normEmail_(s.email) === f.email)) throw new Error('이미 등록된 이메일이에요');
  const map = Object.assign({}, coPwMap_());
  if (PW_ROLES.includes(f.role) && f.pw) {   // 비워 두면 QR로 처음 로그인한 뒤 본인이 만들어요
    if (!validCoPw_(f.pw)) throw new Error(PW_RULE_MSG);
    coSet_(map, f.email, f.pw);
  }
  list.push({ name: f.name, email: f.email, role: f.role, team: f.team, pos: f.pos, duty: f.duty });
  cfgSet_('staff', list);
  if (PW_ROLES.includes(f.role) && f.pw) cfgSet_('coPw', map);
  logRows_(auth, [['', f.name, f.team || '', '이메일', roleName_(f.role, f.team) + (f.pos ? ' · ' + f.pos : '') + ' 이메일 등록: ' + f.email]]);
  return staffView_(auth);
}
/** 임원 로그인 1단계: 등록된 이메일인지 확인하고 이름·구분을 돌려준다. (이메일 추측 대입은 잠금 대상) */
function verifyEmail_(email) {
  if (guardStop_('mail', '', '이메일 로그인')) throw new Error(GUARD_MSG);
  const e = normEmail_(email);
  const hit = staffList_().find(s => normEmail_(s.email) === e);
  if (!hit) { guardFail_('mail', e || '-', '', '이메일 로그인 실패'); throw new Error('등록되지 않은 이메일이에요. 총관리자에게 등록을 요청하세요.'); }
  const role = STAFF_ROLES.includes(hit.role) ? hit.role : 'team';
  return { name: hit.name, role, team: role === 'team' ? (hit.team || '') : '' };
}
/** 관리자 로그인 2단계: 공동관리자·임원의 개인 비밀번호 확인 (총관리자는 관리자 비밀번호로 따로) */
function coLogin_(email, pw) {
  const e = normEmail_(email); const kind = 'co:' + e;
  if (guardStop_(kind, '', '관리자 로그인 (' + mailTag_(e) + ')')) throw new Error(GUARD_MSG);
  const hit = staffList_().find(s => normEmail_(s.email) === e && s.role !== 'admin');
  if (!hit) throw new Error('관리자로 등록된 이메일이 아니에요');
  const nm = roleName_(hit.role, hit.team);
  if (!coPwMap_()[e]) throw new Error('아직 개인 비밀번호가 없어요. 관리자에게 QR을 받아 처음 로그인한 뒤 비밀번호를 만들어 주세요.');
  if (!coCheck_(e, pw)) { guardFail_(kind, String(pw || '-'), hit.team || '', nm + ' 로그인 실패 (' + hit.name + ')'); throw new Error('비밀번호가 맞지 않아요'); }
  guardOk_(kind);
  logLogin_(PW_ROLES.includes(hit.role) ? { admin: adminPw_(), co: true, co2: hit.role === 'co2', coName: hit.name } : { mgr: { [hit.team]: 1 }, email: e }, hit.team || '', nm + ' ' + hit.name + ' 로그인');
  return { ok: true, name: hit.name, role: hit.role, team: hit.team || '' };
}
/** 하나로 합친 로그인 창: 이메일이면 공동관리자·임원(개인 비밀번호), 이름이면 선생님(개인 비밀번호).
 *  총관리자는 여기서 받지 않는다 (숨은 관리자 로그인으로). 성공하면 화면이 번호를 기억하고 다음 요청에서 기기로 등록한다 */
function login_(id, pw) {
  id = String(id == null ? '' : id).trim().slice(0, 100); pw = String(pw == null ? '' : pw).trim();
  if (!id || !pw) throw new Error('이름(임원은 이메일)과 비밀번호를 넣어 주세요');
  if (id.indexOf('@') >= 0) {
    const e = normEmail_(id), hit = staffList_().find(s => normEmail_(s.email) === e);
    if (!hit) { if (guardStop_('mail', '', '로그인 (등록되지 않은 이메일: ' + mailTag_(e) + ')')) throw new Error(GUARD_MSG); guardFail_('mail', e, '', '로그인 실패 (등록되지 않은 이메일: ' + mailTag_(e) + ')'); throw new Error('등록되지 않은 이메일이에요. 이메일을 다시 확인해 주세요.'); }
    if (hit.role === 'admin') throw new Error('총관리자 계정이에요. 관리자 로그인 메뉴를 이용해 주세요.');
    const r = coLogin_(e, pw);
    return { kind: 'staff', email: e, role: r.role, team: r.team || '', name: r.name };
  }
  const tc = cfgAll_().tcodes || {}, tch = activeTeachers_();
  const mine = Object.keys(tc).filter(pid => tch[pid] && TEAMS[tc[pid].t] && tch[pid].t === tc[pid].t && nameMatch_(tch[pid].n, id));
  if (!mine.length) {
    const known = Object.keys(tch).some(pid => nameMatch_(tch[pid].n, id));
    if (known) throw new Error('아직 개인 비밀번호가 없어요. 임원에게 QR을 받아 처음 로그인한 뒤 비밀번호를 만들어 주세요.');
    if (guardStop_('login', '', '로그인 (이름: ' + id.slice(0, 20) + ')')) throw new Error(GUARD_MSG);
    guardFail_('login', id + '|' + pw, '', '로그인 실패 (이름: ' + id.slice(0, 20) + ')');
    throw new Error('이름이나 비밀번호가 맞지 않아요. 교사 명단에 적힌 이름 그대로 넣어 주세요.');
  }
  const teams = mine.map(pid => tc[pid].t).filter((t, i, a) => a.indexOf(t) === i);
  if (teams.every(t => guardLocked_('entry:' + t))) { teams.forEach(t => guardStop_('entry:' + t, t, TEAMS[t].label + ' ' + id.slice(0, 20) + ' 선생님 로그인')); throw new Error(GUARD_MSG); }
  const ok = mine.filter(pid => !guardLocked_('entry:' + tc[pid].t) && tcCheck_(tc[pid], pw));
  if (!ok.length) { teams.filter(t => !guardLocked_('entry:' + t)).forEach(t => guardFail_('entry:' + t, pw, t, TEAMS[t].label + ' ' + id.slice(0, 20) + ' 선생님 로그인 실패')); throw new Error('이름이나 비밀번호가 맞지 않아요'); }
  ok.forEach(pid => { const t = tc[pid].t, nm = String(tch[pid].n || ''); logLogin_({ hrMe: { [t]: { pid, name: nm, own: true } } }, t, TEAMS[t].label + ' ' + nm + ' 선생님 로그인'); });
  return { kind: 'teacher', list: ok.map(pid => ({ pid, team: tc[pid].t, name: String(tch[pid].n || '') })) };
}
/** 공동관리자·임원 본인 비밀번호 만들기·바꾸기. QR로 등록된 기기에서는 지금 비밀번호 없이 정할 수 있다 */
function staffPw_(auth, pw, old) {
  const em = normEmail_(auth.staffDev || auth.staffPw || '');
  const s = em ? staffList_().find(z => normEmail_(z.email) === em) : null;
  if (!s || s.role === 'admin') throw new Error('QR이나 개인 비밀번호로 로그인한 관리자만 바꿀 수 있어요');
  const map = Object.assign({}, coPwMap_());
  if (!auth.staffDev && map[em] && !coCheck_(em, old)) { guardFail_('co:' + em, String(old || '-'), s.team || '', roleName_(s.role, s.team) + ' 비밀번호 변경 (지금 비밀번호 틀림)'); throw new Error('지금 쓰는 비밀번호가 맞지 않아요'); }
  if (!validCoPw_(pw)) throw new Error(PW_RULE_MSG);
  const first = !map[em]; coSet_(map, em, pw); cfgSet_('coPw', map);
  logRows_(auth, [['', s.name, s.team || '', '번호', roleName_(s.role, s.team) + ' ' + s.name + ' 개인 비밀번호를 ' + (first ? '만듦' : '바꿈')]]);
  return true;
}
/** 공동관리자·임원 본인 비밀번호 확인 (연말 학년 등반처럼 큰 작업 전에). 틀리면 잠금 횟수에 센다 */
function checkMyPw_(auth, pw) {
  const em = normEmail_(auth.staffDev || auth.staffPw || '');
  const s = em ? staffList_().find(z => normEmail_(z.email) === em) : null;
  if (!s || s.role === 'admin') throw new Error('QR이나 개인 비밀번호로 로그인한 관리자만 확인할 수 있어요');
  if (!coPwMap_()[em]) throw new Error('아직 개인 비밀번호가 없어요. 메인 화면 [내 비밀번호]에서 먼저 만들어 주세요');
  const kind = 'co:' + em; if (guardStop_(kind, s.team || '', roleName_(s.role, s.team) + ' 비밀번호 확인 (' + s.name + ')')) throw new Error(GUARD_MSG);
  if (coCheck_(em, pw)) return true;
  guardFail_(kind, String(pw || '-'), s.team || '', roleName_(s.role, s.team) + ' 비밀번호 확인 실패 (' + s.name + ')');
  return false;
}
/** [로그인 관리 → 선생님 관리]: 부서 선생님마다 개인 비밀번호 유무와 QR로 등록된 기기 수 (임원·공동관리자·총관리자) */
function tchStatus_(auth, team) {
  if (!TEAMS[team]) throw new Error('부서를 알 수 없어요');
  if (!deptLevel_(auth, team)) throw new Error('임원만 볼 수 있어요');
  try { CacheService.getScriptCache().remove('qr-tch2'); } catch (e) {}
  const tch = activeTeachers_(), tc = cfgAll_().tcodes || {}, devs = {};
  const all = qrProps_().getProperties();
  Object.keys(all).forEach(k => { if (k.indexOf('qd:') !== 0) return; try { (JSON.parse(all[k]).ids || []).forEach(x => { if (x.k === 'hr') devs[x.pid] = (devs[x.pid] || 0) + 1; }); } catch (e) {} });
  // 임원(관리자 등록 명단에 이름이 있는 사람)은 [관리자 QR]에서만 다뤄요
  const staffNames = staffList_().map(s => String(s.name || '')).filter(Boolean);
  return Object.keys(tch).filter(pid => tch[pid].t === team && !staffNames.some(n => sameName_(n, tch[pid].n))).map(pid => ({ pid, name: String(tch[pid].n || ''), hasPw: !!(tc[pid] && tc[pid].t === team), devs: devs[pid] || 0 }))
    .sort((a, b) => a.name.localeCompare(b.name, 'ko'));
}
/** 선생님 개인 비밀번호 초기화 (공동관리자2도 가능: 로그인 관리 권한). 초기화하면 새 QR로 들어와 다시 만든다 */
function tchReset_(auth, team, pid) {
  if (!deptLevel_(auth, team)) throw new Error('임원만 초기화할 수 있어요');
  const tc = Object.assign({}, cfgAll_().tcodes || {}); const r = tc[pid];
  if (r && r.t === team) { delete tc[pid]; cfgSet_('tcodes', Object.keys(tc).length ? tc : null); logRows_(auth, [[pid, r.n, team, '번호', r.n + ' 선생님 개인 비밀번호를 초기화함']]); }
  return tchStatus_(auth, team);
}
/** 예전 공용 번호를 완전히 지운다 (2026-11 QR·개인 비밀번호로 바꾼 뒤). 남아 있으면 첫 boot 때 한 번 지우고 기록한다 */
function dropTeamCodes_() {
  const c = cfgAll_(), keys = ['codes', 'codeSelf', 'gradeCodes', 'gcOwn', 'gcChanged', 'subCodes', 'legacyLogin'];
  const has = keys.filter(k => c[k] !== undefined && c[k] !== null);
  if (!has.length) return;
  try { withLock_(() => { has.forEach(k => cfgSet_(k, null));
    logRows_({ sys: '시스템' }, [['', '', '', '설정', '예전 번호(임원 번호·반 담임 번호·보조교사 비밀번호)를 모두 지움']]); }); } catch (e) {}
}
function clearTeamCodes() { dropTeamCodes_(); Logger.log('팀별 관리자 번호를 지웠어요.'); }
/** 로그인 화면의 [관리자에게 요청하기]: 이메일 등록·비밀번호 분실 등 요청을 총관리자에게 메일로 보냅니다 (MailApp, 무료).
 *  받는 사람: 등록 명단의 총관리자 이메일 + 스크립트 소유자. 같은 사람은 10분에 1번, 전체는 하루 30통까지. */
const ASK_KINDS = { email: '이메일 등록 요청', pw: '비밀번호·번호 분실/오류', etc: '기타 문의' };
function askAdmin_(f) {
  f = f || {};
  const cut = (v, n) => String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, n);
  const name = cut(f.name, 30), pos = cut(f.pos, 30), phone = cut(f.phone, 30), email = normEmail_(cut(f.email, 100));
  const team = TEAMS[f.team] ? f.team : '', co = f.team === 'co', kind = ASK_KINDS[f.kind] ? f.kind : 'etc';
  const msg = String(f.msg == null ? '' : f.msg).trim().slice(0, 1000);
  if (!name) throw new Error('이름을 적어 주세요');
  if (!msg) throw new Error('요청사항을 적어 주세요');
  if (!email && !phone) throw new Error('답을 받을 이메일이나 연락처를 하나는 적어 주세요');
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('이메일 주소를 다시 확인해 주세요');
  const c = CacheService.getScriptCache(), who = 'ask-' + guardDigest_(name + '|' + (email || phone));
  if (c.get(who)) throw new Error('이미 요청을 보냈어요. 10분 뒤에 다시 보낼 수 있어요');
  const day = 'ask-day-' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd'), n = Number(c.get(day) || 0);
  if (n >= 30) throw new Error('오늘은 요청이 많아 더 보낼 수 없어요. 내일 다시 보내거나 관리자에게 직접 연락해 주세요');
  const to = {};
  staffList_().forEach(s => { if (s.role === 'admin' && emailOk_(s.email)) to[normEmail_(s.email)] = 1; });
  try { const own = normEmail_(Session.getEffectiveUser().getEmail()); if (own) to[own] = 1; } catch (e) {}
  const list = Object.keys(to);
  if (!list.length) throw new Error('받을 관리자 이메일이 없어요. 관리자에게 직접 연락해 주세요');
  if (MailApp.getRemainingDailyQuota() < list.length) throw new Error('오늘 보낼 수 있는 메일 수를 다 썼어요. 관리자에게 직접 연락해 주세요');
  const when = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm');
  const rows = [['요청 종류', ASK_KINDS[kind]], ['이름', name], ['부서', team ? TEAMS[team].label : co ? '공동관리자' : '-'], ['직책', pos || '-'],
    ['이메일', email || '-'], ['연락처', phone || '-'], ['보낸 시각', when]].concat(REQ_DEV_ ? [['기기', REQ_DEV_]] : [], [['요청사항', msg]]);
  const h = s => String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
  const mail = { to: list.join(','), name: '신당중앙교회 교육팀 출석부', subject: '[출석부 요청] ' + ASK_KINDS[kind] + ' - ' + name + (pos ? ' ' + pos : ''),
    body: rows.map(r => r[0] + ': ' + r[1]).join('\n') + '\n\n출석부 앱 > 관리 > 이메일 등록에서 처리할 수 있어요.\n' + publicUrl_(),
    htmlBody: '<table style="border-collapse:collapse;font-size:14px">' + rows.map(r => '<tr><th style="text-align:left;padding:6px 12px;background:#f3f4f6;border:1px solid #ddd;white-space:nowrap">' + r[0] + '</th><td style="padding:6px 12px;border:1px solid #ddd;white-space:pre-wrap">' + h(r[1]) + '</td></tr>').join('') + '</table><p style="font-size:13px;color:#666">출석부 앱 &gt; 관리 &gt; 이메일 등록에서 처리할 수 있어요.<br><a href="' + h(publicUrl_()) + '">출석부 열기</a></p>' };
  if (email) mail.replyTo = email;
  MailApp.sendEmail(mail);
  c.put(who, '1', 600); c.put(day, String(n + 1), 21600);
  if (tgOn_('요청')) tgNotify_('📩 ' + mail.subject + '\n\n' + rows.map(r => r[0] + ': ' + r[1]).join('\n'));
  // 〔회신 …〕: 활동 기록의 [답장] 버튼이 이 연락처로 메일·문자를 보낸다. 〔기기 …〕는 답장에 인용하지 않는다
  const back = [email, phone].filter(String).join(' · ');
  try { withLock_(() => logRows_({}, [['', name, team, '요청', ASK_KINDS[kind] + (pos ? ' (' + pos + ')' : '') + (REQ_DEV_ ? ' 〔기기 ' + REQ_DEV_ + '〕' : '') + ' 〔회신 ' + back + '〕: ' + msg.slice(0, 200)]])); } catch (e) {}
  return { ok: true };
}
/** 텔레그램 봇으로 휴대폰 알림을 보낸다. 연결 안 됐거나 실패해도 조용히 넘어간다 (메일은 이미 갔으므로).
 *  btn: { text, url } 을 주면 알림 아래에 링크 버튼을 붙인다 */
function tgNotify_(text, btn) {
  const p = PropertiesService.getScriptProperties(), token = p.getProperty('TG_TOKEN'), chat = p.getProperty('TG_CHAT');
  if (!token || !chat) return;
  try {
    UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', { method: 'post', muteHttpExceptions: true,
      payload: Object.assign({ chat_id: chat, text: String(text).slice(0, 4000), disable_web_page_preview: 'true' },
        btn ? { reply_markup: JSON.stringify({ inline_keyboard: [[{ text: btn.text, url: btn.url }]] }) } : {}) });
  } catch (e) {}
}
/** 예상하지 못한 서버 오류를 텔레그램으로 알린다. 안내용으로 던진 오류(new Error)는 빼고, 같은 오류는 30분에 한 번만 */
function tgErr_(op, e) {
  if (!e || e.name === 'Error') return;
  try {
    const msg = String(e.message || e).slice(0, 300), c = CacheService.getScriptCache();
    const key = 'tgerr-' + Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, op + '|' + msg)).slice(0, 16);
    if (c.get(key)) return; c.put(key, '1', 1800);
    tgNotify_('⚠️ 서버 오류 (' + op + ')' + devTag_() + '\n' + (e.name || '') + ': ' + msg + (e.stack ? '\n' + String(e.stack).split('\n').slice(1, 3).join('\n') : ''));
  } catch (x) {}
}
/** 텔레그램 알림 연결: 편집기에서 한 번 실행하세요.
 *  1) 텔레그램 @BotFather → /newbot → 받은 토큰을 [프로젝트 설정 → 스크립트 속성]에 TG_TOKEN 으로 저장
 *  2) 만든 봇과 대화방을 열어 아무 말이나 보낸 뒤 → 이 함수 실행 → 테스트 알림이 오면 끝 */
function connectTelegram() {
  const p = PropertiesService.getScriptProperties(), token = p.getProperty('TG_TOKEN');
  if (!token) throw new Error('스크립트 속성에 TG_TOKEN(봇 토큰)을 먼저 저장해 주세요');
  const res = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getUpdates', { muteHttpExceptions: true }).getContentText());
  if (!res.ok) throw new Error('토큰이 올바르지 않아요: ' + (res.description || ''));
  const last = (res.result || []).map(u => u.message || u.edited_message).filter(m => m && m.chat).pop();
  if (!last) throw new Error('봇에게 온 메시지가 없어요. 텔레그램에서 봇에게 아무 말이나 보낸 뒤 다시 실행해 주세요');
  p.setProperty('TG_CHAT', String(last.chat.id));
  tgNotify_('✅ 출석부 문의 알림이 연결됐어요. 이제 [관리자에게 요청하기]가 오면 여기로 알려 드려요.');
  Logger.log('연결 완료: ' + (last.chat.first_name || last.chat.title || '') + ' (' + last.chat.id + '). 텔레그램에 테스트 알림이 왔는지 확인하세요.');
}
/** 활동 기록의 요청에 총관리자가 메일로 답장한다. f: { email, name, msg, ask(원래 요청 일부) } */
function askReply_(auth, f) {
  if (!isRoot_(auth)) throw new Error('답장은 총관리자만 보낼 수 있어요');
  return askReplySend_(auth, f, '이메일 답장');
}
/** 답장 메일 보내기 (권한 확인은 부르는 쪽에서). how: 활동 기록에 남길 말 */
function askReplySend_(auth, f, how) {
  f = f || {};
  const email = normEmail_(String(f.email || '').trim()), name = String(f.name || '').replace(/[\r\n\t]+/g, ' ').trim().slice(0, 30);
  const msg = String(f.msg == null ? '' : f.msg).trim().slice(0, 2000), ask = String(f.ask || '').trim().slice(0, 300);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('받는 사람 이메일이 올바르지 않아요');
  if (!msg) throw new Error('답장 내용을 적어 주세요');
  if (MailApp.getRemainingDailyQuota() < 1) throw new Error('오늘 보낼 수 있는 메일 수를 다 썼어요');
  const h = s => String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
  const quote = ask ? '\n\n----- 보내신 요청 -----\n' + ask : '';
  MailApp.sendEmail({ to: email, name: '신당중앙교회 교육팀 출석부', subject: '[출석부] ' + (name ? name + '님 ' : '') + '요청에 대한 답장',
    body: msg + quote + '\n\n출석부 열기: ' + publicUrl_(),
    htmlBody: '<div style="font-size:14px;white-space:pre-wrap">' + h(msg) + '</div>' + (ask ? '<p style="font-size:13px;color:#666;border-left:3px solid #ddd;padding-left:10px;white-space:pre-wrap">보내신 요청: ' + h(ask) + '</p>' : '') + '<p style="font-size:13px"><a href="' + h(publicUrl_()) + '">출석부 열기</a></p>' });
  try { withLock_(() => logRows_(auth, [['', name, '', '답장', how + ' (' + email + '): ' + msg.slice(0, 200)]])); } catch (e) {}
  return { ok: true };
}
/** 텔레그램 답장 확인 트리거(TG_POLL_MIN분마다)가 없거나 간격이 다르면 새로 만든다. ping 때 부르며, 6시간에 한 번만 검사한다. */
const TG_POLL_MIN = 5;   // 구글이 허용하는 값: 1, 5, 10, 15, 30
function ensureTgPoll_() {
  const c = CacheService.getScriptCache(), key = 'tgPollChk' + TG_POLL_MIN; if (c.get(key)) return;
  c.put(key, '1', 21600);
  try {
    const p = PropertiesService.getScriptProperties();
    if (!p.getProperty('TG_CHAT')) return;
    const old = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'tgPoll');
    if (old.length && p.getProperty('TG_POLL_MIN') === String(TG_POLL_MIN)) return;
    old.forEach(t => ScriptApp.deleteTrigger(t));
    ScriptApp.newTrigger('tgPoll').timeBased().everyMinutes(TG_POLL_MIN).create();
    p.setProperty('TG_POLL_MIN', String(TG_POLL_MIN));
  } catch (e) {}
}
/** TG_POLL_MIN분마다: 문의 알림에 [답장]으로 보낸 텔레그램 메시지를 문의한 사람에게 메일로 보낸다. 연결된 대화방 것만 받는다. */
function tgPoll() {
  try { ensureBackupTrigger_(); } catch (e) {}
  try { trimBackupsIfKeepChanged_(); } catch (e) {}
  try { flushAttLog_(); } catch (e) {}   // 18시: 그날 출석 체크 기록
  const p = PropertiesService.getScriptProperties(), token = p.getProperty('TG_TOKEN'), chat = p.getProperty('TG_CHAT');
  if (!token || !chat) return;
  try { tgSundayTick_(p); } catch (e) { tgNotify_('⚠️ 주일 알림을 만들지 못했어요: ' + e.message); }
  try { tgSetCommands_(p, token); } catch (e) {}
  const res = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getUpdates?offset=' + Number(p.getProperty('TG_OFFSET') || 0), { muteHttpExceptions: true }).getContentText());
  if (!res.ok || !res.result || !res.result.length) return;
  p.setProperty('TG_OFFSET', String(res.result[res.result.length - 1].update_id + 1)); // 먼저 저장해 같은 답장을 두 번 보내지 않게
  res.result.forEach(u => {
    const m = u.message;
    if (!m || !m.chat || String(m.chat.id) !== chat || !m.text) return;
    try { tgReply_(m); } catch (e) { tgNotify_('⚠️ 답장을 보내지 못했어요: ' + e.message); }
  });
}
/* ---------------- 주일 텔레그램 알림 (tgPoll이 5분마다 확인해서 주일에 한 번씩) ----------------
 *  TG_BD_HOUR시 : 이번 주(주일~토) 생일
 *  TG_SUM_HOUR시: 부서별 출석 요약 + 출석 체크를 안 한 반
 *  보낸 날짜는 스크립트 속성 TG_SUN_BD / TG_SUN_SUM 에 적어 두어 두 번 보내지 않는다. 편집기에서 tgSundayTest 로 미리 볼 수 있다. */
const TG_BD_HOUR = 8;
const TG_SUM_HOUR = 18;
const DOW_KO = ['일', '월', '화', '수', '목', '금', '토'];
function tgSundayTick_(p) {
  const now = new Date(), tz = tz_();
  if (Utilities.formatDate(now, tz, 'u') !== '7') return;   // 7 = 일요일
  const key = Utilities.formatDate(now, tz, 'yyyy-MM-dd'), h = Number(Utilities.formatDate(now, tz, 'H'));
  if (h >= TG_BD_HOUR && p.getProperty('TG_SUN_BD') !== key) { p.setProperty('TG_SUN_BD', key);
    const rows = []; tgNotify_(tgBirthdays_(key, rows, true));
    if (rows.length) withLock_(() => logRows_({ sys: '자동 알림' }, rows)); }   // 활동 기록에도 한 사람씩
  if (h >= TG_SUM_HOUR && p.getProperty('TG_SUN_SUM') !== key) { p.setProperty('TG_SUN_SUM', key); tgNotify_(tgSunSummary_(key)); }
}
/** 편집기에서 실행: 가장 최근 주일 기준으로 두 알림을 지금 바로 보내 본다 (보낸 날짜 기록은 건드리지 않음) */
function tgSundayTest() { const key = lastSunKey_(); tgNotify_(tgBirthdays_(key, null, true)); tgNotify_(tgSunSummary_(key)); }
/** 오늘이 주일이면 오늘, 아니면 지난 주일 'yyyy-MM-dd' */
function lastSunKey_() {
  const tz = tz_(), now = new Date();
  return keyStr_(keyAdd_(Utilities.formatDate(now, tz, 'yyyy-MM-dd'), -(Number(Utilities.formatDate(now, tz, 'u')) % 7)));
}

/* ---------------- 텔레그램 명령 (봇에게 보내면 다음 tgPoll 때 답한다) ---------------- */
const TG_CMDS = [['att', '출석', '주일 출석 요약 (예: 출석 유소년 10/18)'], ['absent', '결석', '결석자 명단, 연속 결석 표시 (예: 결석 중고등)'],
  ['find', '찾기', '학생·교사 찾기 (예: 찾기 김하나)'], ['bday', '생일', '이번 주 생일'], ['help', '도움말', '명령 목록']];
const TG_CMDS_VER = '1';   // 바꾸면 다음 tgPoll 때 텔레그램 메뉴(/ 버튼)를 다시 등록한다
function tgSetCommands_(p, token) {
  if (p.getProperty('TG_CMDS') === TG_CMDS_VER) return;
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/setMyCommands', { method: 'post', muteHttpExceptions: true,
    payload: { commands: JSON.stringify(TG_CMDS.map(([c, k, d]) => ({ command: c, description: k + ' — ' + d }))) } });
  if (r.getResponseCode() === 200) p.setProperty('TG_CMDS', TG_CMDS_VER);
}
/** 긴 답은 줄 단위로 나눠 여러 통으로 보낸다 (텔레그램 한 통 4096자) */
function tgSend_(text) {
  let buf = '';
  String(text).split('\n').forEach(line => {
    if (buf && buf.length + line.length + 1 > 3800) { tgNotify_(buf); buf = ''; }
    buf += (buf ? '\n' : '') + line.slice(0, 3800);
  });
  if (buf) tgNotify_(buf);
}
/** 말 속의 부서 이름 → 부서 키 */
function tgTeamArg_(s) {
  return /영아/.test(s) ? 'infant' : /유아/.test(s) ? 'toddler' : /유소년|초등/.test(s) ? 'youth' : /중고등|중등|고등|청소년/.test(s) ? 'teen' : '';
}
/** 말 속의 날짜(10/18, 10.18, 10월 18일) → 그 날짜가 든 주의 주일. 없으면 가장 최근 주일 */
function tgDateArg_(s) {
  const m = String(s).match(/(\d{1,2})\s*[\/.월]\s*(\d{1,2})/), last = lastSunKey_();
  if (!m) return last;
  let y = Number(last.slice(0, 4)); const mo = Number(m[1]), d = Number(m[2]);
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return last;
  if (keyStr_([y, mo, d]) > last) y--;   // 앞으로 올 날짜면 작년으로
  const x = new Date(Date.UTC(y, mo - 1, d, 12));
  return keyStr_(keyAdd_(keyStr_([y, mo, d]), -x.getUTCDay()));
}
function tgHelp_() {
  return '🤖 출석부 봇 명령\n' + TG_CMDS.map(([c, k, d]) => '• ' + k + ' (/' + c + '): ' + d).join('\n') +
    '\n\n부서(영아·유아·유소년·중고등)와 날짜(10/18)는 붙여도 되고 빼도 돼요. 이름만 보내도 찾아 줘요.\n답은 최대 ' + TG_POLL_MIN + '분 뒤에 와요.';
}
function tgCommand_(text) {
  const s = String(text || '').trim(); if (!s) return;
  const first = s.split(/\s+/)[0].replace(/^\//, '').replace(/@.*$/, '').toLowerCase(), rest = s.slice(s.split(/\s+/)[0].length).trim();
  const hit = TG_CMDS.find(([c, k]) => first === c || first === k) || (first === 'start' || first === '도움' || first === '?' ? TG_CMDS[4] : null);
  const cmd = hit ? hit[0] : '';
  if (cmd === 'help') return tgSend_(tgHelp_());
  if (cmd === 'att') return tgSend_(tgSunSummary_(tgDateArg_(rest), tgTeamArg_(rest)));
  if (cmd === 'absent') return tgSend_(tgAbsent_(tgDateArg_(rest), tgTeamArg_(rest)));
  if (cmd === 'bday') return tgSend_(tgBirthdays_(tgDateArg_(rest)));
  if (cmd === 'find') return tgSend_(rest ? tgFind_(rest) : '찾을 이름을 같이 보내 주세요. 예: 찾기 김하나');
  // 명령이 아니면 이름으로 찾아 보고, 없으면 도움말
  const f = s.length <= 10 && !/\s/.test(s) ? tgFind_(s, true) : '';
  tgSend_(f || '알아듣지 못했어요. 문의에 답장하려면 문의 알림을 길게 눌러 [답장]으로 보내 주세요.\n\n' + tgHelp_());
}
/** 주일 key 결석자: 반별 명단, 그 부서가 출석을 체크한 주일만 세어 연속 결석 주수를 붙인다 (최근 12주) */
function tgAbsent_(key, only) {
  const from = keyStr_(keyAdd_(key, -77));
  const R = readRoster_(true), W = readWeekly_(SHEET.att, k => k >= from && k <= key);
  const keys = W.use.slice().sort().reverse();
  const on = (p, k) => { const r = W.rows[p.id], c = W.cols[k]; return !!(r && c && Number(W.val(r, c)) > 0); };
  let out = '❌ ' + key.slice(5).replace('-', '/') + ' 결석', longs = [];
  TEAM_ORDER.filter(t => !only || t === only).forEach(t => {
    const st = R.list.filter(p => p.team === t && p.active !== false && p.group !== 'T'); if (!st.length) return;
    const weeks = keys.filter(k => st.some(p => on(p, k)));   // 이 부서가 출석을 체크한 주일
    if (weeks[0] !== key) { out += '\n\n[' + TEAMS[t].label + '] 이 주일 출석 체크가 없어요.'; return; }
    const streak = p => { let n = 0; for (const k of weeks) { if (on(p, k)) break; n++; } return n; };
    const ab = st.filter(p => !on(p, key));
    out += '\n\n[' + TEAMS[t].label + '] 결석 ' + ab.length + '/' + st.length + '명';
    groupsOf_(t).forEach(([g]) => {
      const secs = [...new Set(ab.filter(p => p.group === g).map(p => p.section || ''))].sort();
      secs.forEach(sec => {
        const L = ab.filter(p => p.group === g && (p.section || '') === sec).sort((a, b) => a.name.localeCompare(b.name, 'ko'));
        out += '\n• ' + clsName_(t, g, sec) + ': ' + L.map(p => { const n = streak(p), full = n >= weeks.length;
          if (n >= 3) longs.push(p.name + '(' + clsName_(t, g, sec) + ', ' + n + '주' + (full ? '+' : '') + ')');
          return p.name + (n >= 2 ? '(' + n + '주' + (full ? '+' : '') + ')' : ''); }).join(', ');
      });
    });
  });
  if (longs.length) out += '\n\n⚠️ 3주 이상 연속 결석 ' + longs.length + '명\n' + longs.join('\n');
  return out + '\n\n(괄호 안은 연속 결석 주수)';
}
/** 이름으로 사람 찾기: 반·담임·생일·최근 출석·연락처. quiet=true 면 못 찾았을 때 '' */
function tgFind_(q, quiet) {
  q = String(q).replace(/\s+/g, '');
  const R = readRoster_(true);
  let L = R.list.filter(p => p.name.replace(/\s+/g, '') === q);
  if (!L.length) L = R.list.filter(p => p.name.replace(/\s+/g, '').indexOf(q) >= 0);
  if (!L.length) return quiet ? '' : '🔍 "' + q + '" 이름을 찾지 못했어요.';
  L.sort((a, b) => (b.active !== false) - (a.active !== false));
  const key = lastSunKey_(), W = readWeekly_(SHEET.att, k => k >= keyStr_(keyAdd_(key, -56)) && k <= key), cls = readClasses_(), B = readBook_().map;
  const keys = W.use.slice().sort().reverse();
  const on = (p, k) => { const r = W.rows[p.id], c = W.cols[k]; return !!(r && c && Number(W.val(r, c)) > 0); };
  const out = L.slice(0, 5).map(p => {
    const t = p.team, c = cls[t + '_' + p.group] || {}, who = p.group === 'T' ? '' : ((p.section ? (c.teachers || {})[p.section] : '') || c.teacher || '');
    const weeks = keys.filter(k => R.list.some(x => x.team === t && on(x, k))).slice(0, 6);
    const b = B[p.id] || {};
    return '👤 ' + p.name + (p.active === false ? ' [' + (p.leftType || '제적') + ']' : '') + '\n' +
      TEAMS[t].label + ' · ' + (p.group === 'T' ? '교사' + (p.role ? '(' + p.role + ')' : '') : clsName_(t, p.group, p.section)) + (who ? ' · 담임 ' + who : '') +
      [[p.birthday ? '생일 ' + (p.lunar ? '음력 ' : '') + p.birthday : '', p.gender, p.isNew ? '새친구' : ''].filter(Boolean).join(' · ')].filter(Boolean).map(x => '\n' + x).join('') +
      (weeks.length ? '\n최근 출석: ' + weeks.slice().reverse().map(k => k.slice(5).replace('-', '/') + (on(p, k) ? '○' : '×')).join(' ') : '') +
      (b.phone ? '\n연락처: ' + b.phone : '') + (b.parentPhone ? '\n부모 연락처: ' + b.parentPhone + (b.guardian ? ' (' + b.guardian + ')' : '') : '') +
      (p.note ? '\n메모: ' + p.note : '');
  });
  return out.join('\n\n') + (L.length > 5 ? '\n\n… 외 ' + (L.length - 5) + '명. 이름을 더 정확히 보내 주세요.' : '');
}
/** 'yyyy-MM-dd' 에서 n일 뒤의 [연, 월, 일] (시간대 영향 없이 정오 기준) */
function keyAdd_(key, n) { const [y, m, d] = key.split('-').map(Number); const x = new Date(Date.UTC(y, m - 1, d + n, 12)); return [x.getUTCFullYear(), x.getUTCMonth() + 1, x.getUTCDate()]; }
const keyStr_ = ([y, m, d]) => y + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2);
/** 반 이름: 학년·나이반 + 분반 */
function clsName_(team, group, sec) { return groupLabel_(team, group) + (sec ? ' ' + sec : ''); }
/** y년 음력 '월/일' → 양력 [[월, 일], ...]. 지원하지 않는 환경이면 null */
const LUNAR_ = {};
function lunarMap_(y) {
  if (y in LUNAR_) return LUNAR_[y];
  let map = {};
  try {
    const f = new Intl.DateTimeFormat('en-u-ca-chinese', { month: 'numeric', day: 'numeric', timeZone: 'UTC' });
    for (let i = 0; i < 366; i++) {
      const d = new Date(Date.UTC(y, 0, 1 + i, 12)); if (d.getUTCFullYear() !== y) break;
      let lm = '', ld = '';
      f.formatToParts(d).forEach(x => { if (x.type === 'month') lm = x.value; else if (x.type === 'day') ld = x.value; });
      if (!lm || /[^0-9]/.test(lm)) continue;   // 윤달은 건너뜀
      const k = Number(lm) + '/' + Number(ld);
      (map[k] = map[k] || []).push([d.getUTCMonth() + 1, d.getUTCDate()]);
    }
    if (!Object.keys(map).length) map = null;
  } catch (e) { map = null; }
  return LUNAR_[y] = map;
}
/** 이번 주(주일 key ~ 토요일) 생일. 생일은 'YY/MM/DD' 또는 'MM/DD' */
/** sink(배열)를 주면 생일인 사람마다 활동 기록 줄을 모은다. staffOnly=true 면 교사(관리자 포함)만: 주일 자동 알림 */
function tgBirthdays_(key, sink, staffOnly) {
  const days = []; for (let i = 0; i < 7; i++) days.push(keyAdd_(key, i));
  const R = readRoster_(true), hits = []; let lunarFail = 0;
  R.list.forEach(p => {
    if (p.active === false || (staffOnly && p.group !== 'T')) return;
    const m = String(p.birthday || '').trim().match(/^(?:\d{2}\/)?(\d{1,2})\/(\d{1,2})$/); if (!m) return;
    const bm = Number(m[1]), bd = Number(m[2]);
    days.some(([y, mo, d], i) => {
      let ok;
      if (!p.lunar) ok = bm === mo && bd === d;
      else {
        const L = lunarMap_(y); if (!L) { if (i === 0) lunarFail++; return true; }
        ok = (L[bm + '/' + bd] || (bd === 30 ? L[bm + '/29'] : null) || []).some(([a, b]) => a === mo && b === d);
      }
      if (ok) hits.push({ p, i, mo, d });
      return ok;
    });
  });
  const [, m0, d0] = days[0], [, m6, d6] = days[6];
  let out = '🎂 이번 주 ' + (staffOnly ? '교사 ' : '') + '생일 (' + m0 + '/' + d0 + '~' + m6 + '/' + d6 + ')';
  if (!hits.length) out += '\n이번 주 생일인 ' + (staffOnly ? '교사가' : '사람이') + ' 없어요.';
  TEAM_ORDER.forEach(t => {
    const L = hits.filter(h => h.p.team === t).sort((a, b) => a.i - b.i || a.p.name.localeCompare(b.p.name, 'ko'));
    if (!L.length) return;
    out += '\n\n[' + TEAMS[t].label + ']';
    L.forEach(h => {
      const where = (h.p.group === 'T' ? '교사' + (h.p.role ? '(' + h.p.role + ')' : '') : clsName_(t, h.p.group, h.p.section)) + (h.p.lunar ? ' (음력)' : '');
      out += '\n• ' + h.mo + '/' + h.d + '(' + DOW_KO[h.i] + ') ' + h.p.name + ' · ' + where + (h.i === 0 ? ' 🎉오늘' : '');
      if (sink) sink.push([h.p.id, h.p.name, t, '생일', h.mo + '/' + h.d + '(' + DOW_KO[h.i] + ') 생일 · ' + where]);
    });
  });
  if (lunarFail) out += '\n\n※ 음력 생일 ' + lunarFail + '명은 날짜를 계산하지 못했어요. 앱의 생일 탭에서 확인해 주세요.';
  return out;
}
/** 주일 key 출석 요약: 부서별 학생·교사 출석(지난주 대비), 새친구 출석, 출석 체크가 하나도 없는 반 */
function tgSunSummary_(key, only) {
  const prev = keyStr_(keyAdd_(key, -7));
  const R = readRoster_(true), W = readWeekly_(SHEET.att, k => k === key || k === prev), cls = readClasses_();
  const on = (p, k) => { const r = W.rows[p.id], c = W.cols[k]; return !!(r && c && Number(W.val(r, c)) > 0); };
  const hasCol = !!W.cols[key];
  let out = '📊 ' + key.slice(5).replace('-', '/') + ' 주일 출석 요약', miss = '', sumNow = 0, sumAll = 0;
  TEAM_ORDER.filter(t => !only || t === only).forEach(t => {
    const ps = R.list.filter(p => p.team === t && p.active !== false);
    const st = ps.filter(p => p.group !== 'T'), te = ps.filter(p => p.group === 'T');
    if (!st.length && !te.length) return;
    const n = st.filter(p => on(p, key)).length, n0 = st.filter(p => on(p, prev)).length, tn = te.filter(p => on(p, key)).length;
    sumNow += n; sumAll += st.length;
    const diff = n - n0, newbies = st.filter(p => p.isNew && on(p, key)).map(p => p.name);
    out += '\n\n[' + TEAMS[t].label + ']\n학생 ' + n + '/' + st.length + '명' + (st.length ? ' (' + Math.round(n * 100 / st.length) + '%)' : '') +
      ' · 지난주 ' + (diff > 0 ? '+' + diff : diff < 0 ? String(diff) : '같음') + (te.length ? '\n교사 ' + tn + '/' + te.length + '명' : '') +
      (newbies.length ? '\n새친구 출석: ' + newbies.slice(0, 10).join(', ') + (newbies.length > 10 ? ' 외 ' + (newbies.length - 10) + '명' : '') : '');
    // 반(분반)마다 출석이 하나도 없으면 체크 안 한 반
    const byCls = {};
    st.forEach(p => { const k = p.group + '|' + (p.section || ''); (byCls[k] = byCls[k] || []).push(p); });
    const idle = groupsOf_(t).reduce((acc, [g]) => acc.concat(Object.keys(byCls).filter(k => k.split('|')[0] === g).sort()), [])
      .filter(k => !byCls[k].some(p => on(p, key)));
    if (idle.length) miss += '\n[' + TEAMS[t].label + '] ' + idle.map(k => {
      const [g, sec] = k.split('|'), c = cls[t + '_' + g] || {}, who = (sec ? (c.teachers || {})[sec] : '') || c.teacher || '';
      return clsName_(t, g, sec) + (who ? '(' + who + ')' : '');
    }).join(', ');
  });
  if (!only) out += '\n\n전체 학생 ' + sumNow + '/' + sumAll + '명';
  out += miss ? '\n\n✏️ 출석 체크 안 한 반' + miss : hasCol ? '\n\n✅ 모든 반이 출석을 체크했어요.' : '';
  return out;
}
function tgReply_(m) {
  const src = (m.reply_to_message && m.reply_to_message.text) || '';
  if (src.indexOf('[출석부 요청]') < 0) {
    // 오래된 메시지(트리거가 멈췄다 다시 돈 경우)에는 답하지 않는다
    if (Date.now() / 1000 - m.date < 1200) { try { tgCommand_(m.text); } catch (e) { tgNotify_('⚠️ 명령을 처리하지 못했어요: ' + e.message); } }
    return;
  }
  const field = k => { const r = src.match(new RegExp('^' + k + ': (.*)$', 'm')); return r && r[1].trim() !== '-' ? r[1].trim() : ''; };
  const name = field('이름'), email = field('이메일'), phone = field('연락처');
  const at = src.indexOf('\n요청사항: '), ask = at < 0 ? '' : src.slice(at + 7);
  const sms = tgSmsBtn_(phone, m.text);
  if (!email) {
    if (!sms) { tgNotify_('📵 ' + name + '님은 이메일도 연락처도 없어 답장할 수 없어요.'); return; }
    try { withLock_(() => logRows_({}, [['', name, '', '답장', '텔레그램 답장 문자 링크 (' + phone + '): ' + m.text.slice(0, 200)]])); } catch (e) {}
    tgNotify_('📱 ' + name + '님은 이메일이 없어요. 아래 버튼을 누르면 문자 앱이 열려요. [보내기]만 누르세요. (' + phone + ')', sms);
    return;
  }
  askReplySend_({}, { email: email, name: name, msg: m.text, ask: ask }, '텔레그램 답장');
  tgNotify_('✅ ' + name + '님께 메일로 답장했어요 (' + email + ')' + (sms ? '\n문자로도 보내려면 아래 버튼을 누르세요. (' + phone + ')' : ''), sms);
}
/** 문자 앱을 여는 GitHub 페이지(sms.html) 버튼. 번호가 없으면 null. 링크가 너무 길지 않게 내용은 500자까지 */
function tgSmsBtn_(phone, text) {
  const to = String(phone || '').replace(/[^0-9+]/g, '');
  if (to.length < 8) return null;
  const body = '[신당중앙교회 교육팀 출석부] ' + String(text).slice(0, 500);
  return { text: '📱 문자로 보내기', url: PUBLIC_URL + 'sms.html?to=' + encodeURIComponent(to) + '&body=' + encodeURIComponent(body) };
}
/** 문자 답장은 기기의 문자 앱으로 보내므로 기록만 남긴다 */
function askReplyLog_(auth, f) {
  if (!isRoot_(auth)) throw new Error('총관리자만 할 수 있어요');
  f = f || {};
  const name = String(f.name || '').trim().slice(0, 30), phone = String(f.phone || '').trim().slice(0, 30), msg = String(f.msg || '').trim().slice(0, 200);
  withLock_(() => logRows_(auth, [['', name, '', '답장', '문자 답장 (' + phone + '): ' + msg]]));
  return { ok: true };
}
function staffEdit_(auth, oldEmail, f) {
  const sc = staffMgr_(auth); if (!sc) throw new Error('관리자만 수정할 수 있어요');
  const o = normEmail_(oldEmail); f = staffFields_(f);
  if (auth.co && f.role === 'admin') throw new Error('공동관리자는 총관리자 계정을 지정할 수 없어요');
  const list = staffList_().map(s => ({ name: s.name, email: s.email, role: STAFF_ROLES.includes(s.role) ? s.role : 'team', team: s.team || '', pos: s.pos || '', duty: s.duty || '' }));
  const at = list.findIndex(s => normEmail_(s.email) === o);
  if (at < 0) throw new Error('등록된 이메일을 찾을 수 없어요');
  if (sc !== 'full') {
    if (!staffInScope_(sc, list[at])) throw new Error('이 관리자는 수정할 수 없어요');
    if (f.role !== 'team' || (sc !== 'all' && sc.indexOf(f.team) < 0)) throw new Error('자기 부서 임원으로만 바꿀 수 있어요');
    f.pos = sc === 'all' && (f.pos === '팀장' || f.pos === '') ? f.pos : list[at].pos; f.pw = '';
  }
  if (auth.co && list[at].role === 'admin') throw new Error('공동관리자는 총관리자 계정을 수정할 수 없어요');
  if (list.some((s, i) => i !== at && normEmail_(s.email) === f.email)) throw new Error('이미 등록된 이메일이에요');
  const map = Object.assign({}, coPwMap_()); const prev = map[o]; delete map[o];
  if (f.role !== 'admin') {   // 공동관리자·임원의 개인 비밀번호는 그대로 (공동관리자는 여기서 새로 정할 수도 있어요)
    if (f.pw && PW_ROLES.includes(f.role)) { if (!validCoPw_(f.pw)) throw new Error(PW_RULE_MSG); coSet_(map, f.email, f.pw); }
    else if (prev) map[f.email] = prev;
  }
  list[at] = { name: f.name, email: f.email, role: f.role, team: f.team, pos: f.pos, duty: f.duty };
  cfgSet_('staff', list);
  cfgSet_('coPw', Object.keys(map).length ? map : null);
  if (o !== f.email) {
    const P = qrProps_(), allProps = P.getProperties();
    Object.keys(allProps).forEach(k => {
      if (k.indexOf('qd:') !== 0) return;
      try {
        const rec = JSON.parse(allProps[k]); let hit = false;
        (rec.ids || []).forEach(x => { if (x.e && normEmail_(x.e) === o) { x.e = f.email; hit = true; } });
        if (hit) P.setProperty(k, JSON.stringify(rec));
      } catch (e) {}
    });
  }
  logRows_(auth, [['', f.name, f.team || '', '이메일', '등록 정보 수정: ' + o + ' → ' + f.email + ' (' + roleName_(f.role, f.team) + (f.pos ? ' · ' + f.pos : '') + ')']]);
  return staffView_(auth);
}
function staffDel_(auth, email) {
  const sc = staffMgr_(auth); if (!sc) throw new Error('관리자만 삭제할 수 있어요');
  const e = normEmail_(email); const all = staffList_();
  if (sc !== 'full' && !staffInScope_(sc, all.find(s => normEmail_(s.email) === e))) throw new Error('이 관리자는 삭제할 수 없어요');
  if (auth.co && all.some(s => normEmail_(s.email) === e && s.role === 'admin')) throw new Error('공동관리자는 총관리자 계정을 삭제할 수 없어요');
  const list = all.filter(s => normEmail_(s.email) !== e);
  if (list.length === all.length) throw new Error('등록된 이메일을 찾을 수 없어요');
  cfgSet_('staff', list.length ? list : null);
  const map = Object.assign({}, coPwMap_());
  if (map[e]) { delete map[e]; cfgSet_('coPw', Object.keys(map).length ? map : null); }
  logRows_(auth, [['', '', '', '이메일', '관리자 이메일 삭제: ' + e]]);
  return staffView_(auth);
}
/** 이메일 등록을 잘못해서 못 들어갈 때, 편집기에서 한 번 실행하면 등록 명단을 모두 지워요. */
function clearStaff() { cfgSet_('staff', null); cfgSet_('coPw', null); Logger.log('이메일 등록 명단을 모두 지웠어요.'); }

/**
 * 확인표: 번호가 맞았던 기기에 주는 서명. 누가 일부러 틀린 번호를 넣어 잠금이 걸려도,
 * 이미 들어와 있던 사람은 이 표로 계속 쓸 수 있다. (표는 번호마다 다르고 서버 비밀값 없이는 만들 수 없어요. 번호를 바꾸면 무효)
 */
let TOK_SECRET_ = null;
function tokSecret_() {
  if (TOK_SECRET_) return TOK_SECRET_;
  const p = PropertiesService.getScriptProperties(); let v = p.getProperty('tokSecret');
  if (!v) { v = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('tokSecret', v); }
  return (TOK_SECRET_ = v);
}
function tok_(kind, v) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(kind + '|' + v, tokSecret_())).slice(0, 24); }
function tokOk_(auth, kind, v) { const t = (auth.tk || {})[kind]; return !!t && String(t) === tok_(kind, v); }
/** 요청에 실려 온 번호를 걸러서, 맞는 번호만 남긴다. 틀린 번호는 시도 횟수에 세고, 잠겨 있으면 무시한다.
 *  임원 번호와 총관리자 비밀번호는 등록된 이메일(auth.email)도 맞아야 인정한다. (등록 명단이 비어 있으면 총관리자만 통과) */
function cleanAuth_(auth) {
  auth = auth || {};
  let cfg, pw;
  try { cfg = cfgAll_(); pw = adminPw_(); } catch (e) { return auth; }
  const codes = cfg.codes || {}, gc = cfg.gradeCodes || {};
  const staff = Array.isArray(cfg.staff) ? cfg.staff : [];
  const em = normEmail_(auth.email);
  const mailOk = !!em && staff.some(s => normEmail_(s.email) === em);
  const out = { admin: '', codes: {}, gcodes: {}, email: mailOk ? em : '', co: false, co2: false, coName: '', tk: {}, mgr: {} };
  if (typeof auth.tk !== 'object' || !auth.tk) auth.tk = {};
  // 잠겨 있으면 틀린 번호는 세지 않고, 맞는 번호도 확인표가 있을 때만 받는다
  if (auth.admin) {
    const a = String(auth.admin), locked = guardLocked_('admin');
    if (pw && a === pw) { if ((!locked || tokOk_(auth, 'admin', a)) && (!staff.length || (mailOk && staffRole_(em) === 'admin'))) { out.admin = a; out.tk.admin = tok_('admin', a); } }
    else if (pw) out.adminBad = true; // 기기에 기억된 비밀번호가 틀림 → 화면이 지우고 다시 로그인하라고 안내
    if (pw && a !== pw && !locked) guardFail_('admin', a, '', '총관리자: 기기에 기억된 옛 비밀번호가 자동으로 전송됨 (직접 로그인한 것 아님, 기기 이메일: ' + mailTag_(auth.email) + ')');
  }
  // 공동관리자·임원: 등록된 이메일 + 개인 비밀번호 (공동관리자 = 총관리자와 같은 권한(이메일 등록 제외), 임원 = 담당 부서 전체)
  if (!out.admin && auth.copw && mailOk) {
    const me = staff.find(s => normEmail_(s.email) === em);
    const kind = 'co:' + em;
    if (me && me.role !== 'admin') {
      const locked = guardLocked_(kind);
      const rec = (cfg.coPw || {})[em];
      if (rec && coCheck_(em, auth.copw)) { if (!locked || tokOk_(auth, kind, auth.copw)) {
        if (PW_ROLES.includes(me.role)) { if (pw) { out.admin = pw; out.co = true; if (me.role === 'co2') out.co2 = true; out.coName = me.name; } }
        else if (TEAMS[me.team]) out.mgr[me.team] = 1;
        out.staffPw = em; out.tk[kind] = tok_(kind, auth.copw); } }
      else if (!locked) guardFail_(kind, String(auth.copw), '', roleName_(me.role, me.team) + ' ' + me.name + ': 기기에 기억된 옛 비밀번호가 자동으로 전송됨 (직접 로그인한 것 아님)');
    }
  }
  // 담임 선생님 개인 비밀번호 (tcodes[교사 id] = { t: 부서, c: 번호, n: 이름 }). 맞으면 그 선생님으로 로그인
  const tc = cfg.tcodes || {};
  out.tcodes = {}; out.hrMe = {};
  Object.keys(auth.tcodes || {}).forEach(pid => {
    const v = String(auth.tcodes[pid] || ''), r = tc[pid]; if (!v || !r || !TEAMS[r.t]) return;
    const locked = guardLocked_('entry:' + r.t);
    if (tcCheck_(r, v)) { if (!locked || tokOk_(auth, 'h:' + pid, gcNorm_(v))) { out.tcodes[pid] = v; out.tk['h:' + pid] = tok_('h:' + pid, gcNorm_(v)); out.hrMe[r.t] = { pid, name: String(r.n || ''), own: true }; } }
    else if (!locked) guardFail_('entry:' + r.t, v, r.t, TEAMS[r.t].label + ' ' + String(r.n || '').slice(0, 20) + ' 선생님: 기기에 기억된 옛 비밀번호가 자동으로 전송됨 (직접 로그인한 것 아님)');
  });
  out.scodes = {};
  applyDevs_(auth, out, cfg, pw, staff);
  return out;
}

/* ---------------- QR 로그인 (기기 등록) ----------------
 * 관리자가 만든 1회용 링크(…/#qr=코드, 24시간)를 열면 그 기기에 '기기 표'를 준다. 그 뒤로는 비밀번호 없이 그 사람으로 들어온다.
 *  - 기기 표는 비밀번호와 따로라서 번호를 바꿔도 등록된 기기는 계속 쓴다. 끊으려면 [로그인 기기]에서 해제.
 *  - 누구로 들어오는지(ids): admin(총관리자) / co(공동관리자1·2) / team(임원, 부서) / hr(교사: 담임이면 그 반, 아니면 일정 탭만)
 *  - 저장: 스크립트 속성 qi:<코드 해시> = 링크, qd:<기기 id> = 기기 (서로 다른 칸이라 동시에 써도 겹치지 않아요)
 * 만들 수 있는 사람: 총관리자 = 모두 / 공동관리자1·2 = 임원·교사 / 임원 = 자기 부서 교사 / 누구나 = 본인의 다른 기기
 * 한 사람당 휴대폰(태블릿 포함) 1대 + PC 1대: 같은 종류의 새 기기를 등록하면 그 사람의 이전 기기는 해제된다 */
const QR_HOURS = 24;
const QR_MAX_DEV = 6;   // 한 기기(브라우저)에 기억하는 기기 표 수
function qrHash_(v) { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'qr|' + String(v))).slice(0, 32); }
function qrRand_() { return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''); }
function qrProps_() { return PropertiesService.getScriptProperties(); }
function qrGet_(key) { try { return JSON.parse(qrProps_().getProperty(key) || 'null'); } catch (e) { return null; } }
function qrToday_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function qrNow_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'); }
/** 지금 교사 명단에 있는 선생님 { id: {t, n} } (5분 캐시) */
function activeTeachers_() {
  const c = CacheService.getScriptCache(); const hit = c.get('qr-tch2');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const m = {}; readRoster_(true).list.forEach(p => { if (p.group === 'T' && p.active !== false) m[p.id] = { t: p.team, n: p.name, r: String(p.role || '') }; });
  try { c.put('qr-tch2', JSON.stringify(m), 300); } catch (e) {}
  return m;
}
/** 기기 표가 맞으면 그 사람의 권한을 out 에 더한다 (cleanAuth_ 끝에서) */
function applyDevs_(auth, out, cfg, pw, staff) {
  out.dev = []; out.devBad = [];
  const list = (Array.isArray(auth.dev) ? auth.dev : []).map(String).filter(s => /^[0-9a-f]{12,40}\.[0-9a-f]{32,80}$/.test(s)).slice(0, QR_MAX_DEV);
  if (!list.length) return;
  let tch = null;
  list.forEach(tokn => {
    const [id, sec] = tokn.split('.'); const rec = qrGet_('qd:' + id);
    if (!rec || rec.h !== qrHash_(sec)) { out.devBad.push(tokn); return; }
    out.dev.push(id); out.viaDev = true;
    (rec.ids || []).forEach(x => {
      const em = normEmail_(x.e), s = em ? staff.find(z => normEmail_(z.email) === em) : null;
      if (x.k === 'admin' && pw && (!staff.length || (s && s.role === 'admin'))) { out.admin = pw; out.co = false; out.co2 = false; out.coName = ''; if (em) out.email = em; }   // 총관리자 본인이 [내 다른 기기 추가]로 등록한 기기
      else if (x.k === 'co' && pw && s && PW_ROLES.includes(s.role) && !(out.admin && !out.co)) { out.admin = pw; out.co = true; out.co2 = s.role === 'co2'; out.coName = s.name; out.email = em; out.staffDev = em; }
      else if (x.k === 'team' && s && s.role !== 'admin' && !PW_ROLES.includes(s.role) && TEAMS[x.t] && (!s.team || s.team === x.t)) { out.mgr[x.t] = 1; if (!out.email) out.email = em; if (!out.staffDev) out.staffDev = em; }
      else if (x.k === 'hr' && TEAMS[x.t]) { tch = tch || activeTeachers_(); const p = tch[x.pid]; if (p && p.t === x.t) out.hrMe[x.t] = { pid: x.pid, name: String(p.n || ''), own: true, dev: true }; }
    });
    if (rec.seen !== qrToday_()) { rec.seen = qrToday_(); try { qrProps_().setProperty('qd:' + id, JSON.stringify(rec)); } catch (e) {} }
  });
}
/** 사람 하나(ids)를 읽기 좋은 이름으로: '홍길동 (유소년교육팀 교사)' */
function qrLabel_(ids) {
  const st = staffList_(); let tch = null;
  return (ids || []).map(x => {
    if (x.k === 'hr') { tch = tch || activeTeachers_(); const p = tch[x.pid]; return (p ? p.n : (x.n || '교사')) + ' (' + (TEAMS[x.t] ? TEAMS[x.t].label : '') + ' 교사)'; }
    const s = st.find(z => normEmail_(z.email) === normEmail_(x.e));
    return (s ? s.name : (x.e || '')) + ' (' + (x.k === 'admin' ? '총관리자' : x.k === 'co' ? (s && s.role === 'co2' ? '공동관리자2' : '공동관리자1') : roleName_('team', x.t)) + ')';
  }).join(', ');
}
/** 이 요청을 보낸 사람이 그 ids 를 다룰(링크 만들기·기기 해제) 수 있는지 */
function qrCanManage_(auth, ids) {
  if (!ids || !ids.length) return false;
  if (isRoot_(auth)) return true;
  return ids.every(x => {
    if (x.k === 'hr') return deptLevel_(auth, x.t);
    if (x.k === 'team') return isAdmin_(auth) || staffInScope_(staffMgr_(auth), staffList_().find(s => normEmail_(s.email) === normEmail_(x.e)));   // 공동관리자1·2, 목사·전도사·팀장
    return false;                                // 총관리자·공동관리자 기기는 총관리자만
  });
}
/** 지금 로그인한 본인 (다른 기기 추가용) */
function qrSelfIds_(auth) {
  const ids = [];
  if (isRoot_(auth)) ids.push({ k: 'admin', e: normEmail_(auth.email) });   // 총관리자는 본인 다른 기기 추가만 (남이 만들어 줄 수는 없어요)
  else if (isAdmin_(auth) && auth.co) ids.push({ k: 'co', e: normEmail_(auth.email) });
  else if (auth.email) TEAM_ORDER.forEach(t => { if (deptLevel_(auth, t)) ids.push({ k: 'team', e: normEmail_(auth.email), t }); });
  let cls = null, R = null;
  TEAM_ORDER.forEach(t => {
    const sc = scope_(auth, t); if (sc.level === 'dept' || sc.level === 'none') return;
    let me = sc.level === 'grade' ? meOf_(auth, t, cls || (cls = readClasses_())) : (auth.hrMe || {})[t];
    let pid = me ? String(me.pid || '') : '';
    const name = me ? me.name : sc.name;
    if (!pid && name) { R = R || readRoster_(true); const p = subTeacher_(R, t, name); if (p) pid = p.id; }
    if (pid) ids.push({ k: 'hr', t, pid });
  });
  return ids;
}
/** 링크 만들기. spec: {k:'hr', t, pid} | {k:'staff', e} | {k:'self'} */
function qrMake_(auth, spec) {
  spec = spec || {}; let ids;
  if (spec.k === 'self') { ids = qrSelfIds_(auth); if (!ids.length) throw new Error('먼저 로그인한 뒤에 만들 수 있어요'); }
  else if (spec.k === 'hr') {
    const t = String(spec.t || ''), pid = String(spec.pid || ''); let p = activeTeachers_()[pid];
    if (!p) { try { CacheService.getScriptCache().remove('qr-tch2'); } catch (e) {} p = activeTeachers_()[pid]; }   // 방금 추가한 선생님
    if (!TEAMS[t] || !p || p.t !== t) throw new Error('교사 명단에서 그 선생님을 찾지 못했어요');
    ids = [{ k: 'hr', t, pid }];
  } else if (spec.k === 'staff') {
    const e = normEmail_(spec.e), s = staffList_().find(z => normEmail_(z.email) === e);
    if (!s) throw new Error('등록된 이메일이 아니에요');
    if (s.role === 'admin') throw new Error('총관리자는 비밀번호로 로그인해요. 본인의 다른 기기는 [내 다른 기기 추가]로 등록하세요');
    else if (PW_ROLES.includes(s.role)) ids = [{ k: 'co', e }];
    else { if (!TEAMS[s.team]) throw new Error('이 임원은 담당 부서가 지정되어 있지 않아요. [수정]에서 부서를 먼저 정해 주세요'); ids = [{ k: 'team', e, t: s.team }]; }
  } else throw new Error('누구의 QR인지 알 수 없어요');
  if (spec.k !== 'self' && !qrCanManage_(auth, ids)) throw new Error('이 사람의 QR을 만들 권한이 없어요');
  const P = qrProps_(); const now = Date.now();
  // 지난 링크 정리 + 같은 사람에게 전에 만든(아직 안 쓴) QR은 무효로 (가장 최근 QR 하나만 쓸 수 있어요)
  const keys = qrKeys_(ids); let old = 0;
  const all = P.getProperties(); Object.keys(all).forEach(k => { if (k.indexOf('qi:') === 0) { try { const v = JSON.parse(all[k]);
    if (v.exp < now) P.deleteProperty(k); else if (qrKeys_(v.ids).some(x => keys.indexOf(x) >= 0)) { P.deleteProperty(k); old++; } } catch (e) { P.deleteProperty(k); } } });
  const code = qrRand_(), label = qrLabel_(ids), by = actorLabel_(auth, (ids[0] || {}).t || '');
  P.setProperty('qi:' + qrHash_(code), JSON.stringify({ ids, exp: now + QR_HOURS * 3600e3, by, self: spec.k === 'self' }));
  const team = (ids.find(x => x.t) || {}).t || '';
  try { withLock_(() => { touchVer_('x'); logRows_(auth, [['', label.split(' (')[0], team, '로그인', label + (spec.k === 'self' ? ' 본인 기기 추가용' : '') + ' QR 로그인 링크를 만듦' + (old ? ' (전에 만든 QR ' + old + '개 무효)' : '')]]); }); } catch (e) {}
  const base = publicUrl_().replace(/#.*$/, '');
  return { url: base + '#qr=' + code, label, old, exp: Utilities.formatDate(new Date(now + QR_HOURS * 3600e3), tz_(), 'M월 d일 HH:mm') };
}
/** 링크를 연 사람에게 누구 전용 QR인지 먼저 보여 준다 (쓰지는 않음). 쓸 수 없는 QR이면 사용할 때와 같은 안내 */
function qrPeek_(code) {
  code = String(code || ''); if (!/^[0-9a-f]{64}$/.test(code)) throw new Error('QR 주소가 올바르지 않아요');
  const inv = qrGet_('qi:' + qrHash_(code));
  if (!inv || inv.exp < Date.now()) { Utilities.sleep(500); throw new Error('이미 사용했거나, 기간(24시간)이 지났거나, 새 QR을 만들어서 무효가 된 QR이에요. 관리자에게 가장 최근 QR을 요청해 주세요.'); }
  return { label: qrLabel_(inv.ids), exp: Utilities.formatDate(new Date(inv.exp), tz_(), 'M월 d일 HH:mm') };
}
/** 링크 사용: 맞으면 링크를 지우고 기기 표를 준다 (로그인 없이 부름) */
function qrUse_(code) {
  code = String(code || ''); if (!/^[0-9a-f]{64}$/.test(code)) throw new Error('QR 주소가 올바르지 않아요');
  const P = qrProps_(), key = 'qi:' + qrHash_(code);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let inv;
  try { inv = qrGet_(key); if (inv) P.deleteProperty(key); } finally { lock.releaseLock(); }
  if (!inv || inv.exp < Date.now()) { Utilities.sleep(500); throw new Error('이미 사용했거나, 기간(24시간)이 지났거나, 새 QR을 만들어서 무효가 된 QR이에요. 관리자에게 가장 최근 QR을 요청해 주세요.'); }
  const r = qrRegister_(inv.ids, inv.by || '');
  try { withLock_(() => { touchVer_('x'); logRows_({ sys: 'QR 로그인' }, [['', r.label.split(' (')[0], r.team, '로그인', r.label + ' QR로 기기 등록' + devTag_() + r.goneTx + (inv.by ? ' · 만든 사람: ' + inv.by : '')]]); }); } catch (e) {}
  tgNotify_('🔑 QR 로그인 기기 등록\n' + r.label + devTag_() + r.goneTx + (inv.by ? '\n만든 사람: ' + inv.by : '') + '\n\n본인이 아니라면 [로그인 기기]에서 해제하세요.');
  return { tok: r.tok, label: r.label, gone: r.gone };
}
/** 기기 등록 (QR·비밀번호 로그인 공통). 관리자는 같은 종류(휴대폰/PC) 기기를 해제 (휴대폰 1대 + PC 1대), 선생님은 종류 상관없이 이전 기기를 해제 (1대) */
function qrRegister_(ids, by) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const P = qrProps_(), id = qrRand_().slice(0, 20), sec = qrRand_();
    const label = qrLabel_(ids), cls = qrCls_(REQ_DEV_);
    const keys = qrKeys_(ids), all = P.getProperties(); let gone = 0, goneAny = false;
    // 선생님은 휴대폰·PC 상관없이 1대만. 관리자 명단에도 이름이 있는 선생님(겸 관리자)은 관리자처럼 휴대폰 1대 + PC 1대
    const tch = activeTeachers_(), stNames = staffList_().map(s => String(s.name || '')).filter(Boolean);
    const both = pid => { const t = tch[pid]; return !!t && stNames.some(n => sameName_(n, t.n)); };
    const hrK = ids.filter(x => x.k === 'hr' && !both(x.pid)).map(x => 'hr:' + x.pid), stK = keys.filter(x => hrK.indexOf(x) < 0);
    Object.keys(all).forEach(k => { if (k.indexOf('qd:') !== 0) return; let r; try { r = JSON.parse(all[k]); } catch (e) { return; }
      const rk = qrKeys_(r.ids), same = qrCls_(r.dv) === cls;
      if (rk.some(x => hrK.indexOf(x) >= 0) || (same && rk.some(x => stK.indexOf(x) >= 0))) { P.deleteProperty(k); gone++; if (!same) goneAny = true; } });
    P.setProperty('qd:' + id, JSON.stringify({ h: qrHash_(sec), ids, n: label, dv: REQ_DEV_ || '', at: qrNow_(), by: by || '', seen: qrToday_() }));
    const gn = gone ? (goneAny ? '기기' : cls === 'pc' ? 'PC' : '휴대폰') : '';
    return { tok: id + '.' + sec, label, gone: gn, goneTx: gn ? ' · 이전 ' + gn + ' 해제' : '', team: (ids.find(x => x.t) || {}).t || '' };
  } finally { lock.releaseLock(); }
}
/** 이름 + 개인 비밀번호로 들어온 선생님: 이 기기를 등록해서 기기 표로 바꾼다 (비밀번호는 기기에 남기지 않고, 2대 제한도 똑같이) */
function devBind_(auth) {
  const tc = cfgAll_().tcodes || {}, ids = [];
  Object.keys(auth.tcodes || {}).forEach(pid => { const r = tc[pid], m = r && (auth.hrMe || {})[r.t]; if (m && m.pid === pid && !m.dev) ids.push({ k: 'hr', t: r.t, pid }); });
  if (auth.staffPw && !auth.staffDev) { const s = staffList_().find(z => normEmail_(z.email) === auth.staffPw);
    if (s && PW_ROLES.includes(s.role)) ids.push({ k: 'co', e: auth.staffPw }); else if (s && s.role !== 'admin' && TEAMS[s.team]) ids.push({ k: 'team', e: auth.staffPw, t: s.team }); }
  if (!ids.length) return null;
  const r = qrRegister_(ids, '비밀번호 로그인');
  try { withLock_(() => { touchVer_('x'); logRows_(auth, [['', r.label.split(' (')[0], r.team, '로그인', r.label + ' 비밀번호로 기기 등록' + devTag_() + r.goneTx]]); }); } catch (e) {}
  if (r.gone) tgNotify_('🔑 비밀번호 로그인 기기 등록\n' + r.label + devTag_() + r.goneTx + '\n\n본인이 아니라면 비밀번호를 초기화하고 [로그인 기기]에서 해제하세요.');
  return { tok: r.tok, label: r.label, gone: r.gone };
}
/** 기기 종류: 휴대폰·태블릿 = 'm', 그 밖(PC 등) = 'pc' (화면이 보낸 기기 이름 기준) */
function qrCls_(dv) { return /^(휴대폰|태블릿)/.test(String(dv || '')) ? 'm' : 'pc'; }
/** 사람 구분 열쇠: 교사 = 명단 id, 관리자 = 이메일 */
function qrKeys_(ids) { return (ids || []).map(x => x.k === 'hr' ? 'hr:' + x.pid : 'st:' + normEmail_(x.e)); }
/** QR을 만들 수 있는 관리자 명단 (공동관리자2는 boot 에 명단이 오지 않아서 따로) */
function qrStaff_(auth) {
  if (!isAdmin_(auth) && !staffMgr_(auth)) throw new Error('관리자만 볼 수 있어요');
  const pwm = coPwMap_(), devs = {}, all = qrProps_().getProperties(), tch = activeTeachers_();   // 개인 비밀번호 유무와 QR 기기 수 (선생님 관리처럼)
  Object.keys(all).forEach(k => { if (k.indexOf('qd:') !== 0) return; try { (JSON.parse(all[k]).ids || []).forEach(x => { if (x.e) devs[normEmail_(x.e)] = (devs[normEmail_(x.e)] || 0) + 1; }); } catch (e) {} });
  return staffList_().filter(s => s.role !== 'admin').filter(s => { const k = s.role === 'admin' ? 'admin' : PW_ROLES.includes(s.role) ? 'co' : 'team'; return qrCanManage_(auth, [{ k, e: s.email, t: s.team || '' }]); })
    .map(s => ({ name: s.name, email: s.email, role: STAFF_ROLES.includes(s.role) ? s.role : 'team', team: s.team || '', pos: s.pos || '', duty: s.duty || staffDuty_(s, tch), hasPw: !!pwm[normEmail_(s.email)], devs: devs[normEmail_(s.email)] || 0 }));
}
/** [로그인 관리] 옆 숫자: 부서별 선생님·관리자 중 QR 기기를 등록한 사람 수 [등록, 전체], 로그인 기기는 그 합 */
function qrCounts_(auth) {
  const out = { t: {} }, cnt = l => [l.filter(x => x.devs > 0).length, l.length];
  Object.keys(TEAMS).forEach(t => { if (!deptLevel_(auth, t)) return; try { out.t[t] = cnt(tchStatus_(auth, t)); } catch (e) {} });
  if (isAdmin_(auth) || staffMgr_(auth)) try { out.s = cnt(qrStaff_(auth)); } catch (e) {}
  const v = Object.values(out.t).concat(out.s ? [out.s] : []); out.d = [v.reduce((a, x) => a + x[0], 0), v.reduce((a, x) => a + x[1], 0)];
  return out;
}
/** 로그인 기기 목록: 내가 다룰 수 있는 기기 + 이 기기 */
function qrDevs_(auth) {
  const all = qrProps_().getProperties(), mine = auth.dev || [], out = [];
  Object.keys(all).forEach(k => {
    if (k.indexOf('qd:') !== 0) return; const id = k.slice(3); let r; try { r = JSON.parse(all[k]); } catch (e) { return; }
    const me = mine.indexOf(id) >= 0; if (!me && !qrCanManage_(auth, r.ids)) return;
    out.push({ id, n: qrLabel_(r.ids), dv: r.dv || '', at: r.at || '', seen: r.seen || '', by: r.by || '', me });
  });
  return out.sort((a, b) => String(b.at).localeCompare(String(a.at)));
}
/** 기기 해제 (ids 가 여럿이면 모두). 내 기기는 누구나, 다른 사람 기기는 다룰 권한이 있을 때만 */
function qrRevoke_(auth, ids) {
  ids = (Array.isArray(ids) ? ids : [ids]).map(String).slice(0, 50); const P = qrProps_(); let n = 0;
  ids.forEach(id => {
    const r = qrGet_('qd:' + id); if (!r) return;
    const me = (auth.dev || []).indexOf(id) >= 0;
    if (!me && !qrCanManage_(auth, r.ids)) throw new Error('이 기기를 해제할 권한이 없어요');
    P.deleteProperty('qd:' + id); n++;
    if (!me) try { withLock_(() => { touchVer_('x'); logRows_(auth, [['', String(r.n || '').split(' (')[0], (r.ids.find(x => x.t) || {}).t || '', '로그인', (r.n || '') + ' QR 로그인 기기를 해제함' + (r.dv ? ' (' + r.dv + ')' : '')]]); }); } catch (e) {}
  });
  return n;
}
/** 편집기에서 실행: QR 로그인 기기와 링크를 모두 지운다 (모든 기기가 다시 로그인해야 해요) */
function clearQrDevices() {
  const P = qrProps_(); Object.keys(P.getProperties()).forEach(k => { if (k.indexOf('qd:') === 0 || k.indexOf('qi:') === 0) P.deleteProperty(k); });
  Logger.log('QR 로그인 기기와 링크를 모두 지웠어요.');
}
/** 잠금을 지금 풀고 싶을 때 편집기에서 한 번 실행하세요. */
function clearLoginLocks() {
  const keys = []; ['admin', 'mail', 'login'].concat(TEAM_ORDER.map(t => 'entry:' + t), staffList_().map(s => 'co:' + normEmail_(s.email))).forEach(k => ['lock-', 'fail-', 'strike-'].forEach(p => keys.push(p + k)));
  CacheService.getScriptCache().removeAll(keys);
  Logger.log('로그인 잠금을 모두 풀었어요.');
}
/**
 * 시트 수식 주입 막기: 사람이 적은 글(이름·메모·주소 등)이 = + - @ 로 시작하면 구글 시트가 수식으로 실행해요.
 * 그런 글은 앞에 ' 를 붙여 글자로 저장해요 (시트에는 ' 가 보이지 않고, 읽을 때도 원래 글 그대로예요).
 * 이미 텍스트 서식(@)인 칸(ID, 날짜, 연락처 칸)에는 쓰지 않아요.
 */
function safeCell_(v) { return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v; }

/** 시트 시간대. 엑셀에서 바꾼 시트는 시간대가 비어 있을 수 있어서, 비어 있으면 서울로 정해 둔다. */
let TZ_CACHE_ = null;
function tz_() {
  if (TZ_CACHE_) return TZ_CACHE_;
  const ss = ss_();
  let tz = '';
  try { tz = ss.getSpreadsheetTimeZone(); } catch (e) {}
  if (!tz) {
    tz = 'Asia/Seoul';
    try { ss.setSpreadsheetTimeZone(tz); } catch (e) {}
  }
  return (TZ_CACHE_ = tz);
}

/** 필요한 탭·머리글을 만든다. 쓰기가 많아 느리므로 6시간에 한 번만 실제로 점검한다. */
function ensureSheets_(force) {
  const cache = CacheService.getScriptCache(), key = 'sheets-ok-v1';
  if (!force && cache.get(key)) return;
  ensureSheetsNow_();
  try { cache.put(key, '1', 21600); } catch (e) {}
}
function ensureSheetsNow_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET.roster);
  if (!sh) { sh = ss.insertSheet(SHEET.roster); sh.getRange(1, 1, 1, RCOLS.length).setValues([RCOLS]).setFontWeight('bold'); sh.setFrozenRows(1); }
  else {
    const hdr = sh.getRange(1, 1, 1, RCOLS.length).getValues()[0];
    RCOLS.forEach((h, i) => { if (!hdr[i]) sh.getRange(1, i + 1).setValue(h).setFontWeight('bold'); });
  }
  fmt_(sh.getRange('A:A'), '@'); fmt_(sh.getRange('G:G'), '@');
  [SHEET.att, SHEET.tal].forEach(name => {
    let w = ss.getSheetByName(name);
    if (!w) {
      w = ss.insertSheet(name);
      w.getRange(1, 1, 1, 6).setValues([['ID', '부서', '반', '분반', '이름', name === SHEET.att ? '출석 합계' : '달란트 합계']]).setFontWeight('bold');
      w.setFrozenRows(1);
    }
  });
  let cls = ss.getSheetByName(SHEET.cls);
  if (!cls) { cls = ss.insertSheet(SHEET.cls); cls.getRange(1, 1, 1, 5).setValues([['부서', '반', '분반', '담임', '재적 학생']]).setFontWeight('bold'); }
  let book = ss.getSheetByName(SHEET.book);
  if (!book) { book = ss.insertSheet(SHEET.book); book.getRange(1, 1, 1, BCOLS.length).setValues([BCOLS]).setFontWeight('bold'); book.setFrozenRows(1); }
  fmt_(book.getRange('A:A'), '@'); fmt_(book.getRange('E:F'), '@');
  let memo = ss.getSheetByName(SHEET.memo);
  if (!memo) { memo = ss.insertSheet(SHEET.memo); memo.getRange(1, 1, 1, 6).setValues([['날짜', 'ID', '이름', '부서', '반', '메모']]).setFontWeight('bold'); memo.setFrozenRows(1); }
  fmt_(memo.getRange('A:B'), '@');
  let ev = ss.getSheetByName(SHEET.events);
  if (!ev) { ev = ss.insertSheet(SHEET.events); ev.getRange(1, 1, 1, 7).setValues([['ID', '부서', '날짜', '제목', '내용', '공유문서', '수정일']]).setFontWeight('bold'); ev.setFrozenRows(1); }
  fmt_(ev.getRange('A:C'), '@');
  let hist = ss.getSheetByName(SHEET.hist);
  if (!hist) { hist = ss.insertSheet(SHEET.hist); hist.getRange(1, 1, 1, 7).setValues([['시각', 'ID', '이름', '부서', '구분', '내용', '바꾼 사람']]).setFontWeight('bold'); hist.setFrozenRows(1); }
  fmt_(hist.getRange('A:B'), '@');
  let lib = ss.getSheetByName(SHEET.lib);
  if (!lib) { lib = ss.insertSheet(SHEET.lib); lib.getRange(1, 1, 1, 8).setValues([['파일ID', '부서', '파일 이름', '공개', '크기', '형식', '올린 날', '올린 사람']]).setFontWeight('bold'); lib.setFrozenRows(1); }
  let cfg = ss.getSheetByName(SHEET.cfg);
  if (!cfg) {
    cfg = ss.insertSheet(SHEET.cfg);
    cfg.getRange('A1:C1').setValues([['관리자 비밀번호', '', '← B1 칸에 관리자 비밀번호를 적으세요. 이 탭은 앱 설정용이라 지우거나 고치지 마세요.']]);
    cfg.getRange('A2:B2').setValues([['키', '값(JSON)']]).setFontWeight('bold');
    cfg.hideSheet();
  }
  if (!cfg.getRange('E1').getValue()) {
    cfg.getRange('E1:G2').setValues([
      ['바로가기 주소', '', '← (선택) 선생님들께 보낼 바로가기 주소. 비워 두면 코드의 PUBLIC_URL(없으면 웹앱 주소 …/exec)을 씁니다.'],
      ['아이콘 주소', '', '← (선택) 브라우저 탭 아이콘 그림 주소. 비워 두면 기본 아이콘.']]);
  }
}

/* ---------------- 설정 저장소 (_설정 탭) ---------------- */
function cfgSheet_() { return ss_().getSheetByName(SHEET.cfg); }
let ADMIN_PW_ = null;
function adminPw_() { if (ADMIN_PW_ === null) ADMIN_PW_ = String(cfgSheet_().getRange('B1').getValue() || '').trim(); return ADMIN_PW_; }
// 한 번의 요청 안에서 _설정 탭을 여러 번 읽지 않도록 값을 기억해 둔다 (cfgSet_ 이 지운다).
let CFG_ROWS_ = null, CFG_OBJ_ = null;
function cfgAll_() {
  if (CFG_OBJ_) return CFG_OBJ_;
  if (!CFG_ROWS_) {
    // 설정은 1분 반쯤 캐시해요. 쓰기(withLock_)나 cfgSet_ 이 있으면 바로 비워요. (100KB를 넘으면 캐시되지 않고 매번 읽어요)
    let hit = null; try { hit = CacheService.getScriptCache().get('cfg-rows'); } catch (e) {}
    if (hit) { try { CFG_ROWS_ = JSON.parse(hit); } catch (e) { CFG_ROWS_ = null; } }
    if (!CFG_ROWS_) {
      const sh = cfgSheet_(); const n = sh.getLastRow();
      CFG_ROWS_ = n >= 3 ? sh.getRange(3, 1, n - 2, 2).getValues() : [];
      try { CacheService.getScriptCache().put('cfg-rows', JSON.stringify(CFG_ROWS_), 90); } catch (e) {}
    }
  }
  const out = {};
  CFG_ROWS_.forEach(r => { if (r[0]) { try { out[r[0]] = JSON.parse(r[1]); } catch (e) {} } });
  return (CFG_OBJ_ = out);
}
function cfgSet_(key, val) {
  CFG_ROWS_ = null; CFG_OBJ_ = null;
  try { CacheService.getScriptCache().remove('cfg-rows'); } catch (e) {}
  const sh = cfgSheet_(); const n = sh.getLastRow();
  const keys = n >= 3 ? sh.getRange(3, 1, n - 2, 1).getValues().map(r => r[0]) : [];
  const i = keys.indexOf(key);
  if (val === null) { if (i >= 0) sh.deleteRow(i + 3); return; }
  const json = JSON.stringify(val);
  if (i >= 0) sh.getRange(i + 3, 2).setValue(json); else sh.appendRow([key, json]);
}

/* ---------------- 권한 ---------------- */
function isAdmin_(auth) { const pw = adminPw_(); return !!pw && String(auth.admin || '') === pw; }
function codes_() { return cfgAll_().codes || {}; }
function canAccess_(auth, team, codes) {
  if (isAdmin_(auth)) return true;
  if ((auth.mgr || {})[team]) return true;   // 임원 (QR 기기 또는 개인 비밀번호)
  codes = codes || codes_();
  if (!codes[team]) return false; // 임원 번호가 없는 부서는 총관리자만 (담임 번호로는 자기 반만)
  return String((auth.codes || {})[team] || '') === String(codes[team]);
}
/** 이 요청이 부서 안에서 어디까지 볼 수 있는지: dept(부서 전체) / grade(담임 반만) / none */
function scope_(auth, team, codes, gcodes) {
  if (!TEAMS[team]) return { level: 'none', grades: [] };
  codes = codes || codes_();
  if (canAccess_(auth, team, codes)) return { level: 'dept', grades: groupsOf_(team).map(x => x[0]).concat(['T']) };
  const grades = hrGrades_(auth, team, gcodes);
  if (!grades.length) {
    const me = (auth.hrMe || {})[team]; if (me && me.own) return { level: 'sub', grades: [], name: String(me.name || '') };   // 담임 반이 없는 선생님의 개인 비밀번호
  }
  return { level: grades.length ? 'grade' : 'none', grades };
}
/** 보조교사 비밀번호 (부서마다 하나). 교사 명단에 이름이 있는 선생님이 이름과 함께 넣으면 그 부서 일정 탭만 볼 수 있다 */
function subCodes_() { return cfgAll_().subCodes || {}; }
/** 로그인 이름 비교 (띄어쓰기 무시). 명단 이름이 '정나리(정미유)' 처럼 괄호가 있으면 괄호 앞·안 이름만 적어도 맞다 */
function nameMatch_(stored, typed) {
  const t = String(typed || '').replace(/\s+/g, ''), s = String(stored || '').replace(/\s+/g, ''); if (!t) return false;
  if (s === t) return true;
  const m = s.match(/^(.+?)[(（](.+?)[)）]$/); return !!m && (m[1] === t || m[2] === t);
}
/** 두 이름이 같은 사람인지 (어느 쪽이 괄호 이름이어도) */
const sameName_ = (a, b) => nameMatch_(a, b) || nameMatch_(b, a);
/** homeroomNames_ 결과에 이 선생님 이름이 있는지 (반 설정에 '정미유', 명단에 '정나리(정미유)' 여도 맞다) */
const hrHas_ = (names, name) => !!names[name] || Object.keys(names).some(n => sameName_(n, name));
function subTeacher_(R, team, name) {
  const nm = String(name || '').replace(/\s+/g, ''); if (!nm) return null;
  const L = R.list.filter(p => p.team === team && p.group === 'T' && p.active !== false);
  return L.find(p => String(p.name || '').replace(/\s+/g, '') === nm) || L.find(p => nameMatch_(p.name, nm)) || null;
}
/** 담임으로 열 수 있는 반: 반 담임 번호가 맞는 반 + 개인 비밀번호로 들어온 선생님이 지금 담임으로 배치된 반 */
function hrGrades_(auth, team, gcodes) {
  gcodes = gcodes || gradeCodes_();
  const mine = auth.gcodes || {}, me = (auth.hrMe || {})[team];
  const cls = me && me.own ? readClasses_() : null;
  return groupsOf_(team).map(x => x[0]).filter(g => gcEq_(gcodes[team + '_' + g], mine[team + '_' + g]) || (cls && hrHas_(homeroomNames_(team, [g], cls), me.name)));
}
/** 지금 로그인한 담임 선생님 본인 {pid, name, own}. 개인 비밀번호면 그대로, 이름만 고른 경우는 그 이름이 반 담임 번호로 연 반의 담임일 때만 */
function meOf_(auth, team, cls) {
  const m = (auth.hrMe || {})[team]; if (!m || !m.name) return null;
  if (m.own) return m;
  const g = gradeCodes_(), mine = auth.gcodes || {}; cls = cls || readClasses_();
  return groupsOf_(team).some(([k]) => gcEq_(g[team + '_' + k], mine[team + '_' + k]) && hrHas_(homeroomNames_(team, [k], cls), m.name)) ? m : null;
}
/** 한 사람(학생)을 다룰 권한. 담임은 자기 반 학생만. */
function requirePerson_(auth, team, group, codes, gcodes) {
  const sc = scope_(auth, team, codes, gcodes);
  if (sc.level === 'dept') return sc;
  if (sc.level === 'grade' && sc.grades.indexOf(group) >= 0) return sc;
  if (sc.level === 'grade') throw new Error('담임 번호로는 자기 반만 관리할 수 있어요');
  throw new Error(TEAMS[team] ? TEAMS[team].label + ' 번호가 필요해요. 메인 화면에서 다시 들어가 주세요.' : '부서를 알 수 없어요');
}
/** 담임 번호로 들어온 사람이 출결을 체크할 수 있는 그 반 담임 선생님 이름들 */
function homeroomNames_(team, grades, classes) {
  const names = {};
  grades.forEach(g => {
    const c = classes[team + '_' + g]; if (!c) return;
    [c.teacher].concat(Object.keys(c.teachers || {}).map(k => c.teachers[k])).forEach(v =>
      String(v || '').split(/\s*[,，]\s*/).forEach(n => { n = n.trim(); if (n) names[n] = true; }));
  });
  return names;
}
/** 출결·메모를 적을 권한. 담임 번호로는 자기 반 학생과 그 반 담임 선생님까지. */
function requireAttPerson_(auth, p, codes, gcodes, ctx) {
  if (p.group === 'T') {
    const sc = scope_(auth, p.team, codes, gcodes);
    if (sc.level === 'grade' && hrHas_(homeroomNames_(p.team, sc.grades, ctx.classes || (ctx.classes = readClasses_())), p.name)) return sc;
  }
  return requirePerson_(auth, p.team, p.group, codes, gcodes);
}
/** 연말 등반: 총관리자 또는 임원 번호를 가진 임원만. 번호가 없는 부서는 총관리자만. */
function requirePromoter_(auth, team, codes) {
  if (!TEAMS[team]) throw new Error('부서를 알 수 없어요');
  if (!deptLevel_(auth, team, codes)) throw new Error('임원 이상만 할 수 있어요');
}
function requireDept_(auth, team, codes) {
  if (!TEAMS[team]) throw new Error('부서를 알 수 없어요');
  if (!canAccess_(auth, team, codes)) throw new Error('임원만 할 수 있어요');
}
function requireAccess_(auth, team, codes) {
  if (!TEAMS[team]) throw new Error('부서를 알 수 없어요');
  if (!canAccess_(auth, team, codes)) throw new Error(TEAMS[team].label + ' 고유번호가 필요해요. 메인 화면에서 다시 들어가 주세요.');
}
function adminLogin_(pw, email) {
  if (guardStop_('admin', '', '총관리자 로그인 (입력 이메일: ' + mailTag_(email) + ', 로그인 창)')) throw new Error(GUARD_MSG);
  const real = adminPw_(); const ok = !!real && String(pw || '') === real;
  if (ok) migrateRootStaffRole_(email);
  const need = staffList_().length > 0;
  if (ok && need && staffRole_(email) !== 'admin') {
    guardFail_('admin', String(pw) + '|' + normEmail_(email), '', '총관리자 로그인 이메일 불일치 (입력 이메일: ' + mailTag_(email) + ', 로그인 창)');
    throw new Error('총관리자로 등록된 이메일을 함께 입력해 주세요.');
  }
  if (ok) { guardOk_('admin'); logLogin_({ admin: String(pw) }, '', '총관리자로 로그인'); }
  else if (real) guardFail_('admin', pw, '', '총관리자 로그인 (입력 이메일: ' + mailTag_(email) + ', 로그인 창)');
  return { ok, configured: !!real, emailRequired: need };
}
/** 로그인·입장 기록. 기록에 실패해도 로그인은 막지 않는다. */
/** 로그인 기록. 화면이 알려 준 기기(휴대폰·PC 등)를 끝에 붙인다 (로그인 실패 기록 logFail_ 도 같음) */
let REQ_DEV_ = '';
function devTag_() { return REQ_DEV_ ? ' (' + REQ_DEV_ + ')' : ''; }
function logLogin_(auth, team, text) { try { withLock_(() => logRows_(auth, [['', '', team, '로그인', text + devTag_()]])); } catch (e) {} }
/** 이미 쓰고 있는 번호인지 확인. 겹치면 어느 번호와 겹치는지 알려 준다. skip: 지금 바꾸는 자리
 *  sameTeam: 담임 번호를 정할 때 그 부서. 같은 부서의 다른 반 담임 번호와는 겹쳐도 된다 (한 담임이 여러 반을 맡을 때) */
function codeOwner_(code, skip, sameTeam) {
  code = String(code || '');
  const codes = codes_(), g = gradeCodes_();
  for (const t of TEAM_ORDER) if (('team:' + t) !== skip && String(codes[t] || '') === code) return TEAMS[t].label + ' 임원 번호';
  for (const t of TEAM_ORDER) if (t !== sameTeam) for (const [k, gl] of groupsOf_(t)) if (('grade:' + t + '_' + k) !== skip && gcEq_(g[t + '_' + k], code)) return TEAMS[t].label + ' ' + gl + ' 담임 번호';
  const tc = cfgAll_().tcodes || {};
  for (const pid of Object.keys(tc)) if (('tch:' + pid) !== skip && tcCheck_(tc[pid], code)) return '다른 선생님의 개인 비밀번호';
  const sc = subCodes_();
  for (const t of TEAM_ORDER) if (('sub:' + t) !== skip && gcEq_(sc[t], code)) return TEAMS[t].label + ' 보조교사 비밀번호';
  return '';
}


/* ---------------- 이름표 변환 ---------------- */
const TEAM_BY_LABEL = Object.keys(TEAMS).reduce((m, k) => (m[TEAMS[k].label] = k, m), {});
function groupKey_(team, label) {
  label = String(label).trim();
  if (label === '교사') return 'T';
  const g = groupsOf_(team).find(x => x[1] === label) || (isFlex_(team) ? null : TEAMS[team] && TEAMS[team].groups.find(x => x[1] === label));
  if (g) return g[0];
  // 영아: 지금 반 목록에 없는 이름(예전 '미배정', 없앤 반, 바꾸기 전 이름)은 첫 번째 반으로
  if (isFlex_(team)) { const cls = groupsOf_(team); return cls.some(x => x[0] === label) ? label : cls[0][0]; }
  return label;
}
function groupLabel_(team, key) {
  if (key === 'T') return '교사';
  const g = groupsOf_(team).find(x => x[0] === key);
  return g ? g[1] : key;
}
function fmtBirthday_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'MM/dd');
  return String(v || '').trim();
}
function dateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  const s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

/* ---------------- 명단 ---------------- */
/** lite=true 이면 사진 칸(L)을 읽지 않아요 (사진이 커서 느려져요). lite 로 읽은 명단은 시트에 다시 쓰면 안 돼요. */
function readRoster_(lite) {
  const sh = ss_().getSheetByName(SHEET.roster);
  const n = sh.getLastRow();
  const list = [];
  if (n < 2) return { sh, list, byId: {} };
  let vals;
  if (lite) {
    const a = sh.getRange(2, 1, n - 1, 11).getValues(), x = sh.getRange(2, 13, n - 1, 1).getValues();
    vals = a.map((r, i) => r.concat(['', x[i][0]]));
  } else vals = sh.getRange(2, 1, n - 1, RCOLS.length).getValues();
  vals.forEach((r, i) => {
    const id = String(r[0] || '').trim(); if (!id) return;
    const team = TEAM_BY_LABEL[String(r[1]).trim()] || 'youth';
    let extra = {}; try { extra = r[12] ? JSON.parse(r[12]) : {}; } catch (e) {}
    list.push({
      row: i + 2, id, team, group: groupKey_(team, r[2]), section: String(r[3] || '').trim(), name: String(r[4] || ''),
      birthday: fmtBirthday_(r[6]), role: String(r[7] || ''), isNew: String(r[8]).trim() === 'Y',
      active: ['N', '제적', '졸업'].indexOf(String(r[9]).trim()) < 0, leftType: String(r[9]).trim() === '졸업' ? '졸업' : (['N', '제적'].indexOf(String(r[9]).trim()) >= 0 ? '제적' : ''),
      note: String(r[10] || ''), photo: String(r[11] || ''),
      gender: extra.gender === '남' || extra.gender === '여' ? extra.gender : '', lunar: extra.lunar === true, photoV: Number(extra.pv) || 0, order: extra.order, promotedYear: extra.promotedYear == null ? null : extra.promotedYear, talentAdj: extra.talentAdj || {},
    });
    const p = list[list.length - 1];
    // 나이: 따로 적힌 값, 없으면 나이반(0세~4세)에 있던 아이는 그 나이
    p.age = p.group === 'T' ? '' : (ageKey_(team, extra.age) || (TEAMS[team].groups.some(x => x[0] === p.group) ? ageKey_(team, p.group) : ''));
  });
  const byId = {}; list.forEach(p => byId[p.id] = p);
  return { sh, list, byId };
}
function personRow_(p) {
  const photo = String(p.photo || '');
  if (photo.length > 49000) throw new Error('사진이 너무 커요. 다른 사진을 골라 주세요.');
  const extra = { order: p.order == null ? 999 : p.order, promotedYear: p.promotedYear == null ? null : p.promotedYear, talentAdj: p.talentAdj || {} };
  if (p.age) extra.age = p.age;
  if (p.gender) extra.gender = p.gender;
  if (p.lunar) extra.lunar = true;
  if (p.photoV) extra.pv = p.photoV;
  return [p.id, TEAMS[p.team] ? TEAMS[p.team].label : p.team, safeCell_(groupLabel_(p.team, p.group)), safeCell_(p.section || ''), safeCell_(p.name || ''),
    p.group === 'T' ? '교사' : '학생', p.birthday || '', safeCell_(p.role || ''), p.isNew ? 'Y' : '', p.active === false ? (p.leftType === '졸업' ? '졸업' : '제적') : 'Y',
    safeCell_(p.note || ''), photo, JSON.stringify(extra)];
}
function cleanPerson_(id, body) {
  const team = TEAMS[body.team] ? body.team : 'youth';
  return {
    id, team, group: String(body.group || ''), section: String(body.section || ''), name: String(body.name || '').trim(),
    birthday: String(body.birthday || ''), role: String(body.role || ''), isNew: !!body.isNew, active: body.active !== false,
    leftType: body.active === false ? (body.leftType === '졸업' ? '졸업' : '제적') : '',
    gender: body.gender === '남' || body.gender === '여' ? body.gender : '',
    lunar: body.lunar === true && !!String(body.birthday || '').trim(),
    photoV: Number(body.photoV) || 0,
    note: String(body.note || ''), photo: String(body.photo || ''), order: body.order, promotedYear: body.promotedYear, talentAdj: body.talentAdj || {},
    age: String(body.group || '') === 'T' ? '' : ageKey_(team, body.age),
  };
}
/** 명단에서 교사 이름을 고치면 반 설정의 담임 칸(분반 담임 포함, 여러 명이면 그 이름만)도 같이 고친다.
 *  같은 부서에 옛 이름을 쓰는 다른 교사가 남아 있으면 누구인지 알 수 없어 그대로 둔다. */
function renameHomerooms_(R, renamed) {
  if (!renamed.length) return;
  { const tc = cfgAll_().tcodes || {}; let ch = false; const n = Object.assign({}, tc);
    renamed.forEach(p => { if (n[p.id]) { n[p.id] = Object.assign({}, n[p.id], { n: p.name }); ch = true; } });
    if (ch) cfgSet_('tcodes', n); } // 개인 비밀번호에 적어 둔 이름도 같이
  const ids = {}; renamed.forEach(p => ids[p.id] = 1);
  const todo = renamed.map(p => [p.team, R.byId[p.id].name, p.name])
    .filter(([t, a]) => !R.list.some(x => !ids[x.id] && x.team === t && x.group === 'T' && x.active !== false && x.name === a));
  if (!todo.length) return;
  const map = readClasses_(); let changed = false;
  const fix = (team, v) => { const list = String(v || '').split(/\s*[,，]\s*/).map(n => n.trim()).filter(Boolean);
    const out = list.map(n => { const hit = todo.find(([t, a]) => t === team && a === n); return hit ? hit[2] : n; });
    if (out.join() === list.join()) return v; changed = true; return out.join(', '); };
  Object.keys(map).forEach(k => { const c = map[k];
    c.teacher = fix(c.team, c.teacher);
    Object.keys(c.teachers || {}).forEach(s => c.teachers[s] = fix(c.team, c.teachers[s])); });
  if (changed) { writeClasses_(map); todo.forEach(([t]) => touchTeam_(t)); }
}
function savePersonRows_(R, people) {
  // people: cleaned person objects; 새 사람은 맨 아래에 추가
  people.forEach(p => { touchTeam_(p.team); const o = R.byId[p.id]; if (o) touchTeam_(o.team); });
  renameHomerooms_(R, people.filter(p => { const o = R.byId[p.id]; return o && o.group === 'T' && p.group === 'T' && o.team === p.team && o.name && p.name && o.name !== p.name; }));
  const upd = people.filter(p => R.byId[p.id]), add = people.filter(p => !R.byId[p.id]);
  if (upd.length <= 5) upd.forEach(p => R.sh.getRange(R.byId[p.id].row, 1, 1, RCOLS.length).setValues([personRow_(p)]));
  else {
    // 여러 명(엑셀·등반)이면 한꺼번에 쓴다. 사진 칸(L)은 글자가 커서 건드리지 않고, 사진이 바뀐 사람만 따로 쓴다
    const rows = upd.map(p => [R.byId[p.id], personRow_(p)]);
    writeRows_(R.sh, 1, 11, rows.map(([o, v]) => [o.row, v.slice(0, 11)]));
    writeRows_(R.sh, 13, 1, rows.map(([o, v]) => [o.row, [v[12]]]));
    rows.forEach(([o, v]) => { if (String(o.photo || '') !== String(v[11] || '')) R.sh.getRange(o.row, 12).setValue(v[11]); });
  }
  if (add.length) {
    const start = R.sh.getLastRow() + 1;
    R.sh.getRange(start, 1, add.length, RCOLS.length).setValues(add.map(personRow_));
    add.forEach((p, i) => { R.byId[p.id] = Object.assign({ row: start + i }, p); });
  }
  syncWeeklyInfo_(people);
  try { CacheService.getScriptCache().remove('qr-tch2'); } catch (e) {}
}
/** 여러 줄의 몇 칸을 고칠 때: 적으면 한 줄씩, 많으면 그 범위를 한 번에 읽고 고쳐서 한 번에 쓴다. edits: [[줄, [값...]], ...] */
function writeRows_(sh, col, width, edits) {
  if (!edits.length) return;
  if (edits.length <= 5) { edits.forEach(([r, v]) => sh.getRange(r, col, 1, width).setValues([v])); return; }
  const at = edits.map(e => e[0]); const lo = Math.min.apply(null, at), hi = Math.max.apply(null, at);
  const rng = sh.getRange(lo, col, hi - lo + 1, width);
  const cur = rng.getValues().map(r => r.map(safeCell_)); // 다시 쓸 때 '=' 등으로 시작하는 글자가 수식이 되지 않게
  edits.forEach(([r, v]) => { cur[r - lo] = v; });
  rng.setValues(cur);
}
function syncBookInfo_(people) {
  const sh = ss_().getSheetByName(SHEET.book); if (!sh) return;
  const B = readBook_();
  writeRows_(sh, 2, 3, people.filter(p => B.rows[p.id]).map(p => [B.rows[p.id], [safeCell_(p.name), TEAMS[p.team].label, safeCell_(groupLabel_(p.team, p.group))]]));
}
function syncWeeklyInfo_(people) {
  syncBookInfo_(people);
  [SHEET.att, SHEET.tal].forEach(name => {
    const W = readWeeklyIdx_(name);
    writeRows_(W.sh, 2, 4, people.filter(p => W.rows[p.id]).map(p => [W.rows[p.id], [TEAMS[p.team].label, safeCell_(groupLabel_(p.team, p.group)), safeCell_(p.section || ''), safeCell_(p.name)]]));
  });
}
function deletePerson_(id) {
  const R = readRoster_(true); const old = R.byId[id];
  if (old) R.sh.deleteRow(old.row);
  [SHEET.att, SHEET.tal].forEach(name => { const W = readWeeklyIdx_(name); if (W.rows[id]) W.sh.deleteRow(W.rows[id]); });
  const B = readBook_(); if (B.rows[id]) B.sh.deleteRow(B.rows[id]);
  try { CacheService.getScriptCache().remove('qr-tch2'); } catch (e) {}
}

/* ---------------- 출석 / 달란트 (주일 칸) ---------------- */
/** 머리글(날짜 칸)과 ID 칸만 읽는다 (쓰기 쪽에서 쓰는 가벼운 버전) */
function readWeeklyIdx_(name) {
  const sh = ss_().getSheetByName(name);
  const lastRow = Math.max(sh.getLastRow(), 1), lastCol = Math.max(sh.getLastColumn(), 6);
  const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const cols = {}; for (let c = WEEK_COL - 1; c < lastCol; c++) { const k = dateKey_(hdr[c]); if (k) cols[k] = c + 1; }
  const rows = {};
  if (lastRow >= 2) sh.getRange(2, 1, lastRow - 1, 1).getValues().forEach((r, i) => { const id = String(r[0] || '').trim(); if (id) rows[id] = i + 2; });
  return { sh, cols, rows, lastCol, lastRow };
}
/** 화면에 보낼 값까지 읽는다. yes(날짜키)가 true 인 주일 칸(보통 올해)만 읽는다. W.val(줄, 칸) 으로 꺼낸다. */
function readWeekly_(name, yes) {
  const W = readWeeklyIdx_(name);
  W.use = Object.keys(W.cols).filter(k => !yes || yes(k));
  let minc = 0, maxc = 0;
  W.use.forEach(k => { const c = W.cols[k]; if (!minc || c < minc) minc = c; if (c > maxc) maxc = c; });
  W.minc = minc; W.block = [];
  if (W.use.length && W.lastRow >= 2) W.block = W.sh.getRange(2, minc, W.lastRow - 1, maxc - minc + 1).getValues();
  W.val = (row, col) => { const r = W.block[row - 2]; return r ? r[col - W.minc] : ''; };
  return W;
}
function weeklyRow_(W, p) {
  if (W.rows[p.id]) return W.rows[p.id];
  const row = W.sh.getLastRow() + 1;
  W.sh.getRange(row, 1, 1, 6).setValues([[p.id, TEAMS[p.team].label, safeCell_(groupLabel_(p.team, p.group)), safeCell_(p.section || ''), safeCell_(p.name), '=SUM(G' + row + ':' + row + ')']]);
  fmt_(W.sh.getRange(row, 1), '@');
  W.rows[p.id] = row; return row;
}
function weeklyCol_(W, key) {
  if (W.cols[key]) return W.cols[key];
  const col = Math.max(W.sh.getLastColumn() + 1, WEEK_COL);
  W.sh.getRange(1, col).setValue(key); fmt_(W.sh.getRange(1, col), 'm/d'); W.sh.getRange(1, col).setFontWeight('bold');
  W.cols[key] = col; return col;
}
/** 출석 체크·해제 기록: 바로 쓰지 않고 그날 체크·해제한 것을 (주일·부서·반·체크한 사람)마다 모아 두었다가
 *  ATT_LOG_HOUR시에 반마다 한 줄씩 한 번에 활동 기록에 쓴다 (flushAttLog_. tgPoll 이 5분마다 확인). 텔레그램 알림은 보내지 않는다 */
const ATT_LOG_HOUR = 18;
const ATT_MODES_ = ['출석 체크', '출석 해제', '달란트 해제'];
/** mode: ATT_MODES_ 중 하나 */
function logAttChecks_(auth, key, ps, mode) {
  let cls = {}; try { cls = readClasses_(); } catch (e) {}
  // 스크립트 속성에 (주일|부서|반|체크한 사람)마다 이름 목록을 하나씩 (설정 시트 캐시를 건드리지 않게)
  const P = PropertiesService.getScriptProperties(), by = {}, actors = {};
  ps.forEach(p => { const actor = actors[p.team] || (actors[p.team] = logActor_(auth, p.team, cls)); const k = 'ATTP|' + [key, p.team, p.group, mode || '출석 체크', actor].join('|'); (by[k] || (by[k] = [])).push(String(p.name || '')); });
  Object.keys(by).forEach(k => {
    let L = []; try { L = JSON.parse(P.getProperty(k) || '[]'); } catch (e) {}
    by[k].forEach(nm => { if (nm && L.indexOf(nm) < 0) L.push(nm); });
    P.setProperty(k, JSON.stringify(L).slice(0, 8000));
  });
}
/** 모아 둔 출석 체크를 활동 기록에 쓴다. force 가 아니면 ATT_LOG_HOUR시가 지나고 그날 처음일 때만 */
function flushAttLog_(force) {
  const now = new Date(), tz = tz_(), day = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const p = PropertiesService.getScriptProperties();
  if (!force && (Number(Utilities.formatDate(now, tz, 'H')) < ATT_LOG_HOUR || p.getProperty('ATT_LOG_DAY') === day)) return;
  withLock_(() => {
    if (!force) p.setProperty('ATT_LOG_DAY', day);
    const all = p.getProperties(), pend = {}; Object.keys(all).forEach(k => { if (k.indexOf('ATTP|') === 0) { try { pend[k.slice(5)] = JSON.parse(all[k]); } catch (e) { pend[k.slice(5)] = []; } } });
    const ks = Object.keys(pend); if (!ks.length) return;
    const sh = ss_().getSheetByName(SHEET.hist); if (!sh) return;
    const at = Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm');
    const parts = k => { const x = k.split('|'), m = ATT_MODES_.indexOf(x[3]) >= 0; return { d: x[0], t: x[1], g: x[2], mode: m ? x[3] : '출석 체크', by: x.slice(m ? 4 : 3).join('|') }; };
    const gi = (t, g) => groupsOf_(t).map(y => y[0]).concat(['T']).indexOf(g);
    const rows = ks.map(k => Object.assign(parts(k), { names: pend[k] || [] })).filter(x => TEAMS[x.t] && x.names.length)
      .sort((x, y) => x.d.localeCompare(y.d) || TEAM_ORDER.indexOf(x.t) - TEAM_ORDER.indexOf(y.t) || gi(x.t, x.g) - gi(y.t, y.g) || ATT_MODES_.indexOf(x.mode) - ATT_MODES_.indexOf(y.mode))
      .map(x => [at, '', '', TEAMS[x.t].label, '출석', safeCell_(x.d.slice(5).replace('-', '/') + ' ' + groupLabel_(x.t, x.g) + ' ' + x.mode + ' ' + x.names.length + '명 (' + x.names.join(', ') + ')'), x.by, x.mode]);
    if (rows.length) { const r = sh.getLastRow() + 1; fmt_(sh.getRange(r, 1, rows.length, 2), '@'); sh.getRange(r, 1, rows.length, 7).setValues(rows.map(v => v.slice(0, 7))); }
    ks.forEach(k => p.deleteProperty('ATTP|' + k));
    // 텔레그램으로도 한 번에 (알림 종류 '출석체크'를 꺼 두면 안 보냄)
    const tgRows = rows.filter(v => tgOn_(v[7] === '출석 체크' ? '출석체크' : '출석'));
    if (tgRows.length) try {
      const L = tgRows.map(v => '• [' + v[3] + '] ' + v[5] + (v[6] ? ' — ' + v[6] : ''));
      tgNotify_('📋 출석 기록 ' + Utilities.formatDate(now, tz, 'M/d HH:mm') + '\n' + L.slice(0, 30).map(x => x.slice(0, 300)).join('\n') + (L.length > 30 ? '\n… 외 ' + (L.length - 30) + '건 (활동 기록에서 확인)' : ''));
    } catch (e) {}
  });
}
/** 편집기에서 실행: 모아 둔 출석 체크를 지금 바로 활동 기록에 쓴다 */
function flushAttLogNow() { flushAttLog_(true); }
function attUpdate_(auth, key, body) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(key)) throw new Error('날짜 형식이 맞지 않아요');
  const R = readRoster_(true); const codes = codes_(); const gcodes = gradeCodes_(); const ctx = {};
  const apply = (name, map, toCell) => {
    const ids = Object.keys(map || {}); if (!ids.length) return;
    // 권한은 모두 먼저 확인해요 (번호 없는 요청이 빈 열이나 빈 줄을 만들지 않게)
    const todo = [];
    ids.forEach(id => { const p = R.byId[id]; if (!p) return; requireAttPerson_(auth, p, codes, gcodes, ctx); todo.push(p); touchGroup_(p.team, p.group); });
    if (!todo.length) return;
    if (name === SHEET.att) todo.forEach(p => { if (toCell(map[p.id]) !== '') checked.push(p); });   // 출석 체크 기록용
    const W = readWeeklyIdx_(name);
    // 시트에 줄이 없는 사람은 한꺼번에 만든다
    const missing = todo.filter(p => !W.rows[p.id]);
    if (missing.length) {
      const start = W.sh.getLastRow() + 1;
      W.sh.getRange(start, 1, missing.length, 6).setValues(missing.map((p, i) => [p.id, TEAMS[p.team].label, safeCell_(groupLabel_(p.team, p.group)), safeCell_(p.section || ''), safeCell_(p.name), '=SUM(G' + (start + i) + ':' + (start + i) + ')']));
      fmt_(W.sh.getRange(start, 1, missing.length, 1), '@');
      missing.forEach((p, i) => { W.rows[p.id] = start + i; });
    }
    const col = weeklyCol_(W, key);
    const clears = todo.some(p => toCell(map[p.id]) === '');
    if (todo.length <= 3 && !clears) { todo.forEach(p => W.sh.getRange(W.rows[p.id], col).setValue(toCell(map[p.id]))); return; }
    // 여러 명이거나 지우는 칸이 있으면 그날 칸을 한 번에 읽어서 고치고 한 번에 쓴다 (실제로 지워진 사람을 알 수 있게)
    const last = W.sh.getLastRow();
    const rng = W.sh.getRange(2, col, last - 1, 1); const cur = rng.getValues();
    todo.forEach(p => { const at = W.rows[p.id] - 2; if (!cur[at]) return; const v = toCell(map[p.id]); if (v === '' && cur[at][0] !== '' && cur[at][0] !== null) cleared.push([name, p]); cur[at][0] = v; });
    rng.setValues(cur);
  };
  const cleared = [], checked = [];
  apply(SHEET.att, body.p, v => (v ? 1 : ''));
  apply(SHEET.tal, body.t, v => (Number(v) > 0 ? Number(v) : ''));
  if (checked.length) try { logAttChecks_(auth, key, checked, '출석 체크'); } catch (e) {}
  // 지운 출석·달란트도 바로 쓰지 않고 모았다가 ATT_LOG_HOUR시에 반마다 한 줄로
  if (cleared.length) try {
    const by = {}; cleared.forEach(([name, p]) => (by[name] || (by[name] = [])).push(p));
    Object.keys(by).forEach(name => logAttChecks_(auth, key, by[name], name + ' 해제'));
  } catch (e) {}
  const mids = Object.keys(body.m || {});
  if (mids.length) {
    const sh = ss_().getSheetByName(SHEET.memo); const n = sh.getLastRow();
    const rows = n >= 2 ? sh.getRange(2, 1, n - 1, 2).getValues() : [];
    mids.forEach(id => {
      const p = R.byId[id]; if (!p) return;
      requireAttPerson_(auth, p, codes, gcodes, ctx); touchGroup_(p.team, p.group);
      const text = String(body.m[id] || '').trim().slice(0, 300);
      let at = -1; rows.forEach((r, i) => { if (dateKey_(r[0]) === key && String(r[1]) === id) at = i; });
      if (!text) { if (at >= 0) { sh.deleteRow(at + 2); rows.splice(at, 1); } return; }
      const vals = [[key, id, safeCell_(p.name), TEAMS[p.team].label, safeCell_(groupLabel_(p.team, p.group)), safeCell_(text)]];
      if (at >= 0) sh.getRange(at + 2, 1, 1, 6).setValues(vals);
      else { const r = sh.getLastRow() + 1; fmt_(sh.getRange(r, 1, 1, 2), '@'); sh.getRange(r, 1, 1, 6).setValues(vals); rows.push([key, id]); }
    });
  }
  return true;
}

/* ---------------- 반 설정 ---------------- */
// 한 요청 안에서 반 설정을 여러 번 읽지 않도록 기억해 둔다 (writeClasses_ 가 지운다). 고쳐 쓰는 쪽이 있어 늘 복사본을 돌려준다
let CLS_ = null;
function readClasses_() {
  if (!CLS_) CLS_ = readClassesRaw_();
  return JSON.parse(JSON.stringify(CLS_));
}
function readClassesRaw_() {
  const sh = ss_().getSheetByName(SHEET.cls); const n = sh.getLastRow();
  const map = {};
  if (n < 2) return map;
  sh.getRange(2, 1, n - 1, 4).getValues().forEach(r => {
    const team = TEAM_BY_LABEL[String(r[0]).trim()]; if (!team) return;
    const g = groupKey_(team, r[1]); const key = team + '_' + g; const sec = String(r[2] || '').trim(); const t = String(r[3] || '').trim();
    const c = map[key] || (map[key] = { team, group: g, teacher: '', sections: [], teachers: {} });
    if (sec) { if (c.sections.indexOf(sec) < 0) c.sections.push(sec); c.teachers[sec] = t; } else c.teacher = t;
  });
  return map;
}
function writeClasses_(map) {
  CLS_ = null;
  const sh = ss_().getSheetByName(SHEET.cls);
  const n = sh.getLastRow(); if (n >= 2) sh.getRange(2, 1, n - 1, Math.max(sh.getLastColumn(), 5)).clearContent();
  const rows = [];
  TEAM_ORDER.forEach(team => groupsOf_(team).forEach(([g, gl]) => {
    const c = map[team + '_' + g] || { teacher: '', sections: [], teachers: {} };
    const push = (sec, teacher) => {
      const r = rows.length + 2;
      rows.push([TEAMS[team].label, safeCell_(gl), sec, safeCell_(teacher || ''), sec
        ? '=COUNTIFS(명단!$B:$B,$A' + r + ',명단!$C:$C,$B' + r + ',명단!$D:$D,$C' + r + ',명단!$J:$J,"Y")'
        : '=COUNTIFS(명단!$B:$B,$A' + r + ',명단!$C:$C,$B' + r + ',명단!$J:$J,"Y")']);
    };
    push('', c.teacher);
    (c.sections || []).forEach(sec => push(sec, (c.teachers || {})[sec]));
  }));
  if (rows.length) sh.getRange(2, 1, rows.length, 5).setValues(rows);
}

/* ---------------- 읽기 ---------------- */
/** 화면을 처음 열 때(그리고 바뀐 것이 있을 때) 한 번에 보내는 데이터. years 가 없으면 올해(1~2월엔 작년도)의 출석만 보낸다. */
function defaultYears_() {
  const now = new Date(); const y = Number(Utilities.formatDate(now, tz_(), 'yyyy')), m = Number(Utilities.formatDate(now, tz_(), 'M'));
  return m <= 2 ? [y - 1, y] : [y];
}
/** 이 화면이 지켜볼 버전 이름: 전체 + 볼 수 있는 부서 + 그 부서에서 볼 수 있는 반(교사 줄 T 포함) */
function watchKeys_(scope) {
  const keys = ['g'];
  TEAM_ORDER.forEach(t => {
    const sc = scope[t]; if (!sc || sc.level === 'none') return;
    keys.push(t);
    (sc.level === 'dept' ? groupsOf_(t).map(x => x[0]) : sc.grades).concat(['T']).forEach(g => { if (keys.indexOf(t + '_' + g) < 0) keys.push(t + '_' + g); });
  });
  return keys;
}
/** 한시(12월 ~ 아래 부서 연말 등반을 마칠 때까지): 바로 아래 부서의 맨 위 학년·나이(영아 4세, 유아 7세, 유소년 6학년) 중
 *  등반 전인 학생을 윗부서에 예비 신입생으로 보여준다. 등반 연도는 화면과 같이 1~6월이면 지난해 연말 등반으로 본다.
 *  부서 전체를 보는 사람과 첫 학년(5세·1학년·중1) 담임만. 연락처·메모는 빼고 보낸다. 기간이 아니면 null */
function preNew_(R, scope, cfg) {
  const now = new Date(), m = Number(Utilities.formatDate(now, tz_(), 'M')); if (m !== 12 && m > 6) return null;
  const y = Number(Utilities.formatDate(now, tz_(), 'yyyy')) - (m >= 7 ? 0 : 1); const out = {};
  TEAM_ORDER.forEach((t, i) => {
    const from = TEAM_ORDER[i - 1], sc = scope[t]; if (!from || !sc) return;
    const rec = cfg['promo:' + from + '_' + y]; // 등반을 한 번이라도 실행(되돌리지 않음)했으면 끝
    if (rec && (Array.isArray(rec.runs) ? rec.runs.some(x => !x.undone) : !rec.undone)) return;
    if (sc.level !== 'dept' && !(sc.level === 'grade' && sc.grades.indexOf(TEAMS[t].groups[0][0]) >= 0)) return;
    const last = TEAMS[from].groups[TEAMS[from].groups.length - 1][0];
    const skip = (rec || {}).skip || [];
    out[t] = R.list.filter(p => p.team === from && p.active && p.group !== 'T' && (isFlex_(from) ? p.age : p.group) === last && !(Number(p.promotedYear) >= y) && skip.indexOf(p.id) < 0)
      .map(p => ({ id: p.id, name: p.name, gender: p.gender, birthday: p.birthday, lunar: p.lunar, group: p.group, section: p.section, isNew: p.isNew }));
  });
  return Object.keys(out).length ? out : null;
}
function boot_(auth, years) {
  const ver = ver_();
  ensureSheets_();
  dropTeamCodes_();
  const codes = codes_(); const admin = isAdmin_(auth);
  const access = {}, locked = {}, counts = {}, scope = {}; const gcodes = gradeCodes_();
  TEAM_ORDER.forEach(t => { locked[t] = !!codes[t]; scope[t] = scope_(auth, t, codes, gcodes); access[t] = scope[t].level !== 'none'; counts[t] = 0; });
  const vers = verMap_(watchKeys_(scope)); // 데이터를 읽기 전에 버전부터 (읽는 사이에 바뀐 것은 다음 확인 때 걸러져요)
  const R = readRoster_(true);
  R.list.forEach(p => { if (p.active && p.group !== 'T' && counts[p.team] != null) counts[p.team]++; });
  const hr = {}; let cls = null;
  TEAM_ORDER.forEach(t => { if (scope[t].level === 'grade') hr[t] = homeroomNames_(t, scope[t].grades, cls || (cls = readClasses_())); });
  const people = R.list.filter(p => access[p.team] && (scope[p.team].level === 'dept' || scope[p.team].grades.indexOf(p.group) >= 0 || (p.group === 'T' && hr[p.team] && hrHas_(hr[p.team], p.name)))).map(p => { const o = Object.assign({}, p); delete o.row; delete o.photo; return o; });
  const visible = {}; people.forEach(p => visible[p.id] = true);
  const useYears = (Array.isArray(years) && years.length ? years.map(Number).filter(y => y > 2000 && y < 2200) : defaultYears_());
  const inYears = k => useYears.indexOf(Number(String(k).slice(0, 4))) >= 0;
  const att = {}, seenYears = {};
  [[SHEET.att, 'p'], [SHEET.tal, 't']].forEach(([name, f]) => {
    const W = readWeekly_(name, inYears);
    Object.keys(W.cols).forEach(k => { seenYears[k.slice(0, 4)] = 1; });
    W.use.forEach(k => {
      const col = W.cols[k]; const a = att[k] || (att[k] = { date: k, p: {}, t: {} });
      Object.keys(W.rows).forEach(id => {
        if (!visible[id]) return; const v = W.val(W.rows[id], col);
        if (v === '' || v === null) return;
        if (f === 'p') { if (Number(v) > 0 || v === true) a.p[id] = 1; } else if (Number(v) > 0) a.t[id] = Number(v);
      });
    });
  });
  const msh = ss_().getSheetByName(SHEET.memo); const mn = msh ? msh.getLastRow() : 0;
  if (mn >= 2) msh.getRange(2, 1, mn - 1, 6).getValues().forEach(r => {
    const k = dateKey_(r[0]), id = String(r[1] || ''); if (!k || !visible[id] || !inYears(k)) return;
    const a = att[k] || (att[k] = { date: k, p: {}, t: {} }); (a.m || (a.m = {}))[id] = String(r[5] || '');
  });
  const cfg = cfgAll_(); const promos = {};
  Object.keys(cfg).forEach(k => { if (k.indexOf('promo:') === 0) promos[k.slice(6)] = cfg[k]; });
  const allCls = cls || readClasses_(); const classes = {};
  Object.keys(allCls).forEach(k => { const c = allCls[k]; if (access[c.team] && scope[c.team].grades.indexOf(c.group) >= 0) classes[k] = c; });
  Object.keys(promos).forEach(k => { const t = k.split('_')[0]; if (!scope[t] || scope[t].level !== 'dept') delete promos[k]; });
  // 담임이 직접 바꾼 담임 번호 (최근 14일, 임원 이상에게만)
  const gcChanged = {}; const since = Utilities.formatDate(new Date(Date.now() - 14 * 864e5), tz_(), 'yyyy-MM-dd HH:mm');
  if (!auth.co2) Object.keys(cfg.gcChanged || {}).forEach(k => { const t = k.split('_')[0]; if (scope[t] && scope[t].level === 'dept' && String(cfg.gcChanged[k]) >= since) gcChanged[k] = cfg.gcChanged[k]; });
  const signedStaff = staffList_().find(x => normEmail_(x.email) === normEmail_(auth.email));
  // 본인 이름으로 들어온 담임 선생님 (부서별 { pid, name, own: 개인 비밀번호 사용 })
  const hrMe = {}; TEAM_ORDER.forEach(t => { if (scope[t].level === 'grade') { const m = meOf_(auth, t, allCls); if (m) hrMe[t] = { pid: m.pid, name: m.name, own: !!m.own, dev: !!m.dev, hasPw: !!(cfg.tcodes || {})[m.pid] }; }
    else if (scope[t].level === 'sub') { const m = (auth.hrMe || {})[t]; if (m && m.own) hrMe[t] = { pid: m.pid, name: m.name, own: true, sub: true, dev: !!m.dev, hasPw: !!(cfg.tcodes || {})[m.pid] };
      else { const p = subTeacher_(R, t, scope[t].name); if (p) hrMe[t] = { pid: p.id, name: p.name, own: false, sub: true }; } } });
  const codeTemp = {}; if (!admin) TEAM_ORDER.forEach(t => { if (scope[t].level === 'dept' && codes[t] && !(cfg.codeSelf || {})[t] && !(auth.mgr || {})[t]) codeTemp[t] = true; });
  const smE = normEmail_(auth.staffDev || auth.staffPw || ''), smS = smE ? staffList_().find(z => normEmail_(z.email) === smE) : null;
  const staffMe = smS && smS.role !== 'admin' ? { name: smS.name, role: smS.role, team: smS.team || '', pos: smS.pos || '', dev: !!auth.staffDev, hasPw: !!coPwMap_()[smE] } : null;
  return { hrMe, codeTemp, staffMe, tk: auth.tk || {}, devBad: auth.devBad || [], devOn: (auth.dev || []).length, adminBad: !!auth.adminBad, isAdmin: admin, bibleViews: isRoot_(auth) ? bibleViews_() : null, isCo: !!auth.co, isCo2: !!auth.co2, me: auth.coName || (signedStaff ? String(signedStaff.name || '') : ''), adminSet: !!adminPw_(), locked, access, scope, counts, people, att, classes, promos, settings: cfg.teamSettings || {}, codes: admin && !auth.co2 ? Object.keys(codes).reduce((m, t) => (m[t] = (cfg.codeSelf || {})[t] ? '' : codes[t], m), {}) : {}, // 임원이 직접 바꾼 번호는 보내지 않아요
    codesOwn: admin && !auth.co2 ? Object.keys(codes).filter(t => (cfg.codeSelf || {})[t]) : [], staff: staffView_(auth), staffMgr: staffMgr_(auth),
    gcChanged, preNew: preNew_(R, scope, cfg), rptOk: rptViewer_(auth), v: ver, vers, years: useYears, allYears: Object.keys(seenYears).map(Number).sort() };
}
/** 사진은 따로 보낸다 (명단에 사진을 같이 실으면 열 때마다 느려져요). 화면이 모르는 사람의 사진만 부탁해요. 보이는 권한이 있는 사람만. */
function photos_(auth, ids) {
  if (!Array.isArray(ids) || !ids.length) return {};
  ids = ids.slice(0, 120).map(String);
  const R = readRoster_(true); const codes = codes_(), gcodes = gradeCodes_(); const ctx = {};
  const allowed = [];
  ids.forEach(id => {
    const p = R.byId[id]; if (!p) return;
    const sc = scope_(auth, p.team, codes, gcodes);
    if (sc.level === 'dept' || (sc.level === 'grade' && (sc.grades.indexOf(p.group) >= 0 || (p.group === 'T' && hrHas_(homeroomNames_(p.team, sc.grades, ctx.classes || (ctx.classes = readClasses_())), p.name))))) allowed.push(p);
  });
  const out = {}; if (!allowed.length) return out;
  const sh = R.sh, n = sh.getLastRow();
  if (allowed.length <= 6) allowed.forEach(p => { const v = String(sh.getRange(p.row, 12).getValue() || ''); if (v) out[p.id] = v; });
  else { const col = sh.getRange(2, 12, n - 1, 1).getValues(); allowed.forEach(p => { const v = String((col[p.row - 2] || [''])[0] || ''); if (v) out[p.id] = v; }); }
  return out;
}

/* ---------------- 쓰기 ---------------- */
function write_(auth, op, path, body) {
  const parts = String(path).split('/'); const col = parts[0], id = parts.slice(1).join('/');
  body = body || {};
  if (col === 'att') {
    if (op === 'delete') throw new Error('출석 기록 전체는 지울 수 없어요');
    return attUpdate_(auth, id, body);
  }
  if (col === 'people') {
    const R = readRoster_(op === 'delete'); const codes = codes_(); const gcodes = gradeCodes_(); const old = R.byId[id];
    if (op === 'delete') { if (!old) return true; requirePerson_(auth, old.team, old.group, codes, gcodes); touchTeam_(old.team); deletePerson_(id); logRows_(auth, [[old.id, old.name, old.team, kindOf_(old), '명단에서 삭제 (' + placeLabel_(old) + ')']]); return true; }
    let next, sc;
    if (op === 'update') {
      if (!old) throw new Error('명단에서 이 사람을 찾을 수 없어요');
      sc = requirePerson_(auth, old.team, old.group, codes, gcodes);
      if (promoRevert_(old, body)) requireUndoPromo_(auth);
      next = Object.assign({}, old, body);
      if (body.talentAdj) next.talentAdj = Object.assign({}, old.talentAdj || {}, body.talentAdj);
    } else {
      sc = old ? requirePerson_(auth, old.team, old.group, codes, gcodes) : requirePerson_(auth, body.team, String(body.group || ''), codes, gcodes);
      if (old && promoRevert_(old, body)) requireUndoPromo_(auth);
      next = body;
    }
    if (sc.level === 'grade' && (next.team !== (old ? old.team : body.team) || sc.grades.indexOf(String(next.group)) < 0))
      throw new Error('담임 번호로는 학생을 다른 반으로 옮길 수 없어요. 임원에게 요청하세요.');
    if (sc.level !== 'dept' && (old ? old.section || '' : '') !== String(next.section || ''))
      throw new Error('분반은 임원 이상만 바꿀 수 있어요. 임원에게 요청하세요.');
    if (old && next.team !== old.team) { requirePromoter_(auth, old.team, codes); requireAccess_(auth, next.team, codes); } // 다른 부서로: 원래 부서 임원 이상, 받는 부서도 들어갈 수 있어야
    const p = cleanPerson_(id, next);
    p.photoV = p.photo ? ((old && old.photo === p.photo) ? (Number(old.photoV) || 0) : Date.now()) : 0; // 사진이 바뀌면 표시를 새로
    if (!p.name) throw new Error('이름을 적어 주세요');
    savePersonRows_(R, [p]);
    logPersonChange_(auth, old, p, '');
    return true;
  }
  if (col === 'classes') {
    const team = id.split('_')[0]; const sc = requirePerson_(auth, team, id.slice(team.length + 1));
    // 담임 지정·분반은 임원 이상만
    if (sc.level !== 'dept') throw new Error('담임 지정과 분반은 임원 이상만 할 수 있어요. 임원에게 요청하세요.');
    touchTeam_(team);
    const map = readClasses_(); const prevCls = map[id] ? JSON.parse(JSON.stringify(map[id])) : null;
    if (op === 'delete') delete map[id];
    else map[id] = { team, group: id.slice(team.length + 1), teacher: String(body.teacher || ''),
      sections: (body.sections || []).filter(s => /^[A-D]$/.test(s)), teachers: body.teachers || {} };
    writeClasses_(map);
    logHomeroom_(auth, team, id.slice(team.length + 1), prevCls, map[id]);
    return true;
  }
  if (col === 'promotions') {
    const team = id.split('_')[0]; requirePromoter_(auth, team);
    CFG_ROWS_ = null; CFG_OBJ_ = null; try { CacheService.getScriptCache().remove('cfg-rows'); } catch (e) {}
    const cur = cfgAll_()['promo:' + id];
    const next = op === 'delete' ? null : op === 'update' ? Object.assign({}, cur || {}, body) : body;
    if (promoUndoing_(cur, next)) requireUndoPromo_(auth);
    if (op === 'delete') cfgSet_('promo:' + id, null);
    else cfgSet_('promo:' + id, next);
    return true;
  }
  throw new Error('저장할 수 없는 항목이에요: ' + col);
}

/**
 * 이미 등반한 학생 한 명을 등반에서 뺀다: 등반 전 자리로 되돌리고, 그해 등반 기록(runs)과 '이력' 탭의 등반 줄에서 지우고,
 * 그해 등반 제외(skip)에 넣는다. key = '부서_연도' (예: youth_2026)
 */
/** 등반은 임원 이상, 되돌리기(전체·한 명)는 총관리자·공동관리자만 */
function requireUndoPromo_(auth) { if (!isAdmin_(auth)) throw new Error('등반 되돌리기는 총관리자·공동관리자만 할 수 있어요'); }
/** 등반 기록이 바뀌면서 이미 한 등반이 지워지거나 '되돌림'이 되는지 */
function promoUndoing_(cur, next) {
  if (!cur) return false;
  const runsOf = r => (r && Array.isArray(r.runs)) ? r.runs : (r && r.moves ? [{ id: 'r0', moves: r.moves, undone: !!r.undone }] : []);
  const was = runsOf(cur).filter(r => !r.undone); if (!was.length) return false;
  if (!next) return true;
  const now = runsOf(next);
  return was.some(r => !now.some(x => x.id === r.id && !x.undone && (x.moves || []).length >= (r.moves || []).length));
}
/** 등반한 학생을 등반 전으로 되돌리는 수정인지 (등반 표시 연도가 지워지거나 앞 해로) */
function promoRevert_(old, patch) {
  if (!patch || !('promotedYear' in patch) || old.promotedYear == null || old.promotedYear === '') return false;
  return patch.promotedYear == null || patch.promotedYear === '' || Number(patch.promotedYear) < Number(old.promotedYear);
}
function promoExclude_(auth, key, id) {
  key = String(key || ''); id = String(id || '');
  const team = key.split('_')[0], year = Number(key.split('_')[1]);
  if (!TEAMS[team] || !year) throw new Error('등반 기록을 알 수 없어요');
  requireUndoPromo_(auth);
  CFG_ROWS_ = null; CFG_OBJ_ = null; try { CacheService.getScriptCache().remove('cfg-rows'); } catch (e) {}
  const rec = cfgAll_()['promo:' + key]; if (!rec) throw new Error('등반 기록이 없어요');
  const runs = rec.runs || [{ id: 'r0', at: rec.at, count: rec.count, moves: rec.moves || [], undone: !!rec.undone }];
  let mv = null;
  runs.forEach(r => { if (r.undone || mv) return; const i = (r.moves || []).findIndex(m => m.id === id); if (i < 0) return;
    mv = r.moves[i]; r.moves.splice(i, 1); r.count = r.moves.length; if (!r.count) { r.undone = true; r.undoneAt = new Date().toISOString(); } });
  if (!mv) throw new Error('이 학생의 등반 기록을 찾을 수 없어요');
  // 학생을 등반 전 자리로 (이력에 반 이동으로 남기지 않음)
  const R = readRoster_(); const cur = R.byId[id];
  if (cur) {
    const next = Object.assign({}, cur, { team: mv.team, group: mv.group, section: mv.section || '', active: mv.active !== false, leftType: mv.leftType || '',
      note: mv.note == null ? cur.note : mv.note, promotedYear: mv.promotedYear == null ? null : mv.promotedYear });
    if ('age' in mv) next.age = mv.age;
    savePersonRows_(R, [cleanPerson_(id, next)]);
  }
  const live = runs.filter(r => !r.undone);
  const skip = (Array.isArray(rec.skip) ? rec.skip : []).filter(x => x !== id).concat([id]);
  cfgSet_('promo:' + key, { team, year, at: (live[live.length - 1] || runs[runs.length - 1] || {}).at || rec.at, count: live.reduce((a, r) => a + (r.count || 0), 0), undone: !live.length, runs, skip });
  // '이력' 탭에서 이 학생의 그해 등반 줄을 지운다 (아래에서 위로)
  const sh = ss_().getSheetByName(SHEET.hist); const n = sh ? sh.getLastRow() : 0;
  if (n >= 2) {
    // 등반 줄('연말 등반(연도)')과, 같은 때 함께 남은 분반 해제 줄까지
    const tag = '연말 등반(' + year + ')'; const v = sh.getRange(2, 1, n - 1, 6).getValues();
    const at = {}; v.forEach(r => { if (String(r[1]) === id && String(r[5]).indexOf(tag) >= 0) at[String(r[0])] = 1; });
    for (let i = v.length - 1; i >= 0; i--) if (String(v[i][1]) === id && at[String(v[i][0])] && (String(v[i][5]).indexOf(tag) >= 0 || String(v[i][5]).indexOf('분반:') === 0)) sh.deleteRow(i + 2);
  }
  return true;
}

/** 여러 사람을 한 번에 고칠 때 (학년 등반, 분반). list = [[ 'people/ID', patch ], ...] */
function batch_(auth, list) {
  const R = readRoster_(); const codes = codes_(); const gcodes = gradeCodes_(); const out = [];
  const before = {}; R.list.forEach(p => before[p.id] = Object.assign({}, p));
  (list || []).forEach(([path, patch]) => {
    const id = String(path).replace(/^people\//, ''); const old = R.byId[id];
    if (!old) return;
    const sc = requirePerson_(auth, old.team, old.group, codes, gcodes);
    const moves = patch && (('team' in patch && patch.team !== old.team) || ('group' in patch && String(patch.group) !== String(old.group)) || 'promotedYear' in patch);
    if (moves) requirePromoter_(auth, old.team, codes);
    if (promoRevert_(old, patch)) requireUndoPromo_(auth);
    const next = Object.assign({}, old, patch || {});
    if (sc.level === 'grade' && (next.team !== old.team || sc.grades.indexOf(String(next.group)) < 0))
      throw new Error('담임 번호로는 학생을 다른 반으로 옮길 수 없어요. 임원에게 요청하세요.');
    if (patch && patch.talentAdj) next.talentAdj = Object.assign({}, old.talentAdj || {}, patch.talentAdj);
    out.push(cleanPerson_(id, next));
  });
  if (out.length) savePersonRows_(R, out);
  const logs = [];
  out.forEach(p => { const old = before[p.id]; logPersonChange_(auth, old, p, p.promotedYear && p.promotedYear !== (old && old.promotedYear) ? '연말 등반(' + p.promotedYear + ')' : '', logs); });
  logRows_(auth, logs);
  return out.length;
}

/* ---------------- 주소록 (권한별) ---------------- */
/**
 * 주소록은 명단과 따로 '주소록' 탭에 저장하고, 권한이 확인된 범위만 화면으로 보낸다.
 *   관리자                 : 모든 부서
 *   부서 고유번호를 가진 사람 : 그 부서 전체 (교사 포함) — 부서에 고유번호가 정해져 있어야 함
 *   학년 번호를 가진 사람    : 그 학년 학생만
 */
function gradeCodes_() { return cfgAll_().gradeCodes || {}; }
function deptLevel_(auth, team, codes) {
  if (isAdmin_(auth)) return true;
  if ((auth.mgr || {})[team]) return true;
  codes = codes || codes_();
  return !!codes[team] && String((auth.codes || {})[team] || '') === String(codes[team]);
}
function bookScope_(auth, team) {
  const codes = codes_(), g = gradeCodes_();
  const dept = deptLevel_(auth, team, codes);
  const grades = dept ? groupsOf_(team).map(x => x[0]) : hrGrades_(auth, team, g);
  if (dept) grades.push('T');
  return { dept, grades };
}
function readBook_() {
  const sh = ss_().getSheetByName(SHEET.book); const n = sh.getLastRow();
  const rows = {}, map = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, BCOLS.length).getValues().forEach((r, i) => {
    const id = String(r[0] || '').trim(); if (!id) return;
    rows[id] = i + 2; map[id] = { phone: String(r[4] || ''), parentPhone: String(r[5] || ''), guardian: String(r[6] || ''), address: String(r[7] || '') };
  });
  return { sh, rows, map };
}
function contacts_(auth, team) {
  if (!TEAMS[team]) throw new Error('부서를 알 수 없어요');
  const scope = bookScope_(auth, team); const g = gradeCodes_();
  const R = readRoster_(true); const B = readBook_(); const out = {};
  R.list.forEach(p => { if (p.team === team && scope.grades.indexOf(p.group) >= 0) out[p.id] = B.map[p.id] || { phone: '', parentPhone: '', guardian: '', address: '' }; });
  return { dept: scope.dept, grades: scope.grades, contacts: out };
}
function saveContact_(auth, id, f) {
  const R = readRoster_(true); const p = R.byId[id];
  if (!p) throw new Error('명단에서 이 사람을 찾을 수 없어요');
  if (bookScope_(auth, p.team).grades.indexOf(p.group) < 0) throw new Error('이 사람의 주소록을 고칠 권한이 없어요');
  f = f || {};
  const row = [p.id, safeCell_(p.name), TEAMS[p.team].label, safeCell_(groupLabel_(p.team, p.group)), String(f.phone || '').trim(), String(f.parentPhone || '').trim(), safeCell_(String(f.guardian || '').trim()), safeCell_(String(f.address || '').trim())];
  const B = readBook_();
  const r = B.rows[id] || (B.sh.getLastRow() + 1);
  fmt_(B.sh.getRange(r, 1, 1, 1), '@'); fmt_(B.sh.getRange(r, 5, 1, 2), '@');
  B.sh.getRange(r, 1, 1, BCOLS.length).setValues([row]);
  return true;
}
/**
 * 대표전화: 모든 부서에서 학생마다 고를 수 있다. 정하지 않으면 영아·유아는 부모, 그 외는 학생.
 * 학생별: teamSettings[부서].callFirstByPerson[학생 id] = 'parent' | 'student'.
 * 연락처를 고칠 수 있는 사람(임원, 그 반 담임)만. value 가 비어 있으면 반·부서 설정을 따른다.
 */
const CALL_CHOICE_TEAMS = ['youth'];
function requireCallChoice_(team) {
  if (CALL_CHOICE_TEAMS.indexOf(team) < 0) throw new Error('대표전화는 ' + CALL_CHOICE_TEAMS.map(t => TEAMS[t].label).join(', ') + '만 고를 수 있어요. 영아·유아는 부모님, 중고등은 학생 번호로 정해져 있어요.');
}
function setPersonCallFirst_(auth, id, value) {
  const p = readRoster_(true).byId[id];
  if (!p) throw new Error('명단에서 이 사람을 찾을 수 없어요');
  if (p.group === 'T') throw new Error('선생님은 연락처가 하나라 대표전화를 정하지 않아요');
  if (bookScope_(auth, p.team).grades.indexOf(p.group) < 0) throw new Error('이 학생의 대표전화를 정할 권한이 없어요');
  if (value && ['parent', 'student'].indexOf(value) < 0) throw new Error('대표전화는 부모 또는 학생이에요');
  const all = cfgAll_().teamSettings || {};
  const cur = Object.assign({}, all[p.team] || {});
  const byPerson = Object.assign({}, cur.callFirstByPerson || {});
  if (value) byPerson[id] = value; else delete byPerson[id];
  cur.callFirstByPerson = byPerson;
  all[p.team] = cur;
  cfgSet_('teamSettings', all);
  return all[p.team];
}
/** 새 반 담임 번호: 영문 소문자·숫자 8자리 (헷갈리는 l·o·0·1 빼고, 영문과 숫자 모두 들어감) */
function newGradeCode_() {
  const L = 'abcdefghjkmnpqrstuvwxyz', N = '23456789', A = L + N;
  for (;;) {
    const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Math.random());
    const c = b.slice(0, 8).map(x => A[(x + 256) % A.length]).join('');
    if (/[a-z]/.test(c) && /[0-9]/.test(c)) return c;
  }
}

/* ---------------- 예시 데이터 지우기 (편집기에서 한 번만 실행) ---------------- */
/**
 * Apps Script 편집기 위쪽 함수 목록에서 clearExampleData 를 고르고 [실행]을 누르세요.
 * 지우는 것 : 명단·출석·달란트·주소록·출결메모·이력 탭의 내용(머리글은 남김), 반 설정의 담임·분반, 등반 기록
 * 남기는 것 : 관리자 비밀번호, 부서 고유번호, 학년 번호, 주일 날짜 머리글, 통계 수식
 */
function clearExampleData() {
  ensureSheets_(true);
  const ss = ss_();
  [SHEET.roster, SHEET.att, SHEET.tal, SHEET.book, SHEET.memo, SHEET.hist].forEach(name => {
    const sh = ss.getSheetByName(name); if (!sh) return;
    const n = sh.getLastRow(), c = Math.max(sh.getLastColumn(), 1);
    if (n >= 2) sh.getRange(2, 1, n - 1, c).clearContent();
  });
  writeClasses_({});
  const cfg = cfgAll_();
  Object.keys(cfg).forEach(k => { if (k.indexOf('promo:') === 0) cfgSet_(k, null); });
  const guide = ss.getSheetByName('안내');
  if (guide) {
    const n = guide.getLastRow();
    if (n >= 1) guide.getRange(1, 1, n, 1).getValues().forEach((r, i) => {
      if (String(r[0]).indexOf('예시 데이터') === 0) guide.getRange(i + 1, 1).setValue('실제 명단과 출석을 입력해 사용하세요. (예시 데이터는 지웠어요)');
    });
  }
  const stat = ss.getSheetByName('통계');
  if (stat) stat.getRange('B3').setValue('2026-01-04');
  SpreadsheetApp.flush();
  Logger.log('예시 데이터를 모두 지웠어요.');
}

/**
 * 다른 계정으로 시트를 [사본 만들기] 해서 새로 시작할 때 편집기에서 한 번 실행하세요. (되돌릴 수 없어요)
 * clearExampleData 에 더해 관리자 비밀번호, 임원·담임 번호, 부서 설정, 일정·자료실 목록, 바로가기 주소도 지웁니다.
 * (일정·자료실 파일은 원래 계정의 드라이브에 있어서 새 계정에서는 열 수 없기 때문에 목록만 지워요)
 */
function resetForNewAccount() {
  clearExampleData();
  const ss = ss_(); const cfg = cfgSheet_();
  [SHEET.events, SHEET.lib].forEach(name => {
    const sh = ss.getSheetByName(name); if (!sh) return;
    const n = sh.getLastRow(), c = Math.max(sh.getLastColumn(), 1);
    if (n >= 2) sh.getRange(2, 1, n - 1, c).clearContent();
  });
  cfg.getRange('B1').clearContent(); cfg.getRange('F1:F2').clearContent();
  const n = cfg.getLastRow(); if (n >= 3) cfg.deleteRows(3, n - 2);
  DriveApp.getRootFolder();
  SpreadsheetApp.flush();
  Logger.log('새 계정용으로 초기화했어요. _설정 탭 B1에 새 관리자 비밀번호를 적은 뒤 [배포 → 새 배포]를 해 주세요.');
}

/** 담임 선생님 본인 비밀번호 만들기·바꾸기. 처음에는 반 담임 번호를, 그 뒤로는 지금 개인 비밀번호를 확인한다 */
function changeMyCode_(auth, team, code, old) {
  if (!TEAMS[team]) throw new Error('부서를 알 수 없어요');
  let me = meOf_(auth, team), viaSub = false;
  if (!me) { const sc = scope_(auth, team); if (sc.level === 'sub') { const p = subTeacher_(readRoster_(true), team, sc.name); if (p) { me = { pid: p.id, name: p.name }; viaSub = true; } } }
  if (!me || !me.pid) throw new Error('본인 이름으로 로그인한 선생님만 바꿀 수 있어요. 메인 화면에서 이름과 비밀번호로 다시 들어와 주세요.');
  const kind = 'entry:' + team;
  if (guardStop_(kind, team, TEAMS[team].label + ' ' + me.name + ' 선생님 비밀번호 변경')) throw new Error(GUARD_MSG);
  const tc = Object.assign({}, cfgAll_().tcodes || {});
  const oldOk = me.dev ? true : tc[me.pid] ? tcCheck_(tc[me.pid], old) : false;
  if (!oldOk) { guardFail_(kind, String(old || ''), team, TEAMS[team].label + ' ' + me.name + ' 선생님 비밀번호 변경 (지금 번호 틀림)'); throw new Error(tc[me.pid] ? '지금 쓰는 비밀번호가 맞지 않아요' : viaSub ? '지금 쓰는 번호(보조교사 비밀번호)가 맞지 않아요' : '지금 쓰는 번호(반 담임 번호)가 맞지 않아요'); }
  code = String(code || '').trim();
  if (!PW_RULE.test(code)) throw new Error(PW_RULE_MSG);
  const first = !tc[me.pid];
  const salt = Utilities.getUuid();
  tc[me.pid] = { t: team, n: me.name, s: salt, h: coHash_(salt, code.toLowerCase()) };
  cfgSet_('tcodes', tc);
  logRows_(Object.assign({}, auth, { hrMe: Object.assign({}, auth.hrMe, { [team]: Object.assign({}, me, { own: true }) }) }), [[me.pid, me.name, team, '번호', me.name + ' 선생님 개인 비밀번호를 ' + (first ? '만듦' : '바꿈')]]);
  dropSharedCodes_(auth, team);
  return true;
}
/** 담임이 여러 명인 반에서 담임 모두가 개인 비밀번호를 가지면 같이 쓰던 반 담임 번호를 지운다. 지운 반 이름들 */
function dropSharedCodes_(auth, team) {
  const g = gradeCodes_(); if (!Object.keys(g).length) return [];
  const tc = cfgAll_().tcodes || {}, cls = readClasses_(), R = readRoster_(true), gone = [];
  groupsOf_(team).forEach(([k, gl]) => {
    if (!g[team + '_' + k]) return;
    const names = Object.keys(homeroomNames_(team, [k], cls)); if (names.length < 2) return;
    const all = names.every(n => { const p = R.list.find(x => x.team === team && x.group === 'T' && x.active !== false && x.name === n) || R.list.find(x => x.team === team && x.group === 'T' && x.active !== false && sameName_(x.name, n)); return p && tc[p.id] && tc[p.id].t === team; });
    if (all) { delete g[team + '_' + k]; gone.push(gl); }
  });
  if (!gone.length) return gone;
  cfgSet_('gradeCodes', g);
  logRows_(auth, [['', '', team, '번호', gone.join('·') + ' 담임 번호를 지움 (담임 모두 개인 비밀번호 사용)']]);
  return gone;
}

/* ---------------- 샘플 명단 넣기 (편집기에서 한 번만 실행) ---------------- */
/**
 * 함수 목록에서 resetSampleData 를 고르고 [실행]을 누르세요.
 * 예시 데이터를 모두 지운 뒤(clearExampleData) 샘플 명단만 넣습니다.
 *   부서마다 5명씩 (영아 0~4세 · 유아 5·6·7·5·6세 · 유소년 1~5학년 · 중고등 중1~중3·고1·고2)
 */
function resetSampleData() {
  clearExampleData();
  const plan = [['infant', '영아', ['0', '1', '2', '3', '4']], ['toddler', '유아', ['5', '6', '7', '5', '6']],
    ['youth', '유소년', ['1', '2', '3', '4', '5']], ['teen', '중고등', ['m1', 'm2', 'm3', 'h1', 'h2']]];
  const people = [];
  plan.forEach(([team, short, groups]) => groups.forEach((g, i) => people.push({
    id: 'sample-' + team + '-' + (i + 1), team, group: g, section: '', name: short + '샘플' + (i + 1),
    birthday: ('0' + (i + 3)).slice(-2) + '/' + ('0' + (i * 5 + 2)).slice(-2), role: '', isNew: false, active: true,
    note: '샘플 명단 · 확인 후 지우세요', photo: '', order: i + 1, promotedYear: null, talentAdj: {},
    age: isFlex_(team) ? g : '',
  })));
  savePersonRows_(readRoster_(), people);
  SpreadsheetApp.flush();
  Logger.log('샘플 명단 ' + people.length + '명을 넣었어요.');
}

/* ---------------- 부서 설정 (임원) ---------------- */
/**
 * 부서별 설정. talent(달란트 사용 여부, 기록은 지우지 않고 화면에서만 숨긴다), talentLabel(항목 이름),
 * callFirst(전화 걸 때 먼저 걸 번호: 'parent' | 'student'. 유소년만 고를 수 있고, 없으면 학생)
 */
function setTeamSetting_(auth, team, key, value) {
  requirePromoter_(auth, team);
  if (['talent', 'talentLabel', 'callFirst'].indexOf(key) < 0) throw new Error('바꿀 수 없는 설정이에요');
  if (key === 'talentLabel') { value = String(value || '').trim().slice(0, 10); if (!value) throw new Error('항목 이름을 적어 주세요'); }
  else if (key === 'callFirst') { requireCallChoice_(team); if (['parent', 'student'].indexOf(value) < 0) throw new Error('전화 우선순위는 부모 또는 학생이에요'); }
  else value = !!value;
  const all = cfgAll_().teamSettings || {};
  all[team] = Object.assign({}, all[team] || {}, { [key]: value });
  cfgSet_('teamSettings', all);
  logRows_(auth, [['', '', team, '설정', key === 'talent' ? (value ? '달란트 사용으로 바꿈' : '달란트 숨김으로 바꿈') : key === 'talentLabel' ? '달란트 항목 이름: ' + value : '전화 우선순위: ' + value]]);
  return all[team];
}
/**
 * 반별 전화 우선순위 (담임은 자기 반만, 임원은 모든 반). teamSettings[부서].callFirstByClass[반] = 'parent' | 'student'.
 * value 가 비어 있으면 반 설정을 지우고 부서 설정을 따른다.
 */
function setClassCallFirst_(auth, team, group, value) {
  group = String(group || '');
  if (!group || group === 'T' || !groupsOf_(team).some(x => x[0] === group)) throw new Error('반을 알 수 없어요');
  requireCallChoice_(team);
  requirePerson_(auth, team, group);
  if (value && ['parent', 'student'].indexOf(value) < 0) throw new Error('전화 우선순위는 부모 또는 학생이에요');
  const all = cfgAll_().teamSettings || {};
  const cur = Object.assign({}, all[team] || {});
  const byClass = Object.assign({}, cur.callFirstByClass || {});
  if (value) byClass[group] = value; else delete byClass[group];
  cur.callFirstByClass = byClass;
  all[team] = cur;
  cfgSet_('teamSettings', all);
  return all[team];
}

/**
 * 영아교육팀 반 편성 (임원 이상). list = [[키, 반 이름], ...] 순서대로.
 * 새 반은 키를 비워 보내면 만들어 준다. 재적 아이가 있는 반은 없앨 수 없다(먼저 반 배정으로 옮긴다).
 * 없앤 반에 남은 떠난 아이는 첫 번째 반으로, 그 반 담임·담임 번호는 지운다.
 */
function setClasses_(auth, team, list) {
  requirePromoter_(auth, team);
  if (!Array.isArray(list) || !list.length) throw new Error('반을 하나 이상 만들어 주세요');
  // 유아·유소년·중고등은 이름만: 반 순서·개수는 나이·학년 그대로
  if (!isFlex_(team)) {
    const byKey = {}; list.forEach(([k, l]) => byKey[String(k)] = l);
    list = TEAMS[team].groups.map(([k, l]) => [k, String(byKey[k] == null ? l : byKey[k])]);
  }
  if (list.length > 20) throw new Error('반은 20개까지 만들 수 있어요');
  const old = groupsOf_(team).map(x => x.slice());
  const seen = {}; const next = [];
  list.forEach(([k, l], i) => {
    l = String(l || '').trim().slice(0, 12);
    if (!l) throw new Error((i + 1) + '번째 반 이름을 적어 주세요');
    if (l === '교사' || l === '미배정') throw new Error('"' + l + '"은(는) 반 이름으로 쓸 수 없어요');
    if (seen[l]) throw new Error('반 이름 "' + l + '"이(가) 겹쳐요');
    seen[l] = true;
    k = String(k || '').trim();
    if (isFlex_(team) && (!k || k === 'T' || k === '_' || next.some(x => x[0] === k))) k = 'c' + Date.now().toString(36) + i;
    next.push([k, l]);
  });
  const R = readRoster_(); const classMap = readClasses_();
  const keep = {}; next.forEach(x => keep[x[0]] = x[1]);
  old.forEach(([k, l]) => {
    if (keep[k] != null) return;
    const n = R.list.filter(p => p.team === team && p.group === k && p.active !== false).length;
    if (n) throw new Error('"' + l + '"에 재적 아이가 ' + n + '명 있어요. 먼저 [반 배정]에서 다른 반으로 옮겨 주세요');
  });
  const all = cfgAll_().teamSettings || {};
  all[team] = Object.assign({}, all[team] || {}, { classes: next });
  cfgSet_('teamSettings', all); TEAM_SETTINGS_CACHE_ = null;
  // 이름이 바뀌었거나 없어진 반의 사람들: 시트의 반 칸을 새 이름(또는 첫 번째 반)으로
  const changed = {}; old.forEach(([k, l]) => { if (keep[k] !== l) changed[k] = true; });
  const moved = R.list.filter(p => p.team === team && p.group !== 'T' && changed[p.group])
    .map(p => Object.assign({}, p, keep[p.group] ? {} : { group: next[0][0], section: '' }));
  if (moved.length) savePersonRows_(R, moved);
  Object.keys(classMap).forEach(key => { if (key.indexOf(team + '_') === 0 && !keep[key.slice(team.length + 1)]) delete classMap[key]; });
  writeClasses_(classMap);
  const g = gradeCodes_(); let gc = false;
  Object.keys(g).forEach(key => { if (key.indexOf(team + '_') === 0 && !keep[key.slice(team.length + 1)]) { delete g[key]; gc = true; } });
  if (gc) cfgSet_('gradeCodes', g);
  logRows_(auth, [['', '', team, '설정', '반 편성 저장: ' + next.map(x => x[1]).join(', ')]]);
  return { classes: next };
}

/* ---------------- 엑셀 명단 올리기 (임원 이상) ---------------- */
/** 엑셀 올리기에서 같은 사람인지 알아보는 도우미 */
function phonesOf_(x) { return [x.phone, x.parentPhone].map(v => String(v || '').replace(/[^0-9]/g, '')).filter(v => v.length >= 7); }
function sharePhone_(a, b) { return a.some(v => b.indexOf(v) >= 0); }
function sameBd_(a, b) {
  a = String(a || '').trim(); b = String(b || '').trim();
  if (!a || !b || a.slice(-5) !== b.slice(-5)) return false;
  const ya = a.length > 5 ? a.slice(0, 2) : '', yb = b.length > 5 ? b.slice(0, 2) : '';
  return !ya || !yb || ya === yb;
}
function sameRow_(a, b) {
  const F = ['name', 'gender', 'group', 'age', 'section', 'birthday', 'lunar', 'role', 'isNew', 'note', 'guardian', 'parentPhone', 'phone', 'address'];
  const key = x => F.map(f => String(x[f] == null ? '' : x[f]).replace(/\s/g, '')).join('\u0001');
  if (key(a) === key(b)) return true;
  return !!String(a.name || '').trim() && String(a.name).trim() === String(b.name || '').trim() && sameBd_(a.birthday, b.birthday) && sharePhone_(phonesOf_(a), phonesOf_(b));
}

/**
 * rows: [{ name, gender(남/여), group(반 이름 또는 키), section, birthday, role(교사만), isNew, note, guardian, parentPhone, phone, address }]
 * 같은 부서에서 이름·생일·연락처가 같은 사람(없으면 이름과 반이 같은 사람)은 기존 정보 위에 덮어쓰고, 없으면 새로 추가한다.
 */
function importRoster_(auth, team, rows) {
  if (!TEAMS[team]) throw new Error('부서를 알 수 없어요');
  const admin = deptLevel_(auth, team); const sc = scope_(auth, team);
  if (!admin && sc.level !== 'grade') throw new Error('임원이나 담임만 명단을 올릴 수 있어요');
  if (!Array.isArray(rows) || !rows.length) throw new Error('올릴 명단이 비어 있어요');
  if (rows.length > 500) throw new Error('한 번에 500명까지 올릴 수 있어요');
  const R = readRoster_(); const B = readBook_();
  const labelToKey = {};
  if (!isFlex_(team)) TEAMS[team].groups.forEach(([k, gl]) => { labelToKey[gl] = k; }); // 바꾸기 전 이름(1학년 등)도 받아 준다
  groupsOf_(team).forEach(([k, gl]) => { labelToKey[gl.replace(/\s/g, '')] = k; labelToKey[k] = k; }); labelToKey['교사'] = 'T';
  const out = { created: 0, updated: 0, skipped: [], duplicates: 0 }; const people = []; const contacts = [];
  // 이름·생일·연락처가 같은 줄(또는 모든 칸이 같은 줄)은 한 사람으로 보고 하나로 합친다. 뒤 줄의 값이 앞 줄 위에 덮어쓴다.
  // 줄 번호는 원래 파일 기준으로 유지
  const FLD = ['name', 'gender', 'group', 'age', 'section', 'birthday', 'lunar', 'role', 'isNew', 'note', 'guardian', 'parentPhone', 'phone', 'address'];
  const uniq = [];
  rows.forEach((r, i) => {
    const same = uniq.find(([u]) => sameRow_(u, r));
    if (!same) { uniq.push([Object.assign({}, r), i]); return; }
    out.duplicates++;
    FLD.forEach(f => { if (String(r[f] == null ? '' : r[f]).trim() !== '') same[0][f] = r[f]; });
  });
  uniq.forEach(([r, i]) => {
    const name = String(r.name || '').trim(); const g = labelToKey[String(r.group || '').trim().replace(/\s/g, '')];
    if (!name) { out.skipped.push({ row: i + 2, reason: '이름이 비어 있어요' }); return; }
    if (!g) { out.skipped.push({ row: i + 2, reason: '반 "' + (r.group || '') + '"을(를) 알 수 없어요' }); return; }
    if (!admin && (g === 'T' || sc.grades.indexOf(g) < 0)) { out.skipped.push({ row: i + 2, reason: '담임은 자기 반만 올릴 수 있어요 (' + (r.group || '') + ')' }); return; }
    // 이름·생일·연락처가 같은 기존 사람이면 그 사람 위에 덮어쓴다(반이 달라도). 없으면 예전처럼 이름·반이 같은 사람
    const old = R.list.find(p => p.team === team && p.name === name && sameBd_(p.birthday, r.birthday) && sharePhone_(phonesOf_(B.map[p.id] || {}), phonesOf_(r)))
      || R.list.find(p => p.team === team && p.group === g && p.name === name);
    if (old && !admin && old.group === 'T') { out.skipped.push({ row: i + 2, reason: '선생님 정보는 임원만 고칠 수 있어요 (' + name + ')' }); return; }
    if (old && !admin && sc.grades.indexOf(old.group) < 0) { out.skipped.push({ row: i + 2, reason: '다른 반 학생이라 고칠 수 없어요 (' + name + ')' }); return; }
    const id = old ? old.id : ('p' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7) + i);
    const sec = admin && /^[A-D]$/.test(String(r.section || '').trim().toUpperCase()) ? String(r.section).trim().toUpperCase() : (old ? old.section : ''); // 분반은 임원 이상만
    const base = old ? Object.assign({}, old) : { active: true, order: 999, talentAdj: {}, photo: '' };
    const bday = String(r.birthday || '').trim();
    const sx = String(r.gender || '').trim().toUpperCase();
    const gender = /^(남|남자|M|MALE)$/.test(sx) ? '남' : /^(여|여자|F|FEMALE)$/.test(sx) ? '여' : (old ? old.gender : '');
    people.push(cleanPerson_(id, Object.assign(base, {
      team, group: g, section: sec, name, gender,
      birthday: bday || (old ? old.birthday : ''),
      lunar: String(r.lunar || '').trim() ? /^(y|yes|o|예|음력|true|1)$/i.test(String(r.lunar).trim()) : (old ? !!old.lunar : false),
      role: g === 'T' && String(r.role || '').trim() ? String(r.role).trim() : (old ? old.role : ''),
      age: ageKey_(team, r.age) || (old ? old.age : (TEAMS[team].groups.some(x => x[0] === g) ? g : '')),
      isNew: /^(y|yes|o|예|새친구|true|1)$/i.test(String(r.isNew || '').trim()) || (old ? old.isNew && !String(r.isNew || '').trim() : false),
      note: r.note != null && String(r.note).trim() !== '' ? String(r.note).trim() : (old ? old.note : ''),
      active: true,
    })));
    contacts.push([id, name, g, r]);
    old ? out.updated++ : out.created++;
  });
  const beforeImp = {}; R.list.forEach(p => beforeImp[p.id] = Object.assign({}, p));
  if (people.length) savePersonRows_(R, people);
  const logs = []; people.forEach(p => logPersonChange_(auth, beforeImp[p.id], p, '엑셀', logs)); logRows_(auth, logs);
  // 연락처: 이미 있는 줄은 한꺼번에 고치고, 새 줄은 한꺼번에 붙인다
  const edit = [], fresh = [];
  contacts.forEach(([id, name, g, r]) => {
    const has = ['guardian', 'parentPhone', 'phone', 'address'].some(k => String(r[k] || '').trim());
    if (!has) return;
    const cur = B.map[id] || {};
    const row = [id, safeCell_(name), TEAMS[team].label, safeCell_(groupLabel_(team, g)),
      String(r.phone || cur.phone || '').trim(), String(r.parentPhone || cur.parentPhone || '').trim(),
      safeCell_(String(r.guardian || cur.guardian || '').trim()), safeCell_(String(r.address || cur.address || '').trim())];
    if (B.rows[id]) edit.push([B.rows[id], row]); else fresh.push(row);
  });
  if (edit.length) {
    const at = edit.map(e => e[0]); const lo = Math.min.apply(null, at), hi = Math.max.apply(null, at);
    fmt_(B.sh.getRange(lo, 1, hi - lo + 1, 1), '@'); fmt_(B.sh.getRange(lo, 5, hi - lo + 1, 2), '@');
    writeRows_(B.sh, 1, BCOLS.length, edit);
  }
  if (fresh.length) {
    const start = B.sh.getLastRow() + 1;
    fmt_(B.sh.getRange(start, 1, fresh.length, 1), '@'); fmt_(B.sh.getRange(start, 5, fresh.length, 2), '@');
    B.sh.getRange(start, 1, fresh.length, BCOLS.length).setValues(fresh);
  }
  return out;
}

/* ---------------- 일정 ---------------- */
/** 일정은 부서에 들어올 수 있는 사람(임원·담임)이 보고, 임원 이상이 만들고 고친다. 파일은 스크립트 주인의 드라이브에 비공개로 저장한다. */
function readEvents_() {
  const sh = ss_().getSheetByName(SHEET.events); const n = sh.getLastRow();
  const list = [];
  if (n >= 2) sh.getRange(2, 1, n - 1, 7).getValues().forEach((r, i) => {
    const id = String(r[0] || '').trim(); if (!id) return;
    let files = []; try { files = r[5] ? JSON.parse(r[5]) : []; } catch (e) {}
    list.push({ row: i + 2, id, team: TEAM_BY_LABEL[String(r[1]).trim()] || '', date: dateKey_(r[2]), title: String(r[3] || ''), memo: String(r[4] || ''), files });
  });
  return { sh, list };
}
function requireView_(auth, team) {
  if (!TEAMS[team]) throw new Error('부서를 알 수 없어요');
  if (scope_(auth, team).level === 'none') throw new Error(TEAMS[team].label + ' 번호가 필요해요');
}
function listEvents_(auth, team) {
  requireView_(auth, team);
  const admin = deptLevel_(auth, team);
  return { canEdit: admin, events: cachedEvents_().filter(e => e.team === team).map(e => {
    const o = Object.assign({}, e); delete o.row;
    o.files = (o.files || []).map(f => Object.assign({ vis: 'all' }, f)).filter(f => admin || f.vis !== 'admin');
    return o; }) };
}
function saveEvent_(auth, team, ev) {
  requirePromoter_(auth, team);
  ev = ev || {};
  const date = String(ev.date || ''); if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('날짜를 골라 주세요');
  const title = String(ev.title || '').trim().slice(0, 80); if (!title) throw new Error('일정 제목을 적어 주세요');
  const E = readEvents_(); const old = ev.id ? E.list.find(e => e.id === ev.id && e.team === team) : null;
  const id = old ? old.id : ('e' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6));
  const files = old ? old.files : [];
  const vals = [[id, TEAMS[team].label, date, safeCell_(title), safeCell_(String(ev.memo || '').slice(0, 2000)), JSON.stringify(files), Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm')]];
  const r = old ? old.row : E.sh.getLastRow() + 1;
  fmt_(E.sh.getRange(r, 1, 1, 3), '@'); E.sh.getRange(r, 1, 1, 7).setValues(vals);
  logRows_(auth, [['', '', team, '일정', (old ? '일정 수정: ' : '일정 등록: ') + date + ' ' + title]]);
  return { id, team, date, title, memo: String(ev.memo || '').slice(0, 2000), files };
}
function deleteEvent_(auth, team, id) {
  requirePromoter_(auth, team);
  const E = readEvents_(); const ev = E.list.find(e => e.id === id && e.team === team); if (!ev) return true;
  (ev.files || []).forEach(f => { try { DriveApp.getFileById(f.id).setTrashed(true); } catch (e) {} });
  logRows_(auth, [['', '', team, '일정', '일정 삭제: ' + ev.date + ' ' + ev.title]]);
  E.sh.deleteRow(ev.row); return true;
}
function eventsFolder_(team) {
  const root = (() => { const it = DriveApp.getFoldersByName('교육팀 출석부 공유문서'); return it.hasNext() ? it.next() : DriveApp.createFolder('교육팀 출석부 공유문서'); })();
  const name = TEAMS[team].label; const it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}
function downloadFile_(auth, team, fileId) {
  requireView_(auth, team);
  const L = cachedLib_().find(f => f.team === team && f.id === fileId);
  if (!L) throw new Error('파일을 찾을 수 없어요');
  if (L.vis === 'admin' && !deptLevel_(auth, team)) throw new Error('관리자만 볼 수 있는 문서예요');
  const r = libFileBlob_(L);
  return { name: r.name, mime: r.blob.getContentType(), b64: Utilities.base64Encode(r.blob.getBytes()) };
}
/**
 * 자료실 파일의 최신 내용. 온라인 작성본(구글 시트·문서·슬라이드)이 있으면 그 내용을 원래 형식으로 바꿔 돌려주고,
 * 원본과 형식이 같으면(xlsx·docx·pptx·csv·txt) 원본 파일에도 덮어써서 원본과 작성본이 늘 같게 한다.
 * (.xls·.xlsm·.doc·.ppt 등은 형식이 달라 원본을 덮어쓰지 않고, 새 형식으로 바꿔 돌려준다)
 */
function libFileBlob_(L) {
  const d = (cfgAll_().libDocs || {})[L.id]; const kind = docKind_(L.name);
  if (d && d.id && kind) {
    try {
      const name = kind.wb ? L.name : L.name.replace(/[.][^.]+$/, '') + '.' + kind.out;
      const xb = DriveApp.getFileById(d.id).getAs(kind.exp).setName(name);
      if (kind.wb) {
        try {
          UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files/' + L.id + '?uploadType=media', {
            method: 'patch', contentType: xb.getContentType(), payload: xb.getBytes(),
            headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
        } catch (e) {}
      }
      return { blob: xb, name };
    } catch (e) {}
  }
  const b = DriveApp.getFileById(L.id).getBlob();
  return { blob: b, name: L.name };
}
/** 자료실(파일 올리기·온라인 작성)을 처음 쓸 때 편집기에서 한 번 실행해 드라이브 사용을 허용하세요. */
function authorizeDrive() {
  DriveApp.getRootFolder();
  TEAM_ORDER.forEach(t => eventsFolder_(t));
  Logger.log('드라이브 사용을 허용했어요. 내 드라이브에 "교육팀 출석부 공유문서" 폴더(부서별 4개)가 만들어졌어요. 이제 [배포 → 배포 관리 → 연필 → 새 버전 → 배포]를 해 주세요.');
}


/**
 * 온라인 작성이 가능한 파일 종류. 구글 시트·문서·슬라이드로 바꿔 열고, 고친 내용은 같은 형식으로 다시 내보낸다.
 *   wb: true 이면 원본 파일(같은 형식)에도 덮어쓴다. 아니면 원본은 그대로 두고 out 형식으로 바꿔 돌려준다.
 */
const XLSX_ = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const DOCX_ = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const PPTX_ = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const DOC_KINDS = {
  xlsx: { src: XLSX_, g: 'spreadsheet', out: 'xlsx', exp: XLSX_, wb: true },
  xlsm: { src: 'application/vnd.ms-excel.sheet.macroEnabled.12', g: 'spreadsheet', out: 'xlsx', exp: XLSX_ },
  xls:  { src: 'application/vnd.ms-excel', g: 'spreadsheet', out: 'xlsx', exp: XLSX_ },
  ods:  { src: 'application/vnd.oasis.opendocument.spreadsheet', g: 'spreadsheet', out: 'xlsx', exp: XLSX_ },
  csv:  { src: 'text/csv', g: 'spreadsheet', out: 'csv', exp: 'text/csv', wb: true },
  docx: { src: DOCX_, g: 'document', out: 'docx', exp: DOCX_, wb: true },
  doc:  { src: 'application/msword', g: 'document', out: 'docx', exp: DOCX_ },
  odt:  { src: 'application/vnd.oasis.opendocument.text', g: 'document', out: 'docx', exp: DOCX_ },
  rtf:  { src: 'application/rtf', g: 'document', out: 'docx', exp: DOCX_ },
  txt:  { src: 'text/plain', g: 'document', out: 'txt', exp: 'text/plain', wb: true },
  pptx: { src: PPTX_, g: 'presentation', out: 'pptx', exp: PPTX_, wb: true },
  ppt:  { src: 'application/vnd.ms-powerpoint', g: 'presentation', out: 'pptx', exp: PPTX_ },
  odp:  { src: 'application/vnd.oasis.opendocument.presentation', g: 'presentation', out: 'pptx', exp: PPTX_ },
};
const G_LABEL = { spreadsheet: '구글 시트', document: '구글 문서', presentation: '구글 슬라이드' };
function docKind_(fileName) {
  const ext = (String(fileName || '').toLowerCase().match(/[.]([a-z0-9]+)$/) || [])[1] || '';
  return DOC_KINDS[ext] || null;
}
/** 파일을 구글 시트·문서·슬라이드 작성본으로 바꿔 부서 폴더에 만든다 (내려받지 않고 바로 열어 고치고 자동 저장). 매크로는 사라진다. */
function docFromBlob_(team, blob, kind, title) {
  const boundary = 'ya' + Date.now();
  const meta = JSON.stringify({ name: title, mimeType: 'application/vnd.google-apps.' + kind.g, parents: [eventsFolder_(team).getId()] });
  const head = '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta + '\r\n--' + boundary + '\r\nContent-Type: ' + kind.src + '\r\n\r\n';
  const payload = Utilities.newBlob(head).getBytes().concat(blob.getBytes()).concat(Utilities.newBlob('\r\n--' + boundary + '--').getBytes());
  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', {
    method: 'post', contentType: 'multipart/related; boundary=' + boundary, payload,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (res.getResponseCode() >= 300) throw new Error('온라인 작성본으로 바꾸지 못했어요 (오류 ' + res.getResponseCode() + '). 편집기에서 authorizeMail 을 다시 실행해 권한을 허용하고 새 버전으로 배포했는지 확인해 주세요');
  const id = JSON.parse(res.getContentText()).id;
  const f = DriveApp.getFileById(id);
  let shared = true;
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT); } catch (e) { shared = false; }
  return { id, url: f.getUrl(), shared, g: kind.g };
}
function libDocCreate_(auth, team, fileId) {
  requirePromoter_(auth, team);
  const L = readLib_().list.find(x => x.team === team && x.id === String(fileId || '')); if (!L) throw new Error('자료실에서 파일을 찾을 수 없어요');
  const kind = docKind_(L.name);
  if (!kind) throw new Error('엑셀·워드·파워포인트·텍스트·CSV 파일만 온라인으로 작성할 수 있어요 (PDF·한글·사진은 안 돼요)');
  if ((cfgAll_().libDocs || {})[L.id]) throw new Error('이미 온라인 작성본이 있어요');
  let src; try { src = DriveApp.getFileById(L.id).getBlob(); } catch (e) { throw new Error('파일을 열 수 없어요. 자료실에 다시 올려 주세요'); }
  const title = L.name.replace(/[.][^.]+$/, '') + ' 작성본';
  const made = docFromBlob_(team, src, kind, title);
  const docs = cfgAll_().libDocs || {};
  const info = { id: made.id, url: made.url, name: title, shared: made.shared, g: kind.g, created: Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm') };
  docs[L.id] = info; cfgSet_('libDocs', docs);
  logRows_(auth, [['', '', team, '자료실', '온라인 작성본(' + G_LABEL[kind.g] + ') 만듦: ' + L.name]]);
  return info;
}
/** 메일 보내기 기능을 처음 쓸 때 편집기에서 한 번 실행해 메일 발송을 허용하세요. */
function authorizeMail() {
  UrlFetchApp.fetch('https://www.google.com/generate_204', { muteHttpExceptions: true }); // 외부 접속 권한도 함께 허용
  Logger.log('오늘 남은 메일 발송 수: ' + MailApp.getRemainingDailyQuota() + '. 메일 발송을 허용했어요. 이제 [배포 → 배포 관리 → 연필 → 새 버전 → 배포]를 해 주세요.');
}

/* ---------------- 이력 (학생·교사) ---------------- */
function kindOf_(p) { return p.group === 'T' ? '교사' : '학생'; }
function placeLabel_(p) { return (TEAMS[p.team] ? TEAMS[p.team].label.replace('교육팀', '') + ' ' : '') + groupLabel_(p.team, p.group) + (p.section ? ' ' + p.section + '반' : ''); }
/** 요청에 실린 이메일로 등록 명단에서 이름을 찾는다 */
function staffNameOf_(auth) {
  const e = normEmail_(auth && auth.email); if (!e) return '';
  const hit = staffList_().find(s => normEmail_(s.email) === e);
  return hit ? String(hit.name || '') : '';
}
function actorLabel_(auth, team) {
  if (isAdmin_(auth)) return auth.co ? (auth.co2 ? '공동관리자2 ' : '공동관리자1 ') + (auth.coName || '') : '총관리자';
  const sc = scope_(auth, team);
  if (sc.level === 'grade') return '담임(' + sc.grades.map(g => groupLabel_(team, g)).join(', ') + ')';
  if (sc.level === 'sub') return '보조교사(' + sc.name + ')';
  if (deptLevel_(auth, team)) { const nm = staffNameOf_(auth); return nm ? '임원(' + nm + ')' : '임원'; }
  return '번호 없음';
}
/** 활동기록에 적는 '바꾼 사람'. 담임 번호면 반 설정에 배치한 담임 이름(여러 명, 분반 담임 포함)도 함께: 담임(6학년 김은혜·박요한)
 *  (자료실 '내 파일' 구분은 이름이 바뀌어도 그대로이도록 actorLabel_ 를 그대로 쓴다) */
function logActor_(auth, team, cls) {
  if (auth.sys) return auth.sys;   // 사람이 아닌 자동 작업(주일 생일 알림 등)
  const sc = scope_(auth, team);
  if (sc.level !== 'grade') { const l = actorLabel_(auth, team); return l === '번호 없음' ? '확인 안 된 사용자' : l; }
  const me = meOf_(auth, team, cls); // 본인 이름으로 들어온 선생님이면 그 이름만
  if (me) return '담임(' + sc.grades.map(g => groupLabel_(team, g)).join(', ') + ' ' + me.name + ')';
  return '담임(' + sc.grades.map(g => { const n = Object.keys(homeroomNames_(team, [g], cls)); return groupLabel_(team, g) + (n.length ? ' ' + n.join('·') : ''); }).join(', ') + ')';
}
/** rows: [[id, name, team, 구분, 내용]] */
function logRows_(auth, rows) {
  if (!rows.length) return;
  const sh = ss_().getSheetByName(SHEET.hist); if (!sh) return;
  const now = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm');
  let cls = {}; try { cls = readClasses_(); } catch (e) {}
  const by = {}, actor = t => by[t] || (by[t] = logActor_(auth, t, cls));
  const vals = rows.map(([id, name, team, kind, text]) => [now, id || '', safeCell_(name || ''), TEAMS[team] ? TEAMS[team].label : '', kind, safeCell_(text), actor(team)]);
  const r = sh.getLastRow() + 1;
  fmt_(sh.getRange(r, 1, vals.length, 2), '@');
  sh.getRange(r, 1, vals.length, 7).setValues(vals);
  // 텔레그램으로도 보낼 줄 (요청·텔레그램 답장·생일은 이미 따로 알림이 가므로 뺀다). 요청이 끝날 때 tgFlushLog_ 가 한 번에 보낸다
  vals.forEach(v => { if (v[4] !== '요청' && v[4] !== '생일' && String(v[5]).indexOf('텔레그램 답장') !== 0)
    TG_LOG_.push({ k: v[4], s: '• ' + (v[3] ? '[' + v[3] + '] ' : '') + v[4] + (v[2] ? ' · ' + v[2] : '') + ': ' + v[5] + (v[6] ? ' — ' + v[6] : '') }); });
}
/** 텔레그램으로 받을 알림 종류. 설정(tgKinds)에 false 로 꺼 둔 것만 안 보낸다 (새 종류는 '기타'를 따른다) */
const TG_KINDS = [['요청', '문의 ([관리자에게 요청하기])'], ['로그인', '로그인'], ['보안', '보안 (로그인 실패·잠금)'], ['번호', '번호 변경'], ['이메일', '이메일 등록'],
  ['학생', '학생 명단'], ['교사', '교사 명단'], ['담임', '담임 지정'], ['반 설정', '반 설정'], ['일정', '일정'], ['자료실', '자료실'], ['당회보고', '당회보고'],
  ['설정', '설정'], ['출석', '출석·달란트 해제 (오후 6시 묶음)'], ['출석체크', '출석 체크 (오후 6시 묶음)'], ['답장', '활동 기록에서 보낸 답장'], ['기타', '기타']];
function tgOn_(kind) {
  const off = cfgAll_().tgKinds || {};
  const k = TG_KINDS.some(x => x[0] === kind) ? kind : '기타';
  return off[k] !== false;
}
function tgKinds_(auth) {
  if (!isRoot_(auth)) throw new Error('총관리자만 볼 수 있어요');
  const p = PropertiesService.getScriptProperties();
  return { linked: !!(p.getProperty('TG_TOKEN') && p.getProperty('TG_CHAT')), kinds: TG_KINDS.map(([k, label]) => ({ k: k, label: label, on: tgOn_(k) })) };
}
function tgKindsSet_(auth, on) {
  if (!isRoot_(auth)) throw new Error('총관리자만 바꿀 수 있어요');
  on = on || {}; const off = {};
  TG_KINDS.forEach(([k]) => { if (on[k] === false) off[k] = false; });
  cfgSet_('tgKinds', off);
  const offList = TG_KINDS.filter(x => off[x[0]] === false).map(x => x[1]);
  logRows_(auth, [['', '', '', '설정', '텔레그램 알림 종류 변경' + (offList.length ? ' (끔: ' + offList.join(', ') + ')' : ' (모두 켬)')]]);
  return tgKinds_(auth);
}
/** 이번 요청에서 쌓인 활동 기록을 텔레그램으로 한 번에 보낸다. 꺼 둔 종류는 빼고, 너무 많으면 앞의 30줄만 */
let TG_LOG_ = [];
function tgFlushLog_() {
  if (!TG_LOG_.length) return;
  let L = TG_LOG_; TG_LOG_ = [];
  try { L = L.filter(x => tgOn_(x.k)).map(x => x.s); } catch (e) { return; }
  if (!L.length) return;
  const head = '📝 활동 기록 ' + Utilities.formatDate(new Date(), tz_(), 'HH:mm') + '\n';
  tgNotify_(head + L.slice(0, 30).map(s => s.slice(0, 300)).join('\n') + (L.length > 30 ? '\n… 외 ' + (L.length - 30) + '건 (활동 기록에서 확인)' : ''));
}
/** sink(배열)를 주면 시트에 바로 쓰지 않고 거기에 모은다 (여러 명을 한 번에 기록할 때) */
function logPersonChange_(auth, old, p, reason, sink) {
  const kind = kindOf_(p); const msgs = [];
  if (!old) msgs.push((reason === '엑셀' ? '엑셀로 등록' : '새로 등록') + ' (' + placeLabel_(p) + ')');
  else {
    const tag = reason && reason !== '엑셀' ? reason + ': ' : (reason === '엑셀' ? '엑셀로 수정: ' : '');
    if (old.team !== p.team) msgs.push(tag + '부서 이동: ' + placeLabel_(old) + ' → ' + placeLabel_(p));
    else if (old.group !== p.group) msgs.push(tag + '반 이동: ' + groupLabel_(old.team, old.group) + ' → ' + groupLabel_(p.team, p.group));
    if (old.team === p.team && old.group === p.group && (old.section || '') !== (p.section || '')) msgs.push('분반: ' + (old.section ? old.section + '반' : '없음') + ' → ' + (p.section ? p.section + '반' : '없음'));
    if (old.active !== false && p.active === false) msgs.push(reason && reason.indexOf('등반') === 0 ? reason + ': 졸업' : '출석 → ' + (p.leftType === '졸업' ? '졸업' : '제적'));
    else if (old.active === false && p.active === false && (old.leftType || '제적') !== (p.leftType || '제적')) msgs.push((old.leftType || '제적') + ' → ' + (p.leftType || '제적'));
    if (old.active === false && p.active !== false) msgs.push((old.leftType || '제적') + ' → 출석으로 되돌림');
    if (old.name !== p.name) msgs.push('이름 변경: ' + old.name + ' → ' + p.name);
    if ((old.role || '') !== (p.role || '') && p.group === 'T') msgs.push('직분: ' + (old.role || '없음') + ' → ' + (p.role || '없음'));
  }
  const rows = msgs.map(m => [p.id, p.name, p.team, kind, m]);
  if (sink) rows.forEach(r => sink.push(r)); else logRows_(auth, rows);
}
function logHomeroom_(auth, team, group, prev, next) {
  prev = prev || { teacher: '', teachers: {}, sections: [] }; next = next || { teacher: '', teachers: {}, sections: [] };
  const R = readRoster_(true); const place = sec => TEAMS[team].label.replace('교육팀', '') + ' ' + groupLabel_(team, group) + (sec ? ' ' + sec + '반' : '');
  const findT = name => R.list.find(p => p.team === team && p.group === 'T' && p.name === name) || R.list.find(p => p.team === team && p.group === 'T' && sameName_(p.name, name));
  const rows = [];
  // 담임은 여러 명일 수 있어요 ("김은혜, 박요한"). 빠진 사람은 해제, 새로 들어온 사람은 지정으로 적어요.
  const names = v => String(v || '').split(/\s*[,，]\s*/).map(x => x.trim()).filter(Boolean);
  const cmp = (a, b, sec) => {
    const A = names(a), B = names(b); const y = new Date().getFullYear();
    A.filter(x => B.indexOf(x) < 0).forEach(x => { const t = findT(x); rows.push([t ? t.id : '', x, team, '담임', y + '년 ' + place(sec) + ' 담임 해제']); });
    B.filter(x => A.indexOf(x) < 0).forEach(x => { const t = findT(x); rows.push([t ? t.id : '', x, team, '담임', y + '년 ' + place(sec) + ' 담임 지정']); });
  };
  cmp(prev.teacher, next.teacher, '');
  const secs = {}; Object.keys(prev.teachers || {}).concat(Object.keys(next.teachers || {})).forEach(k => secs[k] = 1);
  Object.keys(secs).forEach(sec => cmp((prev.teachers || {})[sec], (next.teachers || {})[sec], sec));
  const ps = (prev.sections || []).join(','), ns = (next.sections || []).join(',');
  if (ps !== ns) rows.push(['', '', team, '반 설정', place('') + ' 분반: ' + (ps ? ps.replace(/,/g, '·') + '반' : '없음') + ' → ' + (ns ? ns.replace(/,/g, '·') + '반' : '없음')]);
  logRows_(auth, rows);
}
/** 임원 이상. id를 주면 그 사람(교사면 이름이 같은 담임 기록 포함)만. */
function history_(auth, team, id) {
  if (!TEAMS[team]) throw new Error('부서를 알 수 없어요');
  if (!deptLevel_(auth, team)) throw new Error('이력은 임원 이상만 볼 수 있어요');
  const sh = ss_().getSheetByName(SHEET.hist); const n = sh ? sh.getLastRow() : 0;
  const label = TEAMS[team].label; const R = readRoster_(true); const person = id ? R.byId[id] : null;
  const mine = {}; R.list.forEach(p => { if (p.team === team) mine[p.id] = 1; });
  const hit = id ? (x => x.id === id || (person && person.group === 'T' && x.kind === '담임' && !x.id && x.name === person.name && x.team === label))
    : (x => x.team === label || mine[x.id]);
  // 시트 맨 아래(최신)부터 2000줄씩 거꾸로 읽다가 1000건이 모이면 멈춘다 (기록이 쌓여도 느려지지 않게)
  const out = [];
  for (let end = n; end >= 2 && out.length < 1000; end -= 2000) {
    const start = Math.max(2, end - 1999); const rows = sh.getRange(start, 1, end - start + 1, 7).getValues();
    for (let i = rows.length - 1; i >= 0 && out.length < 1000; i--) {
      const r = rows[i]; const x = { at: String(r[0]), id: String(r[1] || ''), name: String(r[2] || ''), team: String(r[3] || ''), kind: String(r[4] || ''), text: String(r[5] || ''), by: String(r[6] || '') };
      if (hit(x)) out.push(x);
    }
  }
  return out;
}

/**
 * 총관리자·공동관리자1: 모든 부서의 활동 기록(이력 탭)을 최근 순으로 200건씩. 시트 맨 아래가 최신. (공동관리자2는 CO2_BLOCK에서 막힘)
 * upto: 처음 불러올 때의 마지막 줄 번호(그 뒤 새로 쌓인 줄 때문에 다음 묶음이 겹치지 않게), before: 이 줄 번호 위부터 읽는다
 * f: { kind: 'all' | 'nologin'(로그인 빼고) | 종류, team: 부서 키, q: 검색어 }. 거른 결과가 200건이 될 때까지 시트 전체를 거슬러 찾는다
 */
function activity_(auth, before, upto, f) {
  if (!isAdmin_(auth) || auth.co2) throw new Error('활동 기록은 총관리자·공동관리자1만 볼 수 있어요');
  f = f || {}; const limit = 200;
  const sh = ss_().getSheetByName(SHEET.hist); const last = sh ? sh.getLastRow() : 1;
  const top = Number(upto) > 0 ? Math.min(last, Math.floor(Number(upto))) : last;
  let end = Number(before) > 0 ? Math.min(top, Math.floor(Number(before)) - 1) : top;   // 시트 줄 번호 (end까지 포함)
  const kind = String(f.kind || 'all'), lab = TEAMS[f.team] ? TEAMS[f.team].label : '', q = String(f.q || '').trim();
  const hit = x => (kind === 'all' || (kind === 'nologin' ? x.kind !== '로그인' : x.kind === kind)) && (!lab || x.team === lab)
    && (!q || [x.name, x.text, x.by, x.kind].some(v => v.indexOf(q) >= 0));
  const list = []; const t0 = Date.now();
  // 2000줄씩 거꾸로 읽는다. 오래 걸리면(20초) 거기까지만 돌려주고 [더 보기]로 이어 간다
  while (end >= 2 && list.length < limit && Date.now() - t0 < 20000) {
    const start = Math.max(2, end - 1999); const rows = sh.getRange(start, 1, end - start + 1, 7).getValues();
    let i = rows.length - 1;
    for (; i >= 0 && list.length < limit; i--) {
      const r = rows[i]; const x = { at: String(r[0]), id: String(r[1] || ''), name: String(r[2] || ''), team: String(r[3] || ''), kind: String(r[4] || ''), text: String(r[5] || ''), by: String(r[6] || '') };
      if (hit(x)) list.push(x);
    }
    end = start + i;   // 다음에 읽을 줄 (i+1번째 줄은 이미 봤다)
  }
  const kinds = TG_KINDS.map(x => x[0]).filter(k => k !== '기타').concat(['생일']);   // 종류 고르기 칸에 늘 보이는 것
  return { list, upto: top, before: end + 1, more: end >= 2, total: Math.max(0, top - 1), kinds };
}

/* ---------------- 자료실 (일정과 별개인 부서 공유 파일) ---------------- */
function readLib_() {
  const sh = ss_().getSheetByName(SHEET.lib); const n = sh.getLastRow(); const list = [];
  if (n >= 2) sh.getRange(2, 1, n - 1, 8).getValues().forEach((r, i) => {
    const id = String(r[0] || '').trim(); if (!id) return;
    list.push({ row: i + 2, id, team: TEAM_BY_LABEL[String(r[1]).trim()] || '', name: String(r[2] || ''), vis: String(r[3]) === '관리자만' ? 'admin' : 'all', size: Number(r[4]) || 0, mime: String(r[5] || ''), updated: String(r[6] || ''), by: String(r[7] || '') });
  });
  return { sh, list };
}
const libOut_ = f => ({ id: f.id, name: f.name, vis: f.vis, size: f.size, mime: f.mime, updated: f.updated, by: f.by });
/** 자료실 편집: 부서에 들어온 사람(담임 포함)이 올리고 바꾸고 지울 수 있다. 임원 여부를 돌려준다. */
function requireLibEdit_(auth, team) {
  if (!TEAMS[team]) throw new Error('부서를 알 수 없어요');
  const lv = scope_(auth, team).level;
  if (lv === 'none') throw new Error(TEAMS[team].label + ' 번호가 필요해요');
  if (lv === 'sub') throw new Error('보조교사는 일정·자료실을 볼 수만 있어요');
  return deptLevel_(auth, team);
}
function libList_(auth, team) {
  requireView_(auth, team);
  const admin = deptLevel_(auth, team);
  const docs = admin ? (cfgAll_().libDocs || {}) : {};
  const me = admin ? '' : actorLabel_(auth, team); // 담임: 자기가 올린 파일만 지울 수 있어서 '내 파일' 표시
  return { canEdit: scope_(auth, team).level !== 'sub', admin, files: cachedLib_().filter(f => f.team === team && (admin || f.vis !== 'admin')).map(f => Object.assign(libOut_(f), docs[f.id] ? { doc: docs[f.id] } : {}, me && f.by === me ? { mine: true } : {})).reverse() };
}
/** file: { name, mime, b64, vis } 한 파일 10MB까지. 담임은 '전체 보기' 파일만 */
function libUpload_(auth, team, file) {
  const admin = requireLibEdit_(auth, team);
  file = file || {};
  const b64 = String(file.b64 || ''); if (!b64) throw new Error('파일이 비어 있어요');
  if (b64.length > 14 * 1024 * 1024) throw new Error('파일은 10MB까지 올릴 수 있어요');
  const L = readLib_();
  const name = String(file.name || '문서').slice(0, 120);
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), String(file.mime || 'application/octet-stream'), name);
  const f = eventsFolder_(team).createFile(blob);
  const vis = !admin ? 'all' : (file.vis === 'admin' ? 'admin' : 'all');
  const row = [f.getId(), TEAMS[team].label, safeCell_(name), vis === 'admin' ? '관리자만' : '전체', f.getSize(), blob.getContentType(), Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'), actorLabel_(auth, team)];
  const r = L.sh.getLastRow() + 1; L.sh.getRange(r, 1, 1, 8).setValues([row]);
  logRows_(auth, [['', '', team, '자료실', '파일 올림: ' + name + (vis === 'admin' ? ' (관리자만 보기)' : '')]]);
  return libList_(auth, team).files;
}
function libDelete_(auth, team, fileId) {
  const admin = requireLibEdit_(auth, team);
  const L = readLib_(); const f = L.list.find(x => x.team === team && x.id === fileId);
  if (f && f.vis === 'admin' && !admin) throw new Error('관리자만 볼 수 있는 파일은 지울 수 없어요');
  if (f && !admin && f.by !== actorLabel_(auth, team)) throw new Error('다른 사람이 올린 파일은 임원만 지울 수 있어요');
  if (f) {
    logRows_(auth, [['', '', team, '자료실', '파일 삭제: ' + f.name]]);
    try { DriveApp.getFileById(f.id).setTrashed(true); } catch (e) {} L.sh.deleteRow(f.row);
    const docs = cfgAll_().libDocs || {};
    if (docs[f.id]) { try { DriveApp.getFileById(docs[f.id].id).setTrashed(true); } catch (e) {} delete docs[f.id]; cfgSet_('libDocs', docs); }
  }
  return libList_(auth, team).files;
}
/** 자료실 파일을 메일로 보낸다. m: { fileId, to, msg } 5명까지. 온라인 작성본이 있으면 그 최신 내용을 보낸다 */
function libMail_(auth, team, m) {
  requireView_(auth, team); m = m || {};
  if (!deptLevel_(auth, team)) throw new Error('메일 보내기는 임원 이상만 할 수 있어요');
  const f = readLib_().list.find(x => x.team === team && x.id === String(m.fileId || ''));
  if (!f) throw new Error('파일을 찾을 수 없어요');
  const admin = deptLevel_(auth, team);
  if (f.vis === 'admin' && !admin) throw new Error('관리자만 볼 수 있는 파일이에요');
  const list = String(m.to || '').split(/[\s,;，]+/).filter(Boolean);
  if (!list.length) throw new Error('받는 메일 주소를 적어 주세요');
  if (list.length > 5) throw new Error('한 번에 5명까지 보낼 수 있어요');
  list.forEach(a => { if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(a)) throw new Error('"' + a + '"은(는) 올바른 메일 주소가 아니에요'); });
  let blob, name;
  try { const r = libFileBlob_(f); blob = r.blob.setName(r.name); name = r.name; } catch (e) { throw new Error('파일을 열 수 없어요. 자료실에 다시 올려 주세요'); }
  if (MailApp.getRemainingDailyQuota() < list.length) throw new Error('오늘 보낼 수 있는 메일 수를 다 썼어요');
  MailApp.sendEmail({ to: list.join(','), name: '신당중앙교회 교육팀 출석부', subject: '[' + TEAMS[team].label + '] ' + name,
    body: ((String(m.msg || '').slice(0, 500)) + '\n\n— 신당중앙교회 ' + TEAMS[team].label + ' 출석부 자료실에서 보냈어요 —').trim(), attachments: [blob] });
  logRows_(auth, [['', '', team, '자료실', '메일 발송: ' + name + ' → ' + list.join(', ')]]);
  return { to: list.join(', '), name };
}

function libVis_(auth, team, fileId, vis) {
  requirePromoter_(auth, team);
  const L = readLib_(); const f = L.list.find(x => x.team === team && x.id === fileId); if (!f) throw new Error('파일을 찾을 수 없어요');
  L.sh.getRange(f.row, 4).setValue(vis === 'admin' ? '관리자만' : '전체');
  logRows_(auth, [['', '', team, '자료실', '공개 범위 변경: ' + f.name + ' → ' + (vis === 'admin' ? '관리자만 보기' : '전체 보기')]]);
  return libList_(auth, team).files;
}

/* ---------------- 보안: 자동 백업 · 공유 점검 (편집기에서 한 번씩 실행) ---------------- */
/**
 * [백업] 매주 월요일 새벽 3시에 자동으로 백업해요. 트리거는 ensureBackupTrigger_ 가 알아서 만들고, 처음 켜질 때와 실패할 때 텔레그램으로 알려요.
 *   내 드라이브의 "교육팀 출석부 백업" 폴더에 시트 사본이 쌓이고, 가장 최근 BACKUP_KEEP 개만 남기고 지워요.
 *   처음 실행할 때 권한 허용 창이 뜨면 허용해 주세요.
 */
const BACKUP_KEEP = 5;
function backupFolder_() {
  const name = '교육팀 출석부 백업';
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}
function backupNow() {
  try {
    const stamp = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm');
    const folder = backupFolder_();
    DriveApp.getFileById(ss_().getId()).makeCopy('출석부 백업 ' + stamp, folder);
    const n = trimBackups_(folder);
    Logger.log('백업했어요: 출석부 백업 ' + stamp + ' (보관 ' + n + '개)');
    return n;
  } catch (e) {
    tgNotify_('⚠️ 출석부 시트 자동백업 실패: ' + e.message);
    throw e;
  }
}
/** 백업 폴더에서 가장 최근 BACKUP_KEEP 개만 남기고 휴지통으로 보낸다. 남은 개수와 지운 개수를 돌려준다 */
function trimBackups_(folder) {
  const list = []; const it = folder.getFiles();
  while (it.hasNext()) { const f = it.next(); if (f.getName().indexOf('출석부 백업 ') === 0) list.push(f); }
  list.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  const old = list.slice(BACKUP_KEEP); old.forEach(f => f.setTrashed(true));
  trimBackups_.removed = old.length;
  return Math.min(list.length, BACKUP_KEEP);
}
/** BACKUP_KEEP 을 바꾼 뒤 다음 백업을 기다리지 않고 한 번 바로 정리한다 (ping 에서 부름) */
function trimBackupsIfKeepChanged_() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('BACKUP_KEPT') === String(BACKUP_KEEP)) return;
  const lock = LockService.getScriptLock(); if (!lock.tryLock(0)) return;
  try {
    const n = trimBackups_(backupFolder_());
    p.setProperty('BACKUP_KEPT', String(BACKUP_KEEP));
    tgNotify_('🧹 출석부 백업을 최근 ' + BACKUP_KEEP + '개만 남기고 정리했어요 (휴지통으로 ' + trimBackups_.removed + '개, 남은 ' + n + '개)');
  } catch (e) {
    tgNotify_('⚠️ 출석부 백업 정리 실패: ' + e.message);
  } finally { lock.releaseLock(); }
}
function setBackupTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'backupNow') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('backupNow').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(3).create();
  PropertiesService.getScriptProperties().setProperty('BACKUP_DAILY', 'weekly');
}
/** 매주 월요일 백업 트리거가 없으면 (예전 매일 트리거('1')면 주간으로 바꿔 단다) 만든다. 편집기에서 직접 실행하지 않아도 되게 ping·tgPoll 에서 부른다. removeBackupTrigger 로 끄면 다시 켜지 않는다 */
function ensureBackupTrigger_() {
  const c = CacheService.getScriptCache(), key = 'backupChk'; if (c.get(key)) return;
  c.put(key, '1', 21600);
  const lock = LockService.getScriptLock(); if (!lock.tryLock(0)) { c.remove(key); return; }
  try {
    const v = PropertiesService.getScriptProperties().getProperty('BACKUP_DAILY');
    if (v === 'off') return;
    if (v === 'weekly' && ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'backupNow')) return;
    setBackupTrigger_();
    const n = backupNow();
    tgNotify_('✅ 출석부 시트 자동백업을 켰어요. 매주 월요일 새벽 3시, 드라이브 \'교육팀 출석부 백업\' 폴더 (지금 ' + n + '개, 최대 ' + BACKUP_KEEP + '개)');
  } catch (e) {
    tgNotify_('⚠️ 출석부 시트 자동백업을 켜지 못했어요: ' + e.message);
  } finally { lock.releaseLock(); }
}
function installBackupTrigger() {
  setBackupTrigger_();
  backupNow();
  Logger.log('매주 월요일 새벽 3시 자동 백업을 켰어요. 끄려면 removeBackupTrigger 를 실행하세요.');
}
function removeBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'backupNow') ScriptApp.deleteTrigger(t); });
  PropertiesService.getScriptProperties().setProperty('BACKUP_DAILY', 'off');
  Logger.log('자동 백업을 껐어요.');
}

/**
 * [백업에서 출석 되살리기] restoreFromBackup 을 실행하면 지난 주일(오늘이 주일이면 오늘) 하루치를 백업에서 되살려요.
 *   - '교육팀 출석부 백업' 폴더에서 그날 기록이 가장 많은 백업을 골라요.
 *   - 출석·달란트는 지금 비어 있는 칸만 채우고, 출결메모는 없는 것만 더해요 (그 뒤에 새로 체크한 것은 그대로 둬요).
 *   - 되살리기 전에 지금 상태를 '출석부 복구 전 …' 이름으로 같은 폴더에 복사해 둬요 (자동 백업 정리에서 지워지지 않아요).
 * 다른 날짜를 되살리려면 RESTORE_DATE 에 'yyyy-MM-dd' 를 적고 실행하세요.
 */
const RESTORE_DATE = '';
function restoreFromBackup() {
  const tz = tz_();
  let key = RESTORE_DATE;
  if (!key) { const d = new Date(); d.setDate(d.getDate() - Number(Utilities.formatDate(d, tz, 'u')) % 7); key = Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }
  // 백업 파일마다 그날 칸을 읽어서 기록이 가장 많은 것을 고른다
  const readDay = (ss) => {
    const out = { cnt: 0 };
    [SHEET.att, SHEET.tal].forEach(name => {
      const sh = ss.getSheetByName(name); const m = out[name] = {}; if (!sh) return;
      const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn(); if (lastRow < 2 || lastCol < WEEK_COL) return;
      const v = sh.getRange(1, 1, lastRow, lastCol).getValues();
      const col = v[0].findIndex((h, c) => c >= WEEK_COL - 1 && dateKey_(h) === key); if (col < 0) return;
      for (let r = 1; r < v.length; r++) { const id = String(v[r][0] || '').trim(), x = v[r][col]; if (id && x !== '' && x !== null) { m[id] = x; out.cnt++; } }
    });
    const msh = ss.getSheetByName(SHEET.memo); out.memo = [];
    if (msh && msh.getLastRow() >= 2) msh.getRange(2, 1, msh.getLastRow() - 1, 6).getValues().forEach(r => { if (dateKey_(r[0]) === key && String(r[1] || '')) out.memo.push(r); });
    return out;
  };
  const folder = backupFolder_(); const it = folder.getFiles(); let best = null;
  while (it.hasNext()) {
    const f = it.next(); if (f.getName().indexOf('출석부 백업 ') !== 0) continue;
    const d = readDay(SpreadsheetApp.openById(f.getId()));
    Logger.log(f.getName() + ' : ' + key + ' 기록 ' + d.cnt + '칸, 메모 ' + d.memo.length + '개');
    if (d.cnt + d.memo.length > 0 && (!best || d.cnt + d.memo.length > best.d.cnt + best.d.memo.length)) best = { f, d };
  }
  if (!best) { Logger.log(key + ' 기록이 있는 백업이 없어요. 시트의 [파일 → 버전 기록]에서 되살려 주세요.'); return; }
  const now = readDay(ss_());
  Logger.log('지금 시트 : ' + key + ' 기록 ' + now.cnt + '칸, 메모 ' + now.memo.length + '개');
  DriveApp.getFileById(ss_().getId()).makeCopy('출석부 복구 전 ' + Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'), folder);
  let filled = 0, memos = 0;
  withLock_(() => {
    const R = readRoster_(true);
    [SHEET.att, SHEET.tal].forEach(name => {
      const src = best.d[name]; const ids = Object.keys(src).filter(id => R.byId[id]); if (!ids.length) return;
      const W = readWeeklyIdx_(name); const col = weeklyCol_(W, key);
      ids.forEach(id => weeklyRow_(W, R.byId[id]));
      const last = W.sh.getLastRow(); const rng = W.sh.getRange(2, col, last - 1, 1); const cur = rng.getValues();
      ids.forEach(id => { const at = W.rows[id] - 2; if (cur[at] && (cur[at][0] === '' || cur[at][0] === null)) { cur[at][0] = src[id]; filled++; } });
      rng.setValues(cur);
      ids.forEach(id => touchGroup_(R.byId[id].team, R.byId[id].group));
    });
    const have = {}; now.memo.forEach(r => have[String(r[1])] = 1);
    const add = best.d.memo.filter(r => !have[String(r[1])]);
    if (add.length) {
      const sh = ss_().getSheetByName(SHEET.memo); const r0 = sh.getLastRow() + 1;
      fmt_(sh.getRange(r0, 1, add.length, 2), '@');
      sh.getRange(r0, 1, add.length, 6).setValues(add.map(r => [key, String(r[1]), r[2], r[3], r[4], r[5]]));
      memos = add.length;
    }
    touchVer_('g');
  });
  Logger.log(best.f.getName() + ' 에서 ' + key + ' 출석·달란트 ' + filled + '칸, 메모 ' + memos + '개를 되살렸어요.');
}

/**
 * [공유 점검] auditSharing 을 실행하면 로그에 아래 내용이 나와요. 읽기만 하고 아무것도 바꾸지 않아요.
 *   - 출석부 시트를 누가 열 수 있는지 (링크 공유 여부, 편집자·보기 권한자 메일)
 *   - 공유문서 폴더 안에서 "링크가 있으면 누구나" 열리는 파일 (온라인 작성본)
 */
function auditSharing() {
  const A = DriveApp.Access, P = DriveApp.Permission;
  const accName = a => a === A.PRIVATE ? '나와 지정한 사람만' : a === A.DOMAIN ? '도메인 전체' : a === A.DOMAIN_WITH_LINK ? '도메인(링크)' : a === A.ANYONE ? '인터넷 누구나' : a === A.ANYONE_WITH_LINK ? '링크가 있으면 누구나' : String(a);
  const permName = p => p === P.EDIT ? '편집' : p === P.VIEW ? '보기' : p === P.COMMENT ? '댓글' : String(p);
  const f = DriveApp.getFileById(ss_().getId());
  const out = [];
  out.push('■ 출석부 시트: ' + f.getName());
  out.push('  소유자: ' + (f.getOwner() ? f.getOwner().getEmail() : '(알 수 없음)'));
  out.push('  링크 공유: ' + accName(f.getSharingAccess()) + (f.getSharingAccess() === A.PRIVATE ? ' (안전)' : ' / 권한 ' + permName(f.getSharingPermission()) + '  ← 꺼 두는 게 안전해요'));
  out.push('  편집자 ' + f.getEditors().length + '명: ' + (f.getEditors().map(u => u.getEmail()).join(', ') || '없음'));
  out.push('  보기 권한자 ' + f.getViewers().length + '명: ' + (f.getViewers().map(u => u.getEmail()).join(', ') || '없음'));
  const roots = DriveApp.getFoldersByName('교육팀 출석부 공유문서');
  if (roots.hasNext()) {
    const open = [];
    const walk = folder => {
      const fi = folder.getFiles();
      while (fi.hasNext()) { const x = fi.next(); if (x.getSharingAccess() !== A.PRIVATE) open.push(x.getName() + ' (' + accName(x.getSharingAccess()) + ', ' + permName(x.getSharingPermission()) + ')'); }
      const fo = folder.getFolders(); while (fo.hasNext()) walk(fo.next());
    };
    walk(roots.next());
    out.push('■ 공유문서 폴더: 링크로 열리는 파일 ' + open.length + '개' + (open.length ? '\n  - ' + open.join('\n  - ') : ''));
    if (open.length) out.push('  (온라인 작성본은 편집을 위해 링크 공유가 켜져 있어요. 링크를 밖으로 알리지 말고, 더 이상 안 쓰는 작성본은 자료실에서 삭제하세요.)');
  } else out.push('■ 공유문서 폴더가 아직 없어요 (자료실을 쓰기 전이에요).');
  Logger.log(out.join('\n'));
}

/* ---------------- 당회보고 (임원 이상) ----------------
 * 달마다 부서별 당회보고(이번 달 보고 / 다음 달 계획)를 모아 '○년 ○월 당회.제직회자료.hwp' 한 파일로 합친다.
 *  - 임원: 우리 부서 보고서를 올리고(.hwp → 화면에서 글자를 꺼내 보냄) 고친다. 다른 부서는 보기만.
 *  - 관리자(총관리자·공동관리자): 모든 부서와 '교육지원' 칸을 고치고, 날짜·제목을 정하고, 메일로 보낸다.
 *  - 내용은 '당회보고' 탭에 달마다 한 줄(JSON)로, 올린 원본 .hwp 는 공유문서 폴더의 '당회보고' 폴더에 둔다.
 */
const RPT_SHEET_ = '당회보고';
const RPT_ROWS_ = TEAM_ORDER.concat(['support']);
const RPT_HEADING_ = '제1교육위원회 [교회학교 파트]';
const RPT_DEFAULT_TO_ = [{ name: '조주영 간사님', email: 'sdcc1945@hanmail.net' }];
/** 메일 '보내는 사람' 목록에 늘 나오는 사람 (최근 보낸 사람 뒤에) */
const RPT_DEFAULT_FROM_ = [{ name: '김도훈(지원팀)', email: 'bo730129@nate.com' }];
/** 당회보고는 총관리자·공동관리자와 임원(어느 부서든 임원 번호로 들어온 사람)가 본다. 고치기는 자기 부서만 (rptCanEdit_) */
function rptViewer_(auth) { if (isAdmin_(auth)) return true; const codes = codes_(); return TEAM_ORDER.some(t => deptLevel_(auth, t, codes)); }
function requireRptView_(auth) { if (!rptViewer_(auth)) throw new Error('당회보고는 임원 이상만 볼 수 있어요'); }
function rptCanEdit_(auth, team, codes) { return isAdmin_(auth) || (team !== 'support' && deptLevel_(auth, team, codes)); }
function rptYm_(ym) { ym = String(ym || ''); if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(ym)) throw new Error('달 형식이 맞지 않아요'); return ym; }
/** 그달 넷째 주일 (YYYY-MM-DD) */
function fourthSunday_(ym) {
  const y = Number(ym.slice(0, 4)), m = Number(ym.slice(5, 7));
  const first = 1 + (7 - new Date(y, m - 1, 1).getDay()) % 7;
  return ym + '-' + String(first + 21);
}
function rptSheet_() {
  const ss = ss_(); let sh = ss.getSheetByName(RPT_SHEET_);
  if (!sh) { sh = ss.insertSheet(RPT_SHEET_); sh.getRange(1, 1, 1, 3).setValues([['달', '내용(JSON)', '고친 시각']]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
/** 앱을 쓰기 전 달의 자료: 부서별 원본 보고서(.hwp)에서 꺼낸 내용. 그달을 처음 저장할 때 시트에 함께 들어간다 */
const RPT_SEED_ = { '2026-09': { date: '2026-09-27', at: '2026-09-18', by: '9월 원본 보고서', teams: {"infant": {"cur": ["1. 나의 버스 만들기 (9/6)", "2. 기도 손 만들기 (9/20)", "3. 말씀 암송, 생일 축하 (9/27)"], "next": ["1. 예수님의 어린 양- 대그룹 활동 (10/4)", "2. 요리 활동 (10/11)", "3. 추수감사절 – 하나님께 드려요 (10/18)", "3. 말씀 암송, 생일 축하 (10/25)"]}, "toddler": {"cur": ["1. 함께 자라가는 어린이", "  / 말씀사랑 달력(8월 시상, 9월 배부)", "  / 우리 마을 지도판 만들기", "", "2. 예수님이 오신 마을", "  / 하트 안경 만들기(예수님처럼 축복하기 게임)", "", "3. 생명의 말씀, 성경", "  / 카드 성경책 만들기", "", "4. 구약, 하나님과의 특별한 약속", "  / 언약 마음 손 만들기", "  / 생일잔치, 말씀암송 대회", "  / 교사 월례회"], "next": ["1. 신약, 새로운 약속", "  / 말씀사랑 달력(9월 시상, 10월 배부)", "  / 신약성경 지도 그리기", "", "2. 함께 듣는 하나님의 말씀", "  / 공과 활동 : 말씀사랑 하나님 사랑", "", "3. 하나님의 사람으로 온전하게 하는 성경", "  / 신체활동 : 말씀 먹고 자라가요!!", "", "4. 오직 믿음으로 살아가는 우리(종교개혁주일)", "  / 생일잔치, 말씀암송 대회", "  / 교사 월례회"]}, "youth": {"cur": ["9/6 : 성령강림주일후 열다섯 번째 주일", "  2부활동 – 9월 생일 축하 / 반별모임 / 교역자와 함께하는 식사(2학년)", "", "9/13 : 성령강림주일후 열여섯 번째 주일", "  2부활동 – 반별모임 / 교역자와 함께하는 식사(3학년)", "", "9/19 : 교육지원팀 교사위로회", "", "9/20 : 성령강림주일후 열일곱 번째 주일", "  2부활동 – 반별모임 / 교역자와 함께하는 식사(4학년)", "", "9/27 : 성령강림주일후 열여덟 번째 주일", "  2부활동 – 반별모임 / 교역자와 함께하는 식사(5, 6학년)", "  교사 – 9월 월례회"], "next": ["10/4 : 성령강림주일후 열아홉 번째 주일", "  가을소풍 – 용산전쟁기념관", "", "10/11 : 성령강림주일후 스무 번째 주일", "  2부활동 – 10월 생일 축하", "  제자학교 1", "", "10/18 : 성령강림주일후 스물한 번째 주일, 추수감사주일", "  예배 – 추수감사주일 예배", "  제자학교 2", "", "10/25 : 성령강림주일후 스물두 번째 주일, 종교개혁주일", "  예배 – 종교개혁주일 예배", "  교사 – 10월 월례회", "  제자학교 3"]}, "teen": {"cur": ["9/5 (토) 중1반 단합대회 (야외활동)", "", "9/6 (주일) 오순절 후 제15주", "  - 예배 후 반별 모임과 나눔", "  - 세례교육 2주차 (선교관 304호)", "", "9/13 (주일) 오순절 후 제16주", "  - 예배 후 반별 모임과 나눔", "  - 세례교육 3주차 (선교관 304호)", "", "9/19 (토) 제1교육위원회 교사 위로회 참석", "", "9/20 (주일) 오순절 후 제17주", "  - 예배 후 반별 모임과 나눔", "  - 세례교육 4주차 (선교관 304호)", "", "9/27 (주일) 오순절 후 제18주", "  - 예배 후 반별 모임과 나눔", "  - 세례교육 5주차 (선교관 304호)"], "next": ["10/4 (주일) 오순절 후 제19주", "  - 예배 후 반별 모임과 나눔", "  - 세례교육 6주차&세례입교문답 (선교관 304호)", "", "10/11 (주일) 오순절 후 제20주", "  - 예배 전 반별 모임과 나눔(10시)", "  - 세례입교예식으로 대예배 참석 예정", "  - 예배 후 단체활동(풋살 예정)", "", "10/18 (주일) 오순절 후 제21주, 추수감사주일", "  - 예배 전 반별 모임과 나눔(10시)", "  - 성찬식으로 대예배 참석 예정", "", "10/24 (토) 교사모임&교사기도회", "", "10/25 (주일) 오순절 후 제22주, 종교개혁주일", "  - 예배 후 반별 모임과 나눔", "  - 생일자 축하"]}, "support": {"cur": ["• 2026년 하반기 전체교사 수련회", "일시: 9월19일(토) 장소: 영종도", "• 유아교육팀 무선마이크 교체 지원"], "next": ["• 2026년 교회학교 예배 지원"]}} } };
function rptSeed_(ym) {
  const z = RPT_SEED_[ym]; if (!z) return null; const teams = {};
  Object.keys(z.teams).forEach(t => { teams[t] = { cur: z.teams[t].cur.slice(), next: z.teams[t].next.slice(), at: z.at, by: z.by, file: null }; });
  return { date: z.date, teams };
}
function rptRead_(ym) {
  const sh = rptSheet_(); const n = sh.getLastRow(); let row = 0, data = null;
  if (n >= 2) {
    const v = sh.getRange(2, 1, n - 1, 2).getValues();
    for (let i = 0; i < v.length; i++) if (String(v[i][0]).trim() === ym) { row = i + 2; try { data = JSON.parse(v[i][1]); } catch (e) {} break; }
  }
  if (!row) data = rptSeed_(ym);
  data = data || {}; data.teams = data.teams || {};
  return { sh, row, data };
}
function rptWrite_(R, ym, data) {
  if (!R.row) R.row = R.sh.getLastRow() + 1;
  fmt_(R.sh.getRange(R.row, 1), '@');
  R.sh.getRange(R.row, 1, 1, 3).setValues([[ym, JSON.stringify(data), Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm')]]);
}
function rptLines_(a) {
  return (Array.isArray(a) ? a : String(a || '').split('\n')).slice(0, 300).map(x => String(x == null ? '' : x).replace(/[\u0000-\u0008\u000b-\u001f]/g, '').slice(0, 300));
}
function rptFolder_() {
  const root = (() => { const it = DriveApp.getFoldersByName('교육팀 출석부 공유문서'); return it.hasNext() ? it.next() : DriveApp.createFolder('교육팀 출석부 공유문서'); })();
  const it = root.getFoldersByName('당회보고'); return it.hasNext() ? it.next() : root.createFolder('당회보고');
}
function rptTeamName_(t) { return t === 'support' ? '교육지원' : TEAMS[t].label; }
/** 지난달에 쓴 '다음 달 계획' (이번 달 보고 내용으로 이월할 것) */
function rptCarry_(ym) {
  let y = Number(ym.slice(0, 4)), m = Number(ym.slice(5, 7)) - 1; if (m < 1) { m = 12; y--; }
  const prev = rptRead_(y + '-' + (m < 10 ? '0' : '') + m).data; const out = {};
  RPT_ROWS_.forEach(t => { const x = prev.teams[t] || {}; out[t] = (x.next || []).filter((v, i, a) => a.slice(i).some(z => String(z).trim())); });
  return out;
}
/** 부서별 내용. 이번 달 보고 내용이 비어 있으면(안 냈거나, 지웠거나, 계획만 썼거나) 언제나 지난달 계획을 보고 내용으로 이월 */
function rptTeams_(ym, d, carry) {
  carry = carry || rptCarry_(ym); const out = {};
  const has = a => (a || []).some(v => String(v).trim());
  RPT_ROWS_.forEach(t => { const x = d.teams[t] || {}; const c = carry[t] || []; const own = has(x.cur);
    out[t] = { cur: own ? x.cur : c.slice(), next: x.next || [], carried: !own && c.length > 0, carry: c, x }; });
  return out;
}
function rptGet_(auth, ym) {
  requireRptView_(auth); ym = rptYm_(ym); const codes = codes_();
  const d = rptRead_(ym).data; const teams = {}; const T = rptTeams_(ym, d);
  RPT_ROWS_.forEach(t => { const v = T[t], x = v.x;
    teams[t] = { cur: v.cur, next: v.next, carried: v.carried, carry: v.carry, saved: !!d.teams[t] && !x.deleted, at: x.at || '', by: x.by || '', file: x.file ? x.file.name : '', edit: rptCanEdit_(auth, t, codes) }; });
  const admin = isAdmin_(auth); const recent = cfgAll_().rptMail; const froms = cfgAll_().rptFrom;
  const meStaff = admin ? staffList_().find(x => normEmail_(x.email) === normEmail_(auth.email)) : null;
  return { ym, date: d.date || fourthSunday_(ym), heading: d.heading || RPT_HEADING_, teams, admin, sent: admin ? (d.sent || []) : [],
    recent: admin ? (Array.isArray(recent) ? recent : RPT_DEFAULT_TO_) : [],
    froms: admin ? rptFroms_() : [], me: meStaff ? { name: String(meStaff.name || ''), email: normEmail_(meStaff.email) } : null,
    fileName: rptFileName_(ym) };
}
function rptSave_(auth, ym, team, body) {
  requireRptView_(auth); ym = rptYm_(ym);
  if (RPT_ROWS_.indexOf(team) < 0) throw new Error('부서를 알 수 없어요');
  if (!rptCanEdit_(auth, team)) throw new Error(team === 'support' ? '교육지원 칸은 관리자만 고칠 수 있어요' : '우리 부서 보고서만 고칠 수 있어요');
  body = body || {}; touchVer_('x');
  const R = rptRead_(ym); const old = R.data.teams[team] || {};
  const x = { cur: rptLines_(body.cur), next: rptLines_(body.next), at: Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'),
    by: actorLabel_(auth, team === 'support' ? TEAM_ORDER[0] : team), file: old.file || null };
  const file = body.file || null;
  if (file && file.b64) {
    if (String(file.b64).length > 14 * 1024 * 1024) throw new Error('파일은 10MB까지 올릴 수 있어요');
    const name = String(file.name || (rptTeamName_(team) + ' 당회보고.hwp')).slice(0, 120);
    const f = rptFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(file.b64), 'application/x-hwp', name));
    if (old.file && old.file.id) { try { DriveApp.getFileById(old.file.id).setTrashed(true); } catch (e) {} }
    x.file = { id: f.getId(), name };
  }
  R.data.teams[team] = x; rptWrite_(R, ym, R.data);
  logRows_(auth, [['', '', team === 'support' ? '' : team, '당회보고', ym.replace('-', '년 ') + '월 ' + rptTeamName_(team) + (file ? ' 보고서 올림 (' + x.file.name + ')' : ' 보고서 고침')]]);
  return rptGet_(auth, ym);
}
/** 이번 달 부서 보고서 삭제: 내용과 올린 원본을 지운다. 보고 내용은 다시 지난달 계획에서 이월된다 */
function rptDelete_(auth, ym, team) {
  requireRptView_(auth); ym = rptYm_(ym);
  if (RPT_ROWS_.indexOf(team) < 0) throw new Error('부서를 알 수 없어요');
  if (!rptCanEdit_(auth, team)) throw new Error(team === 'support' ? '교육지원 칸은 관리자만 지울 수 있어요' : '우리 부서 보고서만 지울 수 있어요');
  touchVer_('x');
  const R = rptRead_(ym); const old = R.data.teams[team] || {};
  if (old.file && old.file.id) { try { DriveApp.getFileById(old.file.id).setTrashed(true); } catch (e) {} }
  R.data.teams[team] = { cur: [], next: [], deleted: true, at: Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'), by: actorLabel_(auth, team === 'support' ? TEAM_ORDER[0] : team), file: null };
  rptWrite_(R, ym, R.data);
  logRows_(auth, [['', '', team === 'support' ? '' : team, '당회보고', ym.replace('-', '년 ') + '월 ' + rptTeamName_(team) + ' 보고서 삭제']]);
  return rptGet_(auth, ym);
}
function rptMeta_(auth, ym, meta) {
  requireRptView_(auth); if (!isAdmin_(auth)) throw new Error('날짜와 제목은 관리자만 바꿀 수 있어요');
  ym = rptYm_(ym); meta = meta || {}; touchVer_('x');
  const R = rptRead_(ym);
  if (meta.date != null) { const s = String(meta.date); if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('날짜 형식이 맞지 않아요'); R.data.date = s; }
  if (meta.heading != null) R.data.heading = String(meta.heading).trim().slice(0, 60) || RPT_HEADING_;
  rptWrite_(R, ym, R.data);
  return rptGet_(auth, ym);
}
function rptFileName_(ym) { return ym.slice(0, 4) + '년 ' + Number(ym.slice(5, 7)) + '월 당회.제직회자료.hwp'; }
function rptBytes_(ym, d) {
  const date = d.date || fourthSunday_(ym);
  const T = rptTeams_(ym, d);
  const rows = RPT_ROWS_.map(t => ({ team: t === 'support' ? ['교육지원'] : [TEAMS[t].label.replace('교육팀', ''), '교육팀'], cur: T[t].cur.slice(), next: T[t].next.slice() }));
  const title = Number(date.slice(0, 4)) + '년 ' + Number(date.slice(5, 7)) + '월 ' + Number(date.slice(8, 10)) + '일 당회 보고 사항';
  return buildReportHwp_(RPT_TPL_, s => Uint8Array.from(Utilities.base64Decode(s), x => x & 255),
    { heading: d.heading || RPT_HEADING_, title, cols: ['팀', Number(ym.slice(5, 7)) + ' 월 보 고 내 용', '계획, 예정사항'], rows });
}
const signedBytes_ = u8 => Array.from(u8, x => x > 127 ? x - 256 : x);
function rptHwp_(auth, ym) {
  requireRptView_(auth); ym = rptYm_(ym);
  return { name: rptFileName_(ym), b64: Utilities.base64Encode(signedBytes_(rptBytes_(ym, rptRead_(ym).data))) };
}
function rptOrig_(auth, ym, team) {
  requireRptView_(auth); ym = rptYm_(ym);
  const x = rptRead_(ym).data.teams[team] || {}; if (!x.file || !x.file.id) throw new Error('올린 원본 파일이 없어요');
  const f = DriveApp.getFileById(x.file.id);
  return { name: x.file.name, b64: Utilities.base64Encode(f.getBlob().getBytes()) };
}
/** 보내는 사람 목록: 최근 보낸 사람 + 고정 목록(지운 사람은 빼고) */
const rptSameFrom_ = (a, b) => String(a.name || '') === String(b.name || '') && normEmail_(a.email) === normEmail_(b.email);
function rptFroms_() {
  const c = cfgAll_(); const recent = Array.isArray(c.rptFrom) ? c.rptFrom : []; const hide = Array.isArray(c.rptFromHide) ? c.rptFromHide : [];
  return recent.concat(RPT_DEFAULT_FROM_.filter(f => !recent.some(x => rptSameFrom_(x, f)) && !hide.some(x => rptSameFrom_(x, f))));
}
/** 보내는 사람 목록에서 한 사람을 지운다. f = { name, email } */
function rptFromDel_(auth, ym, f) {
  requireRptView_(auth); if (!isAdmin_(auth)) throw new Error('관리자만 지울 수 있어요');
  ym = rptYm_(ym); f = { name: String((f && f.name) || '').trim().slice(0, 30), email: normEmail_(f && f.email) };
  if (!f.name && !f.email) throw new Error('지울 사람을 알 수 없어요');
  const c = cfgAll_(); const recent = Array.isArray(c.rptFrom) ? c.rptFrom : [];
  if (recent.some(x => rptSameFrom_(x, f))) cfgSet_('rptFrom', recent.filter(x => !rptSameFrom_(x, f)));
  if (RPT_DEFAULT_FROM_.some(x => rptSameFrom_(x, f))) {
    const hide = Array.isArray(cfgAll_().rptFromHide) ? cfgAll_().rptFromHide : [];
    if (!hide.some(x => rptSameFrom_(x, f))) cfgSet_('rptFromHide', hide.concat([f]));
  }
  logRows_(auth, [['', '', '', '당회보고', '메일 보내는 사람 목록에서 삭제: ' + (f.name || f.email)]]);
  return rptGet_(auth, ym);
}
/** 받는 사람 목록에서 한 사람을 지운다. t = { email } */
function rptToDel_(auth, ym, t) {
  requireRptView_(auth); if (!isAdmin_(auth)) throw new Error('관리자만 지울 수 있어요');
  ym = rptYm_(ym); const email = normEmail_(t && t.email); if (!email) throw new Error('지울 사람을 알 수 없어요');
  const prev = cfgAll_().rptMail; const list = Array.isArray(prev) ? prev : RPT_DEFAULT_TO_;
  const hit = list.find(x => normEmail_(x.email) === email);
  cfgSet_('rptMail', list.filter(x => normEmail_(x.email) !== email));
  logRows_(auth, [['', '', '', '당회보고', '메일 받는 사람 목록에서 삭제: ' + ((hit && hit.name) || email)]]);
  return rptGet_(auth, ym);
}
function rptMail_(auth, ym, m) {
  requireRptView_(auth); if (!isAdmin_(auth)) throw new Error('메일은 관리자만 보낼 수 있어요');
  ym = rptYm_(ym); m = m || {}; touchVer_('x');
  const to = (Array.isArray(m.to) ? m.to : []).map(x => ({ name: String((x && x.name) || '').trim().slice(0, 30), email: normEmail_(x && x.email) })).filter(x => x.email);
  if (!to.length) throw new Error('받는 사람을 골라 주세요');
  if (to.length > 10) throw new Error('한 번에 10명까지 보낼 수 있어요');
  to.forEach(x => { if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(x.email)) throw new Error('"' + x.email + '"은(는) 올바른 메일 주소가 아니에요'); });
  if (MailApp.getRemainingDailyQuota() < to.length) throw new Error('오늘 보낼 수 있는 메일 수를 다 썼어요');
  const R = rptRead_(ym); const name = rptFileName_(ym);
  const blob = Utilities.newBlob(signedBytes_(rptBytes_(ym, R.data)), 'application/x-hwp', name);
  const who = to.map(x => x.name || x.email).join(', ');
  const msg = String(m.msg || '').trim().slice(0, 1000) || (ym.slice(0, 4) + '년 ' + Number(ym.slice(5, 7)) + '월 당회·제직회 자료를 보내 드립니다.');
  // 보내는 사람(선택): 받는 분에게 이 이름으로 보이고, 답장은 이 주소로 간다. 없으면 '교육팀 출석부'
  const from = m.from && typeof m.from === 'object' ? { name: String(m.from.name || '').trim().slice(0, 30), email: normEmail_(m.from.email) } : null;
  if (from && from.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(from.email)) throw new Error('보내는 사람 메일 주소를 확인해 주세요');
  if (from && !from.name && !from.email) throw new Error('보내는 사람 이름을 적어 주세요');
  const mail = { to: to.map(x => x.email).join(','), name: from && from.name ? from.name + ' (신당중앙교회 교육팀)' : '신당중앙교회 교육팀 출석부', subject: name.replace(/\.hwp$/, ''),
    body: who + '께\n\n' + msg + '\n\n' + (from ? '보낸 사람: ' + (from.name || '') + (from.email ? ' <' + from.email + '>' : '') + '\n' : '') + '— 신당중앙교회 교육팀 출석부에서 보냈어요 —', attachments: [blob] };
  if (from && from.email) mail.replyTo = from.email;
  MailApp.sendEmail(mail);
  if (from) { const hd = cfgAll_().rptFromHide; if (Array.isArray(hd) && hd.some(x => rptSameFrom_(x, from))) cfgSet_('rptFromHide', hd.filter(x => !rptSameFrom_(x, from)));
    const pf = cfgAll_().rptFrom; cfgSet_('rptFrom', [from].concat((Array.isArray(pf) ? pf : []).filter(p => !(normEmail_(p.email) === from.email && String(p.name || '') === from.name))).slice(0, 8)); }
  // 최근 받는 사람: 방금 보낸 사람을 맨 앞으로, 10명까지 기억
  const prev = cfgAll_().rptMail; const list = to.concat((Array.isArray(prev) ? prev : RPT_DEFAULT_TO_).filter(p => !to.some(x => x.email === normEmail_(p.email)))).slice(0, 10);
  cfgSet_('rptMail', list);
  const at = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm');
  R.data.sent = [{ at, to: to.map(x => (x.name ? x.name + ' ' : '') + '<' + x.email + '>').join(', '), from: from ? (from.name || from.email) : '', by: actorLabel_(auth, TEAM_ORDER[0]) }].concat(R.data.sent || []).slice(0, 20);
  rptWrite_(R, ym, R.data);
  logRows_(auth, [['', '', '', '당회보고', '메일 발송: ' + name + (from ? ' (보낸 사람 ' + (from.name || from.email) + ')' : '') + ' → ' + to.map(x => (x.name ? x.name + ' ' : '') + x.email).join(', ')]]);
  return rptGet_(auth, ym);
}

/* ---------------- 한글(.hwp) 파일 만들기: 당회보고 통합본 ----------------
 * 지난 통합본(2026년 9월 당회.제직회자료.hwp)의 서식(글꼴·문단 모양·표)을 틀(RPT_TPL_)로 그대로 쓰고, 글자만 새로 채운다.
 * 모든 부서 칸에 같은 문단 모양·글자 모양을 써서 글꼴·크기·스타일이 하나로 맞춰진다.
 * 압축은 deflate 의 '저장(비압축) 블록'으로 써서 별도 압축 도구 없이도 한글이 그대로 읽는다.
 */
const RPT_STY_ = { title: [0, 7], date: [1, 8], head: [7, 9], team: [15, 9], body: [17, 9] }; // [문단 모양, 글자 모양] (틀의 번호)
function u8cat_(list) { let n = 0; list.forEach(a => n += a.length); const out = new Uint8Array(n); let o = 0; list.forEach(a => { out.set(a, o); o += a.length; }); return out; }
function le_(n, bytes) { const a = new Uint8Array(bytes); for (let i = 0; i < bytes; i++) a[i] = Math.floor(n / Math.pow(256, i)) & 255; return a; }
function utf16_(s) { const a = new Uint8Array(s.length * 2); for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); a[i * 2] = c & 255; a[i * 2 + 1] = c >> 8; } return a; }
function hwpRec_(tag, lv, p) {
  const big = p.length >= 0xFFF;
  return u8cat_([le_((tag | (lv << 10) | ((big ? 0xFFF : p.length) << 20)) >>> 0, 4), big ? le_(p.length, 4) : new Uint8Array(0), p]);
}
/** 문단 하나: 문단 머리 + 글자 + 글자 모양. (줄 배치 정보는 비워 두면 한글이 열 때 다시 계산한다) */
function hwpPara_(lv, text, sty, last) {
  text = String(text || '').replace(/\t/g, '    ').replace(/[\u0000-\u001f]/g, '');
  const n = text.length + 1;
  const head = u8cat_([le_((n | (last ? 0x80000000 : 0)) >>> 0, 4), le_(0, 4), le_(sty[0], 2), le_(0, 1), le_(0, 1), le_(1, 2), le_(0, 2), le_(0, 2), le_(0, 4), le_(0, 2)]);
  const out = [hwpRec_(66, lv, head)];
  if (text) out.push(hwpRec_(67, lv + 1, u8cat_([utf16_(text), le_(13, 2)])));
  out.push(hwpRec_(68, lv + 1, u8cat_([le_(0, 4), le_(sty[1], 4)])));
  return u8cat_(out);
}
/** deflate '저장 블록' (압축하지 않은 deflate). zlib 으로 그대로 풀린다 */
function storedDeflate_(d) {
  const out = []; let o = 0;
  do { const n = Math.min(65535, d.length - o); const fin = o + n >= d.length;
    out.push(new Uint8Array([fin ? 1 : 0]), le_(n, 2), le_(n ^ 0xFFFF, 2), d.subarray(o, o + n)); o += n; } while (o < d.length);
  return u8cat_(out);
}
/**
 * r = { heading, title, cols: [팀 머리글, 이번 달 머리글, 다음 달 머리글], rows: [{ team: ['영아', '교육팀'], cur: [줄...], next: [줄...] }] }
 * 표의 줄 수는 틀과 같다(머리글 + 5줄). 모자라면 빈 줄, 넘치면 마지막 줄에 이어 붙인다.
 */
function buildReportHwp_(T, b64, r) {
  const cells = T.cells.map(b64), nRows = T.rows;
  const rows = r.rows.slice(0, nRows - 1);
  if (r.rows.length > nRows - 1) r.rows.slice(nRows - 1).forEach(x => { const L = rows[rows.length - 1]; L.cur = L.cur.concat([''], x.team.join(' ') + ':', x.cur); L.next = L.next.concat([''], x.team.join(' ') + ':', x.next); });
  while (rows.length < nRows - 1) rows.push({ team: [''], cur: [''], next: [''] });
  const parts = [b64(T.pre), hwpPara_(0, r.heading, RPT_STY_.title, false), hwpPara_(0, r.title, RPT_STY_.date, false), b64(T.tbl)];
  const cell = (i, lines, sty) => {
    lines = lines.length ? lines : [''];
    const h = new Uint8Array(cells[i]); h[0] = lines.length & 255; h[1] = (lines.length >> 8) & 255;
    if (i >= 3) h.set(cells[0].subarray(20, 24), 20); // 칸 높이는 머리글 높이로 (내용이 많으면 한글이 늘려 줌)
    parts.push(hwpRec_(72, 2, h));
    lines.forEach((t, j) => parts.push(hwpPara_(2, t, sty, j === lines.length - 1)));
  };
  r.cols.forEach((t, c) => cell(c, [t], RPT_STY_.head));
  rows.forEach((x, i) => { const b = (i + 1) * 3; cell(b, x.team, RPT_STY_.team); cell(b + 1, x.cur, RPT_STY_.body); cell(b + 2, x.next, RPT_STY_.body); });
  const section = u8cat_(parts);
  // 문서 정보: 커서 위치를 맨 앞으로 (틀 문서의 커서가 새 문서에 없는 곳을 가리키지 않게)
  const di = new Uint8Array(b64(T.di)); { const sz = (di[0] | di[1] << 8 | di[2] << 16 | di[3] << 24) >>> 20; if (sz >= 26) for (let i = 4 + 14; i < 4 + 26; i++) di[i] = 0; }
  const prv = utf16_([r.heading, r.title].concat(...rows.map(x => [x.team.join(' ')].concat(x.cur, x.next))).join('\r\n').slice(0, 1000) + '\u0000');
  return cfbWrite_([
    ['FileHeader', b64(T.fh)], ['DocInfo', storedDeflate_(di)], ['BodyText/Section0', storedDeflate_(section)],
    ['\u0005HwpSummaryInformation', b64(T.sum)], ['PrvText', prv], ['DocOptions/_LinkDoc', b64(T.link)],
    ['Scripts/DefaultJScript', b64(T.js)], ['Scripts/JScriptVersion', b64(T.jsv)],
  ]);
}
/** 복합 문서(CFB, 한글·옛 오피스 파일의 바깥 상자) 쓰기. files: [['경로/이름', bytes], ...] */
function cfbWrite_(files) {
  const SS = 512, MSS = 64, CUT = 4096, END = 0xFFFFFFFE, FREE = 0xFFFFFFFF, FATS = 0xFFFFFFFD;
  // 디렉터리 나무
  const ents = [{ name: 'Root Entry', type: 5, kids: [] }];
  const dirOf = path => { let cur = ents[0]; path.forEach(n => { let e = cur.kids.map(i => ents[i]).find(x => x.name === n && x.type === 1); if (!e) { e = { name: n, type: 1, kids: [] }; ents.push(e); cur.kids.push(ents.length - 1); } cur = e; }); return cur; };
  files.forEach(([path, data]) => { const p = path.split('/'); const d = dirOf(p.slice(0, -1)); ents.push({ name: p[p.length - 1], type: 2, data }); d.kids.push(ents.length - 1); });
  // 미니 스트림 (4096바이트보다 작은 것)
  const mini = [], minifat = [];
  ents.forEach(e => { if (e.type !== 2) return; if (e.data.length >= CUT) return;
    const n = Math.ceil(e.data.length / MSS); e.start = n ? minifat.length : END;
    for (let i = 0; i < n; i++) minifat.push(i === n - 1 ? END : minifat.length + 1);
    const pad = new Uint8Array(n * MSS); pad.set(e.data); mini.push(pad); });
  const ministream = u8cat_(mini);
  // 섹터 배치: [FAT][디렉터리][미니FAT][미니스트림][큰 스트림들]
  const nSec = b => Math.ceil(b / SS);
  const dirBytes = Math.ceil(ents.length / 4) * 512, mfBytes = minifat.length * 4;
  const bigs = ents.filter(e => e.type === 2 && e.data.length >= CUT);
  let body = nSec(dirBytes) + nSec(mfBytes) + nSec(ministream.length) + bigs.reduce((s, e) => s + nSec(e.data.length), 0);
  let nFat = 1; while (nFat * 128 < body + nFat) nFat++;
  const fat = []; for (let i = 0; i < nFat; i++) fat.push(FATS);
  const run = n => { if (!n) return END; const s = fat.length; for (let i = 0; i < n; i++) fat.push(i === n - 1 ? END : fat.length + 1); return s; };
  const dirStart = run(nSec(dirBytes)), mfStart = run(nSec(mfBytes)), msStart = run(nSec(ministream.length));
  bigs.forEach(e => { e.start = run(nSec(e.data.length)); });
  while (fat.length % 128) fat.push(FREE);
  ents[0].start = ministream.length ? msStart : END; ents[0].size = ministream.length;
  // 형제는 이름 길이 → 대문자 순으로 정렬한 이진 나무 (모두 검정)
  const cmp = (a, b) => a.name.length - b.name.length || (a.name.toUpperCase() < b.name.toUpperCase() ? -1 : a.name.toUpperCase() > b.name.toUpperCase() ? 1 : 0);
  ents.forEach(e => { e.l = e.r = e.c = FREE; });
  const tree = ids => { if (!ids.length) return FREE; const m = ids.length >> 1; const id = ids[m]; ents[id].l = tree(ids.slice(0, m)); ents[id].r = tree(ids.slice(m + 1)); return id; };
  ents.forEach(e => { if (e.kids) e.c = tree(e.kids.slice().sort((a, b) => cmp(ents[a], ents[b]))); });
  const dir = new Uint8Array(dirBytes);
  ents.forEach((e, i) => { const o = i * 128; const nm = utf16_(e.name); dir.set(nm.subarray(0, 62), o); dir.set(le_(Math.min(nm.length, 62) + 2, 2), o + 64);
    dir[o + 66] = e.type; dir[o + 67] = 1; dir.set(le_(e.l, 4), o + 68); dir.set(le_(e.r, 4), o + 72); dir.set(le_(e.c, 4), o + 76);
    const size = e.type === 2 ? e.data.length : e.type === 5 ? e.size : 0; const start = e.type === 1 ? 0 : e.start;
    dir.set(le_(start, 4), o + 116); dir.set(le_(size, 4), o + 120); });
  for (let i = ents.length; i < dirBytes / 128; i++) { const o = i * 128; dir.set(le_(FREE, 4), o + 68); dir.set(le_(FREE, 4), o + 72); dir.set(le_(FREE, 4), o + 76); }
  const hdr = new Uint8Array(512);
  hdr.set([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]); hdr.set(le_(0x3E, 2), 24); hdr.set(le_(3, 2), 26); hdr.set(le_(0xFFFE, 2), 28); hdr.set(le_(9, 2), 30); hdr.set(le_(6, 2), 32);
  hdr.set(le_(nFat, 4), 44); hdr.set(le_(dirStart, 4), 48); hdr.set(le_(CUT, 4), 56); hdr.set(le_(minifat.length ? mfStart : END, 4), 60); hdr.set(le_(nSec(mfBytes), 4), 64);
  hdr.set(le_(END, 4), 68); hdr.set(le_(0, 4), 72);
  for (let i = 0; i < 109; i++) hdr.set(le_(i < nFat ? i : FREE, 4), 76 + i * 4);
  const pad = a => { const p = new Uint8Array(nSec(a.length) * SS); p.set(a); return p; };
  const fatBytes = new Uint8Array(fat.length * 4); fat.forEach((v, i) => fatBytes.set(le_(v, 4), i * 4));
  const mfb = new Uint8Array(mfBytes); minifat.forEach((v, i) => mfb.set(le_(v, 4), i * 4));
  return u8cat_([hdr, fatBytes, dir, pad(mfb), pad(ministream)].concat(bigs.map(e => pad(e.data))));
}

/** 한글 틀: 2026년 9월 당회.제직회자료.hwp 에서 뽑은 서식 (글자는 쓰지 않음) */
const RPT_TPL_ = {"fh":"SFdQIERvY3VtZW50IEZpbGUAAAAAAAAAAAAAAAAAAAAAAQEFAQAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==","di":"EACgAQEAAQABAAEAAQABAAEAAAAAAAIAAAAVAAAAEQCABAAAAAAEAAAABAAAAAQAAAAEAAAABAAAAAQAAAAEAAAABgAAAA0AAAADAAAAAwAAAAAAAAAVAAAADgAAAAAAAAAAAAAAAAAAABME0AFhAgB0rby5AgsGAAABAQEBAQUARwB1AGwAaQBtABME8AFhAgAUvNXQAgMGAAABAQEBAQYAQgBhAHQAYQBuAGcAEwTwAWEEADTXPLrgrBW1AgEFBAABAQEBAQQANNc8uuCsFbUTBFADYQUASABZAKys4KwVtQIDBgAAAQEBAQEOAEgAWQBHAG8AdABoAGkAYwAtAEUAeAB0AHIAYQATBNABYQIAdK28uQILBgAAAQEBAQEFAEcAdQBsAGkAbQATBPABYQIAFLzV0AIDBgAAAQEBAQEGAEIAYQB0AGEAbgBnABME8AFhBAA01zy64KwVtQIBBQQAAQEBAQEEADTXPLrgrBW1EwRQA2EFAEgAWQCsrOCsFbUCAwYAAAEBAQEBDgBIAFkARwBvAHQAaABpAGMALQBFAHgAdAByAGEAEwTQAWECAHStvLkCCwYAAAEBAQEBBQBHAHUAbABpAG0AEwTwAWECABS81dACAwYAAAEBAQEBBgBCAGEAdABhAG4AZwATBPABYQQANNc8uuCsFbUCAQUEAAEBAQEBBAA01zy64KwVtRMEUANhBQBIAFkArKzgrBW1AgMGAAABAQEBAQ4ASABZAEcAbwB0AGgAaQBjAC0ARQB4AHQAcgBhABME0AFhAgB0rby5AgsGAAABAQEBAQUARwB1AGwAaQBtABME8AFhAgAUvNXQAgMGAAABAQEBAQYAQgBhAHQAYQBuAGcAEwTwAWEEADTXPLrgrBW1AgEFBAABAQEBAQQANNc8uuCsFbUTBFADYQUASABZAKys4KwVtQIDBgAAAQEBAQEOAEgAWQBHAG8AdABoAGkAYwAtAEUAeAB0AHIAYQATBNABYQIAdK28uQILBgAAAQEBAQEFAEcAdQBsAGkAbQATBPABYQIAFLzV0AILBgQCAgICAgQGAEIAYQB0AGEAbgBnABME8AFhBAA01zy64KwVtQILBgQCAgICAgQEADTXPLrgrBW1EwRQA2EFAEgAWQCsrOCsFbUCCwYEAgICAgIEDgBIAFkARwBvAHQAaABpAGMALQBFAHgAdAByAGEAEwTQAWECAHStvLkCCwYAAAEBAQEBBQBHAHUAbABpAG0AEwTwAWECABS81dACAwYAAAEBAQEBBgBCAGEAdABhAG4AZwATBPABYQQANNc8uuCsFbUCAQUEAAEBAQEBBAA01zy64KwVtRMEUANhBQBIAFkArKzgrBW1AgMGAAABAQEBAQ4ASABZAEcAbwB0AGgAaQBjAC0ARQB4AHQAcgBhABME0AFhAgB0rby5AgsGAAABAQEBAQUARwB1AGwAaQBtABME8AFhAgAUvNXQAgMGAAABAQEBAQYAQgBhAHQAYQBuAGcAEwTwAWEEADTXPLrgrBW1AgEFBAABAQEBAQQANNc8uuCsFbUTBFADYQUASABZAKys4KwVtQIDBgAAAQEBAQEOAEgAWQBHAG8AdABoAGkAYwAtAEUAeAB0AHIAYQAUBIACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAABQEUAMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAEAAAD/////AAAAAP////8AAAAAABQEgAIAAAEBAAAAAAEBAAAAAAEBAAAAAAEBAAAAAAEAAAAAAAAAAAAAAAAAFARQAwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAAAAAAAQAAAP////8AAAD//////wAAAAAAFASAAgAAAQEAAAAAAQEAAAAAAQEAAAAAAQEAAAAAAQEAAAAAAAAAAAAAAAAUBFADAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAABAAAA/////5mZmQD/////AAAAAAAVBKAEAQABAAEAAQABAAEAAQBkZGRkZGRkAAAAAAAAAGRkZGRkZGQAAAAAAAAA6AMAAAAAAAAKCgAAAAAAAAAA/////7KysgAEAAAAAAAVBKAEAAAAAAAAAAAAAAAAAABkZGRkZGRkAAAAAAAAAGRkZGRkZGQAAAAAAAAA6AMAAAAAAAAKCgAAAAAAAAAA/////7KysgABAAAAAAAVBKAEAQABAAEAAQABAAEAAQBkZGRkZGRkAAAAAAAAAGRkZGRkZGQAAAAAAAAA6AMAAAAAAAAKCgAAAAAAAAAA/////7KysgABAAAAAAAVBKAEAAAAAAAAAAAAAAAAAABkZGRkZGRkAAAAAAAAAGRkZGRkZGQAAAAAAAAAhAMAAAAAAAAKCgAAAAAAAAAA/////7KysgABAAAAAAAVBKAEAQABAAEAAQABAAEAAQBfX19fX19f+/v7+/v7+2RkZGRkZGQAAAAAAAAAhAMAAAAAAAAKCgAAAAAAAAAA/////7KysgABAAAAAAAVBKAEAAAAAAAAAAAAAAAAAABfX19fX19f+/v7+/v7+2RkZGRkZGQAAAAAAAAAhAMAAAAAAAAKCgAAAAAAAAAA/////7KysgABAAAAAAAVBKAEAwADAAMAAwADAAMAAwBkZGRkZGRkAAAAAAAAAGRkZGRkZGQAAAAAAAAAYAkAAAQAAAAKCgAAAAAAAAAA/////7KysgABAAAAAAAVBKAEAwADAAMAAwADAAMAAwBkZGRkZGRkAAAAAAAAAGRkZGRkZGQAAAAAAAAAmAgAAAQAAAAKCgAAAAAAAAAA/////7KysgACAAAAAAAVBKAEAwADAAMAAwADAAMAAwBkZGRkZGRkAAAAAAAAAGRkZGRkZGQAAAAAAAAAeAUAAAQAAAAKCgAAAAAAAAAA/////7KysgACAAAAAAAVBKAEAgACAAIAAgACAAIAAgBkZGRkZGRkAAAAAAAAAGRkZGRkZGQAAAAAAAAA6AMAAAAAAAAKCgAAAAAAAAAA/////7KysgAEAAAAAAAVBKAEAgACAAIAAgACAAIAAgBiYmJiYmJiAAAAAAAAAGRkZGRkZGQAAAAAAAAA6AMAAAAAAAAKCgAAAAAAAAAA/////7KysgAEAAAAAAAVBKAEAgACAAIAAgACAAIAAgBkZGRkZGRk+Pj4+Pj4+GRkZGRkZGQAAAAAAAAA6AMAAAAAAAAKCgAAAAAAAAAA/////7KysgAEAAAAAAAVBKAEAgACAAIAAgACAAIAAgBkZGRkZGRkAAAAAAAAAGRkZGRkZGQAAAAAAAAA6AMAAAAAAAAKCgAAAAAAAAAA/////7KysgAGAAAAAAAWBIAAAAAAAAAAAAAWBIAAAQAAAAAAAAAWBIAAAgAAAAAAAAAXBCAODAAAAAAAMgD/////AwBeADEALgAMAQAAAAAyAP////8DAF4AMgAuAAwAAAAAADIA/////wMAXgAzACkADAEAAAAAMgD/////AwBeADQAKQAMAAAAAAAyAP////8EACgAXgA1ACkADAEAAAAAMgD/////BAAoAF4ANgApACwAAAAAADIA/////wIAXgA3AAAAAQAAAAEAAAABAAAAAQAAAAEAAAABAAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAABAAAAAQAAABcEYA4MAAAAAAAyAP////8DAF4AMQAuAAwBAAAAADIA/////wMAXgAyAC4ADAAAAAAAMgD/////AwBeADMAKQAMAQAAAAAyAP////8DAF4ANAApAAwAAAAAADIA/////wQAKABeADUAKQAMAQAAAAAyAP////8EACgAXgA2ACkALAAAAAAAMgD/////AgBeADcAAQABAAAAAQAAAAEAAAABAAAAAQAAAAEAAAABAAAALAEAAAAAMgD/////AgBeADgADAAAAAAAMgD/////AAAMAAAAAAAyAP////8AAAEAAAABAAAAAQAAABcEYA4MAAAAAAAyAP////8DAF4AMQAuAAwBAAAAADIA/////wMAXgAyAC4ADAAAAAAAMgD/////AwBeADMAKQAMAQAAAAAyAP////8DAF4ANAApAAwAAAAAADIA/////wQAKABeADUAKQAMAQAAAAAyAP////8EACgAXgA2ACkALAAAAAAAMgD/////AgBeADcAAQABAAAAAQAAAAEAAAABAAAAAQAAAAEAAAABAAAALAEAAAAAMgD/////AgBeADgADAAAAAAAMgD/////AAAMAAAAAAAyAP////8AAAEAAAABAAAAAQAAABkEoAMMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACgAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAKAAAAAAAAAAGQSgA4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAoAAAAAAAAAAZBKADgAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAoAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAACgAAAAAAAAABkEoAOAAAAAAAAAAAAAAADE9f//AAAAAAAAAACCAAAAAQAAAAQAAAAAAAAAAAAAAAAAAAAAAIIAAAAAAAAAGQSgA4AAAAC4CwAAAAAAAAAAAAAAAAAAAAAAAKAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAoAAAAAAAAAAZBKADgCiAANAHAAAAAAAAAAAAAAAAAAAAAAAAoAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAACgAAAAAAAAABkEoAMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACWAAAAAgAAAAQAAAAAAAAAAAAAAAAAAAAAAJYAAAAAAAAAGQSgAwwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAoAAAAAAAAAAZBKADgCiAAqAPAAAAAAAAAAAAAAAAAAAAAAAAoAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAACgAAAAAQAAABkEoAOAKIAEcBcAAAAAAAAAAAAAAAAAAAAAAACgAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAKAAAAACAAAAGQSgA4AogAZAHwAAAAAAAAAAAAAAAAAAAAAAAKAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAoAAAAAMAAAAZBKADgCiACBAnAAAAAAAAAAAAAAAAAAAAAAAAoAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAACgAAAABAAAABkEoAOAKIAK4C4AAAAAAAAAAAAAAAAAAAAAAACgAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAKAAAAAFAAAAGQSgA4AogAywNgAAAAAAAAAAAAAAAAAAAAAAAKAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAoAAAAAYAAAAZBKADDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAggAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAACCAAAAAAAAABkEoAMMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACgAAAAAAAAAAQAAAAAAAAAAAAIAAAAAAAAAKAAAAAAAAAAGQSgA4AAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAJYAAAAAAAIABAAAAAAAAAAAAAAAAAAAAAAAlgAAAAAAAAAZBKADgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlgAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAACWAAAAAAAAABkEoAOAAAABAAAAAAAAAAAAAAAAAAAAAAAAAACWAAAAAAADAAQAAAAAAAAAAAAAAAAAAAAAAJYAAAAAAAAAGQSgAwQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJYAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAlgAAAAAAAAAZBKADBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlgAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAACWAAAAAAAAABoEAAIDABS81dAArgYATgBvAHIAbQBhAGwAAAASBAEAAAAAABoEoAECAPi8OLsEAEIAbwBkAHkAAAESBAQAAgAAABoEgAIEAByslMYgADEACQBPAHUAdABsAGkAbgBlACAAMQAAAhIEBQACAAAAGgSAAgQAHKyUxiAAMgAJAE8AdQB0AGwAaQBuAGUAIAAyAAADEgQIAAIAAAAaBIACBAAcrJTGIAAzAAkATwB1AHQAbABpAG4AZQAgADMAAAQSBAkAAgAAABoEgAIEAByslMYgADQACQBPAHUAdABsAGkAbgBlACAANAAABRIECgACAAAAGgSAAgQAHKyUxiAANQAJAE8AdQB0AGwAaQBuAGUAIAA1AAAGEgQLAAIAAAAaBIACBAAcrJTGIAA2AAkATwB1AHQAbABpAG4AZQAgADYAAAcSBAwAAgAAABoEgAIEAByslMYgADcACQBPAHUAdABsAGkAbgBlACAANwAACBIEDQACAAAAGgTAAgQAvcogAIi8ONYLAFAAYQBnAGUAIABOAHUAbQBiAGUAcgAACRIEAQABAAAAGgQAAgMAOLqsudC5BgBIAGUAYQBkAGUAcgAAChIEBgADAAAAGgQgAgIAAaz8yAgARgBvAG8AdABuAG8AdABlAAALEgQDAAQAAAAaBAACAgD4u/zIBwBFAG4AZABuAG8AdABlAAAMEgQDAAQAAAAaBKABAgBUuqi6BABNAGUAbQBvAAANEgQCAAUAAAAbAIAEHAIBAAAABwIAgAcCBwAAAApABgAAAAAADkAGAAAAAAAGQAYAAAAAAB9ABwBkAAAAGkAGAAAAAAAdQAYAAAAAACBABwBkAAAAXgQAAQAAAAAAAAAAAAAAAAAAAAAeAEAAAAAAAB8EQAEAAAAAAAAAAAAAAAAAAAAAAAAAACAEgEA4AAAAAAAAAAIAAAACAAAAAAAAhQAAA4FAABYAAAAWAAAAFgAAABYAAAAAAICFAAADgUAAFgAAABYAAAAWAAAAFgAAAAAAHocAAAOBQAAWAAAAFgAAABYAAAAWAAAAAADghQAAA4FAABYAAAAWAAAAFgAAABYAAAAAAAEAAAAAAUAABgAAABYAAAACAAAAAgAAAAAAyocAAAOBQAAWAAAAFgAAABYAAAAWAAAAAABQiAAAAgFAABYAAAAWAAAAFgAAABYAAAAJAFS6qLogACOxMK4oACYAVAApAP////8HAETMsMYwrigAJgBJACkAAAAAAAAABwBEzLDGMK4oACYASQApAAQAAACjigAAAgFAABYAAAAWAAAAFgAAABYAAAAJAJDH2bMgAETMsMYwriAAqbpduAEAAAAAAUAABgAAABYAAAACAAAAAgAAAAAAoIoAAAIBQAAWAAAAFgAAABYAAAAWAAAAAAChigAAAgFAABYAAAAWAAAAFgAAABYAAAAAAP////8MAFCtFcggAIC9ONYgACOxMK4oACYAVgApAAAAAAAAAAwAUK0VyCAAgL041iAAI7EwrigAJgBWACkAFgAAAGGIAAADgUAAFgAAABYAAAAWAAAAFgAAAAAAb4gAAAOBQAAWAAAAFgAAABYAAAAWAAAAAABuiAAAA4FAABYAAAAWAAAAFgAAABYAAAAAAGKIAAADgUAAFgAAABYAAAAWAAAAFgAAAAAAY4gAAAOBQAAWAAAAFgAAABYAAAAWAAAAAAAAAH4AAH6XfBIAQLgBAjCk8wEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=","sum":"/v8AAA0AAABgtqKfYRDUEbTGAGCXwJ2MAQAAAGC2op9hENQRtMYAYJfAnYwwAAAAxQEAAA4AAAACAAAAeAAAAAMAAAC0AAAABAAAAMAAAAAUAAAA0AAAAAUAAAAUAQAABgAAACABAAAIAAAALAEAAAkAAAA4AQAADAAAAIQBAAANAAAAkAEAAAsAAACcAQAADgAAAKgBAAAVAAAAsAEAAAAAAAC4AQAAHwAAABkAAACgJSAAMgAwADEAMwBEsSAAMQDUxiAAUK0hxwTH0MaM1iAAANPEvCAA9LzgrCAAoCUAAAAAHwAAAAEAAAAAAAAAHwAAAAMAAAB4AHAAAAAAAB8AAAAdAAAAMgAwADEAMwBEsSAAMQDUxiAAMgAzAHzHIAAYwpTGfMcgACTGBMggADEAMQA6ADQANwA6ADAANgAAAAAAHwAAAAEAAAAAAAAAHwAAAAEAAAAAAAAAHwAAAAIAAABhAAAAHwAAACEAAAAxADIALAAgADAALAAgADAALAAgADMAMQA0ADYAIABXAEkATgAzADIATABFAFcAaQBuAGQAbwB3AHMAXwAxADAAAAAAAEAAAADwE5/uE/nNAUAAAACQ02loTkfdAUAAAAAAAAAAAAAAAAMAAAAAAAAAAwAAAAAAAAABAAAAAAAAAAEAAAAA","link":"AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=","js":"82dgYChjSGQoYlBgcGHIZ0hmKGXIZUhlyGMoYSgGitkCcQSDB0M5QwEWeWsGXgYunPohujF16TE4AkVKGDKBOlMZ4nGaDzEdGfwHAgCR35jcsgAAAA==","jsv":"Y2SAAAD334ipCAAAAA==","pre":"QgCAAREAAAAEAAAAAAAAAwIAAAABAAAAAAAAAEMEIAICAGRjZXMAAAAAAAAAAAIAAgBkbG9jAAAAAAAAAAACAA0ARAQAAQAAAAAHAAAAEAAAAAYAAABFBEACAAAAAAAAAABgCQAAYAkAAPgHAACgBQAAAAAAAGDSAAAAAAYARwTwAmRjZXMAAAAAbgQAAAAAQB8AAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAASQiAAojoAADcSAEAEwsAABMLAABSAwAAUgMAAFIDAABSAwAAAAAAAAAAAABKCMABAAAAAAAAAAApAAEA/////1IDNwIbAQEBAAAAAEoIwAEAAAAAAAAAACkAAQD4L+AAUgM3AgAAAQEAAAAASwjgAAEAAACJBYkFiQWJBQEASwjgAAEAAACJBYkFiQWJBQEASwjgAAEAAACJBYkFiQWJBQEARwQAAWRsb2MEEAAAAAAAAAAAAAA=","tbl":"QgCAAQkAAIAACAAAAQAAAAIAAAABAAAAAAAAAEMEIAELACBsYnQAAAAAAAAAAAsADQBEBAABAAAAAAcAAAAIAAAACAAAAEUEQAIAAAAAgCUAAJgIAACYCAAATgcAACgFAAAAAAAAAAAAAAAABgBHBOACIGxidBAiKghTAAAAJQAAAAfQAAAmEwEAAAAAABsBGwEbARsBgMSeRgAAAAAAAE0IIAIGAAAEBgADAAAA/gH+AY0AjQADAAMAAwADAAMAAwADAAAA","cells":["AQAAACAAAAAAAAAAAQABAPcTAADvBgAA/gH+AY0AjQADAPcTAAAAAAAAAAAAAAA=","AQAAACAAAAABAAAAAQABACNfAADvBgAA/gH+AY0AjQADACNfAAAAAAAAAAAAAAA=","AQAAACAAAAACAAAAAQABAO1cAADvBgAA/gH+AY0AjQADAO1cAAAAAAAAAAAAAAA=","AgAAACAAAAAAAAEAAQABAPcTAACdHAAA/gH+AY0AjQADAPcTAAAAAAAAAAAAAAA=","AwAAACAAAAABAAEAAQABACNfAACdHAAAjQCNAI0AjQAFACNfAAAAAAAAAAAAAAA=","BAAAACAAAAACAAEAAQABAO1cAACdHAAAjQCNAI0AjQAFAO1cAAAAAAAAAAAAAAA=","AgAAACAAAAAAAAIAAQABAPcTAAAeSwAA/gH+AY0AjQADAPcTAAAAAAAAAAAAAAA=","DAAAACAAAAABAAIAAQABACNfAAAeSwAAjQCNAI0AjQAFACNfAAAAAAAAAAAAAAA=","CwAAACAAAAQCAAIAAQABAO1cAAAeSwAAjQCNAI0AjQAFAO1cAAAAAAAAAAAAAAA=","AgAAACAAAAAAAAMAAQABAPcTAADkXAAA/gH+AY0AjQADAPcTAAAAAAAAAAAAAAA=","CQAAACAAAAABAAMAAQABACNfAADkXAAAjQCNAI0AjQAFACNfAAAAAAAAAAAAAAA=","DAAAACAAAAACAAMAAQABAO1cAADkXAAAjQCNAI0AjQAFAO1cAAAAAAAAAAAAAAA=","AgAAACAAAAAAAAQAAQABAPcTAADvYgAA/gH+AY0AjQADAPcTAAAAAAAAAAAAAAA=","EQAAACAAAAABAAQAAQABACNfAADvYgAAjQCNAI0AjQAFACNfAAAAAAAAAAAAAAA=","EQAAACAAAAACAAQAAQABAO1cAADvYgAAjQCNAI0AjQAFAO1cAAAAAAAAAAAAAAA=","AQAAACAAAAEAAAUAAQABAPcTAAA/GQAA/gH+AY0AjQADAPcTAAAAAAAAAAAAAAA=","AwAAACAAAAEBAAUAAQABACNfAAA/GQAAjQCNAI0AjQAFACNfAAAAAAAAAAAAAAA=","AQAAACAAAAECAAUAAQABAO1cAAA/GQAAjQCNAI0AjQAFAO1cAAAAAAAAAAAAAAA="],"rows":6,"cols":3};

// ===== 파일 끝: 이 줄이 보이면 전체가 들어간 것입니다 =====
