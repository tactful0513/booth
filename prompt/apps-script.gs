/**
 * 프롬프트 퀘스트 웹 워크시트 — 구글 시트 연결용 Apps Script
 * 사용법: 새 구글 시트 → 확장 프로그램 → Apps Script → 이 코드를 통째로 붙여넣기 → 배포 → 새 배포
 *        (유형: 웹 앱 / 실행 사용자: 나 / 액세스 권한: 모든 사용자) → 웹 앱 URL을 선생님 화면(board.html) 설정에 붙여넣기
 * 학생 1명 = 1줄. 같은 학생(수업코드·학년·반·모둠·번호)이 다시 보내면 그 줄이 최신 내용으로 바뀝니다.
 */
const SHEET = '응답';

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
// 학생이 쓴 글이 =, +, -, @로 시작하면 수식으로 실행되지 않게 막음
function safe_(v) {
  return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const row = JSON.parse(e.postData.contents);
    if (!row['키'] || !row['수업코드']) return json_({ ok: false, error: 'no key' });
    row['보낸시각'] = new Date();
    const sh = sheet_();
    const lastCol = sh.getLastColumn();
    const head = lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    Object.keys(row).forEach(k => { if (head.indexOf(k) < 0) head.push(k); });
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);

    const keyCol = head.indexOf('키') + 1, n = sh.getLastRow();
    let r = 0;
    if (n > 1) {
      const keys = sh.getRange(2, keyCol, n - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) if (keys[i][0] === row['키']) { r = i + 2; break; }
    }
    const old = r ? sh.getRange(r, 1, 1, head.length).getValues()[0] : [];
    if (!r) r = Math.max(n, 1) + 1;
    sh.getRange(r, 1, 1, head.length).setValues([head.map((h, i) => (h in row) ? safe_(row[h]) : (old[i] ?? ''))]);
    return json_({ ok: true, row: r });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 선생님 화면용: ?code=수업코드 → 그 수업의 학생 기록 목록 (이름은 보내지 않음, ?c=는 구글 예약어라 못 씀)
function doGet(e) {
  const c = (e.parameter.code || '').trim();
  const sh = sheet_(), n = sh.getLastRow(), m = sh.getLastColumn();
  if (n < 2 || !c) return json_({ ok: true, rows: [] });
  const v = sh.getRange(1, 1, n, m).getValues(), head = v[0];
  const ci = head.indexOf('수업코드'), ji = head.indexOf('_json'), ii = head.indexOf('캐릭터그림'), ti = head.indexOf('보낸시각');
  const rows = v.slice(1).filter(x => String(x[ci]) === c).map(x => ({
    json: x[ji], img: ii >= 0 ? x[ii] : '', at: ti >= 0 ? x[ti] : ''
  }));
  return json_({ ok: true, rows });
}
