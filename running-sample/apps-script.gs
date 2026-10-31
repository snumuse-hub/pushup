/**
 * 자유학교 러닝 기록 받기 (구글 앱스 스크립트)
 *
 * 쓰는 방법
 *  1) 구글 시트를 새로 만들고  확장 프로그램 → Apps Script 를 엽니다.
 *  2) 이 코드를 전부 붙여넣고 저장(💾)합니다.
 *  3) 위쪽 함수 선택에서 setup 을 고르고 ▶ 실행 → 권한 허용
 *     → 시트에 '기록', '학생별 요약' 탭이 만들어집니다.
 *  4) 오른쪽 위 배포 → 새 배포 → 유형: 웹 앱
 *       실행 사용자: 나 / 액세스 권한: 모든 사용자
 *     → 배포 → 나오는 '웹 앱 URL'(…/exec)을 복사해서 러닝 앱 SHEET_URL 에 넣습니다.
 *
 * CLASS_CODE 는 러닝 앱(index.html)의 CLASS_CODE 와 같아야 기록을 받습니다.
 */
const CLASS_CODE = 'RUN2026';

const RECORD_SHEET = '기록';
const SUMMARY_SHEET = '학생별 요약';
const HEAD = ['제출 시각', '기록 ID', '학년', '반', '번호', '이름', '러닝 날짜', '거리(km)', '운동 시간', '운동 시간(초)',
  '평균 페이스', '평균 페이스(초)', '평균 속도(km/h)', '목표 거리(km)', '목표 시간', '목표 페이스', '목표 달성',
  '목표 거리 기록', '목표 대비(초)', '구간 페이스', '확인 필요', '기기'];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(RECORD_SHEET);
  if (!sh) sh = ss.insertSheet(RECORD_SHEET, 0);
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#fde68a');
  sh.setFrozenRows(1);
  sh.getRange('G:G').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.setColumnWidth(20, 260);

  let sm = ss.getSheetByName(SUMMARY_SHEET);
  if (!sm) sm = ss.insertSheet(SUMMARY_SHEET, 1);
  sm.clear();
  sm.getRange(1, 1, 1, 11).setValues([['학년', '반', '번호', '이름', '러닝 횟수', '총 거리(km)', '평균 페이스(초)',
    '최근 러닝', '평균 페이스', '목표 달성 횟수', '목표 달성률']]).setFontWeight('bold').setBackground('#bfdbfe');
  sm.setFrozenRows(1);
  // A~H: 학생별 묶어서 계산 / I~K: 보기 좋게 변환
  sm.getRange('A2').setFormula(
    "=IFERROR(QUERY('" + RECORD_SHEET + "'!B2:V,\"select C,D,E,F,count(B),sum(H),avg(L),max(G) where B is not null " +
    "group by C,D,E,F order by C,D,E label count(B) '', sum(H) '', avg(L) '', max(G) ''\",0),\"\")");
  sm.getRange('I2').setFormula('=ARRAYFORMULA(IF(A2:A="",,INT(G2:G/60)&"\'"&TEXT(ROUND(MOD(G2:G,60)),"00")&""""))');
  sm.getRange('J2').setFormula("=ARRAYFORMULA(IF(A2:A=\"\",,COUNTIFS('" + RECORD_SHEET + "'!C:C,A2:A,'" + RECORD_SHEET +
    "'!D:D,B2:B,'" + RECORD_SHEET + "'!E:E,C2:C,'" + RECORD_SHEET + "'!F:F,D2:D,'" + RECORD_SHEET + "'!Q:Q,\"달성\")))");
  sm.getRange('K2').setFormula('=ARRAYFORMULA(IF(A2:A="",,J2:J/E2:E))');
  sm.getRange('F:F').setNumberFormat('0.00');
  sm.getRange('G:G').setNumberFormat('0');
  sm.getRange('H:H').setNumberFormat('yyyy-mm-dd');
  sm.getRange('K:K').setNumberFormat('0%');

  const s1 = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 2) ss.deleteSheet(s1);
}

// 같은 값이 수식으로 실행되지 않도록 (=, +, -, @ 로 시작하는 글자 막기)
function safe(v) {
  if (typeof v !== 'string') return v;
  v = v.slice(0, 300);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.code !== CLASS_CODE) return out({ ok: false, error: 'code' });
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECORD_SHEET);
    if (!sh) return out({ ok: false, error: 'setup 을 먼저 실행해 주세요' });
    // 같은 기록이 두 번 들어오지 않게
    const last = sh.getLastRow();
    if (last > 1) {
      const ids = sh.getRange(2, 2, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) if (ids[i][0] === d.id) return out({ ok: true, dup: true });
    }
    const num = v => (v === '' || v === null || v === undefined || isNaN(+v)) ? '' : +v;
    sh.appendRow([
      new Date(), safe(d.id), num(d.grade), num(d.cls), num(d.no), safe(d.name),
      d.date ? new Date(String(d.date).replace(' ', 'T')) : '', num(d.distKm), safe(d.time), num(d.timeSec),
      safe(d.pace), num(d.paceSec), num(d.speed), num(d.goalKm), safe(d.goalTime), safe(d.targetPace), safe(d.goalResult),
      safe(d.goalRecord), num(d.diffSec), safe(d.splits), safe(d.flags), safe(d.device)
    ]);
    // 글자로 저장되는 시간 칸(예: 12:58)이 시각으로 바뀌지 않게 고정
    const r = sh.getLastRow();
    sh.getRange(r, 9).setNumberFormat('@'); sh.getRange(r, 15).setNumberFormat('@'); sh.getRange(r, 18).setNumberFormat('@');
    sh.getRange(r, 9).setValue(safe(d.time)); sh.getRange(r, 15).setValue(safe(d.goalTime)); sh.getRange(r, 18).setValue(safe(d.goalRecord));
    if (d.flags) sh.getRange(r, 21).setBackground('#fee2e2');
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return out({ ok: true, msg: '자유학교 러닝 기록 받는 중' });
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
