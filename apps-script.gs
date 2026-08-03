/**
 * 위시리스트 앱 — Google Apps Script 백엔드
 *
 * [설치 방법]
 * 1. 구글 스프레드시트 새로 만들기
 * 2. 확장 프로그램 > Apps Script > 이 코드 전체 붙여넣기
 * 3. 상단 함수 선택에서 setup 선택 후 ▶ 실행 (시트/헤더 자동 생성, 최초 1회 권한 승인)
 * 4. 배포 > 새 배포 > 유형: 웹 앱
 *    - 실행 계정: 나
 *    - 액세스 권한: 모든 사용자
 * 5. 발급된 웹 앱 URL을 index.html의 API_URL에 붙여넣기
 *
 * ※ 코드를 수정한 뒤에는 배포 > 배포 관리 > (기존 배포) 편집 > 버전을 "새 버전"으로 저장해야
 *   변경 사항이 반영됩니다. (URL은 그대로 유지됩니다)
 */

const SHEET_NAME = '위시리스트';

// 시트 컬럼: A:id B:이름 C:사진URL D:가격 E:이유 F:상태 G:등록일
//           H:결재자1결과 I:결재자1의견 J:결재자2결과 K:결재자2의견
//           L:결재자1시각 M:결재자2시각  (순서 판정용, epoch ms)
const HEADER = ['id', '이름', '사진URL', '가격', '이유', '상태', '등록일',
                '결재자1결과', '결재자1의견', '결재자2결과', '결재자2의견',
                '결재자1시각', '결재자2시각'];

/** 최초 1회 실행: 시트와 헤더 생성 (재실행 시 헤더 행 갱신) */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { setup(); sheet = ss.getSheetByName(SHEET_NAME); }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function findRow_(sheet, id) {
  const ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

/** 전체 아이템 조회 */
function doGet(e) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return json_({ ok: true, items: [] });

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADER.length).getValues();
  const items = rows.filter(r => r[0]).map(r => ({
    id: String(r[0]),
    name: String(r[1]),
    photo: String(r[2]),
    price: Number(r[3]) || 0,
    reason: String(r[4]),
    status: String(r[5]),
    createdAt: r[6] instanceof Date ? Utilities.formatDate(r[6], 'Asia/Seoul', 'yyyy-MM-dd') : String(r[6]),
    a1: String(r[7]), a1c: String(r[8]),
    a2: String(r[9]), a2c: String(r[10]),
    a1t: Number(r[11]) || '', a2t: Number(r[12]) || ''
  }));
  return json_({ ok: true, items: items });
}

/**
 * 쓰기 작업 — 요청 본문(JSON)의 action으로 분기
 *  add     : { action, name, photo, price, reason, status }
 *  update  : { action, id, name, photo, price, reason }   // 상품 정보 수정 (상태·결재 내역은 유지)
 *  status  : { action, id, status }           // 상태 변경 (결재 내역은 유지)
 *  approve : { action, id, approver(1|2), result('승인'|'반려'|'보류'), comment, at(결재 시각 ms) }
 *  delete  : { action, id }
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const req = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    if (req.action === 'add') {
      const id = String(new Date().getTime());
      sheet.appendRow([
        id, req.name || '', req.photo || '', Number(req.price) || 0,
        req.reason || '', req.status || '위시', new Date(), '', '', '', '', '', ''
      ]);
      return json_({ ok: true, id: id });
    }

    const row = findRow_(sheet, req.id);
    if (row === -1) return json_({ ok: false, error: '아이템을 찾을 수 없어요' });

    if (req.action === 'update') {
      // B:이름 C:사진URL D:가격 E:이유 (상태·등록일·결재 내역은 그대로 둠)
      sheet.getRange(row, 2, 1, 4).setValues([[
        req.name || '', req.photo || '', Number(req.price) || 0, req.reason || ''
      ]]);
      return json_({ ok: true });
    }

    if (req.action === 'status') {
      // 상태만 변경 — 결재 결과·의견·시각은 그대로 유지 (재상신해도 보존)
      sheet.getRange(row, 6).setValue(req.status);
      return json_({ ok: true });
    }

    if (req.action === 'approve') {
      const n = Number(req.approver);
      const col = n === 1 ? 8 : 10;    // 결과·의견 열
      const tcol = n === 1 ? 12 : 13;  // 결재 시각 열
      sheet.getRange(row, col, 1, 2).setValues([[req.result || '', req.comment || '']]);
      // 결재하면 시각 기록, 취소(빈 결과)면 시각도 지움
      sheet.getRange(row, tcol).setValue(req.result ? (Number(req.at) || new Date().getTime()) : '');
      return json_({ ok: true });
    }

    if (req.action === 'delete') {
      sheet.deleteRow(row);
      return json_({ ok: true });
    }

    return json_({ ok: false, error: '알 수 없는 action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
