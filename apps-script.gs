/**
 * 위시리스트 앱 — Google Apps Script 백엔드 (최적화 버전)
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
 *
 * [성능 메모]
 * - 읽기(doGet)는 CacheService로 캐싱해 반복 조회 시 시트 접근을 건너뜁니다(빠름).
 * - 쓰기(add/update/status/approve/delete) 시 캐시를 즉시 무효화해 항상 최신을 유지합니다.
 * - 단, Apps Script 웹앱 특유의 콜드 스타트와 강제 리다이렉트 지연은 코드로 없앨 수 없습니다.
 *   (그게 필요하면 Google Sheets API 직접 연동이 근본 해법입니다.)
 */

const SHEET_NAME = '위시리스트';
const CACHE_KEY  = 'WL_ITEMS_V2';   // 읽기 결과 캐시 키
const CACHE_TTL  = 25;              // 초 — 외부(수동) 편집도 이 시간 내 반영

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
  clearCache();
}

/** 캐시 수동 초기화 (문제 시 실행) */
function clearCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}
function invalidate_() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { setup(); sheet = ss.getSheetByName(SHEET_NAME); }
  return sheet;
}

/** 원본 문자열을 JSON MIME으로 응답 (직렬화 재사용) */
function out_(str) {
  return ContentService.createTextOutput(str).setMimeType(ContentService.MimeType.JSON);
}
function json_(obj) {
  return out_(JSON.stringify(obj));
}

function findRow_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  const target = String(id);
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === target) return i + 2;
  }
  return -1;
}

/** 시트 전체를 items JSON 문자열로 직렬화 */
function buildItemsJson_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return JSON.stringify({ ok: true, items: [] });

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADER.length).getValues();
  const items = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    items.push({
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
    });
  }
  return JSON.stringify({ ok: true, items: items });
}

/** 전체 아이템 조회 — 캐시 우선 (?fresh=1 이면 캐시 건너뜀) */
function doGet(e) {
  const cache = CacheService.getScriptCache();
  const skip = e && e.parameter && (e.parameter.fresh === '1' || e.parameter.nocache === '1');

  if (!skip) {
    const cached = cache.get(CACHE_KEY);
    if (cached) return out_(cached);
  }

  const payload = buildItemsJson_(getSheet_());
  if (payload.length < 95000) cache.put(CACHE_KEY, payload, CACHE_TTL);  // 스크립트 캐시 100KB 한도 여유
  return out_(payload);
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
      invalidate_();
      return json_({ ok: true, id: id });
    }

    const row = findRow_(sheet, req.id);
    if (row === -1) return json_({ ok: false, error: '아이템을 찾을 수 없어요' });

    if (req.action === 'update') {
      // B:이름 C:사진URL D:가격 E:이유 (상태·등록일·결재 내역은 그대로 둠)
      sheet.getRange(row, 2, 1, 4).setValues([[
        req.name || '', req.photo || '', Number(req.price) || 0, req.reason || ''
      ]]);
      invalidate_();
      return json_({ ok: true });
    }

    if (req.action === 'status') {
      // 상태만 변경 — 결재 결과·의견·시각은 그대로 유지 (재상신해도 보존)
      sheet.getRange(row, 6).setValue(req.status);
      invalidate_();
      return json_({ ok: true });
    }

    if (req.action === 'approve') {
      const n = Number(req.approver);
      // H~M(8~13) 6칸을 한 번에 읽고 고쳐 한 번에 기록 (서비스 호출 최소화)
      const range = sheet.getRange(row, 8, 1, 6);
      const cur = range.getValues()[0];   // [a1, a1c, a2, a2c, a1t, a2t]
      const ts = req.result ? (Number(req.at) || new Date().getTime()) : '';
      if (n === 1) { cur[0] = req.result || ''; cur[1] = req.comment || ''; cur[4] = ts; }
      else         { cur[2] = req.result || ''; cur[3] = req.comment || ''; cur[5] = ts; }
      range.setValues([cur]);
      invalidate_();
      return json_({ ok: true });
    }

    if (req.action === 'delete') {
      sheet.deleteRow(row);
      invalidate_();
      return json_({ ok: true });
    }

    return json_({ ok: false, error: '알 수 없는 action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
