/**
 * 연구과제 예산현황 시트 → 참여율 관리 웹 동기화
 *
 * 원본은 드라이브에 있는 기존 .xlsx 파일 그대로 둔다(팀 권한 손댈 필요 없음).
 * 스크립트가 주기적으로 원본을 읽어 웹에 반영한다.
 *   - .xlsx 는 Apps Script 가 직접 못 읽으므로, 매번 임시 구글시트로 변환해 읽고 바로 버린다.
 *   - 원본 수정시각이 그대로면 변환도 전송도 하지 않는다(불필요한 호출 방지).
 *   - 평일(주말·공휴일 제외) 09·12·15·18·21시에만 확인한다.
 *
 * 설치
 *  1) 왼쪽 「서비스 +」 → Drive API 추가 (식별자 Drive, 버전 v3)
 *  2) 톱니(프로젝트 설정) → 스크립트 속성에  SYNC_TOKEN = (준호에게 받은 토큰)
 *  3) 함수 목록에서 setupTriggers 선택 → 실행 → 권한 승인 (공휴일 확인용 캘린더 포함)
 *  4) 확인은 syncNow 실행 후 실행 로그 보기
 */

/* 기존에 쓰던 원본 파일 (연구과제 예산현황v2_2026.xlsx) */
var SRC_FILE_ID = "18jl7dagS5IF_hFHb0rzhoHIr6N7i_OSv";

var FN_URL = "https://pkencmbryzgtnwrxlksz.supabase.co/functions/v1/sync-budget";
/* Edge Function 이 JWT 검증을 켜둔 상태라 공개 anon 키를 같이 보낸다.
   이 키는 웹페이지에도 박혀 있는 공개 값이고, 실제 인가는 x-sync-token 이 한다. */
var ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBrZW5jbWJyeXpndG53cnhsa3N6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzkzNDYyOTIsImV4cCI6MjA5NDkyMjI5Mn0.Jx8qCUlVgGLDzKiscMetgEGjZ4FSs2ptTixvaUO7IvA";

/* 예산 행으로 인정하는 비목 (그 아래 단가/회 상세표를 걸러내기 위함) */
var VALID_BM = ["직접비","간접비","인건비","운영비","여비","업무추진비","현물",
                "위탁연구개발비","연구활동비","연구재료비"];
var SKIP_SHEETS = ["총괄","2026","2025","사용내역","민간부담금 납부 내역"];
var TARGET_YEAR = 2026;

/* 동기화 시각 — 평일 09·12·15·18·21시 (한국시간). 주말·공휴일은 건너뛴다. */
var TZ = "Asia/Seoul";
var HOLIDAY_CAL = "ko.south_korea#holiday@group.v.calendar.google.com";

function num_(v){
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  var n = parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
  return isNaN(n) ? 0 : n;
}
function s_(v){ return v === null || v === undefined ? "" : String(v).trim(); }

/** 과제 시트에서 해당 연도 블록의 비목/세목 행을 뽑는다 */
function readBudget_(sh){
  var v = sh.getDataRange().getValues();
  var rows = [], inBlock = false;
  for (var r = 0; r < v.length; r++){
    var a = s_(v[r][0]), b = s_(v[r][1]);
    if (/^\d+차년도$/.test(a)){                       // "1차년도 | 2026 | ..."
      inBlock = (num_(v[r][1]) === TARGET_YEAR);
      continue;
    }
    if (a === "비목" && b === "세목") continue;
    if (!a) continue;
    var yk = v[r].length > 10 ? num_(v[r][10]) : 0;   // K열 연도
    var take = (yk === TARGET_YEAR) || (yk === 0 && inBlock);
    if (!take) continue;
    if (a === "현금 계" || a === "현물 계") continue;
    if (a === "총합계"){ inBlock = false; continue; } // 아래는 단가/회 상세표
    if (VALID_BM.indexOf(a) < 0) continue;
    rows.push({ bm:a, sm:b,
                budget:num_(v[r][4]), used:num_(v[r][5]),
                pend:num_(v[r][6]),  left:num_(v[r][7]) });
  }
  return rows;
}

/** 시트 하단의 "비목 | 세목 | 세세목" 편성표 */
function readCatalog_(sh){
  var v = sh.getDataRange().getValues(), out = [], hdr = -1;
  for (var r = 0; r < v.length; r++){
    var a = s_(v[r][0]), b = s_(v[r][1]), c = s_(v[r][2]);
    if (a === "비목" && b === "세목" && c === "세세목"){ hdr = r; continue; }
    if (hdr >= 0 && r > hdr){
      if (!a){ if (r > hdr + 1) hdr = -1; continue; }
      if (c) out.push({ bm:a, sm:b, ssm:c });
    }
  }
  return out;
}

/** 사용내역 시트 */
function readHistory_(ss){
  var sh = ss.getSheetByName("사용내역");
  if (!sh) return [];
  var v = sh.getDataRange().getValues(), out = [], started = false;
  for (var r = 0; r < v.length; r++){
    if (!started){
      if (s_(v[r][0]) === "과제명" && s_(v[r][2]) === "비목") started = true;
      continue;
    }
    var p = s_(v[r][0]), bm = s_(v[r][2]);
    if (!p || !bm) continue;
    out.push({ proj:p, bm:bm, sm:s_(v[r][3]), ssm:s_(v[r][4]), desc:s_(v[r][6]).slice(0, 40) });
  }
  return out;
}

function buildPayload_(ss){
  var budgets = {}, catalog = {};
  ss.getSheets().forEach(function(sh){
    var nm = sh.getName();
    if (SKIP_SHEETS.indexOf(nm) >= 0) return;
    var rows = readBudget_(sh);
    if (rows.length) budgets[nm] = rows;
    var cat = readCatalog_(sh);
    if (cat.length) catalog[nm] = cat;
  });
  return {
    asof: Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd"),
    budgets: budgets,
    catalog: catalog,
    history: readHistory_(ss)
  };
}

/** 원본 .xlsx 를 임시 구글시트로 변환해 연다. 반환값은 {ss, tmpId} */
function openSource_(){
  var src = DriveApp.getFileById(SRC_FILE_ID);
  var copy = Drive.Files.copy(
    { name: "[임시-자동생성] " + src.getName(), mimeType: MimeType.GOOGLE_SHEETS },
    SRC_FILE_ID);
  return { ss: SpreadsheetApp.openById(copy.id), tmpId: copy.id, name: src.getName() };
}

function post_(payload){
  var token = PropertiesService.getScriptProperties().getProperty("SYNC_TOKEN");
  if (!token) throw new Error("스크립트 속성에 SYNC_TOKEN 을 먼저 넣어주세요.");
  var res = UrlFetchApp.fetch(FN_URL, {
    method: "post",
    contentType: "application/json",
    headers: { "x-sync-token": token, "Authorization": "Bearer " + ANON_KEY, "apikey": ANON_KEY },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode(), body = res.getContentText();
  Logger.log("HTTP " + code + " " + body);
  if (code !== 200) throw new Error("동기화 실패 (" + code + "): " + body);
  return body;
}

/** 원본을 읽어 전송. force 가 아니면 원본이 안 바뀐 경우 건너뛴다. */
function sync_(force){
  var props = PropertiesService.getScriptProperties();
  var mtime = DriveApp.getFileById(SRC_FILE_ID).getLastUpdated().toISOString();
  if (!force && props.getProperty("SRC_MTIME") === mtime){
    Logger.log("원본 변경 없음 (" + mtime + ") — 건너뜀");
    return "변경 없음";
  }
  var src = openSource_();
  try {
    var payload = buildPayload_(src.ss);
    if (!Object.keys(payload.budgets).length)
      throw new Error("예산 행을 하나도 못 읽었습니다. 원본 시트 구조를 확인해주세요.");
    var body = post_(payload);
    props.setProperty("SRC_MTIME", mtime);
    props.setProperty("LAST_SYNC", new Date().toISOString());
    return body;
  } finally {
    try { DriveApp.getFileById(src.tmpId).setTrashed(true); } catch (e) {}   // 임시본은 반드시 정리
  }
}

/** 수동 실행용 — 변경 여부와 관계없이 무조건 보낸다 */
function syncNow(){ return sync_(true); }

/** 지금이 동기화할 시간인가. 건너뛸 이유가 있으면 그 이유를 돌려준다. */
function skipReason_(){
  var now = new Date();
  var day = Number(Utilities.formatDate(now, TZ, "u"));   // 1=월 … 7=일
  if (day >= 6) return "주말";
  var hh = Number(Utilities.formatDate(now, TZ, "H"));
  if (hh < 9 || hh > 21) return "업무시간 외 (" + hh + "시)";
  if (isHoliday_(now)) return "공휴일";
  return "";
}
function isHoliday_(d){
  try {
    var cal = CalendarApp.getCalendarById(HOLIDAY_CAL);
    return cal ? cal.getEventsForDay(d).length > 0 : false;
  } catch (e) {
    return false;              // 달력 권한이 없으면 공휴일 검사만 건너뛴다
  }
}

/** 트리거용 — 업무시간에만, 원본이 바뀌었을 때만 보낸다 */
function tick(){
  var why = skipReason_();
  if (why){ Logger.log("건너뜀: " + why); return "건너뜀 (" + why + ")"; }
  return sync_(false);
}

/** 아침 첫 회차 — 변경 여부와 무관하게 한 번 전체 갱신(안전망) */
function morningFull(){
  var why = skipReason_();
  if (why){ Logger.log("건너뜀: " + why); return "건너뜀 (" + why + ")"; }
  return sync_(true);
}

function setupTriggers(){
  ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("morningFull").timeBased().atHour(9).everyDays(1).inTimezone(TZ).create();
  [12, 15, 18, 21].forEach(function(h){
    ScriptApp.newTrigger("tick").timeBased().atHour(h).everyDays(1).inTimezone(TZ).create();
  });
  syncNow();
  return "트리거 설치 완료 (평일 09·12·15·18·21시) · 첫 동기화까지 마쳤습니다.";
}

/** 설치 확인용 */
function checkSetup(){
  var props = PropertiesService.getScriptProperties();
  var f = DriveApp.getFileById(SRC_FILE_ID);
  Logger.log("원본: " + f.getName() + " / 최종수정 " + f.getLastUpdated());
  Logger.log("SYNC_TOKEN: " + (props.getProperty("SYNC_TOKEN") ? "설정됨" : "없음"));
  Logger.log("마지막 동기화: " + (props.getProperty("LAST_SYNC") || "없음"));
  Logger.log("트리거: " + ScriptApp.getProjectTriggers().map(function(t){
    return t.getHandlerFunction();
  }).join(", "));
  Logger.log("지금 동기화 가능? " + (skipReason_() || "예"));
}
