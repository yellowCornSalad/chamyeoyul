/**
 * 연구과제 예산현황 시트 → 참여율 관리 웹 동기화
 *
 * 설치
 *  1) 시트에서  확장 프로그램 → Apps Script
 *  2) 이 파일 내용을 통째로 붙여넣고 저장
 *  3) 왼쪽 톱니(프로젝트 설정) → 스크립트 속성 → 속성 추가
 *       SYNC_TOKEN = (준호에게 받은 토큰)
 *  4) 위 함수 목록에서 setupTriggers 선택 → 실행 (최초 1회 권한 승인)
 *  5) 확인은 syncNow 실행 후 실행 로그 보기
 *
 * 이후 시트를 고치면 1분 안에 웹에 반영된다.
 */

var FN_URL = "https://pkencmbryzgtnwrxlksz.supabase.co/functions/v1/sync-budget";
/* Edge Function 이 JWT 검증을 켜둔 상태라 공개 anon 키를 같이 보낸다.
   이 키는 웹페이지에도 박혀 있는 공개 값이고, 실제 인가는 x-sync-token 이 한다. */
var ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBrZW5jbWJyeXpndG53cnhsa3N6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzkzNDYyOTIsImV4cCI6MjA5NDkyMjI5Mn0.Jx8qCUlVgGLDzKiscMetgEGjZ4FSs2ptTixvaUO7IvA";

/* 예산 행으로 인정하는 비목 (그 아래 단가/회 상세표를 걸러내기 위함) */
var VALID_BM = ["직접비","간접비","인건비","운영비","여비","업무추진비","현물",
                "위탁연구개발비","연구활동비","연구재료비"];
var SKIP_SHEETS = ["총괄","2026","2025","사용내역","민간부담금 납부 내역"];
var TARGET_YEAR = 2026;

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

function buildPayload_(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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

function syncNow(){
  var token = PropertiesService.getScriptProperties().getProperty("SYNC_TOKEN");
  if (!token) throw new Error("스크립트 속성에 SYNC_TOKEN 을 먼저 넣어주세요.");
  var payload = buildPayload_();
  var nProj = Object.keys(payload.budgets).length;
  if (!nProj) throw new Error("예산 행을 하나도 못 읽었습니다. 시트 구조를 확인해주세요.");

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
  PropertiesService.getScriptProperties().setProperty("LAST_SYNC", new Date().toISOString());
  return body;
}

/* 편집이 잦을 때 매번 쏘지 않도록, 변경 표시만 남기고 1분 트리거가 실제 전송한다 */
function onSheetChange(){
  PropertiesService.getScriptProperties().setProperty("DIRTY", "1");
}
function tick(){
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty("DIRTY") !== "1") return;
  props.deleteProperty("DIRTY");
  syncNow();
}

function setupTriggers(){
  ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger("onSheetChange").forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger("tick").timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger("syncNow").timeBased().everyHours(6).create();   // 안전망
  syncNow();
  return "트리거 설치 완료 · 첫 동기화까지 마쳤습니다.";
}
