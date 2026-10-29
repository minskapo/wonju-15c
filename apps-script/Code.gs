/* ============================================================
   1.5℃ 라이프스타일 2주 살기 — 구글 시트 서버 (Apps Script)

   붙여넣는 곳: 수집용 구글 시트 → 확장 프로그램 → Apps Script
   배포: 배포 → 배포 관리 → 기존 배포 수정 → 버전: 새 버전 → 배포
         (실행 사용자: 나 / 액세스 권한: 모든 사용자)

   시트 구성
   - 첫 번째 탭: 결과. 참가자 한 명이 '계산완료', '2주완료' 두 줄을 남깁니다.
                 맨 끝 열에 운영단위가 붙습니다.
   - '참가자' 탭: 로그인 정보와 진행 중인 기록 전체(JSON)를 한 사람당 한 줄로 보관합니다.
                 없으면 처음 등록할 때 자동으로 만듭니다.
                 PIN 원문은 저장하지 않고 해시값만 저장합니다.
   ============================================================ */

var SECRET = "wonju2026-9f3k";   // index.html 의 CONFIG.SUBMIT_KEY 와 같은 값
var USERS_SHEET = "참가자";
var MAX_FAILS = 5;               // PIN을 이 횟수만큼 틀리면
var LOCK_SECONDS = 600;          // 이 시간(초) 동안 잠급니다

var COL = { KEY:1, NAME:2, PHONE:3, GROUP:4, PINHASH:5, TOKEN:6, STATE:7, CREATED:8, UPDATED:9 };
var USER_HEAD = ["로그인키","이름","휴대전화","운영단위","PIN해시","토큰","기록(JSON)","등록시각","갱신시각"];

function doPost(e){
  var d;
  try{ d = JSON.parse(e.postData.contents); }catch(err){ return json({ ok:false, error:"bad" }); }
  if(d.key !== SECRET) return ContentService.createTextOutput("denied");

  if(d.action === "register") return json(register(d));
  if(d.action === "login")    return json(login(d));
  if(d.action === "sync")     return json(sync(d));

  // action 이 없으면 결과 한 줄을 추가합니다 (이전 버전과 같은 동작).
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  sh.appendRow([new Date(), d.stage, d.name, d.contact, d.region, d.age, d.gender, d.spend,
    d.food, d.consumption, d.transport, d.housing, d.leisure,
    d.total, d.target, d.gap, d.rate,
    d.action1, d.action2, d.action3,
    d.days, d.reducedKg, d.annualT, d.afterT, d.remainT,
    d.policy1a, d.policy1b, d.policy1c, d.policy2a, d.policy2b, d.policy2c,
    d.barrier1, d.barrier2, d.barrier3, d.social, d.notes, d.group || ""]);
  return ContentService.createTextOutput("ok");
}

/* ---------- 등록: 이름 + 휴대전화 + PIN ---------- */
function register(d){
  var name = normName(d.name), digits = String(d.phone || "").replace(/\D/g, "");
  if(!name || digits.length < 10 || !/^\d{4}$/.test(String(d.pin || ""))) return { ok:false, error:"bad" };
  var key = name + "|" + digits.slice(-4);
  if(isLocked(key)) return { ok:false, error:"locked" };

  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try{
    var sh = usersSheet(), row = findRow(sh, COL.KEY, key);
    if(row){
      // 같은 이름·번호 뒷자리가 이미 있으면 PIN이 맞을 때만 그 기록을 돌려줍니다.
      if(sh.getRange(row, COL.PINHASH).getValue() !== pinHash(key, d.pin)){
        fail(key); return { ok:false, error:"pin" };
      }
      var token = sh.getRange(row, COL.TOKEN).getValue();
      var state = parseState(sh.getRange(row, COL.STATE).getValue());
      return { ok:true, token:token, state:hasProgress(state) ? state : null };
    }
    var t = Utilities.getUuid(), now = new Date();
    sh.appendRow([key, String(d.name).trim(), String(d.phone), d.groupName || d.group || "",
      pinHash(key, d.pin), t, "", now, now]);
    return { ok:true, token:t };
  } finally { lock.releaseLock(); }
}

/* ---------- 로그인: 이름 + 번호 뒷 4자리 + PIN ---------- */
function login(d){
  var key = normName(d.name) + "|" + String(d.last4 || "");
  if(isLocked(key)) return { ok:false, error:"locked" };
  var sh = usersSheet(), row = findRow(sh, COL.KEY, key);
  if(!row || sh.getRange(row, COL.PINHASH).getValue() !== pinHash(key, d.pin)){
    fail(key); return { ok:false, error:"nomatch" };
  }
  CacheService.getScriptCache().remove("fail:" + key);
  var state = parseState(sh.getRange(row, COL.STATE).getValue());
  // 첫 저장이 서버에 닿기 전이라면 등록할 때 받은 이름과 번호로 채웁니다.
  if(!state.profile) state.profile = { name:sh.getRange(row, COL.NAME).getValue(),
                                       contact:String(sh.getRange(row, COL.PHONE).getValue()) };
  return { ok:true, token:sh.getRange(row, COL.TOKEN).getValue(), state:state };
}

/* ---------- 기록 저장: 토큰으로 찾아서 덮어씁니다 ---------- */
function sync(d){
  if(!d.token) return { ok:false, error:"bad" };
  var text = JSON.stringify(d.state || {});
  if(text.length > 49000) return { ok:false, error:"toolarge" };   // 셀 하나에 5만 자까지 들어갑니다
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try{
    var sh = usersSheet(), row = findRow(sh, COL.TOKEN, d.token);
    if(!row) return { ok:false, error:"notoken" };
    sh.getRange(row, COL.STATE, 1, 1).setValue(text);
    sh.getRange(row, COL.UPDATED, 1, 1).setValue(new Date());
    // 운영단위를 나중에 고른 경우에도 시트에 반영합니다.
    var g = d.groupName;
    if(g && !sh.getRange(row, COL.GROUP).getValue()) sh.getRange(row, COL.GROUP).setValue(g);
    return { ok:true };
  } finally { lock.releaseLock(); }
}

/* ---------- 도움 함수 ---------- */
function json(o){
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function normName(n){ return String(n || "").replace(/\s+/g, ""); }

function usersSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(USERS_SHEET);
  if(!sh){
    sh = ss.insertSheet(USERS_SHEET);
    sh.appendRow(USER_HEAD);
    sh.setFrozenRows(1);
    sh.getRange("C:C").setNumberFormat("@");   // 휴대전화번호 앞자리 0이 사라지지 않게
  }
  return sh;
}

function findRow(sh, col, value){
  var last = sh.getLastRow();
  if(last < 2) return 0;
  var vals = sh.getRange(2, col, last - 1, 1).getValues();
  for(var i = 0; i < vals.length; i++) if(String(vals[i][0]) === String(value)) return i + 2;
  return 0;
}

function parseState(v){ try{ return v ? JSON.parse(v) : {}; }catch(e){ return {}; } }
function hasProgress(s){
  return !!(s && ((s.answers && Object.keys(s.answers).length) || (s.picked && s.picked.length)));
}

/* PIN 해시 — 시트에서 보이는 값만으로 PIN을 알아내기 어렵도록 비밀 소금값을 섞습니다.
   소금값은 처음 실행할 때 스크립트 속성에 자동으로 만들어집니다. 바꾸면 기존 PIN이 모두 무효가 됩니다. */
function pinHash(key, pin){
  var props = PropertiesService.getScriptProperties(), salt = props.getProperty("PIN_SALT");
  if(!salt){ salt = Utilities.getUuid(); props.setProperty("PIN_SALT", salt); }
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + "|" + key + "|" + pin,
    Utilities.Charset.UTF_8);
  return bytes.map(function(b){ return ("0" + (b & 255).toString(16)).slice(-2); }).join("");
}

/* 같은 이름·번호로 PIN을 여러 번 틀리면 잠시 잠급니다 (숫자 4자리를 차례로 넣어보는 것을 막습니다). */
function isLocked(key){
  return Number(CacheService.getScriptCache().get("fail:" + key) || 0) >= MAX_FAILS;
}
function fail(key){
  var c = CacheService.getScriptCache(), n = Number(c.get("fail:" + key) || 0) + 1;
  c.put("fail:" + key, String(n), LOCK_SECONDS);
}

/* 담당자용 — 참가자가 PIN을 잊었을 때 Apps Script 편집기에서 직접 실행합니다.
   이름과 번호 뒷자리, 새 PIN을 아래에 적고 resetPin 을 실행한 뒤 참가자에게 새 PIN을 알려주세요. */
function resetPin(){
  var name = "홍길동", last4 = "5678", newPin = "0000";
  var key = normName(name) + "|" + last4, sh = usersSheet(), row = findRow(sh, COL.KEY, key);
  if(!row) throw new Error("해당 참가자를 찾지 못했습니다: " + key);
  sh.getRange(row, COL.PINHASH).setValue(pinHash(key, newPin));
  CacheService.getScriptCache().remove("fail:" + key);
}
