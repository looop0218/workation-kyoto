/*  워케이션 경비+예약 공유 — Apps Script 백엔드 (Code.gs)  [토큰 무중단교체 + sheet 라우팅(경비/예약) 통합본]
 *  ────────────────────────────────────────────────────────────────────────
 *  ⚠️ 재배포 (같은 웹앱 URL 유지 — index.html / trip.json 수정 불필요):
 *     Apps Script 편집기에 이 파일 전체를 붙여넣기 → 저장 →
 *     [배포] → [배포 관리] → 기존 배포의 ✏️(편집) → 버전 [새 버전] → [배포]
 *     ※ [새 배포]로 만들면 /exec URL이 바뀌어 앱 재수정 필요 → 반드시 "기존 배포 편집 → 새 버전"
 *     실행 계정: 본인(나) / 액세스: 모든 사용자
 *  ────────────────────────────────────────────────────────────────────────
 *  이번 변경점(배치A):
 *   (1) 토큰 Script Properties 분리 — 코드에 시크릿 없음(git 편입 가능).
 *       ★★ 배포 전 필수: 편집기 ⚙️프로젝트 설정 → 스크립트 속성 → 속성 "TOKENS" 추가,
 *          값 = 쉼표로 구분한 토큰 목록(공백 무관). 속성 미설정 시 모든 요청이 bad token 거부.
 *          로테이션: 속성값에서 구토큰만 지우면 끝 — 재배포 불필요(런타임에 읽음).
 *          속성은 구 버전 가동 중 미리 설정해도 안전(구 코드는 속성을 안 읽음) → 설정 후 새 버전 배포 = 무중단.
 *   (2) 소프트삭제: delete가 행을 지우지 않고 deleted 플래그(경비 10열/예약 12열) 마킹.
 *       list/update/delete/영수증연결은 deleted 행을 없는 것으로 취급(삭제 항목 부활 금지 유지).
 *   (3) 감사로그: "로그" 시트에 add/update/delete/upload/uploadDoc 기록(토큰 뒤6자) +
 *       토큰별 사용 추적("seen" 행, 시간당 1회) → 로테이션 전 구토큰 트래픽을 로그 시트에서 확인.
 *   (4) cleanupOrphanFiles(): 시트 미참조 영수증·예약문서 파일 휴지통 이동(편집기 수동 실행).
 *   (5) sheet 라우팅(expenses 폴백)·upload/receipt/충전·보안 가드(크기 컷·마스킹·덮어쓰기 방지) 전부 보존.
 *   (6) 맛집후보(places) 시트 추가 — 팀원 발굴 맛집 수집(sheet=places, list/add/update/delete, 소프트삭제·감사로그 동일 패턴).
 *   (7) ★일정 오버레이(itinerary) 시트 추가 — 이 줄이 보이면 최신 버전. 팀 일정 추가분(sheet=itinerary).
 *  ────────────────────────────────────────────────────────────────────────
 */

var SHEET_ID = "1vQNshHQwnkTlJUWm16xYdLJsYHQW8g0zvDcd6PrT_gs";

/* 토큰: Script Properties "TOKENS"(쉼표구분)에서 런타임 로드 — 코드에 토큰 리터럴 금지.
 * 미설정이면 빈 목록 → 모든 요청 거부. 배포 전 속성 설정이 반드시 선행(파일 상단 안내). */
function _tokens(){
  var s = "";
  try { s = PropertiesService.getScriptProperties().getProperty("TOKENS") || ""; } catch(e){}
  var out = [];
  s.split(",").forEach(function(t){ t = String(t||"").replace(/^\s+|\s+$/g,""); if(t) out.push(t); });
  return out;
}

/* ── 경비(expenses) 시트: 기존 스키마 그대로 ── */
var SHEET    = "경비";
var CATS     = ["식비","교통","입장·관광","회식·술","숙박","기타","충전"];

var RECEIPT_FOLDER    = "워케이션영수증";                                  // 없으면 자동 생성
var RECEIPT_MIME_RE   = /^image\/(jpe?g|png|webp|heic|heif)$/;             // 허용 이미지 MIME
var RECEIPT_MAX_BYTES = 8 * 1024 * 1024;                                   // 디코드 후 실바이트 상한
var RECEIPT_B64_MAX   = 12 * 1024 * 1024;                                  // 디코드 前 원시 base64 길이 컷(폭탄 방어)
var RECEIPT_URL_RE    = /^https:\/\/(drive|docs)\.google\.com\//;          // receipt에 저장 허용할 URL

/* ── 예약(bookings) 시트: 스키마 ──
 *  헤더: id | type | title | datetime | sub | detail | voucherUrl | mapUrl | memo | ts | attachments
 *  type ∈ {flight, hotel, train}. PNR·여권영문명 등 민감값은 설계상 저장 안 함.
 *  attachments: JSON 배열 문자열 [{fileId,mime,size}] — 파일 실체는 booking-docs 제한폴더(4명 뷰어)에.
 */
var BK_SHEET   = "예약";
var BK_HEADERS = ["id","type","title","datetime","sub","detail","voucherUrl","mapUrl","memo","ts","attachments"];
var BK_TYPES   = ["flight","hotel","train"];
var BK_URL_RE  = /^https?:\/\//i;                                          // voucherUrl/mapUrl: http(s)만 허용
var BK_LEN     = { title:120, sub:120, detail:200, memo:300 };             // 필드 길이컷

/* ── 맛집후보(places) 시트: 팀원 발굴 맛집 수집 ──
 *  헤더: id | name | region | lat | lng | foundBy | url | memo | status | ts | deleted
 *  status ∈ {candidate(기본, 미채점), scored}. 좌표는 일본 범위 밖이면 저장 안 함(리스트 전용).
 */
var PL_SHEET   = "맛집후보";
var PL_HEADERS = ["id","name","region","lat","lng","foundBy","url","memo","status","ts"];
var PL_LEN     = { name:80, region:40, memo:200 };

/* ── 일정 오버레이(itinerary) 시트: 정적 동선(index.html 임베드)은 그대로, 팀 추가분만 여기 ──
 *  헤더: id | date | time | name | lat | lng | memo | category | createdBy | ts | deleted
 *  date=YYYY-MM-DD, time=HH:MM(정렬용). 좌표는 일본 범위 밖이면 저장 안 함.
 */
var IT_SHEET   = "일정추가";
var IT_HEADERS = ["id","date","time","name","lat","lng","memo","category","createdBy","ts"];
var IT_LEN     = { name:80, memo:200 };
var IT_DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
var IT_TIME_RE = /^\d{1,2}:\d{2}$/;

/* ── 회비(fund) 시트: 여행 자금 풀(원화). 경비(지출)와 완전 분리 ──
 *  헤더: id | date | krw | payer | memo | ts | deleted
 *  전체 회비 = Σ krw. 환전분은 경비 '충전'의 실지불 원화(krw)로 별도 추적 → 앱에서 미환전 계산.
 */
var FUND_SHEET   = "회비";
var FUND_HEADERS = ["id","date","krw","payer","memo","ts"];
var FUND_LEN     = { memo:200 };

/* ── 예약 첨부 문서(booking-docs) ──
 *  ★★★ 아래 BK_DOCS_FOLDER_ID 에 드라이브 폴더 ID를 붙여넣으세요. ★★★
 *    - 그 폴더는 팀원 4명 구글계정에 "뷰어"로만 공유(일반 액세스=제한됨, '링크가 있는 모든 사용자' 금지).
 *    - 스크립트가 그 폴더에 만든 파일은 폴더의 제한공유를 상속 → 4명만 열람(공개 안 됨).
 */
var BK_DOCS_FOLDER_ID = "1fFmhYVphJEDiBJyfdA56uc72hDF-Ju5_";               // booking-docs 폴더(4명 뷰어 공유)
var BK_DOC_MIME_RE    = /^(application\/pdf|image\/(jpe?g|png|webp|heic|heif))$/;  // 허용: PDF + 이미지
var BK_DOC_MAX_BYTES  = 10 * 1024 * 1024;                                  // 디코드 후 실바이트 상한
var BK_DOC_B64_MAX    = 14 * 1024 * 1024;                                  // 디코드 前 base64 길이 컷(폭탄 방어)
var BK_DOC_EXT        = { "application/pdf":".pdf","image/jpeg":".jpg","image/jpg":".jpg","image/png":".png","image/webp":".webp","image/heic":".heic","image/heif":".heif" };

/* ───────────────────────── 공통 유틸 ───────────────────────── */

// 토큰 검사(속성 TOKENS 목록에 포함되면 통과)
function _tokOK(t){ return _tokens().indexOf(String(t||"")) >= 0; }

// http(s) URL만 통과, 그 외 "" (길이도 컷)
function _urlOrBlank(u){ u = String(u||""); return BK_URL_RE.test(u) ? u.slice(0,2000) : ""; }

function _out(obj, cb){
  var s = JSON.stringify(obj);
  if(cb){ return ContentService.createTextOutput(cb+"("+s+")").setMimeType(ContentService.MimeType.JAVASCRIPT); }
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}

// sheet 파라미터 정규화: 미지정/미상 → "expenses"(하위호환 사수)
function _route(v){ v = String(v||""); return (v === "bookings" || v === "places" || v === "itinerary" || v === "fund") ? v : "expenses"; }
// 숫자 범위 검증(밖이면 "" 반환 → 저장 안 함)
function _numRange(v, lo, hi){ var n = Number(v); return (n >= lo && n <= hi) ? n : ""; }

/* ── 감사로그("로그" 시트) — 로그 실패가 본 동작을 깨지 않도록 전부 try/catch ── */
var LOG_SHEET_NAME = "로그";
function _logSheet(){
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(LOG_SHEET_NAME);
  if(!sh){ sh = ss.insertSheet(LOG_SHEET_NAME); sh.appendRow(["ts","action","sheet","id","token","detail"]); }
  return sh;
}
function _audit(action, sheetName, id, token, detail){
  try {
    _logSheet().appendRow([new Date(), String(action||""), String(sheetName||""), String(id||""),
                           String(token||"").slice(-6), String(detail||"").slice(0,180)]);
  } catch(e){}
}
/* 토큰별 마지막 사용(시간 단위, 바뀔 때만 "seen" 1행) — 로테이션 전 구토큰 트래픽 확인용 */
function _seenToken(token){
  try {
    var sfx = String(token||"").slice(-6);
    var hour = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Seoul", "yyyy-MM-dd'T'HH");
    var p = PropertiesService.getScriptProperties(), k = "SEEN_" + sfx;
    if(p.getProperty(k) !== hour){ p.setProperty(k, hour); _audit("seen", "-", "-", token, ""); }
  } catch(e){}
}

/* ───────────────────────── 경비(expenses) ───────────────────────── */

function _sheet(){
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET);
  if(!sh){ sh = ss.insertSheet(SHEET); sh.appendRow(["id","date","cat","amount","payer","parts","memo","ts","receipt","deleted","krw","splits","groups"]); return sh; }
  if(String(sh.getRange(1,10).getValue()) === ""){ sh.getRange(1,10).setValue("deleted"); }  // 소프트삭제 컬럼(10열) 보강 — 빈 셀일 때만(수동 확장열 덮어쓰기 방지)
  if(String(sh.getRange(1,11).getValue()) === ""){ sh.getRange(1,11).setValue("krw"); }        // 충전 실지불 원화(11열) 보강
  if(String(sh.getRange(1,12).getValue()) === ""){ sh.getRange(1,12).setValue("splits"); }      // 커스텀 분할(12열) 보강 — "멤버:엔|멤버:엔"
  if(String(sh.getRange(1,13).getValue()) === ""){ sh.getRange(1,13).setValue("groups"); }      // 그룹 분배(13열) 보강 — "멤버:그룹번호|…"(수정 복원용)
  return sh;
}
function _items(){
  var v = _sheet().getDataRange().getValues(), out = [];
  for(var i=1;i<v.length;i++){ var r=v[i]; if(!r[0] || r[9]) continue;   // 소프트삭제 행 제외
    out.push({ id:String(r[0]), date:String(r[1]), cat:String(r[2]), amount:Number(r[3])||0,
               payer:String(r[4]), parts:String(r[5]||"").split("|").filter(String), memo:String(r[6]||""),
               receipt:String(r[8]||""), krw:(r[10]===""||r[10]==null)?null:Number(r[10]),   // 충전 실지불 원화(없으면 null)
               splits:String(r[11]||""), groups:String(r[12]||"") });   // splits=커스텀 분할, groups=그룹 배정(복원용)
  }
  return out;
}
function _receiptFolder(){
  var it = DriveApp.getFoldersByName(RECEIPT_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(RECEIPT_FOLDER);
}
// 커스텀 분할 정화: "멤버:정수엔" 쌍만 통과(최대 12명), 그 외 "" → 균등 분할
function _splits(s){
  var out = [];
  String(s||"").split("|").forEach(function(pair){
    var m = String(pair).split(":");
    var name = String(m[0]||"").replace(/[|:]/g,"").slice(0,20);
    var val = Math.round(Number(m[1]));
    if(name && val>=0 && val<=20000000 && out.length<12) out.push(name+":"+val);
  });
  return out.join("|");
}
// 그룹 배정 정화: "멤버:그룹번호(1~12)" 쌍만 통과(수정 복원용)
function _groups(s){
  var out = [];
  String(s||"").split("|").forEach(function(pair){
    var m = String(pair).split(":");
    var name = String(m[0]||"").replace(/[|:]/g,"").slice(0,20);
    var g = parseInt(m[1], 10);
    if(name && g>=1 && g<=12 && out.length<12) out.push(name+":"+g);
  });
  return out.join("|");
}
// receipt(9열) 갱신 — URL 화이트리스트 통과 + "빈 셀일 때만"(기존 영수증 덮어쓰기 방지)
function _setReceipt(id, url){
  if(!RECEIPT_URL_RE.test(String(url||""))) return false;
  var sh = _sheet(), v = sh.getDataRange().getValues();
  for(var i=1;i<v.length;i++){
    if(v[i][9]) continue;                              // 소프트삭제 행 제외
    if(String(v[i][0]) === String(id)){
      if(String(v[i][8]||"") !== "") return false;     // 이미 있으면 보존
      sh.getRange(i+1, 9).setValue(String(url)); return true;
    }
  }
  return false;
}

// 경비: GET(list) — 기존 동작 그대로
function _expensesGet(p, cb){
  return _out({ok:true, items:_items()}, cb);
}

// 경비: POST(upload/add/update/delete) — 기존 동작 그대로
function _expensesPost(d){
  var sh = _sheet();

  // ── 영수증 업로드 ──
  if(d.action === "upload"){
    if(!RECEIPT_MIME_RE.test(String(d.mimeType||""))) return _out({ok:false, error:"bad mime"});
    var raw = String(d.dataBase64||"").replace(/^data:[^,]*,/, "");        // 데이터URL 접두어 방어
    if(!raw) return _out({ok:false, error:"empty"});
    if(raw.length > RECEIPT_B64_MAX) return _out({ok:false, error:"too large"}); // 디코드 前 1차 컷(폭탄 방어)
    var bytes;
    try { bytes = Utilities.base64Decode(raw); } catch(e3){ return _out({ok:false, error:"bad base64"}); }
    if(bytes.length > RECEIPT_MAX_BYTES) return _out({ok:false, error:"too large"}); // 실바이트 최종 게이트

    var safe = String(d.filename||"receipt").replace(/[^\w.\-가-힣]/g,"_").replace(/\.{2,}/g,".").slice(0,80) || "receipt";
    var ext = ({ "image/jpeg":".jpg", "image/jpg":".jpg", "image/png":".png",
                 "image/webp":".webp", "image/heic":".heic", "image/heif":".heif" })[d.mimeType] || "";
    if(ext && safe.slice(-ext.length).toLowerCase() !== ext) safe += ext;

    var folder = _receiptFolder();
    var f = folder.createFile(Utilities.newBlob(bytes, d.mimeType, safe));
    try { f.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch(e4){} // 비공개 강제
    var url = f.getUrl(), name = f.getName();
    if(d.id) _setReceipt(String(d.id), url);       // 기존 항목이면 즉시 연결(신규는 add가 함께 기록)
    _audit("upload", "expenses", d.id||"", d.token, f.getId());
    return _out({ok:true, url:url, id:f.getId(), name:name});
  }

  // ── 항목 추가/수정 ──
  if(d.action === "add" || d.action === "update"){
    var en = d.entry || {};
    if(CATS.indexOf(en.cat) < 0) return _out({ok:false, error:"bad cat"});
    var amt = Number(en.amount); if(!(amt>=0 && amt<=20000000)) return _out({ok:false, error:"bad amount"});  // 상한 2천만엔(대형 숙박·그룹 항공·충전 커버)
    var memo  = String(en.memo||"").slice(0,200);
    var parts = (en.parts||[]).join("|");
    var rc = RECEIPT_URL_RE.test(String(en.receipt||"")) ? String(en.receipt) : "";  // 화이트리스트 통과분만
    var krw = (en.krw!=null && Number(en.krw)>=0 && Number(en.krw)<=100000000) ? Math.round(Number(en.krw)) : "";  // 충전 실지불 원화(선택)
    var splits = _splits(en.splits);   // 커스텀 분할 정화("멤버:엔|…"), 빈값이면 균등
    var groups = _groups(en.groups);   // 그룹 배정 정화("멤버:그룹번호|…"), 수정 복원용
    // 10열=deleted 자리는 ""(살아있는 행 갱신/신규라 무삭제), 11열=krw, 12열=splits, 13열=groups
    var row = [String(en.id), String(en.date), en.cat, amt, String(en.payer||""), parts, memo, new Date(), rc, "", krw, splits, groups];
    var v = sh.getDataRange().getValues();
    for(var i=1;i<v.length;i++){
      if(v[i][9]) continue;                                                 // 소프트삭제 행은 없는 것으로(부활 방지)
      if(String(v[i][0]) === String(en.id)){ sh.getRange(i+1,1,1,row.length).setValues([row]); _audit(d.action,"expenses",en.id,d.token,en.cat+" ¥"+amt); return _out({ok:true, id:en.id, receipt:rc}); }
    }
    if(d.action === "update") return _out({ok:false, error:"not found"});   // 삭제된 항목 부활 방지(upsert 금지)
    sh.appendRow(row); _audit("add","expenses",en.id,d.token,en.cat+" ¥"+amt); return _out({ok:true, id:en.id, receipt:rc});
  }

  // ── 삭제 (소프트삭제: deleted=1 마킹 — 행·드라이브 파일 보존, 복구 가능) ──
  if(d.action === "delete"){
    var v2 = sh.getDataRange().getValues();
    for(var j=1;j<v2.length;j++){
      if(v2[j][9]) continue;
      if(String(v2[j][0]) === String(d.id)){
        sh.getRange(j+1,10).setValue(1);
        _audit("delete","expenses",d.id,d.token, String(v2[j][2])+" ¥"+String(v2[j][3])+" "+String(v2[j][4]));
        return _out({ok:true});
      }
    }
    return _out({ok:false, error:"not found"});
  }

  return _out({ok:false, error:"bad action"});
}

/* ───────────────────────── 예약(bookings) ───────────────────────── */

function _bkSheet(){
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(BK_SHEET);
  if(!sh){ sh = ss.insertSheet(BK_SHEET); sh.appendRow(BK_HEADERS.concat(["deleted"])); return sh; }   // 없으면 자동 생성(헤더+소프트삭제)
  if(sh.getLastColumn() < BK_HEADERS.length){ sh.getRange(1, BK_HEADERS.length).setValue("attachments"); }  // 구버전 시트 헤더 보강
  if(String(sh.getRange(1, BK_HEADERS.length+1).getValue()) === ""){ sh.getRange(1, BK_HEADERS.length+1).setValue("deleted"); }  // 소프트삭제 컬럼(12열) 보강 — 빈 셀일 때만(수동 확장열 덮어쓰기 방지)
  return sh;
}
function _bkItems(){
  var v = _bkSheet().getDataRange().getValues(), out = [];
  // 열: 0 id|1 type|2 title|3 datetime|4 sub|5 detail|6 voucherUrl|7 mapUrl|8 memo|9 ts|10 attachments
  for(var i=1;i<v.length;i++){ var r=v[i]; if(!r[0] || r[11]) continue;   // 소프트삭제 행 제외
    out.push({ id:String(r[0]), type:String(r[1]||""), title:String(r[2]||""), datetime:String(r[3]||""),
               sub:String(r[4]||""), detail:String(r[5]||""), voucherUrl:String(r[6]||""),
               mapUrl:String(r[7]||""), memo:String(r[8]||""), attachments:String(r[10]||"") });
  }
  return out;
}
/* attachments 정화: [{fileId,mime,size}]만 통과(시트에 민감값 저장 방지) */
function _bkAtt(s){
  var arr; try{ arr = JSON.parse(String(s||"[]")); }catch(e){ return "[]"; }
  if(!(arr instanceof Array)) return "[]";
  var out=[];
  for(var i=0;i<arr.length && out.length<8;i++){
    var a=arr[i]||{};
    var fid=String(a.fileId||"").replace(/[^A-Za-z0-9_\-]/g,"").slice(0,120); if(!fid) continue;
    var mime=String(a.mime||""); if(!BK_DOC_MIME_RE.test(mime)) mime="";
    var size=Number(a.size); if(!(size>=0)) size=0;
    out.push({fileId:fid, mime:mime, size:size});
  }
  return JSON.stringify(out);
}

// 예약: GET(list)
function _bookingsGet(p, cb){
  return _out({ok:true, items:_bkItems()}, cb);
}

// 예약: POST(add/update/delete)
function _bookingsPost(d){
  var sh = _bkSheet();

  if(d.action === "list"){                                  // POST로 list를 보내도 허용(관용)
    return _out({ok:true, items:_bkItems()});
  }

  // ── 첨부 문서 업로드(PDF/이미지) → booking-docs 제한폴더 ──
  if(d.action === "uploadDoc"){
    if(!BK_DOC_MIME_RE.test(String(d.mimeType||""))) return _out({ok:false, error:"bad mime"});
    var raw = String(d.dataBase64||"").replace(/^data:[^,]*,/, "");
    if(!raw) return _out({ok:false, error:"empty"});
    if(raw.length > BK_DOC_B64_MAX) return _out({ok:false, error:"too large"});     // 디코드 前 1차 컷
    var dbytes; try { dbytes = Utilities.base64Decode(raw); } catch(e6){ return _out({ok:false, error:"bad base64"}); }
    if(dbytes.length > BK_DOC_MAX_BYTES) return _out({ok:false, error:"too large"}); // 실바이트 최종 게이트
    if(!BK_DOCS_FOLDER_ID || BK_DOCS_FOLDER_ID === "PASTE_FOLDER_ID_HERE") return _out({ok:false, error:"folder not set"});
    var dfolder; try { dfolder = DriveApp.getFolderById(BK_DOCS_FOLDER_ID); } catch(e7){ return _out({ok:false, error:"folder error"}); }
    var dext  = BK_DOC_EXT[String(d.mimeType)] || "";
    var dbase = String(d.id||"doc").replace(/[^\w\-]/g,"_").slice(0,24) || "doc";
    var dsafe = dbase + "__" + Utilities.getUuid().slice(0,8) + dext;   // 실명·여권번호 없는 불투명 파일명
    var df = dfolder.createFile(Utilities.newBlob(dbytes, String(d.mimeType), dsafe));
    // ★ setSharing 호출 안 함 → 폴더의 제한공유(4명 뷰어)를 상속. anyone-with-link 절대 설정 금지.
    _audit("uploadDoc", "bookings", d.id||"", d.token, df.getId());
    return _out({ok:true, fileId:df.getId(), url:"https://drive.google.com/file/d/"+df.getId()+"/view", mime:String(d.mimeType), size:dbytes.length});
  }

  if(d.action === "add" || d.action === "update"){
    var en = d.entry || {};
    if(BK_TYPES.indexOf(String(en.type)) < 0) return _out({ok:false, error:"bad type"});
    var id       = String(en.id||"");
    if(!id) return _out({ok:false, error:"bad id"});
    var title    = String(en.title||"").slice(0, BK_LEN.title);
    var datetime = String(en.datetime||"").slice(0, 40);     // ISO 등 짧은 문자열
    var sub      = String(en.sub||"").slice(0, BK_LEN.sub);
    var detail   = String(en.detail||"").slice(0, BK_LEN.detail);
    var voucher  = _urlOrBlank(en.voucherUrl);               // http(s)만, 그 외 ""
    var mapUrl   = _urlOrBlank(en.mapUrl);                   // http(s)만, 그 외 ""
    var memo     = String(en.memo||"").slice(0, BK_LEN.memo);
    var attachments = _bkAtt(en.attachments);                // [{fileId,mime,size}]만 정화 저장
    // 열 순서(BK_HEADERS)와 정확히 일치
    var row = [id, String(en.type), title, datetime, sub, detail, voucher, mapUrl, memo, new Date(), attachments];
    var v = sh.getDataRange().getValues();
    for(var i=1;i<v.length;i++){
      if(v[i][11]) continue;                                                // 소프트삭제 행은 없는 것으로(부활 방지)
      if(String(v[i][0]) === id){ sh.getRange(i+1,1,1,row.length).setValues([row]); _audit(d.action,"bookings",id,d.token,String(en.type)+" "+title); return _out({ok:true, id:id}); }
    }
    if(d.action === "update") return _out({ok:false, error:"not found"});   // 삭제된 예약 부활 방지(upsert 금지)
    sh.appendRow(row); _audit("add","bookings",id,d.token,String(en.type)+" "+title); return _out({ok:true, id:id});
  }

  if(d.action === "delete"){
    var v2 = sh.getDataRange().getValues();
    for(var j=1;j<v2.length;j++){
      if(v2[j][11]) continue;
      if(String(v2[j][0]) === String(d.id)){
        sh.getRange(j+1, BK_HEADERS.length+1).setValue(1);
        _audit("delete","bookings",d.id,d.token, String(v2[j][1])+" "+String(v2[j][2]));
        return _out({ok:true});
      }
    }
    return _out({ok:false, error:"not found"});
  }

  return _out({ok:false, error:"bad action"});
}

/* ───────────────────────── 맛집후보(places) ───────────────────────── */

function _plSheet(){
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(PL_SHEET);
  if(!sh){ sh = ss.insertSheet(PL_SHEET); sh.appendRow(PL_HEADERS.concat(["deleted"])); return sh; }
  if(String(sh.getRange(1, PL_HEADERS.length+1).getValue()) === ""){ sh.getRange(1, PL_HEADERS.length+1).setValue("deleted"); }  // 소프트삭제 컬럼(11열) 보강
  return sh;
}
function _plItems(){
  var v = _plSheet().getDataRange().getValues(), out = [];
  // 열: 0 id|1 name|2 region|3 lat|4 lng|5 foundBy|6 url|7 memo|8 status|9 ts|10 deleted
  for(var i=1;i<v.length;i++){ var r=v[i]; if(!r[0] || r[10]) continue;   // 소프트삭제 행 제외
    out.push({ id:String(r[0]), name:String(r[1]||""), region:String(r[2]||""),
               lat:(r[3]===""||r[3]==null)?null:Number(r[3]), lng:(r[4]===""||r[4]==null)?null:Number(r[4]),
               foundBy:String(r[5]||""), url:String(r[6]||""), memo:String(r[7]||""), status:String(r[8]||"candidate") });
  }
  return out;
}

// 맛집후보: GET(list)
function _placesGet(p, cb){ return _out({ok:true, items:_plItems()}, cb); }

// 맛집후보: POST(add/update/delete)
function _placesPost(d){
  var sh = _plSheet();

  if(d.action === "list"){ return _out({ok:true, items:_plItems()}); }

  if(d.action === "add" || d.action === "update"){
    var en = d.entry || {};
    var id = String(en.id||""); if(!id) return _out({ok:false, error:"bad id"});
    var name = String(en.name||"").slice(0, PL_LEN.name); if(!name) return _out({ok:false, error:"bad name"});
    var region  = String(en.region||"").slice(0, PL_LEN.region);
    var lat = _numRange(en.lat, 24, 46);       // 일본 위도
    var lng = _numRange(en.lng, 122, 154);     // 일본 경도
    if(lat === "" || lng === ""){ lat = ""; lng = ""; }  // 한쪽만 있으면 둘 다 버림(마커 불가)
    var foundBy = String(en.foundBy||"").slice(0, 12);
    var url  = BK_URL_RE.test(String(en.url||"")) ? String(en.url).slice(0,2000) : "";  // http(s)만
    var memo = String(en.memo||"").slice(0, PL_LEN.memo);
    var status = (String(en.status||"") === "scored") ? "scored" : "candidate";
    var row = [id, name, region, lat, lng, foundBy, url, memo, status, new Date()];
    var v = sh.getDataRange().getValues();
    for(var i=1;i<v.length;i++){
      if(v[i][10]) continue;                                                // 소프트삭제 행은 없는 것으로(부활 방지)
      if(String(v[i][0]) === id){ sh.getRange(i+1,1,1,row.length).setValues([row]); _audit(d.action,"places",id,d.token,name); return _out({ok:true, id:id}); }
    }
    if(d.action === "update") return _out({ok:false, error:"not found"});
    sh.appendRow(row); _audit("add","places",id,d.token,name+(foundBy?(" ·"+foundBy):"")); return _out({ok:true, id:id});
  }

  if(d.action === "delete"){
    var v2 = sh.getDataRange().getValues();
    for(var j=1;j<v2.length;j++){
      if(v2[j][10]) continue;
      if(String(v2[j][0]) === String(d.id)){
        sh.getRange(j+1, PL_HEADERS.length+1).setValue(1);
        _audit("delete","places",d.id,d.token, String(v2[j][1]));
        return _out({ok:true});
      }
    }
    return _out({ok:false, error:"not found"});
  }

  return _out({ok:false, error:"bad action"});
}

/* ───────────────────────── 일정 오버레이(itinerary) ───────────────────────── */

function _itSheet(){
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(IT_SHEET);
  if(!sh){ sh = ss.insertSheet(IT_SHEET); sh.appendRow(IT_HEADERS.concat(["deleted"])); return sh; }
  if(String(sh.getRange(1, IT_HEADERS.length+1).getValue()) === ""){ sh.getRange(1, IT_HEADERS.length+1).setValue("deleted"); }  // 소프트삭제 컬럼(11열) 보강
  return sh;
}
function _itDate(v){  // 시트가 Date로 자동 변환한 값 → "yyyy-MM-dd" 정규화
  if(Object.prototype.toString.call(v) === "[object Date]") return Utilities.formatDate(v, Session.getScriptTimeZone() || "Asia/Seoul", "yyyy-MM-dd");
  return String(v||"");
}
function _itTime(v){  // Date/문자열 → "HH:mm" 정규화
  if(Object.prototype.toString.call(v) === "[object Date]") return Utilities.formatDate(v, Session.getScriptTimeZone() || "Asia/Seoul", "HH:mm");
  var m = String(v||"").match(/(\d{1,2}):(\d{2})/); return m ? (("0"+m[1]).slice(-2) + ":" + m[2]) : "";
}
function _itItems(){
  var v = _itSheet().getDataRange().getValues(), out = [];
  // 0 id|1 date|2 time|3 name|4 lat|5 lng|6 memo|7 category|8 createdBy|9 ts|10 deleted
  for(var i=1;i<v.length;i++){ var r=v[i]; if(!r[0] || r[10]) continue;   // 소프트삭제 행 제외
    out.push({ id:String(r[0]), date:_itDate(r[1]), time:_itTime(r[2]), name:String(r[3]||""),
               lat:(r[4]===""||r[4]==null)?null:Number(r[4]), lng:(r[5]===""||r[5]==null)?null:Number(r[5]),
               memo:String(r[6]||""), category:String(r[7]||""), createdBy:String(r[8]||"") });
  }
  return out;
}

// 일정 오버레이: GET(list)
function _itineraryGet(p, cb){ return _out({ok:true, items:_itItems()}, cb); }

// 일정 오버레이: POST(add/update/delete)
function _itineraryPost(d){
  var sh = _itSheet();

  if(d.action === "list"){ return _out({ok:true, items:_itItems()}); }

  if(d.action === "add" || d.action === "update"){
    var en = d.entry || {};
    var id = String(en.id||""); if(!id) return _out({ok:false, error:"bad id"});
    var date = String(en.date||""); if(!IT_DATE_RE.test(date)) return _out({ok:false, error:"bad date"});
    var name = String(en.name||"").slice(0, IT_LEN.name); if(!name) return _out({ok:false, error:"bad name"});
    var time = IT_TIME_RE.test(String(en.time||"")) ? String(en.time) : "";
    var lat = _numRange(en.lat, 24, 46);
    var lng = _numRange(en.lng, 122, 154);
    if(lat === "" || lng === ""){ lat = ""; lng = ""; }
    var memo = String(en.memo||"").slice(0, IT_LEN.memo);
    var category = String(en.category||"").slice(0, 12);
    var createdBy = String(en.createdBy||"").slice(0, 12);
    var row = [id, date, time, name, lat, lng, memo, category, createdBy, new Date()];
    var v = sh.getDataRange().getValues();
    for(var i=1;i<v.length;i++){
      if(v[i][10]) continue;
      if(String(v[i][0]) === id){ sh.getRange(i+1,1,1,row.length).setValues([row]); _audit(d.action,"itinerary",id,d.token,date+" "+name); return _out({ok:true, id:id}); }
    }
    if(d.action === "update") return _out({ok:false, error:"not found"});
    sh.appendRow(row); _audit("add","itinerary",id,d.token,date+" "+name); return _out({ok:true, id:id});
  }

  if(d.action === "delete"){
    var v2 = sh.getDataRange().getValues();
    for(var j=1;j<v2.length;j++){
      if(v2[j][10]) continue;
      if(String(v2[j][0]) === String(d.id)){
        sh.getRange(j+1, IT_HEADERS.length+1).setValue(1);
        _audit("delete","itinerary",d.id,d.token, String(v2[j][1])+" "+String(v2[j][3]));
        return _out({ok:true});
      }
    }
    return _out({ok:false, error:"not found"});
  }

  return _out({ok:false, error:"bad action"});
}

/* ───────────────────────── 회비(fund) ───────────────────────── */

function _fundSheet(){
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(FUND_SHEET);
  if(!sh){ sh = ss.insertSheet(FUND_SHEET); sh.appendRow(FUND_HEADERS.concat(["deleted"])); return sh; }
  if(String(sh.getRange(1, FUND_HEADERS.length+1).getValue()) === ""){ sh.getRange(1, FUND_HEADERS.length+1).setValue("deleted"); }  // 소프트삭제 컬럼(7열) 보강
  return sh;
}
function _fundItems(){
  var v = _fundSheet().getDataRange().getValues(), out = [];
  // 0 id|1 date|2 krw|3 payer|4 memo|5 ts|6 deleted
  for(var i=1;i<v.length;i++){ var r=v[i]; if(!r[0] || r[6]) continue;   // 소프트삭제 행 제외
    out.push({ id:String(r[0]), date:_itDate(r[1]), krw:Number(r[2])||0, payer:String(r[3]||""), memo:String(r[4]||"") });
  }
  return out;
}

// 회비: GET(list)
function _fundGet(p, cb){ return _out({ok:true, items:_fundItems()}, cb); }

// 회비: POST(add/update/delete)
function _fundPost(d){
  var sh = _fundSheet();

  if(d.action === "list"){ return _out({ok:true, items:_fundItems()}); }

  if(d.action === "add" || d.action === "update"){
    var en = d.entry || {};
    var id = String(en.id||""); if(!id) return _out({ok:false, error:"bad id"});
    var date = String(en.date||""); if(!IT_DATE_RE.test(date)) return _out({ok:false, error:"bad date"});
    var krw = Number(en.krw); if(!(krw>0 && krw<=100000000)) return _out({ok:false, error:"bad krw"});  // 원화 입금액(양수 필수)
    krw = Math.round(krw);
    var payer = String(en.payer||"").slice(0, 12);
    var memo = String(en.memo||"").slice(0, FUND_LEN.memo);
    var row = [id, date, krw, payer, memo, new Date()];
    var v = sh.getDataRange().getValues();
    for(var i=1;i<v.length;i++){
      if(v[i][6]) continue;
      if(String(v[i][0]) === id){ sh.getRange(i+1,1,1,row.length).setValues([row]); _audit(d.action,"fund",id,d.token,"₩"+krw+(payer?(" "+payer):"")); return _out({ok:true, id:id}); }
    }
    if(d.action === "update") return _out({ok:false, error:"not found"});
    sh.appendRow(row); _audit("add","fund",id,d.token,"₩"+krw+(payer?(" "+payer):"")); return _out({ok:true, id:id});
  }

  if(d.action === "delete"){
    var v2 = sh.getDataRange().getValues();
    for(var j=1;j<v2.length;j++){
      if(v2[j][6]) continue;
      if(String(v2[j][0]) === String(d.id)){
        sh.getRange(j+1, FUND_HEADERS.length+1).setValue(1);
        _audit("delete","fund",d.id,d.token, "₩"+String(v2[j][2]));
        return _out({ok:true});
      }
    }
    return _out({ok:false, error:"not found"});
  }

  return _out({ok:false, error:"bad action"});
}

/* ───────────────────────── 엔트리포인트 ───────────────────────── */

function doGet(e){
  var p = e.parameter || {}, cb = p.callback;
  if(!_tokOK(p.token)) return _out({ok:false, error:"bad token"}, cb);   // 속성 TOKENS 목록 기준
  _seenToken(p.token);
  try {
    // sheet 미지정 시 "expenses"로 폴백 → 기존 경비/잔액/영수증 호출(sheet 없이 옴) 안 깨짐
    if(_route(p.sheet) === "bookings") return _bookingsGet(p, cb);
    if(_route(p.sheet) === "places") return _placesGet(p, cb);
    if(_route(p.sheet) === "itinerary") return _itineraryGet(p, cb);
    if(_route(p.sheet) === "fund") return _fundGet(p, cb);
    return _expensesGet(p, cb);
  } catch(err){ try{ console.error(err); }catch(e2){} return _out({ok:false, error:"server error"}, cb); }
}

function doPost(e){
  var d; try{ d = JSON.parse(e.postData.contents); }catch(err){ return _out({ok:false, error:"bad json"}); }
  if(!_tokOK(d.token)) return _out({ok:false, error:"bad token"});       // 속성 TOKENS 목록 기준
  _seenToken(d.token);
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(eLock){ return _out({ok:false, error:"busy"}); }  // 동시 쓰기 직렬화(오삭제·덮어쓰기 방지)
  try {
    // sheet 미지정 시 "expenses"로 폴백(하위호환 사수)
    if(_route(d.sheet) === "bookings") return _bookingsPost(d);
    if(_route(d.sheet) === "places") return _placesPost(d);
    if(_route(d.sheet) === "itinerary") return _itineraryPost(d);
    if(_route(d.sheet) === "fund") return _fundPost(d);
    return _expensesPost(d);
  } catch(err){ try{ console.error(err); }catch(e5){} return _out({ok:false, error:"server error"}); }
  finally { try{ lock.releaseLock(); }catch(eRel){} }
}

/* ───────────────────────── 유지보수(편집기에서 수동 실행) ─────────────────────────
 * cleanupOrphanFiles: 영수증·booking-docs 폴더에서 시트가 참조하지 않는 파일을 휴지통으로.
 *  - 소프트삭제된 행의 파일도 "참조 있음"으로 취급(복구 대비 보존).
 *  - 생성 1시간 이내 파일은 스킵(업로드→행 저장 사이 레이스 보호).
 *  - 휴지통 이동이라 30일 내 복구 가능. 실행: 함수 선택 → ▶실행, 결과는 "로그" 시트 확인.
 */
function cleanupOrphanFiles(){
  var refs = "";
  var ev = _sheet().getDataRange().getValues();
  for(var i=1;i<ev.length;i++){ refs += String(ev[i][8]||"") + "\n"; }        // 경비 receipt URL(파일ID 포함)
  var bv = _bkSheet().getDataRange().getValues();
  for(var j=1;j<bv.length;j++){ refs += String(bv[j][10]||"") + "\n"; }       // 예약 attachments JSON(fileId 포함)
  // 안전장치: 참조 fileId가 하나도 없으면(시트 이름변경·유실 등 이상 상태) 폴더 전체 오삭제 방지 위해 중단
  if(!/[A-Za-z0-9_-]{20,}/.test(refs)){ _audit("orphan-abort", "-", "-", "", "empty refs"); return -1; }
  var n = 0;
  n += _trashUnref(_receiptFolder(), refs);
  try { n += _trashUnref(DriveApp.getFolderById(BK_DOCS_FOLDER_ID), refs); } catch(e){}
  _audit("orphan-cleanup", "-", "-", "", "trashed=" + n);
  return n;
}
function _trashUnref(folder, refs){
  var it = folder.getFiles(), n = 0, now = new Date().getTime();
  while(it.hasNext()){
    var f = it.next();
    if(now - f.getDateCreated().getTime() < 3600000) continue;                // 1시간 이내 신규 파일 보호
    if(refs.indexOf(f.getId()) < 0){
      try { f.setTrashed(true); _audit("orphan-trash", "-", f.getId(), "", f.getName()); n++; } catch(e){}
    }
  }
  return n;
}
