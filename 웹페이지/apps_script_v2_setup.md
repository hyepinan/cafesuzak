# Apps Script v2 업그레이드 가이드 (매칭 기능 추가)

기존 후기 수집 기능 + 소개팅 신청 + 수작부리기 게시판까지 한 Apps Script로 처리합니다.

---

## STEP 1 · 구글 시트에 새 시트 3개 추가

기존 "카페수작 페이지 후기" 시트 파일 안에 시트 3개를 추가합니다.

### 1-1. 시트 추가 방법
- 시트 하단 좌측 **+** 버튼 클릭 → 새 시트 생성
- 시트 탭 우클릭 → **이름 바꾸기** 로 정확한 이름 입력

### 1-2. 새 시트 이름 (정확히 입력!)
| 시트 이름 | 용도 |
|-----------|------|
| `applications` | 소개팅 신청서 |
| `suzak_board` | 수작부리기 게시글 |
| `suzak_replies` | 수작부리기 답장 |

### 1-3. 각 시트 헤더 (1행에 입력)

**applications 시트 1행** (탭 구분, 복붙 가능):
```
제출시각	이름	연락처	나이	학과	이상형	MBTI	어필	결제시각	익명ID	공개여부
```

> 💡 `결제시각`은 신청자가 카페수작에 방문한 결제 시간 (예: "11월 18일 14:35"). 카드사 결제 내역 / 카페 POS 내역으로 본인이 직접 매칭 확인하시면 됩니다.
> 💡 `공개여부`는 신청자가 익명 카드 공개에 동의했는지 ("yes" 또는 빈칸).

**suzak_board 시트 1행**:
```
게시ID	작성시각	내용	작성자닉	비밀번호해시	답장수
```

**suzak_replies 시트 1행**:
```
답장ID	게시ID	답장시각	만남년월일	만남시간	답장자닉	메시지
```

> 💡 `메시지`는 답장 시 작성하는 짧은 한 마디 (선택 입력).

---

## STEP 2 · Apps Script 코드 전체 교체

기존 Apps Script 편집창에서 **코드 전체 삭제**하고 아래를 통째로 붙여넣기:

```
// ========================================
//   Cafe Suzak — 통합 백엔드 (v2)
// ========================================
const FEEDBACK_SHEET = ""; // 첫 번째 시트 (기존 후기)
const APPLICATIONS_SHEET = "applications";
const BOARD_SHEET = "suzak_board";
const REPLIES_SHEET = "suzak_replies";

function doPost(e) {
  try {
    if (e.parameter["bot-field"]) return jsonResponse({result:"spam_ignored"}, e);
    const action = e.parameter.action || "feedback";
    let result;
    switch (action) {
      case "feedback":  result = saveFeedback(e); break;
      case "apply":     result = saveApplication(e); break;
      case "board_add": result = addBoardPost(e); break;
      case "reply_add": result = addReply(e); break;
      default: result = {result:"error", error:"unknown action: "+action};
    }
    return jsonResponse(result, e);
  } catch (err) {
    return jsonResponse({result:"error", error:err.toString()}, e);
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action || "status";
    let result;
    switch (action) {
      case "status":     result = {status:"ok", message:"Cafe Suzak v2 alive"}; break;
      case "list_anon":  result = listAnonApplications(); break;
      case "board_list": result = boardList(); break;
      case "reply_view": result = viewReply(e); break;
      default: result = {result:"error", error:"unknown action: "+action};
    }
    return jsonResponse(result, e);
  } catch (err) {
    return jsonResponse({result:"error", error:err.toString()}, e);
  }
}

function jsonResponse(data, e) {
  const callback = e && e.parameter && e.parameter.callback;
  const json = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return name ? ss.getSheetByName(name) : ss.getSheets()[0];
}

function sha256(str) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(str || ""),
    Utilities.Charset.UTF_8
  );
  return bytes.map(b => ("0" + (b & 0xff).toString(16)).slice(-2)).join("");
}

function uid() {
  return Utilities.getUuid().replace(/-/g, "").slice(0, 12);
}

// ========================================
//   기존 후기 저장 (기능 유지)
// ========================================
function saveFeedback(e) {
  const sheet = getSheet(FEEDBACK_SHEET);
  const cols = [
    () => new Date(),
    () => e.parameter.rating || "",
    () => e.parameter.impression || "",
    () => e.parameter.feature_mood || "",
    () => e.parameter.feature_weather || "",
    () => e.parameter.feature_status || "",
    () => e.parameter.feature_menu || "",
    () => e.parameter.feature_visit || "",
    () => e.parameter.feature_design || "",
    () => e.parameter.message || "",
    () => e.parameter.improve || "",
    () => e.parameter.visited || "",
    () => e.parameter.cafe_message || "",
    () => e.parameter.favorite_menu || "",
    () => e.parameter.nickname || "",
    () => e.parameter.contact || "",
  ];
  sheet.appendRow(cols.map(fn => fn()));
  return {result:"success"};
}

// ========================================
//   소개팅 신청서 저장
// ========================================
function saveApplication(e) {
  const sheet = getSheet(APPLICATIONS_SHEET);
  const anonId = uid();
  const isPublic = (e.parameter.is_public || "").toLowerCase() === "yes" ? "yes" : "";
  sheet.appendRow([
    new Date(),
    e.parameter.name || "",
    e.parameter.contact || "",
    e.parameter.age || "",
    e.parameter.department || "",
    e.parameter.ideal_type || "",
    e.parameter.mbti || "",
    e.parameter.appeal || "",
    e.parameter.payment_time || "",
    anonId,
    isPublic,
  ]);
  return {result:"success", anonId:anonId};
}

// 익명 신청자 목록 (공개 동의한 신청자만 / 이름/연락처/결제시각 제외)
function listAnonApplications() {
  const sheet = getSheet(APPLICATIONS_SHEET);
  const last = sheet.getLastRow();
  if (last < 2) return {result:"success", items:[]};
  const data = sheet.getRange(2, 1, last - 1, 11).getValues();
  const items = data
    .filter(row => row[9] && String(row[10] || "").toLowerCase() === "yes")
    .map(row => ({
      timestamp: row[0],
      age: row[3],
      department: row[4],
      ideal_type: row[5],
      mbti: row[6],
      appeal: row[7],
      anonId: row[9],
    }))
    .reverse();
  return {result:"success", items:items};
}

// ========================================
//   수작부리기 게시판 — 글 등록
// ========================================
function addBoardPost(e) {
  const sheet = getSheet(BOARD_SHEET);
  const content = (e.parameter.content || "").trim();
  const nickname = (e.parameter.nickname || "").trim();
  const password = (e.parameter.password || "").trim();
  if (!content || !nickname || !password) {
    return {result:"error", error:"필수 항목 누락"};
  }
  const id = uid();
  sheet.appendRow([
    id,
    new Date(),
    content,
    nickname,
    sha256(password),
    0,
  ]);
  return {result:"success", id:id};
}

// 게시글 목록 (비밀번호 해시 제외)
function boardList() {
  const sheet = getSheet(BOARD_SHEET);
  const last = sheet.getLastRow();
  if (last < 2) return {result:"success", items:[]};
  const data = sheet.getRange(2, 1, last - 1, 6).getValues();
  const items = data.map(row => ({
    id: row[0],
    timestamp: row[1],
    content: row[2],
    nickname: row[3],
    replyCount: row[5] || 0,
  })).filter(item => item.id).reverse();
  return {result:"success", items:items};
}

// ========================================
//   답장 등록 — "저인거같아요"
// ========================================
function addReply(e) {
  const repliesSheet = getSheet(REPLIES_SHEET);
  const boardSheet = getSheet(BOARD_SHEET);
  const boardId = (e.parameter.board_id || "").trim();
  const meetDate = (e.parameter.meet_date || "").trim();
  const meetTime = (e.parameter.meet_time || "").trim();
  const replyNick = (e.parameter.reply_nickname || "").trim();
  const replyMessage = (e.parameter.reply_message || "").trim();
  if (!boardId || !meetDate || !meetTime || !replyNick) {
    return {result:"error", error:"필수 항목 누락"};
  }
  repliesSheet.appendRow([
    uid(),
    boardId,
    new Date(),
    meetDate,
    meetTime,
    replyNick,
    replyMessage,
  ]);
  // 답장수 +1
  const last = boardSheet.getLastRow();
  if (last >= 2) {
    const ids = boardSheet.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === boardId) {
        const cell = boardSheet.getRange(i + 2, 6);
        cell.setValue((cell.getValue() || 0) + 1);
        break;
      }
    }
  }
  return {result:"success"};
}

// ========================================
//   답장 보기 (비밀번호 검증)
// ========================================
function viewReply(e) {
  const boardSheet = getSheet(BOARD_SHEET);
  const repliesSheet = getSheet(REPLIES_SHEET);
  const boardId = (e.parameter.board_id || "").trim();
  const password = (e.parameter.password || "").trim();
  if (!boardId || !password) return {result:"error", error:"필수 항목 누락"};

  // 1) 게시글 찾기 + 비밀번호 검증
  const last = boardSheet.getLastRow();
  if (last < 2) return {result:"error", error:"게시글 없음"};
  const data = boardSheet.getRange(2, 1, last - 1, 5).getValues();
  let foundHash = null;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === boardId) { foundHash = data[i][4]; break; }
  }
  if (!foundHash) return {result:"error", error:"게시글 없음"};
  if (foundHash !== sha256(password)) {
    return {result:"error", error:"비밀번호 불일치"};
  }

  // 2) 답장 가져오기
  const lastR = repliesSheet.getLastRow();
  if (lastR < 2) return {result:"success", items:[]};
  const replies = repliesSheet.getRange(2, 1, lastR - 1, 7).getValues();
  const items = replies
    .filter(row => row[1] === boardId)
    .map(row => ({
      timestamp: row[2],
      meet_date: row[3],
      meet_time: row[4],
      reply_nickname: row[5],
      reply_message: row[6] || "",
    }))
    .reverse();
  return {result:"success", items:items};
}
```

붙여넣기 후 **Ctrl+S 저장**.

---

## STEP 3 · 새 버전으로 재배포

1. 우측 상단 **배포** → **배포 관리**
2. 우측 연필(✏) 아이콘 클릭
3. **버전** 드롭다운 → **새 버전** 선택
4. **배포** 클릭
5. URL은 그대로 유지 (`https://script.google.com/macros/s/.../exec`)

---

## STEP 4 · 동작 확인

브라우저 주소창에 다음 URL 입력해서 응답 확인:
```
https://script.google.com/macros/s/AKfycbwPHHbOIffeAJqlbjfyZ3NfoID2DPSJCWghTx3MgekMrbxY0U5D9SRopGtaIkjCD_2H8w/exec?action=status
```

화면에 `{"status":"ok","message":"Cafe Suzak v2 alive"}` 이런 응답 보이면 정상.

---

## 🔧 데이터 관리 팁

### 시트별 보는 법
- **applications**: 이름·연락처는 본인만 본다는 약속. 매칭 후 인스타로 직접 연결
- **suzak_board**: 게시글 자체는 누구나 볼 수 있는 공개 데이터
- **suzak_replies**: 답장 — 작성자 비밀번호로 조회 가능

### 부적절한 글 삭제
- 시트에서 해당 행 우클릭 → 삭제 → 시트에서 즉시 사라짐
- 페이지 새로고침하면 게시판에서도 안 보임

### 백업
- 파일 → 다운로드 → Microsoft Excel(.xlsx)로 주기적 백업 권장

---

준비 끝나면 알려주세요. 다음 단계는 `match.html` 페이지 제작입니다.
