# 후기 → 구글 시트 자동 저장 설정 가이드

이 가이드를 따라하면 카페수작 페이지 후기가 본인 구글 시트에 자동으로 한 줄씩 추가됩니다. 약 15분 소요.

---

## STEP 1 · 구글 시트 만들기

1. https://sheets.google.com 접속
2. 좌측 상단 **빈 스프레드시트** 클릭
3. 파일명을 `카페수작 페이지 후기` 로 변경 (좌측 상단에서 클릭하면 수정 가능)
4. **A1셀부터 첫 행에 컬럼명을 입력** (복붙 가능, 탭으로 구분돼 있어요):

```
제출시각	별점	첫인상	기능_메뉴추천	기능_날씨	기능_영업상태	기능_메뉴판	기능_매장정보	기능_디자인	페이지후기	개선점	방문경험	매장후기	자주시키는메뉴	닉네임	연락처
```

→ A1셀에 `제출시각`을 적고 Tab을 누르면 B1로 넘어가서 `별점`, 그렇게 P1까지 채우시면 돼요.
→ 또는 위 한 줄을 통째로 복사해서 A1에 붙여넣으면 자동으로 컬럼별로 분배됩니다.

---

## STEP 2 · Apps Script 열기

1. 같은 시트에서 상단 메뉴 **확장 프로그램** → **Apps Script** 클릭
2. 새 탭에 `코드.gs` 편집창이 열림
3. 기본으로 들어있는 `function myFunction() { ... }` 전부 **지우고**
4. 아래 코드를 **통째로 붙여넣기**:

```javascript
const SHEET_NAME = ""; // 비워두면 첫 번째 시트 사용

function doPost(e) {
  try {
    // 허니팟 — 봇 차단
    if (e.parameter["bot-field"]) {
      return json({result: "spam_ignored"});
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];

    // 컬럼 순서 — 시트의 1행 헤더와 동일하게 유지하세요
    const cols = [
      () => new Date(),               // 제출시각
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

    const row = cols.map(fn => fn());
    sheet.appendRow(row);

    return json({result: "success"});
  } catch (err) {
    return json({result: "error", error: err.toString()});
  }
}

function doGet(e) {
  return json({status: "ok", message: "Cafe Suzak feedback endpoint is alive."});
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
```

5. 좌측 상단 프로젝트 이름 `제목 없는 프로젝트` 클릭 → `카페수작 후기 수신` 으로 변경
6. 상단 디스켓 아이콘(💾) 클릭해서 저장 (Ctrl+S)

---

## STEP 3 · 웹앱으로 배포하기

1. 우측 상단 파란색 **배포** 버튼 → **새 배포** 클릭
2. 좌측 톱니바퀴(⚙) 아이콘 클릭 → **웹 앱** 선택
3. 아래 옵션 설정:
   - **설명**: `카페수작 후기 수신 v1` (자유)
   - **다음 사용자로 실행**: **나** (본인 구글 계정)
   - **액세스 권한이 있는 사용자**: **모든 사용자** ← 중요!
4. **배포** 버튼 클릭
5. "권한 검토" → 본인 구글 계정 선택 → "고급" → "(안전하지 않음) 카페수작 후기 수신(으)로 이동" → "허용"
   (Google이 본인이 만든 스크립트라 경고를 띄우는 거예요. 정상)
6. 배포 완료되면 **웹 앱 URL** 이 나옴 (`https://script.google.com/macros/s/...../exec` 형태)
7. **이 URL을 복사**

---

## STEP 4 · index.html에 URL 붙여넣기

`웹페이지/index.html` 파일에서 다음 줄을 찾으세요:

```javascript
const GOOGLE_SCRIPT_URL = "PASTE_YOUR_WEB_APP_URL_HERE";
```

따옴표 안의 `PASTE_YOUR_WEB_APP_URL_HERE` 부분을 STEP 3에서 복사한 URL로 교체하세요.

예시 (URL은 본인 것으로):
```javascript
const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfyc...../exec";
```

저장 후 `웹페이지` 폴더 통째로 Netlify Drop에 다시 업로드하세요.

→ 혜빈님이 URL만 알려주시면 제가 코드에 직접 넣어드릴게요.

---

## STEP 5 · 테스트

1. 시크릿창에서 https://meet-at-suzak.vercel.app/#voice 접속
2. 별점 + 첫인상 + 한 줄 후기 + 연락처 입력 후 제출
3. "후기 잘 받았어요" 화면 나오면 성공
4. 구글 시트 새로고침 → 한 줄이 자동으로 추가됐는지 확인

---

## ❓ 자주 발생하는 문제

**Q. 시트에 데이터가 안 들어와요**
- A. STEP 3의 "액세스 권한이 있는 사용자"가 **모든 사용자**로 되어 있는지 확인
- A. 웹 앱 URL이 `.../exec` 로 끝나는지 (`.../dev` 가 아니라)
- A. 브라우저 콘솔(F12) 에러 메시지 확인

**Q. 컬럼 순서가 어긋났어요**
- A. STEP 1의 헤더 순서 = STEP 2의 `cols` 배열 순서가 같아야 해요. 한쪽만 바꾸면 깨짐.

**Q. 코드를 수정했는데 반영이 안 돼요**
- A. Apps Script는 수정 후 **새 버전으로 재배포** 해야 반영돼요.
  → 배포 → 배포 관리 → 우측 연필(✏) → 버전: 새 버전 → 배포

**Q. 결과를 누가 볼 수 있나요?**
- A. 본인 구글 계정으로 만든 시트라 본인만 봐요. 공유 안 누르면 외부 노출 없음.
- A. 단, 누구나 후기를 "제출"은 할 수 있음 (액세스 권한이 모든 사용자라서)

---

## 📊 보고서용 데이터 분석 팁

시트에 데이터가 쌓이면 보고서 작성에 유용한 분석:

- **별점 평균**: `=AVERAGE(B2:B)` → 페이지 만족도
- **별점 분포**: `=COUNTIF(B2:B, 5)` (5점 응답 수)
- **첫인상 분포**: 피벗 테이블로 C열 그룹화 → 가장 많이 나온 인상
- **기능별 선호도**: D~I 열의 빈 셀 제외 카운트 → 어떤 기능이 가장 호평
- **방문 경험 분포**: L열 피벗

이 통계를 PPT 슬라이드(예: "VOICE OF USERS")에 그래프로 넣으면 보고서 설득력 ↑

---

준비 끝나면 알려주세요. URL 받으면 코드에 바로 적용해드릴게요.
