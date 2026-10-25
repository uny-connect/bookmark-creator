# BOOKMARK CREATORS (GAS)

일본 후쿠오카 매장 ↔ 한국 인플루언서(크리에이터) 방문·리뷰 매칭 서비스. Google Apps Script 웹앱 + Google Sheets(DB) + Gmail. **실제 운영 중** — 사용자 영향이 큰 변경은 작게, 되돌리기 쉽게.
GitHub: `uny-connect/bookmark-creator` (main). 아임웹(주문/회원) → 웹훅 → 시트 → 크리에이터/점주/운영진 화면·메일.

## 파일 구조 (GAS는 전 .js가 하나의 전역 스코프)
| 파일 | 역할 |
|---|---|
| `Router.js` | `doGet`(mode 라우팅·결과 화면), `doPost`(웹훅 위임). 로직 금지 |
| `Triggers.js` | `onOpen` 시트 메뉴, `onEdit` 자동 보정(점포ID↔점포명, 마감일·보증금, 슬롯 리셋) |
| `Service_Creator.js` | 미션 조회, 슬롯 계산(`checkAvailability`), 예약(`bookTimeSlot`), 방문확인, 리뷰 제출 |
| `Service_Store.js` | 점주: 링크 확정(`confirmBookingByLink`), 일정변경 요청, PIN 로그인, 대시보드, 매장 설정 |
| `Service_Admin.js` | 관리자 API/인증, 노쇼 배치, 빈칸 채우기, PIN 생성, 취소·재발송(시트 메뉴) |
| `Service_Imweb.js` | 아임웹 웹훅(주문/취소/가입) `handleImwebWebhook`, 회원 동기화 |
| `Service_Mail.js` | **모든 메일** 템플릿·발송(`send*Mail`), 예약 신청 알림 `_sendBookingNotifications` |
| `Service_Auth.js` | 세션 토큰·로그인 실패 제한 |
| `Config.js` | `SHEET`, `COL_*`/`IDX_*`, `MAIL`, 행 시작 상수 |
| `Utils.js` | 날짜/문자 정규화, `_getSheetsData`, Settings 조회, `_escapeHtml` |
| `View_*.html` | Admin / Creator(기본) / Store / Feedback / TestTool |

웹 라우트: `?mode=admin` · `?mode=store`(또는 `?id=`) · `?mode=store_confirm&row=&o=`(메일 링크) · `?mode=feedback&row=&o=` · `?mode=test_tool` · 기본=크리에이터 화면(`?o=&p=`).

## 시트 (스키마 = `Config.js`)
- `Master_Log`: 1~2행 헤더, 3행~ 데이터. 한 주문 항목 = 한 행. 핵심 열: A 주문번호, B 멤버코드, E 점포ID, G 구매옵션, H 방문일시, L 상태, M 피드백, Q/R 리뷰·구글 링크, U 환불, Y 시스템 로그(`[yyyy-MM-dd HH:mm]` 타임스탬프 포함)
- `Restaurant_List`: 1~2행 헤더, 3행~ (`REST_DATA_START_INDEX = 2`). 이메일 K, PIN I, 슬롯/휴무/영업시간 L~Q, 옵션별 영업시간 AA~AD, 무응답 카운트 Z
- `User_DB`: 1행 헤더. A 고유키, N 이메일, E~I SNS 채널, J 등급(YELLOW/RED/BLACK)
- `Settings`: A 키, B 값 — `ADMIN_PASSWORD`, `ADMIN_NOTIFICATION_EMAILS`, `WEB_APP_URL`, `START_DATE`/`END_DATE`/`TARGET_COUNT`(챌린지)
- `Imweb_Raw`(회원 동기화 원본), `Log`(웹훅 수신 디버그)

## 도메인 용어
- **진행 상태(L열)**: `예약대기` → (크리에이터 예약) `예약확인중` → (점주 확정) `방문전` → (매장 직원 확인) `방문완료` → (리뷰 제출) `제출완료`. 예외: `일정조율필요`(점주 일정변경 요청/어드민), `노쇼`, `취소(아임웹)`, `취소(운영자)`, `취소(기한만료)`. 취소류는 `includes('취소')` 로 판별
- **슬롯**: 매장 × 날짜 × 시간대의 예약 가능 건수. 시간대당 `MAX_PER_SLOT`, 일일 `DAILY_CAP`(숫자 또는 `기본:3,토:2` 규칙), 휴무 `BLACKOUTS`(날짜/요일/`第n요일`). 취소·노쇼는 점유하지 않음
- **24시간 규칙**: 점주 확인 메일 발송 후 24h 내 미응답이면 `취소(기한만료)` + 슬롯 반환 + 매장 무응답 카운트 +1 (대상: `예약확인중`만)
- **마감일**: 방문일 + 10일. **보증금**: 기본 10,000원, 제출완료 시 `환급대기` → 어드민이 `환급완료`
- **구매옵션(G열)**: 아임웹 옵션명 원문. 시트의 키워드 매칭으로 옵션별 영업시간 선택
- **PIN**: 점주 파트너센터 로그인 코드(6자리). **파트너/점주 = 매장 관리자**, **크리에이터 = 방문·리뷰하는 인플루언서**

## 코딩 규칙
- 전역 이름 중복 금지(파일이 달라도 충돌). 새 함수는 `node test/run-checks.js` 로 확인
- **열 번호·시트명 리터럴 금지**: `sheet.getRange(row, COL_MASTER.STATUS)`, 배열은 `row[IDX_MASTER.STATUS]`, 시트는 `SHEET.MASTER`. 컬럼이 바뀌면 `Config.js`만 수정
- 시트 읽기는 `_getSheetsData([SHEET.X, …])` 로 일괄 조회 후 메모리에서 처리. 루프 안 `getRange().getValue()` 반복 금지
- 메일은 `Service_Mail.js` 의 `send*Mail` 만 사용(`GmailApp` 직접 호출 금지). 발신/주소/URL은 `MAIL` 상수. 사용자 입력은 `_escapeHtml`
- 점주 대상 메일·화면은 **일본어**, 크리에이터·운영 화면은 한국어(Admin은 ko/ja 토글). 한글/일어 혼입 주의
- 서버 쓰기 함수는 `LockService` 로 동시성 보호(예: `bookTimeSlot`, `doPost`)
- **google.script.run 에서 호출되는 공개 함수는 누구나 호출 가능**(웹앱 `ANYONE_ANONYMOUS`):
  관리자/점주 API는 첫 줄에서 `requireToken_(token, scope)` / `requireStoreAccess_`. 클라이언트에 노출하면 안 되는 함수는 이름 끝에 `_`
- 상태 문자열은 현재 리터럴로 사용 중(`'방문전'` 등) — 오타 주의, 새 상태 추가 시 Admin/Creator HTML의 매핑도 함께 수정
- 주석·UI 문구는 기존 톤(이모지 섹션 주석, 한국어) 유지. 요청 범위 밖 리팩터링 금지

## 인증
- 관리자: `verifyAdminPassword` → 토큰(localStorage `adminToken`) → `getAdminData/adminUpdateMission/saveChallengeSettings(token…)`
- 점주: `loginStoreByPin` → `{storeId, token}` → `getStoreData/getStoreConfig(id, token)`, `updateStoreConfig(token, …)`(본인 매장 행만)
- 메일 링크(`store_confirm`, `feedback`)는 링크 소지 자체가 자격(row + 주문번호). 토큰 만료/무효 시 서버가 `AUTH_REQUIRED` 를 던지고 화면이 재로그인

## 배포 · 검증
- 검증: `node test/run-checks.js` — 구문, 전역 중복, 템플릿 파일 존재, `google.script.run` ↔ 서버 함수, 열/시트 리터럴 재유입, 메일 렌더(`mail-smoke`), 인증·확정 로직(`logic-smoke`). GAS 런타임/Gmail은 로컬 검증 불가
- 반영: `clasp push` 후 웹앱 **새 배포(버전)** 생성. `.clasp.json` 은 커밋하지 않음
- 운영 변경은 **테스트 시트 사본**에서 먼저: 예약 신청 → 점주 메일 → 확정/일정변경 → 방문확인 → 제출, 24h 만료, 어드민 수정/취소, `testCancelWebhook`
- 커밋은 단계별로 작게. 푸시는 사용자 승인 후에만
- 시간 기반 트리거(노쇼 배치 `checkAndMarkNoShow` 등) 설정은 코드에 없음 — GAS 콘솔에서 확인

## 알려진 후속 과제
- 메일 링크 GET 으로 상태가 바뀜(메일 보안 스캐너의 선조회로 자동 확정될 수 있음) → 확인 버튼 경유 권장
- View_Admin.html 이 크리에이터 입력값(후기·이름)을 `innerHTML` 에 이스케이프 없이 삽입(저장형 XSS) → 관리자 토큰 탈취 위험, 클라이언트 이스케이프 필요
- 관리자 비밀번호가 Settings 에 평문 저장
- `appsscript.json` 스코프가 넓음(`mail.google.com`, 전체 `drive`, `script.deployments`) — 변경 시 재승인 필요
- `getMissions` 는 주문번호+전화번호 뒷4자리만으로 조회(시도 제한 없음)

## 작업 원칙 (요약)
1. 가정하지 말고 모호하면 먼저 질문. 2. 요청한 만큼만, 가장 단순하게. 3. 바꾼 줄은 요청에 연결되어야 함(주변 코드 정리 금지, 내가 만든 잔여물만 제거). 4. 성공 기준을 정하고 테스트로 확인.
