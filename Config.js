/*****************************************************************************************
 * 📌 시스템 표준 시트 행 인덱스 설정
 * - Master_Log: 1~2행 헤더, 3행부터 데이터 (인덱스 i = 2, 행 번호 row = 3)
 * - Restaurant_List: 1~2행 헤더, 3행부터 데이터 (인덱스 k = 2, 행 번호 row = 3)
 * - User_DB: 1행 헤더, 2행부터 데이터 (인덱스 u = 1, 행 번호 row = 2)
 ****************************************************************************************/
const REST_DATA_START_INDEX = 2; // Restaurant_List 시트 3행부터 순회 (k = 2)
const MASTER_LOG_START_ROW = 3;  // Master_Log 실제 데이터 시작 행 번호

/** 시트 이름 */
const SHEET = {
  MASTER: 'Master_Log',
  REST: 'Restaurant_List',
  USER: 'User_DB',
  SETTINGS: 'Settings',
  IMWEB_RAW: 'Imweb_Raw',
  DEBUG_LOG: 'Log',
};

/** 열 번호 맵 (1-based, sheet.getRange(row, COL_X.Y) 용). 시트 열이 이동하면 이곳만 수정. */
// Master_Log (A=1 …). V~X(22~24)는 미사용
const COL_MASTER = {
  ORDER: 1,            // A 주문번호
  MEMBER: 2,           // B 멤버코드(아임웹 고유키)
  NAME: 3,             // C 성함(영문)
  CHANNEL: 4,          // D 참여 채널
  STORE_ID: 5,         // E 점포 ID
  STORE_NAME: 6,       // F 점포명
  OPTION: 7,           // G 구매옵션(런치/디너/코스 등)
  VISIT: 8,            // H 방문예정일시
  PEOPLE: 9,           // I 방문인원
  DEADLINE: 10,        // J 리뷰 마감 기한
  DEPOSIT: 11,         // K 보증금액
  STATUS: 12,          // L 진행 상태
  FEEDBACK: 13,        // M 점주 피드백/메모
  VISITED: 14,         // N 방문 확인(Y)
  RECEIPT: 15,         // O 매장 영수증
  SUBMIT_DATE: 16,     // P 콘텐츠 제출일
  REVIEW_URL: 17,      // Q 제출 콘텐츠 URL
  GOOGLE_URL: 18,      // R 구글맵 리뷰 URL
  SHORT_REVIEW: 19,    // S 한 줄 후기
  SHORT_REVIEW_JP: 20, // T 한 줄 후기(JP)
  REFUND: 21,          // U 환불 처리
  SYS_LOG: 25,         // Y 시스템 로그
};
// Restaurant_List
const COL_REST = {
  ID: 1,               // A 점포 ID
  NAME_KO: 2,          // B 매장명(KO)
  NAME_JP: 3,          // C 매장명(JP)
  STORE_TYPE: 8,       // H RESTAURANT / RETAIL
  PIN: 9,              // I 파트너 PIN
  EMAIL: 11,           // K 점주 수신 이메일
  MAX_PER_SLOT: 12,    // L 슬롯당 최대 예약
  DAILY_CAP: 13,       // M 일일 한도(숫자 또는 요일별 규칙)
  BLACKOUTS: 14,       // N 휴무일
  BIZ_HOURS: 15,       // O 영업시간
  INTERVAL_MIN: 16,    // P 슬롯 간격(분)
  BLOCKED_TIMES: 17,   // Q 예약 차단 시간
  MAP_URL: 18,         // R 지도 URL
  GUIDE_URL: 19,       // S 가이드 URL
  BOOKING_URL: 21,     // U 외부 예약현황 URL
  MAX_PEOPLE: 22,      // V 최대 방문 인원
  BENEFIT: 23,         // W 제공 메뉴/혜택
  MIN_DAYS: 24,        // X 최소 사전 예약일
  MAX_DAYS: 25,        // Y 최대 사전 예약일
  NORESPONSE_COUNT: 26,// Z 점주 24시간 무응답 누적
  OPT1_KEYWORDS: 27,   // AA 옵션1 매칭 키워드
  OPT1_HOURS: 28,      // AB 옵션1 운영시간
  OPT2_KEYWORDS: 29,   // AC 옵션2 매칭 키워드
  OPT2_HOURS: 30,      // AD 옵션2 운영시간
};
// User_DB
const COL_USER = {
  KEY: 1,              // A 고유키(회원코드)
  ENG_NAME: 2,         // B 영문명
  KO_NAME: 3,          // C 이름
  PHONE: 4,            // D 연락처
  BLOG: 5,             // E 네이버 블로그
  INSTA: 6,            // F 인스타그램
  YOUTUBE: 7,          // G 유튜브
  TIKTOK: 8,           // H 틱톡
  GOOGLE_GUIDE: 9,     // I 구글 로컬 가이드
  GROUP: 10,           // J 회원 그룹(등급)
  RESERVED_K: 11,      // K 미사용
  RESERVED_L: 12,      // L 미사용
  SIGNUP_DATE: 13,     // M 가입일
  EMAIL: 14,           // N 이메일
  RESERVED_O: 15,      // O 미사용
  ADMIN_MEMO: 16,      // P 관리자 메모
};
// Settings (A=키, B=값)
const COL_SETTINGS = { KEY: 1, VALUE: 2 };

/** 배열(getValues) 접근용 0-based 인덱스: row[IDX_MASTER.STATUS] */
const _toIdx = cols => Object.fromEntries(Object.entries(cols).map(([k, v]) => [k, v - 1]));
const IDX_MASTER = _toIdx(COL_MASTER);
const IDX_REST = _toIdx(COL_REST);
const IDX_USER = _toIdx(COL_USER);
const IDX_SETTINGS = _toIdx(COL_SETTINGS);

/** 한 번에 비우는 열 범위 폭 */
const MASTER_SLOT_COLS = COL_MASTER.DEADLINE - COL_MASTER.VISIT + 1;     // H~J 방문일시/인원/마감일
const MASTER_SUBMIT_COLS = COL_MASTER.REFUND - COL_MASTER.RECEIPT + 1;   // O~U 영수증~환불

/** 메일 공통 상수 */
const MAIL = {
  NAME: 'BOOKMARK CREATORS',            // 발신자 표시명
  NOTI_NAME: 'BOOKMARK NOTI',           // 운영진 알림 발신자 표시명
  FROM: 'info@bookmarkfukuoka.jp',      // 발신 주소(Gmail 별칭)
  DEFAULT_ADMIN: 'bookmarkjapan.info@gmail.com', // Settings 에 ADMIN_NOTIFICATION_EMAILS 가 없을 때
  KAKAO_URL: 'http://pf.kakao.com/_vFSxfX/chat', // 크리에이터 문의 카카오 채널
};
