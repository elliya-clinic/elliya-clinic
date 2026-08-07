// ═══════════════════════════════════════════════════════════
//  Elliya Clinic — Code.gs 변경 이력
// ═══════════════════════════════════════════════════════════
// v1  2026.07.23   예약 차단(구글시트 메뉴) 기능 신설
//                  중복된 "예약 차단 기능" 주석 블록(5개) 정리
// v2  2026.07.23   서버측 이메일 인증 신설: action=request_verification,
//                  action=verify_code 추가. doPost에 verify_token 검증 추가
//                  (기존엔 클라이언트에서만 인증번호 생성/검증 → 화면을
//                  거치지 않고 직접 요청 시 인증 우회 가능했던 보안 구멍)
// v3  2026.07.23   MailApp.sendEmail에 name:"Elliya Clinic" 추가
//                  (발신자 표시 이름 문제 — 이후 v4에서 개인계정 alias로 전환)
// v4  2026.07.23   MailApp.sendEmail에 from:"elliyaclinic@gmail.com" 추가
//                  (biggerjohny 계정에 elliyaclinic을 Gmail 별칭으로 등록 후 사용)
// v5  2026.07.24   [중요/장애대응] checkNaverBooking 검색조건에서
//                  -label:처리완료 제거, newer_than:30d로 변경.
//                  원인: 네이버 예약메일 제목이 동일해 Gmail이 서로 다른
//                  손님의 메일을 같은 스레드로 묶는 경우가 있었음. 기존
//                  라벨 기준 제외 방식은 스레드 일부가 처리완료되면 그
//                  뒤에 들어온 다른 손님의 새 예약메일까지 통째로 건너뛰어
//                  구글시트 누락 사고 발생(실사고, 7/23~7/24 사이 5건 누락
//                  확인 후 수동 복구). 이후 중복 방지는 라벨이 아닌
//                  bookingNoExists()(시트 내 실제 존재 여부)로만 판단하도록 변경
// v6  2026.07.24   [중요] 시트 참조 방식 전면 변경: getActiveSheet() 11곳
//                  전부 getSheetByName()으로 교체. 원인: getActiveSheet()는
//                  "마지막으로 열려있던 탭"을 가리키므로, 보기용 탭을 추가로
//                  만들 경우 그 탭이 열려있는 동안 자동화가 엉뚱한 시트에
//                  쓸 위험이 있었음(실제 장애로 이어지진 않았으나 사전 발견)
// v7  2026.07.24   원본 시트 탭 이름을 "시트1" → "예약현황원본"으로 확정,
//                  v6의 getSheetByName() 인자를 이 이름에 맞춰 최종 반영
// v8  2026.07.24   [버그] cancel_booking(예약취소)이 간헐적으로 "실패"로
//                  뜨는 문제 — 원인은 fetch 응답이 CORS 정책에 막혀 브라우저가
//                  못 읽은 것(서버 자체는 200 정상 처리, 시트도 실제로는
//                  갱신됨. 화면만 실패로 착각). cancel_booking에 callback
//                  파라미터를 받는 JSONP 응답 방식 추가(list_bookings와
//                  동일한 기존 검증된 패턴). index.html도 같은 버전에서
//                  스크립트태그 방식으로 함께 수정
// v9  2026.07.24   [진단 로그 삽입 시작] JSONP 전환(v8) 후에도 예약취소가
//                  계속 실패한다는 보고가 있었으나 재현/원인 확인 안 됨.
//                  cancel_booking 블록에 Logger.log 다수 추가: 받은
//                  email/booked_at 값, 매칭 실패 시 "이메일은 같은데
//                  시각이 다른" 근접 후보 목록(debug_near), try/catch로
//                  감싸 예외 발생 시에도 원인이 응답에 담기도록 함.
//                  ※ 이 로그들은 원장님이 "지워도 된다"고 명시하기 전까지
//                  임의로 제거하지 않음(요청사항)
// v10 2026.07.24   [원인 확정] 예약취소 실패의 진짜 원인 발견: Apps
//                  Script 실행 로그에서 "잠금 시간초과(178행)" 예외 확인.
//                  checkNaverBooking이 1분마다 자동 실행되며 스크립트 전체
//                  공유 잠금(LockService.getScriptLock())을 잡는데, 그
//                  타이밍과 cancel_booking 요청이 겹치면 기존 10초 대기가
//                  클라이언트 8초 타임아웃보다 길어져 계속 실패로 보였음.
//                  waitLock을 3초로 단축하고, 잠금 획득 실패 시에도 예외로
//                  죽지 않고 "서버가 다른 작업을 처리 중입니다. 잠시 후
//                  다시 시도해 주세요"라는 명확한 응답을 반환하도록 수정
// v11 2026.07.24   [근본 수정] 원장님 지적대로, 락을 "쓰기 그 순간"만
//                  잡는 게 정석인데(SQL INSERT/UPDATE와 동일한 원리) 기존
//                  checkNaverBooking은 Gmail 검색·메일 읽기·파싱까지
//                  포함한 전체 과정(최대 4분30초)을 통째로 락으로 감싸고
//                  있던 게 v10 문제의 진짜 근본 원인이었음. checkNaverBooking
//                  전체를 감싸던 큰 락 제거, 대신 handleConfirm/handleChange/
//                  markStatusByBookingNo 안에서 실제 시트 쓰기(중복확인+
//                  appendRow/setValue) 구간만 5초 짧은 락으로 감싸도록 전면
//                  재구성. registerBlock의 대기시간도 5초로 통일. 이제 락
//                  경합이 나더라도 최악 5초 안에 결론(성공 또는 "다시 시도"
//                  안내)이 나며, Gmail 처리 시간과 무관해짐
// v12 2026.07.24   [원인 확정/진짜 버그] 인증메일 발신자가 계속
//                  biggerjohny(개인계정)로 오던 문제. v3~v4에서 시도했던
//                  MailApp.sendEmail의 from 옵션이 애초에 구글 공식 문서상
//                  MailApp이 지원하지 않는 옵션이었음(공식 지원: name/cc/
//                  bcc/replyTo/noReply/attachments 뿐). 몰라도 오류 없이
//                  조용히 무시되고 항상 기본 계정으로 발송되어, 겉으로는
//                  "성공(ok:true)"으로 보였던 것. from 옵션은 GmailApp.
//                  sendEmail에서만 공식 지원(검증된 별칭 필요)되므로
//                  MailApp → GmailApp으로 교체. 참고: 지금까지 한국어
//                  테스트도 발신자를 직접 확인한 적이 없어, 사실 모든
//                  언어에서 동일하게 개인계정으로 나가고 있었을 가능성이 큼
//                  (언어별 차이가 아니라 애초에 전부 안 되고 있었던 것)
// v13 2026.07.25   [인증메일 CORS 간헐실패 수정] request_verification /
//                  verify_code 응답에 JSONP(callback) 지원 추가(_out 헬퍼).
//                  기존엔 JSON만 반환 → index.html이 fetch로 직접 호출 →
//                  CORS 리다이렉트로 간헐 실패(브라우저만 실패로 뜨고 메일은
//                  이미 발송돼 중복 발송 위험). list_bookings/cancel_booking과
//                  동일한 JSONP 패턴으로 통일. 배포 후 index.html의 인증 호출도
//                  스크립트태그(JSONP) 방식으로 교체 예정.
// v14 2026.07.29   [네이버 동기화 장애 근본 수정 + 자가 경고]
//                  장애 원인: "1분마다 30일 전체 재스캔"이 Gmail 일일 할당량을
//                  소진(Service invoked too many times) + 실행이 4~5개씩 겹쳐
//                  락 타임아웃 도배 → 예약이 시트에 안 들어오고 메일함에만 쌓임.
//                  (1) 부하 낮춤: 트리거 5분 + 평소 스캔 3일, 03시·10시 무렵
//                      하루 2회만 14일 전체 스윕(백필 안전망). 겹침 방지용 큰 락은
//                      넣지 않음(v10~v11에서 웹앱 쓰기를 막던 그 버그 재발 방지).
//                  (2) 수동 백필: naverFullSweep() = 14일 전체 강제 스캔.
//                  (3) 자가 경고: 미처리(처리완료 라벨 없음) 3건 이상이면 "발생",
//                      정상 복구 시 "해결"을 elliyaclinic/biggerjohny/serepina95
//                      3곳에 메일. 상태 판단은 숨은 플래그 없이 "경고이력" 시트의
//                      마지막 줄로만(stateless), 상태 전환 순간에만 1줄 기록 + 1통.
//                  ── 변경 함수 (2026.07.27) ──
//                  · 수정  checkNaverBooking(force30→forceSweep)
//                          · 스캔범위: 30d 고정 → 평소 3d / (03·10시 or 강제)엔 14d
//                          · 끝에 checkBacklogAndAlert() 호출 추가
//                          · 겹침용 큰 락 미도입 유지(getScriptLock 미사용)
//                  · 추가  naverFullSweep()                     — 14d 강제 스캔(수동 백필)
//                  · 추가  checkBacklogAndAlert()               — 미처리 3건↑ 발생/해결 판정·기록
//                  · 추가  sendBacklogAlert(kind,count,oldest)  — 3곳 경고메일 발송
//                  · 무변경 handleConfirm/handleCancel/handleChange/markStatusByBookingNo/
//                          bookingNoExists/extract/ensureHeader/doGet/doPost 등 그대로
//                  ※ checkNaverBooking은 트리거가 최신 저장코드를 돌리므로 이 파일은
//                    "저장"만 하면 반영됨(웹앱 재배포 불필요 → SHEET_URL·index 무관).
// v15 2026.07.29   [지난 예약 아카이브 + 스윕 시각 조정]
//                  데이터 누적 대비: 시술일(F열)이 오늘보다 과거인 예약을
//                  "지난예약_YYYY" 탭으로 옮겨 예약현황원본을 항상 가볍게 유지.
//                  오늘 포함 미래 예약은 원본에 남김("지난 예약은 취소 와도 무시").
//                  · 추가  moveOldBookings()  — 매일 03시 트리거로 실행 예정.
//                          예약번호 기준 중복방지(재실행 안전) / 아래→위 삭제(행밀림
//                          방지) / 쓰기 구간만 락 / 시술일 파싱 불가 행은 남김 /
//                          연도 탭(지난예약_YYYY) 없으면 자동 생성(원본과 동일 헤더).
//                  · 수정  checkNaverBooking 전체 스윕 시각 03시→04시(10시는 유지).
//                          03시 moveOldBookings가 시트를 정리한 뒤 04시 스윕이
//                          가벼운 시트를 읽도록 순서 배치. (hr===4 || hr===10)
//                  ※ 트리거: 기존 checkNaverBooking(5분) 1개 + moveOldBookings(매일
//                    03시) 1개 = 총 2개. moveOldBookings 트리거는 원장님이 추가.
//                    첫 실행(누적분 많을 때)은 시간이 좀 걸릴 수 있음(1회성). 저장만 하면 반영.
// v16 2026.07.29   [취소/변경 중복 재처리 방지] 라벨과 무관하게 3일치를 매번
//                  재검색하는 구조(스레드 병합 누락 방지)에서, 확정은 bookingNoExists
//                  로 중복을 막았지만 취소/변경(markStatusByBookingNo)엔 중복방지가
//                  없어, 이미 취소/변경된 행을 매 5분 실행마다 다시 덮어써 처리시각
//                  (R열)이 계속 갱신되고 불필요한 시트 쓰기가 반복되던 문제.
//                  · 수정  markStatusByBookingNo: 찾은 행의 현재 상태(C열)가 바꾸려는
//                          상태와 같으면 즉시 return(재기록 안 함). 취소·변경 공통 적용.
//                          → 처리시각은 최초 상태변경 1회로 고정, 확정·취소·변경이
//                          모두 중복방지 대칭이 됨. (3일 재검색 자체는 그대로 유지)
// v17 2026.07.29   [홈페이지 차단을 폰 대시보드에서 — doGet 액션 노출]
//                  기존 홈페이지 차단은 PC 커스텀 메뉴(registerBlock/deleteBlockRows)로만
//                  가능해 폰에선 못 썼음. 대시보드(PWA)가 부를 수 있게 doGet 액션 3개 추가.
//                  차단은 예약경로="차단" 행으로 저장되고 홈페이지 달력(get_booked_times)이
//                  이를 "찬 시간"으로 읽어 막는 기존 메커니즘을 그대로 재사용.
//                  · 추가  doGet "day_slots"   : 날짜의 요일별 슬롯 + 상태(open/booked/blocked)
//                  · 추가  doGet "add_block"   : 슬롯 1개 차단 등록(registerBlock 재사용)
//                  · 추가  doGet "remove_block": 슬롯 1개 차단 해제(deleteBlockRows 재사용)
//                  · 3개 모두 JSONP(callback) + key="elliya2026" 검사(mark_naver_done과 동일).
//                  · SLOT_TABLE(요일별 슬롯)은 index.html 슬롯 정의와 동일하게 유지해야 함.
//                  ※ 이번엔 doGet을 바꿨으므로 "새 배포" 필요 → /exec URL 변경됨 →
//                    index.html·dashboard.html SHEET_URL 교체 필요(v13 때와 동일 절차).
// v19 2026.07.29   [예약시간 조회 속도 개선 — 컬럼 축소] 날짜 클릭 후 예약가능
//                  시간이 뜨기까지 느린 문제. get_booked_times가 시트 전체(18개
//                  컬럼)를 읽던 것을 → 실제 사용하는 A~F(6개)만 읽도록 변경.
//                  이름·전화·이메일 등 불필요한 데이터 전송이 빠져 응답이 가벼워짐.
//                  · 수정  get_booked_times: getDataRange().getValues() →
//                          getRange(1,1,lastRow,6).getValues() (로직·결과 동일).
//                  ※ get_booked_times는 doGet(웹앱)이므로 "새 배포" 필요 → /exec URL
//                    변경 → index.html·dashboard.html SHEET_URL 교체 필요.
// v20 2026.07.30   [★버그수정: 지난 예약이 매일 새벽 원본에 다시 채워지던 문제]
//                  증상: 03시 moveOldBookings가 지난 예약을 지난예약_YYYY로 옮긴 뒤,
//                  04시 checkNaverBooking 14일 전체 스윕이 같은(지난) 네이버 메일을
//                  다시 읽어 예약현황원본에 재삽입(처리시각 03:55~04:01로 무더기 유입).
//                  원인: bookingNoExists가 원본만 보므로, 옮겨진 건을 "없음"으로 오판.
//                  moveOldBookings(이동)와 스윕(재읽기)이 충돌한 설계 결함.
//                  · 추가  isPastServiceDate(): 시술일(F)이 오늘보다 과거인지 판정.
//                  · 수정  handleConfirm / handleChange(신규): append 직전에
//                          isPastServiceDate면 원본에 넣지 않고 return.
//                          → 스윕이 지난 예약을 다시 읽어도 원본에 재유입 안 됨.
//                          신규 확정/변경은 시술일이 항상 미래라 영향 없음.
//                  ※ handleConfirm/handleChange는 트리거(checkNaverBooking) 경로이므로
//                    "저장"만 하면 반영됨(새 배포·URL 변경 불필요).
// v21 2026.07.31   [지난예약 아카이브에 '이동시각' 열 추가] moveOldBookings가
//                  지난 예약을 지난예약_YYYY로 옮길 때, 원본 행(A~R) 뒤에
//                  '이동시각'(맨 뒤 열, yyyy.MM.dd. HH:mm:ss)을 덧붙여 저장.
//                  언제 아카이브됐는지 추적용. 헤더도 맨 뒤 열에 '이동시각' 보장.
//                  기존에 이미 옮겨진 행은 이 열이 비어 있음(이번 변경 이후분만 기록).
//                  ※ moveOldBookings는 트리거 경로 → "저장"만 하면 반영(새 배포 불필요).
// v22 2026.08.07   [예약번호 자동생성 — 데이터 소실 버그 수정] 홈페이지·차단 예약에
//                  D열(예약번호)이 비어 있어 moveOldBookings의 중복제거 필터
//                  ( !existing[String(v[3])] )에서 빈 문자열 ""이 하나의 키로
//                  뭉개지는 문제가 있었음. 지난예약_YYYY에 예약번호 없는 행이
//                  단 1건이라도 생기면 existing[""]=true가 되어, 이후 모든
//                  홈페이지·차단 예약이 필터에서 제외 → 아카이브에 기록되지
//                  않은 채 원본에서는 deleteRow로 삭제 → 데이터 완전 소실.
//                  (2026.08.05 홈페이지 예약 1건이 실제로 이렇게 유실됨)
//                  · 추가  genBookingNo(prefix, dateObj) — 접수시각 기반 고유번호.
//                          형식: H20260807153042 (홈페이지) / K20260807153042 (차단)
//                          네이버 예약번호(10자리 순수 숫자)와 절대 충돌하지 않도록
//                          알파벳 접두어로 네임스페이스 분리. 같은 초 동시 삽입까지
//                          대비해 뒤에 2자리 랜덤 부여.
//                  · doPost 홈페이지 예약 저장 시 D열에 H번호 기록.
//                  · register_block_web 차단 행 저장 시 D열에 K번호 기록.
//                    (같은 요청의 여러 시간대 행은 각각 다른 번호를 가짐)
//                  ※ 기존에 이미 저장된 D열 공란 행은 소급 적용되지 않음.
//                    → moveOldBookings 필터도 함께 보강: 예약번호가 비어 있는
//                      행은 중복검사 대상에서 제외(항상 통과)시켜 과거 데이터도
//                      더 이상 유실되지 않게 함.
//                  ※ doPost(웹앱 진입점) 변경이므로 반드시 배포 필요.
//                    단, [배포 관리 > 기존 배포 연필(✏️) > 버전:새 버전 > 배포]로
//                    하면 /exec URL이 유지되어 index.html·dashboard.html의
//                    SHEET_URL 교체와 Cloudflare 재배포가 모두 불필요함.
// ═══════════════════════════════════════════════════════════

// ══════════════════════════════════════════════
// Elliya Clinic 예약 시스템 - Apps Script (v3)
// 컬럼 순서: 예약일시/예약경로/상태/예약번호/이름/시술일시/시술명/이메일/전화/국적/특이사항/예상비용
// (M~Q열은 추후 사용을 위해 비워둠, R열 = 처리시각)
// ══════════════════════════════════════════════

// ══════════════════════════════════════════════
// 예약번호 자동생성 (v22)
// ──────────────────────────────────────────────
// 네이버 예약번호는 10자리 순수 숫자(예: 1313796338).
// 홈페이지/차단 건은 알파벳 접두어를 붙여 네임스페이스를 완전히 분리한다.
//   H + yyyyMMddHHmmss + 2자리랜덤  → 홈페이지 예약
//   K + yyyyMMddHHmmss + 2자리랜덤  → 차단(Kill/Block)
// 접수시각 기준이라 정렬하면 시간순이 되고, 사람이 봐도 경로 구분이 된다.
// 2자리 랜덤은 같은 초에 2건이 동시 삽입되는 경우까지 대비한 것.
// ══════════════════════════════════════════════
function genBookingNo(prefix, dateObj) {
  var d = dateObj || new Date();
  var ts = Utilities.formatDate(d, 'Asia/Seoul', 'yyyyMMddHHmmss');
  var rnd = ('0' + Math.floor(Math.random() * 100)).slice(-2);
  return prefix + ts + rnd;
}


function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("예약현황원본");
    var data = JSON.parse(e.postData.contents);
    ensureHeader(sheet);

    // ── 이메일 인증 토큰 검증 (없거나 틀리면 예약 자체를 거부) ──
    var cache = CacheService.getScriptCache();
    var expectedToken = cache.get("vtoken_" + data.patient_email);
    if (!expectedToken || !data.verify_token || expectedToken !== data.verify_token) {
      return ContentService
        .createTextOutput(JSON.stringify({result: 'error', message: '이메일 인증이 필요합니다.'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    cache.remove("vtoken_" + data.patient_email); // 1회용 소진(재사용 방지)

    var now = new Date();
    var bookedAt = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy.MM.dd. HH:mm:ss');
    var processedAt = bookedAt; // 홈페이지 예약은 요청 즉시 저장되므로 예약일시와 동일

    sheet.appendRow([
      bookedAt,                    // A: 예약일시
      '홈페이지',                   // B: 예약경로
      '확정',                      // C: 상태
      genBookingNo('H', now),      // D: 예약번호 (v22 — 자동생성. 공란이면 아카이브 시 유실됨)
      data.patient_name,           // E: 이름
      data.date_time,              // F: 시술일시
      data.service,                // G: 시술명
      data.patient_email,          // H: 이메일
      "'" + data.phone,            // I: 전화
      data.nationality,            // J: 국적
      data.notes,                  // K: 특이사항
      data.price,                  // L: 예상비용
      '', '', '', '', '',          // M~Q: (추후 사용을 위해 비움)
      processedAt                  // R: 처리시각
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({result: 'success'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({result: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var action = e.parameter.action;

  // JSONP 지원: callback 파라미터가 있으면 callback(json)으로 감싸 반환(CORS 우회),
  // 없으면 순수 JSON. request_verification/verify_code의 CORS 간헐 실패 방지용(v13).
  function _out(obj){
    var cb = e.parameter.callback || "";
    var json = JSON.stringify(obj);
    if (cb) return ContentService.createTextOutput(cb + "(" + json + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "request_verification") {
    var vEmail = e.parameter.email || "";
    var vLang = e.parameter.lang || "ko";
    var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailPattern.test(vEmail)) {
      return _out({ok:false, error:"invalid_email"});
    }
    var code = String(Math.floor(100000 + Math.random()*900000));
    CacheService.getScriptCache().put("vcode_" + vEmail, code, 180); // 3분간 유효

    var VS = {
      ko:{s:"[Elliya Clinic] 예약 인증번호", b:"안녕하세요, 엘리야의원입니다.\n\n예약 인증번호는 " + code + " 입니다.\n3분 이내에 입력해 주세요."},
      en:{s:"[Elliya Clinic] Your Verification Code", b:"Hello, this is Elliya Clinic.\n\nYour verification code is " + code + ".\nPlease enter it within 3 minutes."},
      zh:{s:"[Elliya Clinic] 您的驗證碼", b:"您好，這裡是 Elliya Clinic。\n\n您的驗證碼是 " + code + "。\n請於 3 分鐘內輸入。"},
      cn:{s:"[Elliya Clinic] 您的验证码", b:"您好，这里是 Elliya Clinic。\n\n您的验证码是 " + code + "。\n请于 3 分钟内输入。"},
      ja:{s:"[Elliya Clinic] 認証番号のご案内", b:"こんにちは、Elliya Clinicです。\n\n認証番号は " + code + " です。\n3分以内にご入力ください。"},
      ar:{s:"[Elliya Clinic] رمز التحقق الخاص بك", b:"مرحباً، هذه عيادة Elliya Clinic.\n\nرمز التحقق الخاص بك هو " + code + ".\nيرجى إدخاله خلال 3 دقائق."}
    };
    var pick = VS[vLang] || VS.en;

    try {
      GmailApp.sendEmail(vEmail, pick.s, pick.b, { name: "Elliya Clinic", from: "elliyaclinic@gmail.com" });
      return _out({ok:true});
    } catch (mailErr) {
      return _out({ok:false, error: mailErr.message});
    }
  }

  if (action === "verify_code") {
    var vcEmail = e.parameter.email || "";
    var vcCode = e.parameter.code || "";
    var cacheV = CacheService.getScriptCache();
    var stored = cacheV.get("vcode_" + vcEmail);
    if (stored && vcCode && stored === vcCode) {
      var token = Utilities.getUuid();
      cacheV.put("vtoken_" + vcEmail, token, 600); // 10분간 유효(그 사이 예약접수 완료해야 함)
      cacheV.remove("vcode_" + vcEmail);
      return _out({ok:true, token: token});
    }
    return _out({ok:false});
  }

  if (action === "search") {
    var name  = e.parameter.name  || "";
    var email = e.parameter.email || "";
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("예약현황원본");
    var data  = sheet.getDataRange().getValues();
    var bookings = [];
    for (var i = 1; i < data.length; i++) {
      var rowName   = String(data[i][4] || "").trim();
      var rowEmail  = String(data[i][7] || "").trim();
      var rowStatus = String(data[i][2] || "").trim();
      if (rowName.includes(name) && rowEmail === email && rowStatus !== '취소') {
        bookings.push({
          booked_at: data[i][0],
          service:   data[i][6],
          date_time: data[i][5],
          price:     data[i][11],
          status:    rowStatus
        });
      }
    }
    return ContentService
      .createTextOutput(JSON.stringify({bookings: bookings}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "cancel_booking") {
    var cEmail = e.parameter.email || "";
    var cBookedAt = e.parameter.booked_at || "";
    var cCallback = e.parameter.callback || ""; // JSONP 지원(스크립트태그 로딩 방식) — CORS 우회
    var sheetC = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("예약현황원본");
    Logger.log("[취소진단] 받은 email=[" + cEmail + "] booked_at=[" + cBookedAt + "]");
    var lockC = LockService.getScriptLock();
    try {
      lockC.waitLock(3000); // 3초로 단축: checkNaverBooking(1분마다 자동실행)과
                             // 겹치면 기존 10초 대기가 클라이언트 8초 타임아웃보다
                             // 길어져 "응답 없음"으로 보이는 문제가 있었음. 짧게
                             // 기다리다 안 되면 아래 catch에서 "잠시 후 재시도"로
                             // 명확히 안내하도록 변경(무한정 기다리지 않음)
    } catch (lockErrC) {
      Logger.log("[취소진단] 잠금 획득 실패: " + lockErrC.message);
      var lockFailJsonC = JSON.stringify({ok: false, error: "서버가 다른 작업을 처리 중입니다. 잠시 후 다시 시도해 주세요."});
      if (cCallback) {
        return ContentService.createTextOutput(cCallback + "(" + lockFailJsonC + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService.createTextOutput(lockFailJsonC).setMimeType(ContentService.MimeType.JSON);
    }
    try {
      var dataC = sheetC.getDataRange().getValues();
      var foundC = false;
      var nearMatches = []; // 이메일은 같은데 시각이 다른 행들(진단용)
      for (var ci = 1; ci < dataC.length; ci++) {
        var rowEmailC = String(dataC[ci][7] || "").trim();      // H열: 이메일
        var rowBookedAtC = String(dataC[ci][0] || "").trim();   // A열: 예약일시
        var rowStatusC = String(dataC[ci][2] || "").trim();     // C열: 상태
        if (rowEmailC === cEmail && rowBookedAtC !== cBookedAt) {
          nearMatches.push(rowBookedAtC); // 이메일 일치, 시각 불일치 기록(진단용)
        }
        if (rowEmailC === cEmail && rowBookedAtC === cBookedAt && rowStatusC !== "취소") {
          var naverDoneAtC = String(dataC[ci][16] || "").trim(); // Q열: 네이버처리시각(취소 전 값)
          var nowC = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy.MM.dd. HH:mm:ss");

          sheetC.getRange(ci + 1, 3).setValue("취소"); // C열 = 상태 변경

          if (naverDoneAtC) {
            // 케이스1: 직원이 이미 네이버에서 차단해둔 상태였음
            // → Q열을 다시 비워서 "차단 해제 필요"를 대시보드가 알아채게 함
            sheetC.getRange(ci + 1, 17).setValue("");
          } else {
            // 케이스2: 직원이 아직 처리 안 한 상태였음
            // → 어차피 할 일이 없으므로 Q열에 지금 시각을 채워 "처리 불필요"로 표시
            sheetC.getRange(ci + 1, 17).setValue(nowC);
          }

          sheetC.getRange(ci + 1, 18).setValue(nowC); // R열 = 처리시각(취소 처리 시각으로 갱신)
          foundC = true;
          break;
        }
      }
      if (!foundC) {
        Logger.log("[취소진단] 매칭 실패. email은 같은데 시각이 다른 행들(참고): " + JSON.stringify(nearMatches));
      }
      var resultJsonC = JSON.stringify({ok: foundC, debug_near: nearMatches});
      if (cCallback) {
        return ContentService.createTextOutput(cCallback + "(" + resultJsonC + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService.createTextOutput(resultJsonC).setMimeType(ContentService.MimeType.JSON);
    } catch (errC) {
      Logger.log("[취소진단] 예외 발생: " + errC.message);
      var errJsonC = JSON.stringify({ok: false, error: errC.message});
      if (cCallback) {
        return ContentService.createTextOutput(cCallback + "(" + errJsonC + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService.createTextOutput(errJsonC).setMimeType(ContentService.MimeType.JSON);
    } finally {
      lockC.releaseLock();
    }
  }

  if (action === "get_booked_times") {
  var date = e.parameter.date || "";
  var callback = e.parameter.callback || "";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("예약현황원본");
  // 성능 최적화(v19): 이 함수는 C열(상태)·F열(시술일시)만 쓰므로 A~F(6개 컬럼)만 읽는다.
  // 기존 getDataRange()는 18개 컬럼 전부를 읽어 이름·전화·이메일 등 불필요한 데이터까지
  // 전송했음. 범위만 줄인 것이라 로직·결과는 동일. (data[i][2]=C, data[i][5]=F 그대로 유효)
  var lastRow = sheet.getLastRow();
  var data = (lastRow >= 1) ? sheet.getRange(1, 1, lastRow, 6).getValues() : [];
  var busyMinutes = []; // 그 날짜에 확정되어 있는 예약들의 '자정 기준 분(minute)' 목록

  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][2] || "").trim();
    var dateTime = String(data[i][5] || "").trim();

    // 취소된 예약은 제외 (취소, 취소(변경) 모두 포함)
    if (status.indexOf('취소') !== -1) continue;

    // 날짜 포함 여부 확인
    if (dateTime.indexOf(date) === -1) continue;

    // 시간 부분 추출 후 '분' 단위 숫자로 변환 (예: 오후 2:00 → 840)
    var timeMatch = dateTime.match(/(오전|오후)\s+(\d+):(\d+)/);
    if (!timeMatch) continue;
    var h = parseInt(timeMatch[2], 10);
    var m = parseInt(timeMatch[3], 10);
    if (timeMatch[1] === '오후' && h !== 12) h += 12;
    if (timeMatch[1] === '오전' && h === 12) h = 0;
    busyMinutes.push(h * 60 + m);
  }

  var resultJson = JSON.stringify({busyMinutes: busyMinutes});

  // callback 파라미터가 있으면 JSONP 형태로 감싸서 반환 (CORS 우회용)
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + resultJson + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(resultJson)
    .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "list_bookings") {
    var key = e.parameter.key || "";
    var callback3 = e.parameter.callback || "";
    var STAFF_KEY = "elliya2026"; // 직원용 접근 키 (필요시 변경하세요)

    if (key !== STAFF_KEY) {
      var errJson = JSON.stringify({error: "unauthorized"});
      if (callback3) {
        return ContentService.createTextOutput(callback3 + "(" + errJson + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService.createTextOutput(errJson).setMimeType(ContentService.MimeType.JSON);
    }

    var sheetL = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("예약현황원본");
    var dataL = sheetL.getDataRange().getValues();
    var todayStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy.MM.dd');
    var bookingsL = [];

    for (var j = 1; j < dataL.length; j++) {
      var statusL = String(dataL[j][2] || "").trim();
      var isCancelledL = statusL.indexOf('취소') !== -1;
      var naverDoneAtL = String(dataL[j][16] || "").trim(); // Q열: 네이버처리시각

      // 취소된 건인데 Q열이 이미 채워져 있으면(=처리 불필요로 확정된 케이스) 목록에서 제외
      if (isCancelledL && naverDoneAtL) continue;

      var sourceL = String(dataL[j][1] || "").trim();
      if (sourceL.indexOf('홈페이지') === -1) continue; // 홈페이지 예약만 (네이버는 이미 원장님이 확인 가능하므로 제외)

      var dateTimeL = String(dataL[j][5] || "").trim();
      var dateMatchL = dateTimeL.match(/\d{4}\.\d{2}\.\d{2}/);
      var dateOnlyL = dateMatchL ? dateMatchL[0] : "";
      if (dateOnlyL && dateOnlyL < todayStr) continue; // 지난 날짜 제외 (오늘 이후만)

      bookingsL.push({
        source: dataL[j][1],
        status: statusL,
        is_cancelled: isCancelledL,       // true면 "차단 해제 필요" 케이스
        booking_no: dataL[j][3],
        name: dataL[j][4],
        date_time: dateTimeL,
        date_only: dateOnlyL,
        service: dataL[j][6],
        email: dataL[j][7],
        phone: dataL[j][8],
        nationality: dataL[j][9],
        naver_done_at: dataL[j][16] || "",   // Q열: 네이버처리시각
        processed_at: dataL[j][17] || ""     // R열: 처리시각 (행을 특정하는 고유 식별자로 사용)
      });
    }

    bookingsL.sort(function(a, b) { return a.date_only.localeCompare(b.date_only); });

    var resultJsonL = JSON.stringify({bookings: bookingsL});
    if (callback3) {
      return ContentService.createTextOutput(callback3 + "(" + resultJsonL + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(resultJsonL).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "mark_naver_done") {
    var keyM = e.parameter.key || "";
    var callback4 = e.parameter.callback || "";
    var STAFF_KEY_M = "elliya2026"; // list_bookings와 동일한 접근 키

    if (keyM !== STAFF_KEY_M) {
      var errJsonM = JSON.stringify({error: "unauthorized"});
      if (callback4) {
        return ContentService.createTextOutput(callback4 + "(" + errJsonM + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService.createTextOutput(errJsonM).setMimeType(ContentService.MimeType.JSON);
    }

    var processedAtKey = e.parameter.processed_at || ""; // R열 값으로 행을 특정
    var markValue = e.parameter.value || ""; // "1"이면 체크, ""이면 해제

    var sheetM = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("예약현황원본");
    var dataM = sheetM.getDataRange().getValues();
    var found = false;

    for (var k = 1; k < dataM.length; k++) {
      var rVal = String(dataM[k][17] || "").trim(); // R열 (처리시각)
      if (rVal === processedAtKey.trim() && processedAtKey) {
        if (markValue === "1") {
          var naverDoneAt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy.MM.dd. HH:mm:ss');
          sheetM.getRange(k + 1, 17).setValue(naverDoneAt); // Q열 (17번째 컬럼)
        } else {
          sheetM.getRange(k + 1, 17).setValue(''); // 체크 해제 시 Q열 비움
        }
        found = true;
        break;
      }
    }

    var resultJsonM = JSON.stringify({ok: found});
    if (callback4) {
      return ContentService.createTextOutput(callback4 + "(" + resultJsonM + ")").setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(resultJsonM).setMimeType(ContentService.MimeType.JSON);
  }
  // ── 홈페이지 차단 (대시보드용, v17) ──
  // day_slots: 해당 날짜의 예약 슬롯 목록 + 각 슬롯 상태(open/booked/blocked)
  if (action === "day_slots") {
    if ((e.parameter.key || "") !== "elliya2026") return _out({ ok:false, error:"unauthorized" });
    var dsDate = e.parameter.date || ""; // YYYY-MM-DD
    var dp = dsDate.split('-');
    if (dp.length !== 3) return _out({ ok:false, error:"bad_date" });
    var dObj = new Date(parseInt(dp[0],10), parseInt(dp[1],10)-1, parseInt(dp[2],10));
    var wd = dObj.getDay();
    // 요일별 슬롯 테이블(index.html의 슬롯 정의와 반드시 동일하게 유지할 것)
    var SLOT_TABLE = {
      1: ["10:30","11:30","12:30","13:30","14:30","15:30","16:30","17:30","18:30","19:30"],
      5: ["10:30","11:30","12:30","13:30","14:30","15:30","16:30","17:30","18:30","19:30"],
      2: ["10:30","11:30","12:30","13:30","14:30","15:30","16:30","17:30"],
      4: ["10:30","11:30","12:30","13:30","14:30","15:30","16:30","17:30"],
      6: ["9:30","10:30","11:30","13:00","14:00","15:00"]
    };
    var slotList = SLOT_TABLE[wd] || [];
    var dateDot = dp[0] + "." + dp[1] + "." + dp[2]; // 2026.08.01
    var dsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("예약현황원본");
    var dsData = dsSheet.getDataRange().getValues();
    var busyMap = {}; // minute -> 'blocked' | 'booked'
    for (var di = 1; di < dsData.length; di++) {
      var dsStatus = String(dsData[di][2] || "").trim();
      if (dsStatus.indexOf('취소') !== -1) continue;
      var dsDt = String(dsData[di][5] || "").trim();
      if (dsDt.indexOf(dateDot) === -1) continue;
      var dsTm = dsDt.match(/(오전|오후)\s+(\d+):(\d+)/);
      if (!dsTm) continue;
      var dsh = parseInt(dsTm[2],10), dsm = parseInt(dsTm[3],10);
      if (dsTm[1] === '오후' && dsh !== 12) dsh += 12;
      if (dsTm[1] === '오전' && dsh === 12) dsh = 0;
      var dsMin = dsh*60 + dsm;
      var dsKind = String(dsData[di][1] || "").trim() === '차단' ? 'blocked' : 'booked';
      if (busyMap[dsMin] !== 'booked') busyMap[dsMin] = dsKind; // 실예약이 우선(보호)
    }
    var slotsOut = slotList.map(function(t){
      var tp = t.split(':'); var tmin = parseInt(tp[0],10)*60 + parseInt(tp[1],10);
      return { time: t, status: busyMap[tmin] || 'open' };
    });
    return _out({ ok:true, date: dsDate, weekday: wd, closed: (slotList.length === 0), slots: slotsOut });
  }

  // add_block: 슬롯 하나를 차단으로 등록
  if (action === "add_block") {
    if ((e.parameter.key || "") !== "elliya2026") return _out({ ok:false, error:"unauthorized" });
    var abDate = e.parameter.date || "";  // YYYY-MM-DD
    var abTime = e.parameter.time || "";  // HH:MM (슬롯 시각)
    var abAdmin = e.parameter.admin || "관리자";
    var atp = abTime.split(':');
    if (atp.length !== 2) return _out({ ok:false, error:"bad_time" });
    var abEndMin = parseInt(atp[0],10)*60 + parseInt(atp[1],10) + 60;
    var abEndStr = pad2(Math.floor(abEndMin/60)) + ":" + pad2(abEndMin%60);
    return _out(registerBlock(abDate, abTime, abEndStr, abAdmin, "차단"));
  }

  // remove_block: 슬롯 하나의 차단 해제
  if (action === "remove_block") {
    if ((e.parameter.key || "") !== "elliya2026") return _out({ ok:false, error:"unauthorized" });
    var rbDate = e.parameter.date || "";  // YYYY-MM-DD
    var rbTime = e.parameter.time || "";  // HH:MM
    var rdp = rbDate.split('-');
    var rtp = rbTime.split(':');
    if (rdp.length !== 3 || rtp.length !== 2) return _out({ ok:false, error:"bad_param" });
    var rbDateDot = rdp[0] + "." + rdp[1] + "." + rdp[2];
    var rbMin = parseInt(rtp[0],10)*60 + parseInt(rtp[1],10);
    var rbSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("예약현황원본");
    var rbData = rbSheet.getDataRange().getValues();
    var rbRows = [];
    for (var ri = 1; ri < rbData.length; ri++) {
      if (String(rbData[ri][1] || "").trim() !== '차단') continue;
      var rbDt = String(rbData[ri][5] || "").trim();
      if (rbDt.indexOf(rbDateDot) === -1) continue;
      var rbTm = rbDt.match(/(오전|오후)\s+(\d+):(\d+)/);
      if (!rbTm) continue;
      var rbh = parseInt(rbTm[2],10), rbm = parseInt(rbTm[3],10);
      if (rbTm[1] === '오후' && rbh !== 12) rbh += 12;
      if (rbTm[1] === '오전' && rbh === 12) rbh = 0;
      if (rbh*60 + rbm === rbMin) rbRows.push(ri + 1); // 실제 시트 행번호
    }
    if (rbRows.length === 0) return _out({ ok:false, error:"not_found" });
    return _out({ ok:true, removed: deleteBlockRows(rbRows).count });
  }

  return ContentService
    .createTextOutput(JSON.stringify({result: 'no action'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      '예약일시', '예약경로', '상태', '예약번호',
      '이름', '시술일시', '시술명', '이메일',
      '전화', '국적', '특이사항', '예상비용',
      '', '', '', '',               // M~P: (추후 사용을 위해 비움)
      '네이버처리시각',              // Q
      '처리시각'                    // R
    ]);
  }
}

// ══════════════════════════════════════════════
// 네이버 예약 이메일 자동 처리
// ══════════════════════════════════════════════

function checkNaverBooking(forceSweep) {
  // ※ v11부터: 이 함수 전체를 감싸던 큰 락을 제거함(원장님 지적: 락은 "쓰기 그 순간"만).
  //   Gmail 검색·읽기·파싱 구간까지 락을 쥐면 그 사이 예약취소 등 웹앱 쓰기가 락을
  //   기다리다 타임아웃됨(v10~v11의 그 버그). 그래서 여기엔 큰 락을 두지 않는다.
  //   → v14: "겹침 방지"를 큰 락으로 넣으면 위 버그가 되살아나므로 넣지 않고, 대신
  //     트리거를 5분·스캔을 3일로 가볍게 만들어 실행이 겹치지 않도록 부하를 낮춘다.
  //     실제 시트 쓰기는 handleConfirm/markStatusByBookingNo 안 5초 짧은 락 + 중복방지로 안전.
  var startTime = Date.now();
  var MAX_RUNTIME_MS = 4.5 * 60 * 1000; // 4분30초 넘으면 안전하게 중단(실제 한도 6분보다 여유)

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("예약현황원본");
    ensureHeader(sheet);

    var doneLabel = GmailApp.getUserLabelByName("처리완료");
    if (!doneLabel) doneLabel = GmailApp.createLabel("처리완료");

    // ── 스캔 범위 (v14) ──
    // 평소엔 최근 3일만 가볍게(신규 예약은 5분 내 반영). 매일 04시·10시 무렵(하루 2회)
    // 또는 수동 naverFullSweep() 실행 시엔 14일 전체 스윕으로, 스레드 병합 등으로 놓친
    // 건을 백필(안전망). 예전 "1분마다 30일 전체"가 Gmail 일일 할당량을 소진시켜 동기화가
    // 멈추던 문제(Service invoked too many times)를 근본적으로 줄이기 위함.
    // ※ v15: 스윕 03→04시로 변경(03시엔 moveOldBookings가 지난예약을 먼저 정리해
    //   시트를 가볍게 만든 뒤, 04시 스윕이 가벼운 시트를 읽도록 순서 배치).
    // 라벨(-label:처리완료)로 스레드를 건너뛰지 않음: 네이버 메일은 제목이 같아 서로 다른
    // 손님이 한 스레드로 병합될 수 있어, 실제 중복 방지는 bookingNoExists()로만 한다.
    var mins = new Date().getMinutes();
    var hr = new Date().getHours();
    var fullSweep = (forceSweep === true) || ((hr === 4 || hr === 10) && mins < 5);
    var win = fullSweep ? '14d' : '3d';

    var threads = GmailApp.search(
      'from:naverbooking_noreply@navercorp.com -in:trash -in:spam label:"엘리야 예약알림" newer_than:' + win
    );

    for (var t = 0; t < threads.length; t++) {
      if (Date.now() - startTime > MAX_RUNTIME_MS) {
        Logger.log('실행시간 임박으로 중단. 남은 ' + (threads.length - t) + '개 스레드는 다음 실행에서 처리됨.');
        break;
      }
      var thread = threads[t];
      try {
        thread.getMessages().forEach(function(message) {
          var subject = message.getSubject();
          var body = message.getPlainBody();
          if (subject.indexOf('취소') !== -1 && subject.indexOf('변경') === -1) {
            handleCancel(sheet, body);
          } else if (subject.indexOf('변경') !== -1) {
            handleChange(sheet, body);
          } else if (subject.indexOf('확정') !== -1) {
            handleConfirm(sheet, body);
          }
        });
        thread.addLabel(doneLabel);
      } catch (threadErr) {
        Logger.log('스레드 처리 중 에러: ' + threadErr);
      }
    }

    // ── 밀림 감지 & 경고(발생/해결) ── (v14)
    checkBacklogAndAlert();

  } catch (outerErr) {
    Logger.log('checkNaverBooking 처리 중 오류: ' + outerErr);
  }
}

// 수동 전체 백필: 편집기에서 이 함수를 직접 실행하면 14일 전체를 훑어 밀린 예약을 복구한다.
function naverFullSweep() {
  checkNaverBooking(true);
}

// ── 밀림 감지 & 경고 (v14) ──
// "미처리"(= "처리완료" 라벨이 안 붙은 네이버 메일)가 3건 이상이면 "발생", 정상 복구되면 "해결".
// 숨은 플래그를 두지 않고 "경고이력" 시트의 마지막 줄(=현재 상태)로만 판단(stateless).
// 상태가 바뀌는 순간에만 한 줄 기록 + 메일 1통(도배 방지). 3곳 이메일로 발송.
function checkBacklogAndAlert() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var log = ss.getSheetByName("경고이력");
    if (!log) { Logger.log('경고이력 시트 없음 — 경고 스킵'); return; }

    // 완전 빈 시트면 헤더 한 줄만 넣어줌(원장님이 이미 헤더를 넣었으면 건너뜀)
    if (log.getLastRow() === 0) {
      log.appendRow(['시각', '구분', '미처리건수', '가장오래된미처리시각', '사유']);
    }

    var THRESHOLD = 3;
    var pending = GmailApp.search('from:naverbooking_noreply@navercorp.com -in:trash -in:spam label:"엘리야 예약알림" -label:"처리완료" newer_than:14d');
    var count = pending.length;
    var isProblem = (count >= THRESHOLD);

    // 현재 상태: 마지막 줄의 B열(구분)이 "발생"이면 이미 경고 발송된 상태
    var lastStatus = "";
    var lastRow = log.getLastRow();
    if (lastRow >= 2) lastStatus = String(log.getRange(lastRow, 2).getValue()).trim();
    var alreadyAlerting = (lastStatus === "발생");

    var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy.MM.dd. HH:mm:ss');

    if (isProblem && !alreadyAlerting) {
      // 정상 → 문제: 발생 기록 + 메일 (가장 오래된 미처리 시각 계산은 이 전환 순간에만)
      var oldest = "-";
      var oldestDate = null;
      for (var i = 0; i < pending.length; i++) {
        var d = pending[i].getLastMessageDate();
        if (!oldestDate || d < oldestDate) oldestDate = d;
      }
      if (oldestDate) oldest = Utilities.formatDate(oldestDate, 'Asia/Seoul', 'yyyy.MM.dd. HH:mm');
      log.appendRow([now, "발생", count, oldest, "미처리 " + THRESHOLD + "건 이상"]);
      sendBacklogAlert("발생", count, oldest);
    } else if (!isProblem && alreadyAlerting) {
      // 문제 → 정상: 해결 기록 + 메일
      log.appendRow([now, "해결", count, "-", "정상 복구"]);
      sendBacklogAlert("해결", count, "-");
    }
    // 문제 지속 / 정상 지속 → 아무것도 안 함(도배 방지)
  } catch (e) {
    Logger.log('checkBacklogAndAlert 오류: ' + e);
  }
}

function sendBacklogAlert(kind, count, oldest) {
  var TO = "elliyaclinic@gmail.com,biggerjohny@gmail.com,serepina95@gmail.com";
  var subject, body;
  if (kind === "발생") {
    subject = "[엘리야 예약] \u26A0 네이버 예약 동기화 지연 발생";
    body = "네이버 예약 자동 동기화가 지연되고 있습니다.\n\n"
         + "· 미처리 메일: " + count + "건\n"
         + "· 가장 오래된 미처리: " + oldest + "\n\n"
         + "구글시트 예약현황에 아직 반영되지 않은 네이버 예약이 쌓여 있습니다.\n"
         + "Apps Script 실행 기록(Executions) 또는 Gmail '엘리야 예약알림' 라벨을 확인해 주세요.\n\n"
         + "(이 메일은 상태가 정상→지연으로 바뀌는 순간 1회만 발송됩니다.)";
  } else {
    subject = "[엘리야 예약] \u2705 네이버 예약 동기화 정상 복구";
    body = "네이버 예약 자동 동기화가 정상으로 돌아왔습니다.\n\n"
         + "· 현재 미처리 메일: " + count + "건\n";
  }
  try {
    GmailApp.sendEmail(TO, subject, body);
  } catch (e) {
    Logger.log('경고메일 발송 실패(할당량 소진 등일 수 있음): ' + e);
  }
}

// ── 지난 예약 아카이브 (v15) ──
// 시술일(F열)이 오늘보다 과거인 예약을 "지난예약_YYYY" 탭으로 이동한다.
// 오늘 포함 미래 예약은 예약현황원본에 그대로 남긴다("지난 예약은 취소 와도 무시" 정책).
// 매일 03시 트리거로 실행 → 04시 네이버 스윕이 가벼워진 시트를 읽도록 하는 순서.
// 안전장치: 예약번호 기준 중복방지(재실행해도 아카이브 중복 없음) / 아래→위 삭제(행밀림 방지)
//          / 쓰기 구간만 락 / 시술일 파싱 불가 행은 안전하게 남김.
function moveOldBookings() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var src = ss.getSheetByName("예약현황원본");
    if (!src) { Logger.log('moveOldBookings: 예약현황원본 없음'); return; }

    var lastRow = src.getLastRow();
    var lastCol = src.getLastColumn();
    if (lastRow < 2) { Logger.log('moveOldBookings: 이동할 데이터 없음'); return; }

    // ── 읽기·판단 (락 없이) ──
    var COL_SVC_DATE = 6; // F열 = 시술일시
    var COL_BOOKING_NO = 4; // D열 = 예약번호
    var data = src.getRange(2, 1, lastRow - 1, lastCol).getValues();

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var toMove = []; // { rowNum, year, values }
    for (var i = 0; i < data.length; i++) {
      var svcStr = String(data[i][COL_SVC_DATE - 1] || "");
      var m = svcStr.match(/(\d{4})\.(\d{1,2})\.(\d{1,2})/); // YYYY.MM.DD 추출(요일·시간 무관)
      if (!m) continue; // 파싱 불가 → 안전하게 남김
      var svcDate = new Date(parseInt(m[1],10), parseInt(m[2],10) - 1, parseInt(m[3],10));
      if (svcDate < today) {
        toMove.push({ rowNum: i + 2, year: parseInt(m[1],10), values: data[i] });
      }
    }
    if (toMove.length === 0) { Logger.log('moveOldBookings: 이동 대상 없음'); return; }

    // 연도별로 묶기
    var byYear = {};
    toMove.forEach(function(r){ (byYear[r.year] = byYear[r.year] || []).push(r.values); });

    // ── 쓰기 (락) : 아카이브 append + 원본 delete ──
    var lock = LockService.getScriptLock();
    try { lock.waitLock(10000); }
    catch (e) { Logger.log('moveOldBookings 락 획득 실패 — 이번 실행 건너뜀(내일 재시도): ' + e); return; }
    try {
      var movedAt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy.MM.dd. HH:mm:ss'); // 이동(아카이브)시각
      // 1) 연도 탭에 append (예약번호 기준 중복 제거)
      Object.keys(byYear).forEach(function(yr){
        var tabName = "지난예약_" + yr;
        var tab = ss.getSheetByName(tabName);
        if (!tab) { tab = ss.insertSheet(tabName); ensureHeader(tab); }
        else if (tab.getLastRow() === 0) { ensureHeader(tab); }

        var existing = {};
        if (tab.getLastRow() >= 2) {
          var ex = tab.getRange(2, COL_BOOKING_NO, tab.getLastRow() - 1, 1).getValues();
          for (var k = 0; k < ex.length; k++) {
            var exNo = String(ex[k][0] || '').trim();
            if (exNo) existing[exNo] = true; // v22: 빈 예약번호는 키로 등록하지 않음
          }
        }
        // 원본 행(A~R) 뒤에 '이동시각'(맨 뒤 열)을 덧붙여 저장 (v21)
        var rows = byYear[yr]
          .filter(function(v){
            // v22: 예약번호가 비어 있으면(과거 데이터) 중복검사를 건너뛰고 항상 통과시킨다.
            //      빈 값을 키로 쓰면 ""가 하나로 뭉개져 2건째부터 전부 유실됨.
            var no = String(v[COL_BOOKING_NO - 1] || '').trim();
            if (!no) return true;
            return !existing[no];
          })
          .map(function(v){ return v.concat([movedAt]); });
        if (rows.length > 0) {
          var movedCol = rows[0].length; // 이동시각이 들어갈 맨 뒤 열
          if (String(tab.getRange(1, movedCol).getValue()).trim() === '') {
            tab.getRange(1, movedCol).setValue('이동시각'); // 헤더 보장(맨 뒤 열)
          }
          tab.getRange(tab.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
        }
      });

      // 2) 원본에서 삭제 (아래→위 순서로 행밀림 방지)
      var rowNums = toMove.map(function(r){ return r.rowNum; }).sort(function(a, b){ return b - a; });
      for (var j = 0; j < rowNums.length; j++) src.deleteRow(rowNums[j]);

      var summary = Object.keys(byYear).map(function(y){ return y + "년 " + byYear[y].length + "건"; }).join(', ');
      Logger.log('moveOldBookings 완료: 총 ' + toMove.length + '건 이동 (' + summary + ')');
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    Logger.log('moveOldBookings 오류: ' + err);
  }
}

// ── 확정 처리 ──
// ── 지난 예약 재유입 방지 (v20) ──
// 시술일(F)이 오늘보다 과거인 예약은 예약현황원본에 넣지 않는다.
// 이유: 03시 moveOldBookings가 지난 예약을 지난예약_YYYY로 옮긴 뒤, 04시 14일 전체
//   스윕이 같은(지난) 네이버 메일을 다시 읽어 원본에 재삽입하던 충돌을 막기 위함.
//   (bookingNoExists는 원본만 보므로, 옮겨진 건을 "없음"으로 오판해 다시 넣었음)
// 신규 확정/변경 예약은 시술일이 항상 미래라 이 가드에 안 걸리고, 오직 스윕이 재읽는
//   '이미 지난' 예약만 걸러진다. 파싱 불가 시엔 막지 않음(예약 유실 방지).
function isPastServiceDate(dateTimeStr) {
  var m = String(dateTimeStr || "").match(/(\d{4})\.(\d{1,2})\.(\d{1,2})/);
  if (!m) return false;
  var svc = new Date(parseInt(m[1],10), parseInt(m[2],10) - 1, parseInt(m[3],10));
  var today = new Date(); today.setHours(0, 0, 0, 0);
  return svc < today;
}

function handleConfirm(sheet, body) {
  var name      = extract(body, '예약자명', '예약신청');
  var bookedAt  = extract(body, '예약신청 일시', '예약내역');
  var bookingNo = extract(body, '예약번호', '예약상품');
  var service   = extract(body, '예약상품', '이용일시');
  var dateTime  = extract(body, '이용일시', '결제상태');
  var menu      = extract(body, '선택메뉴', '요청사항');

  // 지난 예약 재유입 방지(v20): 시술일이 과거면 원본에 넣지 않음(지난예약_YYYY 관리 대상)
  if (isPastServiceDate(dateTime)) {
    Logger.log('handleConfirm: 시술일 과거 — 원본에 넣지 않음: ' + bookingNo + ' / ' + dateTime);
    return;
  }

  // 시트에 쓰는 순간만 짧게 락을 잡음(중복확인+기록을 하나의 원자적 작업으로)
  var lockW = LockService.getScriptLock();
  try {
    lockW.waitLock(5000);
  } catch (lockErr) {
    Logger.log('handleConfirm 락 획득 실패, 다음 실행에서 재시도됨: ' + lockErr);
    return; // 다음 트리거 실행(1분 뒤) 때 bookingNoExists 재확인 후 안전하게 재처리됨
  }
  try {
    // 같은 예약번호가 이미 시트에 있으면 (라벨 타이밍 등으로 중복 처리되더라도) 다시 쓰지 않음
    if (bookingNo && bookingNoExists(sheet, bookingNo)) {
      return;
    }
    var processedAt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy.MM.dd. HH:mm:ss');
    sheet.appendRow([
      bookedAt, '네이버', '확정', bookingNo,
      name, dateTime, service,
      '', '', '', '', menu,
      '', '', '', '', '',          // M~Q: (추후 사용을 위해 비움)
      processedAt                  // R: 처리시각
    ]);
  } finally {
    lockW.releaseLock();
  }
}

// ── 예약번호가 시트(D열)에 이미 존재하는지 확인 ──
function bookingNoExists(sheet, bookingNo) {
  // 예약번호가 비어있으면("추출 실패" 포함) 절대 "이미 있다"고 판단하지 않음.
  // (차단/취소 등으로 D열이 빈 값인 행이 많아지면서, 빈 문자열끼리 매칭되어
  //  정상적인 새 네이버 예약이 "이미 기록됨"으로 잘못 판정되어 누락되는 사고를 방지)
  if (!bookingNo || !String(bookingNo).trim()) return false;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var dColumn = sheet.getRange(2, 4, lastRow - 1, 1).getValues(); // D열 전체
  for (var i = 0; i < dColumn.length; i++) {
    var cell = String(dColumn[i][0]).trim();
    if (cell && cell === String(bookingNo).trim()) return true;
  }
  return false;
}

// ── 취소 처리 (네이버 이메일 기반) ──
function handleCancel(sheet, body) {
  var bookingNo = extract(body, '예약번호', '예약상품');
  markStatusByBookingNo(sheet, bookingNo, '취소');
}

// ── 변경 처리 (네이버 이메일 기반) ──
function handleChange(sheet, body) {
  var name       = extract(body, '예약자명', '예약변경');
  var bookedAt   = extract(body, '예약변경 일시', '신규예약내역');
  var newBlock   = body.indexOf('신규예약내역') !== -1 ? body.split('신규예약내역')[1] : '';
  var cancelBlock = body.indexOf('예약취소내역') !== -1 ? body.split('예약취소내역')[1] : '';

  if (cancelBlock) {
    var oldNo = extract(cancelBlock, '예약번호', '예약상품');
    markStatusByBookingNo(sheet, oldNo, '취소(변경)');
  }

  if (newBlock) {
    var newNo      = extract(newBlock, '예약번호', '예약상품');
    var newService = extract(newBlock, '예약상품', '이용일시');
    var newDate    = extract(newBlock, '이용일시', '결제상태');
    var newMenu    = extract(newBlock, '선택메뉴', '요청사항');

    // 지난 예약 재유입 방지(v20): 변경된 신규 시술일이 과거면 원본에 넣지 않음
    if (isPastServiceDate(newDate)) {
      Logger.log('handleChange: 신규 시술일 과거 — 원본에 넣지 않음: ' + newNo + ' / ' + newDate);
      return;
    }

    var lockW2 = LockService.getScriptLock();
    try {
      lockW2.waitLock(5000);
    } catch (lockErr) {
      Logger.log('handleChange(신규) 락 획득 실패, 다음 실행에서 재시도됨: ' + lockErr);
      return;
    }
    try {
      if (newNo && bookingNoExists(sheet, newNo)) {
        return;
      }
      var processedAt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy.MM.dd. HH:mm:ss');
      sheet.appendRow([
        bookedAt, '네이버', '확정(변경)', newNo,
        name, newDate, newService,
        '', '', '', '', newMenu,
        '', '', '', '', '',          // M~Q: (추후 사용을 위해 비움)
        processedAt                  // R: 처리시각
      ]);
    } finally {
      lockW2.releaseLock();
    }
  }
}

// ── 예약번호로 상태 업데이트 ──
function markStatusByBookingNo(sheet, bookingNo, status) {
  if (!bookingNo) return;
  var lockW = LockService.getScriptLock();
  try {
    lockW.waitLock(5000);
  } catch (lockErr) {
    Logger.log('markStatusByBookingNo 락 획득 실패, 다음 실행에서 재시도됨: ' + lockErr);
    return;
  }
  try {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][3]).trim() === bookingNo.trim()) {
        // 이미 같은 상태면 재처리하지 않음(중복 방지). 라벨과 무관하게 3일치를 매번 재검색하는
        // 구조라, 이게 없으면 이미 취소/변경된 행을 매 실행마다 다시 덮어써 처리시각(R열)이
        // 계속 갱신됐음(v16 수정). 처리시각은 최초 상태변경 1회로 고정.
        if (String(data[i][2]).trim() === status) return; // C열 = 현재 상태
        sheet.getRange(i + 1, 3).setValue(status); // C열 = 상태
        var processedAt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy.MM.dd. HH:mm:ss');
        sheet.getRange(i + 1, 18).setValue(processedAt); // R열 = 처리시각(최초 상태변경 시각)
        return;
      }
    }
  } finally {
    lockW.releaseLock();
  }
}

// ── 텍스트 파싱 헬퍼 ──
function extract(text, startKey, endKey) {
  var start = text.indexOf(startKey);
  if (start === -1) return '';
  start += startKey.length;
  while (start < text.length &&
         (text[start] === ':' || text[start] === ' ' || text[start] === '\n' || text[start] === '\r')) {
    start++;
  }
  if (!endKey) {
    var result = text.substring(start).trim().split('\n')[0].trim();
    return result.replace(/<!--.*?-->/g, '').trim();
  }
  var end = text.indexOf(endKey, start);
  if (end === -1) {
    var result = text.substring(start).trim().split('\n')[0].trim();
    return result.replace(/<!--.*?-->/g, '').trim();
  }
  var result = text.substring(start, end).trim().replace(/[\n\r]/g, ' ');
  return result.replace(/<!--.*?-->/g, '').trim();
}

// ══════════════════════════════════════════════
// 중복 행 제거 (일회성 - 실행 후 삭제해도 됨)
// ══════════════════════════════════════════════
function removeDuplicateBookingRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("예약현황원본");
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return; // 헤더만 있으면 할 일 없음

  var seen = {};
  var rowsToDelete = []; // 시트 상의 실제 행 번호 (1-based)

  for (var i = 1; i < data.length; i++) { // i=0은 헤더라서 건너뜀
    // R열(처리시각, 마지막 값)은 매번 달라서 비교 키에서 제외 (A~Q열, 즉 처리시각 빼고 비교)
    var rowKey = data[i].slice(0, -1).join('|||');
    if (seen[rowKey]) {
      rowsToDelete.push(i + 1); // 시트 행 번호는 1부터 시작 + 헤더 1행
    } else {
      seen[rowKey] = true;
    }
  }

  // 뒤에서부터 삭제해야 앞 행 번호가 안 꼬임
  for (var j = rowsToDelete.length - 1; j >= 0; j--) {
    sheet.deleteRow(rowsToDelete[j]);
  }

  Logger.log(rowsToDelete.length + '개의 중복 행을 삭제했습니다.');
}

// ══════════════════════════════════════════════
// 예약 차단 기능 (구글시트 메뉴)
// - 기존 doGet/doPost는 건드리지 않음. 재배포 불필요.
// - 차단 데이터는 기존 시트에 예약경로="차단"으로 직접 추가되어,
//   기존 get_booked_times 로직이 자동으로 그 시간을 막아줌.
// ══════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('[예약 차단]')
    .addItem('날짜 지정해서 차단', 'showBlockDialog')
    .addItem('차단 해제', 'showUnblockDialog')
    .addToUi();
}

// ── 차단 등록 ──

function showBlockDialog() {
  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Arial,sans-serif;font-size:13px;padding:4px 8px;}' +
    'label{display:block;margin:10px 0 4px;font-weight:bold;}' +
    'input{width:100%;padding:6px;box-sizing:border-box;font-size:13px;}' +
    'button{margin-top:16px;padding:8px 14px;background:#1C1916;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:13px;}' +
    'button:disabled{opacity:.5;}' +
    '#msg{margin-top:10px;color:#c0392b;font-size:12px;}' +
    '</style>' +
    '<label>날짜</label><input type="date" id="d">' +
    '<label>시작 시간</label><input type="time" id="s" value="12:30">' +
    '<label>종료 시간</label><input type="time" id="e" value="13:30">' +
    '<label>관리자 이름</label><input type="text" id="a" placeholder="예: 홍준택 원장">' +
    '<label>메모 (선택)</label><input type="text" id="m" placeholder="예: 점심시간, 학회 참석">' +
    '<button id="btn" onclick="submitBlock()">차단 등록</button>' +
    '<div id="msg"></div>' +
    '<script>' +
    'function submitBlock(){' +
    '  var d=document.getElementById("d").value;' +
    '  var s=document.getElementById("s").value;' +
    '  var e=document.getElementById("e").value;' +
    '  var a=document.getElementById("a").value;' +
    '  var m=document.getElementById("m").value;' +
    '  if(!d||!s||!e){document.getElementById("msg").innerText="날짜/시작/종료를 모두 입력해 주세요.";return;}' +
    '  if(!a){document.getElementById("msg").innerText="관리자 이름을 입력해 주세요.";return;}' +
    '  if(s>=e){document.getElementById("msg").innerText="종료 시간이 시작 시간보다 늦어야 합니다.";return;}' +
    '  document.getElementById("btn").disabled=true;' +
    '  google.script.run.withSuccessHandler(function(res){' +
    '    if(res && res.ok){ document.getElementById("msg").style.color="#2a7"; document.getElementById("msg").innerText="등록 완료 (" + res.count + "건). 창을 닫습니다."; setTimeout(google.script.host.close, 1200); }' +
    '    else { document.getElementById("btn").disabled=false; document.getElementById("msg").innerText="오류: " + (res && res.error); }' +
    '  }).withFailureHandler(function(err){' +
    '    document.getElementById("btn").disabled=false;' +
    '    document.getElementById("msg").innerText="오류: " + err.message;' +
    '  }).registerBlock(d, s, e, a, m);' +
    '}' +
    '</script>'
  ).setWidth(340).setHeight(340);
  SpreadsheetApp.getUi().showModalDialog(html, '예약 차단 등록');
}

function registerBlock(dateStr, startStr, endStr, adminName, memo) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(5000); // 다른 락들과 대기시간 통일(5초)

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("예약현황원본");
    ensureHeader(sheet);

    var dParts = dateStr.split('-');
    var dateObj = new Date(parseInt(dParts[0],10), parseInt(dParts[1],10)-1, parseInt(dParts[2],10));

    var sParts = startStr.split(':');
    var eParts = endStr.split(':');
    var startMin = parseInt(sParts[0],10)*60 + parseInt(sParts[1],10);
    var endMin   = parseInt(eParts[0],10)*60 + parseInt(eParts[1],10);

    var regNow = new Date();
    var regStr = Utilities.formatDate(regNow, Session.getScriptTimeZone(), "yyyy.MM.dd. HH:mm:ss");
    var memoFinal = memo && memo.trim() ? memo.trim() : '차단';

    var rows = [];
    for (var t = startMin; t < endMin; t += 60) {
      var hh = Math.floor(t/60), mm = t%60;
      var dtStr = formatKoreanDateTime(dateObj, hh, mm);
      rows.push([
        regStr,        // A 예약일시(등록시각)
        '차단',         // B 예약경로
        '확정',         // C 상태
        genBookingNo('K', regNow), // D 예약번호 (v22 — 시간대별로 각각 다른 번호)
        adminName,     // E 이름 (관리자 이름)
        dtStr,         // F 시술일시
        memoFinal,     // G 시술명
        '', '', '', '', '', // H~L
        '', '', '', '',     // M~P
        '',            // Q 네이버처리시각
        regStr         // R 처리시각 (대시보드 등에서 행 식별자로 쓰이므로 채워둠)
      ]);
    }

    if (rows.length === 0) return { ok:false, error:'등록할 시간대가 없습니다.' };

    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

    return { ok:true, count: rows.length };
  } catch (err) {
    return { ok:false, error: err.message };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

function formatKoreanDateTime(dateObj, hh24, mm) {
  var days = ['일','월','화','수','목','금','토'];
  var y = dateObj.getFullYear();
  var mo = dateObj.getMonth()+1;
  var d = dateObj.getDate();
  var dow = days[dateObj.getDay()];
  var ampm = hh24 < 12 ? '오전' : '오후';
  var h12 = hh24 % 12; if (h12 === 0) h12 = 12;
  return y+'.'+pad2(mo)+'.'+pad2(d)+'.('+dow+') '+ampm+' '+h12+':'+pad2(mm);
}
function pad2(n){ return (n<10?'0':'')+String(n); }

// ── 차단 해제 ──

function showUnblockDialog() {
  var groups = getBlockGroups_();
  if (groups.length === 0) {
    SpreadsheetApp.getUi().alert('오늘 이후로 등록된 차단이 없습니다.');
    return;
  }
  var itemsHtml = groups.map(function(g, i) {
    return '<label style="display:block;padding:8px 0;border-bottom:1px solid #eee;">' +
      '<input type="checkbox" class="gb" value="' + i + '"> ' + g.label + '</label>';
  }).join('');

  var html = HtmlService.createHtmlOutput(
    '<style>' +
    'body{font-family:Arial,sans-serif;font-size:13px;padding:4px 8px;}' +
    'button{margin-top:14px;padding:8px 14px;background:#c0392b;color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:13px;}' +
    'button:disabled{opacity:.5;}' +
    '#msg{margin-top:10px;font-size:12px;}' +
    '</style>' +
    '<div>' + itemsHtml + '</div>' +
    '<button id="btn" onclick="submitUnblock()">선택 항목 해제</button>' +
    '<div id="msg"></div>' +
    '<script>' +
    'var GROUPS = ' + JSON.stringify(groups) + ';' +
    'function submitUnblock(){' +
    '  var checked = Array.prototype.slice.call(document.querySelectorAll(".gb:checked")).map(function(el){return parseInt(el.value,10);});' +
    '  if(checked.length===0){document.getElementById("msg").innerText="해제할 항목을 선택해 주세요.";return;}' +
    '  document.getElementById("btn").disabled=true;' +
    '  var rowsToDelete = [];' +
    '  checked.forEach(function(i){ rowsToDelete = rowsToDelete.concat(GROUPS[i].rows); });' +
    '  google.script.run.withSuccessHandler(function(res){' +
    '    document.getElementById("msg").style.color="#2a7";' +
    '    document.getElementById("msg").innerText="해제 완료 (" + res.count + "건). 창을 닫습니다.";' +
    '    setTimeout(google.script.host.close, 1200);' +
    '  }).withFailureHandler(function(err){' +
    '    document.getElementById("btn").disabled=false;' +
    '    document.getElementById("msg").innerText="오류: " + err.message;' +
    '  }).deleteBlockRows(rowsToDelete);' +
    '}' +
    '</script>'
  ).setWidth(380).setHeight(Math.min(500, 120 + groups.length*40));
  SpreadsheetApp.getUi().showModalDialog(html, '차단 해제');
}

function getBlockGroups_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("예약현황원본");
  var data = sheet.getDataRange().getValues();
  var today = new Date(); today.setHours(0,0,0,0);

  var map = {}; // key -> {rows:[], date:Date, times:[], memo:''}
  for (var i = 1; i < data.length; i++) {
    var path = String(data[i][1] || '').trim();
    if (path !== '차단') continue;

    var dtStr = String(data[i][5] || '').trim();
    var m = dtStr.match(/(\d{4})\.(\d{2})\.(\d{2})\.\([^)]+\)\s+(오전|오후)\s+(\d+):(\d+)/);
    if (!m) continue;
    var rowDate = new Date(parseInt(m[1],10), parseInt(m[2],10)-1, parseInt(m[3],10));
    if (rowDate < today) continue; // 오늘 이후만

    var regKey = String(data[i][0] || '');
    var memo = String(data[i][6] || '');
    var key = regKey + '|' + m[1]+m[2]+m[3] + '|' + memo;

    var h = parseInt(m[5],10), mm = parseInt(m[6],10);
    if (m[4] === '오후' && h !== 12) h += 12;
    if (m[4] === '오전' && h === 12) h = 0;
    var minutes = h*60+mm;

    if (!map[key]) map[key] = { rows: [], date: rowDate, memo: memo, times: [] };
    map[key].rows.push(i+1); // 실제 시트 행번호(1-indexed, 헤더 포함이므로 +1)
    map[key].times.push(minutes);
  }

  var groups = Object.keys(map).map(function(k){
    var g = map[k];
    g.times.sort(function(a,b){return a-b;});
    var startMin = g.times[0];
    var endMin = g.times[g.times.length-1] + 30;
    var fmt = function(mn){
      var h=Math.floor(mn/60), mi=mn%60;
      var ap = h<12?'오전':'오후'; var h12=h%12; if(h12===0)h12=12;
      return ap+' '+h12+':'+pad2(mi);
    };
    var days = ['일','월','화','수','목','금','토'];
    var dLabel = g.date.getFullYear()+'.'+pad2(g.date.getMonth()+1)+'.'+pad2(g.date.getDate())+'('+days[g.date.getDay()]+')';
    return {
      label: dLabel + ' ' + fmt(startMin) + '~' + fmt(endMin) + ' · ' + g.memo,
      rows: g.rows
    };
  });

  groups.sort(function(a,b){ return a.label < b.label ? -1 : 1; });
  return groups;
}

function deleteBlockRows(rowNumbers) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("예약현황원본");
  // 아래에서 위로 삭제해야 행 번호가 안 밀림
  var sorted = rowNumbers.slice().sort(function(a,b){ return b-a; });
  sorted.forEach(function(r){ sheet.deleteRow(r); });
  return { ok:true, count: rowNumbers.length };
}
