// ============================================================
// ML/CNN 레벨 테스트 — Google Apps Script 수신 코드
// ============================================================
// 설치 방법:
// 1. Google Sheets 새 파일 생성
// 2. 상단 메뉴 → 확장 프로그램 → Apps Script
// 3. 아래 코드 전체를 붙여넣기
// 4. 저장 → 배포 → 새 배포 → 웹 앱
//    - 실행 사용자: 나
//    - 액세스 권한: 모든 사용자 (익명 포함)
// 5. 배포 후 나오는 URL을 복사해서
//    quiz HTML 파일의 APPS_SCRIPT_URL 변수에 붙여넣기
// ============================================================

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = JSON.parse(e.postData.contents);

    // 헤더가 없으면 첫 행에 생성
    if (sheet.getLastRow() === 0) {
      const headers = buildHeaders();
      sheet.appendRow(headers);
      // 헤더 스타일
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#1a1714');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setFontSize(11);
      sheet.setFrozenRows(1);
    }

    // 데이터 행 구성
    const row = buildRow(data);
    sheet.appendRow(row);

    // 자동 채점 결과 열 강조
    const lastRow = sheet.getLastRow();
    colorizeScore(sheet, lastRow, data);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// GET 요청도 허용 (배포 테스트용)
function doGet(e) {
  return ContentService
    .createTextOutput('ML/CNN Quiz Sheet is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ──────────────────────────────────────────────
// 헤더 구성
// ──────────────────────────────────────────────
function buildHeaders() {
  return [
    '제출 시간', '이름',
    // Q1
    'Q1-로직A(학생)', 'Q1-로직B(학생)', 'Q1-로직C(학생)',
    'Q1-로직A(정답)', 'Q1-로직B(정답)', 'Q1-로직C(정답)',
    'Q1-1번(학생)', 'Q1-2번(학생)', 'Q1-3번(학생)', 'Q1-4번(학생)', 'Q1-5번(학생)',
    'Q1-1번(정답)', 'Q1-2번(정답)', 'Q1-3번(정답)', 'Q1-4번(정답)', 'Q1-5번(정답)',
    // Q2
    'Q2-로직A(학생)', 'Q2-로직B(학생)', 'Q2-로직C(학생)',
    'Q2-로직A(정답)', 'Q2-로직B(정답)', 'Q2-로직C(정답)',
    'Q2-1번(학생)', 'Q2-2번(학생)', 'Q2-3번(학생)', 'Q2-4번(학생)', 'Q2-5번(학생)', 'Q2-6번(학생)',
    'Q2-1번(정답)', 'Q2-2번(정답)', 'Q2-3번(정답)', 'Q2-4번(정답)', 'Q2-5번(정답)', 'Q2-6번(정답)',
    // Q3
    'Q3-로직A(학생)', 'Q3-로직B(학생)', 'Q3-로직C(학생)',
    'Q3-로직A(정답)', 'Q3-로직B(정답)', 'Q3-로직C(정답)',
    'Q3-1번(학생)', 'Q3-2번(학생)', 'Q3-3번(학생)', 'Q3-4번(학생)', 'Q3-5번(학생)', 'Q3-6번(학생)', 'Q3-7번(학생)',
    'Q3-1번(정답)', 'Q3-2번(정답)', 'Q3-3번(정답)', 'Q3-4번(정답)', 'Q3-5번(정답)', 'Q3-6번(정답)', 'Q3-7번(정답)',
    // Q4
    'Q4-로직A(학생)', 'Q4-로직B(학생)', 'Q4-로직C(학생)',
    'Q4-로직A(정답)', 'Q4-로직B(정답)', 'Q4-로직C(정답)',
    'Q4-1번(학생)', 'Q4-2번(학생)', 'Q4-3번(학생)', 'Q4-4번(학생)', 'Q4-5번(학생)', 'Q4-6번(학생)',
    'Q4-1번(정답)', 'Q4-2번(정답)', 'Q4-3번(정답)', 'Q4-4번(정답)', 'Q4-5번(정답)', 'Q4-6번(정답)',
    // Q5
    'Q5-로직A(학생)', 'Q5-연산자1(학생)', 'Q5-로직B(학생)', 'Q5-연산자2(학생)', 'Q5-로직C(학생)',
    'Q5-로직A(정답)', 'Q5-연산자1(정답)', 'Q5-로직B(정답)', 'Q5-연산자2(정답)', 'Q5-로직C(정답)',
    'Q5-1번(학생)', 'Q5-2번(학생)', 'Q5-3번(학생)', 'Q5-4번(학생)', 'Q5-5번(학생)', 'Q5-6번(학생)',
    'Q5-1번(정답)', 'Q5-2번(정답)', 'Q5-3번(정답)', 'Q5-4번(정답)', 'Q5-5번(정답)', 'Q5-6번(정답)',
    // Q6
    'Q6(학생)', 'Q6(정답)',
    // Q7
    'Q7-1(학생)', 'Q7-2(학생)', 'Q7-1(정답)', 'Q7-2(정답)',
    // Q8
    'Q8-경로(학생)', 'Q8-경로(정답)', 'Q8-비용(학생)', 'Q8-비용(정답)',
    // Q9
    'Q9-최대값(학생)', 'Q9-최소값(학생)', 'Q9-로직설명(학생)',
    'Q9-최대값(정답)', 'Q9-최소값(정답)',
    // Q10
    'Q10-1등(학생)', 'Q10-2등(학생)', 'Q10-3등(학생)', 'Q10-4등(학생)',
    'Q10-로직설명(학생)',
    'Q10-1등(정답)', 'Q10-2등(정답)', 'Q10-3등(정답)', 'Q10-4등(정답)',
    // Q★
    'Q★-1(학생)', 'Q★-2설명(학생)', 'Q★-3(학생)',
    'Q★-1(정답)', 'Q★-3(정답)',
    // 채점
    '자동채점 정답수', '자동채점 총문항'
  ];
}

// ──────────────────────────────────────────────
// 데이터 행 구성
// ──────────────────────────────────────────────
function buildRow(d) {
  const AUTO_GRADE_KEYS = [
    'q1_r1','q1_r2','q1_r3','q1_r4','q1_r5',
    'q2_r1','q2_r2','q2_r3','q2_r4','q2_r5','q2_r6',
    'q3_r1','q3_r2','q3_r3','q3_r4','q3_r5','q3_r6','q3_r7',
    'q4_r1','q4_r2','q4_r3','q4_r4','q4_r5','q4_r6',
    'q5_r1','q5_r2','q5_r3','q5_r4','q5_r5','q5_r6',
    'q6',
    'q7_1','q7_2',
    'q8_path','q8_cost',
    'q9_max','q9_min',
    'q10_1','q10_2','q10_3','q10_4',
    'qs_1','qs_3'
  ];

  let correct = 0;
  AUTO_GRADE_KEYS.forEach(key => {
    const ansKey = 'ans_' + key;
    if (d[ansKey] && d[key]) {
      if (d[key].toLowerCase().replace(/\s/g,'') === d[ansKey].toLowerCase().replace(/\s/g,'')) correct++;
    }
  });

  return [
    d.timestamp, d.name,
    d.q1_a, d.q1_b, d.q1_c,
    d.ans_q1_a, d.ans_q1_b, d.ans_q1_c,
    d.q1_r1, d.q1_r2, d.q1_r3, d.q1_r4, d.q1_r5,
    d.ans_q1_r1, d.ans_q1_r2, d.ans_q1_r3, d.ans_q1_r4, d.ans_q1_r5,
    d.q2_a, d.q2_b, d.q2_c,
    d.ans_q2_a, d.ans_q2_b, d.ans_q2_c,
    d.q2_r1, d.q2_r2, d.q2_r3, d.q2_r4, d.q2_r5, d.q2_r6,
    d.ans_q2_r1, d.ans_q2_r2, d.ans_q2_r3, d.ans_q2_r4, d.ans_q2_r5, d.ans_q2_r6,
    d.q3_a, d.q3_b, d.q3_c,
    d.ans_q3_a, d.ans_q3_b, d.ans_q3_c,
    d.q3_r1, d.q3_r2, d.q3_r3, d.q3_r4, d.q3_r5, d.q3_r6, d.q3_r7,
    d.ans_q3_r1, d.ans_q3_r2, d.ans_q3_r3, d.ans_q3_r4, d.ans_q3_r5, d.ans_q3_r6, d.ans_q3_r7,
    d.q4_a, d.q4_b, d.q4_c,
    d.ans_q4_a, d.ans_q4_b, d.ans_q4_c,
    d.q4_r1, d.q4_r2, d.q4_r3, d.q4_r4, d.q4_r5, d.q4_r6,
    d.ans_q4_r1, d.ans_q4_r2, d.ans_q4_r3, d.ans_q4_r4, d.ans_q4_r5, d.ans_q4_r6,
    d.q5_a, d.q5_op1, d.q5_b, d.q5_op2, d.q5_c,
    d.ans_q5_a, d.ans_q5_op1, d.ans_q5_b, d.ans_q5_op2, d.ans_q5_c,
    d.q5_r1, d.q5_r2, d.q5_r3, d.q5_r4, d.q5_r5, d.q5_r6,
    d.ans_q5_r1, d.ans_q5_r2, d.ans_q5_r3, d.ans_q5_r4, d.ans_q5_r5, d.ans_q5_r6,
    d.q6, d.ans_q6,
    d.q7_1, d.q7_2, d.ans_q7_1, d.ans_q7_2,
    d.q8_path, d.ans_q8_path, d.q8_cost, d.ans_q8_cost,
    d.q9_max, d.q9_min, d.q9_logic,
    d.ans_q9_max, d.ans_q9_min,
    d.q10_1, d.q10_2, d.q10_3, d.q10_4,
    d.q10_logic,
    d.ans_q10_1, d.ans_q10_2, d.ans_q10_3, d.ans_q10_4,
    d.qs_1, d.qs_2, d.qs_3,
    d.ans_qs_1, d.ans_qs_3,
    correct, AUTO_GRADE_KEYS.length
  ];
}

// ──────────────────────────────────────────────
// 채점 결과 시각화
// ──────────────────────────────────────────────
function colorizeScore(sheet, row, data) {
  const lastCol = sheet.getLastColumn();
  // 채점 점수 열 (마지막 2열)
  const scoreCell = sheet.getRange(row, lastCol - 1);
  const totalCell = sheet.getRange(row, lastCol);
  scoreCell.setFontWeight('bold');
  scoreCell.setBackground('#e6f4ec');
  scoreCell.setFontColor('#1a6b3a');
}
