// ============================================================
// 구글 시트 설문 응답 저장 Apps Script (업데이트본)
// ------------------------------------------------------------
//  - NFC(Q4) 15문항 반영
//  - 시나리오 4번째 추가
//  - HTML 서빙용 doGet() 추가
//  - google.script.run 용 submitSurvey() 추가
// ============================================================

const SHEET_NAME = '설문응답';
const HTML_FILE  = 'Index';   // Index.html 파일명과 동일하게 유지
const NOTIFY_EMAIL = '54ryh123@pusan.ac.kr'; // 알림 받을 이메일

// HTML 웹앱 서빙
function doGet() {
  return HtmlService.createHtmlOutputFromFile(HTML_FILE)
    .setTitle('연구 설문조사')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// HTML에서 google.script.run.submitSurvey(data) 로 직접 호출
function submitSurvey(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) { sheet = ss.insertSheet(SHEET_NAME); createHeaders(sheet); }
    if (sheet.getLastRow() === 0) createHeaders(sheet);

    // 6.2 중복 응답 방지 (SONA ID가 입력된 경우에만 체크)
    if (data.sonaId && data.sonaId.trim() !== '') {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const sonaIds = sheet.getRange(2, 4, lastRow - 1, 1).getValues(); // D열: SONA_ID
        const isDuplicate = sonaIds.some(row => String(row[0]).trim() === data.sonaId.trim());
        if (isDuplicate) {
          return { result: 'error', error: 'DUPLICATE_SONA_ID' };
        }
      }
    }

    appendDataRow(sheet, data);
    sendNewResponseEmail(data);

    return { result: 'success', row: sheet.getLastRow() };
  } finally {
    lock.releaseLock();
  }
}

// 6.3 새 응답 제출 시 이메일 알림
function sendNewResponseEmail(data) {
  try {
    const subject = '[설문 응답] 새 응답이 제출되었습니다';
    const body =
      '새로운 설문 응답이 접수되었습니다.\n\n' +
      '제출시각: ' + (data.timestamp || '') + '\n' +
      '조건번호: ' + (data.condition || '') + '\n' +
      'SONA ID: ' + (data.sonaId || '(미입력)') + '\n' +
      '성별: ' + (data.gender || '') + '\n' +
      '연령: ' + (data.age || '');
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  } catch (err) {
    Logger.log('이메일 발송 실패: ' + err.toString());
  }
}

// ============================================================
// 균등 배분 — 시트 기존 응답을 읽고 가장 적게 배정된 조건을 반환
//   - 8개 조건 카운트를 세고 최소 카운트인 조건들 중 무작위 선택
//   - LockService로 동시 접속 시 중복/편향 방지
// ============================================================
function assignCondition() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) { sheet = ss.insertSheet(SHEET_NAME); createHeaders(sheet); }
    if (sheet.getLastRow() === 0) createHeaders(sheet);

    const MAX_PER_CONDITION = 40;
    const cache = CacheService.getScriptCache();
    const cachedCounts = cache.get('condition_counts');
    let counts;

    // 캐시 있으면 사용, 없으면 시트에서 읽고 캐시에 저장
    if (cachedCounts) {
      counts = JSON.parse(cachedCounts);
    } else {
      counts = [0,0,0,0,0,0,0,0];
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
        for (let i = 0; i < values.length; i++) {
          const c = parseInt(values[i][0], 10);
          if (c >= 1 && c <= 8) counts[c-1]++;
        }
      }
      cache.put('condition_counts', JSON.stringify(counts), 60); // 60초 캐시
    }

    // 40명 미만 조건들 중 카운트 최소값 찾기
    let minCount = MAX_PER_CONDITION;
    for (let i = 0; i < 8; i++) {
      if (counts[i] < MAX_PER_CONDITION && counts[i] < minCount) {
        minCount = counts[i];
      }
    }
    const candidates = [];
    for (let i = 0; i < 8; i++) {
      if (counts[i] < MAX_PER_CONDITION && counts[i] === minCount) {
        candidates.push(i + 1);
      }
    }

    // 모든 조건이 40명 도달 → 마감
    if (candidates.length === 0) return -1;

    const chosen = candidates[Math.floor(Math.random() * candidates.length)];

    // 캐시 업데이트 (방금 배정한 조건 +1) — 다음 사용자도 빠르게 응답
    counts[chosen - 1]++;
    cache.put('condition_counts', JSON.stringify(counts), 60);

    return chosen;
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result;
    if (data.action === 'assignCondition') {
      result = { condition: assignCondition() };
    } else {
      result = submitSurvey(data);
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error', error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function createHeaders(sheet) {
  const headers = [
    // 메타
    '제출시각', '조건번호',
    // 기본정보
    '동의여부', 'SONA_ID', '성별', '연령',
    // Q3 — AI 활용습관
    'Q3_habit_1','Q3_habit_2','Q3_habit_3',
    // Q4 — NFC
    'Q4_noc_1','Q4_noc_2','Q4_noc_3','Q4_noc_4','Q4_noc_5','Q4_noc_6','Q4_noc_7','Q4_noc_8',
    // 시나리오 1
    '시나리오1_초기선택','시나리오1_AI확인','시나리오1_최종선택','시나리오1_선택변경여부',
    // 시나리오 2
    '시나리오2_초기선택','시나리오2_AI확인','시나리오2_최종선택','시나리오2_선택변경여부',
    // 시나리오 3
    '시나리오3_초기선택','시나리오3_AI확인','시나리오3_최종선택','시나리오3_선택변경여부',
    // 시나리오 4
    '시나리오4_초기선택','시나리오4_AI확인','시나리오4_최종선택','시나리오4_선택변경여부',
    // Q14
    'Q14_conf_1','Q14_conf_2','Q14_conf_3','Q14_trust_1','Q14_trust_2','Q14_trust_3','Q14_frame_1','Q14_frame_2',
    // Q15
    'Q15_resp_1','Q15_resp_2','Q15_resp_3','Q15_resp_4',
    // Q16
    'Q16_advice_1','Q16_advice_2','Q16_advice_3'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4f46e5');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  for (let i = 1; i <= headers.length; i++) sheet.autoResizeColumn(i);
  sheet.setFrozenRows(1);
}

function appendDataRow(sheet, data) {
  const q3  = data.q3  || {};
  const q4  = data.q4  || {};
  const q14 = data.q14 || {};
  const q15 = data.q15 || {};
  const q16 = data.q16 || {};

  const row = [
    // 메타
    Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
    data.condition || '',
    // 기본정보
    data.consent || '',
    data.sonaId  || '',
    data.gender  || '',
    data.age     || '',
    // Q3
    q3.habit_1 || '', q3.habit_2 || '', q3.habit_3 || '',
    // Q4 (8)
    q4.noc_1 || '', q4.noc_2 || '', q4.noc_3 || '', q4.noc_4 || '',
    q4.noc_5 || '', q4.noc_6 || '', q4.noc_7 || '', q4.noc_8 || '',
    // 시나리오 1
    data.scenario1_choice || '',
    data.scenario1_ai_confirmed ? 'TRUE' : 'FALSE',
    data.scenario1_final_choice || '',
    data.scenario1_changed ? 'TRUE' : 'FALSE',
    // 시나리오 2
    data.scenario2_choice || '',
    data.scenario2_ai_confirmed ? 'TRUE' : 'FALSE',
    data.scenario2_final_choice || '',
    data.scenario2_changed ? 'TRUE' : 'FALSE',
    // 시나리오 3
    data.scenario3_choice || '',
    data.scenario3_ai_confirmed ? 'TRUE' : 'FALSE',
    data.scenario3_final_choice || '',
    data.scenario3_changed ? 'TRUE' : 'FALSE',
    // 시나리오 4
    data.scenario4_choice || '',
    data.scenario4_ai_confirmed ? 'TRUE' : 'FALSE',
    data.scenario4_final_choice || '',
    data.scenario4_changed ? 'TRUE' : 'FALSE',
    // Q14
    q14.conf_1||'', q14.conf_2||'', q14.conf_3||'',
    q14.trust_1||'', q14.trust_2||'', q14.trust_3||'',
    q14.frame_1||'', q14.frame_2||'',
    // Q15
    q15.resp_1||'', q15.resp_2||'', q15.resp_3||'', q15.resp_4||'',
    // Q16
    q16.advice_1||'', q16.advice_2||'', q16.advice_3||''
  ];

  sheet.appendRow(row);
  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(lastRow, 1, 1, row.length);
  if (lastRow % 2 === 0) dataRange.setBackground('#f8fafc');
  dataRange.setHorizontalAlignment('center');
  dataRange.setVerticalAlignment('middle');
}

// 테스트용

function testSubmit() {
  const testData = {
    timestamp: new Date().toISOString(),
    condition: 1,
    consent: 'yes',
    sonaId: ' ',
    gender: 'female',
    age: 22,
    q3: { habit_1:'5', habit_2:'6', habit_3:'7' },
    q4: { noc_1:'3', noc_2:'4', noc_3:'5', noc_4:'3', noc_5:'4', noc_6:'5', noc_7:'3', noc_8:'4' },
    scenario1_choice:'A', scenario1_ai_confirmed:true, scenario1_final_choice:'B', scenario1_changed:true,
    scenario2_choice:'B', scenario2_ai_confirmed:true, scenario2_final_choice:'B', scenario2_changed:false,
    scenario3_choice:'A', scenario3_ai_confirmed:true, scenario3_final_choice:'A', scenario3_changed:false,
    scenario4_choice:'B', scenario4_ai_confirmed:true, scenario4_final_choice:'A', scenario4_changed:true,
    q14: { conf_1:'5', conf_2:'6', conf_3:'7', trust_1:'5', trust_2:'6', trust_3:'7', frame_1:'5', frame_2:'6' },
    q15: { resp_1:'5', resp_2:'6', resp_3:'7', resp_4:'5' },
    q16: { advice_1:'5', advice_2:'6', advice_3:'7' }
  };
  const result = submitSurvey(testData);
 Logger.log('결과: ' + JSON.stringify(result));
}
