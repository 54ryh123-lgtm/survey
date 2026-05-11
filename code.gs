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

// HTML 웹앱 서빙
function doGet() {
  return HtmlService.createHtmlOutputFromFile(HTML_FILE)
    .setTitle('연구 설문조사')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// HTML에서 google.script.run.submitSurvey(data) 로 직접 호출
function submitSurvey(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME); createHeaders(sheet); }
  if (sheet.getLastRow() === 0) createHeaders(sheet);
  appendDataRow(sheet, data);
  return { result: 'success', row: sheet.getLastRow() };
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

    const counts = [0,0,0,0,0,0,0,0]; // 조건 1~8
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // B열(조건번호)
      for (let i = 0; i < values.length; i++) {
        const c = parseInt(values[i][0], 10);
        if (c >= 1 && c <= 8) counts[c-1]++;
      }
    }
    const minCount = Math.min.apply(null, counts);
    const candidates = [];
    for (let i = 0; i < 8; i++) if (counts[i] === minCount) candidates.push(i + 1);
    return candidates[Math.floor(Math.random() * candidates.length)];
  } finally {
    lock.releaseLock();
  }
}

// 외부 fetch POST 백업 — HTML Service가 아니라 별도 webapp으로 배포 시 사용
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = submitSurvey(data);
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
    // Q3 (7점, 3문항) — AI 활용습관
    'Q3_1','Q3_2','Q3_3',
    // Q4 (5점, 8문항) — NFC
    'Q4_1','Q4_2','Q4_3','Q4_4','Q4_5','Q4_6','Q4_7','Q4_8',
    // 시나리오 1
    '시나리오1_초기선택','시나리오1_AI확인','시나리오1_최종선택','시나리오1_선택변경여부',
    // 시나리오 2
    '시나리오2_초기선택','시나리오2_AI확인','시나리오2_최종선택','시나리오2_선택변경여부',
    // 시나리오 3
    '시나리오3_초기선택','시나리오3_AI확인','시나리오3_최종선택','시나리오3_선택변경여부',
    // 시나리오 4
    '시나리오4_초기선택','시나리오4_AI확인','시나리오4_최종선택','시나리오4_선택변경여부',
    // Q14 (7점, 8문항)
    'Q14_1','Q14_2','Q14_3','Q14_4','Q14_5','Q14_6','Q14_7','Q14_8',
    // Q15 (7점, 4문항)
    'Q15_1','Q15_2','Q15_3','Q15_4',
    // Q16 (7점, 3문항)
    'Q16_1','Q16_2','Q16_3'
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
    data.timestamp || new Date().toISOString(),
    data.condition || '',
    // 기본정보
    data.consent || '',
    data.sonaId  || '',
    data.gender  || '',
    data.age     || '',
    // Q3
    q3.q1 || '', q3.q2 || '', q3.q3 || '',
    // Q4 (8)
    q4.q1 || '', q4.q2 || '', q4.q3 || '', q4.q4 || '',
    q4.q5 || '', q4.q6 || '', q4.q7 || '', q4.q8 || '',
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
    q14.q1||'', q14.q2||'', q14.q3||'', q14.q4||'',
    q14.q5||'', q14.q6||'', q14.q7||'', q14.q8||'',
    // Q15
    q15.q1||'', q15.q2||'', q15.q3||'', q15.q4||'',
    // Q16
    q16.q1||'', q16.q2||'', q16.q3||''
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
    sonaId: '123456',
    gender: 'female',
    age: 22,
    q3: { q1:'5', q2:'6', q3:'7' },
    q4: { q1:'3', q2:'4', q3:'5', q4:'3', q5:'4', q6:'5', q7:'3', q8:'4' },
    scenario1_choice:'A', scenario1_ai_confirmed:true, scenario1_final_choice:'B', scenario1_changed:true,
    scenario2_choice:'B', scenario2_ai_confirmed:true, scenario2_final_choice:'B', scenario2_changed:false,
    scenario3_choice:'A', scenario3_ai_confirmed:true, scenario3_final_choice:'A', scenario3_changed:false,
    scenario4_choice:'B', scenario4_ai_confirmed:true, scenario4_final_choice:'A', scenario4_changed:true,
    q14: { q1:'5', q2:'6', q3:'7', q4:'5', q5:'6', q6:'7', q7:'5', q8:'6' },
    q15: { q1:'5', q2:'6', q3:'7', q4:'5' },
    q16: { q1:'5', q2:'6', q3:'7' }
  };
  submitSurvey(testData);
  Logger.log('테스트 데이터 추가 완료');
}