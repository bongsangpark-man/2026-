/**
 * [트리거 함수] 월 선택 시 OR 데이터(납부내역/원본현황) 수정 시 자동 업데이트
 * -> 트리거 설정: 이벤트 소스(스프레드시트), 이벤트 유형(수정 시)
 */
function autoUpdateRent(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();
  
  // 1. '월별 임대료 납부 현황' 시트에서 '월(A1)'을 변경했을 때 실행
  if (sheetName === "월별 임대료 납부 현황" && range.getA1Notation() === "A1") {
    var selectedMonth = range.getValue();
    updateAndSortDashboard(selectedMonth);
  }

  // 2. '임대료 납부내역' 또는 '임대 현황표' 수정 시 실시간 반영
  if (sheetName === "임대료 납부내역" || sheetName === "임대 현황표") {
    var dashSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("월별 임대료 납부 현황");
    if (dashSheet) {
      var currentMonth = dashSheet.getRange("A1").getValue();
      updateAndSortDashboard(currentMonth);
    }
  }
}

/**
 * [핵심 로직] 동적 열 탐색을 통한 미납 현황 및 실시간 반영
 */
function updateAndSortDashboard(monthStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashSheet = ss.getSheetByName("월별 임대료 납부 현황");
  var dbSheet = ss.getSheetByName("임대료 납부내역");

  if (!dbSheet || !dashSheet) return;

  // 1. DB 헤더 정보를 읽어 선택된 월의 위치 파악 (이월된 데이터 대응)
  var maxColDB = dbSheet.getLastColumn();
  var headers = dbSheet.getRange(1, 1, 1, maxColDB).getValues()[0];
  
  var targetMonthIdx = -1;
  for (var k = 0; k < headers.length; k++) {
    // 드롭다운에서 선택한 월(예: "1월")과 일치하는 헤더 열 탐색
    if (headers[k].toString().trim() === String(monthStr).trim()) {
      targetMonthIdx = k;
      break;
    }
  }

  // 월을 찾지 못하면 종료 (잘못된 선택 등)
  if (targetMonthIdx === -1) return;

  // 2. DB 데이터 가져오기
  var lastRowDB = dbSheet.getLastRow();
  if (lastRowDB < 2) return;

  var dbRange = dbSheet.getRange(2, 1, lastRowDB - 1, maxColDB);
  var dbValues = dbRange.getValues();
  var dbBackgrounds = dbRange.getBackgrounds(); // 배경색 (공실 확인용)
  var dbFontColors = dbRange.getFontColors();   // 글자색 (금액 예외 확인용)

  var processedData = [];

  // 3. 데이터 가공 루프
  for (var i = 0; i < dbValues.length; i++) {
    var row = dbValues[i];
    var rowBackgrounds = dbBackgrounds[i];
    var rowFontColors = dbFontColors[i];
    
    var unit = row[0];        // A열: 호수
    var name = row[1];        // B열: 이름
    var payDay = row[2];      // C열: 납부일
    var rentTypeRaw = row[3]; // D열: 임대유형
    var stdAmount = row[4];   // E열: 기준 월세
    
    var dbRowIndex = i + 2;
    if (rentTypeRaw === "전세" || name === "공실") continue;
    var rentTypeClean = String(rentTypeRaw).replace(/월세|\(|\)/g, "").trim();

    // [A] 선택 월의 납부유무 수식 (찾은 열의 다음 열인 '날짜' 열 참조)
    var targetDateColIdx = targetMonthIdx + 1;
    var targetColLetter = getColumnLetter(targetDateColIdx + 1);
    var statusFormula = "=IF('임대료 납부내역'!" + targetColLetter + dbRowIndex + "<>\"\",\"○\",\"\")";

    // [B] 미납 현황 계산 (이월된 데이터 포함 과거 모든 열 동적 스캔)
    var unpaidList = [];
    
    // 데이터 시작점인 F열(인덱스 5)부터 선택한 월 이전까지 탐색
    for (var mIdx = 5; mIdx < targetMonthIdx; mIdx += 2) {
      var headerName = headers[mIdx].toString().trim(); // "25.11", "1월" 등 헤더 이름 그대로 사용
      if (!headerName) continue;

      var mAmount = row[mIdx];        // 금액
      var mDateIdx = mIdx + 1;        // 날짜 열 인덱스
      var mDate = row[mDateIdx];      // 날짜 값
      var mDateBgColor = rowBackgrounds[mDateIdx];
      var mAmountFontColor = rowFontColors[mIdx];
      
      // 공실 체크: 날짜칸 배경이 흰색(#ffffff)일 때만 미납으로 간주
      if (mDateBgColor !== "#ffffff" && mDateBgColor !== "white") continue;
      
      // 날짜가 없으면 미납
      if (mDate === "" || mDate == null) {
        unpaidList.push(headerName + " 미납");
      } 
      else {
        // 금액 칸의 글자 색이 검정(#000000)인 경우에만 엄격 비교 (빨간색 등은 정상납부 간주)
        if (mAmountFontColor === "#000000" || mAmountFontColor === "black") {
          if (typeof mAmount === 'number' && typeof stdAmount === 'number') {
             if (mAmount > 0 && mAmount < stdAmount) {
                unpaidList.push(headerName + " 일부미납");
             }
          }
        }
      }
    }

    var unpaidStatus = unpaidList.join(", ");
    processedData.push([unit, name, rentTypeClean, payDay, stdAmount, statusFormula, unpaidStatus]);
  }

  // 4. 정렬: 납부일 오름차순
  processedData.sort(function(a, b) {
    var dayA = parseInt(String(a[3]).replace(/[^0-9]/g, '')) || 99;
    var dayB = parseInt(String(b[3]).replace(/[^0-9]/g, '')) || 99;
    return dayA - dayB; 
  });

  // 5. 결과 입력
  var startRow = 4;
  var lastRowDash = dashSheet.getLastRow();
  if (lastRowDash >= startRow) {
    dashSheet.getRange(startRow, 1, lastRowDash - startRow + 1, 7).clearContent();
  }
  if (processedData.length > 0) {
    dashSheet.getRange(startRow, 1, processedData.length, 7).setValues(processedData);
  }
}

/**
 * [보조 함수] 열 번호를 알파벳으로 변환
 */
function getColumnLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - 1) / 26 | 0;
  }
  return letter;
}