/**
 * 파일명: Code.gs
 * 기능: 장부 생성/이월 자동화 (완납 세대 금액+날짜 열 쌍으로 색칠 구분 및 미납 이월)
 */

// ==========================================
// [1] 메뉴 생성
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🏢 임대현황관리')
    .addItem('임대 관리 시스템 열기', 'showRentalSidebar') 
    .addSeparator() 
    .addItem('🔒 시트 잠금 (수정 방지)', 'lockRentalSheet') 
    .addItem('🔓 시트 잠금 해제', 'unlockRentalSheet')     
    .addToUi();

  ui.createMenu('📂 장부만들기')
    .addItem('📅 금년 장부 생성하기(1월 2일 이후 생성)', 'createNextYearSheet')
    .addSeparator()
    .addItem('⚙️ 자동화(트리거) 생성하기', 'setupTriggersForNewYear')
    .addToUi();

  ui.createMenu('📊 부가세 신고자료')
    .addItem('부가세 메뉴 열기', 'showVatSidebar') 
    .addToUi();
}

// ==========================================
// [2] 금년 장부 자동 생성 및 이월 메인 로직
// ==========================================

function createNextYearSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentFileName = ss.getName();
  const currentFileId = ss.getId(); 
  
  const yearMatch = currentFileName.match(/\d{4}/);
  const currentYear = yearMatch ? parseInt(yearMatch[0]) : 2025;
  const nextYear = currentYear + 1;

  const response = ui.alert(
    `📅 ${nextYear}년 장부 생성`, 
    `데이터를 이월하고 새 장부를 생성하시겠습니까?\n(완납 세대는 금액/날짜 열 모두 회색으로 표시됩니다.)`, 
    ui.ButtonSet.YES_NO
  );
  
  if (response != ui.Button.YES) return;

  try {
    // 1. 파일 복제
    const newFileName = currentFileName.replace(String(currentYear), String(nextYear)) + " (새해 장부)";
    const newFile = DriveApp.getFileById(currentFileId).makeCopy(newFileName);
    const newSS = SpreadsheetApp.openById(newFile.getId());
    const newUrl = newSS.getUrl(); 
    
    // 2. 설정값 저장
    const targetSheet = newSS.getSheetByName('임대 현황표');
    if (targetSheet) {
      targetSheet.getRange('A1').setNote(JSON.stringify({ prevId: currentFileId, year: String(nextYear) }));
    }

    // 3. 이월 로직 실행 (추출 -> 삭제 -> 삽입/색칠)
    processRentCarryOver(ss, newSS, currentYear); 
    processMaintCarryOver(ss, newSS, currentYear); 

    // 4. 퇴실 내역 정리
    ['임대료 납부내역(퇴실)', '관리비 납부내역(퇴실)'].forEach(name => {
      const s = newSS.getSheetByName(name);
      if (s && s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
    });

    cleanUpOldTriggers();

    // 5. 결과 안내 모달
    const htmlOutput = HtmlService.createHtmlOutput(
      `<div style="font-family: sans-serif; padding: 10px; text-align: center;">` +
      `  <h3 style="margin-top: 0; color: #188038;">✅ 생성 및 이월 완료!</h3>` +
      `  <p>전년도 완납 세대는 금액/날짜 열이 모두 회색으로 표시되었습니다.</p>` +
      `  <div style="margin-top: 20px;">` +
      `    <a href="${newUrl}" target="_blank" style="background-color: #1a73e8; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px; font-weight: bold; display: inline-block;">🚀 새 장부로 이동하기</a>` +
      `  </div>` +
      `</div>`
    ).setWidth(400).setHeight(300);

    ui.showModalDialog(htmlOutput, '장부 생성 결과');

  } catch (e) {
    ui.alert('오류 발생', e.toString(), ui.ButtonSet.OK);
  }
}

// ==========================================
// [3] 임대료 이월: 금액+날짜 열(쌍) 색칠 로직
// ==========================================

function processRentCarryOver(currentSS, newSS, currentYear) {
  const oldSheet = currentSS.getSheetByName('임대료 납부내역');
  const newSheet = newSS.getSheetByName('임대료 납부내역');
  if (!oldSheet || !newSheet) return;

  const lastRow = oldSheet.getLastRow();
  const headers = oldSheet.getRange(1, 1, 1, oldSheet.getLastColumn()).getValues()[0];
  const targetMonths = [];

  headers.forEach((h, i) => {
    const headStr = h.toString().trim();
    if (i+1 >= 6 && headStr.includes('월') && !headStr.includes('.')) {
      const range = oldSheet.getRange(2, i+1, lastRow - 1, 1);
      const vals = range.getValues();
      const bgs = range.getBackgrounds();
      if (vals.some((v, idx) => v[0].toString().trim() === "" && (bgs[idx][0] === "#ffffff" || bgs[idx][0] === "white"))) {
        targetMonths.push({index: i + 1, month: headStr.replace(/[^0-9]/g, "")});
      }
    }
  });

  if (newSheet.getLastRow() > 1) {
    newSheet.getRange(2, 6, newSheet.getLastRow() - 1, newSheet.getLastColumn() - 5).clearContent().setBackground(null);
  }

  targetMonths.sort((a, b) => a.month - b.month).reverse().forEach(obj => {
    newSheet.insertColumns(6, 2);
    const headerTitle = `${String(currentYear).slice(-2)}.${obj.month}`;
    newSheet.getRange(1, 6, 1, 2).setValues([[headerTitle, ""]]).setBackground("#eeeeee").setHorizontalAlignment("center").setFontWeight("bold");

    const oldData = oldSheet.getRange(2, obj.index, lastRow - 1, 2).getValues();
    const newBgs = [];

    oldData.forEach(row => {
      const date = row[1] ? row[1].toString().trim() : "";
      const color = (date !== "") ? "#d9d9d9" : null;
      newBgs.push([color, color]); // 금액열과 날짜열 세트로 색칠
    });

    newSheet.getRange(2, 6, newBgs.length, 2).setBackgrounds(newBgs);
  });
}

// ==========================================
// [4] 관리비 이월: 12월 금액+날짜 열(쌍) 색칠 및 미납 이월
// ==========================================

function processMaintCarryOver(currentSS, newSS, currentYear) {
  const oldSheet = currentSS.getSheetByName('관리비 납부내역');
  const newSheet = newSS.getSheetByName('관리비 납부내역');
  if (!oldSheet || !newSheet) return;

  const lastRow = oldSheet.getLastRow();
  const headers = oldSheet.getRange(1, 1, 1, oldSheet.getLastColumn()).getValues()[0];
  let decCol = -1;
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toString().trim() === "12월") { decCol = i + 1; break; }
  }
  if (decCol === -1) return;

  const oldData = oldSheet.getRange(2, decCol, lastRow - 1, 2).getValues();
  const carryVals = [];
  const carryBgs = [];

  oldData.forEach(row => {
    const cost = row[0];
    const date = row[1] ? row[1].toString().trim() : "";
    if (date !== "") {
      carryVals.push(["", ""]); // 완납은 금액 비움
      carryBgs.push(["#d9d9d9", "#d9d9d9"]); // 완납은 세트로 회색 색칠
    } else {
      carryVals.push([cost, ""]); // 미납은 금액만 이월
      carryBgs.push([null, null]);
    }
  });

  if (newSheet.getLastRow() > 1) {
    newSheet.getRange(2, 3, newSheet.getLastRow() - 1, newSheet.getLastColumn() - 2).clearContent().setBackground(null);
  }

  newSheet.insertColumns(3, 2);
  const headerTitle = `${String(currentYear).slice(-2)}.12`;
  newSheet.getRange(1, 3, 1, 2).setValues([[headerTitle, ""]]).setBackground("#dfebf7").setHorizontalAlignment("center").setFontWeight("bold");
  
  newSheet.getRange(2, 3, carryVals.length, 2).setValues(carryVals).setBackgrounds(carryBgs);
}

// ==========================================
// [5] 트리거 및 유틸리티
// ==========================================

function cleanUpOldTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction() === 'sendExtensionCheckEmails') {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }
}

function setupTriggersForNewYear() {
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  try {
    ScriptApp.newTrigger('sendExtensionCheckEmails').timeBased().atHour(9).everyDays(1).create();
    ScriptApp.newTrigger('autoUpdateRent').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
    ui.alert('✅ 자동화 설정 완료');
  } catch (e) {
    ui.alert('설정 실패: ' + e.toString());
  }
}

function showVatSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('VatSidebar').setTitle('📊 부가세 신고 관리').setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}