/**
 * 파일명: VatReport.gs
 * 기능: 부가세 신고용 '임대현황표' 양식 자동 작성
 * 수정사항: '매매' 건에 대한 잔금일 기준 필터링 및 입주일 기재 로직 추가
 * [최종 수정] 이미지 확인 결과 반영: 용도(Q열, 인덱스 16) 기준 인덱스 및 범위 조정
 */

// 메뉴 연결용 함수
function runVatRentalReport1() {
  return generateRentalStatusReport(1, 6, "1기");
}

function runVatRentalReport2() {
  return generateRentalStatusReport(7, 12, "2기");
}

/**
 * 메인 로직
 */
function generateRentalStatusReport(startMonth, endMonth, periodName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. 연도 파악
  const fileName = ss.getName(); 
  const yearMatch = fileName.match(/\d{4}/);
  const targetYear = yearMatch ? parseInt(yearMatch[0]) : new Date().getFullYear();
  
  // 조회 기간 설정
  const periodStart = new Date(targetYear, startMonth - 1, 1);
  const periodEnd = new Date(targetYear, endMonth, 0, 23, 59, 59);

  // 2. 시트 로드
  const sheetCurrent = ss.getSheetByName("임대 현황표");
  const sheetExit = ss.getSheetByName("임대 현황표(퇴실)"); 
  
  const targetSheetName = "부가세신고양식(임대현황표)";
  let targetSheet = ss.getSheetByName(targetSheetName);
  
  if (!targetSheet) {
    ui.alert(`⚠️ '${targetSheetName}' 시트가 없습니다.`);
    return null;
  }

  // 3. 데이터 읽기
  const lastRowCur = sheetCurrent.getLastRow();
  // ★ [수정 1] 범위 17 (A~Q열까지)로 조정 (이미지 확인 결과 Q열이 마지막)
  const curData = lastRowCur > 1 ? sheetCurrent.getRange(2, 1, lastRowCur - 1, 17).getValues() : [];

  const lastRowExit = sheetExit ? sheetExit.getLastRow() : 0;
  // ★ [수정 2] 범위 18 (A~R열까지)로 조정 (퇴실일 추가로 1칸 밀림)
  const exitData = (sheetExit && lastRowExit > 1) ? sheetExit.getRange(2, 1, lastRowExit - 1, 18).getValues() : [];

  let finalRows = [];

  // ==================================================
  // [Step 1] 현재 임대 현황
  // ==================================================
  for (let i = 0; i < curData.length; i++) {
    const row = curData[i];
    
    let item = {
      isExit: false,
      unit: String(row[0]).trim(),
      type: String(row[1]).trim(),
      bizNo: String(row[2]).trim(),
      area1: row[3],
      area2: row[4],
      deposit: row[5],
      rent: row[6],
      periodStr: String(row[9]).trim(),
      tenant: String(row[11]).trim(),
      resNo: String(row[12]).trim(), // M열(12)은 주민번호
      
      // ★ [수정 3] 용도 인덱스 16 (Q열)으로 확정
      usage: String(row[16]).trim()
    };

    if (isValidItem(item, periodStart, periodEnd, false)) {
      // 매매인 경우 잔금일을 입주일로 설정
      if (item.type.includes("매매")) {
         const balanceDate = parseDateSmart(item.periodStr);
         item.moveInDate = balanceDate ? formatDate(balanceDate) : "";
      } else {
         const dates = parsePeriodString(item.periodStr);
         item.moveInDate = dates.start ? formatDate(dates.start) : "";
      }
      item.exitDate = ""; 
      finalRows.push(item);
    }
  }

  // ==================================================
  // [Step 2] 퇴실자 현황
  // ==================================================
  for (let i = 0; i < exitData.length; i++) {
    const row = exitData[i];
    
    let exitDateRaw = row[0];
    let exitDateObj = parseDateSmart(exitDateRaw);

    // 날짜 없으면 패스
    if (!exitDateObj) continue; 

    // 1. 작년 퇴실 제외
    const fileStartObject = new Date(targetYear, 0, 1); 
    if (exitDateObj < fileStartObject) continue;

    // 2. 기수 범위 밖 제외
    if (exitDateObj < periodStart || exitDateObj > periodEnd) continue;

    let item = {
      isExit: true,
      exitDateObj: exitDateObj,
      
      unit: String(row[1]).trim(),
      type: String(row[2]).trim(),
      bizNo: String(row[3]).trim(),
      area1: row[4],
      area2: row[5],
      deposit: row[6],
      rent: row[7],
      periodStr: String(row[10]).trim(),
      tenant: String(row[12]).trim(),
      resNo: String(row[13]).trim(), // N열(13)은 주민번호 (퇴실시트라 1칸 밀림)
      
      // ★ [수정 4] 용도 인덱스 17 (R열)으로 확정 (퇴실시트라 1칸 밀림)
      usage: String(row[17]).trim()
    };

    if (isValidItem(item, periodStart, periodEnd, true)) {
      if (item.type.includes("매매")) {
         const balanceDate = parseDateSmart(item.periodStr);
         item.moveInDate = balanceDate ? formatDate(balanceDate) : "";
      } else {
         const dates = parsePeriodString(item.periodStr);
         item.moveInDate = dates.start ? formatDate(dates.start) : "";
      }
      item.exitDate = formatDate(exitDateObj);
      finalRows.push(item);
    }
  }

  // ==================================================
  // [Step 3] 정렬 및 출력
  // ==================================================
  finalRows.sort((a, b) => {
    const numA = parseInt(a.unit.replace(/[^0-9]/g, "")) || 0;
    const numB = parseInt(b.unit.replace(/[^0-9]/g, "")) || 0;
    if (numA !== numB) return numA - numB;

    if (a.isExit && !b.isExit) return -1;
    if (!a.isExit && b.isExit) return 1;
    return 0;
  });

  const lastRowTarget = targetSheet.getLastRow();
  if (lastRowTarget >= 3) {
    targetSheet.getRange(3, 1, lastRowTarget - 2, 12).clearContent();
  }

  if (finalRows.length === 0) {
    ui.alert(`조건에 맞는 데이터가 없습니다.\n(기준연도: ${targetYear})`);
    return null;
  }

  const outputValues = finalRows.map(r => {
    // 사업자 없으면 주민번호 사용
    const finalIdNum = r.bizNo ? r.bizNo : r.resNo;

    return [
      r.usage, r.unit, r.area1, r.area2, r.exitDate, r.moveInDate,
      r.tenant, finalIdNum, r.periodStr, r.type, r.deposit, r.rent
    ];
  });

  targetSheet.getRange(3, 1, outputValues.length, 12).setValues(outputValues);
  targetSheet.getRange(3, 11, outputValues.length, 2).setNumberFormat("#,##0");

  // 임대료 납부내역 시트도 생성
  const rentPaymentResult = generateRentPaymentReport(ss, startMonth, endMonth, targetYear);

  // 관리비 납부내역 시트도 생성
  const mgmtPaymentResult = generateMgmtPaymentReport(ss, startMonth, endMonth, targetYear);

  // === 엑셀 파일 생성 (3개 시트만 포함) ===
  const rentSheet = ss.getSheetByName("임대료 납부내역(부가세)");
  const mgmtSheet = ss.getSheetByName("관리비 납부내역(부가세)");

  // 1. 임시 스프레드시트 생성
  const tempSS = SpreadsheetApp.create(`부가세신고자료_${targetYear}년_${periodName}_temp`);

  // 2. 3개 시트를 임시 스프레드시트로 복사 & 이름 설정
  targetSheet.copyTo(tempSS).setName("임대현황표");
  if (rentSheet) rentSheet.copyTo(tempSS).setName("임대료 납부내역");
  if (mgmtSheet) mgmtSheet.copyTo(tempSS).setName("관리비 납부내역");

  // 3. 기본 시트 삭제 (Sheet1 / 시트1)
  const defaultSheet = tempSS.getSheets()[0];
  if (tempSS.getSheets().length > 1 && defaultSheet.getName() !== "임대현황표") {
    tempSS.deleteSheet(defaultSheet);
  }

  // 4. xlsx로 내보내기 → Drive에 저장
  const exportUrl = `https://docs.google.com/spreadsheets/d/${tempSS.getId()}/export?format=xlsx`;
  const blob = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob().setName(`부가세신고자료_${targetYear}년_${periodName}.xlsx`);

  const file = DriveApp.createFile(blob);
  const downloadUrl = `https://drive.google.com/uc?export=download&id=${file.getId()}`;

  // 5. 임시 스프레드시트 & 원본의 부가세 시트 정리
  DriveApp.getFileById(tempSS.getId()).setTrashed(true);
  if (rentSheet) ss.deleteSheet(rentSheet);
  if (mgmtSheet) ss.deleteSheet(mgmtSheet);

  const rentMsg = rentPaymentResult ? `\n임대료 납부내역: ${rentPaymentResult.count}건` : '';
  const mgmtMsg = mgmtPaymentResult ? `\n관리비 납부내역: ${mgmtPaymentResult.count}건` : '';
  ui.alert(`✅ 생성 완료 (${targetYear}년 ${periodName})\n임대현황표: ${outputValues.length}건${rentMsg}${mgmtMsg}`);
  
  return downloadUrl;
}

// ----------------------------------------------------
// Helper Functions
// ----------------------------------------------------
function isValidItem(item, periodStart, periodEnd, isExitRow) {
  if (!item.unit) return false;
  if (item.usage.includes("근린생활")) return false;

  if (item.type.includes("매매")) {
    if (item.periodStr.toLowerCase().includes("sh")) return false;
    const balanceDate = parseDateSmart(item.periodStr);
    if (!balanceDate) return false;
    if (balanceDate < periodStart || balanceDate > periodEnd) {
      return false; 
    }
    return true;
  } else {
    if (!isExitRow) {
      const dates = parsePeriodString(item.periodStr);
      if (dates.start && dates.end) {
        if (dates.end < periodStart || dates.start > periodEnd) return false;
      }
    }
  }
  return true;
}

function parsePeriodString(str) {
  if (!str) return { start: null, end: null };
  const p = str.split("~");
  return p.length < 2 ? { start: null, end: null } : { start: parseDateSmart(p[0]), end: parseDateSmart(p[1]) };
}

function parseDateSmart(val) {
  if (!val) return null;
  if (Object.prototype.toString.call(val) === "[object Date]") {
    if (isNaN(val.getTime())) return null;
    return val;
  }
  const str = String(val).trim();
  const parts = str.replace(/[^0-9.\-]/g, "").split(/[.\-]/);
  
  if (parts.length === 3) {
    let y = parseInt(parts[0]);
    let m = parseInt(parts[1]) - 1;
    let d = parseInt(parts[2]);
    if (y < 100) y += 2000;
    const dateObj = new Date(y, m, d);
    if (!isNaN(dateObj.getTime())) return dateObj;
  }
  const simpleDate = new Date(str);
  if (!isNaN(simpleDate.getTime())) return simpleDate;
  return null;
}

function formatDate(d) {
  if(!d) return "";
  return `${d.getFullYear()}.${String(d.getMonth()+1).padStart(2,"0")}.${String(d.getDate()).padStart(2,"0")}`;
}

// ============================================================
// [추가] 임대료 납부내역 생성 (부가세 신고용)
// ============================================================

/**
 * 임대료 납부내역을 기수(1기/2기)에 맞춰 필터링하여 시트를 생성합니다.
 * 
 * 원본 시트 구조: A=호수, B=임차인, C=납부일(공식), D=유형, E=월세, F~=[금액,납부일] 쌍
 * 
 * 처리 로직:
 * 1. D열이 "전세"인 행 제외
 * 2. A열(호수), B열(임차인)만 유지 / C,D,E열 제거
 * 3. F열부터 [금액, 납부일] 쌍 순회:
 *    - 해당 열에 기수 기간 내(1~6월 or 7~12월) 납부일이 하나라도 있으면 열 쌍 유지
 *    - 유지된 열에서도 기간 외 납부일의 행 데이터는 삭제
 * 4. 임대료 납부내역(퇴실)도 동일 필터 후 하단에 추가
 */
function generateRentPaymentReport(ss, startMonth, endMonth, targetYear) {
  const outputSheetName = "임대료 납부내역(부가세)";

  // 기존 출력 시트 삭제 후 재생성
  let outputSheet = ss.getSheetByName(outputSheetName);
  if (outputSheet) ss.deleteSheet(outputSheet);
  outputSheet = ss.insertSheet(outputSheetName);

  // 기간 범위
  const periodStart = new Date(targetYear, startMonth - 1, 1);
  const periodEnd = new Date(targetYear, endMonth, 0, 23, 59, 59);

  // === 0. 근린생활시설 호수 목록 조회 (임대 현황표 Q열) ===
  const excludeUnits = new Set();
  const rentalSheet = ss.getSheetByName("임대 현황표");
  if (rentalSheet && rentalSheet.getLastRow() > 1) {
    const rentalData = rentalSheet.getRange(2, 1, rentalSheet.getLastRow() - 1, 17).getValues();
    rentalData.forEach(row => {
      const usage = String(row[16] || "").trim(); // Q열 (인덱스 16)
      if (usage.includes("근린생활")) {
        excludeUnits.add(String(row[0]).trim()); // A열 호수
      }
    });
  }
  if (excludeUnits.size > 0) {
    console.log("[부가세] 근린생활시설 제외 호수: " + [...excludeUnits].join(", "));
  }

  // === 1. 임대료 납부내역 시트 읽기 ===
  const rentSheet = ss.getSheetByName("임대료 납부내역");
  if (!rentSheet || rentSheet.getLastRow() < 2) {
    outputSheet.getRange(1, 1).setValue("임대료 납부내역 데이터가 없습니다.");
    return { sheetId: outputSheet.getSheetId(), count: 0 };
  }

  const headers = rentSheet.getRange(1, 1, 1, rentSheet.getLastColumn()).getValues()[0];
  const allData = rentSheet.getRange(2, 1, rentSheet.getLastRow() - 1, rentSheet.getLastColumn()).getValues();

  // ★ 전처리: 모든 날짜 셀을 targetYear 기준 Date 객체로 변환
  normalizePaymentDates_(allData, headers, targetYear);

  // 행 제외 조건: 전세, 근린생활시설, 합계 행
  function shouldExcludeRow_(row) {
    const hosu = String(row[0]).trim();
    if (!hosu || hosu === "") return true;
    if (hosu === "합계" || hosu === "계") return true;
    if (String(row[3]).trim() === "전세") return true;
    if (excludeUnits.has(hosu)) return true;
    return false;
  }

  // === 2. 기수 기간 내 납부일이 있는 열 쌍(인덱스) 찾기 ===
  const validColPairs = [];

  for (let c = 5; c < headers.length; c += 2) {
    const amountIdx = c;
    const dateIdx = c + 1;
    if (dateIdx >= headers.length) break;

    let hasValidDate = false;
    for (let r = 0; r < allData.length; r++) {
      if (shouldExcludeRow_(allData[r])) continue;
      const dateVal = allData[r][dateIdx];
      if (dateVal instanceof Date && dateVal >= periodStart && dateVal <= periodEnd) {
        hasValidDate = true;
        break;
      }
    }

    if (hasValidDate) {
      validColPairs.push({
        amountIdx: amountIdx,
        dateIdx: dateIdx,
        headerName: String(headers[amountIdx]).trim()
      });
    }
  }

  // === 3. 현재 임대료 데이터 필터링 ===
  const resultRows = [];

  for (let r = 0; r < allData.length; r++) {
    const row = allData[r];
    if (shouldExcludeRow_(row)) continue;

    const outputRow = [String(row[0]).trim(), String(row[1]).trim()];

    validColPairs.forEach(pair => {
      const amount = row[pair.amountIdx];
      const dateVal = row[pair.dateIdx];

      if (dateVal instanceof Date && dateVal >= periodStart && dateVal <= periodEnd) {
        outputRow.push(amount || "");
        outputRow.push(formatDate(dateVal));
      } else {
        outputRow.push("");
        outputRow.push("");
      }
    });

    resultRows.push(outputRow);
  }

  // === 4. 퇴실 데이터 추가 (구분선 없이 바로 이어붙임) ===
  const exitSheet = ss.getSheetByName("임대료 납부내역(퇴실)");

  if (exitSheet && exitSheet.getLastRow() >= 2) {
    const exitHeaders = exitSheet.getRange(1, 1, 1, exitSheet.getLastColumn()).getValues()[0];
    const exitData = exitSheet.getRange(2, 1, exitSheet.getLastRow() - 1, exitSheet.getLastColumn()).getValues();

    normalizePaymentDates_(exitData, exitHeaders, targetYear);

    const exitColMap = [];
    validColPairs.forEach(pair => {
      let found = false;
      for (let ec = 5; ec < exitHeaders.length; ec += 2) {
        if (String(exitHeaders[ec]).trim() === pair.headerName) {
          exitColMap.push({ amountIdx: ec, dateIdx: ec + 1 });
          found = true;
          break;
        }
      }
      if (!found) exitColMap.push(null);
    });

    for (let r = 0; r < exitData.length; r++) {
      const row = exitData[r];
      if (shouldExcludeRow_(row)) continue;
      if (!row[0] || String(row[0]).trim() === "") continue;

      const outputRow = [String(row[0]).trim(), String(row[1]).trim()];
      let hasAnyData = false;

      exitColMap.forEach(mapping => {
        if (!mapping) {
          outputRow.push("");
          outputRow.push("");
          return;
        }
        const amount = row[mapping.amountIdx];
        const dateVal = row[mapping.dateIdx];

        if (dateVal instanceof Date && dateVal >= periodStart && dateVal <= periodEnd) {
          outputRow.push(amount || "");
          outputRow.push(formatDate(dateVal));
          hasAnyData = true;
        } else {
          outputRow.push("");
          outputRow.push("");
        }
      });

      if (hasAnyData) {
        resultRows.push(outputRow); // ★ exitRows 대신 resultRows에 바로 추가
      }
    }
  }

  // === 5. 출력 시트에 쓰기 ===
  if (validColPairs.length === 0 && resultRows.length === 0) {
    outputSheet.getRange(1, 1).setValue("해당 기간에 납부 내역이 없습니다.");
    return { sheetId: outputSheet.getSheetId(), count: 0 };
  }

  const headerRow = ["호수", "임차인"];
  validColPairs.forEach(pair => {
    headerRow.push(pair.headerName);
    headerRow.push("");
  });
  outputSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow])
    .setFontWeight("bold")
    .setBackground("#e8f0fe")
    .setHorizontalAlignment("center");

  let currentRow = 2;

  if (resultRows.length > 0) {
    outputSheet.getRange(currentRow, 1, resultRows.length, resultRows[0].length).setValues(resultRows);
    currentRow += resultRows.length;
  }

  // === 6. 합계 행 추가 ===
  const totalRow = ["합계", ""];
  for (let i = 0; i < validColPairs.length; i++) {
    const colIdx = 2 + (i * 2); // resultRows 내 금액 열 인덱스
    let sum = 0;
    resultRows.forEach(row => {
      const val = Number(row[colIdx]);
      if (!isNaN(val)) sum += val;
    });
    totalRow.push(sum > 0 ? sum : "");
    totalRow.push(""); // 납부일 열은 비움
  }

  outputSheet.getRange(currentRow, 1, 1, totalRow.length).setValues([totalRow])
    .setFontWeight("bold")
    .setBackground("#d9ead3");
  currentRow++;

  // 금액 열 서식 (숫자 포맷) — 합계 행 포함
  for (let i = 0; i < validColPairs.length; i++) {
    const col = 3 + (i * 2);
    if (currentRow > 2) {
      outputSheet.getRange(2, col, currentRow - 2, 1).setNumberFormat("#,##0");
    }
  }

  return {
    sheetId: outputSheet.getSheetId(),
    count: resultRows.length
  };
}

// ============================================================
// [추가] 관리비 납부내역 생성 (부가세 신고용)
// 관리비 시트 구조: A=호수, B=임차인, C열~=[금액,납부일] 쌍 (인덱스 2부터)
// ============================================================
function generateMgmtPaymentReport(ss, startMonth, endMonth, targetYear) {
  const outputSheetName = "관리비 납부내역(부가세)";
  const DATA_START_COL = 2; // C열(인덱스 2)부터 [금액,납부일] 시작

  // 기존 출력 시트 삭제 후 재생성
  let outputSheet = ss.getSheetByName(outputSheetName);
  if (outputSheet) ss.deleteSheet(outputSheet);
  outputSheet = ss.insertSheet(outputSheetName);

  // 기간 범위
  const periodStart = new Date(targetYear, startMonth - 1, 1);
  const periodEnd = new Date(targetYear, endMonth, 0, 23, 59, 59);

  // === 0. 제외 호수 목록 조회 (임대 현황표) ===
  // 제외 조건 1: 근린생활시설 (Q열 용도)
  // 제외 조건 2: 오피스텔(Q열) + 사업자등록번호(C열) 있는 호수
  const excludeUnits = new Set();
  const rentalSheet = ss.getSheetByName("임대 현황표");
  if (rentalSheet && rentalSheet.getLastRow() > 1) {
    const rentalData = rentalSheet.getRange(2, 1, rentalSheet.getLastRow() - 1, 17).getValues();
    rentalData.forEach(row => {
      const hosu = String(row[0]).trim();
      const bizNo = String(row[2] || "").trim();  // C열: 사업자등록번호
      const usage = String(row[16] || "").trim();  // Q열: 용도
      if (usage.includes("근린생활")) {
        excludeUnits.add(hosu);
      } else if (usage.includes("오피스텔") && bizNo !== "") {
        excludeUnits.add(hosu);
        console.log(`[관리비] 오피스텔+사업자등록번호 제외: ${hosu} (${bizNo})`);
      }
    });
  }

  // === 1. 관리비 납부내역 시트 읽기 ===
  const mgmtSheet = ss.getSheetByName("관리비 납부내역");
  if (!mgmtSheet || mgmtSheet.getLastRow() < 2) {
    outputSheet.getRange(1, 1).setValue("관리비 납부내역 데이터가 없습니다.");
    return { sheetId: outputSheet.getSheetId(), count: 0 };
  }

  const headers = mgmtSheet.getRange(1, 1, 1, mgmtSheet.getLastColumn()).getValues()[0];
  const allData = mgmtSheet.getRange(2, 1, mgmtSheet.getLastRow() - 1, mgmtSheet.getLastColumn()).getValues();

  // ★ 전처리: 날짜 셀 변환 (인덱스 3부터 짝수 = 납부일 열)
  normalizeMgmtDates_(allData, headers, targetYear, DATA_START_COL);

  // 행 제외 조건: 근린생활시설, 합계 행 (관리비는 전세 필터 없음)
  function shouldExcludeRow_(row) {
    const hosu = String(row[0]).trim();
    if (!hosu || hosu === "") return true;
    if (hosu === "합계" || hosu === "계") return true;
    if (excludeUnits.has(hosu)) return true;
    return false;
  }

  // === 2. 기수 기간 내 납부일이 있는 열 쌍(인덱스) 찾기 ===
  const validColPairs = [];

  for (let c = DATA_START_COL; c < headers.length; c += 2) {
    const amountIdx = c;
    const dateIdx = c + 1;
    if (dateIdx >= headers.length) break;

    let hasValidDate = false;
    for (let r = 0; r < allData.length; r++) {
      if (shouldExcludeRow_(allData[r])) continue;
      const dateVal = allData[r][dateIdx];
      if (dateVal instanceof Date && dateVal >= periodStart && dateVal <= periodEnd) {
        hasValidDate = true;
        break;
      }
    }

    if (hasValidDate) {
      validColPairs.push({
        amountIdx: amountIdx,
        dateIdx: dateIdx,
        headerName: String(headers[amountIdx]).trim()
      });
    }
  }

  // === 3. 현재 관리비 데이터 필터링 ===
  const resultRows = [];

  for (let r = 0; r < allData.length; r++) {
    const row = allData[r];
    if (shouldExcludeRow_(row)) continue;

    const outputRow = [String(row[0]).trim(), String(row[1]).trim()];

    validColPairs.forEach(pair => {
      const amount = row[pair.amountIdx];
      const dateVal = row[pair.dateIdx];

      if (dateVal instanceof Date && dateVal >= periodStart && dateVal <= periodEnd) {
        outputRow.push(amount || "");
        outputRow.push(formatDate(dateVal));
      } else {
        outputRow.push("");
        outputRow.push("");
      }
    });

    resultRows.push(outputRow);
  }

  // === 4. 퇴실 데이터 추가 (구분선 없이 이어붙임) ===
  const exitSheet = ss.getSheetByName("관리비 납부내역(퇴실)");

  if (exitSheet && exitSheet.getLastRow() >= 2) {
    const exitHeaders = exitSheet.getRange(1, 1, 1, exitSheet.getLastColumn()).getValues()[0];
    const exitData = exitSheet.getRange(2, 1, exitSheet.getLastRow() - 1, exitSheet.getLastColumn()).getValues();

    normalizeMgmtDates_(exitData, exitHeaders, targetYear, DATA_START_COL);

    const exitColMap = [];
    validColPairs.forEach(pair => {
      let found = false;
      for (let ec = DATA_START_COL; ec < exitHeaders.length; ec += 2) {
        if (String(exitHeaders[ec]).trim() === pair.headerName) {
          exitColMap.push({ amountIdx: ec, dateIdx: ec + 1 });
          found = true;
          break;
        }
      }
      if (!found) exitColMap.push(null);
    });

    for (let r = 0; r < exitData.length; r++) {
      const row = exitData[r];
      if (shouldExcludeRow_(row)) continue;

      const outputRow = [String(row[0]).trim(), String(row[1]).trim()];
      let hasAnyData = false;

      exitColMap.forEach(mapping => {
        if (!mapping) {
          outputRow.push("");
          outputRow.push("");
          return;
        }
        const amount = row[mapping.amountIdx];
        const dateVal = row[mapping.dateIdx];

        if (dateVal instanceof Date && dateVal >= periodStart && dateVal <= periodEnd) {
          outputRow.push(amount || "");
          outputRow.push(formatDate(dateVal));
          hasAnyData = true;
        } else {
          outputRow.push("");
          outputRow.push("");
        }
      });

      if (hasAnyData) {
        resultRows.push(outputRow);
      }
    }
  }

  // === 5. 출력 시트에 쓰기 ===
  if (validColPairs.length === 0 && resultRows.length === 0) {
    outputSheet.getRange(1, 1).setValue("해당 기간에 납부 내역이 없습니다.");
    return { sheetId: outputSheet.getSheetId(), count: 0 };
  }

  const headerRow = ["호수", "임차인"];
  validColPairs.forEach(pair => {
    headerRow.push(pair.headerName);
    headerRow.push("");
  });
  outputSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow])
    .setFontWeight("bold")
    .setBackground("#e8f0fe")
    .setHorizontalAlignment("center");

  let currentRow = 2;

  if (resultRows.length > 0) {
    outputSheet.getRange(currentRow, 1, resultRows.length, resultRows[0].length).setValues(resultRows);
    currentRow += resultRows.length;
  }

  // === 6. 합계 행 추가 ===
  const totalRow = ["합계", ""];
  for (let i = 0; i < validColPairs.length; i++) {
    const colIdx = 2 + (i * 2);
    let sum = 0;
    resultRows.forEach(row => {
      const val = Number(row[colIdx]);
      if (!isNaN(val)) sum += val;
    });
    totalRow.push(sum > 0 ? sum : "");
    totalRow.push("");
  }

  outputSheet.getRange(currentRow, 1, 1, totalRow.length).setValues([totalRow])
    .setFontWeight("bold")
    .setBackground("#d9ead3");
  currentRow++;

  // 금액 열 서식
  for (let i = 0; i < validColPairs.length; i++) {
    const col = 3 + (i * 2);
    if (currentRow > 2) {
      outputSheet.getRange(2, col, currentRow - 2, 1).setNumberFormat("#,##0");
    }
  }

  return {
    sheetId: outputSheet.getSheetId(),
    count: resultRows.length
  };
}

/**
 * 관리비용 날짜 전처리 (C열부터 시작하므로 납부일 인덱스가 다름)
 * 납부일 열 = DATA_START_COL + 1, +3, +5, ... (홀수 인덱스)
 */
function normalizeMgmtDates_(data, headers, targetYear, dataStartCol) {
  for (let r = 0; r < data.length; r++) {
    for (let c = dataStartCol + 1; c < data[r].length; c += 2) {
      const raw = data[r][c];
      if (!raw && raw !== 0) continue;

      const dates = convertCellToDate_(raw, targetYear);
      if (dates.length > 0) {
        data[r][c] = dates[0];
      }
    }
  }
}

// ============================================================
// ★★★ 핵심: 날짜 전처리 함수 ★★★
// 모든 날짜 셀을 targetYear 기준 Date 객체로 통일 변환
// ============================================================
function normalizePaymentDates_(data, headers, targetYear) {
  // 처음 3행의 디버그 로그 출력
  let debugCount = 0;

  for (let r = 0; r < data.length; r++) {
    for (let c = 6; c < data[r].length; c += 2) {
      const raw = data[r][c];
      if (!raw && raw !== 0) continue;

      // 디버그: 처음 몇 개 셀의 타입/값 로깅
      if (debugCount < 10 && raw !== "") {
        console.log(`[디버그] 행${r+2} 열${c+1}: typeof=${typeof raw}, value=${raw}, isDate=${raw instanceof Date}`);
        debugCount++;
      }

      // 변환 시도
      const dates = convertCellToDate_(raw, targetYear);
      if (dates.length > 0) {
        // 기간 내 첫 번째 유효한 날짜를 저장
        data[r][c] = dates[0];
      }
    }
  }
}

/**
 * 셀 값을 Date 객체 배열로 변환 (모든 형식 대응)
 * ★ Date 객체는 월/일만 추출하여 targetYear로 재설정
 */
function convertCellToDate_(raw, targetYear) {
  if (!raw && raw !== 0) return [];

  // 1. Date 객체 → 월/일 추출 후 targetYear 적용
  if (raw instanceof Date) {
    if (isNaN(raw.getTime())) return [];
    const fixed = new Date(targetYear, raw.getMonth(), raw.getDate());
    return [fixed];
  }

  // 2. 숫자 (8.24 → 8월24일)
  if (typeof raw === "number") {
    const month = Math.floor(raw);
    const str = String(raw);
    const dotIdx = str.indexOf(".");
    if (dotIdx === -1) return []; // 정수는 날짜 아님

    const dayStr = str.substring(dotIdx + 1);
    const day = parseInt(dayStr, 10);

    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return [new Date(targetYear, month - 1, day)];
    }
    return [];
  }

  // 3. 문자열 — 다중 날짜 가능
  const str = String(raw).trim();
  if (str === "") return [];

  const results = [];

  // 쉼표 분리 ("01.20,02.24" → ["01.20", "02.24"])
  const segments = str.split(",").map(s => s.trim()).filter(s => s);

  for (const seg of segments) {
    // '.'과 '/' 동시 → 다중 ("9.11/10.4", "11.24/12.03")
    if (seg.includes(".") && seg.includes("/")) {
      seg.split("/").forEach(sub => {
        const d = parseDateStr_(sub.trim(), targetYear);
        if (d) results.push(d);
      });
    } else {
      const d = parseDateStr_(seg, targetYear);
      if (d) results.push(d);
    }
  }

  return results;
}

/**
 * 단일 문자열 → Date
 * "01.13" → 1/13, "25/12/13" → 2025/12/13, "7.25" → 7/25
 */
function parseDateStr_(str, targetYear) {
  if (!str) return null;
  str = str.trim();
  if (str === "") return null;

  const parts = str.split(/[.\-\/]/);

  // 3파트: 연/월/일 (25/12/13 → 2025.12.13)
  if (parts.length === 3) {
    let y = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10);
    const d = parseInt(parts[2], 10);
    if (y < 100) y += 2000;
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return new Date(y, m - 1, d);
    }
  }

  // 2파트: 월.일 (01.13, 7.25, 11/11)
  if (parts.length === 2) {
    const m = parseInt(parts[0], 10);
    const d = parseInt(parts[1], 10);
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return new Date(targetYear, m - 1, d);
    }
  }

  return null;
}