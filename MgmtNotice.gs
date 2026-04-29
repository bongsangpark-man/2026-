/**
 * MgmtNotice.gs
 * 관리비 고지서 생성기 (Apps Script)
 *
 * 기능: 구글 공유드라이브의 관리비 세부내역 + 스프레드시트 내 고지서 양식 시트를 사용하여
 *       세대별 관리비 고지서 이미지(JPG)를 생성하고 월별 폴더에 저장
 *
 * 원본: 관리비.py (Python/openpyxl/win32com) → Apps Script 변환
 */

// ==========================================
// [설정] 상수 정의
// ==========================================

const MGMT_DATA_ID       = '11yCNh-O-vfLa2GI2ekFVrcYjWt-8H4Ar';  // 관리비_세부내역 (공유드라이브 xlsx)
const OUTPUT_FOLDER_ID   = '1ctcsP5jlk2itBPdEeliqRFJYyMzV_42K';   // 저장 폴더 (루트)
const CLOUD_FUNCTION_URL = 'https://asia-northeast3-mgmt-notice.cloudfunctions.net/pdf-to-jpg';
const NOTICE_ADDRESS     = '서울시 중랑구 면목로 92라길 2';
const TEMPLATE_SHEET_NAME = '고지서_양식';                         // 스프레드시트 내 고지서 템플릿 시트명
const RENTAL_SHEET_NAME   = '임대 현황표';                          // 임대 현황표 시트명 (A열=호수, Q열=용도)

// 관리비 항목 → 고지서 템플릿 셀 위치 매핑
// ※ 항목이 추가/변경되면 이 매핑만 수정하면 됩니다
const ITEM_CELL_MAP = {
  '기타':            'I13',
  '청소비':          'L13',
  '일반관리비':       'I14',
  '승강기안전관리비':  'L14',
  '공동전기료':       'I15',
  '소방안전관리비':    'L15',
  '공동수도료':       'I16',
  '전기안전관리비':    'L16',
  '수선유지비':       'I17',
  '소독비':          'L17',
  '저수조':          'I18',
  '정화조':          'L18',
  '화재보험료':       'I19',
  '주차비':          'L19',
  '승강기보험료':     'I20',
  '장기수선충당금':    'L20',
  '인터넷&TV':       'L21',
  '기타항목':        'L23',
  '미납액':          'K25',
  '미납연체료':       'K26',
};

// 데이터만 직접 입력하는 셀 (수식 셀은 제외 — 양식의 수식이 자동 계산)
// 수식 셀: H10(=E10), E11(=B11), E12(=B12), H11(=B12), B30(=B10),
//          C20(=C14+C16+C18),
//          F14~F19, I21~I23, L22, K27
const DATA_CELLS = [
  'B4',                        // 주소
  'B10', 'B11', 'B12',         // 납입영수증: 전월년월, 호수, 면적
  'E10',                       // 관리비 납입 통지서: 당월년월
  'B14', 'C14',                // 전월 납부액 라벨/값
  'C16', 'C18',                // 미납액/미납연체료 (전월 데이터에서 직접 입력)
  'E28', 'E30',                // 납부기한 라벨/날짜
];

// ==========================================
// [유틸리티] 헬퍼 함수
// ==========================================

/**
 * 시트명 생성 (예: 2026년 3월 → "26.03")
 */
function makeNoticeSheetName_(year, month) {
  const yy = year % 100;
  return `${yy}.${String(month).padStart(2, '0')}`;
}

/**
 * 전월 연/월 계산
 */
function getPrevYearMonth_(year, month) {
  if (month === 1) return { year: year - 1, month: 12 };
  return { year: year, month: month - 1 };
}

/**
 * 하위 폴더 생성/조회
 */
function getOrCreateSubFolder_(parent, folderName) {
  const folders = parent.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(folderName);
}

/**
 * 연도/월별 하위 폴더 생성/조회 (예: "2026년/3월" 폴더)
 */
function getOrCreateYearMonthFolder_(parentId, year, month) {
  const root = DriveApp.getFolderById(parentId);
  const yearFolder = getOrCreateSubFolder_(root, `${year}년`);
  return getOrCreateSubFolder_(yearFolder, `${month}월`);
}

/**
 * 납부기한 계산 (다음달 10일)
 */
function calcDueDate_(currYear, currMonth) {
  if (currMonth === 12) {
    return { year: currYear + 1, month: 1, day: 10 };
  }
  return { year: currYear, month: currMonth + 1, day: 10 };
}

/**
 * Excel 파일(.xlsx)을 Google Sheets로 변환하여 열기
 * 공유드라이브의 .xlsx 파일은 SpreadsheetApp.openById()로 직접 열 수 없으므로
 * Drive API를 통해 Google Sheets 형식으로 변환 후 열기
 */
function openExcelAsSheet_(fileId, tempName) {
  const blob = DriveApp.getFileById(fileId).getBlob();
  const resource = {
    name: tempName || 'temp_converted',
    mimeType: 'application/vnd.google-apps.spreadsheet'
  };
  const convertedFile = Drive.Files.create(resource, blob, {
    supportsAllDrives: true
  });
  return {
    spreadsheet: SpreadsheetApp.openById(convertedFile.id),
    tempFileId: convertedFile.id
  };
}

// ==========================================
// [메인] 관리비 고지서 생성
// ==========================================

function generateMgmtNotice() {
  const ui = SpreadsheetApp.getUi();

  // 1. 당월/전월 계산
  const now = new Date();
  const currYear  = now.getFullYear();
  const currMonth = now.getMonth() + 1;
  const prev = getPrevYearMonth_(currYear, currMonth);

  const currSheetName = makeNoticeSheetName_(currYear, currMonth);
  const prevSheetName = makeNoticeSheetName_(prev.year, prev.month);
  const due = calcDueDate_(currYear, currMonth);

  // 2. 확인 다이얼로그
  const response = ui.alert(
    `📄 ${currMonth}월 관리비 고지서 생성`,
    `당월 시트: ${currSheetName}\n전월 시트: ${prevSheetName}\n납부기한: ${due.year}년 ${String(due.month).padStart(2,'0')}월 ${due.day}일\n\n고지서를 생성하시겠습니까?`,
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  // 임시 변환 파일 ID 추적 (정리용)
  const tempFileIds = [];

  try {
    // 3. 관리비 세부내역 열기 (xlsx → Google Sheets 변환)
    const mgmtResult = openExcelAsSheet_(MGMT_DATA_ID, 'temp_관리비_세부내역');
    const mgmtSS = mgmtResult.spreadsheet;
    tempFileIds.push(mgmtResult.tempFileId);

    const currSheet = mgmtSS.getSheetByName(currSheetName);
    const prevSheet = mgmtSS.getSheetByName(prevSheetName);

    if (!currSheet) {
      ui.alert('오류', `'${currSheetName}' 시트를 찾을 수 없습니다.`, ui.ButtonSet.OK);
      cleanupTempFiles_(tempFileIds);
      return;
    }

    // A1: 건물명
    const buildingName = String(currSheet.getRange('A1').getValue() || '');

    // 4. 당월 데이터 읽기 (2행=헤더, 3행부터 데이터)
    const currData = currSheet.getDataRange().getValues();
    const currHeaders = currData[1]; // 인덱스 1 = 2행
    const currColMap = buildColumnMap_(currHeaders);

    // 5. 전월 데이터 읽기
    let prevData = [];
    let prevColMap = {};
    if (prevSheet) {
      prevData = prevSheet.getDataRange().getValues();
      const prevHeaders = prevData[1];
      prevColMap = buildColumnMap_(prevHeaders);
    }

    // 6. 월별 출력 폴더
    const monthFolder = getOrCreateYearMonthFolder_(OUTPUT_FOLDER_ID, currYear, currMonth);

    // 7. 고지서 템플릿: 현재 스프레드시트 내 '고지서_양식' 시트를 복사하여 사용
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = ss.getSheetByName(TEMPLATE_SHEET_NAME);
    if (!templateSheet) {
      ui.alert('오류', `'${TEMPLATE_SHEET_NAME}' 시트를 찾을 수 없습니다.\n스프레드시트에 고지서 양식 시트를 추가해주세요.`, ui.ButtonSet.OK);
      cleanupTempFiles_(tempFileIds);
      return;
    }
    const ws = templateSheet.copyTo(ss).setName('temp_고지서_작업중');

    // 8. 임대 현황표에서 근린생활시설 호수 파악 (고지서 생성 제외)
    const skipHosuSet = buildSkipHosuSet_(ss, RENTAL_SHEET_NAME);

    // 9. 세대별 루프
    let successCount = 0;
    let skipCount = 0;
    let errorList = [];

    for (let i = 2; i < currData.length; i++) {
      const row = currData[i];
      const hosu = String(row[currColMap['호수']] || '').trim();
      if (!hosu) continue;

      // 근린생활시설은 고지서 생성 제외
      if (skipHosuSet.has(hosu)) {
        console.log(`[호수 ${hosu}] 근린생활시설 → 고지서 생성 제외`);
        skipCount++;
        continue;
      }

      // 당월 항목 수집
      const houseInfo = collectHouseInfo_(row, currColMap, hosu);

      // 전월 정보
      const prevInfo = findPrevInfo_(prevData, prevColMap, hosu);

      try {
        // 템플릿에 데이터 채우기
        fillNoticeSheet_(ws, houseInfo, prevInfo,
          prev.year, prev.month,
          currYear, currMonth,
          due, buildingName
        );
        SpreadsheetApp.flush();

        // JPG 내보내기 (현재 스프레드시트의 임시 시트에서 export)
        const hosuDisplay = hosu.endsWith('호') ? hosu : hosu + '호';
        const filename = `${currMonth}월 고지서_${hosuDisplay}`;
        exportSheetAsImage_(ss.getId(), ws.getSheetId(), monthFolder, filename);

        // 셀 초기화 (다음 세대를 위해)
        clearNoticeSheet_(ws);
        SpreadsheetApp.flush();

        successCount++;

      } catch (e) {
        errorList.push(`${hosu}: ${e.toString()}`);
        // 오류 시에도 셀 초기화
        try { clearNoticeSheet_(ws); SpreadsheetApp.flush(); } catch (ex) {}
      }
    }

    // 9. 임시 시트 및 변환 파일 삭제
    try { ss.deleteSheet(ws); } catch (ex) {}
    cleanupTempFiles_(tempFileIds);

    // 10. 결과 안내
    let resultMsg = `✅ ${successCount}건 생성 완료!\n저장 위치: ${monthFolder.getUrl()}`;
    if (skipCount > 0) {
      resultMsg += `\n🏪 근린생활시설 ${skipCount}건 제외`;
    }
    if (errorList.length > 0) {
      resultMsg += `\n\n⚠️ ${errorList.length}건 오류:\n${errorList.join('\n')}`;
    }

    const htmlOutput = HtmlService.createHtmlOutput(
      `<div style="font-family: sans-serif; padding: 10px; text-align: center;">` +
      `  <h3 style="margin-top: 0; color: #188038;">📄 관리비 고지서 생성 완료</h3>` +
      `  <p>${successCount}건의 고지서가 생성되었습니다.</p>` +
      (skipCount > 0 ? `<p style="color: #666;">🏪 근린생활시설 ${skipCount}건 제외</p>` : '') +
      (errorList.length > 0 ? `<p style="color: red;">⚠️ ${errorList.length}건 오류 발생</p>` : '') +
      `  <div style="margin-top: 20px;">` +
      `    <a href="${monthFolder.getUrl()}" target="_blank" ` +
      `       style="background-color: #1a73e8; color: white; padding: 10px 20px; ` +
      `              text-decoration: none; border-radius: 4px; font-weight: bold; ` +
      `              display: inline-block;">📂 ${currMonth}월 폴더 열기</a>` +
      `  </div>` +
      `</div>`
    ).setWidth(400).setHeight(250);

    ui.showModalDialog(htmlOutput, '관리비 고지서');

  } catch (e) {
    cleanupTempFiles_(tempFileIds);
    ui.alert('오류 발생', e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 임시 변환 파일 일괄 삭제
 */
function cleanupTempFiles_(fileIds) {
  fileIds.forEach(id => {
    try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
  });
}

/**
 * 임대 현황표에서 근린생활시설 호수 목록을 Set으로 반환
 * A열 = 호수, Q열 = 용도
 */
function buildSkipHosuSet_(ss, sheetName) {
  const skipSet = new Set();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    console.log(`⚠️ '${sheetName}' 시트를 찾을 수 없어 용도 필터를 건너뜀니다.`);
    return skipSet;
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {   // 1행부터 (헤더 제외)
    const hosu  = String(data[i][0] || '').trim();   // A열
    const usage = String(data[i][16] || '').trim();  // Q열
    if (hosu && usage === '근린생활시설') {
      skipSet.add(hosu);
    }
  }

  if (skipSet.size > 0) {
    console.log(`근린생활시설 제외 호수: ${[...skipSet].join(', ')}`);
  }
  return skipSet;
}

// ==========================================
// [내부] 데이터 처리 함수
// ==========================================

/**
 * 헤더 배열 → { 열이름: 인덱스 } 매핑 생성
 */
function buildColumnMap_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    if (h) map[String(h).trim()] = i;
  });
  return map;
}

/**
 * 한 세대의 관리비 항목 정보를 행에서 추출
 */
function collectHouseInfo_(row, colMap, hosu) {
  const info = {
    '호수': hosu,
    '면적': colMap['면적'] !== undefined ? (row[colMap['면적']] || 0) : 0,
  };

  // ITEM_CELL_MAP의 모든 항목을 데이터에서 찾기
  for (const itemName of Object.keys(ITEM_CELL_MAP)) {
    if (colMap[itemName] !== undefined) {
      let val = row[colMap[itemName]];
      info[itemName] = (val === '' || val === null || val === undefined) ? 0 : Number(val) || 0;
    }
  }

  return info;
}

/**
 * 전월 데이터에서 해당 호수의 합계/미납 정보 추출
 */
function findPrevInfo_(prevData, prevColMap, hosu) {
  const result = { '합계': 0, '미납액': 0, '연체료': 0 };
  if (!prevData.length || prevColMap['호수'] === undefined) return result;

  const hosuIdx = prevColMap['호수'];
  for (let j = 2; j < prevData.length; j++) {
    if (String(prevData[j][hosuIdx] || '').trim() === hosu) {
      result['합계']   = Number(prevData[j][prevColMap['합계']] || 0) || 0;
      result['미납액'] = Number(prevData[j][prevColMap['미납액']] || 0) || 0;
      result['연체료'] = Number(prevData[j][prevColMap['미납연체료']] || 0) || 0;
      break;
    }
  }
  return result;
}

// ==========================================
// [내부] 템플릿 채우기 / 초기화
// ==========================================

/**
 * 고지서 템플릿 시트에 한 세대의 데이터를 채움
 */
function fillNoticeSheet_(ws, houseInfo, prevInfo,
  prevYear, prevMonth, currYear, currMonth,
  due, buildingName) {

  const ho = houseInfo['호수'];
  const hoDisplay = ho.endsWith('호') ? ho : ho + '호';        // "201" → "201호"
  const prevYM = `${prevYear}년 ${String(prevMonth).padStart(2,'0')}월분`;
  const currYM = `${currYear}년 ${String(currMonth).padStart(2,'0')}월분`;

  // ─── 직접 입력 셀만 기입 (수식 셀은 건드리지 않음) ───
  // B10 → B30(=B10), E10 → H10(=E10), B11 → E11(=B11), B12 → E12/H11(=B12)
  ws.getRange('B10').setValue(prevYM);                          // 전월 년월
  ws.getRange('B11').setValue(`${buildingName} ${hoDisplay}`);       // 호수
  ws.getRange('B12').setValue(`${houseInfo['면적']}㎡`);         // 면적
  ws.getRange('E10').setValue(currYM);                           // 당월 년월

  // ─── 주소 (상단) ───
  ws.getRange('B4').setValue(`${NOTICE_ADDRESS}, ${hoDisplay} (${buildingName})`);

  // ─── 전월 납부 (좌측 납입영수증) ───
  // C20(=C14+C16+C18) 수식이 자동 합산
  const currUnpaid = houseInfo['미납액'] || 0;
  ws.getRange('B14').setValue(`${prevMonth}월분 납부액`);            // 라벨은 항상 입력

  if (currUnpaid > 0) {
    // 당월 미납액이 있으면 납입영수증 금액(C14) 부분은 입력하지 않음
    // (날짜, 호수, 면적, 항목 라벨만 입력 — 위에서 이미 설정됨)
  } else {
    // 일반 처리: 전월 합계에서 미납액·미납연체료를 차감하여 순수 납부액만 입력
    const prevTotal   = prevInfo['합계']   || 0;
    const prevUnpaid  = prevInfo['미납액'] || 0;
    const prevLateFee = prevInfo['연체료'] || 0;
    const prevPayment = prevTotal - prevUnpaid - prevLateFee;
    ws.getRange('C14').setValue(prevPayment || '');

    // 전월 미납액·미납연체료가 있으면 납입영수증에 직접 입력
    if (prevUnpaid > 0)  ws.getRange('C16').setValue(prevUnpaid);
    if (prevLateFee > 0) ws.getRange('C18').setValue(prevLateFee);
  }

  // ─── 당월 관리비 세부내역 항목 (값 셀만) ───
  for (const [itemName, cellAddr] of Object.entries(ITEM_CELL_MAP)) {
    if (itemName in houseInfo) {
      ws.getRange(cellAddr).setValue(houseInfo[itemName]);
    }
  }

  // ─── 납부기한 ───
  ws.getRange('E28').setValue(`${currMonth}월분    납 부 기 한`);
  ws.getRange('E30').setValue(
    `${due.year}년 ${String(due.month).padStart(2,'0')}월 ${due.day}일`
  );
}

/**
 * 템플릿 셀 초기화 (다음 세대를 위해 데이터 클리어)
 */
function clearNoticeSheet_(ws) {
  // 데이터 셀만 클리어 (수식 셀은 건드리지 않음)
  DATA_CELLS.forEach(addr => ws.getRange(addr).clearContent());

  // 관리비 항목 값 셀 클리어 (ITEM_CELL_MAP의 값 셀만)
  Object.values(ITEM_CELL_MAP).forEach(addr => ws.getRange(addr).clearContent());
}

// ==========================================
// [내부] 이미지 내보내기
// ==========================================

/**
 * 시트를 JPG 이미지로 내보내기
 *
 * 방법 1 (우선): Cloud Function + PyMuPDF → 200 DPI (2338×1654) JPG
 * 방법 2 (폴백): Drive 썸네일 API → ~1400px JPG
 * 방법 3 (최종): PDF 파일 그대로 저장
 */
function exportSheetAsImage_(ssId, sheetId, folder, filenameBase) {
  const token = ScriptApp.getOAuthToken();

  // Step 1: PDF 생성
  const pdfUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export`
    + `?format=pdf&gid=${sheetId}&portrait=false&fitw=true&size=A4`
    + `&gridlines=false&top_margin=0.2&bottom_margin=0.2&left_margin=0.2&right_margin=0.2`;

  const pdfResponse = UrlFetchApp.fetch(pdfUrl, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (pdfResponse.getResponseCode() !== 200) {
    throw new Error(`PDF 생성 실패 (HTTP ${pdfResponse.getResponseCode()})`);
  }

  const pdfBlob = pdfResponse.getBlob();

  // Step 2: Cloud Function으로 고해상도 JPG 변환 시도
  if (CLOUD_FUNCTION_URL) {
    try {
      const pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
      const cfResponse = UrlFetchApp.fetch(CLOUD_FUNCTION_URL, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ pdf_base64: pdfBase64 }),
        muteHttpExceptions: true
      });

      if (cfResponse.getResponseCode() === 200) {
        const result = JSON.parse(cfResponse.getContentText());
        if (result.jpg_base64) {
          const jpgBytes = Utilities.base64Decode(result.jpg_base64);
          const jpgBlob = Utilities.newBlob(jpgBytes, 'image/jpeg', filenameBase + '.jpg');
          folder.createFile(jpgBlob);
          console.log(`[${filenameBase}] Cloud Function으로 고해상도 JPG 생성 완료`);
          return;
        }
      }
      console.log(`[${filenameBase}] Cloud Function 실패 (HTTP ${cfResponse.getResponseCode()}) → 썸네일 방식으로 폴백`);
    } catch (e) {
      console.log(`[${filenameBase}] Cloud Function 오류: ${e.toString()} → 썸네일 방식으로 폴백`);
    }
  }

  // Step 3 (폴백): Drive 썸네일 API로 JPG 생성
  const tempPdf = folder.createFile(pdfBlob.setName('temp_' + filenameBase + '.pdf'));
  const tempPdfId = tempPdf.getId();

  for (let attempt = 1; attempt <= 5; attempt++) {
    Utilities.sleep(3000);

    try {
      const meta = Drive.Files.get(tempPdfId, {
        fields: 'hasThumbnail,thumbnailLink',
        supportsAllDrives: true
      });

      if (meta.hasThumbnail && meta.thumbnailLink) {
        const imgUrl = meta.thumbnailLink.replace(/=s\d+/, '=s1400');
        const imgResponse = UrlFetchApp.fetch(imgUrl, {
          muteHttpExceptions: true,
          followRedirects: true
        });

        if (imgResponse.getResponseCode() === 200) {
          const contentType = imgResponse.getHeaders()['Content-Type'] || '';
          if (contentType.includes('image')) {
            folder.createFile(imgResponse.getBlob().setName(filenameBase + '.jpg'));
            tempPdf.setTrashed(true);
            return;
          }
        }
      }
    } catch (e) {
      console.log(`[${filenameBase}] 썸네일 JPG 시도 ${attempt}/5: ${e.toString()}`);
    }
  }

  // Step 4 (최종 폴백): JPG 실패 → PDF 유지
  tempPdf.setName(filenameBase + '.pdf');
  console.log(`[${filenameBase}] JPG 변환 실패 → PDF로 저장`);
}
