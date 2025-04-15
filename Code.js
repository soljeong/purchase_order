// 🧩 1. 공통 설정
function getSharedContext() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    ss,
    orderSheet: ss.getSheetByName('발주초안'),
    requestSheet: ss.getSheetByName('구매의뢰'),
    orderListSheet: ss.getSheetByName('발주목록'),
    orderForm : ss.getSheetByName("발주서")
  };
}

function confirmOrder() {
  const { orderSheet, orderListSheet, ss } = getSharedContext();

  const timestamp = new Date();
  const orderNumber = 'PO-' + Utilities.formatDate(timestamp, ss.getSpreadsheetTimeZone(), 'yyyyMMdd') + orderSheet.getRange('A1').getValue().toString().trim();

  // C2부터 시작된 발주 리스트 읽기
  const startRow = 2;
  const startCol = 3;
  const numCols = 6;
  const dataRange = orderSheet.getRange(startRow, startCol, orderSheet.getLastRow() - 1, numCols);
  const data = dataRange.getValues().filter(row => row.join('') !== ''); // 빈 줄 제외

  if (data.length === 0) {
    SpreadsheetApp.getUi().alert('발주할 항목이 없습니다.');
    return;
  }

  // 발주번호를 각 행에 추가하여 전송
  const dataWithOrderNumber = data.map(row => [orderNumber, ...row]);
  const targetStartRow = orderListSheet.getLastRow() + 1;
  orderListSheet.getRange(targetStartRow, 1, dataWithOrderNumber.length, dataWithOrderNumber[0].length)
    .setValues(dataWithOrderNumber);

  // 발주번호를 구매의뢰 시트에도 반영
  recordOrderNumberToRequests(data, orderNumber);

  SpreadsheetApp.getUi().alert(`발주가 확정되었습니다.\n발주번호: ${orderNumber}`);
}

function recordOrderNumberToRequests(filteredData, orderNumber) {
  const { requestSheet } = getSharedContext();
  const allData = requestSheet.getDataRange().getValues();
  const headers = allData[0];
  const targetData = allData.slice(1);

  const matchKeyIndex = headers.indexOf('ID');
  const vendorIndex = headers.indexOf('구매처');
  const orderNoIndex = headers.indexOf('발주번호');

  if (matchKeyIndex === -1 || vendorIndex === -1) {
    SpreadsheetApp.getUi().alert('"품목코드" 또는 "구매처" 열이 필요합니다.');
    return;
  }

  // 구매처 기준으로 대상 구매의뢰 항목에 발주번호 입력
  filteredData.forEach(row => {
    const itemCode = row[0]; // 구매의뢰id
    const vendor = row[5];   // 구매처 (예시로 두번째 열)
    targetData.forEach((reqRow, i) => {
      if (reqRow[matchKeyIndex] === itemCode && reqRow[vendorIndex] === vendor) {
        const targetRow = i + 2; // 헤더 포함이므로 +2
        const writeCol = orderNoIndex === -1 ? headers.length + 1 : orderNoIndex + 1;
        requestSheet.getRange(targetRow, writeCol).setValue(orderNumber);
      }
    });
  });
}

function setDropdownFromVendors() {
  const { requestSheet, orderSheet } = getSharedContext();
  const targetCell = orderSheet.getRange('A1');

  // 1. 구매의뢰 시트에서 "구매처" 열의 데이터 가져오기
  const columnIndex = requestSheet.getRange("1:1").getValues()[0].indexOf("구매처") + 1;
  if (columnIndex === 0) {
    SpreadsheetApp.getUi().alert('"구매의뢰" 시트에서 "구매처" 열을 찾을 수 없습니다.');
    return;
  }
  const lastRow = requestSheet.getLastRow();
  const vendorData = requestSheet.getRange(2, columnIndex, lastRow - 1, 1).getValues(); // 헤더 제외

  // 2. 유니크한 값만 필터링
  const vendors = [...new Set(vendorData.flat().filter(v => v !== ''))];

  if (vendors.length === 0) {
    SpreadsheetApp.getUi().alert('구매처 목록이 비어있습니다.');
    return;
  }

  // 3. 유효성 규칙 만들기
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(vendors, true)
    .setAllowInvalid(false)
    .build();

  // 4. M8 셀에 적용
  targetCell.setDataValidation(rule);
  targetCell.setValue(''); // 기존 값 초기화
}

function copyFilteredVendorRows() {
  const { requestSheet, orderSheet } = getSharedContext();
  const vendorValue = orderSheet.getRange('A1').getValue().toString().trim();
  if (!vendorValue) {
    SpreadsheetApp.getUi().alert('구매처가 선택되지 않았습니다.');
    return;
  }

  // "구매의뢰" 시트 데이터 불러오기
  const dataRange = requestSheet.getDataRange();
  const data = dataRange.getValues();
  const headers = data[0];
  const vendorColumnIndex = headers.indexOf('구매처');
  const orderNumberColumnIndex = headers.indexOf('발주번호');


  if (vendorColumnIndex === -1) {
    SpreadsheetApp.getUi().alert('"구매처" 열을 찾을 수 없습니다.');
    return;
  }

  // ✅ 구매처가 일치하고 발주번호가 비어 있는 행만 필터링 (헤더 제외)
  const filteredRows = data.slice(1).filter(row =>
    row[vendorColumnIndex] === vendorValue &&
    (!row[orderNumberColumnIndex] || row[orderNumberColumnIndex].toString().trim() === '')
  );

  if (filteredRows.length === 0) {
    SpreadsheetApp.getUi().alert(`"${vendorValue}"에 해당하는 데이터가 없습니다.`);
    return;
  }
  // 🔸 기존 C2:H 데이터 삭제
  const clearRange = orderSheet.getRange('C2:H');
  clearRange.clearContent();

  // 결과 붙여넣기: 시트의 A2 셀 기준으로
  const startRow = 2;
  const startCol = 3; // C열 = 3
  const outputRange = orderSheet.getRange(startRow, startCol, filteredRows.length, filteredRows[0].length);
  outputRange.setValues(filteredRows);
}

function fillPOForm() {
  const { orderListSheet, orderForm } = getSharedContext();



  // 1. 구매처 값 읽기 (M8)
  const vendorValue = orderForm.getRange("M8").getValue().toString().trim();
  if (!vendorValue) {
    SpreadsheetApp.getUi().alert("M8 셀에 구매처가 입력되지 않았습니다.");
    return;
  }

  // 2. 구매의뢰 시트에서 데이터 불러오기
  const data = orderListSheet.getDataRange().getValues();
  const headers = data[0];

  const vendorIndex = headers.indexOf("구매처");
  const poNumberIndex = headers.indexOf("발주번호");
  // const fieldsToExtract = ["SKU", "품명", "색상", "수량", "단가", "금액"];
  const fieldsToExtract = ["SKU", "수량"];
  const columnIndexes = fieldsToExtract.map(field => headers.indexOf(field));

  if (vendorIndex === -1 || poNumberIndex === -1 || columnIndexes.includes(-1)) {
    SpreadsheetApp.getUi().alert("필수 컬럼이 없습니다. '구매처', '발주번호', 'SKU', '품명' 등 확인해주세요.");
    return;
  }

  // // 3. 조건에 맞는 행 필터링
  // const filteredRows = data.slice(1).filter(row =>
  //   row[vendorIndex] === vendorValue && !row[poNumberIndex] // 발주번호가 비어있는 행
  // ).map(row => columnIndexes.map(i => row[i]));

  // if (filteredRows.length === 0) {
  //   SpreadsheetApp.getUi().alert(`'${vendorValue}'에 해당하며 발주번호가 비어있는 항목이 없습니다.`);
  //   return;
  // }

  // // 4. 기존 데이터 삭제 (C2:H 범위)
  // const startRow = 2;
  // const startCol = 3; // C열
  // const numRowsToClear = orderForm.getLastRow() - 1;
  // orderForm.getRange(startRow, startCol, numRowsToClear, 6).clearContent();

  // // 5. 데이터 입력
  // const outputRange = orderForm.getRange(startRow, startCol, filteredRows.length, 6);
  // outputRange.setValues(filteredRows);

  // SpreadsheetApp.getUi().alert(`${vendorValue} 구매처의 ${filteredRows.length}건 발주 데이터가 입력되었습니다.`);
}



// 📌 4. 커스텀 메뉴 등록
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 발주 자동화')
    .addItem('① 드롭다운 구매처 생성', 'setDropdownFromVendors')
    .addItem('② 선택된 구매처 데이터 복사', 'copyFilteredVendorRows')
    .addItem('③ 발주 확정 및 발주번호 발행', 'confirmOrder')
    .addToUi();
}

function onEdit(e) {
  const { orderSheet } = getSharedContext();

  const sheet = e.range.getSheet();
  const editedCell = e.range;

  // A1 셀을 수정했을 때만 동작 (시트도 '발주서'여야 함)
  if (sheet.getName() === orderSheet.getName() &&
    editedCell.getA1Notation() === 'A1') {
    copyFilteredVendorRows();
  }
}

