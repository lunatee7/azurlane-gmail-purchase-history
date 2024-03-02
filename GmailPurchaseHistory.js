function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Apps Script')
      .addItem('Gmail에서 결제내역 가져오기', 'getEmailsToSpreadsheet')
      .addToUi();
}

function getEmailsToSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet");

  // Query의 before에 해당하는 날짜 계산(오늘 날짜)
  var today = new Date();
  var todayDateEpoch = Math.floor(today.getTime() / 1000); // Unix Time -> second

  // 시트에서 마지막 행 인덱스 찾기
  var lastRow = sheet.getLastRow();
  var rowIdx = null; // 행별 데이터 추가용 idx 변수
  
  // 이메일 검색용 Query 선언
  var searchQuery = null;

  if (sheet.getRange(lastRow, 1).getValue() === "날짜") {
    // 시트가 비어있을 경우(처음 실행)
    searchQuery = "from:googleplay subject:(주문 영수증) 벽람항로";
    rowIdx = 2;
  }
  else  {
    var lastDateCell = sheet.getRange(lastRow, 1);
    var lastDate = lastDateCell.getValue();
    
    // 공백 제거 후 년, 월, 일 변수 초기화
    var year = lastDate.getFullYear();
    var month = lastDate.getMonth() + 1;
    var day = lastDate.getDate();

    var afterDateString = year + '-' + month + '-' + day;
    var afterDate = new Date(afterDateString);
    var afterDateEpoch = Math.floor(afterDate.getTime() / 1000);  // Unix Time -> second

    // Query 설정
    searchQuery = "from:googleplay subject:(주문 영수증) 벽람항로 ";
    searchQuery += "after:" + afterDateEpoch + " ";
    searchQuery += "before:" + todayDateEpoch;

    // 마지막 행의 날짜 값과 동일한 행 찾기
    // 이전에 동일한 날짜가 있을 경우 덮어쓰기 위함임
    for (var row = 2; row <= lastRow; row++) {
      var dateValue = sheet.getRange(row, 1).getValue();
      var spreadsheetDate = new Date(dateValue);  // 스프레드시트에서 가져온 날짜 값을 날짜 객체로 변환
    if (spreadsheetDate.getFullYear() === year &&
        spreadsheetDate.getMonth() + 1 === month &&
        spreadsheetDate.getDate() === day) {
          rowIdx = row;
          break;
      }
    }

    // 날짜 값을 찾지 못한 경우 종료
    if (rowIdx === null) {
      Browser.msgBox('날짜를 찾을 수 없습니다.');
      return;
    }
  }

  var threads = GmailApp.search(searchQuery);

  // Thread 역순으로 순회하며 데이터 가져오기 (오래된 날짜부터)
  for (var i = threads.length - 1; i >= 0; i--) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var body = message.getPlainBody();
      
     // 날짜 데이터 가져오기
      var dateMatch = body.match(/\d{4}\. \d{1,2}\. \d{1,2}\./);
      if (dateMatch !== null) {
        var date = dateMatch[0];
        if (date.endsWith('.')) {
          date = date.slice(0, -1);
        }
        sheet.getRange(rowIdx, 1).setValue(date);
      }

      // 상품명 데이터 가져오기
      var startIndex = body.indexOf("가격");
      var endIndex = body.indexOf("(벽람항로)");
      if (startIndex !== -1 && endIndex !== -1) {
        var product = body.substring(startIndex + 2, endIndex).trim();
        sheet.getRange(rowIdx, 2).setValue(product);
      }

      // 가격 데이터 가져오기
      var startIndex = body.indexOf("합계: ");
      if (startIndex !== -1) {
        var priceString = body.substring(startIndex + 4).trim();
        var price = priceString.match(/\₩[\d,]+/);
        if (price) {
          sheet.getRange(rowIdx, 3).setValue(price[0]);
        }
      }
      if (sheet.getRange(rowIdx, 1).getValue() === '' || 
          sheet.getRange(rowIdx, 2).getValue() === '' || 
          sheet.getRange(rowIdx, 3).getValue() === '')  {
            continue;
          }
      rowIdx++; // 다음 행으로 이동
    }
  }
}