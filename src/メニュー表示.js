function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menu = [{ name: '請求書送信', functionName: 'sendInvoicesInFolder' }];
  ss.addMenu('請求書送信', menu);
}
