function sendInvoicesInFolder() {
  // Browser.msgBox('msgBox');

  try {
    // Customerシートから請求書情報を取得する
    const customerInfo = getCustomer();

    customerInfo.forEach((customer) => {
      // 対象の請求書を取得する
      const invoices = getInvoices(customer);

      invoices.forEach((invoice) => {
        // 請求先情報に紐づく請求書をメールで送付する
        const result = sendMail(customer, invoice);

        // Emailシートに送った履歴を記載する
        writeHitory(customer, invoice, result);

        // メール送付が成功した請求書を送信済みフォルダに移動する
        if (result) {
          moveInvoice(invoice);
        }
      });
    });
  } catch (e) {
    Browser.msgBox(e);
  }
}

const getCustomer = () => {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customer');
  const lastRow = sheet.getLastRow();
  if (lastRow === 1)
    throw new Error('Customerシートに請求先情報がありません。');

  const usedRange = sheet.getDataRange().getValues().slice(1);
  const arr = usedRange.map((row) => ({
    invoiceCd: row[0],
    invoiceName: row[1],
    addressTo: row[2],
    addressCC: row[3],
    addressBCC: row[4],
  }));

  return arr;
};

const getInvoices = (customer) => {
  const folderId =
    PropertiesService.getScriptProperties().getProperty('awaitingSendFolder');
  const files = DriveApp.getFolderById(folderId).getFiles();

  const regex = /_([\w\d]{10})\.pdf$/;
  let invoiceFiles = [];
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName(); // ファイル名
    const match = fileName.match(regex);
    const invoiceCd = match ? match[1] : null;
    if (!invoiceCd || invoiceCd !== customer.invoiceCd) continue;

    const fileId = file.getId(); // ファイルID
    const fileURL = file.getUrl(); // ファイルURL

    invoiceFiles.push({
      fileName,
      fileId,
      fileURL,
      invoiceCd,
    });
  }

  return invoiceFiles;
};

const sendMail = (customer, invoice) => {
  const to = customer.addressTo;
  const subject = 'ご請求書の送付（千鳥饅頭総本舗）';
  const body = `
${customer.invoiceName}　御中

お世話になっております。
千鳥饅頭総本舗でございます。

当月締分のご請求書をお送りします。
不明点等ございましたら、お気軽にお問い合わせください。

何卒よろしくお願い申し上げます。

株式会社千鳥饅頭総本舗
経理部`;

  const options = {
    cc: customer.addressCC,
    bcc: customer.addressBCC,
    attachments: DriveApp.getFileById(invoice.fileId).getBlob(),
  };

  try {
    MailApp.sendEmail(to, subject, body, options);
    return true;
  } catch (e) {
    return false;
  }
};

const writeHitory = (customer, invoice, result) => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email');
  let rowIndex = sheet.getLastRow() + 1;
  let range = sheet.getRange(rowIndex, 1, 1, 9);
  range.setValues([
    [
      new Date(),
      customer.invoiceCd,
      customer.invoiceName,
      customer.addressTo,
      customer.addressCC,
      customer.addressBCC,
      invoice.fileURL,
      invoice.fileId,
      result ? '成功' : '失敗',
    ],
  ]);
  SpreadsheetApp.flush();
};

const moveInvoice = (invoice) => {
  const file = DriveApp.getFileById(invoice.fileId);
  const folder = DriveApp.getFolderById(
    PropertiesService.getScriptProperties().getProperty('sentFolder')
  );
  file.moveTo(folder);
};
