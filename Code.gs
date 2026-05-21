/**
 * portfolio-invoice-automation
 * GAS 毎月自動請求書PDF生成 → Gmail送信ツール
 *
 * 作成者: hariku0212
 */

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    templateId: props.getProperty('INVOICE_TEMPLATE_ID'),
    sheetId: props.getProperty('INVOICE_SHEET_ID'),
    outputFolderId: props.getProperty('OUTPUT_FOLDER_ID'),
    senderName: props.getProperty('SENDER_NAME') || '株式会社サンプル',
    bankInfo: props.getProperty('BANK_INFO') || '〇〇銀行 〇〇支店 普通 1234567 カ）サンプル',
    taxRate: parseFloat(props.getProperty('TAX_RATE') || '0.10')
  };
}

function runMonthlyInvoice() {
  const config = getConfig();
  validateConfig(config);

  const ss = SpreadsheetApp.openById(config.sheetId);
  const sheet = ss.getSheetByName('請求リスト');

  if (!sheet) throw new Error('スプレッドシートに「請求リスト」シートが見つかりません。');

  const rows = getUnsentRows(sheet);
  if (rows.length === 0) { console.log('送付対象なし'); return; }

  const results = { success: 0, failure: 0, errors: [] };

  rows.forEach(({ rowIndex, data }) => {
    try {
      processOneInvoice(config, sheet, rowIndex, data);
      results.success++;
    } catch (err) {
      results.failure++;
      results.errors.push({ row: rowIndex + 1, message: err.message });
    }
  });

  sendSummaryReport(config, results);
}

function processOneInvoice(config, sheet, rowIndex, data) {
  const [customerName, email, serviceName, amount, dueDate] = data;
  const now = new Date();
  const invoiceNumber = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMM') +
                        '-' + String(rowIndex + 1).padStart(3, '0');

  const baseAmount = parseFloat(String(amount).replace(/,/g, '')) || 0;
  const tax = Math.floor(baseAmount * config.taxRate);
  const totalAmount = baseAmount + tax;

  const replacements = {
    '笠頓名鼷': customerName,
    '笠魔笪差': invoiceNumber,
    '笠鯜年月': Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年M月'),
    '笠鯞日': Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年M月d日'),
    '笠鰄期限': formatDueDate(dueDate),
    '笠鰔小名鼷': serviceName,
    '笠鰔额': formatCurrency(baseAmount),
    '笠消費税': formatCurrency(tax),
    '笠合計金額': formatCurrency(totalAmount)
  };

  const docId = createInvoiceDoc(config.templateId, replacements, invoiceNumber, customerName);
  const pdfBlob = convertDocToPdf(docId, config.outputFolderId, invoiceNumber);

  sendInvoiceEmail(email, customerName, invoiceNumber, pdfBlob,
                   config.senderName, config.bankInfo, totalAmount, replacements['笠鰄期限']);

  markAsSent(sheet, rowIndex, now);
  DriveApp.getFileById(docId).setTrashed(true);
}

function createInvoiceDoc(templateId, replacements, invoiceNumber, customerName) {
  const templateFile = DriveApp.getFileById(templateId);
  const newFile = templateFile.makeCopy('請求書_' + invoiceNumber + '_' + customerName);
  const doc = DocumentApp.openById(newFile.getId());
  const body = doc.getBody();

  for (const [placeholder, value] of Object.entries(replacements)) {
    body.replaceText(escapeRegex(placeholder), value);
  }

  doc.saveAndClose();
  return newFile.getId();
}

function convertDocToPdf(docId, outputFolderId, invoiceNumber) {
  const docFile = DriveApp.getFileById(docId);
  const pdfBlob = docFile.getAs('application/pdf');
  pdfBlob.setName('請求書_' + invoiceNumber + '.pdf');

  if (outputFolderId) {
    DriveApp.getFolderById(outputFolderId).createFile(pdfBlob);
  }

  return pdfBlob;
}

function sendInvoiceEmail(email, customerName, invoiceNumber, pdfBlob,
                          senderName, bankInfo, totalAmount, dueDate) {
  const subject = '【請求書】' + invoiceNumber + ' ' + senderName;
  const body = [
    customerName + ' 様', '',
    'お世話になっております。' + senderName + ' でございます。', '',
    '先月分のご請求書を添付にてお送りいたします。',
    'ご査収のほどよろしくお願いいたします。', '',
    '─────────────────',
    '請求番号: ' + invoiceNumber,
    '合計金額: ' + formatCurrency(totalAmount) + '円（税込）',
    '振込期限: ' + dueDate,
    '─────────────────', '',
    '【振込先】', bankInfo, '',
    '※振込手数料はご負担いただきますようお願いいたします。', '',
    '─────────────────', senderName
  ].join('\n');

  GmailApp.sendEmail(email, subject, body, { attachments: [pdfBlob], name: senderName });
}

function sendSummaryReport(config, results) {
  const adminEmail = Session.getActiveUser().getEmail();
  if (!adminEmail) return;

  const lines = ['月次請求書の自動送付が完了しました。', '',
                 '送付成功: ' + results.success + '件',
                 '送付失敗: ' + results.failure + '件'];
  if (results.errors.length > 0) {
    lines.push('', 'エラー詳細:');
    results.errors.forEach(e => lines.push('  行' + e.row + ': ' + e.message));
  }
  MailApp.sendEmail(adminEmail, '[invoice-automation] 月次請求書送付レポート', lines.join('\n'));
}

function getUnsentRows(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const allData = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  return allData
    .map((row, i) => ({ rowIndex: i + 1, data: row }))
    .filter(({ data }) => !data[5] && data[1]);
}

function markAsSent(sheet, rowIndex, sentAt) {
  const sheetRow = rowIndex + 2;
  sheet.getRange(sheetRow, 6).setValue(true);
  sheet.getRange(sheetRow, 7).setValue(
    Utilities.formatDate(sentAt, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
  );
}

function formatCurrency(amount) {
  return Number(amount).toLocaleString('ja-JP');
}

function formatDueDate(dueDate) {
  if (!dueDate) return '月末';
  if (dueDate instanceof Date) return Utilities.formatDate(dueDate, 'Asia/Tokyo', 'yyyy年M月d日');
  return String(dueDate);
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function validateConfig(config) {
  ['templateId', 'sheetId'].forEach(key => {
    if (!config[key]) throw new Error('スクリプトプロパティに ' + key + ' を設定してください。');
  });
}

function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runMonthlyInvoice')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('runMonthlyInvoice').timeBased().onMonthDay(1).atHour(9).create();
  console.log('月次トリガーを設定しました（毎月１日 09:00）');
}

function testRunSingle() {
  const config = getConfig();
  validateConfig(config);

  const ss = SpreadsheetApp.openById(config.sheetId);
  const sheet = ss.getSheetByName('請求リスト');

  if (!sheet || sheet.getLastRow() < 2) {
    console.log('テストデータがありません。');
    return;
  }

  const firstRow = sheet.getRange(2, 1, 1, 7).getValues()[0];
  processOneInvoice(config, sheet, 1, firstRow);
  console.log('テスト処理完了');
}

function createSampleSheet() {
  const config = getConfig();
  const ss = SpreadsheetApp.openById(config.sheetId);
  let sheet = ss.getSheetByName('請求リスト') || ss.insertSheet('請求リスト');

  const headers = ['顧客名', 'メールアドレス', 'サービス名', '金額（税抜）', '振込期限', '送付済みフラグ', '送付日時'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#4285F4').setFontColor('#FFFFFF').setFontWeight('bold');

  sheet.getRange(2, 1, 1, 7).setValues([
    ['株式会社サンプル', 'test@example.com', 'Webサイト保守費', 50000, '当月末', false, '']
  ]);
  sheet.autoResizeColumns(1, headers.length);
  console.log('サンプルシートを作成しました');
}
