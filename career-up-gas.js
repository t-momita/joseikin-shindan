/**
 * キャリアアップ助成金 申請フォーム - Google Apps Script
 *
 * 【セットアップ手順】
 * 1. https://script.google.com を開く
 * 2. 「新しいプロジェクト」を作成
 * 3. このファイルの内容をすべてコピーして貼り付け
 * 4. 「デプロイ」→「新しいデプロイ」→「種類: ウェブアプリ」
 *    - 実行するユーザー: 自分
 *    - アクセスできるユーザー: 全員（匿名含む）
 * 5. 表示されたURLを career-up.html の GAS_URL に貼り付ける
 */

const SPREADSHEET_NAME = '助成金申請データ_キャリアアップ';
const DRIVE_FOLDER_NAME = '助成金申請書類_キャリアアップ';

const HEADERS = [
  '受付日時',
  '担当者名', '電話番号', 'メールアドレス', '従業員代表者名',
  '対象者氏名', '転換前の働き方', '有期雇用開始日', '正社員転換日', '転換後月給',
  '履歴事項全部証明書（Drive URL）',
  '雇用保険適用事業所設置届（Drive URL）',
  '雇用保険被保険者資格喪失届（Drive URL）',
  '転換前雇用契約書（Drive URL）',
  '賃金台帳（Drive URL）',
  '出勤簿（Drive URL）',
  '就業規則（Drive URL）',
  'キャリアアップ計画書（Drive URL）',
  '転換後雇用契約書（Drive URL）'
];

function doOptions() {
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = getOrCreateSpreadsheet();
    const sheet = getOrCreateSheet(ss);
    const folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
    const subFolder = createSubFolder(folder, data.person);
    const fileUrls = saveFilesToDrive(subFolder, data.files || {});
    const employees = (data.employees && data.employees.length > 0) ? data.employees : [{}];
    employees.forEach(emp => sheet.appendRow(buildRow(data, emp, fileUrls)));
    return jsonResponse({ success: true });
  } catch (err) {
    console.error(err);
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function getOrCreateSpreadsheet() {
  const files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  return SpreadsheetApp.create(SPREADSHEET_NAME);
}

function getOrCreateSheet(ss) {
  let sheet = ss.getSheetByName('キャリアアップ助成金');
  if (!sheet) sheet = ss.insertSheet('キャリアアップ助成金');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    const range = sheet.getRange(1, 1, 1, HEADERS.length);
    range.setBackground('#534AB7').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateFolder(name) {
  const f = DriveApp.getFoldersByName(name);
  return f.hasNext() ? f.next() : DriveApp.createFolder(name);
}

function createSubFolder(parent, label) {
  const date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmm');
  return parent.createFolder((label || '未入力') + '_' + date);
}

function saveFilesToDrive(folder, filesData) {
  const urls = {};
  Object.keys(filesData).forEach(label => {
    const f = filesData[label];
    if (!f || !f.content) return;
    try {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(f.content),
        f.mimeType || 'application/pdf',
        f.name || label + '.pdf'
      );
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urls[label] = file.getUrl();
    } catch (err) {
      urls[label] = 'エラー: ' + err.message;
    }
  });
  return urls;
}

function buildRow(data, emp, fileUrls) {
  return [
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
    data.person     || '',
    data.phone      || '',
    data.email      || '',
    data.workerRep  || '',
    emp.name        || '',
    emp.empType     || '',
    emp.hireDate    || '',
    emp.convDate    || '',
    emp.wageAfter   || '',
    fileUrls['履歴事項全部証明書']         || '',
    fileUrls['雇用保険適用事業所設置届']   || '',
    fileUrls['雇用保険被保険者資格喪失届'] || '',
    fileUrls['転換前雇用契約書']           || '',
    fileUrls['賃金台帳']                   || '',
    fileUrls['出勤簿']                     || '',
    fileUrls['就業規則']                   || '',
    fileUrls['キャリアアップ計画書']       || '',
    fileUrls['転換後雇用契約書']           || ''
  ];
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
