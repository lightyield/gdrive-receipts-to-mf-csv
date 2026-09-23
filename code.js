/**
 * gdrive-receipts-to-mf-csv
 * 
 * Gmailに転送された領収書メールおよび手動アップロード領収書をGoogleドライブへ集約し、
 * スプレッドシート台帳をベースにマネーフォワード(MF)用仕訳CSVの出力を行います。
 */

// ==========================================
// 設定情報（ご自身の環境に合わせて書き換えてください）
// ==========================================
const FOLDER_ID = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
const TARGET_LABEL = '自動保存_処理待ち';
const SUCCESS_LABEL = '自動保存_完了';

// ==========================================
// 1. Gmailから領収書を取得して「01.受付」に保存
// ==========================================
function importGmailReceipts() {
  let folders;
  try {
    folders = getReceiptFolders();
  } catch (e) {
    Logger.log('エラー: フォルダの取得に失敗しました。' + e.toString());
    if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet()) {
      SpreadsheetApp.getUi().alert('エラー', 'フォルダの取得に失敗しました:\n' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return;
  }
  const folder = folders.inboxFolder;
  const targetLabelObj = GmailApp.getUserLabelByName(TARGET_LABEL);
  const successLabelObj = GmailApp.getUserLabelByName(SUCCESS_LABEL);
  
  if (!targetLabelObj || !successLabelObj) {
    Logger.log('エラー: 必要なラベル（' + TARGET_LABEL + ' または ' + SUCCESS_LABEL + '）がGmail側に見つかりません。');
    return;
  }
  
  const threads = targetLabelObj.getThreads(0, 30);
  let savedFileCount = 0;
  
  for (let i = 0; i < threads.length; i++) {
    const messages = threads[i].getMessages();
    
    for (let j = 0; j < messages.length; j++) {
      const msg = messages[j];
      const date = msg.getDate();
      const attachments = msg.getAttachments();
      const timeStamp = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
      
      if (attachments.length > 0) {
        // 【パターンA】添付ファイル保存
        for (let k = 0; k < attachments.length; k++) {
          const file = attachments[k];
          if (file.getSize() < 5120 && file.getContentType().indexOf('image/') !== -1) {
            continue; // 小さい画像（ロゴ等）は無視
          }
          
          const tempName = `${timeStamp}_${file.getName()}`;
          folder.createFile(file.copyBlob()).setName(tempName);
          savedFileCount++;
        }
      } else {
        // 【パターンB】本文PDF化
        const body = msg.getBody();
        const tempName = `${timeStamp}.pdf`;
        
        // HTMLボディを一時的にGoogleドキュメントにインポートしてPDF化（文字化け防止とレイアウト維持）
        const htmlBlob = Utilities.newBlob(body, 'text/html', 'temp.html');
        
        // Drive API (v3) を使用して、HTMLをGoogleドキュメント形式に自動変換して作成
        const fileMetadata = {
          name: 'temp_receipt_doc',
          mimeType: 'application/vnd.google-apps.document'
        };
        const tempDoc = Drive.Files.create(fileMetadata, htmlBlob);
        const tempDocId = tempDoc.id;
        
        // GoogleドキュメントからPDFを取得
        const pdfBlob = DriveApp.getFileById(tempDocId).getAs('application/pdf').setName(tempName);
        folder.createFile(pdfBlob);
        
        // 一時ファイルの削除
        DriveApp.getFileById(tempDocId).setTrashed(true);
        savedFileCount++;
      }
    }
    
    // スレッドのラベル更新
    threads[i].addLabel(successLabelObj);
    threads[i].removeLabel(targetLabelObj);
  }
  
  Logger.log(`Gmailから ${savedFileCount} 件のファイルを「01.受付」に保存しました。`);
  try {
    if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet()) {
      SpreadsheetApp.getUi().alert('Gmail取込完了', `${threads.length}件のスレッドから ${savedFileCount}件 のファイルを「01.受付」フォルダに保存しました。`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (e) {
    // トリガー実行等でUIがない場合は無視
  }
}

// ==========================================
// 2. 「02.リネーム済み」の領収書を取得してスプレッドシートに登録 ＆ 「03.インポート済み」へ移動
// ==========================================
function importRenamedReceipts() {
  let folders;
  try {
    folders = getReceiptFolders();
  } catch (e) {
    Logger.log('エラー: フォルダの取得に失敗しました。' + e.toString());
    SpreadsheetApp.getUi().alert('エラー', '領収書フォルダの取得に失敗しました:\n' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const renamedFolder = folders.renamedFolder;
  const importedFolder = folders.importedFolder;
  const files = renamedFolder.getFiles();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  let importedCount = 0;
  let skippedCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const originalName = file.getName();
    
    // 拡張子の取得
    const extMatch = originalName.match(/\.[^.]+$/);
    const ext = extMatch ? extMatch[0] : '';
    const baseName = originalName.substring(0, originalName.length - ext.length);
    
    // ファイル名のパース: 日付_勘定科目_取引先名_金額円_メモ（メモは省略可能）
    // 日付は YYYYMMDD のほか YYYY.MM.DD や YYYY/MM/DD, YYYY-MM-DD にも対応
    // 金額はカンマ区切り（2,731円など）にも対応
    const match = baseName.match(/^(\d{4}[./-]?\d{2}[./-]?\d{2})_([^_]+)_([^_]+)_([\d,]+)円(?:_(.*))?$/);
    
    if (!match) {
      Logger.log(`スキップ: ファイル名形式が不適合です: ${originalName}`);
      skippedCount++;
      continue;
    }
    
    const rawDate = match[1].replace(/[./-]/g, ''); // 記号を除去して8桁の数字にする
    const category = match[2];
    const vendor = match[3];
    const amountStr = match[4].replace(/,/g, '');   // カンマを除去
    const amount = parseInt(amountStr, 10);
    const memo = match[5] || ''; // マッチしない場合は空文字
    
    // 取引日付のフォーマット (YYYYMMDD -> YYYY/MM/DD)
    const formattedDate = rawDate.replace(/(\d{4})(\d{2})(\d{2})/, '$1/$2/$3');
    
    try {
      // ファイルを「03.インポート済み」フォルダに移動
      file.moveTo(importedFolder);
      
      // 統一ルール名でのリネーム（正規化）
      const memoPart = memo ? `_${memo}` : '';
      const dotDate = formattedDate.replace(/\//g, '.');
      const newName = `${dotDate}_${category}_${vendor}_${amount}円${memoPart}${ext}`;
      if (originalName !== newName) {
        file.setName(newName);
      }
      
      // スプレッドシートへ行追加
      // 列構成: 登録日時(A), 取引日付(B), 勘定科目(C), 取引先名(D), 取引金額(E), メモ(F), ファイル名(G), ファイルID(H), 領収書リンク(I), CSV出力(J)
      const now = new Date();
      sheet.appendRow([
        now,             // 登録日時
        formattedDate,   // 取引日付
        category,        // 勘定科目
        vendor,          // 取引先名
        amount,          // 取引金額
        memo,            // メモ
        newName,         // ファイル名
        file.getId(),    // ファイルID
        file.getUrl(),   // 領収書リンク
        ''               // CSV出力
      ]);
      
      // E列(5列目)に数値フォーマット（カンマ区切り）を適用
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 5).setNumberFormat("#,##0");
      
      importedCount++;
      Logger.log(`取込成功: ${originalName} -> 03.インポート済みへ移動`);
      
    } catch (e) {
      Logger.log(`エラー: ファイル ${originalName} の処理中にエラーが発生しました: ` + e.toString());
    }
  }
  
  let msg = `${importedCount}件のリネーム済み領収書を取り込みました。`;
  if (skippedCount > 0) {
    msg += `\n※ 適合しないファイル名の画像等 ${skippedCount}件 をスキップしました（02.リネーム済みフォルダに残されています）。`;
  }
  
  SpreadsheetApp.getUi().alert('取込完了', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// 領収書フォルダ（01.受付・02.リネーム済み・03.インポート済み・04.CSV出力済み）を取得するヘルパー関数
function getReceiptFolders() {
  const inboxId = PropertiesService.getScriptProperties().getProperty('INBOX_FOLDER_ID');
  const renamedId = PropertiesService.getScriptProperties().getProperty('RENAMED_FOLDER_ID');
  const importedId = PropertiesService.getScriptProperties().getProperty('IMPORTED_FOLDER_ID');
  const exportedId = PropertiesService.getScriptProperties().getProperty('EXPORTED_FOLDER_ID');

  let inboxFolder = null;
  let renamedFolder = null;
  let importedFolder = null;
  let exportedFolder = null;

  // 1. スクリプトプロパティで個別に設定されている場合はそれを優先
  if (inboxId) {
    try {
      inboxFolder = DriveApp.getFolderById(inboxId);
    } catch (e) {
      Logger.log('警告: INBOX_FOLDER_ID で指定されたフォルダの取得に失敗しました: ' + e.toString());
    }
  }
  if (renamedId) {
    try {
      renamedFolder = DriveApp.getFolderById(renamedId);
    } catch (e) {
      Logger.log('警告: RENAMED_FOLDER_ID で指定されたフォルダの取得に失敗しました: ' + e.toString());
    }
  }
  if (importedId) {
    try {
      importedFolder = DriveApp.getFolderById(importedId);
    } catch (e) {
      Logger.log('警告: IMPORTED_FOLDER_ID で指定されたフォルダの取得に失敗しました: ' + e.toString());
    }
  }
  if (exportedId) {
    try {
      exportedFolder = DriveApp.getFolderById(exportedId);
    } catch (e) {
      Logger.log('警告: EXPORTED_FOLDER_ID で指定されたフォルダの取得に失敗しました: ' + e.toString());
    }
  }

  // 2. 設定されていない場合は、FOLDER_ID から探索（存在しない場合は自動作成）
  if (!inboxFolder || !renamedFolder || !importedFolder || !exportedFolder) {
    if (!FOLDER_ID) {
      throw new Error('FOLDER_ID または各フォルダIDが設定されていません。');
    }
    const parentFolder = DriveApp.getFolderById(FOLDER_ID);

    if (!inboxFolder) {
      const inboxFolders = parentFolder.getFoldersByName('01.受付');
      inboxFolder = inboxFolders.hasNext() ? inboxFolders.next() : parentFolder.createFolder('01.受付');
    }

    if (!renamedFolder) {
      const renamedFolders = parentFolder.getFoldersByName('02.リネーム済み');
      renamedFolder = renamedFolders.hasNext() ? renamedFolders.next() : parentFolder.createFolder('02.リネーム済み');
    }

    if (!importedFolder) {
      const importedFolders = parentFolder.getFoldersByName('03.インポート済み');
      importedFolder = importedFolders.hasNext() ? importedFolders.next() : parentFolder.createFolder('03.インポート済み');
    }

    if (!exportedFolder) {
      const exportedFolders = parentFolder.getFoldersByName('04.CSV出力済み');
      exportedFolder = exportedFolders.hasNext() ? exportedFolders.next() : parentFolder.createFolder('04.CSV出力済み');
    }
  }

  return { inboxFolder, renamedFolder, importedFolder, exportedFolder };
}

// ==========================================
// 3. 「リネーム済み」データをマネーフォワードCSVに出力 ＆ 「04.CSV出力済み」へ実ファイルを移動
// ==========================================
function exportMFSheetsCSV() {
  if (!FOLDER_ID) {
    Logger.log('エラー: プロジェクトの設定 ＞ スクリプトのプロパティ に FOLDER_ID が設定されていません。');
    SpreadsheetApp.getUi().alert('エラー', 'スクリプトのプロパティに FOLDER_ID が設定されていません。プロジェクトの設定から設定してください。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  let folders;
  try {
    folders = getReceiptFolders();
  } catch (e) {
    Logger.log('エラー: フォルダの取得に失敗しました。' + e.toString());
    SpreadsheetApp.getUi().alert('エラー', 'フォルダの取得に失敗しました:\n' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  const parentFolder = DriveApp.getFolderById(FOLDER_ID);
  const exportedFolder = folders.exportedFolder;
  
  const exportRows = [];
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowNum = i + 1;
    
    const fileNameVal = row[6]; // G列: ファイル名 (7列目)
    const fileId = row[7];      // H列: ファイルID (8列目)
    const csvStatus = row[9] || ""; // J列: CSV出力ステータス (10列目)
    
    // G列にファイル名が入っており、かつJ列（CSV出力）が空である行を対象とする
    if (fileNameVal && csvStatus === "") {
      const rawDate = row[1]; // B列: 取引日付
      const debit = row[2];   // C列: 勘定科目
      const vendor = row[3];  // D列: 取引先名
      const amount = row[4];  // E列: 取引金額
      
      // 必須項目のバリデーション
      if (!rawDate || !debit || !vendor || amount === '') {
        const errorMsg = `行 ${rowNum}: 必須情報（取引日付、勘定科目、取引先名、取引金額）が不足しています。内容を確認してください。`;
        Logger.log(errorMsg);
        SpreadsheetApp.getUi().alert('入力エラー', errorMsg, SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
      
      exportRows.push({
        rowNum: rowNum,
        rawDate: rawDate,
        debit: debit,
        vendor: vendor,
        amount: amount,
        memo: row[5], // F列: メモ
        fileId: fileId
      });
    }
  }
  
  if (exportRows.length === 0) {
    SpreadsheetApp.getUi().alert('確認', 'CSV出力対象のデータ（リネーム済みかつCSV未出力）がありませんでした。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // マネーフォワード用CSVデータの作成
  const csvRows = [
    ["取引No", "取引日", "借方勘定科目", "借方補助科目", "借方部門", "借方取引先", "借方税区分", "借方インボイス", "借方金額(円)", "借方税額", "貸方勘定科目", "貸方補助科目", "貸方部門", "貸方取引先", "貸方税区分", "貸方インボイス", "貸方金額(円)", "貸方税額", "摘要", "仕訳メモ", "タグ", "MF仕訳タイプ", "決算整理仕訳", "作成日時", "作成者", "最終更新日時", "最終更新者"]
  ];
  
  let csvTransactionNo = 1;
  
  for (let i = 0; i < exportRows.length; i++) {
    const item = exportRows[i];
    
    const dateObj = new Date(item.rawDate);
    const formattedDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    
    const csvRow = [
      csvTransactionNo++,  // 取引No
      formattedDate,       // 取引日 (YYYY/MM/DD)
      item.debit,          // 借方勘定科目
      "", "", "", "", "",  // 補助, 部門, 取引先, 税区分, インボイス (空欄)
      item.amount,         // 借方金額
      "",                  // 借方税額
      "未払金",            // 貸方勘定科目 (未払金固定)
      "", "", "", "", "",  // 補助, 部門, 取引先, 税区分, インボイス (空欄)
      item.amount,         // 貸方金額
      "",                  // 貸方税額
      item.vendor,         // 摘要 (取引先名)
      item.memo,           // 仕訳メモ (スプレッドシートのメモ)
      "", "", "", "", "", "", "" // 残り空欄
    ];
    csvRows.push(csvRow);
  }
  
  // CSVデータの文字列化
  const csvContent = csvRows.map(row => 
    row.map(value => {
      let str = String(value);
      if (str.indexOf('"') !== -1) {
        str = str.replace(/"/g, '""');
      }
      if (str.indexOf(',') !== -1 || str.indexOf('\n') !== -1 || str.indexOf('"') !== -1) {
        str = `"${str}"`;
      }
      return str;
    }).join(',')
  ).join('\r\n');
  
  try {
    // CSVファイルのファイル名定義
    const csvFileName = `mf_journal_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}.csv`;
    
    // Shift_JISに変換（マネーフォワード取り込み時の日本語文字化け防止）
    const blob = Utilities.newBlob(csvContent, 'text/csv', csvFileName);
    const sjisBlob = blob.getAs('text/csv').setDataFromString(csvContent, 'Shift_JIS');
    
    // Googleドライブ親フォルダにCSVを保存
    const csvFile = parentFolder.createFile(sjisBlob);
    
    // 対象行のJ列にCSVファイル名を書き込み ＆ 実ファイルを「04.CSV出力済み」フォルダへ移動
    for (let i = 0; i < exportRows.length; i++) {
      const item = exportRows[i];
      sheet.getRange(item.rowNum, 10).setValue(csvFileName);
      
      if (item.fileId) {
        try {
          const file = DriveApp.getFileById(item.fileId);
          file.moveTo(exportedFolder);
        } catch (e) {
          Logger.log(`警告: ファイルID ${item.fileId} の「04.CSV出力済み」フォルダへの移動に失敗しました: ` + e.toString());
        }
      }
    }
    
    SpreadsheetApp.getUi().alert(
      '出力完了', 
      `${exportRows.length}件のデータをマネーフォワード用CSVとして出力し、対象ファイルを「04.CSV出力済み」へ移動しました。\n\nCSVファイルURL:\n${csvFile.getUrl()}`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log('エラー: CSVファイルの保存またはステータス更新に失敗しました: ' + e.toString());
    SpreadsheetApp.getUi().alert('エラー', 'CSV出力中にエラーが発生しました:\n' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==========================================
// 4. 「CSV出力済み」レコードを一括削除（04.CSV出力済みフォルダ内のファイルのみ対象）
// ==========================================
function deleteExportedReceipts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  let folders;
  try {
    folders = getReceiptFolders();
  } catch (e) {
    Logger.log('エラー: フォルダの取得に失敗しました。' + e.toString());
    ui.alert('エラー', 'フォルダの取得に失敗しました:\n' + e.toString(), ui.ButtonSet.OK);
    return;
  }
  
  const exportedFolderId = folders.exportedFolder.getId();
  
  // ユーザーに確認
  const response = ui.alert(
    '確認',
    '「04.CSV出力済み」フォルダに移動完了しているレコードを台帳から削除しますか？\n（Googleドライブ上の実ファイルは削除されません）',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  let deleteCount = 0;
  let skippedCount = 0;
  
  // スプレッドシートの行削除による行ずれを防ぐため、下からループを回す
  // 2行目 (インデックス 1) までループ
  for (let i = values.length - 1; i >= 1; i--) {
    const rowNum = i + 1;
    const fileId = values[i][7];          // H列: ファイルID (8列目)
    const csvStatus = values[i][9] || ""; // J列: CSV出力 (10列目)
    
    if (csvStatus !== "") {
      let isInOutFolder = false;
      if (fileId) {
        try {
          const file = DriveApp.getFileById(fileId);
          const parents = file.getParents();
          while (parents.hasNext()) {
            if (parents.next().getId() === exportedFolderId) {
              isInOutFolder = true;
              break;
            }
          }
        } catch (e) {
          Logger.log(`警告: 行 ${rowNum} のファイル取得に失敗しました: ` + e.toString());
        }
      }
      
      if (isInOutFolder) {
        sheet.deleteRow(rowNum);
        deleteCount++;
      } else {
        Logger.log(`スキップ: 行 ${rowNum} のファイルは「04.CSV出力済み」フォルダに存在しないため削除しませんでした。`);
        skippedCount++;
      }
    }
  }
  
  if (deleteCount === 0) {
    let msg = '削除対象のレコードはありませんでした。';
    if (skippedCount > 0) {
      msg += `\n※ CSV出力済みフラグはあるものの「04.CSV出力済み」フォルダに存在しないレコードが ${skippedCount}件 スキップされました。`;
    }
    ui.alert('確認', msg, ui.ButtonSet.OK);
  } else {
    let msg = `${deleteCount}件のレコードを削除しました。`;
    if (skippedCount > 0) {
      msg += `\n※ 「04.CSV出力済み」フォルダに存在しない ${skippedCount}件 はスキップされました。`;
    }
    ui.alert('処理完了', msg, ui.ButtonSet.OK);
  }
}

// ==========================================
// 5. スプレッドシートメニューの追加
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('領収書管理')
    .addItem('Gmailから領収書を取得', 'importGmailReceipts')
    .addItem('リネーム済み領収書の取込', 'importRenamedReceipts')
    .addItem('MFクラウド会計向けCSV出力', 'exportMFSheetsCSV')
    .addItem('CSV出力済みレコードを削除', 'deleteExportedReceipts')
    .addSeparator()
    .addItem('スプレッドシート初期設定', 'setupCategoryValidation')
    .addToUi();
}

// ==========================================
// 6. 勘定科目プルダウン ＆ 集計シートの初期設定
// ==========================================
function setupCategoryValidation() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // ユーザー指定の勘定科目リスト
  const categories = [
    '接待交際費',
    '備品・消耗品費',
    '旅費交通費',
    '通信費',
    '新聞図書費',
    '車両費',
    '荷造運賃',
    '支払手数料',
    '租税公課'
  ];
  
  // C列（C2以降のデータ入力範囲として、C2:C1000 を設定）
  const range = sheet.getRange("C2:C1000");
  
  // データの入力規則を作成
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categories, true) // リストから選択（ドロップダウン表示）
    .setAllowInvalid(true)                // リスト外の値の手入力も許可する
    .setHelpText('リストから勘定科目を選択するか、直接入力してください。')
    .build();
  
  range.setDataValidation(rule);
  
  // E列（取引金額）の数値フォーマット設定（カンマ区切り）
  sheet.getRange("E2:E1000").setNumberFormat("#,##0");
  
  // ヘッダー（A1:J1）の一括設定・更新
  const headers = [
    ["登録日時", "取引日付", "勘定科目", "取引先名", "取引金額", "メモ", "ファイル名", "ファイルID", "領収書リンク", "CSV出力"]
  ];
  sheet.getRange("A1:J1").setValues(headers);
  sheet.getRange("A1:J1").setHorizontalAlignment("left");
  
  // ==========================================
  // 別シート「集計」の作成・自動設定（勘定科目別集計）
  // ==========================================
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheetName = sheet.getName();
  
  let summarySheet = ss.getSheetByName("集計");
  if (!summarySheet) {
    summarySheet = ss.insertSheet("集計");
  }
  
  summarySheet.clear();
  
  // タイトル
  summarySheet.getRange("A1").setValue("勘定科目別集計").setFontSize(14).setFontWeight("bold");
  
  // ヘッダー
  summarySheet.getRange("A3:C3").setValues([["勘定科目", "件数", "合計金額"]]);
  summarySheet.getRange("A3:C3").setFontWeight("bold").setBackground("#e6f2ff").setHorizontalAlignment("left");
  
  // 各勘定科目の集計行設定
  for (let i = 0; i < categories.length; i++) {
    const cat = categories[i];
    const rowNum = i + 4;
    summarySheet.getRange(`A${rowNum}`).setValue(cat);
    summarySheet.getRange(`B${rowNum}`).setFormula(`=COUNTIF('${dataSheetName}'!C2:C, "${cat}")`).setNumberFormat('#,##0"件"');
    summarySheet.getRange(`C${rowNum}`).setFormula(`=SUMIF('${dataSheetName}'!C2:C, "${cat}", '${dataSheetName}'!E2:E)`).setNumberFormat('#,##0"円"');
  }
  
  // 総合計
  const totalRowNum = categories.length + 4;
  summarySheet.getRange(`A${totalRowNum}`).setValue("総合計");
  summarySheet.getRange(`B${totalRowNum}`).setFormula(`=SUM(B4:B${totalRowNum - 1})`).setNumberFormat('#,##0"件"');
  summarySheet.getRange(`C${totalRowNum}`).setFormula(`=SUM(C4:C${totalRowNum - 1})`).setNumberFormat('#,##0"円"');
  summarySheet.getRange(`A${totalRowNum}:C${totalRowNum}`).setFontWeight("bold");
  
  // 幅調整
  summarySheet.autoResizeColumn(1);
  summarySheet.autoResizeColumn(2);
  summarySheet.autoResizeColumn(3);
  
  SpreadsheetApp.getUi().alert(
    '設定・修復完了', 
    'A1:J1のヘッダー再設定、C列のプルダウン設定、E列の金額フォーマット（カンマ区切り）の設定、および「集計」シート（勘定科目別集計）の作成・更新が完了しました。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
