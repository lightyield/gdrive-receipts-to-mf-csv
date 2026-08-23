/**
 * gdrive-receipts-to-mf-csv
 * 
 * Gmailに転送された領収書メールをGoogleドライブへ自動保存し、
 * スプレッドシート台帳をベースにリネームとマネーフォワード(MF)用CSVの出力を行います。
 */

// ==========================================
// 設定情報（ご自身の環境に合わせて書き換えてください）
// ==========================================
const FOLDER_ID = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
const TARGET_LABEL = '自動保存_処理待ち';
const SUCCESS_LABEL = '自動保存_完了';

// ==========================================
// 1. Gmailから領収書を取得してスプレッドシートに登録
// ==========================================
function importGmailReceipts() {
  let gmailFolders;
  try {
    gmailFolders = getGmailFolders();
  } catch (e) {
    Logger.log('エラー: Gmailフォルダの取得に失敗しました。' + e.toString());
    return;
  }
  const folder = gmailFolders.queueFolder;
  const targetLabelObj = GmailApp.getUserLabelByName(TARGET_LABEL);
  const successLabelObj = GmailApp.getUserLabelByName(SUCCESS_LABEL);
  
  if (!targetLabelObj || !successLabelObj) {
    Logger.log('エラー: 必要なラベル（' + TARGET_LABEL + ' または ' + SUCCESS_LABEL + '）がGmail側に見つかりません。');
    return;
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const threads = targetLabelObj.getThreads(0, 30);
  
  for (let i = 0; i < threads.length; i++) {
    const messages = threads[i].getMessages();
    
    for (let j = 0; j < messages.length; j++) {
      const msg = messages[j];
      const subject = msg.getSubject();
      const date = msg.getDate();
      const attachments = msg.getAttachments();
      
      const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      const timeStamp = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
      
      if (attachments.length > 0) {
        // 【パターンA】添付ファイル保存
        for (let k = 0; k < attachments.length; k++) {
          const file = attachments[k];
          if (file.getSize() < 5120 && file.getContentType().indexOf('image/') !== -1) {
            continue; // 小さい画像（ロゴ等）は無視
          }
          
          const tempName = `${timeStamp}_${file.getName()}`;
          const newFile = folder.createFile(file.copyBlob()).setName(tempName);
          
          // スプレッドシートへ行追加
          sheet.appendRow([
            'Gmail', // A列: 取込経路
            date,    // B列: 受信日時
            formattedDate, // C列: 取引日付の初期値として受信日を設定
            '', // D列: 勘定科目 (手動入力用)
            '', // E列: 取引先名 (手動入力用)
            '', // F列: 取引金額 (手動入力用)
            '', // G列: メモ (手動入力用)
            '', // H列: 数式により自動表示されるため空白
            newFile.getId(), // I列: ファイルID
            newFile.getUrl() // J列: 領収書リンク
          ]);
          setFilenameFormula(sheet);
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
        const newFile = folder.createFile(pdfBlob);
        
        // 一時ファイルの削除
        DriveApp.getFileById(tempDocId).setTrashed(true);
        
        // スプレッドシートへ行追加
        sheet.appendRow([
          'Gmail', // A列: 取込経路
          date,    // B列: 受信日時
          formattedDate, // C列: 取引日付
          '', // D列: 勘定科目
          '', // E列: 取引先名
          '', // F列: 取引金額
          '', // G列: メモ
          '', // H列: ファイル名（数式により自動表示されるため空白）
          newFile.getId(), // I列: ファイルID
          newFile.getUrl() // J列: 領収書リンク
        ]);
        setFilenameFormula(sheet);
      }
    }
    
    // スレッドのラベル更新
    threads[i].addLabel(successLabelObj);
    threads[i].removeLabel(targetLabelObj);
  }
}

// H列にファイル名組み立て数式を設定するヘルパー関数
function setFilenameFormula(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  // H列(8列目)に数式を挿入: YYYY.MM.DD_勘定科目_取引先名_取引金額円_メモ（メモが空なら末尾にアンダースコアを付けない）
  const formula = `=TEXT(C${lastRow}, "yyyy.MM.dd")&"_"&D${lastRow}&"_"&E${lastRow}&"_"&F${lastRow}&"円"&IF(G${lastRow}<>"", "_"&G${lastRow}, "")`;
  sheet.getRange(lastRow, 8).setFormula(formula);
  
  // F列(6列目)に数値フォーマット（カンマ区切り）を適用
  sheet.getRange(lastRow, 6).setNumberFormat("#,##0");
}

// ==========================================
// 1.5. 手動アップロードされた領収書を取得してスプレッドシートに登録
// ==========================================
function importManualReceipts() {
  let folders;
  try {
    folders = getManualFolders();
  } catch (e) {
    Logger.log('エラー: フォルダの取得に失敗しました。' + e.toString());
    SpreadsheetApp.getUi().alert('エラー', '手動取込用フォルダの取得に失敗しました:\n' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const queueFolder = folders.queueFolder;
  const doneFolder = folders.doneFolder;
  const files = queueFolder.getFiles();
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
      // ファイルを「処理済み」フォルダに移動
      file.moveTo(doneFolder);
      
      // 統一ルール名でのリネーム（手動取り込み時点での確定名とする）
      const memoPart = memo ? `_${memo}` : '';
      const dotDate = formattedDate.replace(/\//g, '.');
      const newName = `${dotDate}_${category}_${vendor}_${amount}円${memoPart}${ext}`;
      file.setName(newName);
      
      // スプレッドシートへ行追加
      // 列構成: 取込経路(A), 受信日時(B), 取引日付(C), 勘定科目(D), 取引先名(E), 取引金額(F), メモ(G), ファイル名(H), ファイルID(I), 領収書リンク(J)
      const now = new Date();
      sheet.appendRow([
        '手動',           // 取込経路
        now,             // 受信日時（取り込み日時）
        formattedDate,   // 取引日付
        category,        // 勘定科目
        vendor,          // 取引先名
        amount,          // 取引金額
        memo,            // メモ
        newName,         // ファイル名 (最初から確定文字列を直接設定)
        file.getId(),    // ファイルID
        file.getUrl()    // 領収書リンク
      ]);
      
      // F列(6列目)に数値フォーマット（カンマ区切り）を適用
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 6).setNumberFormat("#,##0");
      
      importedCount++;
      Logger.log(`取込成功: ${originalName} -> 処理済みへ移動`);
      
    } catch (e) {
      Logger.log(`エラー: ファイル ${originalName} の処理中にエラーが発生しました: ` + e.toString());
    }
  }
  
  let msg = `${importedCount}件の手動領収書を取り込みました。`;
  if (skippedCount > 0) {
    msg += `\n※ 適合しないファイル名の画像等 ${skippedCount}件 をスキップしました（処理待ちフォルダに残されています）。`;
  }
  
  SpreadsheetApp.getUi().alert('手動取込完了', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// 手動取込用のフォルダ（処理待ち・処理済み）を取得するヘルパー関数
function getManualFolders() {
  const manualQueueId = PropertiesService.getScriptProperties().getProperty('MANUAL_QUEUE_FOLDER_ID');
  const manualDoneId = PropertiesService.getScriptProperties().getProperty('MANUAL_DONE_FOLDER_ID');

  let queueFolder = null;
  let doneFolder = null;

  // 1. スクリプトプロパティで個別に設定されている場合はそれを優先
  if (manualQueueId) {
    try {
      queueFolder = DriveApp.getFolderById(manualQueueId);
    } catch (e) {
      Logger.log('警告: MANUAL_QUEUE_FOLDER_ID で指定されたフォルダの取得に失敗しました: ' + e.toString());
    }
  }
  if (manualDoneId) {
    try {
      doneFolder = DriveApp.getFolderById(manualDoneId);
    } catch (e) {
      Logger.log('警告: MANUAL_DONE_FOLDER_ID で指定されたフォルダの取得に失敗しました: ' + e.toString());
    }
  }

  // 2. 設定されていない場合は、FOLDER_ID から相対的に探索
  if (!queueFolder || !doneFolder) {
    if (!FOLDER_ID) {
      throw new Error('FOLDER_ID または MANUAL_QUEUE_FOLDER_ID / MANUAL_DONE_FOLDER_ID が設定されていません。');
    }
    const parentFolder = DriveApp.getFolderById(FOLDER_ID);
    const manualFolders = parentFolder.getFoldersByName('手動');
    if (!manualFolders.hasNext()) {
      throw new Error('親フォルダの下に「手動」フォルダが見つかりません。');
    }
    const manualFolder = manualFolders.next();

    if (!queueFolder) {
      const queueFolders = manualFolder.getFoldersByName('処理待ち');
      if (!queueFolders.hasNext()) {
        throw new Error('「手動」フォルダの下に「処理待ち」フォルダが見つかりません。');
      }
      queueFolder = queueFolders.next();
    }

    if (!doneFolder) {
      const doneFolders = manualFolder.getFoldersByName('処理済み');
      if (!doneFolders.hasNext()) {
        throw new Error('「手動」フォルダの下に「処理済み」フォルダが見つかりません。');
      }
      doneFolder = doneFolders.next();
    }
  }

  return { queueFolder, doneFolder };
}

// Gmail取込用のフォルダ（処理待ち・処理済み）を取得するヘルパー関数
function getGmailFolders() {
  const gmailQueueId = PropertiesService.getScriptProperties().getProperty('GMAIL_QUEUE_FOLDER_ID');
  const gmailDoneId = PropertiesService.getScriptProperties().getProperty('GMAIL_DONE_FOLDER_ID');

  let queueFolder = null;
  let doneFolder = null;

  // 1. スクリプトプロパティで個別に設定されている場合はそれを優先
  if (gmailQueueId) {
    try {
      queueFolder = DriveApp.getFolderById(gmailQueueId);
    } catch (e) {
      Logger.log('警告: GMAIL_QUEUE_FOLDER_ID で指定されたフォルダの取得に失敗しました: ' + e.toString());
    }
  }
  if (gmailDoneId) {
    try {
      doneFolder = DriveApp.getFolderById(gmailDoneId);
    } catch (e) {
      Logger.log('警告: GMAIL_DONE_FOLDER_ID で指定されたフォルダの取得に失敗しました: ' + e.toString());
    }
  }

  // 2. 設定されていない場合は、FOLDER_ID から相対的に探索
  if (!queueFolder || !doneFolder) {
    if (!FOLDER_ID) {
      throw new Error('FOLDER_ID または GMAIL_QUEUE_FOLDER_ID / GMAIL_DONE_FOLDER_ID が設定されていません。');
    }
    const parentFolder = DriveApp.getFolderById(FOLDER_ID);
    const gmailFolders = parentFolder.getFoldersByName('Gmail');
    if (!gmailFolders.hasNext()) {
      throw new Error('親フォルダの下に「Gmail」フォルダが見つかりません。');
    }
    const gmailFolder = gmailFolders.next();

    if (!queueFolder) {
      const queueFolders = gmailFolder.getFoldersByName('処理待ち');
      if (!queueFolders.hasNext()) {
        throw new Error('「Gmail」フォルダの下に「処理待ち」フォルダが見つかりません。');
      }
      queueFolder = queueFolders.next();
    }

    if (!doneFolder) {
      const doneFolders = gmailFolder.getFoldersByName('処理済み');
      if (!doneFolders.hasNext()) {
        throw new Error('「Gmail」フォルダの下に「処理済み」フォルダが見つかりません。');
      }
      doneFolder = doneFolders.next();
    }
  }

  return { queueFolder, doneFolder };
}

// ==========================================
// 2. 「Gmail未処理」データをリネーム ＆ 処理済みへ移動
// ==========================================
function renameGmailReceipts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  let gmailFolders;
  try {
    gmailFolders = getGmailFolders();
  } catch (e) {
    Logger.log('エラー: Gmailフォルダの取得に失敗しました。' + e.toString());
    SpreadsheetApp.getUi().alert('エラー', 'Gmailフォルダの取得に失敗しました:\n' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  let processedCount = 0;
  
  // 2行目からループ
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // H列(8列目)のセルを取得し、数式が入っているか確認する
    const cellRange = sheet.getRange(i + 1, 8);
    const hasFormula = cellRange.getFormula() !== "";
    
    // A列(取込経路)が「Gmail」かつ H列(8列目)に数式が入っている行（＝Gmail未処理の行）を対象とする
    const sourcePath = row[0]; // A列: 取込経路
    if (sourcePath === 'Gmail' && hasFormula) {
      const rowNum = i + 1;
      const rawDate = row[2]; // C列: 取引日付
      const debit = row[3];   // D列: 勘定科目
      const vendor = row[4];  // E列: 取引先名
      const amount = row[5];  // F列: 取引金額
      const fileId = row[8];  // I列: ファイルID
      
      if (!fileId) continue;
      
      // 必須項目のいずれかが空の場合はエラーで止める
      if (!rawDate || !debit || !vendor || amount === '') {
        const errorMsg = `行 ${rowNum}: 必須情報（取引日付、勘定科目、取引先名、取引金額）が不足しています。すべての項目を入力してから再度実行してください。`;
        Logger.log(errorMsg);
        SpreadsheetApp.getUi().alert('入力エラー', errorMsg, SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
      
      try {
        const file = DriveApp.getFileById(fileId);
        
        // ファイルの親フォルダを確認し、Gmail処理待ちに存在する場合のみ処理
        const parents = file.getParents();
        if (parents.hasNext()) {
          const parent = parents.next();
          if (parent.getId() === gmailFolders.queueFolder.getId()) {
            // Gmail/処理済み フォルダに移動
            file.moveTo(gmailFolders.doneFolder);
            
            // リネーム
            const originalName = file.getName();
            const extMatch = originalName.match(/\.[^.]+$/);
            const ext = extMatch ? extMatch[0] : '';
            
            const dateObj = new Date(rawDate);
            const dotDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy.MM.dd');
            
            const memoVal = row[6]; // G列: メモ
            const memoPart = memoVal ? `_${memoVal}` : '';
            const newName = `${dotDate}_${debit}_${vendor}_${amount}円${memoPart}${ext}`;
            
            file.setName(newName);
            
            // G列（ファイル名）に実際のファイル名を上書き設定（数式を上書きして完了フラグとする）
            cellRange.setValue(newName);
            processedCount++;
            Logger.log(`行 ${rowNum}: リネーム完了 -> ${newName}`);
          }
        }
      } catch (e) {
        Logger.log(`行 ${rowNum} の処理中にエラーが発生しました: ` + e.toString());
      }
    }
  }
  
  if (processedCount === 0) {
    SpreadsheetApp.getUi().alert('確認', 'リネーム対象のGmail未処理データがありませんでした。', SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('処理完了', `${processedCount}件のGmail領収書をリネームし、処理済みへ移動しました。`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==========================================
// 2.5. 「リネーム済み」データをマネーフォワードCSVに出力
// ==========================================
function exportMFSheetsCSV() {
  if (!FOLDER_ID) {
    Logger.log('エラー: プロジェクトの設定 ＞ スクリプトのプロパティ に FOLDER_ID が設定されていません。');
    SpreadsheetApp.getUi().alert('エラー', 'スクリプトのプロパティに FOLDER_ID が設定されていません。プロジェクトの設定から設定してください。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  const folder = DriveApp.getFolderById(FOLDER_ID);
  
  const exportRows = [];
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowNum = i + 1;
    
    const cellRange = sheet.getRange(rowNum, 8);
    const hasFormula = cellRange.getFormula() !== "";
    const fileNameVal = row[7]; // H列: ファイル名
    const csvStatus = row[10] || ""; // K列: CSV出力ステータス（11列目）
    
    // H列に値が入っており（＝数式ではない＝リネーム・手動取込完了）、かつK列（CSV出力）が空であるものを対象とする
    if (!hasFormula && fileNameVal && csvStatus === "") {
      const rawDate = row[2]; // C列: 取引日付
      const debit = row[3];   // D列: 勘定科目
      const vendor = row[4];  // E列: 取引先名
      const amount = row[5];  // F列: 取引金額
      
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
        memo: row[6] // G列: メモ
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
    
    // GoogleドライブにCSVを保存
    const csvFile = folder.createFile(sjisBlob);
    
    // 対象行 of K列にCSVファイル名を書き込む
    for (let i = 0; i < exportRows.length; i++) {
      sheet.getRange(exportRows[i].rowNum, 11).setValue(csvFileName);
    }
    
    SpreadsheetApp.getUi().alert(
      '出力完了', 
      `${exportRows.length}件のデータをマネーフォワード用CSVとして出力し、J列（CSV出力状況）を更新しました。\n\nCSVファイルURL:\n${csvFile.getUrl()}`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log('エラー: CSVファイルの保存またはステータス更新に失敗しました: ' + e.toString());
    SpreadsheetApp.getUi().alert('エラー', 'CSV出力中にエラーが発生しました:\n' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==========================================
// 2.7. 「CSV出力済み」レコードを一括削除
// ==========================================
function deleteExportedReceipts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  // ユーザーに確認
  const response = ui.alert(
    '確認',
    'CSV出力済みのレコードを削除しますか？\n（Googleドライブ上の実ファイルは削除されません）',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  let deleteCount = 0;
  
  // スプレッドシートの行削除による行ずれを防ぐため、下からループを回す
  // 2行目 (インデックス 1) までループ
  for (let i = values.length - 1; i >= 1; i--) {
    const rowNum = i + 1;
    const csvStatus = values[i][10] || ""; // K列: CSV出力
    
    if (csvStatus !== "") {
      sheet.deleteRow(rowNum);
      deleteCount++;
    }
  }
  
  if (deleteCount === 0) {
    ui.alert('確認', '削除対象のCSV出力済みレコードはありませんでした。', ui.ButtonSet.OK);
  } else {
    ui.alert('処理完了', `${deleteCount}件のレコードを削除しました。`, ui.ButtonSet.OK);
  }
}

// ==========================================
// 3. スプレッドシートメニューの追加
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('領収書管理')
    .addItem('Gmailから領収書を取込', 'importGmailReceipts')
    .addItem('手動アップロード領収書を取込', 'importManualReceipts')
    .addItem('Gmail未処理データのリネーム', 'renameGmailReceipts')
    .addItem('MFクラウド会計向けCSV出力', 'exportMFSheetsCSV')
    .addItem('CSV出力済みレコードを削除', 'deleteExportedReceipts')
    .addSeparator()
    .addItem('勘定科目プルダウンを設定', 'setupCategoryValidation')
    .addToUi();
}

// ==========================================
// 4. 勘定科目プルダウンの設定
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
  
  // D列（D2以降のデータ入力範囲として、D2:D1000 を設定）
  const range = sheet.getRange("D2:D1000");
  
  // データの入力規則を作成
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categories, true) // リストから選択（ドロップダウン表示）
    .setAllowInvalid(true)                // リスト外の値の手入力も許可する
    .setHelpText('リストから勘定科目を選択するか、直接入力してください。')
    .build();
  
  range.setDataValidation(rule);
  
  // F列（取引金額）の数値フォーマット設定（カンマ区切り）
  sheet.getRange("F2:F1000").setNumberFormat("#,##0");
  
  // ヘッダー（A1:K1）の一括設定・更新
  const headers = [
    ["取込経路", "受信日時", "取引日付", "勘定科目", "取引先名", "取引金額", "メモ", "ファイル名", "ファイルID", "領収書リンク", "CSV出力"]
  ];
  sheet.getRange("A1:K1").setValues(headers);
  sheet.getRange("A1:K1").setHorizontalAlignment("left");
  
  // H列（ファイル名）の古い数式のアップデート（未処理の数式セルが対象）
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const formulas = sheet.getRange(2, 8, lastRow - 1, 1).getFormulas();
    for (let i = 0; i < formulas.length; i++) {
      const rowNum = i + 2;
      // 数式が入っている行のみアップデート
      if (formulas[i][0] !== "") {
        const formula = `=TEXT(C${rowNum}, "yyyy.MM.dd")&"_"&D${rowNum}&"_"&E${rowNum}&"_"&F${rowNum}&"円"&IF(G${rowNum}<>"", "_"&G${rowNum}, "")`;
        sheet.getRange(rowNum, 8).setFormula(formula);
      }
    }
  }
  
  SpreadsheetApp.getUi().alert(
    '設定・修復完了', 
    'A1:K1のヘッダー再設定、D列のプルダウン設定、F列の金額フォーマット（カンマ区切り）の設定、およびH列の未処理ファイル名数式のアップデートが完了しました。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
