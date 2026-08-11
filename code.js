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
  if (!FOLDER_ID) {
    Logger.log('エラー: プロジェクトの設定 ＞ スクリプトのプロパティ に FOLDER_ID が設定されていません。');
    return;
  }
  const folder = DriveApp.getFolderById(FOLDER_ID);
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
          
          const tempName = `[未処理]_${timeStamp}_${file.getName()}`;
          const newFile = folder.createFile(file.copyBlob()).setName(tempName);
          
          // スプレッドシートへ行追加
          sheet.appendRow([
            date, 
            formattedDate, // 取引日付の初期値として受信日を設定
            '', // 勘定科目 (手動入力用)
            '', // 取引先名 (手動入力用)
            '', // 取引金額 (手動入力用)
            '', // メモ (手動入力用)
            '', // 数式により自動表示されるため空白
            newFile.getId(), 
            newFile.getUrl()
          ]);
          setFilenameFormula(sheet);
        }
      } else {
        // 【パターンB】本文PDF化
        const body = msg.getBody();
        const tempName = `[未処理]_${timeStamp}_本文.pdf`;
        
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
          date, 
          formattedDate,
          '', 
          '', 
          '', 
          '', 
          '', 
          newFile.getId(), 
          newFile.getUrl()
        ]);
        setFilenameFormula(sheet);
      }
    }
    
    // スレッドのラベル更新
    threads[i].addLabel(successLabelObj);
    threads[i].removeLabel(targetLabelObj);
  }
}

// G列にファイル名組み立て数式を設定するヘルパー関数
function setFilenameFormula(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  // G列(7列目)に数式を挿入: YYYY.MM.DD_勘定科目_取引先名_取引金額円_メモ
  const formula = `=TEXT(B${lastRow}, "yyyy.MM.dd")&"_"&C${lastRow}&"_"&D${lastRow}&"_"&E${lastRow}&"円_"&F${lastRow}`;
  sheet.getRange(lastRow, 7).setFormula(formula);
}

// ==========================================
// 2. 「確定」データをリネーム ＆ マネーフォワードCSV出力
// ==========================================
function processConfirmedReceipts() {
  if (!FOLDER_ID) {
    Logger.log('エラー: プロジェクトの設定 ＞ スクリプトのプロパティ に FOLDER_ID が設定されていません。');
    SpreadsheetApp.getUi().alert('エラー', 'スクリプトのプロパティに FOLDER_ID が設定されていません。プロジェクトの設定から設定してください。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  const folder = DriveApp.getFolderById(FOLDER_ID);
  
  // 1. 未処理レコードのバリデーションチェック（1行でも未入力があれば停止）
  const pendingRows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    
    // G列(7列目)のセルを取得し、数式が入っているか確認する
    const cellRange = sheet.getRange(i + 1, 7);
    const hasFormula = cellRange.getFormula() !== "";
    
    // 数式が入っている行（＝未処理の行）を対象とする
    if (hasFormula) {
      const rowNum = i + 1;
      const rawDate = row[1]; // B列: 取引日付
      const debit = row[2];   // C列: 勘定科目
      const vendor = row[3];  // D列: 取引先名
      const amount = row[4];  // E列: 取引金額
      const fileId = row[7];  // H列: ファイルID
      
      // まだファイルすら取り込まれていない行（fileIdが無い等）はスキップ
      if (!fileId) continue;
      
      // 必須項目のいずれかが空の場合はエラーで止める
      if (!rawDate || !debit || !vendor || amount === '') {
        const errorMsg = `行 ${rowNum}: 必須情報（取引日付、勘定科目、取引先名、取引金額）が不足しています。すべての項目を入力してから再度実行してください。`;
        Logger.log(errorMsg);
        SpreadsheetApp.getUi().alert('入力エラー', errorMsg, SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
      
      pendingRows.push({
        rowNum: rowNum,
        rawDate: rawDate,
        debit: debit,
        vendor: vendor,
        amount: amount,
        memo: row[5],   // F列: メモ
        fileId: fileId
      });
    }
  }
  
  if (pendingRows.length === 0) {
    SpreadsheetApp.getUi().alert('確認', '未処理（未リネーム）のデータがありませんでした。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // 2. マネーフォワード用CSVデータの準備
  const csvRows = [
    ["取引No", "取引日", "借方勘定科目", "借方補助科目", "借方部門", "借方取引先", "借方税区分", "借方インボイス", "借方金額(円)", "借方税額", "貸方勘定科目", "貸方補助科目", "貸方部門", "貸方取引先", "貸方税区分", "貸方インボイス", "貸方金額(円)", "貸方税額", "摘要", "仕訳メモ", "タグ", "MF仕訳タイプ", "決算整理仕訳", "作成日時", "作成者", "最終更新日時", "最終更新者"]
  ];
  
  let processedCount = 0;
  let csvTransactionNo = 1;
  
  // 3. リネーム＆CSVデータ生成処理
  for (let i = 0; i < pendingRows.length; i++) {
    const item = pendingRows[i];
    
    // 取引日付のフォーマット (YYYY/MM/DD)
    const dateObj = new Date(item.rawDate);
    const formattedDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    const dotDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy.MM.dd');
    
    try {
      // -----------------------------
      // ドライブ上のファイルリネーム
      // -----------------------------
      const file = DriveApp.getFileById(item.fileId);
      const originalName = file.getName();
      
      // 拡張子の取得
      const extMatch = originalName.match(/\.[^.]+$/);
      const ext = extMatch ? extMatch[0] : '';
      
      // 新しいファイル名の作成: YYYY.MM.DD_勘定科目_取引先名_金額円_メモ.拡張子
      const memoPart = item.memo ? `_${item.memo}` : '';
      const newName = `${dotDate}_${item.debit}_${item.vendor}_${item.amount}円${memoPart}${ext}`;
      
      file.setName(newName);
      
      // -----------------------------
      // マネーフォワード用CSV行の生成
      // -----------------------------
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
      
      // G列（ファイル名）に実際のファイル名を上書き設定（数式を上書きして完了フラグとする）
      sheet.getRange(item.rowNum, 7).setValue(newName);
      processedCount++;
      
    } catch (e) {
      Logger.log(`行 ${item.rowNum} のファイルリネームに失敗しました: ` + e.toString());
    }
  }
  
  // -----------------------------
  // CSVファイルの保存
  // -----------------------------
  if (processedCount > 0) {
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
    
    // Shift_JISに変換（マネーフォワード取り込み時の日本語文字化け防止）
    const blob = Utilities.newBlob(csvContent, 'text/csv', `mf_journal_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss')}.csv`);
    const sjisBlob = blob.getAs('text/csv').setDataFromString(csvContent, 'Shift_JIS');
    
    // GoogleドライブにCSVを保存
    const csvFile = folder.createFile(sjisBlob);
    
    SpreadsheetApp.getUi().alert(
      '処理完了', 
      `${processedCount}件の領収書をリネームし、マネーフォワード用CSVを作成しました。\n\nCSVファイルURL:\n${csvFile.getUrl()}`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ==========================================
// 3. スプレッドシートメニューの追加
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('領収書管理')
    .addItem('Gmailから領収書を取込', 'importGmailReceipts')
    .addItem('未処理データのリネーム＆CSV出力', 'processConfirmedReceipts')
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
  
  // C列（C2以降のデータ入力範囲として、C2:C1000 を設定）
  const range = sheet.getRange("C2:C1000");
  
  // データの入力規則を作成
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categories, true) // リストから選択（ドロップダウン表示）
    .setAllowInvalid(true)                // リスト外の値の手入力も許可する
    .setHelpText('リストから勘定科目を選択するか、直接入力してください。')
    .build();
  
  range.setDataValidation(rule);
  
  SpreadsheetApp.getUi().alert(
    '設定完了', 
    'C列に勘定科目のプルダウンを設定しました。\nリストにない値も直接手入力が可能です。', 
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
