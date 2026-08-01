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
            '未処理', 
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
          '未処理', 
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

// H列にファイル名組み立て数式を設定するヘルパー関数
function setFilenameFormula(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  // H列(8列目)に数式を挿入: YYYY.MM.DD_勘定科目_取引先名_取引金額円_メモ.拡張子
  const formula = `=IF(A${lastRow}="完了", "リネーム済", TEXT(C${lastRow}, "yyyy.MM.dd")&"_"&D${lastRow}&"_"&E${lastRow}&"_"&F${lastRow}&"円_"&G${lastRow})`;
  sheet.getRange(lastRow, 8).setFormula(formula);
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
  
  // マネーフォワード用CSVヘッダー
  const csvRows = [
    ["取引No", "取引日", "借方勘定科目", "借方補助科目", "借方部門", "借方取引先", "借方税区分", "借方インボイス", "借方金額(円)", "借方税額", "貸方勘定科目", "貸方補助科目", "貸方部門", "貸方取引先", "貸方税区分", "貸方インボイス", "貸方金額(円)", "貸方税額", "摘要", "仕訳メモ", "タグ", "MF仕訳タイプ", "決算整理仕訳", "作成日時", "作成者", "最終更新日時", "最終更新者"]
  ];
  
  let processedCount = 0;
  let csvTransactionNo = 1;
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const status = row[0]; // A列: ステータス
    
    if (status === '確定') {
      const rowNum = i + 1;
      
      const rawDate = row[2]; // C列: 取引日付
      const debit = row[3];   // D列: 勘定科目
      const vendor = row[4];  // E列: 取引先名
      const amount = row[5];  // F列: 取引金額
      const memo = row[6];    // G列: メモ
      const fileId = row[8];  // I列: ファイルID
      
      if (!rawDate || !debit || !vendor || !amount) {
        Logger.log(`行 ${rowNum}: 必須情報（取引日付、勘定科目、取引先名、取引金額）が不足しているためスキップします。`);
        continue;
      }
      
      // 取引日付のフォーマット (YYYY/MM/DD)
      const dateObj = new Date(rawDate);
      const formattedDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      const dotDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy.MM.dd');
      
      try {
        // -----------------------------
        // ドライブ上のファイルリネーム
        // -----------------------------
        const file = DriveApp.getFileById(fileId);
        const originalName = file.getName();
        
        // 拡張子の取得
        const extMatch = originalName.match(/\.[^.]+$/);
        const ext = extMatch ? extMatch[0] : '';
        
        // 新しいファイル名の作成: YYYY.MM.DD_勘定科目_取引先名_金額円_メモ.拡張子
        const memoPart = memo ? `_${memo}` : '';
        const newName = `${dotDate}_${debit}_${vendor}_${amount}円${memoPart}${ext}`;
        
        file.setName(newName);
        
        // -----------------------------
        // マネーフォワード用CSV行の生成
        // -----------------------------
        const csvRow = [
          csvTransactionNo++,  // 取引No
          formattedDate,       // 取引日 (YYYY/MM/DD)
          debit,               // 借方勘定科目
          "", "", "", "", "",  // 補助, 部門, 取引先, 税区分, インボイス (空欄)
          amount,              // 借方金額
          "",                  // 借方税額
          "未払金",            // 貸方勘定科目 (未払金固定)
          "", "", "", "", "",  // 補助, 部門, 取引先, 税区分, インボイス (空欄)
          amount,              // 貸方金額
          "",                  // 貸方税額
          vendor,              // 摘要 (取引先名)
          memo,                // 仕訳メモ (スプレッドシートのメモ)
          "", "", "", "", "", "", "" // 残り空欄
        ];
        csvRows.push(csvRow);
        
        // ステータスを完了に変更
        sheet.getRange(rowNum, 1).setValue('完了');
        processedCount++;
        
      } catch (e) {
        Logger.log(`行 ${rowNum} のファイルリネームに失敗しました: ` + e.toString());
      }
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
  } else {
    SpreadsheetApp.getUi().alert('確認', 'ステータスが「確定」になっているデータがありませんでした。', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ==========================================
// 3. スプレッドシートメニューの追加
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('領収書管理')
    .addItem('Gmailから領収書を取込', 'importGmailReceipts')
    .addItem('確定データのリネーム＆CSV出力', 'processConfirmedReceipts')
    .addToUi();
}
