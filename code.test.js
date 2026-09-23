const fs = require('fs');
const path = require('path');

// GAS グローバルモックの定義 (evalする前に設定する必要がある)
let mockSheet;
let mockSpreadsheet;
let mockGmailThreads;
let mockGmailMessages;
let mockInboxFolder;
let mockRenamedFolder;
let mockImportedFolder;
let mockParentFolder;
let mockFile;
let mockProperties;
let mockUi;

// テスト前に一度だけグローバルをセットアップ
mockProperties = {
  FOLDER_ID: 'mock-folder-id',
  INBOX_FOLDER_ID: 'inbox-folder-id',
  RENAMED_FOLDER_ID: 'renamed-folder-id',
  IMPORTED_FOLDER_ID: 'imported-folder-id'
};

mockFile = {
  getId: jest.fn().mockReturnValue('mock-file-id'),
  getUrl: jest.fn().mockReturnValue('https://mock-file-url'),
  getName: jest.fn().mockReturnValue('mock-file-name.pdf'),
  copyBlob: jest.fn().mockReturnThis(),
  setName: jest.fn().mockReturnThis(),
  moveTo: jest.fn().mockReturnThis(),
  getParents: jest.fn().mockReturnValue({
    hasNext: jest.fn().mockReturnValue(true),
    next: jest.fn().mockReturnValue({
      getId: jest.fn().mockReturnValue('inbox-folder-id')
    })
  })
};

mockInboxFolder = {
  getId: jest.fn().mockReturnValue('inbox-folder-id'),
  createFile: jest.fn().mockReturnValue(mockFile),
  getFiles: jest.fn().mockReturnValue({
    hasNext: jest.fn().mockReturnValue(false),
    next: jest.fn()
  })
};

mockRenamedFolder = {
  getId: jest.fn().mockReturnValue('renamed-folder-id'),
  createFile: jest.fn().mockReturnValue(mockFile),
  getFiles: jest.fn().mockReturnValue({
    hasNext: jest.fn()
      .mockReturnValueOnce(true)
      .mockReturnValueOnce(false),
    next: jest.fn().mockReturnValue(mockFile)
  })
};

mockImportedFolder = {
  getId: jest.fn().mockReturnValue('imported-folder-id'),
  createFile: jest.fn().mockReturnValue(mockFile),
  getFiles: jest.fn().mockReturnValue({
    hasNext: jest.fn().mockReturnValue(false),
    next: jest.fn()
  })
};

mockParentFolder = {
  getId: jest.fn().mockReturnValue('mock-folder-id'),
  createFolder: jest.fn().mockImplementation(name => {
    if (name === '01.受付') return mockInboxFolder;
    if (name === '02.リネーム済み') return mockRenamedFolder;
    if (name === '03.インポート済み') return mockImportedFolder;
    return {};
  }),
  getFoldersByName: jest.fn().mockImplementation(name => {
    let target = null;
    if (name === '01.受付') target = mockInboxFolder;
    if (name === '02.リネーム済み') target = mockRenamedFolder;
    if (name === '03.インポート済み') target = mockImportedFolder;
    return {
      hasNext: jest.fn().mockReturnValue(!!target),
      next: jest.fn().mockReturnValue(target)
    };
  }),
  createFile: jest.fn().mockReturnValue(mockFile)
};

const mockRangeSummary = {
  setValue: jest.fn().mockReturnThis(),
  setFontSize: jest.fn().mockReturnThis(),
  setFontWeight: jest.fn().mockReturnThis(),
  setValues: jest.fn().mockReturnThis(),
  setBackground: jest.fn().mockReturnThis(),
  setHorizontalAlignment: jest.fn().mockReturnThis(),
  setFormula: jest.fn().mockReturnThis(),
  setNumberFormat: jest.fn().mockReturnThis()
};

const mockSummarySheet = {
  clear: jest.fn(),
  getRange: jest.fn().mockReturnValue(mockRangeSummary),
  autoResizeColumn: jest.fn()
};

mockSheet = {
  getName: jest.fn().mockReturnValue('シート1'),
  appendRow: jest.fn(),
  getLastRow: jest.fn().mockReturnValue(2),
  deleteRow: jest.fn(),
  getRange: jest.fn().mockReturnValue({
    setFormula: jest.fn(),
    setNumberFormat: jest.fn(),
    setValue: jest.fn(),
    getFormula: jest.fn().mockReturnValue(''),
    setValues: jest.fn(),
    setHorizontalAlignment: jest.fn()
  }),
  getDataRange: jest.fn().mockReturnValue({
    getValues: jest.fn().mockReturnValue([
      // ヘッダー行 (10列構成)
      ['登録日時', '取引日付', '勘定科目', '取引先名', '取引金額', 'メモ', 'ファイル名', 'ファイルID', '領収書リンク', 'CSV出力'],
      // 2行目
      [new Date(), '2026/08/18', '旅費交通費', 'タクシー', 1500, '出張', '2026.08.18_旅費交通費_タクシー_1500円_出張.pdf', 'mock-file-id', 'https://mock-file-url', '']
    ])
  })
};

mockSpreadsheet = {
  getActiveSheet: jest.fn().mockReturnValue(mockSheet),
  getSheetByName: jest.fn().mockReturnValue(mockSummarySheet),
  insertSheet: jest.fn().mockReturnValue(mockSummarySheet)
};

mockGmailMessages = [{
  getSubject: jest.fn().mockReturnValue('領収書メール'),
  getDate: jest.fn().mockReturnValue(new Date('2026-08-18T12:00:00Z')),
  getAttachments: jest.fn().mockReturnValue([{
    getSize: jest.fn().mockReturnValue(10000),
    getContentType: jest.fn().mockReturnValue('application/pdf'),
    getName: jest.fn().mockReturnValue('receipt.pdf'),
    copyBlob: jest.fn()
  }]),
  getBody: jest.fn().mockReturnValue('<html><body>領収書本文</body></html>')
}];

mockGmailThreads = [{
  getMessages: jest.fn().mockReturnValue(mockGmailMessages),
  addLabel: jest.fn(),
  removeLabel: jest.fn()
}];

mockUi = {
  alert: jest.fn(),
  ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' },
  Button: { YES: 'YES', NO: 'NO' }
};

global.Logger = { log: jest.fn() };
global.PropertiesService = {
  getScriptProperties: jest.fn().mockReturnValue({
    getProperty: jest.fn().mockImplementation(key => mockProperties[key])
  })
};
global.GmailApp = {
  getUserLabelByName: jest.fn().mockReturnValue({
    getThreads: jest.fn().mockReturnValue(mockGmailThreads)
  })
};
global.DriveApp = {
  getFolderById: jest.fn().mockImplementation(id => {
    if (id === 'inbox-folder-id') return mockInboxFolder;
    if (id === 'renamed-folder-id') return mockRenamedFolder;
    if (id === 'imported-folder-id') return mockImportedFolder;
    return mockParentFolder;
  }),
  getFileById: jest.fn().mockReturnValue(mockFile)
};
global.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn().mockReturnValue(mockSpreadsheet),
  getUi: jest.fn().mockReturnValue(mockUi),
  newDataValidation: jest.fn().mockReturnValue({
    requireValueInList: jest.fn().mockReturnThis(),
    setAllowInvalid: jest.fn().mockReturnThis(),
    setHelpText: jest.fn().mockReturnThis(),
    build: jest.fn().mockReturnValue('mock-validation-rule')
  })
};
global.Utilities = {
  formatDate: jest.fn().mockImplementation((date, tz, format) => {
    const d = new Date(date);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    if (format === 'yyyy/MM/dd') return `${y}/${m}/${day}`;
    if (format === 'yyyy.MM.dd') return `${y}.${m}.${day}`;
    if (format === 'yyyyMMdd_HHmmss') return `${y}${m}${day}_120000`;
    return '';
  }),
  newBlob: jest.fn().mockReturnValue({
    getAs: jest.fn().mockReturnThis(),
    setDataFromString: jest.fn().mockReturnThis()
  })
};
global.Session = {
  getScriptTimeZone: jest.fn().mockReturnValue('Asia/Tokyo')
};
global.Drive = {
  Files: {
    create: jest.fn().mockReturnValue({ id: 'temp-doc-id' })
  }
};

// インダイレクトevalでグローバルスコープにロードする
const codePath = path.resolve(__dirname, 'code.js');
const codeContent = fs.readFileSync(codePath, 'utf8');
(0, eval)(codeContent);

beforeEach(() => {
  // テストごとのモック履歴のリセット
  jest.clearAllMocks();

  // ファイル名の初期値リセット
  mockFile.getName = jest.fn().mockReturnValue('mock-file-name.pdf');

  // getFilesのhasNextの初期状態のリセット
  mockRenamedFolder.getFiles = jest.fn().mockReturnValue({
    hasNext: jest.fn()
      .mockReturnValueOnce(true)
      .mockReturnValueOnce(false),
    next: jest.fn().mockReturnValue(mockFile)
  });
});

describe('code.js テストスイート（ワークフロー一本化）', () => {
  test('importGmailReceipts() - 添付ファイルありのメールが「01.受付」に保存され、スプレッドシートへは記帳されないこと', () => {
    importGmailReceipts();

    // 01.受付フォルダにファイルが保存されること
    expect(mockInboxFolder.createFile).toHaveBeenCalled();
    // スプレッドシートには行追加されないこと
    expect(mockSheet.appendRow).not.toHaveBeenCalled();
    // ラベルの更新が行われること
    expect(mockGmailThreads[0].addLabel).toHaveBeenCalled();
    expect(mockGmailThreads[0].removeLabel).toHaveBeenCalled();
  });

  test('importRenamedReceipts() - 「02.リネーム済み」のファイル名がパースされ、「03.インポート済み」へ移動＆スプレッドシートに行追加されること', () => {
    mockFile.getName = jest.fn().mockReturnValue('20260815_旅費交通費_タクシー_1500円_東京出張.jpg');

    const mockFormatRange = { setNumberFormat: jest.fn() };
    mockSheet.getRange = jest.fn().mockImplementation((row, col) => {
      if (col === 5) return mockFormatRange;
      return {};
    });

    importRenamedReceipts();

    // スプレッドシートに10列で追加されること
    expect(mockSheet.appendRow).toHaveBeenCalledWith([
      expect.any(Date),
      '2026/08/15',
      '旅費交通費',
      'タクシー',
      1500,
      '東京出張',
      '2026.08.15_旅費交通費_タクシー_1500円_東京出張.jpg',
      'mock-file-id',
      'https://mock-file-url',
      ''
    ]);
    // 03.インポート済みフォルダへ移動されること
    expect(mockFile.moveTo).toHaveBeenCalledWith(mockImportedFolder);
    // E列（5列目）に金額フォーマットが適用されること
    expect(mockSheet.getRange).toHaveBeenCalledWith(2, 5);
    expect(mockFormatRange.setNumberFormat).toHaveBeenCalledWith('#,##0');
  });

  test('exportMFSheetsCSV() - リネーム済みでCSV未出力のデータがCSV出力され、J列(10)にファイル名が書き込まれること', () => {
    const mockStatusRange = {
      setValue: jest.fn()
    };
    mockSheet.getRange = jest.fn().mockImplementation((row, col) => {
      if (row === 2 && col === 10) return mockStatusRange;
      return {};
    });

    exportMFSheetsCSV();

    // J列（10列目）にステータス（CSVファイル名）が書き込まれることを確認
    expect(mockSheet.getRange).toHaveBeenCalledWith(2, 10);
    expect(mockStatusRange.setValue).toHaveBeenCalledWith(expect.stringContaining('mf_journal_'));
    expect(mockParentFolder.createFile).toHaveBeenCalled();
  });

  test('setupCategoryValidation() - A1:J1にヘッダーがセットされ、C列(3)にプルダウン、E列(5)に金額フォーマット、集計シートが設定されること', () => {
    const mockRangeHeader = { setValues: jest.fn(), setHorizontalAlignment: jest.fn() };
    const mockRangeValidation = { setDataValidation: jest.fn() };
    const mockRangeFormat = { setNumberFormat: jest.fn() };

    mockSheet.getRange = jest.fn().mockImplementation((arg1) => {
      if (arg1 === 'A1:J1') return mockRangeHeader;
      if (arg1 === 'C2:C1000') return mockRangeValidation;
      if (arg1 === 'E2:E1000') return mockRangeFormat;
      return {};
    });

    setupCategoryValidation();

    expect(mockRangeHeader.setValues).toHaveBeenCalledWith([
      ["登録日時", "取引日付", "勘定科目", "取引先名", "取引金額", "メモ", "ファイル名", "ファイルID", "領収書リンク", "CSV出力"]
    ]);
    expect(mockRangeValidation.setDataValidation).toHaveBeenCalledWith('mock-validation-rule');
    expect(mockRangeFormat.setNumberFormat).toHaveBeenCalledWith('#,##0');

    // 集計シートの設定に関するアサーション
    expect(mockSpreadsheet.getSheetByName).toHaveBeenCalledWith("集計");
    expect(mockSummarySheet.clear).toHaveBeenCalled();
    expect(mockSummarySheet.getRange).toHaveBeenCalledWith("A1");
    expect(mockRangeSummary.setValue).toHaveBeenCalledWith("勘定科目別集計");
    expect(mockSummarySheet.getRange).toHaveBeenCalledWith("B4");
    expect(mockRangeSummary.setFormula).toHaveBeenCalledWith("=COUNTIF('シート1'!C2:C, \"接待交際費\")");
    expect(mockSummarySheet.autoResizeColumn).toHaveBeenCalledWith(1);
  });

  describe('deleteExportedReceipts() テスト', () => {
    test('ユーザーが「いいえ」を選択した場合、削除が実行されないこと', () => {
      mockUi.alert.mockReturnValue('NO'); // ui.Button.NO

      deleteExportedReceipts();

      expect(mockUi.alert).toHaveBeenCalledWith(
        '確認',
        expect.stringContaining('CSV出力済みのレコードを削除しますか？'),
        'YES_NO'
      );
      expect(mockSheet.deleteRow).not.toHaveBeenCalled();
    });

    test('ユーザーが「はい」を選択した場合、CSV出力済みの行が下から順に削除されること', () => {
      mockUi.alert.mockReturnValue('YES'); // ui.Button.YES

      // J列(10列目)に値がある行とない行を混在させる (10列構成)
      mockSheet.getDataRange = jest.fn().mockReturnValue({
        getValues: jest.fn().mockReturnValue([
          ['登録日時', '取引日付', '勘定科目', '取引先名', '取引金額', 'メモ', 'ファイル名', 'ファイルID', '領収書リンク', 'CSV出力'],
          [new Date(), '2026/08/18', '旅費交通費', 'タクシー', 1500, '', 'file1.pdf', 'id1', 'url1', ''],
          [new Date(), '2026/08/18', '旅費交通費', 'タクシー', 2000, '', 'file2.pdf', 'id2', 'url2', 'mf_journal_1.csv'],
          [new Date(), '2026/08/19', '通信費', 'インターネット', 5000, '', 'file3.pdf', 'id3', 'url3', 'mf_journal_1.csv'],
          [new Date(), '2026/08/19', '会議費', 'カフェ', 800, '', 'file4.pdf', 'id4', 'url4', '']
        ])
      });

      deleteExportedReceipts();

      expect(mockSheet.deleteRow).toHaveBeenCalledTimes(2);
      expect(mockSheet.deleteRow.mock.calls[0][0]).toBe(4);
      expect(mockSheet.deleteRow.mock.calls[1][0]).toBe(3);
      expect(mockUi.alert).toHaveBeenLastCalledWith(
        '処理完了',
        '2件のレコードを削除しました。',
        'OK'
      );
    });
  });
});
