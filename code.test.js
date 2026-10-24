const fs = require('fs');
const path = require('path');

// GAS イテレータヘルパー
function createIterator(items) {
  let index = 0;
  return {
    hasNext: jest.fn().mockImplementation(() => index < items.length),
    next: jest.fn().mockImplementation(() => items[index++])
  };
}

// GAS グローバルモックの定義 (evalする前に設定する必要がある)
let mockSheet;
let mockSpreadsheet;
let mockGmailThreads;
let mockGmailMessages;
let mockInboxFolder;
let mockRenamedFolder;
let mockImportedFolder;
let mockExportedFolder;
let mockParentFolder;
let mockFile;
let mockProperties;
let mockUi;

// テスト前に一度だけグローバルをセットアップ
mockProperties = {
  FOLDER_ID: 'mock-folder-id',
  INBOX_FOLDER_ID: 'inbox-folder-id',
  RENAMED_FOLDER_ID: 'renamed-folder-id',
  IMPORTED_FOLDER_ID: 'imported-folder-id',
  EXPORTED_FOLDER_ID: 'exported-folder-id'
};

mockFile = {
  getId: jest.fn().mockReturnValue('mock-file-id'),
  getUrl: jest.fn().mockReturnValue('https://mock-file-url'),
  getName: jest.fn().mockReturnValue('mock-file-name.pdf'),
  copyBlob: jest.fn().mockReturnThis(),
  setName: jest.fn().mockReturnThis(),
  moveTo: jest.fn().mockReturnThis(),
  getParents: jest.fn().mockImplementation(() => createIterator([{ getId: () => 'exported-folder-id' }]))
};

mockInboxFolder = {
  getId: jest.fn().mockReturnValue('inbox-folder-id'),
  createFile: jest.fn().mockReturnValue(mockFile),
  getFiles: jest.fn().mockImplementation(() => createIterator([]))
};

mockRenamedFolder = {
  getId: jest.fn().mockReturnValue('renamed-folder-id'),
  createFile: jest.fn().mockReturnValue(mockFile),
  getFiles: jest.fn().mockImplementation(() => createIterator([mockFile]))
};

mockImportedFolder = {
  getId: jest.fn().mockReturnValue('imported-folder-id'),
  createFile: jest.fn().mockReturnValue(mockFile),
  getFiles: jest.fn().mockImplementation(() => createIterator([]))
};

mockExportedFolder = {
  getId: jest.fn().mockReturnValue('exported-folder-id'),
  createFile: jest.fn().mockReturnValue(mockFile),
  getFiles: jest.fn().mockImplementation(() => createIterator([]))
};

mockParentFolder = {
  getId: jest.fn().mockReturnValue('mock-folder-id'),
  createFolder: jest.fn().mockImplementation(name => {
    if (name === '01.受付') return mockInboxFolder;
    if (name === '02.リネーム済み') return mockRenamedFolder;
    if (name === '03.インポート済み') return mockImportedFolder;
    if (name === '04.CSV出力済み') return mockExportedFolder;
    return {};
  }),
  getFoldersByName: jest.fn().mockImplementation(name => {
    let target = null;
    if (name === '01.受付') target = mockInboxFolder;
    if (name === '02.リネーム済み') target = mockRenamedFolder;
    if (name === '03.インポート済み') target = mockImportedFolder;
    if (name === '04.CSV出力済み') target = mockExportedFolder;
    return createIterator(target ? [target] : []);
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

const mockFixedExpensesSheet = {
  getLastRow: jest.fn().mockReturnValue(6),
  getRange: jest.fn().mockReturnValue(mockRangeSummary),
  autoResizeColumn: jest.fn(),
  getDataRange: jest.fn().mockReturnValue({
    getValues: jest.fn().mockReturnValue([
      ['勘定科目', '取引先名', '取引金額', '発生日', 'メモ'],
      ['新聞図書費', 'サンプル出版A', 1000, '毎月1日', 'メルマガ代'],
      ['新聞図書費', 'サンプル出版B', 880, '毎月末', 'メルマガ代'],
      ['通信費', 'サンプル通信', 5000, '毎月12日', ''],
      ['車両費', 'サンプル車両', 35000, '毎月2日', 'プライベート兼用のため折半'],
      ['旅費交通費', 'サンプル交通', 15000, '毎月末', '定期代']
    ])
  })
};

mockSheet = {
  getName: jest.fn().mockReturnValue('シート1'),
  setName: jest.fn().mockReturnThis(),
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
  getSheets: jest.fn().mockReturnValue([mockSheet, mockSummarySheet, mockFixedExpensesSheet]),
  getSheetByName: jest.fn().mockImplementation(name => {
    if (name === '固定経費') return mockFixedExpensesSheet;
    if (name === '集計') return mockSummarySheet;
    if (name === '領収書' || name === 'シート1') return mockSheet;
    return null;
  }),
  insertSheet: jest.fn().mockImplementation(name => {
    if (name === '固定経費') return mockFixedExpensesSheet;
    return mockSummarySheet;
  })
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

const mockMenu = {
  addItem: jest.fn().mockReturnThis(),
  addSeparator: jest.fn().mockReturnThis(),
  addToUi: jest.fn().mockReturnThis()
};

mockUi = {
  alert: jest.fn(),
  ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' },
  Button: { YES: 'YES', NO: 'NO' },
  createMenu: jest.fn().mockReturnValue(mockMenu)
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
    if (id === 'exported-folder-id') return mockExportedFolder;
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

  // getParents のデフォルト（exported-folder-id を親に持つ）
  mockFile.getParents = jest.fn().mockImplementation(() => createIterator([{ getId: () => 'exported-folder-id' }]));

  // getFilesの初期状態のリセット
  mockRenamedFolder.getFiles = jest.fn().mockImplementation(() => createIterator([mockFile]));
});

describe('code.js テストスイート（ワークフロー一本化・4フォルダ構成）', () => {
  test('importGmailReceipts() - 添付ファイルありのメールが「01.受付」に保存され、スプレッドシートへは記帳されないこと', () => {
    importGmailReceipts();

    expect(mockInboxFolder.createFile).toHaveBeenCalled();
    expect(mockSheet.appendRow).not.toHaveBeenCalled();
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
    expect(mockFile.moveTo).toHaveBeenCalledWith(mockImportedFolder);
    expect(mockSheet.getRange).toHaveBeenCalledWith(2, 5);
    expect(mockFormatRange.setNumberFormat).toHaveBeenCalledWith('#,##0');
  });

  test('exportMFSheetsCSV() - リネーム済みでCSV未出力のデータがCSV出力され、J列(10)にファイル名が書き込まれ、実ファイルが「04.CSV出力済み」へ移動されること', () => {
    const mockStatusRange = {
      setValue: jest.fn()
    };
    mockSheet.getRange = jest.fn().mockImplementation((row, col) => {
      if (row === 2 && col === 10) return mockStatusRange;
      return {};
    });

    exportMFSheetsCSV();

    expect(mockSheet.getRange).toHaveBeenCalledWith(2, 10);
    expect(mockStatusRange.setValue).toHaveBeenCalledWith(expect.stringContaining('mf_journal_'));
    expect(mockParentFolder.createFile).toHaveBeenCalled();
    expect(mockFile.moveTo).toHaveBeenCalledWith(mockExportedFolder);
  });

  test('exportMFSheetsCSV() - 実ファイルIDがない固定経費レコードも正しくCSV出力対象となり、ファイル移動なしでステータス更新されること', () => {
    const mockStatusRange = { setValue: jest.fn() };
    mockSheet.getRange = jest.fn().mockImplementation((row, col) => {
      if (row === 2 && col === 10) return mockStatusRange;
      return {};
    });

    mockSheet.getDataRange = jest.fn().mockReturnValue({
      getValues: jest.fn().mockReturnValue([
        ['登録日時', '取引日付', '勘定科目', '取引先名', '取引金額', 'メモ', 'ファイル名', 'ファイルID', '領収書リンク', 'CSV出力'],
        [new Date(), '2026/10/01', '新聞図書費', 'サンプル出版A', 1000, 'メルマガ代', '固定経費', '', '', '']
      ])
    });

    exportMFSheetsCSV();

    expect(mockSheet.getRange).toHaveBeenCalledWith(2, 10);
    expect(mockStatusRange.setValue).toHaveBeenCalledWith(expect.stringContaining('mf_journal_'));
    expect(mockParentFolder.createFile).toHaveBeenCalled();
    // ファイル移動は呼ばれないこと
    expect(mockFile.moveTo).not.toHaveBeenCalled();
    expect(mockUi.alert).toHaveBeenCalledWith(
      '出力完了',
      expect.stringContaining('1件のデータをマネーフォワード用CSVとして出力'),
      'OK'
    );
  });

  test('setupCategoryValidation() - A1:J1にヘッダーがセットされ、不要な列の入力規則がクリアされ、C列(3)にプルダウン、E列(5)に金額フォーマット、集計シート・固定経費シートが設定されること', () => {
    const mockRangeHeader = { setValues: jest.fn(), setHorizontalAlignment: jest.fn() };
    const mockRangeValidation = { setDataValidation: jest.fn() };
    const mockRangeFormat = { setNumberFormat: jest.fn() };
    const mockRangeClear = { clearDataValidations: jest.fn() };

    mockSheet.getRange = jest.fn().mockImplementation((arg1) => {
      if (arg1 === 'A1:J1') return mockRangeHeader;
      if (arg1 === 'C2:C1000') return mockRangeValidation;
      if (arg1 === 'E2:E1000') return mockRangeFormat;
      if (arg1 === 'A2:B1000' || arg1 === 'D2:J1000') return mockRangeClear;
      return {};
    });

    setupCategoryValidation();

    expect(mockRangeClear.clearDataValidations).toHaveBeenCalledTimes(2);
    expect(mockRangeHeader.setValues).toHaveBeenCalledWith([
      ["登録日時", "取引日付", "勘定科目", "取引先名", "取引金額", "メモ", "ファイル名", "ファイルID", "領収書リンク", "CSV出力"]
    ]);
    expect(mockRangeValidation.setDataValidation).toHaveBeenCalledWith('mock-validation-rule');
    expect(mockRangeFormat.setNumberFormat).toHaveBeenCalledWith('#,##0');

    expect(mockSpreadsheet.getSheetByName).toHaveBeenCalledWith("集計");
    expect(mockSpreadsheet.getSheetByName).toHaveBeenCalledWith("固定経費");
    expect(mockSummarySheet.clear).toHaveBeenCalled();
    expect(mockSummarySheet.getRange).toHaveBeenCalledWith("A1");
    expect(mockRangeSummary.setValue).toHaveBeenCalledWith("勘定科目別集計");
    expect(mockSummarySheet.getRange).toHaveBeenCalledWith("B4");
    expect(mockRangeSummary.setFormula).toHaveBeenCalledWith("=COUNTIF('シート1'!C2:C, \"接待交際費\")");
    expect(mockSummarySheet.autoResizeColumn).toHaveBeenCalledWith(1);
  });

  test('setupCategoryValidation() - 「領収書」シートがなく「シート1」が存在する場合、「領収書」にリネームされること', () => {
    mockSpreadsheet.getSheetByName = jest.fn().mockImplementation(name => {
      if (name === 'シート1') return mockSheet;
      if (name === '固定経費') return mockFixedExpensesSheet;
      if (name === '集計') return mockSummarySheet;
      return null;
    });

    setupCategoryValidation();

    expect(mockSheet.setName).toHaveBeenCalledWith('領収書');
  });

  describe('deleteExportedReceipts() テスト', () => {
    test('ユーザーが「いいえ」を選択した場合、削除が実行されないこと', () => {
      mockUi.alert.mockReturnValue('NO'); // ui.Button.NO

      deleteExportedReceipts();

      expect(mockUi.alert).toHaveBeenCalledWith(
        '確認',
        expect.stringContaining('「04.CSV出力済み」フォルダに移動完了しているレコードを台帳から削除しますか？'),
        'YES_NO'
      );
      expect(mockSheet.deleteRow).not.toHaveBeenCalled();
    });

    test('ユーザーが「はい」を選択した場合、「04.CSV出力済み」フォルダに実ファイルが存在する行のみ削除されること', () => {
      mockUi.alert.mockReturnValue('YES'); // ui.Button.YES

      // ファイルモックをIDごとに分岐
      global.DriveApp.getFileById = jest.fn().mockImplementation(id => {
        if (id === 'id-in-exported') {
          return {
            getParents: jest.fn().mockImplementation(() => createIterator([{ getId: () => 'exported-folder-id' }]))
          };
        }
        if (id === 'id-in-other-folder') {
          return {
            getParents: jest.fn().mockImplementation(() => createIterator([{ getId: () => 'other-folder-id' }]))
          };
        }
        return mockFile;
      });

      // 10列構成
      // 2行目: CSV未出力 (J列='')
      // 3行目: CSV出力済み ＆ 04.CSV出力済みに存在 (id-in-exported) -> 削除対象
      // 4行目: CSV出力済み だが 04.CSV出力済みに存在しない (id-in-other-folder) -> スキップ
      // 5行目: CSV未出力 (J列='')
      mockSheet.getDataRange = jest.fn().mockReturnValue({
        getValues: jest.fn().mockReturnValue([
          ['登録日時', '取引日付', '勘定科目', '取引先名', '取引金額', 'メモ', 'ファイル名', 'ファイルID', '領収書リンク', 'CSV出力'],
          [new Date(), '2026/08/18', '旅費交通費', 'タクシー', 1500, '', 'file1.pdf', 'id1', 'url1', ''],
          [new Date(), '2026/08/18', '旅費交通費', 'タクシー', 2000, '', 'file2.pdf', 'id-in-exported', 'url2', 'mf_journal_1.csv'],
          [new Date(), '2026/08/19', '通信費', 'インターネット', 5000, '', 'file3.pdf', 'id-in-other-folder', 'url3', 'mf_journal_1.csv'],
          [new Date(), '2026/08/19', '会議費', 'カフェ', 800, '', 'file4.pdf', 'id4', 'url4', '']
        ])
      });

      deleteExportedReceipts();

      // 3行目のみ削除されること
      expect(mockSheet.deleteRow).toHaveBeenCalledTimes(1);
      expect(mockSheet.deleteRow).toHaveBeenCalledWith(3);
      expect(mockUi.alert).toHaveBeenLastCalledWith(
        '処理完了',
        expect.stringContaining('1件のレコードを削除しました。'),
        'OK'
      );
    });

    test('実ファイルIDがないレコード（固定経費等）も、CSV出力済みであれば削除対象となること', () => {
      mockUi.alert.mockReturnValue('YES');

      mockSheet.getDataRange = jest.fn().mockReturnValue({
        getValues: jest.fn().mockReturnValue([
          ['登録日時', '取引日付', '勘定科目', '取引先名', '取引金額', 'メモ', 'ファイル名', 'ファイルID', '領収書リンク', 'CSV出力'],
          [new Date(), '2026/10/01', '新聞図書費', 'サンプル出版A', 1000, 'メルマガ代', '固定経費', '', '', 'mf_journal_1.csv'],
          [new Date(), '2026/10/02', '車両費', 'サンプル車両', 35000, 'プライベート兼用のため折半', '固定経費', '', '', '']
        ])
      });

      deleteExportedReceipts();

      // 2行目（固定経費でCSV出力済み）のみ削除されること
      expect(mockSheet.deleteRow).toHaveBeenCalledTimes(1);
      expect(mockSheet.deleteRow).toHaveBeenCalledWith(2);
      expect(mockUi.alert).toHaveBeenLastCalledWith(
        '処理完了',
        expect.stringContaining('1件のレコードを削除しました。'),
        'OK'
      );
    });
  });

  describe('固定経費の一括追加機能テスト', () => {
    test('calculateFixedExpenseDate() - 計上期間（前月26日〜今月25日）に基づく正しい取引日付が計算されること', () => {
      const baseDate = new Date('2026-10-24T12:00:00Z');

      expect(calculateFixedExpenseDate('毎月1日', baseDate)).toBe('2026/10/01');
      expect(calculateFixedExpenseDate('毎月2日', baseDate)).toBe('2026/10/02');
      expect(calculateFixedExpenseDate('毎月12日', baseDate)).toBe('2026/10/12');
      expect(calculateFixedExpenseDate('毎月末', baseDate)).toBe('2026/09/30');

      // 年跨ぎ（1月実行時）のテスト
      const janDate = new Date('2026-01-15T12:00:00Z');
      expect(calculateFixedExpenseDate('毎月1日', janDate)).toBe('2026/01/01');
      expect(calculateFixedExpenseDate('毎月末', janDate)).toBe('2025/12/31');
    });

    test('addFixedExpenses() - 固定経費シートにデータが存在しない場合はアラートが表示されること', () => {
      mockFixedExpensesSheet.getDataRange = jest.fn().mockReturnValue({
        getValues: jest.fn().mockReturnValue([
          ['勘定科目', '取引先名', '取引金額', '発生日', 'メモ']
        ])
      });

      addFixedExpenses();

      expect(mockUi.alert).toHaveBeenCalledWith(
        '固定経費データがありません',
        expect.stringContaining('「固定経費」シートに取引データが登録されていません'),
        'OK'
      );
      expect(mockSheet.appendRow).not.toHaveBeenCalled();
    });

    test('addFixedExpenses() - 「固定経費」シートの5件の取引が正しい項目・値で追加されること', () => {
      const mockFormatRange = { setNumberFormat: jest.fn() };
      mockSheet.getRange = jest.fn().mockReturnValue(mockFormatRange);
      mockSheet.appendRow = jest.fn();
      mockSheet.getLastRow = jest.fn().mockReturnValue(2);

      // 固定経費シートのモックデータを再設定
      mockFixedExpensesSheet.getDataRange = jest.fn().mockReturnValue({
        getValues: jest.fn().mockReturnValue([
          ['勘定科目', '取引先名', '取引金額', '発生日', 'メモ'],
          ['新聞図書費', 'サンプル出版A', 1000, '毎月1日', 'メルマガ代'],
          ['新聞図書費', 'サンプル出版B', 880, '毎月末', 'メルマガ代'],
          ['通信費', 'サンプル通信', 5000, '毎月12日', ''],
          ['車両費', 'サンプル車両', 35000, '毎月2日', 'プライベート兼用のため折半'],
          ['旅費交通費', 'サンプル交通', 15000, '毎月末', '定期代']
        ])
      });

      // 台帳がヘッダーのみ（既存取引なし）の状態
      mockSheet.getDataRange = jest.fn().mockReturnValue({
        getValues: jest.fn().mockReturnValue([
          ['登録日時', '取引日付', '勘定科目', '取引先名', '取引金額', 'メモ', 'ファイル名', 'ファイルID', '領収書リンク', 'CSV出力']
        ])
      });

      addFixedExpenses();

      // 5件追加されること
      expect(mockSheet.appendRow).toHaveBeenCalledTimes(5);

      const addedRows = mockSheet.appendRow.mock.calls.map(call => call[0]);

      // 1. 新聞図書費: サンプル出版A
      const pubA = addedRows.find(r => r[3] === 'サンプル出版A');
      expect(pubA).toBeDefined();
      expect(pubA[2]).toBe('新聞図書費');
      expect(pubA[4]).toBe(1000);
      expect(pubA[5]).toBe('メルマガ代');
      expect(pubA[6]).toBe('固定経費');

      // 2. 新聞図書費: サンプル出版B
      const pubB = addedRows.find(r => r[3] === 'サンプル出版B');
      expect(pubB).toBeDefined();
      expect(pubB[2]).toBe('新聞図書費');
      expect(pubB[4]).toBe(880);
      expect(pubB[5]).toBe('メルマガ代');

      // 3. 車両費: サンプル車両
      const car = addedRows.find(r => r[3] === 'サンプル車両');
      expect(car).toBeDefined();
      expect(car[2]).toBe('車両費');
      expect(car[4]).toBe(35000);
      expect(car[5]).toBe('プライベート兼用のため折半');

      // 4. 通信費: サンプル通信
      const comm = addedRows.find(r => r[3] === 'サンプル通信');
      expect(comm).toBeDefined();
      expect(comm[2]).toBe('通信費');
      expect(comm[4]).toBe(5000);
      expect(comm[5]).toBe('');

      // 5. 旅費交通費: サンプル交通
      const train = addedRows.find(r => r[3] === 'サンプル交通');
      expect(train).toBeDefined();
      expect(train[2]).toBe('旅費交通費');
      expect(train[4]).toBe(15000);
      expect(train[5]).toBe('定期代');

      // 金額フォーマット設定が各行に対して呼ばれること
      expect(mockFormatRange.setNumberFormat).toHaveBeenCalledWith('#,##0');

      // 完了ダイアログの確認
      expect(mockUi.alert).toHaveBeenCalledWith(
        '追加完了',
        expect.stringContaining('5件の固定経費を台帳に追加しました。'),
        'OK'
      );
    });

    test('addFixedExpenses() - 既に台帳に追加済みの固定経費がある場合はスキップされること（二重登録防止）', () => {
      mockSheet.appendRow = jest.fn();

      // 現在の固定経費と同じ取引日の日付を算出
      const now = new Date();
      const pubADate = calculateFixedExpenseDate('毎月1日', now);

      mockFixedExpensesSheet.getDataRange = jest.fn().mockReturnValue({
        getValues: jest.fn().mockReturnValue([
          ['勘定科目', '取引先名', '取引金額', '発生日', 'メモ'],
          ['新聞図書費', 'サンプル出版A', 1000, '毎月1日', 'メルマガ代'],
          ['新聞図書費', 'サンプル出版B', 880, '毎月末', 'メルマガ代']
        ])
      });

      mockSheet.getDataRange = jest.fn().mockReturnValue({
        getValues: jest.fn().mockReturnValue([
          ['登録日時', '取引日付', '勘定科目', '取引先名', '取引金額', 'メモ', 'ファイル名', 'ファイルID', '領収書リンク', 'CSV出力'],
          // すでにサンプル出版Aが未出力で存在
          [new Date(), pubADate, '新聞図書費', 'サンプル出版A', 1000, 'メルマガ代', '固定経費', '', '', '']
        ])
      });

      addFixedExpenses();

      // サンプル出版Aがスキップされ、サンプル出版Bのみ追加されること
      expect(mockSheet.appendRow).toHaveBeenCalledTimes(1);
      expect(mockUi.alert).toHaveBeenCalledWith(
        '追加完了',
        expect.stringContaining('1件の固定経費を台帳に追加しました。'),
        'OK'
      );
      expect(mockUi.alert).toHaveBeenCalledWith(
        '追加完了',
        expect.stringContaining('すでに追加済みの固定経費 1件 をスキップしました。'),
        'OK'
      );
    });
  });

  describe('getReceiptDataSheet() テスト', () => {
    test('「領収書」シートが存在する場合はそれを返すこと', () => {
      mockSpreadsheet.getSheetByName = jest.fn().mockImplementation(name => {
        if (name === '領収書') return mockSheet;
        return null;
      });

      const sheet = getReceiptDataSheet(mockSpreadsheet);
      expect(sheet).toBe(mockSheet);
    });

    test('「領収書」がなく「シート1」が存在する場合はリネームせず「シート1」を返すこと', () => {
      mockSpreadsheet.getSheetByName = jest.fn().mockImplementation(name => {
        if (name === 'シート1') return mockSheet;
        return null;
      });

      const sheet = getReceiptDataSheet(mockSpreadsheet);
      expect(sheet).toBe(mockSheet);
      expect(mockSheet.setName).not.toHaveBeenCalled();
    });

    test('アクティブシートが固定経費や集計の場合でも、領収書シートが正しく特定されること', () => {
      mockSpreadsheet.getActiveSheet = jest.fn().mockReturnValue(mockFixedExpensesSheet);
      mockSpreadsheet.getSheetByName = jest.fn().mockImplementation(name => {
        if (name === '領収書') return mockSheet;
        if (name === '固定経費') return mockFixedExpensesSheet;
        return null;
      });

      const sheet = getReceiptDataSheet(mockSpreadsheet);
      expect(sheet).toBe(mockSheet);
    });
  });

  test('onOpen() - メニューに「固定経費の一括追加」が含まれていること', () => {
    onOpen();
    expect(mockUi.createMenu).toHaveBeenCalledWith('領収書管理');
    expect(mockMenu.addItem).toHaveBeenCalledWith('固定経費の一括追加', 'addFixedExpenses');
    expect(mockMenu.addToUi).toHaveBeenCalled();
  });
});
