const fs = require('fs');
const path = require('path');

// GAS グローバルモックの定義 (evalする前に設定する必要がある)
let mockSheet;
let mockSpreadsheet;
let mockGmailThreads;
let mockGmailMessages;
let mockFolder;
let mockFile;
let mockProperties;
let mockUi;

// テスト前に一度だけグローバルをセットアップ
mockProperties = {
  FOLDER_ID: 'mock-folder-id',
  GMAIL_QUEUE_FOLDER_ID: 'gmail-queue-id',
  GMAIL_DONE_FOLDER_ID: 'gmail-done-id',
  MANUAL_QUEUE_FOLDER_ID: 'manual-queue-id',
  MANUAL_DONE_FOLDER_ID: 'manual-done-id'
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
      getId: jest.fn().mockReturnValue('gmail-queue-id')
    })
  })
};

mockFolder = {
  getId: jest.fn().mockReturnValue('gmail-queue-id'),
  createFile: jest.fn().mockReturnValue(mockFile),
  getFiles: jest.fn().mockReturnValue({
    hasNext: jest.fn()
      .mockReturnValueOnce(true)
      .mockReturnValueOnce(false),
    next: jest.fn().mockReturnValue(mockFile)
  }),
  getFoldersByName: jest.fn().mockReturnValue({
    hasNext: jest.fn().mockReturnValue(true),
    next: jest.fn().mockReturnValue({
      getFoldersByName: jest.fn().mockReturnValue({
        hasNext: jest.fn().mockReturnValue(true),
        next: jest.fn().mockReturnValue({
          getId: jest.fn().mockReturnValue('sub-folder-id')
        })
      })
    })
  })
};

mockSheet = {
  appendRow: jest.fn(),
  getLastRow: jest.fn().mockReturnValue(3),
  deleteRow: jest.fn(),
  getRange: jest.fn().mockReturnValue({
    setFormula: jest.fn(),
    setNumberFormat: jest.fn(),
    setValue: jest.fn(),
    getFormula: jest.fn().mockReturnValue(''),
    setValues: jest.fn(),
    setHorizontalAlignment: jest.fn(),
    getFormulas: jest.fn().mockReturnValue([['=TEXT(C3, ...)']])
  }),
  getDataRange: jest.fn().mockReturnValue({
    getValues: jest.fn().mockReturnValue([
      // 1行目: 集計行
      ['合計件数: 1件', '', '', '', '', '合計金額: 1,500円', '', '', '', '', ''],
      // 2行目: ヘッダー行
      ['取込経路', '受信日時', '取引日付', '勘定科目', '取引先名', '取引金額', 'メモ', 'ファイル名', 'ファイルID', '領収書リンク', 'CSV出力'],
      // 3行目 (Gmail未処理データ)
      ['Gmail', new Date(), '2026/08/18', '旅費交通費', 'タクシー', 1500, '出張', '=TEXT(C3, ...)', 'mock-file-id', 'https://mock-file-url', '']
    ])
  })
};

mockSpreadsheet = {
  getActiveSheet: jest.fn().mockReturnValue(mockSheet)
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
  getFolderById: jest.fn().mockReturnValue(mockFolder),
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
    // 簡易フォーマッタ
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
  mockFolder.getFiles = jest.fn().mockReturnValue({
    hasNext: jest.fn()
      .mockReturnValueOnce(true)
      .mockReturnValueOnce(false),
    next: jest.fn().mockReturnValue(mockFile)
  });
});

describe('code.js テストスイート', () => {
  test('importGmailReceipts() - 添付ファイルありのメールが正常に取り込まれ、"Gmail"がA列にセットされること', () => {
    importGmailReceipts();

    expect(mockSheet.appendRow).toHaveBeenCalledWith([
      'Gmail',
      expect.any(Date),
      '2026/08/18',
      '', '', '', '', '',
      'mock-file-id',
      'https://mock-file-url'
    ]);
  });

  test('setFilenameFormula() - H列(8)に数式が設定され、F列(6)に金額フォーマットが適用されること', () => {
    const mockRangeFormula = { setFormula: jest.fn() };
    const mockRangeFormat = { setNumberFormat: jest.fn() };
    mockSheet.getRange = jest.fn().mockImplementation((row, col) => {
      if (col === 8) return mockRangeFormula;
      if (col === 6) return mockRangeFormat;
      return {};
    });

    setFilenameFormula(mockSheet);

    expect(mockSheet.getRange).toHaveBeenCalledWith(3, 8);
    expect(mockRangeFormula.setFormula).toHaveBeenCalledWith(
      `=TEXT(C3, "yyyy.MM.dd")&"_"&D3&"_"&E3&"_"&F3&"円"&IF(G3<>"", "_"&G3, "")`
    );
    expect(mockSheet.getRange).toHaveBeenCalledWith(3, 6);
    expect(mockRangeFormat.setNumberFormat).toHaveBeenCalledWith('#,##0');
  });

  test('importManualReceipts() - 手動アップロードのファイル名がパースされ、"手動"がA列にセットされること', () => {
    // ファイル名を手動アップロード形式にする
    mockFile.getName = jest.fn().mockReturnValue('20260815_旅費交通費_タクシー_1500円_東京出張.jpg');

    importManualReceipts();

    expect(mockSheet.appendRow).toHaveBeenCalledWith([
      '手動',
      expect.any(Date),
      '2026/08/15',
      '旅費交通費',
      'タクシー',
      1500,
      '東京出張',
      '2026.08.15_旅費交通費_タクシー_1500円_東京出張.jpg',
      'mock-file-id',
      'https://mock-file-url'
    ]);
    expect(mockFile.setName).toHaveBeenCalledWith('2026.08.15_旅費交通費_タクシー_1500円_東京出張.jpg');
    expect(mockFile.moveTo).toHaveBeenCalledWith(mockFolder);
  });

  test('renameGmailReceipts() - H列(8)に数式がある行（未処理）が正しくリネームされ、H列に確定ファイル名が書き込まれること', () => {
    const mockCellRange = {
      getFormula: jest.fn().mockReturnValue('=TEXT(...)'),
      setValue: jest.fn()
    };
    mockSheet.getRange = jest.fn().mockImplementation((row, col) => {
      if (row === 3 && col === 8) return mockCellRange;
      return {};
    });

    renameGmailReceipts();

    expect(mockFile.setName).toHaveBeenCalledWith('2026.08.18_旅費交通費_タクシー_1500円_出張.pdf');
    expect(mockCellRange.setValue).toHaveBeenCalledWith('2026.08.18_旅費交通費_タクシー_1500円_出張.pdf');
  });

  test('renameGmailReceipts() - A列（取込経路）が"Gmail"ではない場合は無視されること', () => {
    // A列を「手動」にしたモックデータを返すように一時的に設定
    mockSheet.getDataRange = jest.fn().mockReturnValue({
      getValues: jest.fn().mockReturnValue([
        ['合計件数: 1件', '', '', '', '', '合計金額: 1,500円', '', '', '', '', ''],
        ['取込経路', '受信日時', '取引日付', '勘定科目', '取引先名', '取引金額', 'メモ', 'ファイル名', 'ファイルID', '領収書リンク', 'CSV出力'],
        ['手動', new Date(), '2026/08/18', '旅費交通費', 'タクシー', 1500, '出張', '=TEXT(C3, ...)', 'mock-file-id', 'https://mock-file-url', '']
      ])
    });

    const mockCellRange = {
      getFormula: jest.fn().mockReturnValue('=TEXT(...)'),
      setValue: jest.fn()
    };
    mockSheet.getRange = jest.fn().mockImplementation((row, col) => {
      if (row === 3 && col === 8) return mockCellRange;
      return {};
    });

    renameGmailReceipts();

    expect(mockFile.setName).not.toHaveBeenCalled();
    expect(mockCellRange.setValue).not.toHaveBeenCalled();
  });

  test('exportMFSheetsCSV() - リネーム済みでCSV未出力のデータがCSV出力され、K列(11)にファイル名が書き込まれること', () => {
    // 2行目の数式を空（リネーム完了済み）にする
    const mockCellRange = {
      getFormula: jest.fn().mockReturnValue(''), // 数式なし
    };
    const mockStatusRange = {
      setValue: jest.fn()
    };
    mockSheet.getRange = jest.fn().mockImplementation((row, col) => {
      if (row === 3 && col === 8) return mockCellRange;
      if (row === 3 && col === 11) return mockStatusRange;
      return {};
    });

    exportMFSheetsCSV();

    // K列（11列目）にステータス（CSVファイル名）が書き込まれることを確認
    expect(mockSheet.getRange).toHaveBeenCalledWith(3, 11);
    expect(mockStatusRange.setValue).toHaveBeenCalledWith(expect.stringContaining('mf_journal_'));
    expect(mockFolder.createFile).toHaveBeenCalled();
  });

  test('setupCategoryValidation() - A2:K2にヘッダーがセットされ、D列(4)にプルダウン、F列(6)に金額フォーマット、H列の数式がアップデートされること', () => {
    const mockRangeHeader = { setValues: jest.fn(), setHorizontalAlignment: jest.fn() };
    const mockRangeValidation = { setDataValidation: jest.fn() };
    const mockRangeFormat = { setNumberFormat: jest.fn() };
    const mockRangeFormulaCol = {
      getFormulas: jest.fn().mockReturnValue([['=TEXT(...)']]),
    };
    const mockRangeFormulaCell = {
      setFormula: jest.fn()
    };
    const mockRangeDefault = {
      setFormula: jest.fn().mockReturnThis(),
      setNumberFormat: jest.fn().mockReturnThis(),
      setFontWeight: jest.fn().mockReturnThis(),
      setBackground: jest.fn().mockReturnThis(),
      setHorizontalAlignment: jest.fn().mockReturnThis()
    };

    mockSheet.getRange = jest.fn().mockImplementation((arg1, arg2, arg3, arg4) => {
      if (arg1 === 'A2:K2') return mockRangeHeader;
      if (arg1 === 'D3:D1000') return mockRangeValidation;
      if (arg1 === 'F3:F1000') return mockRangeFormat;
      if (arg1 === 3 && arg2 === 8 && arg3 !== undefined) return mockRangeFormulaCol;
      if (arg1 === 3 && arg2 === 8 && arg3 === undefined) return mockRangeFormulaCell;
      return mockRangeDefault;
    });

    setupCategoryValidation();

    expect(mockRangeHeader.setValues).toHaveBeenCalledWith([
      ["取込経路", "受信日時", "取引日付", "勘定科目", "取引先名", "取引金額", "メモ", "ファイル名", "ファイルID", "領収書リンク", "CSV出力"]
    ]);
    expect(mockRangeValidation.setDataValidation).toHaveBeenCalledWith('mock-validation-rule');
    expect(mockRangeFormat.setNumberFormat).toHaveBeenCalledWith('#,##0');
    expect(mockRangeFormulaCell.setFormula).toHaveBeenCalledWith(
      `=TEXT(C3, "yyyy.MM.dd")&"_"&D3&"_"&E3&"_"&F3&"円"&IF(G3<>"", "_"&G3, "")`
    );
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

      // K列(11列目)に値がある行とない行を混在させる
      // 3行目: CSV未出力 (K列='')
      // 4行目: CSV出力済み (K列='mf_journal_...')
      // 5行目: CSV出力済み (K列='mf_journal_...')
      // 6行目: CSV未出力 (K列='')
      mockSheet.getDataRange = jest.fn().mockReturnValue({
        getValues: jest.fn().mockReturnValue([
          ['合計件数: 2件', '', '', '', '', '合計金額: 7,000円', '', '', '', '', ''],
          ['取込経路', '受信日時', '取引日付', '勘定科目', '取引先名', '取引金額', 'メモ', 'ファイル名', 'ファイルID', '領収書リンク', 'CSV出力'],
          ['Gmail', new Date(), '2026/08/18', '旅費交通費', 'タクシー', 1500, '', 'file1.pdf', 'id1', 'url1', ''],
          ['Gmail', new Date(), '2026/08/18', '旅費交通費', 'タクシー', 2000, '', 'file2.pdf', 'id2', 'url2', 'mf_journal_1.csv'],
          ['手動', new Date(), '2026/08/19', '通信費', 'インターネット', 5000, '', 'file3.pdf', 'id3', 'url3', 'mf_journal_1.csv'],
          ['手動', new Date(), '2026/08/19', '会議費', 'カフェ', 800, '', 'file4.pdf', 'id4', 'url4', '']
        ])
      });

      deleteExportedReceipts();

      expect(mockSheet.deleteRow).toHaveBeenCalledTimes(2);
      expect(mockSheet.deleteRow.mock.calls[0][0]).toBe(5);
      expect(mockSheet.deleteRow.mock.calls[1][0]).toBe(4);
      expect(mockUi.alert).toHaveBeenLastCalledWith(
        '処理完了',
        '2件のレコードを削除しました。',
        'OK'
      );
    });
  });
});
