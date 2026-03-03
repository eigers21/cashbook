/**
 * 現金出納帳 Webアプリ — バックエンド
 * Google Apps Script（GAS）で動作するエンドポイント群
 */

// ========================================
// 定数
// ========================================

/** スプレッドシートID（デプロイ前に設定してください） */
const SPREADSHEET_ID = '1aFVeZ7crB3NTg5g9LjmB1MojdS4VP85GUq_sBXlfeKk';

/** マスタシート名 */
const MASTER_SHEET_NAME = 'マスタ_項目';

/** ヘッダー行の定義 */
const HEADERS = ['年', '月', '日', '摘要', '収入金額', '支払金額', '差引残高', '備考'];

/** 金庫シートのヘッダー */
const VAULT_HEADERS = ['金種', '枚数', '小計'];

/** 金種一覧（降順） */
const DENOMINATIONS = [10000, 5000, 1000, 500, 100, 50, 10, 5, 1];

// ========================================
// WebApp エントリーポイント
// ========================================

/**
 * WebApp の HTML を返す / GETパラメータでAPI呼び出しにも対応
 * @param {Object} e - リクエストパラメータ
 */
function doGet(e) {
  // API呼び出し（GitHub Pages からの fetch 用）
  if (e && e.parameter && e.parameter.action) {
    return handleApiRequest(e.parameter);
  }

  // 通常のWebApp配信（GAS WebApp用、後方互換）
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('現金出納帳')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * POST リクエストを処理する（GitHub Pages からの API 呼び出し用）
 * @param {Object} e - リクエストパラメータ
 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    return handleApiRequest(body);
  } catch (err) {
    return jsonResponse({ error: 'リクエストの解析に失敗: ' + err.message });
  }
}

/**
 * API リクエストをルーティングする
 * @param {Object} params - { action: string, ... }
 */
function handleApiRequest(params) {
  try {
    var action = params.action;
    var result;

    switch (action) {
      case 'getBalance':
        result = getBalance(params.sheetName);
        break;
      case 'getCategories':
        result = getCategories();
        break;
      case 'addRecord':
        result = addRecord(params.data);
        break;
      case 'saveVault':
        result = saveVault(params.data);
        break;
      case 'getVault':
        result = getVault();
        break;
      case 'getAvailableMonths':
        result = getAvailableMonths();
        break;
      case 'getMonthlyRecords':
        result = getMonthlyRecords(params.sheetName);
        break;
      case 'updateRecord':
        result = updateRecord(params.data);
        break;
      case 'deleteRecord':
        result = deleteRecord(params.data);
        break;
      case 'bulkSync':
        result = bulkSyncTransactions(params.data);
        break;
      default:
        result = { error: '不明なアクション: ' + action };
    }

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

/**
 * JSON レスポンスを生成する
 * @param {Object} data - レスポンスデータ
 */
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========================================
// ヘルパー関数
// ========================================

/**
 * スプレッドシートを取得する
 * @return {Spreadsheet} スプレッドシートオブジェクト
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * 当月のシート名を取得する（YYYY_MM形式）
 * @param {Date} [date] - 対象日付（省略時は現在日時）
 * @return {string} シート名（例: 2026_03）
 */
function getCurrentSheetName(date) {
  const d = date || new Date();
  const year = d.getFullYear();
  const month = ('0' + (d.getMonth() + 1)).slice(-2);
  return year + '_' + month;
}

/**
 * 前月のシート名を取得する
 * @param {Date} [date] - 対象日付（省略時は現在日時）
 * @return {string} 前月のシート名
 */
function getPreviousSheetName(date) {
  const d = date || new Date();
  const prevDate = new Date(d.getFullYear(), d.getMonth() - 1, 1);
  return getCurrentSheetName(prevDate);
}

/**
 * 当月シートを取得する。存在しない場合は自動生成する
 * @return {Sheet} シートオブジェクト
 */
function getOrCreateCurrentSheet() {
  const ss = getSpreadsheet();
  const sheetName = getCurrentSheetName();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    // シートが存在しない場合は新規作成
    sheet = createNewMonthlySheet(ss, sheetName);
  }

  return sheet;
}

/**
 * 新しい月次シートを作成する
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} sheetName - 作成するシート名
 * @return {Sheet} 作成されたシート
 */
function createNewMonthlySheet(ss, sheetName) {
  const sheet = ss.insertSheet(sheetName);

  // ヘッダー行を設定
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  // ヘッダー行の書式設定
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4a86c8');
  headerRange.setFontColor('#ffffff');

  // 列幅の調整
  sheet.setColumnWidth(1, 50);  // 年
  sheet.setColumnWidth(2, 40);  // 月
  sheet.setColumnWidth(3, 40);  // 日
  sheet.setColumnWidth(4, 120); // 摘要
  sheet.setColumnWidth(5, 100); // 収入金額
  sheet.setColumnWidth(6, 100); // 支払金額
  sheet.setColumnWidth(7, 120); // 差引残高
  sheet.setColumnWidth(8, 150); // 備考

  // 繰越行を設定
  const carryOver = getCarryOverBalance(ss, sheetName);
  const now = new Date();
  const year = now.getFullYear() % 100; // 2桁の年
  const month = now.getMonth() + 1;

  sheet.getRange(2, 1, 1, HEADERS.length).setValues([
    [year, month, 1, '繰越', '', '', carryOver, '']
  ]);

  // 金額列のフォーマット設定
  sheet.getRange(2, 5, 1000, 3).setNumberFormat('#,##0');

  return sheet;
}

/**
 * 前月シートの最終残高を取得する（繰越金額）
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {string} currentSheetName - 当月シート名
 * @return {number} 繰越金額
 */
function getCarryOverBalance(ss, currentSheetName) {
  // 当月のシート名から前月を計算
  const parts = currentSheetName.split('_');
  const year = parseInt(parts[0]);
  const month = parseInt(parts[1]);

  let prevYear, prevMonth;
  if (month === 1) {
    prevYear = year - 1;
    prevMonth = 12;
  } else {
    prevYear = year;
    prevMonth = month - 1;
  }

  const prevSheetName = prevYear + '_' + ('0' + prevMonth).slice(-2);
  const prevSheet = ss.getSheetByName(prevSheetName);

  if (!prevSheet) {
    // 前月シートが存在しない場合は0を返す
    return 0;
  }

  return getLastBalance(prevSheet);
}

/**
 * シートの最終行の差引残高を取得する
 * @param {Sheet} sheet - シートオブジェクト
 * @return {number} 差引残高
 */
function getLastBalance(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 0;
  }

  const balance = sheet.getRange(lastRow, 7).getValue(); // G列 = 差引残高
  return balance || 0;
}

// ========================================
// 公開エンドポイント
// ========================================

/**
 * 指定した月、または当月の残高情報を取得する
 * @param {string} [sheetName] - 対象シート名（例: 2026_03）
 * @return {Object} { balance: number, income: number, expense: number, month: string }
 */
function getBalance(sheetName) {
  try {
    const ss = getSpreadsheet();
    const name = sheetName || getCurrentSheetName();
    let sheet = ss.getSheetByName(name);

    if (!sheet) {
      if (!sheetName) {
        sheet = getOrCreateCurrentSheet();
      } else {
        throw new Error('指定された月のデータが見つかりません');
      }
    }

    const lastRow = sheet.getLastRow();
    let balance = 0;
    let income = 0;
    let expense = 0;

    if (lastRow >= 2) {
      // 最終行の差引残高を取得
      balance = sheet.getRange(lastRow, 7).getValue() || 0;

      // 収入・支出の合計を取得（3行目以降、ヘッダーと繰越行を除く）
      if (lastRow >= 3) {
        const incomeRange = sheet.getRange(3, 5, lastRow - 2, 1).getValues();
        const expenseRange = sheet.getRange(3, 6, lastRow - 2, 1).getValues();

        income = incomeRange.reduce(function(sum, row) {
          return sum + (Number(row[0]) || 0);
        }, 0);

        expense = expenseRange.reduce(function(sum, row) {
          return sum + (Number(row[0]) || 0);
        }, 0);
      }
    }

    return {
      balance: balance,
      income: income,
      expense: expense,
      month: name
    };
  } catch (e) {
    throw new Error('残高の取得に失敗しました: ' + e.message);
  }
}

/**
 * マスタ_項目シートから項目一覧を取得する
 * @return {Object[]} { name: string, type: string } の配列
 */
function getCategories() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(MASTER_SHEET_NAME);

    if (!sheet) {
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }

    // A列(項目名)とB列(区分)をまとめて取得
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    return values
      .map(function(row) {
        return {
          name: row[0],
          type: row[1] // "収入" または "支出"
        };
      })
      .filter(function(item) { return item.name !== ''; });
  } catch (e) {
    throw new Error('項目一覧の取得に失敗しました: ' + e.message);
  }
}

/**
 * 収支データを当月シートに追記する
 * @param {Object} data - { type: "income"|"expense", category: string, amount: number, note: string }
 * @return {Object} { success: boolean, balance: number }
 */
function addRecord(data) {
  try {
    // バリデーション
    if (!data || !data.type || !data.category || !data.amount) {
      throw new Error('必須項目が入力されていません');
    }

    const amount = Number(data.amount);
    if (isNaN(amount) || amount <= 0) {
      throw new Error('金額が不正です');
    }

    const sheet = getOrCreateCurrentSheet();
    const lastRow = sheet.getLastRow();

    // 前行の差引残高を取得
    const prevBalance = getLastBalance(sheet);

    // 収入・支出・新残高を計算
    let incomeAmount = 0;
    let expenseAmount = 0;
    let newBalance = prevBalance;

    if (data.type === 'income') {
      incomeAmount = amount;
      newBalance = prevBalance + amount;
    } else {
      expenseAmount = amount;
      newBalance = prevBalance - amount;
    }

    // 日付情報
    const now = new Date();
    const year = now.getFullYear() % 100; // 2桁の年
    const month = now.getMonth() + 1;
    const day = now.getDate();

    // 新しい行を追記
    const newRow = [
      year,
      month,
      day,
      data.category,
      incomeAmount || '',
      expenseAmount || '',
      newBalance,
      data.note || ''
    ];

    sheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);

    // 金額列のフォーマット設定
    sheet.getRange(lastRow + 1, 5, 1, 3).setNumberFormat('#,##0');

    return {
      success: true,
      balance: newBalance
    };
  } catch (e) {
    throw new Error('記録の追加に失敗しました: ' + e.message);
  }
}

/**
 * 金庫の枚数データを保存する
 * @param {Object} data - { denominations: { "10000": n, "5000": n, ... } }
 * @return {Object} { success: boolean, total: number }
 */
function saveVault(data) {
  try {
    if (!data || !data.denominations) {
      throw new Error('金庫データが不正です');
    }

    const ss = getSpreadsheet();
    const now = new Date();
    const vaultSheetName = '金庫_' + getCurrentSheetName(now);

    // 金庫シートを取得または作成
    let vaultSheet = ss.getSheetByName(vaultSheetName);
    if (!vaultSheet) {
      vaultSheet = ss.insertSheet(vaultSheetName);
    } else {
      // 既存データをクリア
      vaultSheet.clear();
    }

    // ヘッダーを設定
    vaultSheet.getRange(1, 1, 1, VAULT_HEADERS.length).setValues([VAULT_HEADERS]);
    const headerRange = vaultSheet.getRange(1, 1, 1, VAULT_HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4a86c8');
    headerRange.setFontColor('#ffffff');

    // 金種データを書き込み
    let total = 0;
    const rows = [];

    DENOMINATIONS.forEach(function(denom) {
      const count = Number(data.denominations[String(denom)]) || 0;
      const subtotal = denom * count;
      total += subtotal;
      rows.push([denom.toLocaleString() + '円', count, subtotal]);
    });

    // 合計行
    rows.push(['合計', '', total]);

    vaultSheet.getRange(2, 1, rows.length, 3).setValues(rows);

    // 金額フォーマット
    vaultSheet.getRange(2, 3, rows.length, 1).setNumberFormat('#,##0');

    // 列幅設定
    vaultSheet.setColumnWidth(1, 100);
    vaultSheet.setColumnWidth(2, 60);
    vaultSheet.setColumnWidth(3, 120);

    // 更新日時を記録
    vaultSheet.getRange(rows.length + 3, 1).setValue('最終更新: ' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));

    return {
      success: true,
      total: total
    };
  } catch (e) {
    throw new Error('金庫データの保存に失敗しました: ' + e.message);
  }
}

/**
 * 当月の金庫データを取得する
 * @return {Object} { denominations: { "10000": n, ... }, total: number }
 */
function getVault() {
  try {
    const ss = getSpreadsheet();
    const vaultSheetName = '金庫_' + getCurrentSheetName();
    const vaultSheet = ss.getSheetByName(vaultSheetName);

    const result = {
      denominations: {},
      total: 0
    };

    // 初期値を設定
    DENOMINATIONS.forEach(function(denom) {
      result.denominations[String(denom)] = 0;
    });

    if (!vaultSheet) {
      return result;
    }

    const lastRow = vaultSheet.getLastRow();
    if (lastRow < 2) {
      return result;
    }

    // 金種データを読み込み（2行目から金種数分）
    const dataRows = Math.min(lastRow - 1, DENOMINATIONS.length);
    const values = vaultSheet.getRange(2, 2, dataRows, 1).getValues();

    let total = 0;
    DENOMINATIONS.forEach(function(denom, index) {
      if (index < values.length) {
        const count = Number(values[index][0]) || 0;
        result.denominations[String(denom)] = count;
        total += denom * count;
      }
    });

    result.total = total;
    return result;
  } catch (e) {
    throw new Error('金庫データの取得に失敗しました: ' + e.message);
  }
}

// ========================================
// 月別明細・修正機能
// ========================================

/**
 * 利用可能な月シートの一覧を取得する（新しい順）
 * @return {Object[]} { name: string, label: string } の配列
 */
function getAvailableMonths() {
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets();
    const monthSheets = [];

    sheets.forEach(function(sheet) {
      const name = sheet.getName();
      // YYYY_MM 形式のシートだけを抽出
      if (/^\d{4}_\d{2}$/.test(name)) {
        const parts = name.split('_');
        monthSheets.push({
          name: name,
          label: parts[0] + '年 ' + parseInt(parts[1]) + '月'
        });
      }
    });

    // 新しい順にソート
    monthSheets.sort(function(a, b) {
      return b.name.localeCompare(a.name);
    });

    return monthSheets;
  } catch (e) {
    throw new Error('月一覧の取得に失敗しました: ' + e.message);
  }
}

/**
 * 指定月のレコード一覧を取得する
 * @param {string} sheetName - シート名（例: 2026_03）
 * @return {Object[]} レコードの配列
 */
function getMonthlyRecords(sheetName) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error('シート "' + sheetName + '" が見つかりません');
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }

    // 2行目以降のデータを取得（繰越行も含む）
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const records = [];

    data.forEach(function(row, index) {
      records.push({
        rowIndex: index + 2, // スプレッドシート上の実際の行番号
        year: row[0],
        month: row[1],
        day: row[2],
        category: row[3],
        income: row[4] || 0,
        expense: row[5] || 0,
        balance: row[6] || 0,
        note: row[7] || '',
        isCarryOver: row[3] === '繰越'
      });
    });

    return records;
  } catch (e) {
    throw new Error('明細の取得に失敗しました: ' + e.message);
  }
}

/**
 * 指定行のレコードを更新する
 * @param {Object} data - { sheetName, rowIndex, category, income, expense, note }
 * @return {Object} { success: boolean }
 */
function updateRecord(data) {
  try {
    if (!data || !data.sheetName || !data.rowIndex) {
      throw new Error('更新データが不正です');
    }

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(data.sheetName);

    if (!sheet) {
      throw new Error('シートが見つかりません');
    }

    const row = data.rowIndex;

    // 摘要（D列）を更新
    if (data.category !== undefined) {
      sheet.getRange(row, 4).setValue(data.category);
    }

    // 収入金額（E列）を更新
    if (data.income !== undefined) {
      sheet.getRange(row, 5).setValue(data.income || '');
    }

    // 支出金額（F列）を更新
    if (data.expense !== undefined) {
      sheet.getRange(row, 6).setValue(data.expense || '');
    }

    // 備考（H列）を更新
    if (data.note !== undefined) {
      sheet.getRange(row, 8).setValue(data.note);
    }

    // 差引残高を再計算（変更行以降すべて）
    recalcBalances(sheet, row);

    return { success: true };
  } catch (e) {
    throw new Error('レコードの更新に失敗しました: ' + e.message);
  }
}

/**
 * 指定行以降の差引残高を再計算する
 * @param {Sheet} sheet - シートオブジェクト
 * @param {number} startRow - 再計算を開始する行番号
 */
function recalcBalances(sheet, startRow) {
  const lastRow = sheet.getLastRow();
  if (startRow < 2 || startRow > lastRow) return;

  // startRow の1つ前の行の残高を取得
  let prevBalance = 0;
  if (startRow > 2) {
    prevBalance = sheet.getRange(startRow - 1, 7).getValue() || 0;
  }

  // startRow から最終行まで再計算
  for (let r = startRow; r <= lastRow; r++) {
    const income = Number(sheet.getRange(r, 5).getValue()) || 0;
    const expense = Number(sheet.getRange(r, 6).getValue()) || 0;
    const category = sheet.getRange(r, 4).getValue();

    if (category === '繰越') {
      // 繰越行の残高はそのまま（手動設定値を維持）
      prevBalance = sheet.getRange(r, 7).getValue() || 0;
    } else {
      prevBalance = prevBalance + income - expense;
      sheet.getRange(r, 7).setValue(prevBalance);
    }
  }
}

/**
 * 指定行のレコードを削除する
 * @param {Object} data - { sheetName, rowIndex }
 * @return {Object} { success: boolean }
 */
function deleteRecord(data) {
  try {
    if (!data || !data.sheetName || !data.rowIndex) {
      throw new Error('削除データが不正です');
    }

    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(data.sheetName);

    if (!sheet) {
      throw new Error('シートが見つかりません');
    }

    // 繰越行は削除不可
    const category = sheet.getRange(data.rowIndex, 4).getValue();
    if (category === '繰越') {
      throw new Error('繰越行は削除できません');
    }

    // 行を削除
    sheet.deleteRow(data.rowIndex);

    // 残高を再計算（削除された行の位置から）
    recalcBalances(sheet, data.rowIndex);

    return { success: true };
  } catch (e) {
    throw new Error('レコードの削除に失敗しました: ' + e.message);
  }
}
/**
 * Firebase から受け取ったデータを元に、指定した月のシートを一括更新する
 * @param {Object} data - { sheetName, transactions: [], openingBalance }
 * @return {Object} { success: boolean, count: number }
 */
function bulkSyncTransactions(data) {
  try {
    if (!data || !data.sheetName) {
      throw new Error('同期データが不足しています');
    }

    const ss = getSpreadsheet();
    const sheetName = data.sheetName;
    const transactions = data.transactions || [];
    const openingBalance = data.openingBalance;

    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = createNewMonthlySheet(ss, sheetName);
    }

    // 繰越行（2行目）の更新（開始残高が指定されている場合）
    if (openingBalance !== undefined) {
      // D列(4)が'繰越'であることを前提とする
      sheet.getRange(2, 4).setValue('繰越');
      sheet.getRange(2, 7).setValue(openingBalance);
    }

    // 3行目以降の既存データをクリア（最大行数まで確実にクリア）
    const lastRow = sheet.getLastRow();
    if (lastRow >= 3) {
      sheet.getRange(3, 1, lastRow - 2, HEADERS.length).clearContent();
    }

    // 新しいデータを書き込み
    if (transactions.length > 0) {
      const p = sheetName.split('_');
      const year = parseInt(p[0]) % 100;
      const month = parseInt(p[1]);

      const rows = transactions.map(function(t) {
        return [
          year,
          month,
          t.day,
          t.category,
          t.income || '',
          t.expense || '',
          0, // 差引残高（recalcBalances で計算）
          t.note || ''
        ];
      });

      sheet.getRange(3, 1, rows.length, HEADERS.length).setValues(rows);
    }

    // 全体の残高を再計算
    recalcBalances(sheet, 2);

    // 【追加】カテゴリ（マスタ）の更新（データが含まれている場合）
    if (data.categories && Array.isArray(data.categories)) {
      updateMasterCategories(ss, data.categories);
    }

    return {
      success: true,
      count: transactions.length
    };
  } catch (e) {
    throw new Error('一括同期に失敗しました: ' + e.message);
  }
}

/**
 * マスタ_項目シートを最新のカテゴリリストで更新する
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {Object[]} categories - { name, type } の配列
 */
function updateMasterCategories(ss, categories) {
  let masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    masterSheet = ss.insertSheet(MASTER_SHEET_NAME);
  }

  // ヘッダー以外の既存データをクリア
  const lastRow = masterSheet.getLastRow();
  if (lastRow >= 2) {
    masterSheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  } else {
    // ヘッダーがない場合は作成
    masterSheet.getRange(1, 1, 1, 2).setValues([['項目名', '区分']]).setFontWeight('bold');
  }

  if (categories.length > 0) {
    const rows = categories.map(function(c) {
      return [c.name, c.type];
    });
    masterSheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}
