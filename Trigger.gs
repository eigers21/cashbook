/**
 * 現金出納帳 — 月次シート自動生成トリガー
 * 毎月1日 0時〜1時に自動実行される
 */

/**
 * 当月の月次シートを自動生成する
 * トリガーから毎月1日に呼び出される
 * シートが既に存在する場合はスキップする
 */
function createMonthlySheet() {
  const ss = getSpreadsheet();
  const sheetName = getCurrentSheetName();

  // 既に存在する場合はスキップ
  if (ss.getSheetByName(sheetName)) {
    Logger.log('シート "' + sheetName + '" は既に存在します。スキップします。');
    return;
  }

  // 新規シートを作成（Code.gs のヘルパー関数を利用）
  const sheet = createNewMonthlySheet(ss, sheetName);
  Logger.log('シート "' + sheetName + '" を作成しました。');
}

/**
 * 月次シート自動生成トリガーを設定する
 * ※ この関数を手動で一度だけ実行してください
 * 既存の同名トリガーがある場合は削除してから再設定します
 */
function setupMonthlyTrigger() {
  // 既存の createMonthlySheet トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'createMonthlySheet') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('既存のトリガーを削除しました。');
    }
  });

  // 新しいトリガーを設定: 毎月1日 0時〜1時に実行
  ScriptApp.newTrigger('createMonthlySheet')
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .create();

  Logger.log('月次シート自動生成トリガーを設定しました（毎月1日 0:00〜1:00）');
}
