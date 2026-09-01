/**
 * Sales Report Automation Demo
 * menu.gs
 *
 * スプレッドシートのカスタムメニューと
 * 月次管理表の作成処理を管理します。
 */

/**
 * スプレッドシートを開いたときに実行メニューを追加します。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('営業管理 自動化')
    .addItem('月次管理表を作成', 'createMonthlyReports')
    .addToUi();
}

/**
 * 年月を入力し、日次シートと週次シートを作成します。
 */
function createMonthlyReports() {
  const ui = SpreadsheetApp.getUi();

  const yearResponse = ui.prompt(
    '月次管理表の作成',
    '作成する年を入力してください（例：2026）',
    ui.ButtonSet.OK_CANCEL
  );

  if (yearResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const year = Number(yearResponse.getResponseText().trim());

  if (!Number.isInteger(year) || year < 2020 || year > 2100) {
    ui.alert('年は2020～2100の整数で入力してください。');
    return;
  }

  const monthResponse = ui.prompt(
    '月次管理表の作成',
    '作成する月を入力してください（例：9）',
    ui.ButtonSet.OK_CANCEL
  );

  if (monthResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const monthNumber = Number(monthResponse.getResponseText().trim());

  if (
    !Number.isInteger(monthNumber) ||
    monthNumber < 1 ||
    monthNumber > 12
  ) {
    ui.alert('月は1～12の整数で入力してください。');
    return;
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const dailySheetName = `${year}年${monthNumber}月`;
  const weeklySheetName = `${year}年${monthNumber}月（週次）`;

  const existingDailySheet =
    spreadsheet.getSheetByName(dailySheetName);

  const existingWeeklySheet =
    spreadsheet.getSheetByName(weeklySheetName);

  if (existingDailySheet || existingWeeklySheet) {
    const overwriteResponse = ui.alert(
      '上書き確認',
      '同じ年月の管理表が存在します。内容を初期化して作り直しますか？',
      ui.ButtonSet.YES_NO
    );

    if (overwriteResponse !== ui.Button.YES) {
      return;
    }
  }

  try {
    const dailySheet = prepareSheet_(
      spreadsheet,
      dailySheetName
    );

    const weeklySheet = prepareSheet_(
      spreadsheet,
      weeklySheetName
    );

    const targetMonthIndex = monthNumber - 1;

    const result = buildDailySheet(
      dailySheet,
      year,
      targetMonthIndex
    );

    buildWeeklySheet(
      weeklySheet,
      dailySheetName,
      year,
      targetMonthIndex,
      result.weeks
    );

    spreadsheet.setActiveSheet(dailySheet);

    ui.alert(
      '作成完了',
      `「${dailySheetName}」と「${weeklySheetName}」を作成しました。`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    console.error(error);

    ui.alert(
      'エラー',
      `管理表の作成中にエラーが発生しました。\n\n${error.message}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 既存シートは初期化し、存在しない場合は新規作成します。
 *
 * シート自体を削除しないため、最低1枚しかない場合でも
 * 安全に実行できます。
 */
function prepareSheet_(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    return sheet;
  }

  const dataRange = sheet.getDataRange();

  try {
    dataRange.breakApart();
  } catch (error) {
    console.warn(`結合解除をスキップしました: ${error.message}`);
  }

  sheet.clear();
  sheet.clearConditionalFormatRules();

  sheet.getCharts().forEach(chart => {
    sheet.removeChart(chart);
  });

  return sheet;
}
