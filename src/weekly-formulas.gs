/**
 * Sales Report Automation Demo
 * weekly-formulas.gs
 *
 * 日次管理表を参照し、
 * 週次・担当者別・月次の数値を集計します。
 */

/**
 * 週次シートの担当者行に数式を設定します。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 週次シート
 * @param {number} row 設定対象行
 * @param {string} dailySheetName 日次シート名
 * @param {number[]} dailyRows 日次シートの担当者行番号
 */
function setWeeklyMemberFormulas(
  sheet,
  row,
  dailySheetName,
  dailyRows
) {
  if (!dailyRows || dailyRows.length === 0) {
    return;
  }

  const escapedSheetName =
    escapeSheetName_(dailySheetName);

  // E：フォロー目標＋商談化目標
  sheet
    .getRange(row, 5)
    .setFormula(
      `=IF(COUNTA($C${row}:$D${row})=0,"",SUM($C${row}:$D${row}))`
    );

  // 日次シートから実績値を集計
  setCrossSheetRowSum_(
    sheet,
    row,
    6,
    escapedSheetName,
    2,
    dailyRows
  );

  setCrossSheetRowSum_(
    sheet,
    row,
    7,
    escapedSheetName,
    33,
    dailyRows
  );

  setCrossSheetRowSum_(
    sheet,
    row,
    8,
    escapedSheetName,
    31,
    dailyRows
  );

  setCrossSheetRowSum_(
    sheet,
    row,
    11,
    escapedSheetName,
    3,
    dailyRows
  );

  setCrossSheetRowSum_(
    sheet,
    row,
    12,
    escapedSheetName,
    4,
    dailyRows
  );

  setWeeklyRateFormulas_(sheet, row);
}

/**
 * 担当者ごとの月合計を計算します。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 週次シート
 * @param {number} row 月合計行
 * @param {number[]} weekRows 担当者の各週の行
 */
function setWeeklyMemberMonthTotal(
  sheet,
  row,
  weekRows
) {
  setWeeklySummaryFormulas_(
    sheet,
    row,
    weekRows
  );
}

/**
 * 月全体の合計を計算します。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 週次シート
 * @param {number} row 月合計行
 * @param {number[]} allRows 全担当者の各週の行
 */
function setWeeklyMonthTotal(
  sheet,
  row,
  allRows
) {
  setWeeklySummaryFormulas_(
    sheet,
    row,
    allRows
  );
}

/**
 * 週次シート内の複数行を集計します。
 */
function setWeeklySummaryFormulas_(
  sheet,
  targetRow,
  sourceRows
) {
  if (!sourceRows || sourceRows.length === 0) {
    return;
  }

  // B～D：目標値
  for (let column = 2; column <= 4; column++) {
    setSameSheetRowSum_(
      sheet,
      targetRow,
      column,
      sourceRows
    );
  }

  // F～H：活動実績
  for (let column = 6; column <= 8; column++) {
    setSameSheetRowSum_(
      sheet,
      targetRow,
      column,
      sourceRows
    );
  }

  // K～L：成果実績
  for (let column = 11; column <= 12; column++) {
    setSameSheetRowSum_(
      sheet,
      targetRow,
      column,
      sourceRows
    );
  }

  // E：フォロー目標＋商談化目標
  sheet
    .getRange(targetRow, 5)
    .setFormula(
      `=IF(COUNTA($C${targetRow}:$D${targetRow})=0,"",SUM($C${targetRow}:$D${targetRow}))`
    );

  setWeeklyRateFormulas_(
    sheet,
    targetRow
  );
}

/**
 * 接触率・成果率を計算します。
 */
function setWeeklyRateFormulas_(sheet, row) {
  const formulas = {
    // 担当者接触数 ÷ 総活動数
    I: `=IFERROR($H${row}/$F${row},"")`,

    // 担当者接触数 ÷ 有効活動数
    J: `=IFERROR($H${row}/$G${row},"")`,

    // 成果数 ÷ 総活動数
    M: `=IFERROR(SUM($K${row}:$L${row})/$F${row},"")`,

    // 成果数 ÷ 担当者接触数
    N: `=IFERROR(SUM($K${row}:$L${row})/$H${row},"")`,

    // 成果数 ÷ 有効活動数
    O: `=IFERROR(SUM($K${row}:$L${row})/$G${row},"")`
  };

  Object.entries(formulas).forEach(
    ([columnLetter, formula]) => {
      sheet
        .getRange(`${columnLetter}${row}`)
        .setFormula(formula);
    }
  );
}

/**
 * 別シートにある複数行の同じ列を合計します。
 */
function setCrossSheetRowSum_(
  targetSheet,
  targetRow,
  targetColumn,
  sourceSheetName,
  sourceColumn,
  sourceRows
) {
  const sourceColumnLetter =
    columnToLetter(sourceColumn);

  const references = sourceRows.map(
    sourceRow =>
      `'${sourceSheetName}'!$${sourceColumnLetter}${sourceRow}`
  );

  const formula =
    `=SUM(${references.join(',')})`;

  targetSheet
    .getRange(targetRow, targetColumn)
    .setFormula(formula);
}

/**
 * 同じシート内にある複数行の同じ列を合計します。
 */
function setSameSheetRowSum_(
  sheet,
  targetRow,
  column,
  sourceRows
) {
  const columnLetter =
    columnToLetter(column);

  const references = sourceRows.map(
    sourceRow => `$${columnLetter}${sourceRow}`
  );

  const formula =
    `=SUM(${references.join(',')})`;

  sheet
    .getRange(targetRow, column)
    .setFormula(formula);
}
