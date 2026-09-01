/**
 * Sales Report Automation Demo
 * daily-formulas.gs
 *
 * 日次管理表の担当者行・日付行・月合計行に
 * 集計式と割合計算式を設定します。
 */

/**
 * 担当者の入力行に計算式を設定します。
 *
 * J～AEは利用者が入力し、それ以外の集計列は自動計算します。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象シート
 * @param {number} row 行番号
 */
function setMemberRowFormulas(sheet, row) {
  const formulas = {
    // 活動結果の合計
    B: `=SUM($J${row}:$AD${row})`,

    // J～Mの合計
    C: `=SUM($J${row}:$M${row})`,

    // N～Qの合計
    D: `=SUM($N${row}:$Q${row})`,

    // 担当者接触数 ÷ 総活動数
    E: `=IFERROR($AE${row}/$B${row},"")`,

    // 担当者接触数 ÷ 有効活動数
    F: `=IFERROR($AE${row}/$AG${row},"")`,

    // 成果数 ÷ 総活動数
    G: `=IFERROR(SUM($J${row}:$Q${row})/$B${row},"")`,

    // 成果数 ÷ 担当者接触数
    H: `=IFERROR(SUM($J${row}:$Q${row})/$AE${row},"")`,

    // 成果数 ÷ 有効活動数
    I: `=IFERROR(SUM($J${row}:$Q${row})/$AG${row},"")`,

    // 無効項目の合計
    AF: `=SUM($Z${row}:$AD${row})`,

    // 総活動数から無効数を差し引く
    AG: `=MAX($B${row}-$AF${row},0)`
  };

  setFormulaMap_(sheet, row, formulas);
}

/**
 * 1日分の集計行に計算式を設定します。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象シート
 * @param {number} dateRow 日付行
 * @param {number[]} memberRows 担当者行
 */
function setDateRowFormulas(
  sheet,
  dateRow,
  memberRows
) {
  setAggregateRowFormulas_(
    sheet,
    dateRow,
    memberRows
  );
}

/**
 * 月全体の合計行に計算式を設定します。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象シート
 * @param {number} row 合計行
 * @param {number[]} allDateRows 日付行
 */
function setDailyMonthTotalFormulas(
  sheet,
  row,
  allDateRows
) {
  setAggregateRowFormulas_(
    sheet,
    row,
    allDateRows
  );
}

/**
 * 担当者ごとの月合計行に計算式を設定します。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象シート
 * @param {number} row 合計行
 * @param {number[]} memberRows 担当者行
 */
function setDailyMemberMonthTotalFormulas(
  sheet,
  row,
  memberRows
) {
  setAggregateRowFormulas_(
    sheet,
    row,
    memberRows
  );
}

/**
 * 複数の参照行を集計し、指定行へ数式を設定します。
 *
 * B～D、J～AEを集計した後、
 * 割合・無効数・有効活動数を計算します。
 */
function setAggregateRowFormulas_(
  sheet,
  targetRow,
  sourceRows
) {
  if (!sourceRows || sourceRows.length === 0) {
    return;
  }

  // B～Dを集計
  for (let column = 2; column <= 4; column++) {
    setRowListSumFormula_(
      sheet,
      targetRow,
      column,
      sourceRows
    );
  }

  // J～AEを集計
  for (let column = 10; column <= 31; column++) {
    setRowListSumFormula_(
      sheet,
      targetRow,
      column,
      sourceRows
    );
  }

  const formulas = {
    E: `=IFERROR($AE${targetRow}/$B${targetRow},"")`,

    F: `=IFERROR($AE${targetRow}/$AG${targetRow},"")`,

    G: `=IFERROR(SUM($J${targetRow}:$Q${targetRow})/$B${targetRow},"")`,

    H: `=IFERROR(SUM($J${targetRow}:$Q${targetRow})/$AE${targetRow},"")`,

    I: `=IFERROR(SUM($J${targetRow}:$Q${targetRow})/$AG${targetRow},"")`,

    AF: `=SUM($Z${targetRow}:$AD${targetRow})`,

    AG: `=MAX($B${targetRow}-$AF${targetRow},0)`
  };

  setFormulaMap_(
    sheet,
    targetRow,
    formulas
  );
}

/**
 * 離れた複数行の同じ列を合計する数式を設定します。
 *
 * 例：
 * =SUM($B9,$B14,$B19)
 */
function setRowListSumFormula_(
  sheet,
  targetRow,
  column,
  sourceRows
) {
  const columnLetter = columnToLetter(column);

  const references = sourceRows.map(
    row => `$${columnLetter}${row}`
  );

  const formula = `=SUM(${references.join(',')})`;

  sheet
    .getRange(targetRow, column)
    .setFormula(formula);
}

/**
 * 列記号と数式の組み合わせを対象行へ設定します。
 */
function setFormulaMap_(sheet, row, formulaMap) {
  Object.entries(formulaMap).forEach(
    ([columnLetter, formula]) => {
      sheet
        .getRange(`${columnLetter}${row}`)
        .setFormula(formula);
    }
  );
}
