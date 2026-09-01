/**
 * Sales Report Automation Demo
 * weekly-sheet.gs
 *
 * 日次管理表を参照し、
 * 週次・担当者別・月次の集計表を生成します。
 */

/**
 * 週次管理表を構築します。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 週次シート
 * @param {string} dailySheetName 日次シート名
 * @param {number} year 年
 * @param {number} monthIndex 月。0が1月、11が12月
 * @param {Date[][]} weeks 週ごとの営業日
 */
function buildWeeklySheet(
  sheet,
  dailySheetName,
  year,
  monthIndex,
  weeks
) {
  const totalColumns =
    APP_CONFIG.totalWeeklyColumns;

  const monthNumber = monthIndex + 1;

  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(
    APP_CONFIG.frozenWeeklyColumns
  );

  buildWeeklyHeader_(
    sheet,
    year,
    monthNumber,
    totalColumns
  );

  // 月合計行
  sheet
    .getRange(3, 1)
    .setValue(`${monthNumber}月合計`);

  sheet
    .getRange(3, 1, 1, totalColumns)
    .setBackground(APP_CONFIG.white)
    .setFontWeight('bold');

  // 担当者別の月合計行
  APP_CONFIG.members.forEach((member, index) => {
    const row = 4 + index;

    sheet
      .getRange(row, 1)
      .setValue(member);

    sheet
      .getRange(row, 1, 1, totalColumns)
      .setBackground(APP_CONFIG.white)
      .setFontWeight('bold');
  });

  const dailyRowsByWeek =
    getDailyMemberRowsByWeek_(weeks);

  const weekMemberStartRows = [];

  let currentRow = APP_CONFIG.dailyStartRow;

  weeks.forEach((week, weekIndex) => {
    const weekColor =
      APP_CONFIG.weekColors[weekIndex] ||
      APP_CONFIG.weekColors[
        APP_CONFIG.weekColors.length - 1
      ];

    const firstDay = week[0];
    const lastDay = week[week.length - 1];

    const periodLabel =
      `第${weekIndex + 1}週 ` +
      `(${monthNumber}/${firstDay.getDate()}` +
      `～${monthNumber}/${lastDay.getDate()})`;

    // 週見出し
    sheet
      .getRange(
        currentRow,
        1,
        1,
        totalColumns
      )
      .setBackground(APP_CONFIG.sectionColor)
      .setFontColor(APP_CONFIG.white)
      .setFontWeight('bold');

    sheet
      .getRange(currentRow, 1)
      .setValue(periodLabel);

    currentRow++;

    weekMemberStartRows.push(currentRow);

    APP_CONFIG.members.forEach(member => {
      const memberRow = currentRow;

      sheet
        .getRange(memberRow, 1)
        .setValue(member)
        .setBackground(weekColor);

      // B～Dは目標値の入力欄
      formatWeeklyTargetCells_(
        sheet,
        memberRow
      );

      const dailyRows =
        dailyRowsByWeek[weekIndex][member];

      setWeeklyMemberFormulas(
        sheet,
        memberRow,
        dailySheetName,
        dailyRows
      );

      currentRow++;
    });
  });

  // 担当者別の月合計
  APP_CONFIG.members.forEach(
    (member, memberIndex) => {
      const summaryRow = 4 + memberIndex;

      const weeklyRows =
        weekMemberStartRows.map(
          startRow => startRow + memberIndex
        );

      setWeeklyMemberMonthTotal(
        sheet,
        summaryRow,
        weeklyRows
      );
    }
  );

  // 月全体の合計
  const allWeeklyMemberRows = [];

  weekMemberStartRows.forEach(startRow => {
    APP_CONFIG.members.forEach(
      (_, memberIndex) => {
        allWeeklyMemberRows.push(
          startRow + memberIndex
        );
      }
    );
  });

  setWeeklyMonthTotal(
    sheet,
    3,
    allWeeklyMemberRows
  );

  const lastRow = currentRow - 1;

  // 割合の表示形式
  sheet
    .getRange(3, 9, lastRow - 2, 2)
    .setNumberFormat('0.0%');

  sheet
    .getRange(3, 13, lastRow - 2, 3)
    .setNumberFormat('0.0%');

  sheet
    .getRange(1, 1, lastRow, totalColumns)
    .setVerticalAlignment('middle');

  sheet.setColumnWidth(1, 210);

  for (
    let column = 2;
    column <= totalColumns;
    column++
  ) {
    sheet.setColumnWidth(column, 115);
  }

  sheet.setRowHeight(1, 30);
  sheet.setRowHeight(2, 65);
}

/**
 * 週次管理表のタイトルと見出しを作成します。
 */
function buildWeeklyHeader_(
  sheet,
  year,
  monthNumber,
  totalColumns
) {
  sheet
    .getRange(1, 1)
    .setValue(
      `${year}年${monthNumber}月 営業活動管理表（週次）`
    )
    .setFontSize(14)
    .setFontWeight('bold');

  sheet
    .getRange(1, 1, 1, totalColumns)
    .setBackground(APP_CONFIG.white);

  sheet
    .getRange(2, 1, 1, WEEKLY_HEADERS.length)
    .setValues([WEEKLY_HEADERS])
    .setBackground(APP_CONFIG.headerColor)
    .setFontWeight('bold')
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // 目標欄
  sheet
    .getRange(2, 2, 1, 4)
    .setBackground(APP_CONFIG.subHeaderColor);
}

/**
 * 週ごとの目標入力欄を設定します。
 */
function formatWeeklyTargetCells_(sheet, row) {
  const targetRange =
    sheet.getRange(row, 2, 1, 3);

  targetRange.setBackground('#fffdf2');

  const numberRule = SpreadsheetApp
    .newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(false)
    .setHelpText('0以上の目標値を入力してください。')
    .build();

  targetRange.setDataValidation(numberRule);
}

/**
 * 日次管理表の担当者行番号を、
 * 週ごと・担当者ごとに整理します。
 */
function getDailyMemberRowsByWeek_(weeks) {
  const result = {};
  let currentRow = APP_CONFIG.dailyStartRow;

  weeks.forEach((week, weekIndex) => {
    result[weekIndex] = {};

    APP_CONFIG.members.forEach(member => {
      result[weekIndex][member] = [];
    });

    // 週見出し行
    currentRow++;

    week.forEach(() => {
      // 日付合計行
      currentRow++;

      APP_CONFIG.members.forEach(member => {
        result[weekIndex][member].push(
          currentRow
        );

        currentRow++;
      });
    });
  });

  return result;
}
