/**
 * Sales Report Automation Demo
 * daily-sheet.gs
 *
 * 日次管理表のレイアウト、営業日、担当者行を生成します。
 */

/**
 * 日次管理表を構築します。
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象シート
 * @param {number} year 年
 * @param {number} monthIndex 月。0が1月、11が12月
 * @return {{
 *   weeks: Date[][],
 *   memberRowMap: Object,
 *   allDateRows: number[]
 * }}
 */
function buildDailySheet(sheet, year, monthIndex) {
  const totalColumns = APP_CONFIG.totalDailyColumns;
  const businessDays = getBusinessDays(year, monthIndex);
  const weeks = groupByWeek(businessDays);

  if (businessDays.length === 0) {
    throw new Error('対象月の営業日を取得できませんでした。');
  }

  sheet.setFrozenRows(2);
  sheet.setFrozenColumns(
    APP_CONFIG.frozenDailyColumns
  );

  buildDailyHeader_(
    sheet,
    year,
    monthIndex,
    totalColumns
  );

  let currentRow = APP_CONFIG.dailyStartRow;
  const allDateRows = [];

  const memberRowMap = {};

  APP_CONFIG.members.forEach(member => {
    memberRowMap[member] = [];
  });

  weeks.forEach((week, weekIndex) => {
    const weekColor =
      APP_CONFIG.weekColors[weekIndex] ||
      APP_CONFIG.weekColors[
        APP_CONFIG.weekColors.length - 1
      ];

    // 週見出し
    const weekLabelRow = currentRow;

    sheet
      .getRange(
        weekLabelRow,
        1,
        1,
        totalColumns
      )
      .setBackground(APP_CONFIG.sectionColor)
      .setFontColor(APP_CONFIG.white)
      .setFontWeight('bold');

    sheet
      .getRange(weekLabelRow, 1)
      .setValue(`第${weekIndex + 1}週`);

    currentRow++;

    week.forEach(day => {
      const dateRow = currentRow;
      allDateRows.push(dateRow);

      sheet
        .getRange(dateRow, 1)
        .setValue(new Date(day))
        .setNumberFormat('yyyy/mm/dd')
        .setBackground(weekColor)
        .setFontWeight('bold');

      currentRow++;

      const memberRows = [];

      APP_CONFIG.members.forEach(member => {
        const memberRow = currentRow;

        memberRows.push(memberRow);
        memberRowMap[member].push(memberRow);

        sheet
          .getRange(memberRow, 1)
          .setValue(member)
          .setBackground(weekColor);

        formatDailyInputCells_(sheet, memberRow);
        setMemberRowFormulas(sheet, memberRow);

        currentRow++;
      });

      setDateRowFormulas(
        sheet,
        dateRow,
        memberRows
      );
    });
  });

  // 月全体の合計行
  sheet.getRange(3, 1).setValue(
    `${monthIndex + 1}月合計`
  );

  sheet
    .getRange(3, 1, 1, totalColumns)
    .setBackground(APP_CONFIG.white)
    .setFontWeight('bold');

  setDailyMonthTotalFormulas(
    sheet,
    3,
    allDateRows
  );

  // 担当者ごとの月合計行
  APP_CONFIG.members.forEach((member, index) => {
    const row = 4 + index;

    sheet.getRange(row, 1).setValue(member);

    sheet
      .getRange(row, 1, 1, totalColumns)
      .setBackground(APP_CONFIG.white)
      .setFontWeight('bold');

    setDailyMemberMonthTotalFormulas(
      sheet,
      row,
      memberRowMap[member]
    );
  });

  const lastRow = currentRow - 1;

  // 割合をパーセント表示
  sheet
    .getRange(3, 5, lastRow - 2, 5)
    .setNumberFormat('0.0%');

  // 表全体の基本設定
  sheet
    .getRange(1, 1, lastRow, totalColumns)
    .setVerticalAlignment('middle');

  sheet.setColumnWidth(1, 130);

  for (
    let column = 2;
    column <= totalColumns;
    column++
  ) {
    sheet.setColumnWidth(column, 95);
  }

  sheet.setRowHeight(1, 30);
  sheet.setRowHeight(2, 60);

  return {
    weeks,
    memberRowMap,
    allDateRows
  };
}

/**
 * 日次管理表のタイトルと列見出しを作成します。
 */
function buildDailyHeader_(
  sheet,
  year,
  monthIndex,
  totalColumns
) {
  sheet
    .getRange(1, 1)
    .setValue(
      `${year}年${monthIndex + 1}月 営業活動管理表`
    )
    .setFontSize(14)
    .setFontWeight('bold');

  sheet
    .getRange(1, 1, 1, totalColumns)
    .setBackground(APP_CONFIG.white);

  sheet
    .getRange(2, 1, 1, DAILY_HEADERS.length)
    .setValues([DAILY_HEADERS])
    .setBackground(APP_CONFIG.headerColor)
    .setFontWeight('bold')
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // フォロー・商談に関する入力項目
  sheet
    .getRange(2, 10, 1, 4)
    .setBackground('#ffe599');

  sheet
    .getRange(2, 14, 1, 4)
    .setBackground(APP_CONFIG.subHeaderColor);

  // 担当者接触数
  sheet
    .getRange(2, 31)
    .setBackground('#d9ead3');
}

/**
 * 担当者が入力するセルの見た目と入力規則を設定します。
 */
function formatDailyInputCells_(sheet, row) {
  // J～AEを入力欄として表示
  const inputRange = sheet.getRange(row, 10, 1, 22);

  inputRange.setBackground('#fffdf2');

  const numberRule = SpreadsheetApp
    .newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .setAllowInvalid(false)
    .setHelpText('0以上の数値を入力してください。')
    .build();

  inputRange.setDataValidation(numberRule);
}
