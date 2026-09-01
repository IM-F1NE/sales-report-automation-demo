/**
 * Sales Report Automation Demo
 * date-utils.gs
 *
 * 日付・営業日・週区分・列記号に関する共通関数です。
 */

/**
 * 指定月の平日を取得します。
 *
 * @param {number} year 年
 * @param {number} monthIndex 月。0が1月、11が12月
 * @return {Date[]} 土日を除いた日付の配列
 */
function getBusinessDays(year, monthIndex) {
  const businessDays = [];
  const currentDate = new Date(year, monthIndex, 1);

  while (currentDate.getMonth() === monthIndex) {
    const dayOfWeek = currentDate.getDay();
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

    if (!isWeekend) {
      businessDays.push(new Date(currentDate));
    }

    currentDate.setDate(currentDate.getDate() + 1);
  }

  return businessDays;
}

/**
 * 日付を月曜日始まりの週に分けます。
 *
 * 月をまたぐ場合でも、対象月内の日付だけでグループ化します。
 *
 * @param {Date[]} days 日付の配列
 * @return {Date[][]} 週ごとの日付配列
 */
function groupByWeek(days) {
  const groupedWeeks = [];
  let currentWeek = [];
  let previousWeekKey = '';

  days.forEach(day => {
    const weekKey = getMondayKey_(day);

    if (previousWeekKey && weekKey !== previousWeekKey) {
      groupedWeeks.push(currentWeek);
      currentWeek = [];
    }

    currentWeek.push(new Date(day));
    previousWeekKey = weekKey;
  });

  if (currentWeek.length > 0) {
    groupedWeeks.push(currentWeek);
  }

  return groupedWeeks;
}

/**
 * 指定日の週の月曜日を、比較用文字列として返します。
 *
 * @param {Date} date 対象日
 * @return {string} yyyy-MM-dd形式の週識別子
 */
function getMondayKey_(date) {
  const monday = new Date(date);

  monday.setHours(0, 0, 0, 0);

  const dayOfWeek = monday.getDay();
  const differenceFromMonday =
    dayOfWeek === 0 ? -6 : 1 - dayOfWeek;

  monday.setDate(monday.getDate() + differenceFromMonday);

  return Utilities.formatDate(
    monday,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}

/**
 * 列番号をスプレッドシートの列記号に変換します。
 *
 * 例：
 * 1 → A
 * 26 → Z
 * 27 → AA
 *
 * @param {number} columnNumber 列番号
 * @return {string} 列記号
 */
function columnToLetter(columnNumber) {
  let result = '';
  let currentNumber = columnNumber;

  while (currentNumber > 0) {
    const remainder = (currentNumber - 1) % 26;

    result =
      String.fromCharCode(65 + remainder) +
      result;

    currentNumber =
      Math.floor((currentNumber - 1) / 26);
  }

  return result;
}

/**
 * シート名を数式内で安全に使用できる形へ変換します。
 *
 * @param {string} sheetName シート名
 * @return {string} 数式用のシート名
 */
function escapeSheetName_(sheetName) {
  return String(sheetName).replace(/'/g, "''");
}
