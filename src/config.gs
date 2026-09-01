/**
 * Sales Report Automation Demo
 * config.gs
 *
 * 公開用デモの設定値とヘッダーを管理します。
 * 実在する企業・顧客・従業員の情報は含まれていません。
 */

const APP_CONFIG = {
  members: [
    '担当者A',
    '担当者B',
    '担当者C',
    '担当者D'
  ],

  totalDailyColumns: 33,
  totalWeeklyColumns: 15,

  dailyStartRow: 8,
  frozenDailyColumns: 9,
  frozenWeeklyColumns: 5,

  weekColors: [
    '#dce8f8',
    '#b8d4f0',
    '#7fb3e8',
    '#4d94d8',
    '#1a6fc4'
  ],

  sectionColor: '#1a56db',
  headerColor: '#ffff00',
  subHeaderColor: '#ffd966',
  white: '#ffffff'
};

/**
 * 日次管理表の列名です。
 *
 * A～I：自動集計
 * J～AD：活動結果の入力
 * AE：担当者接触数の入力
 * AF～AG：自動集計
 */
const DAILY_HEADERS = [
  '日付・担当者',                 // A
  '総活動数',                     // B
  'フォロー獲得数',               // C
  '商談化数',                     // D
  '全体接触率',                   // E
  '有効接触率',                   // F
  '全体成果率',                   // G
  '接触後成果率',                 // H
  '有効活動成果率',               // I
  '資料送付',                     // J
  '再提案',                       // K
  '要フォロー',                   // L
  '商談済み',                     // M
  '要望あり',                     // N
  '社内連携',                     // O
  'オンライン商談',               // P
  '訪問商談',                     // Q
  '再連絡',                       // R
  '担当者不在',                   // S
  '検討中',                       // T
  '受付終了',                     // U
  'ニーズなし',                   // V
  '本部管理',                     // W
  '対象外',                       // X
  '折り返し待ち',                 // Y
  'その他',                       // Z
  '不通',                         // AA
  '番号無効',                     // AB
  '閉業',                         // AC
  '自動音声',                     // AD
  '担当者接触数',                 // AE
  '無効数',                       // AF
  '有効活動数'                    // AG
];

/**
 * 週次管理表の列名です。
 */
const WEEKLY_HEADERS = [
  '期間・担当者',
  '目標活動数',
  'フォロー目標数',
  '商談化目標数',
  '合計目標数',
  '総活動数',
  '有効活動数',
  '担当者接触数',
  '全体接触率',
  '有効接触率',
  'フォロー獲得数',
  '商談化数',
  '全体成果率',
  '接触後成果率',
  '有効活動成果率'
];
