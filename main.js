function writeCurrentTime() {
  // 1. 新しいスプレッドシートを作成
  const ss = SpreadsheetApp.create("時間記録テスト");
  const sheet = ss.getActiveSheet();

  // 2. 現在の時刻を取得して書き込む
  const now = new Date();
  sheet.getRange("A1").setValue("実行時刻:");
  sheet.getRange("B1").setValue(now);

  // 3. ログにスプレッドシートのURLを表示
  console.log("作成されたスプレッドシートのURL:");
  console.log(ss.getUrl());
}
// スプレッドシートのA1からC3までの背景色を黄色にする関数
function setBackgroundColor() {
    const ss = SpreadsheetApp.create("背景色変更テスト");
    const sheet = ss.getActiveSheet();
    sheet.getRange("A1:C3").setBackground("yellow");
}
// 現在の時刻を書き込み、さらにA1からC3までの背景色を黄色にする関数
function writeTimeAndSetBackground() {
  const ss = SpreadsheetApp.create("合体テスト");
  const sheet = ss.getActiveSheet();

  // 1. 時間を書き込む
  const now = new Date();
  sheet.getRange("A1").setValue("実行時刻:");
  sheet.getRange("B1").setValue(now);
  
  // ★追加：B1セルの表示形式を「秒」まで表示するように設定
  sheet.getRange("B1").setNumberFormat("yyyy/MM/dd HH:mm:ss");
  
  // B列（2列目）の幅を140ピクセルに広げる
  sheet.setColumnWidth(2, 140);

  // 2. 背景色を黄色にする
  sheet.getRange("A1:C3").setBackground("yellow");

  console.log("作成された合体スプレッドシートのURL:");
  console.log(ss.getUrl());
}