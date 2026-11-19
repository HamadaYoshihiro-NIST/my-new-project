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
  // ★修正：createをやめ、取得したIDを指定して既存ファイルを開く
  const ss = SpreadsheetApp.openById("17N3uNBemPNpi1m0Jcaqf-Q2q4mVFSiXOdgfLo-7KGyM");
  const sheet = ss.getActiveSheet();

  // 2. 時間を書き込む
  const now = new Date();
  sheet.getRange("A1").setValue("実行時刻:");
  sheet.getRange("B1").setValue(now);
  
  // B1セルの表示形式と幅の調整
  sheet.getRange("B1").setNumberFormat("yyyy/MM/dd HH:mm:ss");
  sheet.setColumnWidth(2, 140);

  // 3. 背景色を黄色にする
  sheet.getRange("A1:C3").setBackground("yellow");

  console.log("更新されたスプレッドシートのURL:");
  console.log(ss.getUrl());
}