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