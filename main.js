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
  const ss = SpreadsheetApp.openById("17N3uNBemPNpi1m0Jcaqf-Q2q4mVFSiXOdgfLo-7KGyM");
  const sheet = ss.getSheets()[0];
  const now = new Date();

  sheet.getRange("A1").setValue("実行時刻:");
  sheet.getRange("B1")
    .setValue(now)
    .setNumberFormat("yyyy/MM/dd HH:mm:ss");
  sheet.setColumnWidth(2, 140);
  sheet.getRange("A1:C3").setBackground("yellow");

  console.log("更新されたスプレッドシートのURL:");
  console.log(ss.getUrl());
}
// writeTimeAndSetBackgroundを実行したあと、A1セルの文字、B1セルの日付型、B1セルの背景色(黄色)を自動検証し、すべて正しければ「TEST PASS」、違えば「TEST FAIL」とログに出力するテスト関数
function testWriteTimeAndSetBackground() {
  writeTimeAndSetBackground();

  const ss = SpreadsheetApp.openById("17N3uNBemPNpi1m0Jcaqf-Q2q4mVFSiXOdgfLo-7KGyM");
  const sheet = ss.getSheets()[0];

  const label = sheet.getRange("A1").getValue();
  const timeValue = sheet.getRange("B1").getValue();
  const backgroundColor = sheet.getRange("B1").getBackground().toLowerCase();

  const isLabelOk = label === "実行時刻:";
  const isTimeOk = Object.prototype.toString.call(timeValue) === "[object Date]"
    && !isNaN(timeValue.getTime());
  const isColorOk = backgroundColor === "#ffff00";

  if (isLabelOk && isTimeOk && isColorOk) {
    console.log("TEST PASS");
  } else {
    console.error("TEST FAIL", { isLabelOk, isTimeOk, isColorOk });
  }
}
