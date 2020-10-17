/**
 * 予定を作成する
 */
function createSchedule() {

  /**********************
  テスト投稿の時はTrueにする
  **********************/
  var test = false;

  // テスト投稿に使用するカレンダーIDをセット
  const gAccount = "★★";  // ★★ここに連携するカレンダーのIDをいれる
  /**********************
  テスト用の設定ここまで
  **********************/

  // 開いているシートを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // 開いているシートのシート名を取得
  var sheetName = sheet.getSheetName();

  console.log({sheetName})

  // 開いているシートに応じて、読み取り範囲と列を指定
  try {
    // event_listシートを対象に実行する場合
    if(sheetName == "event_list"){
      // 読み取り範囲（表の始まり行と終わり列）
      const topRow = 42;
      const lastCol = 14;

      // A列=0始まりで列を指定
      const titleCellNum = 1;
      const startDateCellNum = 2;
      const startTimeCellNum = 3;
      const endDateCellNum = 4;
      const endTimeCellNum = 5;
      const descriptionCellNum = 6;
      const calendarIDCellNum = 8;
      const calStatusCellNum = 9;
      const startAlertTimeCellNum = 10;
      const startAlertStatusCellNum = 11;
      const endAlertTimeCellNum = 12;
      const endAlertStatusCellNum = 13;

    // daily_event_listシートを対象に実行する場合
    } elseif (sheetName == "daily_event_list"){
      // 読み取り範囲（表の始まり行と終わり列）
      const topRow = 42;
      const lastCol = 15;

      // A列=0始まりで列を指定
      const cycleCellNum = 1;
      const titleCellNum = 2;
      const startDateCellNum = 3;
      const startTimeCellNum = 4;
      const endDateCellNum = 5;
      const endTimeCellNum = 6;
      const descriptionCellNum = 7;
      const calendarIDCellNum = 9;
      const calStatusCellNum = 10;
      const startAlertTimeCellNum = 11;
      const startAlertStatusCellNum = 12;
      const endAlertTimeCellNum = 13;
      const endAlertStatusCellNum = 14;

    // event_listまたはdaily_event_list以外のシートで実行しようとした場合はエラーで終了
    } else {
      throw new error("event_listまたはdaily_event_listを表示した状態で実行してください")
    }
  } catch(e) {
    // エラーの場合はブラウザへ表示
    Browser.msgBox(e);
  }

  // リストの最終行を取得
  var lastRow = sheet.getLastRow();

  //予定の一覧を取得
  var contents = sheet.getRange(topRow, 1, sheet.getLastRow(), lastCol).getValues();

  //順に予定を作成（今回は正しい値が来ることを想定）
  for (i = 0; i < lastRow - topRow+2; i++) {

      //「OK」または開始日が空の場合は飛ばす
      var status = contents[i][calStatusCellNum];
      if (
        status == "OK" ||
        contents[i][startDateCellNum] == ""
      ) {
        continue;
      }

      // 予定を追加するgoogleカレンダーを取得
      // ※GoogleAppsScriptのプロジェクトオーナーのマイカレンダーに対象カレンダーが追加されていないと取得できない（取得するカレンダーへの参照権限が必要）
      // テスト以外では指定されたカレンダーを取得
      if (test == false) {
      var calendar = CalendarApp.getCalendarById(contents[i][calendarIDCellNum].toString());
      // テストの時はテスト用に指定されたカレンダーを取得
      } else {
      var calendar = CalendarApp.getCalendarById(gAccount);
      }

      console.log({calendar})


      // 予定投稿に必要な値をセット 日時はフォーマットして保持
      var startDate = new Date(contents[i][startDateCellNum]);
      var startTime = contents[i][startTimeCellNum];
      var endDate = new Date(contents[i][endDateCellNum]);
      var endTime = contents[i][endTimeCellNum];
      var title = contents[i][titleCellNum];
      // 場所と詳細をセット
      var options = {description: contents[i][descriptionCellNum]};


      try {
          // 開始日時をフォーマット
        var startDateTime = new Date(startDate);
        startDateTime.setHours(startTime.getHours())
        startDateTime.setMinutes(startTime.getMinutes());
          // 終了日時をフォーマット
        var endDateTime = new Date(endDate);
        endDateTime.setHours(endTime.getHours())
        endDateTime.setMinutes(endTime.getMinutes());

          console.log(startDateTime, endDateTime)

          // 予定を作成
          calendar.createEvent(
            title,
            startDateTime,
            endDateTime,
            options
          );

        //無事に予定が作成されたらステータスを「OK」にする
        sheet.getRange(topRow + i, calStatusCellNum + 1).setValue("OK");

      // エラーの場合はログを出力する）
      } catch(e) {
        Logger.log(e);
      }

 } //forここまで

  // 処理が全て完了したら、ブラウザへ通知
  Browser.msgBox("完了");
}
