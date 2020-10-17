/**
 * daily_event_listの予定を作成する
 * 以下のfunctionに対してトリガーを設定して定期実行する
 */
function createScheduleAuto() {

  /**********************
  テスト投稿の時はTrueにする
  **********************/
  var test = false;

  // テスト投稿に使用するカレンダーIDをセット
  const gAccount = "★★";  // ★★ここに連携するカレンダーのIDをいれる
  /**********************
  テスト用の設定ここまで
  **********************/

  // スプレッドシートとシートを取得
  // スプレッドシートのURL  https://docs.google.com/spreadsheets/d/{ss_ID}/edit#gid=0
  var ss_id = "★★★";  //★★★ここにss_IDを入力
  var sheet = SpreadsheetApp.openById(ss_id).getSheetByName("daily_event_list");

  // 読み取り範囲（表の始まり行と終わり列）を指定
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
      var startDate = new Date(contents[i][startDateCellNum]);  //イベント開始日
      var startTime = contents[i][startTimeCellNum];  //イベント開始時刻
      var endDate = new Date(contents[i][endDateCellNum]);  //イベント終了日
      var endTime = contents[i][endTimeCellNum];  //イベント終了時刻
      var title = contents[i][titleCellNum];  //イベントタイトル
      // 場所と詳細をセット
      var options = {description: contents[i][descriptionCellNum]}; //説明


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

} // createScheduleAutoここまで
