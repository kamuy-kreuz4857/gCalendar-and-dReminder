/**
 * daily_event_listのイベント開始日・終了日の更新
 * 開始通知Statusおよび終了通知Statusをリセット
 * ※イベント終了時刻と終了通知時刻が同一の場合、終了通知StatusはOKのまま
 * createScheduleAutoを使用している場合、予定追加Statusをリセット
 *
 * 以下のfunctionに対してトリガーを設定して定期実行する
 */
function resetCycleReminder() {

  /**************************************
  createScheduleAutoを使用している場合はtrue
  **************************************/
  var createScheduleAutoIsActive = false;
  /**************************************
  createScheduleAuto用の設定ここまで
  **************************************/

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

 // スプレッドシートとシートを取得
 // スプレッドシートのURL  https://docs.google.com/spreadsheets/d/{ss_ID}/edit#gid=0
 var ss_id = "★★★";  //★★★ここにss_IDを入力
 var sheet = SpreadsheetApp.openById(ss_id).getSheetByName("daily_event_list");

  // 予定の最終行を取得
  var lastRow = sheet.getLastRow();

  //予定の一覧を取得
  var contents = sheet.getRange(topRow, 1, sheet.getLastRow(), lastCol).getValues();

  //現在日時を取得
  var currentDateTime = new Date();

  /***********************
  /順にリマインダーの開始日を更新/
  ************************/
  for (i = 0; i < lastRow - topRow+2; i++) {

      //開始リマインダーのステータスを取得
      var StartAlertStatus = contents[i][startAlertStatusCellNum];

      //「OK」以外、空行の場合は飛ばす
      if (
        StartAlertStatus == "" ||
        contents[i][startDateCellNum] == ""
      ) {
        continue;
      }

      // イベント開始日の値をセット
      var startDate = new Date(contents[i][startDateCellNum]);
      // 開始通知時刻の値をセット
      var startAlertTime = new Date(contents[i][startAlertTimeCellNum]);
      // 開始通知日時の値をセット
      var StartAlertDateTime = new Date(startDate);
        StartAlertDateTime.setHours(startAlertTime.getHours())
        StartAlertDateTime.setMinutes(startAlertTime.getMinutes());

      //イベント開始通知日時が現在日時より未来の場合は飛ばす
      if (
        StartAlertDateTime > currentDateTime
      ) {
        continue;
      }

      // 値をセット
      var endDate = new Date(contents[i][endDateCellNum]);  //イベント終了日
      var endTime = new Date(contents[i][endTimeCellNum]);  //イベント終了時刻
      var currentEndDate = new Date(endDate);  //現在のイベント終了日時
        currentEndDate.setHours(endTime.getHours())
        currentEndDate.setMinutes(endTime.getMinutes());
      var cycle = contents[i][cycleCellNum];  //イベントの発生周期(日数)
      var newDate = new Date(currentEndDate); //更新後のイベント開始・終了日
        newDate.setDate(newDate.getDate()+cycle);

        console.log(newDate)

      //イベント開始日を上書き入力
      sheet.getRange(topRow + i, startDateCellNum + 1).setValue(newDate);
      //イベント終了日を上書き入力
      sheet.getRange(topRow + i, endDateCellNum + 1).setValue(newDate);

      //開始通知Statusを空に更新
      sheet.getRange(topRow + i, startAlertStatusCellNum + 1).setValue("");


      //　終了通知時刻の値をセット
      var endAlertTime = new Date(contents[i][endAlertTimeCellNum]);
      // 終了通知日時の値をセット
      var EndAlertDateTime = new Date(endDate);
        EndAlertDateTime.setHours(endAlertTime.getHours())
        EndAlertDateTime.setMinutes(endAlertTime.getMinutes())

        console.log({i, currentEndDate, EndAlertDateTime})

      //イベント終了時刻が終了通知時刻と同一の場合はスキップ
      if(currentEndDate.getHours() == EndAlertDateTime.getHours() && currentEndDate.getMinutes() == EndAlertDateTime.getMinutes()) {
        console.log(i, "currentEndDateとEndAlertDateTimeの時：分は一致")
        continue;
      }

      console.log(i, "currentEndDateとEndAlertDateTimeの時：分は不一致")
      //終了通知Statusを空に更新
      sheet.getRange(topRow + i, endAlertStatusCellNum + 1).setValue("");

      //createScheduleAutoを使用している場合
      if (createScheduleAutoIsActive == true){
        sheet.getRange(topRow + i, calStatusCellNum + 1).setValue("");
      }

    } //forここまで

} //resetCycleReminderここまで
