/**
* event_listシートの開始・終了リマインダーを投稿する
*/
function postReminder() {

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

 // スプレッドシートとシートを取得
 // スプレッドシートのURL  https://docs.google.com/spreadsheets/d/{ss_ID}/edit#gid=0
 var ss_id = "★★★";  //★★★ここにss_IDを入力
 var sheet = SpreadsheetApp.openById(ss_id).getSheetByName("event_list");

 // 予定の最終行を取得
 var lastRow = sheet.getLastRow();

 //予定の一覧を取得
 var contents = sheet.getRange(topRow, 1, sheet.getLastRow(), lastCol).getValues();

 //現在日時を取得
 var currentDateTime = new Date();

 /***********************
 /順に開始リマインダーを投稿/
 ************************/
 for (i = 0; i < lastRow - topRow+2; i++) {

     // 「OK」、空行の場合は飛ばす
     var StartAlertStatus = contents[i][startAlertStatusCellNum];

     if (
       StartAlertStatus == "OK" ||
       contents[i][startDateCellNum] == ""
     ) {
       continue;
     }

     // 開始日の値をセット
     var startDate = new Date(contents[i][startDateCellNum]);
     // 開始通知時刻の値をセット
     var startAlertTime = new Date(contents[i][startAlertTimeCellNum]);
     // 開始通知日時の値をセット
     var StartAlertDateTime = new Date(startDate);
       StartAlertDateTime.setHours(startAlertTime.getHours())
       StartAlertDateTime.setMinutes(startAlertTime.getMinutes());

     // 開始通知時刻が現在日時より未来の場合は飛ばす
     if (
       StartAlertDateTime > currentDateTime
     ) {
       continue;
     }

     // 値をセット
     var title = contents[i][titleCellNum]; //タイトル
     var startTime = new Date(contents[i][startTimeCellNum]); // イベント開始時刻
     // イベント開始日時の値をセット
     var StartDateTime = new Date(startDate);
       StartDateTime.setHours(startTime.getHours())
       StartDateTime.setMinutes(startTime.getMinutes());

     // リマインダーの投稿文を生成
     try {
     // イベント開始時刻(時・分)と開始通知時刻(時・分)が一致する場合
       if(StartDateTime.getHours() == StartAlertDateTime.getHours() && StartDateTime.getMinutes() == StartAlertDateTime.getMinutes()){
         var remindScript = title + "が始まりました！";
     // イベント開始時刻(時・分)と開始通知時刻(時・分)が不一致の場合
       } else {
         var remindScript = "まもなく" + title + "が始まります！";
       }

       console.log(remindScript)

       //IFTTTを介してリマインダーを投稿
       callIFTTT(remindScript);

     // エラー処理（ログ出力のみ）
       } catch(e) {
         console.log(e);
       }

       //無事に開始リマインダーが投稿されたら「OK」にする
       sheet.getRange(topRow + i, startAlertStatusCellNum + 1).setValue("OK");

     } //forここまで

   /***********************
   /順に終了リマインダーを投稿/
   ************************/
   for (i = 0; i < lastRow - topRow+2; i++) {

       //「OK」、空行の場合は飛ばす
       var EndAlertStatus = contents[i][endAlertStatusCellNum];

       if (
         EndAlertStatus == "OK" ||
         contents[i][endDateCellNum] == ""
       ) {
         continue;
       }

       // 終了日の値をセット
       var endDate = new Date(contents[i][endDateCellNum]);
       // 終了通知時刻の値をセット
       var endAlertTime = new Date(contents[i][endAlertTimeCellNum]);
       // 終了通知日時の値をセット
       var EndAlertDateTime = new Date(endDate)
         EndAlertDateTime.setHours(endAlertTime.getHours())
         EndAlertDateTime.setMinutes(endAlertTime.getMinutes());

       //終了通知日時が現在日時より未来の場合は飛ばす
       if (
         EndAlertDateTime > currentDateTime
       ) {
         continue;
       }

       // 値をセット
       var title = contents[i][titleCellNum]; //タイトル
       var endTime = new Date(contents[i][endTimeCellNum]);  //イベント終了時刻
       // イベント終了日時の値をセット
       var EndDateTime = new Date(endDate);
         EndDateTime.setHours(endTime.getHours())
         EndDateTime.setMinutes(endTime.getMinutes());

       // リマインダーの投稿文を生成
       try {
       var remindScript2 = "まもなく" + title + "が終了します！";

         console.log(remindScript2)

         //IFTTTを介してリマインダーを投稿
         callIFTTT(remindScript2);

       // エラー処理（ログ出力のみ）
         } catch(e) {
           console.log(e);
         }

         //無事に終了リマインダーが投稿されたら「OK」にする
         sheet.getRange(topRow + i, endAlertStatusCellNum + 1).setValue("OK");

       } //forここまで

} //postReminderここまで
