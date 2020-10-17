/**
 * スプレッドシート表示の際にカレンダー連携の項目を追加
 */
function onOpen() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //スプレッドシートのメニューにカスタムメニュー「カレンダー連携 > 実行」を作成
  var subMenus = [];
  subMenus.push({
    name: "カレンダー登録実行",
    functionName: "createSchedule"  //実行で呼び出す関数を指定
  });
  ss.addMenu("カレンダー連携", subMenus);
}



/**
* リマインダー用スクリプトをIFTTTへPOSTする
*/
function callIFTTT(value1) {
 // トリガーのURL  https://maker.ifttt.com/trigger/{event}/with/key/{key_Num}'
 // event：　IFTTT上で設定した、event名称を入力
 // key_Num:　IFTTTのwebhook documentに表示される、自分のkeyNumberを入力

 var url = 'https://maker.ifttt.com/trigger/{event}/with/key/{key_Num}';

 var headers = {
   "Content-Type": "application/json"
 };

 // post内容 post内容にvalue1を設定している前提
 // post内容をvalue1以外でIFTTT上変更した場合、"value1"の部分を変更する
 var data = {
   "value1": value1
 };

 console.log(data)

 var options = {
   "method" : "post",
   "headers": headers,
   "payload" : JSON.stringify(data)
 };

 UrlFetchApp.fetch(url, options);
}
