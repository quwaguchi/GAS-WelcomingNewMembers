var activeSheet = SpreadsheetApp.getActiveSpreadsheet()

//各シートを読み込む
var respSheet = activeSheet.getSheetByName('入会フォーム回答')
var balanceListSheet = activeSheet.getSheetByName('会計シート一覧')
var memberListSheet = activeSheet.getSheetByName('名簿')


//-----下準備-----

var respLR = respSheet.getLastRow()　//回答シートの最終行
var respLC = respSheet.getLastColumn()　//回答シートの最終列
var data = respSheet.getRange(respLR,1,1,respLC).getValues() //最新回答を取得

var timestamp,newMailAddress,newName,newSex
timestamp = data[0][0]
newMailAddress = data[0][1]
newName = data[0][2]
newSex = data[0][7]

//-----新入生が何期生かを求める（newGen期目とする）-----

var inputYear, inputMonth, newGen

inputYear = timestamp.getFullYear()
inputMonth = timestamp.getMonth() +1
//日本では４月始まり３月終わりが１年度とされているので、入力時点の月が4~12月か1~3月かで場合分け
if (inputMonth<=3){
  inputYear = inputYear -1  
}

//(加入時の西暦)-(加入時の学年)-2018で何期目かが求まる。ちなみに、2022年4月に加入した3年生が1期目。
newGen = inputYear - data[0][6] - 2018
if (newGen<1){
  newGen = 1
}


function forNewMember() {

  //-----新入生用会計シートを作る-----

  //テンプレートを会計フォルダーに追加して、名前を「第1期_川口晴人_会計シート」のようにする
  var folder = DriveApp.getFolderById("1Z_Q8jk_wPdk_k-z7zdKj9cHSL_SJLvDV")
  var tmplSheet = DriveApp.getFileById("1YlNmpXvKJGTkjy3NPv_afsdve3vPpYDX0soL6pwyev0")
  var newSheet = tmplSheet.makeCopy('第'+ newGen + '期_' + newName + '_会計シート', folder)
  var newSheetID = newSheet.getId()

  //新入生以外から見えないようにする
  newSheet.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE)
  newSheet.addViewer(newMailAddress)


  //-----新入生にストークのドライブリンクを送る-----
  var address = newMailAddress+',utstoke@gmail.com'
  MailApp.sendEmail(
    address,
    '東大STOKEへようこそ',
    newName+'さん、ストークへようこそ。\n\nストークでは、連絡の大半をDiscordで行っています。このリンク(/*リンク*/)からサーバーに入り、自己紹介チャンネルに簡単な自己紹介を載せてください。\nまた、合宿情報や備品情報、会計ファイルはGoogle Driveで管理しています。/*ドライブのリンク*/からご確認ください。\n\nご不明な点は utstoke@gmail.com までお尋ねください。\n\n\n2022 UT STOKE SKI TEAM'
  )


  //-----データベースを更新-----
  
  //会計シート一覧に追加
  balanceListSheet.appendRow([newGen,newName,newSheetID])

  //名簿に追加
  memberListSheet.appendRow([newGen,newName,newSex])

  //名簿を並び替える

  var balanceListSheetLR = balanceListSheet.getLastRow()
  balanceListSheet.getRange(2,1,balanceListSheetLR-1,3).sort({column:1, ascending:false})

  var memberListSheetLR = memberListSheet.getLastRow()
  memberListSheet.getRange(2,1,memberListSheetLR-1,3).sort({column:1, ascending:false})
  
}
