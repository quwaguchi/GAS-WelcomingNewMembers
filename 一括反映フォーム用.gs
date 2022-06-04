//一括反映用フォームの選択肢を更新する
function refleshChoices() {
  
  //名簿一覧を1次元配列として取得
  var memberListSheetLR = memberListSheet.getLastRow()
  var members = memberListSheet.getRange(2,2,memberListSheetLR-1,1).getValues().flat()

  var form = FormApp.openById('1jdKNFxxqjamP_DRHdTi479d9gdH3U9P78PpYV48A3GE')

  //セクションを追加した場合は、そのタイトルもitemとしてカウントされる。選択肢を更新したい質問は0,1,..と数えて5番目。
  var allocateSection = form.getItems()[5]
  allocateSection.asCheckboxGridItem().setColumns(members)
  
}
