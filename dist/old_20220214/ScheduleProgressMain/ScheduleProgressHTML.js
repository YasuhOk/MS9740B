// クリックイベント
window.onload = function(){
  var HtmlClick = document.getElementById("data");
  HtmlClick.addEventListener("click", function(evt){
    var target = evt.target
    var borderSet = "solid 3px #f79c7b"
    var borderR = "3px solid rgb(247, 156, 123)"
    // var borderSet2 = "solid 4px #f79c7b"
    // var borderR2 = "4px solid rgb(247, 156, 123)"
    var bgColorSet = "#599ac9"
    var bgColorR = "rgb(89, 154, 201)"
    // 何も設定されていないセルでクリック→枠色変更
    if (target.style.backgroundColor == "" && target.style.borderColor == ""){
        target.style.border = borderSet; 
        target.style.textAlign = "center";
        console.log(target.style.backgroundColor,target.style.border)
    // 枠色が変更されているセルでクリック→セル色変更
    }else if (target.style.backgroundColor == "" && target.style.border == borderR){
        target.style.backgroundColor = bgColorSet;
        target.style.border = borderSet; 
        target.style.textAlign = "left";
        console.log(target.style.backgroundColor,target.style.border)
    // 枠色とセル色が変更されているセルでクリック→リセット
    }else if (target.style.backgroundColor == bgColorR && target.style.border == borderR){
        target.style.backgroundColor = ""
        target.style.border = ""
        target.style.textAlign = "left";
        console.log(target.style.backgroundColor,target.style.border)
    }
    // target.style.backgroundColor = "#599ac9";
  }, false);

};

function ExcelWrite(){
  document.getElementById("button");
  // var workbook = excel.Workbooks.Open(
  //     fso.getAbsolutePathName("X:\Python\MS9740B\MS9740A チェックシート進捗管理版(原本).xlsm"), 0, false, 5, "ms9740");
      console.log("buttonOn");
      // fetch("X:\Python\MS9740B進捗管理アプリケーション\Config.json")
      var fPath = readJSON();
      console.log(fPath);

    }
// JSONファイルの読み込み。
function readJSON(){
 
  var f = "X:\Python\MS9740B進捗管理アプリケーション\Config.json";
  var retJson;
 
  var obj = new XMLHttpRequest();
 
  obj.open( 'get', f, true ); //ファイルオープン : 同期モード
  obj.onload = function() {
    try {
      retJson = JSON.parse(this.responseText); //JSON型でパース。
    } catch (e) {
      alert("コマンド定義ファイルの読み込み、解析に失敗しました。");
    }
  }
  obj.send(null); //ここで読込実行。
  return retJson;
}

