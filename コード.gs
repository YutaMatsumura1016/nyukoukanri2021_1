//Underscoreのロードが必要だよ！！
function doGet(e) {

var rowData = {};  

  if(e.parameter == undefined) {
    //エラーを返す
    var getError = "読み取りエラーが発生しました。もう一度タッチしてください。"
    rowData.value = getError;
    return ContentService.createTextOutput(rowData.value);  
  }else{
    
    var id = '1MBeZEEVi1RIv1L32XN7Zws0vP0Ri5k8x9bJhC8EELMw';
    var sheet = SpreadsheetApp.openById(id).getSheetByName("data1");

    //idmをアプリから受け取る
    var idm = e.parameter.idm;
    var gate = e.parameter.gate;
    
    if(gate === "waseda"){
      gate = "早稲田";
    }else if(gate === "toyama"){
      gate = "戸山";
    }else if(gate === "kougai"){
      gate = "構外";
    }else if(gate === "gakkan"){
      gate = "学館";
    }
    
    //-----------検索--------------
    var array = sheet.getDataRange().getValues();
    var _ = Underscore.load();
    var arrayRoll = _.zip.apply(_, array);
    var nameArray = arrayRoll[1];
    var menberArray = arrayRoll[2];
    var idmArray = arrayRoll[5];
//    var statusArray = arrayRoll[9];
//    var gateArray = arrayRoll[10];
//    var timeArray = arrayRoll[11];
    var statusIn = "入 構";
    var statusRe = "再入構";
    var error = "エラーが発生しました";
    var treatment = "係員は処理を行ってください。";
    var unregistered = "登録されていないカードです。\n入構できません。";
    var outlier = "\nシートに異常な値が記録されています。\nスプレッドシートを確認してください。";
    
    //現在時刻
    var date = new Date();
    var dateLog = (Utilities.formatDate(date, 'Asia/Tokyo', 'MM/dd HH:mm'));
    
    //idmをSSから探す
    var searchIdm = (idmArray.indexOf(idm)) + 1;//IDmの行数が出る
    
    if(searchIdm != ""){//IDmが見つかったら
 
      var nameRange = (nameArray[searchIdm -1]);//IDmに対応した名前を探す
      var statusMember = (menberArray[searchIdm -1]);//IDmに対応した団体名を探す
//      var statusRange = (statusArray[searchIdm -1]);//IDmに対応したステータスを探す
//      var statusGate = (gateArray[searchIdm -1]);//IDmに対応したキャンパスを探す
//      var statusTime = (timeArray[searchIdm -1]);//IDmに対応した前回入構時刻を探す
      
      
//      if(statusRange == "" || statusRange == statusIn){
        sheet.getRange(searchIdm, 10).setValue(statusIn);//セルに記入
        sheet.getRange(searchIdm, 11).setValue(gate);//セルに記入
        var range = sheet.getRange(searchIdm, 12);
        range.insertCells(SpreadsheetApp.Dimension.COLUMNS);
        sheet.getRange(searchIdm, 12).setValue(dateLog + " @" + gate);//時刻を記入
        

        //団体名を区切る
        var memberSplitArray = statusMember.split("/");
        var memberNumber = memberSplitArray.length;
        var member = "";
        var i = 0;
        
        while(i<memberNumber-1){
          member = member + memberSplitArray[i] + "\n" + "　 　　　";
          i++
        }
        member = member + memberSplitArray[i];

        //アプリに返す
        var htmlTemplate = HtmlService.createTemplateFromFile("result");
        htmlTemplate.nameRange = nameRange;
        htmlTemplate.member = member;
//        htmlTemplate.statusIn = statusIn;
//        htmlTemplate.gate = gate;
//        htmlTemplate.dateLog = dateLog;
        return htmlTemplate.evaluate();
    
//      }else{
//        var htmlTemplate = HtmlService.createTemplateFromFile("outlier");
//        htmlTemplate.error = error;
//        htmlTemplate.treatment = treatment;
//        htmlTemplate.error = outlier;
//        htmlTemplate.statusRange = statusRange;
//        return htmlTemplate.evaluate();    
//      }
      
  
    }else{//IDmが見つからなかったら
      var htmlTemplate = HtmlService.createTemplateFromFile("unregistered");
      htmlTemplate.error = error;
      htmlTemplate.treatment = treatment;
      htmlTemplate.unregistered = unregistered;
      return htmlTemplate.evaluate();
      
      }//IDmが登録されているかどうか

  }//読み取りエラーかどうか

}//全体





//データリセット
function reset(){
  var id = '1MBeZEEVi1RIv1L32XN7Zws0vP0Ri5k8x9bJhC8EELMw';
  var sheet = SpreadsheetApp.openById(id).getSheetByName("data1");
  var i = 2;
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  var timeRange = sheet.getRange(2, 10, lastRow, lastColumn);
  timeRange.setValue(null);
}


//間違い探し
function digits(){
  var id = '1MBeZEEVi1RIv1L32XN7Zws0vP0Ri5k8x9bJhC8EELMw';
  var sheet = SpreadsheetApp.openById(id).getSheetByName("data1");
  var array = sheet.getDataRange().getValues();
  var _ = Underscore.load();
  var arrayRoll = _.zip.apply(_, array);
  var gakusekiArray = arrayRoll[4];
  var idmArray = arrayRoll[5];
  var TELArray = arrayRoll[6];
  var lastRow = (idmArray.length)-1;
  var i = 4429;
  
  while(i<=lastRow){
    var itsGakuseki = gakusekiArray[i];
    var itsIdm = idmArray[i]; 
    var itsTEL = TELArray[i];
    
    if(String(itsGakuseki).length != 8 && itsGakuseki != ""){
      if(String(itsGakuseki).length === 9){
        sheet.getRange(i+1, 5).setBackground("#f4e25d");
      }else{sheet.getRange(i+1, 5).setBackground("#519965");
      }
    }
    if(String(itsIdm).length != 16){
      sheet.getRange(i+1, 6).setBackground("#c15050");
    }
    if(String(itsIdm).charAt(0) != 1 && String(itsIdm).charAt(0) != 0){
      sheet.getRange(i+1, 6).setBackground("#c15050");
    }
    if(String(itsIdm).substring(0,3) === "FE00" || String(itsIdm).substring(0,3) === "fe00" || String(itsIdm).substring(0,3) === "0003"){
      sheet.getRange(i+1, 6).setBackground("#e89e9e");
    }
    if(String(itsTEL).length != 11){
      sheet.getRange(i+1, 7).setBackground("#274c82");
    }
    i++;
  }
 } 
 
 
 
 
//IDmの修正
function modify(){
  var id = '1MBeZEEVi1RIv1L32XN7Zws0vP0Ri5k8x9bJhC8EELMw';
  var sheet = SpreadsheetApp.openById(id).getSheetByName("data1");
  var i = 4429;
  var lastRow = sheet.getLastRow() + 1;


  while(i<lastRow){
    var IDmRange = sheet.getRange(i, 6);
    var bgColor = IDmRange.getBackground();
    
    if(bgColor != "#c15050"){
      var IDm = IDmRange.getValue();
      var IDm2 = "0" + IDm.slice(1);
      var IDm3 = IDm2.toUpperCase();
  
      if(IDmRange != null){
        IDmRange.setValue(IDm3);  
      }
    }
    i++
  }
}
