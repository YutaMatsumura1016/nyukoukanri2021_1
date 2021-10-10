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
    var timeArray = arrayRoll[9];
    var statusArray = arrayRoll[10];
    var gateArray = arrayRoll[11];
    var statusIn = "入 構";
    var statusRe = "再入構";
    var error = "エラーが発生しました";
    var treatment = "係員は処理を行ってください。";
    var unregistered = "登録されていないカードです。\n入構できません。";
    
    //現在時刻
    var date = new Date();
    var dateLog = (Utilities.formatDate(date, 'Asia/Tokyo', 'MM/dd HH:mm'));
    
    //idmをSSから探す
    var searchIdm = (idmArray.indexOf(idm)) + 1;//IDmの行数が出る
    
    
    if(searchIdm != ""){//IDmが見つかったら
 
      var nameRange = (nameArray[searchIdm -1]);//IDmに対応した名前を探す
      var statusMember = (menberArray[searchIdm -1]);//IDmに対応した団体名を探す
      var timeRange = (timeArray[searchIdm -1]);//IDmに対応した入構時間を探す
      
            
      //団体名を区切る
      var memberSplitArray = statusMember.split("/");
      var memberNumber = memberSplitArray.length;
      var member = "";
      var i = 0;
      
      while(i<memberNumber-1){
        member = member + memberSplitArray[i] + "\n";
        i++
      }
      member = member + memberSplitArray[i];
      
      
      
      
      
      //入構時間を区切る
      var timeSplitArray = timeRange.split("/");
      var jumbiArray = timeSplitArray.filter(item => item.match(/A/));
      var ichinichiArray = timeSplitArray.filter(item => item.match(/B/));
      var futsukaArray = timeSplitArray.filter(item => item.match(/C/));
      
      var jumbiNumber = jumbiArray.length;
      var ichinichiNumber = ichinichiArray.length;
      var futsukaNumber = futsukaArray.length;
      var j = 0;
      var k = 0;
      var l = 0;
      var jumbiArray2 = [];
      var ichinichiArray2 = [];
      var futsukaArray2 = [];
      
      
      while(j<jumbiNumber){
        jumbiArray2[j] = jumbiArray[j].slice(1);
        j++
      }

      while(k<ichinichiNumber){
        ichinichiArray2[k] = ichinichiArray[k].slice(1);
        k++
      }
      while(l<futsukaNumber){
        futsukaArray2[l] = futsukaArray[l].slice(1);
        l++
      }
      
      
      var jumbiMin = Math.min.apply(null, jumbiArray2);
      var ichinichiMin = Math.min.apply(null, ichinichiArray2);
      var futsukaMin = Math.min.apply(null, futsukaArray2);
      
      
      if(jumbiMin.length !=4){
        jumbiMin = "0" + jumbiMin;
        var ja = String(jumbiMin).substr(0, 2);
        var jb = ":";
        var jc = String(jumbiMin).substr(2);
        var jstrTime = "準備日：" + ja + jb + jc + "\n";
      }else{
        var ja = String(jumbiMin).substr(0, 2);
        var jb = ":";
        var jc = String(jumbiMin).substr(2);
        var jstrTime = "準備日：" + ja + jb + jc + "\n";
      }
      
      if(ichinichiMin.length !=4){
        ichinichiMin = "0" + ichinichiMin;
        var ka = String(ichinichiMin).substr(0, 2);
        var kb = ":";
        var kc = String(ichinichiMin).substr(2);
        var kstrTime = "一日目：" + ka + kb + kc + "\n";
      }else{
        var ka = String(ichinichiMin).substr(0, 2);
        var kb = ":";
        var kc = String(ichinichiMin).substr(2);
        var kstrTime = "一日目：" + ka + kb + kc + "\n";
      }
      
      if(futsukaMin.length !=4){
      futsukaMin = "0" + futsukaMin;
        var la =  String(futsukaMin).substr(0, 2);
        var lb = ":";
        var lc =  String(futsukaMin).substr(2);
        var lstrTime = "二日目：" + la + lb + lc;
      }else{
        var la =  String(futsukaMin).substr(0, 2);
        var lb = ":";
        var lc =  String(futsukaMin).substr(2);
        var lstrTime = "二日目：" + la + lb + lc;
      }
      
      
      
      if(jumbiNumber >= 1 && ichinichiNumber >=1 && futsukaNumber >=1){
        var time = jstrTime + kstrTime + lstrTime;
      }else if(jumbiNumber >= 1 && ichinichiNumber === 0 && futsukaNumber === 0){
        var time = jstrTime;
      }else if(jumbiNumber >= 1 && ichinichiNumber >=1 && futsukaNumber === 0){
        var time = jstrTime + kstrTime;
      }else if(jumbiNumber >= 1 && ichinichiNumber === 0 && futsukaNumber >=1){
        var time = jstrTime + lstrTime;
      }else if(jumbiNumber === 0 && ichinichiNumber >=1 && futsukaNumber === 0){
        var time = kstrTime;
      }else if(jumbiNumber === 0 && ichinichiNumber >=1 && futsukaNumber >=1){
        var time = kstrTime + lstrTime;
      }else if(jumbiNumber === 0 && ichinichiNumber === 0 && futsukaNumber >=1){
        var time = lstrTime;
      }else{
        var time = "入構時間が記録されていないか不正です";
      }

      
      
      //アプリに返す
      var htmlTemplate = HtmlService.createTemplateFromFile("result");
      htmlTemplate.nameRange = nameRange;
      htmlTemplate.member = member;
      htmlTemplate.time = time;
      return htmlTemplate.evaluate();
      
      
      
      //SSに記入
      if(statusArray[searchIdm -1] != "" && gateArray[searchIdm -1] === gate){
        var range = sheet.getRange(searchIdm, 13);
        range.insertCells(SpreadsheetApp.Dimension.COLUMNS);
        sheet.getRange(searchIdm, 13).setValue(dateLog + " @" + gate);//時刻を記入
      }else if(statusArray[searchIdm -1] != "" && gateArray[searchIdm -1] != gate){
        sheet.getRange(searchIdm, 12).setValue(gate);//セルに記入
        var range = sheet.getRange(searchIdm, 13);
        range.insertCells(SpreadsheetApp.Dimension.COLUMNS);
        sheet.getRange(searchIdm, 13).setValue(dateLog + " @" + gate);//時刻を記入
      }else{
        sheet.getRange(searchIdm, 11).setValue(statusIn);//セルに記入
        sheet.getRange(searchIdm, 12).setValue(gate);//セルに記入
        var range = sheet.getRange(searchIdm, 13);
        range.insertCells(SpreadsheetApp.Dimension.COLUMNS);
        sheet.getRange(searchIdm, 13).setValue(dateLog + " @" + gate);//時刻を記入
      }
      
      
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

  var timeRange = sheet.getRange(2, 11, lastRow, lastColumn);
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


//重複確認
function duplicate(){

  var id = '1MBeZEEVi1RIv1L32XN7Zws0vP0Ri5k8x9bJhC8EELMw';
  var sheet = SpreadsheetApp.openById(id).getSheetByName("data1");
  var array = sheet.getDataRange().getValues();
  var _ = Underscore.load();
  var arrayRoll = _.zip.apply(_, array);
  var nameArray = arrayRoll[1];
  var gakusekiArray = arrayRoll[4];
  var idmArray = arrayRoll[5];
  var TELArray = arrayRoll[6];
  var lastRow = (idmArray.length)-1;
  var i = 1;
  
  var deplicate = idmArray.filter(function (x, i, self) {
    return self.indexOf(x) === i && i !== self.lastIndexOf(x);
  });
        
  console.log(deplicate);
  
}
        