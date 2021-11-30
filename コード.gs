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
    var lastColumn = sheet.getLastColumn();
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
      var last = sheet.getLastColumn;
      var lashyaku = last + 100;
      var lashyakucell = "(5, lashyaku)";
      
      sheet.deleteColumn(60);
      
            
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
      
      
      if(String(member) === "早稲田大学チアダンスチームMYNX"){
        var htmlTemplate = HtmlService.createTemplateFromFile("MYNX");
        htmlTemplate.nameRange = nameRange;
        htmlTemplate.error = error;
        htmlTemplate.member = member;
        return htmlTemplate.evaluate();
      };

      
      
      //入構時間を区切る
      var timeSplitArray = timeRange.split("/");
      var jumbiArray = timeSplitArray.filter(item => item.match(/A/));
      var ichinichiArray = timeSplitArray.filter(item => item.match(/B/));
      var futsukaArray = timeSplitArray.filter(item => item.match(/C/));

      var jumbiNumber = jumbiArray.length;
      var ichinichiNumber = ichinichiArray.length;
      var futsukaNumber = futsukaArray.length;
      var j = 0;
      var i = 0;
      var f = 0;
      var jumbiArray2 = [];
      var ichinichiArray2 = [];
      var futsukaArray2 = [];
      var noNyukou = "入構はありません";
     
      
      while(j<jumbiNumber){
        jumbiArray2[j] = jumbiArray[j].slice(1, 2) + jumbiArray[j].slice(3);
        j++
      }

      while(i<ichinichiNumber){
        ichinichiArray2[i] = ichinichiArray[i].slice(1, 2) + ichinichiArray[i].slice(3);
        i++
      }
      while(f<futsukaNumber){
        futsukaArray2[f] = futsukaArray[f].slice(1, 2) + futsukaArray[f].slice(3);
        f++
      }
      
      
      var jumbiWasedaArray = jumbiArray2.filter(item => item.match(/W/));
      var jumbiToyamaArray = jumbiArray2.filter(item => item.match(/T/));
      var jumbiKougaiArray = jumbiArray2.filter(item => item.match(/K/));
      
      var ichinichiWasedaArray = ichinichiArray2.filter(item => item.match(/W/));
      var ichinichiToyamaArray = ichinichiArray2.filter(item => item.match(/T/));
      var ichinichiKougaiArray = ichinichiArray2.filter(item => item.match(/K/));
      
      var futsukaWasedaArray = futsukaArray2.filter(item => item.match(/W/));
      var futsukaToyamaArray = futsukaArray2.filter(item => item.match(/T/));
      var futsukaKougaiArray = futsukaArray2.filter(item => item.match(/K/));
      
            
      var jumbiWasedaNumber = jumbiWasedaArray.length;
      var jumbiToyamaNumber = jumbiToyamaArray.length;
      var jumbiKougaiNumber = jumbiKougaiArray.length;
      
      var ichinichiWasedaNumber = ichinichiWasedaArray.length;
      var ichinichiToyamaNumber = ichinichiToyamaArray.length;
      var ichinichiKougaiNumber = ichinichiKougaiArray.length;

      var futsukaWasedaNumber = futsukaWasedaArray.length;
      var futsukaToyamaNumber = futsukaToyamaArray.length;
      var futsukaKougaiNumber = futsukaKougaiArray.length;
      
      var jw = 0;
      var jt = 0;
      var jk = 0;
      var iw = 0;
      var it = 0;
      var ik = 0;
      var fw = 0;
      var ft = 0;
      var fk = 0;
      var jumbiWasedaArray2 = [];
      var jumbiToyamaArray2 = [];
      var jumbiKougaiArray2 = [];
      
      var ichinichiWasedaArray2 = [];
      var ichinichiToyamaArray2 = [];
      var ichinichiKougaiArray2 = [];
      
      var futsukaWasedaArray2 = [];
      var futsukaToyamaArray2= [];
      var futsukaKougaiArray2 = [];
      
      
      while(jw<jumbiWasedaNumber){
        jumbiWasedaArray2[jw] = jumbiWasedaArray[jw].slice(1);
        jw++
      }
      while(jt<jumbiToyamaNumber){
        jumbiToyamaArray2[jt] = jumbiToyamaArray[jt].slice(1);
        jt++
      }
      while(jk<jumbiKougaiNumber){
        jumbiKougaiArray2[jk] = jumbiKougaiArray[jk].slice(1);
        jk++
      }
      while(iw<ichinichiWasedaNumber){
        ichinichiWasedaArray2[iw] = ichinichiWasedaArray[iw].slice(1);
        iw++
      }
      while(it<ichinichiToyamaNumber){
        ichinichiToyamaArray2[it] = ichinichiToyamaArray[it].slice(1);
        it++
      }
      while(ik<ichinichiKougaiNumber){
        ichinichiKougaiArray2[ik] = ichinichiKougaiArray[ik].slice(1);
        ik++
      }
      while(fw<futsukaWasedaNumber){
        futsukaWasedaArray2[fw] = futsukaWasedaArray[fw].slice(1);
        fw++
      }
      while(ft<futsukaToyamaNumber){
        futsukaToyamaArray2[ft] = futsukaToyamaArray[ft].slice(1);
        ft++
      }
      while(fk<futsukaKougaiNumber){
        futsukaKougaiArray2[fk] = futsukaKougaiArray[fk].slice(1);
        fk++
      }
      
      
      var jumbiWasedaMin = String(Math.min.apply(null, jumbiWasedaArray2));
      var jumbiToyamaMin = String(Math.min.apply(null, jumbiToyamaArray2));
      var jumbiKougaiMin = String(Math.min.apply(null, jumbiKougaiArray2));
      var ichinichiWasedaMin = String(Math.min.apply(null, ichinichiWasedaArray2));
      var ichinichiToyamaMin = String(Math.min.apply(null, ichinichiToyamaArray2));
      var ichinichiKougaiMin = String(Math.min.apply(null, ichinichiKougaiArray2));
      var futsukaWasedaMin = String(Math.min.apply(null, futsukaWasedaArray2));
      var futsukaToyamaMin = String(Math.min.apply(null, futsukaToyamaArray2));
      var futsukaKougaiMin = String(Math.min.apply(null, futsukaKougaiArray2));
      var colon = "：";
      
      //準備日
      if(jumbiWasedaMin.length === 3){
        jumbiWasedaMin = "0" + jumbiWasedaMin;
        var jwa = String(jumbiWasedaMin).substr(0, 2);
        var jwc = String(jumbiWasedaMin).substr(2);
        var jwstrTime = jwa + colon + jwc + "[早]" + "\n";
      }else if(jumbiWasedaMin.length ===4){
        var jwa = String(jumbiWasedaMin).substr(0, 2);
        var jwc = String(jumbiWasedaMin).substr(2);
        var jwstrTime = jwa + colon + jwc + "[早]" + "\n";
      }else{
        var jwstrTime = "";
      }
      if(jumbiToyamaMin.length === 3){
        jumbiToyamaMin = "0" + jumbiToyamaMin;
        var jta = String(jumbiToyamaMin).substr(0, 2);
        var jtc = String(jumbiToyamaMin).substr(2);
        var jtstrTime = jta + colon + jtc + "[戸]" + "\n";
      }else  if(jumbiToyamaMin.length === 4){
        var jta = String(jumbiToyamaMin).substr(0, 2);
        var jtc = String(jumbiToyamaMin).substr(2);
        var jtstrTime = jta + colon + jtc + "[戸]" + "\n";
      }else{
        var jtstrTime = "";
      }
      if(jumbiKougaiMin.length === 3){
        jumbiKougaiMin = "0" + jumbiKougaiMin;
        var jka = String(jumbiKougaiMin).substr(0, 2);
        var jkc = String(jumbiKougaiMin).substr(2);
        var jkstrTime = jka + colon + jkc + "[外]" + "\n";
      }else if(jumbiKougaiMin.length === 4){
        var jka = String(jumbiKougaiMin).substr(0, 2);
        var jkc = String(jumbiKougaiMin).substr(2);
        var jkstrTime = jka + colon + jkc + "[外]" + "\n";
      }else{
        var jkstrTime = "";
      }
      
      //一日目
      if(ichinichiWasedaMin.length === 3){
        ichinichiWasedaMin = "0" + ichinichiWasedaMin;
        var iwa = String(ichinichiWasedaMin).substr(0, 2);
        var iwc = String(ichinichiWasedaMin).substr(2);
        var iwstrTime = iwa + colon + iwc + "[早]" + "\n";
      }else if(ichinichiWasedaMin.length === 4){
        var iwa = String(ichinichiWasedaMin).substr(0, 2);
        var iwc = String(ichinichiWasedaMin).substr(2);
        var iwstrTime = iwa + colon + iwc + "[早]" + "\n";
      }else{
        var iwstrTime = "";
      }
      if(ichinichiToyamaMin.length === 3){
        ichinichiToyamaMin = "0" + ichinichiToyamaMin;
        var ita = String(ichinichiToyamaMin).substr(0, 2);
        var itc = String(ichinichiToyamaMin).substr(2);
        var itstrTime = ita + colon + itc + "[戸]" + "\n";
      }else if(ichinichiToyamaMin.length === 4){
        var ita = String(ichinichiToyamaMin).substr(0, 2);
        var itc = String(ichinichiToyamaMin).substr(2);
        var itstrTime = ita + colon + itc + "[戸]" + "\n";
      }else{
        var itstrTime = "";
      }
      if(ichinichiKougaiMin.length === 3){
        ichinichiKougaiMin = "0" + ichinichiKougaiMin;
        var ika = String(ichinichiKougaiMin).substr(0, 2);
        var ikc = String(ichinichiKougaiMin).substr(2);
        var ikstrTime = ika + colon + ikc + "[外]" + "\n";
      }else if(ichinichiKougaiMin.length === 4){
        var ika = String(ichinichiKougaiMin).substr(0, 2);
        var ikc = String(ichinichiKougaiMin).substr(2);
        var ikstrTime = ika + colon + ikc + "[外]" + "\n";
      }else{
        var ikstrTime = ""
      }
      
      //二日目
      if(futsukaWasedaMin.length === 3){
        futsukaWasedaMin = "0" + futsukaWasedaMin;
        var fwa = String(futsukaWasedaMin).substr(0, 2);
        var fwc = String(futsukaWasedaMin).substr(2);
        var fwstrTime = fwa + colon + fwc + "[早]" + "\n";
      }else if(futsukaWasedaMin.length === 4){
        var fwa = String(futsukaWasedaMin).substr(0, 2);
        var fwc = String(futsukaWasedaMin).substr(2);
        var fwstrTime = fwa + colon + fwc + "[早]" + "\n";
      }else{
        var fwstrTime = "";
      }
      if(futsukaToyamaMin.length === 3){
        futsukaToyamaMin = "0" + futsukaToyamaMin;
        var fta = String(futsukaToyamaMin).substr(0, 2);
        var ftc = String(futsukaToyamaMin).substr(2);
        var ftstrTime = fta + colon + ftc + "[戸]" + "\n";
      }else if(futsukaToyamaMin.length === 4){
        var fta = String(futsukaToyamaMin).substr(0, 2);
        var ftc = String(futsukaToyamaMin).substr(2);
        var ftstrTime = fta + colon + ftc + "[戸]" + "\n";
      }else{
        var ftstrTime = "";
      }
      if(futsukaKougaiMin.length === 3){
        futsukaKougaiMin = "0" + futsukaKougaiMin;
        var fka = String(futsukaKougaiMin).substr(0, 2);
        var fkc = String(futsukaKougaiMin).substr(2);
        var fkstrTime = fka + colon + fkc + "[外]" + "\n";
      }else if(futsukaKougaiMin.length === 4){
        var fka = String(futsukaKougaiMin).substr(0, 2);
        var fkc = String(futsukaKougaiMin).substr(2);
        var fkstrTime = fka + colon + fkc + "[外]" + "\n";
      }else{
        var fkstrTime = "";
      }
      
      

      //このままだとキャンパス順に出るので時間順に変える
      var jTime = [jwstrTime, jtstrTime, jkstrTime];
      var iTime = [iwstrTime, itstrTime, ikstrTime];
      var fTime = [fwstrTime, ftstrTime, fkstrTime];

      //配列を時間順に並び替え
      //このままだとなぜか「,」も表示されちゃう
      jTime.sort((a,b) => parseInt(a) - parseInt(b));
      iTime.sort((a,b) => parseInt(a) - parseInt(b));
      fTime.sort((a,b) => parseInt(a) - parseInt(b));
      
      //「,」を消すために配列を結合して文字列に
      var jTime2 = jTime.toString();
      var iTime2 = iTime.toString()
      var fTime2 = fTime.toString()
      
      //文字列から「,」を削除
      //正規表現を使わないと最初のしか削除されないから注意！
      var jTime3 = jTime2.replace(/,/g, '');
      var iTime3 = iTime2.replace(/,/g, '');
      var fTime3 = fTime2.replace(/,/g, '');
      
      
      //その日に入構がない人を「入構なし」に
      //ここに置かないと配列の並び替えとか結合でエラーが出る
      if(jwstrTime === ""　&& jtstrTime === "" && jkstrTime === ""){
        jTime3 = noNyukou;
      }
      if(iwstrTime === ""　&& itstrTime === "" && ikstrTime === ""){
        iTime3 = noNyukou;
      }
      if(fwstrTime === ""　&& ftstrTime === "" && fkstrTime === ""){
        fTime3 = noNyukou;
      }

      
      //アプリに返す
      var htmlTemplate = HtmlService.createTemplateFromFile("result");
      htmlTemplate.nameRange = nameRange;
      htmlTemplate.member = member;
      htmlTemplate.jTime = jTime3;
      htmlTemplate.iTime = iTime3;
      htmlTemplate.fTime = fTime3;


  
      //SSに記入
      if(statusArray[searchIdm -1] != "" && gateArray[searchIdm -1] === gate){
        var range = sheet.getRange(searchIdm, 13);
        range.insertCells(SpreadsheetApp.Dimension.COLUMNS);
        sheet.getRange(searchIdm, 13).setValue(dateLog + " @" + gate);//時刻を記入
        
//          if(last + 50 == null){
//            //range.insertCells(SpreadsheetApp.Dimension.COLUMNS);
//          }else if((3, lashyaku) === null){
//            sheet.deleteColumn(lashyaku);
//            };

      }else if(statusArray[searchIdm -1] != "" && gateArray[searchIdm -1] != gate){
        sheet.getRange(searchIdm, 12).setValue(gate);//セルに記入
        var range = sheet.getRange(searchIdm, 13);
        range.insertCells(SpreadsheetApp.Dimension.COLUMNS);
        sheet.getRange(searchIdm, 13).setValue(dateLog + " @" + gate);//時刻を記入
        
//          if(last + 50 == null){
//            //range.insertCells(SpreadsheetApp.Dimension.COLUMNS);
//          }else if((3, lashyaku) == null){
//            sheet.deleteColumn(lashyaku);
//          };
          

      }else{
        sheet.getRange(searchIdm, 11).setValue(statusIn);//セルに記入
        sheet.getRange(searchIdm, 12).setValue(gate);//セルに記入
        var range = sheet.getRange(searchIdm, 13);
        range.insertCells(SpreadsheetApp.Dimension.COLUMNS);
        sheet.getRange(searchIdm, 13).setValue(dateLog + " @" + gate);//時刻を記入

//        if(last + 50 == null){
//          //range.insertCells(SpreadsheetApp.Dimension.COLUMNS);
//      }else if((3, lashyaku) === null){
//          sheet.deleteColumn(lashyaku);
//      };
        
      }
      
      //HTMLのリターン、ここに置くとなぜか速い。SSの前に置くとSSに記入されなくなる。
      return htmlTemplate.evaluate();
      
      
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
  var i = 1;
  
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
  var i = 6300;
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


//入構人数確認
function people(){

  var id = '1MBeZEEVi1RIv1L32XN7Zws0vP0Ri5k8x9bJhC8EELMw';
  var sheet1 = SpreadsheetApp.openById(id).getSheetByName("data1");
  var sheet2 = SpreadsheetApp.openById(id).getSheetByName("入構者数");
  var textFinderW5 = sheet1.createTextFinder("^(?=.*11/05)(?=.*早稲田)").useRegularExpression(true);
  var textFinderW6 = sheet1.createTextFinder("^(?=.*11/06)(?=.*早稲田)").useRegularExpression(true);
  var textFinderW7 = sheet1.createTextFinder("^(?=.*11/07)(?=.*早稲田)").useRegularExpression(true);
  var textFinderT5 = sheet1.createTextFinder("^(?=.*11/05)(?=.*戸山)").useRegularExpression(true);
  var textFinderT6 = sheet1.createTextFinder("^(?=.*11/06)(?=.*戸山)").useRegularExpression(true);
  var textFinderT7 = sheet1.createTextFinder("^(?=.*11/07)(?=.*戸山)").useRegularExpression(true);
  var textFinderK5 = sheet1.createTextFinder("^(?=.*11/05)(?=.*構外)").useRegularExpression(true);
  var textFinderK6 = sheet1.createTextFinder("^(?=.*11/06)(?=.*構外)").useRegularExpression(true);
  var textFinderK7 = sheet1.createTextFinder("^(?=.*11/07)(?=.*構外)").useRegularExpression(true);
  var cellsW5 = textFinderW5.findAll();
  var cellsW6 = textFinderW6.findAll();
  var cellsW7 = textFinderW7.findAll();
  var cellsT5 = textFinderT5.findAll();
  var cellsT6 = textFinderT6.findAll();
  var cellsT7 = textFinderT7.findAll();
  var cellsK5 = textFinderK5.findAll();
  var cellsK6 = textFinderK6.findAll();
  var cellsK7 = textFinderK7.findAll();
  var itsukaWaseda = cellsW5.length;
  var muikaWaseda = cellsW6.length;
  var nanokaWaseda = cellsW7.length;
  var itsukaToyama = cellsT5.length;
  var muikaToyama = cellsT6.length;
  var nanokaToyama = cellsT7.length;
  var itsukaKougai = cellsK5.length;
  var muikaKougai = cellsK6.length;
  var nanokaKougai = cellsK7.length;
  

  sheet2.getRange("B9").setValue(itsukaWaseda);
  sheet2.getRange("B11").setValue(muikaWaseda);
  sheet2.getRange("B13").setValue(nanokaWaseda);
  sheet2.getRange("C9").setValue(itsukaToyama);
  sheet2.getRange("C11").setValue(muikaToyama);
  sheet2.getRange("C13").setValue(nanokaToyama);
  sheet2.getRange("E9").setValue(itsukaKougai);
  sheet2.getRange("E11").setValue(muikaKougai);
  sheet2.getRange("E13").setValue(nanokaKougai);

}
    

function peopleBU(){

  var id = '1c5ncM9uL7BzTm4yGGWb3CCqNA1FGVDV7VsOohA1Jf3A';
  var sheet1 = SpreadsheetApp.openById(id).getSheetByName("data1");
  var sheet2 = SpreadsheetApp.openById(id).getSheetByName("入構者数");
  var textFinderW5 = sheet1.createTextFinder("^(?=.*11/05)(?=.*早稲田)").useRegularExpression(true);
  var textFinderW6 = sheet1.createTextFinder("^(?=.*11/06)(?=.*早稲田)").useRegularExpression(true);
  var textFinderW7 = sheet1.createTextFinder("^(?=.*11/07)(?=.*早稲田)").useRegularExpression(true);
  var textFinderT5 = sheet1.createTextFinder("^(?=.*11/05)(?=.*戸山)").useRegularExpression(true);
  var textFinderT6 = sheet1.createTextFinder("^(?=.*11/06)(?=.*戸山)").useRegularExpression(true);
  var textFinderT7 = sheet1.createTextFinder("^(?=.*11/07)(?=.*戸山)").useRegularExpression(true);
  var textFinderK5 = sheet1.createTextFinder("^(?=.*11/05)(?=.*構外)").useRegularExpression(true);
  var textFinderK6 = sheet1.createTextFinder("^(?=.*11/06)(?=.*構外)").useRegularExpression(true);
  var textFinderK7 = sheet1.createTextFinder("^(?=.*11/07)(?=.*構外)").useRegularExpression(true);
  var cellsW5 = textFinderW5.findAll();
  var cellsW6 = textFinderW6.findAll();
  var cellsW7 = textFinderW7.findAll();
  var cellsT5 = textFinderT5.findAll();
  var cellsT6 = textFinderT6.findAll();
  var cellsT7 = textFinderT7.findAll();
  var cellsK5 = textFinderK5.findAll();
  var cellsK6 = textFinderK6.findAll();
  var cellsK7 = textFinderK7.findAll();
  var itsukaWaseda = cellsW5.length;
  var muikaWaseda = cellsW6.length;
  var nanokaWaseda = cellsW7.length;
  var itsukaToyama = cellsT5.length;
  var muikaToyama = cellsT6.length;
  var nanokaToyama = cellsT7.length;
  var itsukaKougai = cellsK5.length;
  var muikaKougai = cellsK6.length;
  var nanokaKougai = cellsK7.length;
  

  sheet2.getRange("B9").setValue(itsukaWaseda);
  sheet2.getRange("B11").setValue(muikaWaseda);
  sheet2.getRange("B13").setValue(nanokaWaseda);
  sheet2.getRange("C9").setValue(itsukaToyama);
  sheet2.getRange("C11").setValue(muikaToyama);
  sheet2.getRange("C13").setValue(nanokaToyama);
  sheet2.getRange("E9").setValue(itsukaKougai);
  sheet2.getRange("E11").setValue(muikaKougai);
  sheet2.getRange("E13").setValue(nanokaKougai);

}
  