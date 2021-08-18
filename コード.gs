//Underscoreのロードが必要だよ！！
function doGet(e) {

var rowData = {};  

  if(e.parameter == undefined) {
    //エラーを返す
    var getError = "読み取りエラーが発生しました。もう一度タッチしてください。"
    rowData.value = getError;
    return ContentService.createTextOutput(rowData.value);  
  }else{
    
    var id = '1MXd7VHuOGZDwVmCjQbJk5np87-SLz-dkp3IA0YsGvMs';
    var sheet = SpreadsheetApp.openById(id).getSheetByName("data1");
    var sheet3 = SpreadsheetApp.openById(id).getSheetByName("表示用");

    //idmをandroidから受け取る
    var idm = e.parameter.idm;
    var gate = e.parameter.gate;
    
    //-----------検索--------------
    var array = sheet.getDataRange().getValues();
    var _ = Underscore.load();
    var arrayRoll = _.zip.apply(_, array);
    var nameArray = arrayRoll[1];
    var menberArray = arrayRoll[2];
    var idmArray = arrayRoll[5];
    var statusArray = arrayRoll[6];
    var gateArray = arrayRoll[7];
    var timeArray = arrayRoll[8];
    var statusIn = "入構";
    var statusRe = "再入構";
    var error = "エラーが発生しました。係員は処理を行ってください。\n"
    var unregistered = "登録されていないカードです。入構できません。"
    var mismatch = "\nスプレッドでIDmを検索し、\n入退構状況を正しく入力し直してください。\nこのエラーは入構キャンパスと退構キャンパスが\n一致しない場合に表示されます。"
    var outlier = "\nシートに異常な値が記録されています。\nスプレッドシートを確認してください。"
    
    //現在時刻
    var date = new Date();
    var dateLog = (Utilities.formatDate(date, 'Asia/Tokyo', 'MM/dd HH:mm'));
    
    //idmをSSから探す
    var searchIdm = (idmArray.indexOf(idm)) + 1;//IDmの行数が出る
    
    if(searchIdm != ""){//IDmが見つかったら
 
      var nameRange = (nameArray[searchIdm -1]);//IDmに対応した名前を探す
      var statusMenber = (menberArray[searchIdm -1]);//IDmに対応した団体名を探す
      var statusRange = (statusArray[searchIdm -1]);//IDmに対応したステータスを探す
      var statusGate = (gateArray[searchIdm -1]);//IDmに対応したキャンパスを探す
      var statusTime = (timeArray[searchIdm -1]);//IDmに対応した前回入構時刻を探す
      

      if(statusRange == ""){//未記入または退構状態だったら
        var status = statusIn;//入構をセット
        sheet.getRange(searchIdm, 4).setValue(status);//セルに記入
        sheet.getRange(searchIdm, 5).setValue(gate);//セルに記入
        sheet.getRange(searchIdm, 6).setValue(dateLog);//時刻を記入
        
        //アプリに返す
        var returnText = "名　前：" + nameRange + "\n" + "団　体：" + statusMenber + "\n" + "状　態：" + statusIn + "\n" + "入構門：" + statusGate + "\n" + "時　刻：" + dateLog;
        rowData.value = returnText;
        return ContentService.createTextOutput(rowData.value).setMimeType(ContentService.MimeType.TEXT);

    
      }else if(statusRange == statusIn || statusRange == statusRe){//入構状態だったら      
        var range = sheet.getRange(searchIdm, 6);
        range.insertCells(SpreadsheetApp.Dimension.COLUMNS);
        sheet.getRange(searchIdm, 6).setValue(dateLog);//時刻を記入
        
        //アプリに返す
        var returnText = "名　前　：" + nameRange + "\n" + "団　体　：" + statusMenber + "\n" + "状　態　：" + statusRe + "\n" + "再入構門：" + statusGate + "\n" + "時　刻　：" + dateLog;
        rowData.value = returnText;
        return ContentService.createTextOutput(rowData.value).setMimeType(ContentService.MimeType.TEXT);
    
      }else{
        var status = error;
        rowData.value =  error + outlier;
        return ContentService.createTextOutput(rowData.value);

    
      }
      
  
    }else{//IDmが見つからなかったら
      var status = error;
      
      //アプリに返す
      rowData.value = error + unregistered;
      sheet3.getRange(1, 1).setValue(rowData.value);
      return ContentService.createTextOutput(rowData.value);
      
      }//IDmが登録されているかどうか

  }//読み取りエラーかどうか

}//全体


