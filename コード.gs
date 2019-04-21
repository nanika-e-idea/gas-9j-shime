/* 支払い明細入力表　締処理スクリプト
 * Function List
 * -PaymentStatment : 支払い明細入力表の入力内容より、個人ごと月ごとの支払明細書を発行する処理
 * -createPDF : PDF作成関数　”https://www.virment.com/create-pdf-google-apps-script/”よりm(__)m
 * -cllst : 支払い明細書の入力内容をクリアする処理
 * [直近の更新履歴]
 * 2018/09/18 - setTrash処理不具合対応
 */
function PaymentStatement() {
  const //入力表のセル位置関係定数…処理に余裕があれば変数化してシートから取得の方が望ましい
      inputYMRow = 4,
      inputPayDateCol = 3,
      inputYearCol = 6,
      inputMonthCol = 8,
      maxRow = 999,
      inputSeqCol = 1,
      inputDateCol = 2,
      inputNameCol = 3,
      inputItemCol = 4,
      inputValueCol = 7,
      inputUnitCol = 8,
      inputFeeCol = 9,
      inputPathCol = 10,
      inputFolderIdCol = 11,
      inputNameSeqCol = 12,
      inputAccountCol = 13,
      inputSRow = 7,
      pathLabelCol = 1,
      pathValueCol = 2;

  const //支払明細書のセル位置関定数…こちらも処理に余裕があれば変数化が望ましい
      statementAccountRow = 39,
      statementAccountCol = 1,
      statementPayDateRow = 19,
      statementPayDateCol = 5,
      statementNameRow = 9,
      statementNameCol = 1,
      statementSRow = 23,
      statementERow = 34,
      statementSCol = 1,
      statementECol = 11,
      statementDateCol = 1,
      statementItemCol = 2,
      statementValueCol = 8,
      statementUnitCol = 9,
      statementFeeCol =11,
      statementExRange = 'A1%3AL39';
  
  var activeSheet = SpreadsheetApp.getActiveSheet(),
      flgCreateFolder = 1,
      flgCreateAllFile = 0,
      goSign,
      inputERow,
      maxDate,
      maxInputCol,
      minDate,
      minInputCol,
      objCopyFile = {},
      objCount = {},
      objFinish = {},
      objFiles,
      objFolder,
      objFolders,
      objName = {},
      objOutputDatePath,
      objOutputFilePath,
      objPathSheet,
      objSaveFile,
      objTemplateFile,
      objTemplatePath,
      objTotal = {},
      objWkFile = {},
      objWLog = {},
      objWork,
      outputFileName,
      outputFilePath,
      saveFileName,
      saveFilePath,
      seqName,
      statementSheetID,
      statementSSID,
      storeFolderRoot,
      storeFolderBranch,
      strAccount,
      strOutputDate,
      strSectionName,
      templateFileName,
      templateFilePath,
      templateBasePath,
      trgFileExist,
      trgFolderPath,
      valDate,
      valFee,
      valFileCount = 0,
      valItem,
      valMonth = activeSheet.getRange(inputYMRow,inputMonthCol).getValue(),
      valName,
      valPayDate　=activeSheet.getRange(inputYMRow,inputPayDateCol).getValue(),
      valUnit,
      valValue,
      valYear = activeSheet.getRange(inputYMRow,inputYearCol).getValue(),
      zzzVarEnd;
  
  //処理月判定用に開始日と終了日を生成
  minDate = new Date(valYear,valMonth-1,1);
  maxDate = new Date();
  maxDate.setYear(minDate.getFullYear());
  maxDate.setMonth(minDate.getMonth() + 1);
  maxDate.setDate(0);
  strOutputDate = valPayDate.getYear() + "/" + (valPayDate.getMonth() + 1) + "/" + valPayDate.getDate() + "支払い分";
  
  
  //PDF出力先フォルダID,基本テンプレートID,経理フォルダID,部門名取得
  objPathSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フォルダID");
  for(var i = 1; i < maxRow; i++){
    var tmpStr = objPathSheet.getRange(i,pathLabelCol).getValue();
    var tmpValue = objPathSheet.getRange(i,pathValueCol).getValue();
    if( tmpStr== "支払明細書"){
      outputFilePath = tmpValue;
    }else if( tmpStr == "基本テンプレート"){
      templateBasePath =  tmpValue;
    }else if( tmpStr == "経理フォルダ"){
      storeFolderRoot =  tmpValue;
    }else if( tmpStr == "部門名"){
      strSectionName =  tmpValue;
    }
  }
  
  goSign = Browser.msgBox(strOutputDate + "の支払明細書を作成しますか？\\n（処理対象月：" + valYear + "年"+ valMonth + "月）",Browser.Buttons.OK_CANCEL);
  if (goSign == 'ok'){
    //エンドマーク位置取得
    var dat = activeSheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
    for(var i=inputSRow;i<dat.length;i++){
      if(dat[i][inputSeqCol-1] === "!処理ココまで!"){
        inputERow = i - 1;
        break;
      }
    };
  
    //（ループ前）経理フォルダ書き込み側用意
    ////日付フォルダ有無判定
    flgCreateFolder = 1;
    objFolders = DriveApp.getFolderById(storeFolderRoot).getFolders();
    while(objFolders.hasNext()){
      objFolder = objFolders.next();
      if(objFolder.getName() == strOutputDate){
        trgFolderPath = objFolder.getId();
        flgCreateFolder = 0;
      }
    }
    //////なければ作る
    if(flgCreateFolder == 1){
      trgFolderPath = DriveApp.getFolderById(storeFolderRoot).createFolder(strOutputDate).getId();
    }
    //////あれば次へ
    ////部門別フォルダ有無確認
     flgCreateFolder = 1;
    objFolders = DriveApp.getFolderById(trgFolderPath).getFolders();
    while(objFolders.hasNext()){
      objFolder = objFolders.next();
      if(objFolder.getName() == strSectionName){
        storeFolderBranch = objFolder.getId();
        flgCreateFolder = 0;
      }
    }
    //////なければ作る
    if(flgCreateFolder == 1){
      storeFolderBranch = DriveApp.getFolderById(trgFolderPath).createFolder(strSectionName).getId();
    }
    trgFolderPath = "";　//別処理で事故らないように空にしとく
    //入力表範囲指定
    maxInputCol = Math.max(inputDateCol,inputNameCol,inputItemCol,inputValueCol);
    minInputCol = Math.min(inputDateCol,inputNameCol,inputItemCol,inputValueCol);
    //入力表を日付昇順ソート
    activeSheet.getRange(inputSRow,minInputCol,inputERow - inputSRow + 1,maxInputCol - minInputCol + 1).sort(inputDateCol);
    //支払明細書転記処理
    for(var i = inputSRow; i < inputERow + 1; i ++){
      strAccount = activeSheet.getRange(i,inputAccountCol).getValue();
      valName = activeSheet.getRange(i,inputNameCol).getValue();
      valDate = activeSheet.getRange(i,inputDateCol).getValue();
      seqName = activeSheet.getRange(i,inputNameSeqCol).getValue();
      if(valDate == ""){break;};　//日付欄が空白ならばそこで処理終了…日付でソート済ならこれでいいはず
      if(valName == ""){break;};  //念のためネーム欄でも同様の処理
      //日付が処理月内でなければcontinue ←実際の使い方とそぐわないようなのでオミット
      //if(valDate < minDate || valDate > maxDate){ continue };
      //数量が0ならcontinue
      valValue = activeSheet.getRange(i,inputValueCol).getValue();
      if(valValue == 0){ continue };
      //テンプレートを開く
      templateFileName = activeSheet.getRange(i,inputNameCol).getValue().replace(/\s|　/g, ""); //空白を削除
      templateFilePath = activeSheet.getRange(i,inputFolderIdCol).getValue();
      objFiles = DriveApp.getFolderById(templateFilePath).getFiles();
      trgFileExist = 0;
      while(objFiles.hasNext()){
        　var tempObj = objFiles.next();
        if(!tempObj.getName().indexOf(templateFileName)){
          objTemplateFile = tempObj;
          trgFileExist = 1;
        }
      }
      //テンプレートがなければcontenue
      if(objFinish[seqName] === void 0){
        objFinish[seqName] = 0;
      }else if(objFinish[seqName] == 1){
        continue;
      }
      if(trgFileExist == 0){
　　　　　var objTemplateBase = DriveApp.getFileById(templateBasePath);
        objTemplateFile = objTemplateBase.makeCopy(templateFileName, DriveApp.getFolderById(templateFilePath));
        objWkFile[seqName] = objTemplateFile;
      };
      objWork = SpreadsheetApp.open(objTemplateFile).getSheets()[0];
      if(trgFileExist == 0){
        objWork.getRange(statementNameRow,statementNameCol).setValue(valName + " 様");
        trgFileExist = 1;
      }
      //ひとりに複数の明細があるケースを想定してひとりあたり明細件数をカウント
      if(objCount[seqName] === void 0){
        objCount[seqName] = 0;
        //テンプレートを初期化
        objWork.getRange(statementSRow,statementSCol,statementERow - statementSRow + 1, statementECol - statementSCol + 1).clearContent();
        //なぜか日付がずれるようなので対策(処理時刻＋8Hで時刻設定されるっぽい)
        var wrPayDate = Utilities.formatDate(valPayDate, "JST", "yyyy年M月d日");
        
        //テンプレートに支払日入力
        objWork.getRange(statementPayDateRow,statementPayDateCol).setValue(wrPayDate);
        //テンプレートに口座情報入力
        objWork.getRange(statementAccountRow,statementAccountCol).setValue(strAccount);
        
        objName[seqName] = valName;
      }; //end IF()
      //値取得
      //ValDate,valName取得済み
      valItem = activeSheet.getRange(i,inputItemCol).getValue();
      valValue = activeSheet.getRange(i,inputValueCol).getValue();
      valUnit = activeSheet.getRange(i,inputUnitCol).getValue();
      valFee = activeSheet.getRange(i,inputFeeCol).getValue();
      
      var wrDate = Utilities.formatDate(valDate, "JST", "M/d");
      
      //ひとり当たりの合計支払金額をカウント
      if(objTotal[seqName] === void 0){
        objTotal[seqName] = 0;
      };
      objTotal[seqName] += valFee;
      //テンプレートに値記入
      objWork.getRange(statementSRow + objCount[seqName],statementDateCol).setValue(wrDate);
      objWork.getRange(statementSRow + objCount[seqName],statementItemCol).setValue(valItem);
      objWork.getRange(statementSRow + objCount[seqName],statementValueCol).setValue(valValue);
      objWork.getRange(statementSRow + objCount[seqName],statementUnitCol).setValue(valUnit);
      objWork.getRange(statementSRow + objCount[seqName],statementFeeCol).setValue(valFee);
      //後の処理のためobjWorkをobjWLog[]に格納うまく動かない場合はフォルダIDとファイル名をそれぞれ格納
      objWLog[seqName] = objWork;
      
      //明細カウント加算
      objCount[seqName] ++;
    }; //end FOR()
    
    //PDF出力処理
    var numVer = activeSheet.getRange(4,4).getValue();
    objOutputFilePath = DriveApp.getFolderById(outputFilePath);
    //印刷後PDF保管用フォルダ生成
    //目的フォルダ存在確認、なければ生成
    objFolders = DriveApp.getFolderById(outputFilePath).getFolders();
    while(objFolders.hasNext()){
      objFolder = objFolders.next();
      if(objFolder.getName() == strOutputDate){
        trgFolderPath = objFolder.getId();
        flgCreateFolder = 0;
      }
    }
    if(flgCreateFolder == 1){
      trgFolderPath = DriveApp.getFolderById(outputFilePath).createFolder(strOutputDate).getId();
    }
    
    //objCount[]の値が0より大きければ
    //key値を頼りに
    for(key in objCount){
      if(objCount[key] > 0){
        outputFileName = objName[key] + "_" + strOutputDate + "-" + numVer;
        //同名ファイルがあれば上書き
        objFiles = DriveApp.getFolderById(trgFolderPath).getFiles();
        while(objFiles.hasNext()){
          var file = objFiles.next();
        　if(file.getName() == outputFileName + ".pdf" ){
           //既存ファイル削除
           file.setTrashed(true);
          }
        }
        objFiles = DriveApp.getFolderById(storeFolderBranch).getFiles();
        while(objFiles.hasNext()){
          var file = objFiles.next();
        　if(file.getName() == outputFileName + ".pdf" ){
           //既存ファイル削除
           file.setTrashed(true);
          }
        }
      }
      //PDF出力
      //【未実装】ターゲットフォルダに同名ファイルがある場合は、出力ファイル名を変更
      //もしくは、はじめからファイル名に識別子をつけておくか（FileName_pubYYYYMMDDHHMMS）
      
      var tmpFPath = objWLog[key].getParent().getId();
      var tmpSId = objWLog[key].getSheetId();
      var tmpstr = objWLog[key].getRange(statementPayDateRow,statementPayDateCol).getValue();
      //出力先選択
      var pdfFolderPath;
      var tlgSelectFoler = activeSheet.getRange(4,10).getValue();
      if (tlgSelectFoler == "経理フォルダ"){
        pdfFolderPath = storeFolderBranch;
      }else{
        pdfFolderPath = trgFolderPath;
      }
      createPDF(pdfFolderPath, tmpFPath, tmpSId, outputFileName, statementExRange);
    }
  } else {
    Browser.msgBox("処理を中止しました");
    return false;
  } //end If(goSign == 'ok')
  activeSheet.getRange(4,4).setValue(parseInt(numVer, 10) + 1);
  Browser.msgBox("処理を終了しました");
} //end Function PaymentStatement()

function createPDF(folderid, ssid, sheetid, filename ,range){

  // PDFファイルの保存先フォルダID指定
  var folder = DriveApp.getFolderById(folderid);

  // スプレッドシートをPDFにエクスポートするURL
  var url = "https://docs.google.com/spreadsheets/d/SSID/export?".replace("SSID", ssid);

  // PDF作成のオプションを指定 ※課題：メモの非表示オプションが不明
  var opts = {
    exportFormat: "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
    format:       "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
    size:         "A4",     // 用紙サイズの指定 legal / letter / A4
    portrait:     "true",   // true → 縦向き、false → 横向き
    fitw:         "true",   // 幅を用紙に合わせるか
    sheetnames:   "false",  // シート名をPDF上部に表示するか
    printtitle:   "false",  // スプレッドシート名をPDF上部に表示するか
    pagenumbers:  "false",  // ページ番号の有無
    gridlines:    "false",  // グリッドラインの表示有無
    fzr:          "false",  // 固定行の表示有無
    range:        range,    // 出力範囲を指定
    gid:          sheetid   // シートIDを指定 sheetidは引数で取得
  };
  
  var url_ext = [];
  
  // 上記のoptsのオプション名と値を「=」で繋げて配列url_extに格納
  for( optName in opts ){
    url_ext.push( optName + "=" + opts[optName] );
  }

  // url_extの各要素を「&」で繋げる
  var options = url_ext.join("&");

  // API使用のためのOAuth認証
  var token = ScriptApp.getOAuthToken();
    // PDF作成
    var response = UrlFetchApp.fetch(url + options, {
      headers: {
        'Authorization': 'Bearer ' +  token
      },
    "muteHttpExceptions" : true
    });

    // 
    var blob = response.getBlob().setName(filename + '.pdf');

  //}

  //　PDFを指定したフォルダに保存
  folder.createFile(blob);
  
} //end Function createPDF()


function cllst(){
  const
      sCol = 2,
      sRow = 7,
      eCol = 7,
      eRow = 500,
      shimeCol = 3,
      shimeRow = 4;
  var activeSheet = SpreadsheetApp.getActiveSheet(),
      trgCleraVer;
  activeSheet.getRange(sRow,sCol,eRow,eCol).clearContent();
  activeSheet.getRange(shimeRow,shimeCol).clearContent();
  trgCleraVer = Browser.msgBox("バージョンナンバーを「１」に戻しますか？",Browser.Buttons.OK_CANCEL);
  if (trgCleraVer == 'ok'){
    //バージョンクリア
    activeSheet.getRange(4,4).setValue(parseInt(1));
  }
}