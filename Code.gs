/**
*Создаем меню для обновления списка импорта
*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  var subMenu  = ui.createMenu('Импорт')
      .addItem('Транзакций', 'importData')
      .addItem('Меток Отправки', 'importDataSend')
      .addItem('Меток Получения', 'importDataReceiv')
  
  ui.createMenu('Импорт BTC')
      .addItem('Загрузить btcData.csv', 'doGet')
      .addItem('Обновить Курс (Poloniex)', 'updatePoloniex')
      .addSeparator()
      .addSubMenu(subMenu)
      .addToUi();
}

/**
* Загружаем файл тарнзакций  для импорта в папку BtcImport на гугле Диске
*/
function doGet() {
  var html = HtmlService.createHtmlOutputFromFile('Form').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  return SpreadsheetApp.getUi().showModalDialog(html, 'Загрузка btcData.csv');
}


/**
* Возвращает Id Папки если она есть или false(если не нашла).
*/
function getFolderId(strFolder) {
  //strFolder = (strFolder || "BtcImport"); //Для теста.
  var folders = DriveApp.getFolders();
  
  while(folders.hasNext())
  {
      var folder = folders.next();

      // Находим папку "BtcImport"
      if(folder.getName() === strFolder)
        return folder.getId();
  }
  return false;
}

/**
* Форма для загрузки файла
*/
function processForm(theForm) {
  var fileBlob = theForm.picToLoad;

  Logger.log("fileBlob Name: " + fileBlob.getName())
  Logger.log("fileBlob type: " + fileBlob.getContentType())
  Logger.log('fileBlob: ' + fileBlob);
  
  var fldrSssn = DriveApp.getFolderById(getFolderId("BtcImport"));
    fldrSssn.createFile(fileBlob);

  return true;
}


/**
* Обновляет курсы валют в Листе Курсы с биржи poloniex.com
*/
function updatePoloniex()
{
  var response = UrlFetchApp.fetch("https://poloniex.com/public?command=returnTicker");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Курсы");
  var json = JSON.parse(response.getContentText());
  var rateETH = parseFloat(json.BTC_ETH.last);
  var rateDOGE = parseFloat(json.BTC_DOGE.last);
  var rateRDD = parseFloat(json.BTC_RDD.last);
  
  sheet.getRange(2, 1).setValue(rateETH);
  sheet.getRange(3, 1).setValue(rateDOGE);
  sheet.getRange(4, 1).setValue(rateRDD);
  
}

/**
* Импорт меток получения
*/
function importDataReceiv() {
    var fSource = DriveApp.getFolderById(getFolderId("BtcImport")); // id папки где лежит CSV фалик для импорта
    var fReceiv = fSource.getFilesByName('receivAddr.csv'); // Текущий файл для импорта
    
    if ( fReceiv.hasNext() ) 
    { // проверяем есть ли "receivAddr.csv" файл в папке для импорта, то есть "BtcImport"
        var file = fReceiv.next();
        var csv = file.getBlob().getDataAsString("CP1251");
        var csvData = CSVToArray(csv); // Передаю как параметр CVS в функцию  CSVToArray
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var btcImportSheetDohod = ss.getSheetByName('Доход');
        var labelTranzact = "";
        
        // Пустой масив для наполнения нужными данными
        var parsArrTranzactLabel = [];
        
        
        // В цикле перебираю даные и строю новый масив для листа, с нужными мне данными 
       for (var i = 0, lenCsv = csvData.length; i < lenCsv; i++ )
       {
           if (i != 0)
           {
               labelTranzact = csvData[i][0];
               parsArrTranzactLabel[i] = labelTranzact;
           }
       }
       
       parsArrTranzactLabel = array_unique(parsArrTranzactLabel);
       parsArrTranzactLabel = parsArrTranzactLabel.sort();
       
       for (var x = 0, lenTrz = parsArrTranzactLabel.length; x < lenTrz; x++ )
           if (x != 0)
           { 
             parsArrTranzactLabel[x] = new Array(parsArrTranzactLabel[x]);
             btcImportSheetDohod.getRange(x+1, 1, 1, parsArrTranzactLabel[x].length).setValues(new Array(parsArrTranzactLabel[x]));
           }
    }      
} 

/**
* Импорт меток отправки 
*/
function importDataSend() {
    var fSource = DriveApp.getFolderById(getFolderId("BtcImport")); // id папки где лежит CSV фалик для импорта
    var fSend = fSource.getFilesByName('sendAddr.csv'); // Текущий файл для импорта
    
    if ( fSend.hasNext() ) 
    { // проверяем есть ли "sendAddr.csv" файл в папке для импорта, то есть "BtcImport"
        var file = fSend.next();
        var csv = file.getBlob().getDataAsString("CP1251");
        var csvData = CSVToArray(csv); // Передаю как параметр CVS в функцию  CSVToArray
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var btcImportSheetRazhod = ss.getSheetByName('Расход');
        var labelTranzact = "";
        
        // Пустой масив для наполнения нужными данными
        var parsArrTranzactLabel = [];
        
        
        // В цикле перебираю даные и строю новый масив для листа, с нужными мне данными 
       for (var i = 0, lenCsv = csvData.length; i < lenCsv; i++ )
       {
           if (i != 0)
           {
               labelTranzact = csvData[i][0];
               parsArrTranzactLabel[i] = labelTranzact;
           }
       }
       
       parsArrTranzactLabel = array_unique(parsArrTranzactLabel);
       parsArrTranzactLabel = parsArrTranzactLabel.sort();
       
       for (var x = 0, lenTrz = parsArrTranzactLabel.length; x < lenTrz; x++ )
           if (x != 0)
           { 
             parsArrTranzactLabel[x] = new Array(parsArrTranzactLabel[x]);
             btcImportSheetRazhod.getRange(x+1, 1, 1, parsArrTranzactLabel[x].length).setValues(new Array(parsArrTranzactLabel[x]));
           }
    }      
} 

/**
* Импортирует данные транзакций CSV с Гугле Диска в  BtcImport Лист.
*/
function importData() {
  var btcFileImport;
  var fSource = DriveApp.getFolderById(getFolderId("BtcImport")); // id папки где лежит CSV фалик для импорта
  var fi = fSource.getFilesByName('btcData.csv'); // Текущий файл для импорта
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Какую табицу обновляем, текущуюю, активную
  var isExist = false;

  if ( fi.hasNext() ) 
  { // проверяем есть ли "btcData..csv" файл в папке для импорта, то есть "BtcImport"
    var file = fi.next();
    
   //Utilities.newBlob("").setDataFromString(csvData[i][2], "UTF-8").getDataAsString("CP1251"); Для теста. 
   var csv = file.getBlob().getDataAsString("CP1251");
    
   var csvData = CSVToArray(csv); // Передаю как параметр CVS в функцию  CSVToArray
   var btcImportSheet = ss.getSheetByName('BtcImport');
   
   if (null != btcImportSheet)
   {
       btcImportSheet.clearContents();
       isExist = true;
   }
   else
   {
     var newsheet = ss.insertSheet('BtcImport'); // создаю 'BtcImport' лист с импортированных даных
     ss.setNamedRange('BtcImp', newsheet.getRange("A2:E2000")); // Создаю именованный диапазон
     isExist = false; // Флаг на созданый или нет.

     // Переимновую имнованый диапазн.
     var namedRanges = ss.getNamedRanges();
     if (namedRanges.length > 1)
         namedRanges[0].setName("BtcImp");
   }
   // в цикле данные csv переношу в масив  и вставляю как строки  в 'BtcImport'
   // заношу даные в временные строки для наполнения нового масива.
   var dateTranzact = ""; //Дата Транзакции
   var typeTranzact = ""; // Сумма -//-
   var metkaTranzact = ""; //Метка -//-
   var summaTranzactDohod = ""; //Сумма -//- (BTC) Доход-Получено
   var summaTranzactRazhod = ""; //Сумма -//- (BTC) Расход-Отправлено
   // var idTranzact = ""; -- Пока не использую, может когда то, что бы смотреть сразу на блокчкйне операцию ;-) 
   
   // Пустой масив для наполнения нужными данными
   var parsArrTranzact = [];
   
   
   // В цикле перебираю даные и строю новый масив для листа, с нужными мне данными 
   for (var i = 0, lenCsv = csvData.length; i < lenCsv; i++ )
   {
      if (0 == i)
        parsArrTranzact[i] = new Array("Дата", "Тип", "Метка", "Сумма (BTC) Доход", "Сумма (BTC) Расход"); 
      else
      {
        dateTranzact = isoToDate(csvData[i][1]);
        
        // Если транзакция получения, то тогда в одну переменную для прихода 
        if (csvData[i][2] == "Получено")
          summaTranzactDohod = parseFloat(csvData[i][5]);
        else
        {
          if (!csvData[i][5] || 0 === csvData[i][5].length)
             summaTranzactRazhod = "Not Float"
          else
            summaTranzactRazhod = parseFloat((csvData[i][5]).substring(1));
        }
        
        typeTranzact = csvData[i][2];
        metkaTranzact = csvData[i][3];
        
        parsArrTranzact[i] = new Array (dateTranzact, typeTranzact, metkaTranzact, summaTranzactDohod, summaTranzactRazhod);
        
        // Обнуаляю переменные для новый даных в цикле
        summaTranzactRazhod ='';
        summaTranzactDohod ='';
       }


    }

    // Наполняю построчно лист импорта
    for (var x = 0, lenTrz = parsArrTranzact.length; x < lenTrz; x++ )
      if (isExist)
          btcImportSheet.getRange(x+1, 1, 1, parsArrTranzact[x].length).setValues(new Array(parsArrTranzact[x]));
      else
          newsheet.getRange(x+1, 1, 1, parsArrTranzact[x].length).setValues(new Array(parsArrTranzact[x]));
   
    // Для ТЕСТА-- удаляю Лист импорта, что бы не удалять вручную.
    //ss.deleteSheet(newsheet);
    // Переименовую файл "btcData..csv" что бы загружать новый и сохранять старые.
    file.setName("btcData-"+(new Date().toString())+".csv");
    }
}


// Функцию взял с интернета и переделал  немного, ссылку и коментарии осталяю.
// http://www.bennadel.com/blog/1504-Ask-Ben-Parsing-CSV-Strings-With-Javascript-Exec-Regular-Expression-Command.htm
// This will parse a delimited string into an array of
// arrays. The default delimiter is the comma, but this
// can be overriden in the second argument.

// Тут получаем данные и разделитель, по умолчанию запятая.
function CSVToArray( strData, strDelimiter ) {
  // Check to see if the delimiter is defined. If not,
  // then default to COMMA.
  strDelimiter = (strDelimiter || ",");
  
  // Create a regular expression to parse the CSV values.
  var objPattern = new RegExp(
    (
      // Delimiters.
      "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +

      // Quoted fields.
      "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +

      // Standard fields.
      "([^\"\\" + strDelimiter + "\\r\\n]*))"
    ),
    "gi"
  );

  // Create an array to hold our data. Give the array
  // a default empty first row.
  var arrData = [[]];

  // Create an array to hold our individual pattern
  // matching groups.
  var arrMatches = null;

  // Keep looping over the regular expression matches
  // until we can no longer find a match.
  while (arrMatches = objPattern.exec( strData ))
  {
  
    // Get the delimiter that was found.
    var strMatchedDelimiter = arrMatches[ 1 ];

    // Check to see if the given delimiter has a length
    // (is not the start of string) and if it matches
    // field delimiter. If id does not, then we know
    // that this delimiter is a row delimiter.
    if (strMatchedDelimiter.length && (strMatchedDelimiter != strDelimiter))
    {

      // Since we have reached a new row of data,
      // add an empty row to our data array.
      arrData.push( [] );

    }

    // Now that we have our delimiter out of the way,
    // let's check to see which kind of value we
    // captured (quoted or unquoted).
    if (arrMatches[2])
    {

      // We found a quoted value. When we capture
      // this value, unescape any double quotes.
      var strMatchedValue = arrMatches[2].replace(
        new RegExp( "\"\"", "g" ),
        "\""
      );

    } 
    else
    {
      // We found a non-quoted value.
      var strMatchedValue = arrMatches[3];
    }

    // Now that we have our value string, let's add
    // it to the data array.
    arrData[arrData.length - 1].push(strMatchedValue);
  }

  // Return the parsed data.
  return(arrData);
};

/**
*Обрезаем и Конвертируем дату.
*/
function isoToDate(dateStr) {// аргумент = данные в формате строки
  if (RegExp(/\d{4}-[01]\d-[0-3]\dT[0-2]\d:[0-5]\d:[0-5]\d/).exec(dateStr))
  {
    var str = dateStr.substring(0, dateStr.length - 9).replace(/-/,'/').replace(/-/,'/');
    return new Date(str);
  }
  else
    return "No Date"; //new Date().getTime();
}


/**
*В масиве оставляет только уникальные значения.
*/
function array_unique(inArr){
  var uniHash={}, outArr=[], i=inArr.length;
  while (i--) 
      uniHash[inArr[i]]=i;
  for (i in uniHash) 
      outArr.push(i);
  return outArr
}
