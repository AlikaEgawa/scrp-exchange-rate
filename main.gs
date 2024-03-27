function myFunction() {
  var date_id = generateID();
  let base_URL = "https://www.murc-kawasesouba.jp/fx/past/index.php?id=";
  var id_URL = base_URL + date_id;
  // console.log(id_URL);
  let response = UrlFetchApp.fetch(id_URL);
  let content = response.getContentText("utf-8");

  var text = Parser.data(content).from('<tr>').to('</tr>').iterate();
  var textArray = [];

  for(var i=0; i<31; i++){
    textArray.push(text[i]);
  }

  var pattern1 = /<td class="t_center">[^<]*<\/td>/g;
  var pattern2 = /<td class="t_right">[^<]*<\/td>/g;
  var matchesArray = [];

  for(var i=0; i<textArray.length; i++){
    var column = []
    column.push(textArray[i].match(pattern1)[0]);
    column.push(textArray[i].match(pattern2)[0]);
    matchesArray.push(column);
  }

  var column_text = transformData(matchesArray);

  column_text.unshift(["Currency", "Start Date", "Expiry Date", "Rates", "2.0%UP"]);

  for(var i=1; i<column_text.length; i++) {
    var two_percent_up = column_text[i][1]*1.02;
    column_text[i].push(two_percent_up.toFixed(2));

    var start_date = createTomorrow();
    var expiry_date = createNextWeek();
    column_text[i].splice(1, 0, start_date);
    column_text[i].splice(2, 0, expiry_date);
  }

  // Logger.log(column_text);
  // Logger.log("================================");

  writeToSpreadsheet(column_text);

  sendHTMLemail(generateFileName(), id_URL);

}

function writeToSpreadsheet(array) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  for (var i = 0; i < array.length; i++) {
    var rowData = array[i];
    sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
  }
}

function createTomorrow(){
  var today = new Date();
  var tomorrow = new Date(today);
  tomorrow.setDate(today.getDate()+1);
  var tomorrowStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return tomorrowStr;
}

function createNextWeek(){
  var today = new Date();
  var next_week = new Date(today);
  next_week.setDate(today.getDate()+7);
  var nextWeekStr = Utilities.formatDate(next_week, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return nextWeekStr;
}

function generateFileName(){
  var today = new Date();
  var tomorrow = new Date(today);
  var next_week = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  next_week.setDate(today.getDate() + 7);
  var tomorrow_year = tomorrow.getFullYear() % 100;
  var tomorrow_month = ("0" + (tomorrow.getMonth() + 1)).slice(-2);
  var tomorrow_day = ("0" + tomorrow.getDate()).slice(-2);
  var next_week_year = next_week.getFullYear() % 100;
  var next_week_month = ("0" + (next_week.getMonth() + 1)).slice(-2);
  var next_week_day = ("0" + next_week.getDate()).slice(-2);

  var file_name = "Japan Exchange Rate_" + tomorrow_year + tomorrow_month + tomorrow_day + "-" + next_week_year + next_week_month + next_week_day;
  return file_name;
}

function sendHTMLemail(title, url) {
  var body = "Here is Spread Sheet URL!" + "\nhttps://docs.google.com/spreadsheets/d/1Kco6H8iUIEHM0KB4PNUTOlzfNBfeKA7RddS1NaXbrjs/edit?usp=sharing" + "\nHere is the URL based on : " + url;
  GmailApp.sendEmail('akiya.hikami@hillebrandgori.com' , title, body, {htmlBody:body});
  Logger.log("Completed")
}

function transformData(data) {
  var transformedData = [];
  for (var i = 0; i < data.length; i++) {
    var innerArray = [];
    for (var j = 0; j < data[i].length; j++) {
      var value = data[i][j].match(/>([^<]+)<\/td>/)[1];
      innerArray.push(value);
    }
    transformedData.push(innerArray);
  }
  return transformedData;
}

function generateID(){
  var today = new Date();
  var yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  var year = yesterday.getFullYear() % 100;
  var month = ("0" + (yesterday.getMonth() + 1)).slice(-2);
  var day = ("0" + yesterday.getDate()).slice(-2);

  var id = year + month + day;
  return id;
}
