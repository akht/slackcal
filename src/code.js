var CALENDAR_ID = PropertiesService.getScriptProperties().getProperty('CALENDAR_ID');
var CALENDAR_URL = PropertiesService.getScriptProperties().getProperty('CALENDAR_URL');
var SPREADSHEET_URL = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_URL');
var SLACK_WEBHOOK_URL = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
var SLACK_CHANNEL_NAME = PropertiesService.getScriptProperties().getProperty('SLACK_CHANNEL_NAME');
var SLACK_USER_NAME = PropertiesService.getScriptProperties().getProperty('SLACK_USER_NAME');
var LACK_USER_ICON = PropertiesService.getScriptProperties().getProperty('SLACK_USER_ICON');

function myFunction() {
  var calendarEventList = listCalendarEvents();
  var spreadsheetEventList = listSpreadsheetEvents();
  var eventList = calendarEventList.concat(spreadsheetEventList);
  var sortedEventList = eventList.sort(function(a,b){if (a < b){return -1;} else {return 1;}});
  var string = formatEventList(addDaysDiffToEventList(sortedEventList));

  postSlack(string);
}

function listCalendarEvents() {
  var events = [];
  var cal = CalendarApp.getCalendarById(CALENDAR_ID);
  var startDate = new Date();
  var endDate = new Date();
  endDate.setMonth(endDate.getMonth()+6);
  var calEvents = cal.getEvents(startDate, endDate);

  var list   = [];  // [actualStartDate, displayStartDate, title]
  var titles = [];

  calEvents.forEach (function (e) {
    var title       = e.getTitle();
    var startTime   = e.getStartTime();
    if (titles.indexOf(title) != -1) {
      return;
    }
    titles.push(title);

    list.push([
      Utilities.formatDate(startTime, "GMT+0900", "yyyy/MM/dd"),
      Utilities.formatDate(startTime, "GMT+0900", "yyyy/MM/dd"),
      title
    ]);
  });

  return list;
}

function listSpreadsheetEvents() {
  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var sheet = ss.getSheetByName("events");

  var startRow = 2, startColumn = 1;
  var numRows = sheet.getLastRow() - startRow + 1;
  var numColumns = sheet.getLastColumn() - startColumn + 1;
  var sheetValues = sheet.getSheetValues(startRow, startColumn, numRows, numColumns);

  var list = [];  // [actualStartDate, displayStartDate, title]
  var currentDate = new Date();

  for (var i = 0, i_len = sheetValues.length; i < i_len; ++i) {
    if (sheetValues[i][2] === "上旬") {
      actualStartDate = new Date(sheetValues[i][0], sheetValues[i][1] - 1, 1);
      if (actualStartDate >= currentDate) {
        list.push([
          Utilities.formatDate(actualStartDate, "GMT+0900", "yyyy/MM/dd"),
          sheetValues[i][0] + "/" + ('00' + sheetValues[i][1]).slice(-2) + " 上旬",
          sheetValues[i][3]
        ])
      }
    } else if (sheetValues[i][2] === "中旬") {
      actualStartDate = new Date(sheetValues[i][0], sheetValues[i][1] - 1, 11);
      if (actualStartDate >= currentDate) {
        list.push([
          Utilities.formatDate(actualStartDate, "GMT+0900", "yyyy/MM/dd"),
          sheetValues[i][0] + "/" + ('00' + sheetValues[i][1]).slice(-2) + " 中旬",
          sheetValues[i][3]
        ])
      }
    } else if (sheetValues[i][2] === "下旬") {
      actualStartDate = new Date(sheetValues[i][0], sheetValues[i][1] - 1, 21);
      if (actualStartDate >= currentDate) {
        list.push([
          Utilities.formatDate(actualStartDate, "GMT+0900", "yyyy/MM/dd"),
          sheetValues[i][0] + "/" + ('00' + sheetValues[i][1]).slice(-2) + " 下旬",
          sheetValues[i][3]
        ])
      }
    }　else if (sheetValues[i][2] === "末") {
      actualStartDate = new Date(sheetValues[i][0], sheetValues[i][1], 0);
      if (actualStartDate >= currentDate) {
        list.push([
          Utilities.formatDate(actualStartDate, "GMT+0900", "yyyy/MM/dd"),
          sheetValues[i][0] + "/" + ('00' + sheetValues[i][1]).slice(-2) + " 末",
          sheetValues[i][3]
        ])
      }
    }
  }
  return list;
}


function getDaysDiff(startDate) {
  var now = new Date()
  var currentDate = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0)

  var msDiff = new Date(startDate).getTime() - currentDate.getTime();
  var daysDiff = Math.floor(msDiff / (1000 * 60 * 60 *24));
  if (daysDiff < 0) {
    daysDiff = 0;
  }

  return daysDiff;
}


function addDaysDiffToEventList(eventList) {
  list = [];
  for (var i = 0; i < eventList.length; i++) {
    daysDiff = getDaysDiff(eventList[i][0])
    list.push([
      eventList[i][1],
      daysDiff,
      eventList[i][2]
    ])
  }

  return list;
}

function formatEventList(eventList) {
  string = "カレンダー：" + CALENDAR_URL + "\n"
  string += "アバウト予定：" + SPREADSHEET_URL + "\n\n";
  eventList.forEach (function (e) {
    var tmpString = "";

    tmpString += e[0];
    tmpString += " (" + e[1] + "日) ";
    tmpString += e[2];

    string += tmpString + "\n";
  });

  return string;
}


function postSlack(list) {
  if (list == "") {
    return;
  }

  var payload = {
    "text" : list,
    "channel" : SLACK_CHANNEL_NAME,
    "username" : SLACK_USER_NAME,
    "icon_url" : LACK_USER_ICON
  }

  var options = {
    "method" : "POST",
    "payload" : JSON.stringify(payload)
  }

  var response = UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
  var content  = response.getContentText("UTF-8");
}
