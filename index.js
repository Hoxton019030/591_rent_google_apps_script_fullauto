//調整你要搜尋的網址
const url = "https://rent.591.com.tw/?rentprice=3000,8000&order=posttime&orderType=desc&section=220&searchtype=1"
//寫入Line權杖
const token = "把你的LineToken放到這邊"
//調整你要搜尋的縣市，務必跟下面的吻合
const city = "台南市"
const cities = new Map([
  ["台北市", 1],
  ["新北市", 3],
  ["桃園市", 6],
  ["新竹市", 4],
  ["新竹縣", 5],
  ["宜蘭縣", 21],
  ["基隆市", 2],
  ["台中市", 8],
  ["彰化縣", 10],
  ["雲林縣", 14],
  ["苗栗縣", 7],
  ["南投縣", 11],
  ["高雄市", 17],
  ["台南市", 15],
  ["嘉義市", 12],
  ["嘉義縣", 13],
  ["屏東縣", 19],
  ["台東縣", 22],
  ["花蓮縣", 23],
  ["花蓮縣", 23],
  ["澎湖縣", 24],
  ["金門縣", 25],
  ["連江縣", 26]
]);



const titles = ['物件Id', '租屋標題', '區域', '路名', '社區', '月租', '坪數', '距離最近的地點', '網址'];
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('591租屋資料表')




/**
 * 主程序
 */
function main() {
  //取得目前Sheet的資料
  const historicalIdList = getHistoricalId();
  //取得591新的資料
  const jsonData = get591Result();
  //把591新的資料讀取出來
  const newIdList = getThisTimeId(jsonData)
  //把sheet的資料跟591出來的資料作對比，找出新的資料
  const uniqueNewIds = getDiffId(historicalIdList, newIdList)
  //如果有新的uniqueNewIds資料，發送通知給用戶
  if (uniqueNewIds.length != 0) {
    console.log(`${new Date()} 本次執行有${uniqueNewIds.length}筆資料不同，分別是${uniqueNewIds}`)
    // snedNotifyToUser(uniqueNewIds,jsonData)
    // 把Sheet的table清空，把新的資料寫進去
    initSheet();
    writeDataIntoSheet(jsonData);
  }
  //如果沒有新的uniqueNewIds資料，不發送通知給用戶
  if (uniqueNewIds.length == 0) {
    console.log(`${new Date()} 本次執行時沒有任何新的資料`)
  }
}
/**
 * 取得A欄的資料
 */
function getHistoricalId() {
  const dataRange = sheet.getRange('A2:A');
  const values = dataRange.getValues();
  const nonEmptyValues = values.filter(row => row[0] !== '');
  const pureNumbers = nonEmptyValues.map(row => row[0]);
  return pureNumbers;
}


function snedNotifyToUser(uniqueNewIds, jsonData) {
  const filtedJsonData = jsonData.filter(item => uniqueNewIds.includes(item.post_id));
  for (var i = 0; i < filtedJsonData.length; i++) {
    var item = filtedJsonData[i]
    let title = item.title
    let price = item.price
    let kind_name = item.kind_name
    let area = item.area
    let floor_str = item.floor_str
    let community = item.floor_str
    let location = item.location
    let url = "https://rent.591.com.tw/" + item.post_id
    let dest = item.surrounding.desc + item.surrounding.distance
    let photo = item.photo_list[0]

    message =
      `
     標題: ${title}
     價錢: ${price}
     房型: ${kind_name}
     坪數: ${area}
     樓層: ${floor_str}
     社區: ${community}
     地點: ${location}
     距離: ${dest}
     網址: ${url}
    `

    sendLineNotify(message, photo)


  }

}
function sendLineNotify(message) {
  var token = "zS6xp8rVUtCiL9Dr2NBskvW98rAoIMrFBoJ1NVMUJjQ"
  var options =
  {
    "method": "post",
    "payload": { "message": message },
    "headers": { "Authorization": "Bearer " + token }
  };
  UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}


function sendLineNotify(message, imagePath) {
  var options =
  {
    "method": "post",
    "payload": { "message": message, "imageFullsize": imagePath, "imageThumbnail": imagePath },
    "headers": { "Authorization": "Bearer " + token },
  };
  var response = UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}



function getDiffId(historicalIdList, newIdList) {
  var uniqueNewIds = [];
  var historicalIdSet = new Set(historicalIdList);
  for (var i = 0; i < newIdList.length; i++) {
    if (!historicalIdSet.has(newIdList[i])) {
      uniqueNewIds.push(newIdList[i]);
    }
  }
  return uniqueNewIds;
}

function getThisTimeId(jsonData) {
  var postIdList = [];
  for (var i = 0; i < jsonData.length; i++) {
    var item = jsonData[i]
    postIdList.push(item.post_id);
  }
  return postIdList;
}

function writeDataIntoSheet(jsonData) {
  for (var i = 0; i < jsonData.length; i++) {
    var item = jsonData[i];
    var row = i + 2; // 從第二行開始，所以column +2
    var a = 0;
    sheet.getRange(row, a += 1).setValue(item.post_id);
    sheet.getRange(row, a += 1).setValue(item.title);
    sheet.getRange(row, a += 1).setValue(item.section_name);
    sheet.getRange(row, a += 1).setValue(item.street_name);
    sheet.getRange(row, a += 1).setValue(item.community);
    sheet.getRange(row, a += 1).setValue(item.price);
    sheet.getRange(row, a += 1).setValue(item.area);
    sheet.getRange(row, a += 1).setValue(item.surrounding.desc + item.surrounding.distance);
    sheet.getRange(row, a += 1).setValue("https://rent.591.com.tw/" + item.post_id);
  }
}


function getCSRFTokenAndCookie() {
  var rootPath = "https://rent.591.com.tw";
  var response = UrlFetchApp.fetch(rootPath);
  var headers = response.getHeaders();
  var cookies = headers['Set-Cookie'];
  var regionValue = cities.get(city)


  cookies += "; urlJumpIp=" + regionValue;

  var tokenMatch = response.getContentText().match(/<meta name="csrf-token" content="([^"]+)"/);
  var csrfToken = tokenMatch ? tokenMatch[1] : null;

  return {
    csrfToken: csrfToken,
    cookies: cookies
  };
}

function get591Result() {


  var csrfInfo = getCSRFTokenAndCookie();
  var csrfToken = csrfInfo.csrfToken;
  var cookies = csrfInfo.cookies;

  if (!csrfToken) {
    throw new Error("CSRF token not found");
  }



  var headers = {
    "Cache-Control": "no-store, no-cache, must-revalidate",
    "Cache-Control": "no-cache, private",
    "Connection": "keep-alive",
    "Content-Encoding": "gzip",
    "Content-Type": "application/json",
    "Date": "Sun, 26 May 2024 06:24:36 GMT",
    "Expires": "Thu, 19 Nov 1981 08:52:00 GMT",
    "Pragma": "no-cache",
    "Server": "openresty",
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Accept-Language": "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7,zh-CN;q=0.6",
    "Cookie": cookies,
    "Referer": "https://rent.591.com.tw/?region=3&kind=2&section=37&searchtype=1",
    "X-CSRF-TOKEN": csrfToken,
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.85 Safari/537.36"
  };

  var options = {
    method: "GET",
    headers: headers,
    muteHttpExceptions: true
  };

  var requestUrl = "https://rent.591.com.tw/home/search/rsList?is_format_data=1&is_new_list=1&type=1&"
  const queryParams = url.match(/\?(.+)/);
  requestUrl = requestUrl + queryParams[1]

  // 使用正則表達式檢查URL中是否包含region參數
  var regionParamPattern = /[?&]region=\d+/;
  if (!regionParamPattern.test(requestUrl)) {
    // 如果不包含region參數，則將 region=3 添加到URL
    // 檢查URL是否已經有查詢參數，如果有就添加&，否則添加?
    if (requestUrl.indexOf('?') === -1) {
      requestUrl += 'region=' + cities.get(city);
    } else {
      requestUrl += 'region=' + cities.get(city);
    }
  }
  console.log(queryParams[1])
  console.log(requestUrl)
  var response = UrlFetchApp.fetch(requestUrl, options);

  var jsonData = JSON.parse(response.getContentText())

  return jsonData.data.data;
}


function initSheet() {
  // var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
  sheet.clear();

  // 恢復預設
  var lastRow = sheet.getMaxRows();
  var lastCol = sheet.getMaxColumns();
  var defaultRowHeight = 20.25
  var defaultColWidth = 64

  // 设置所有行的高度为默认值
  sheet.setRowHeights(1, lastRow, defaultRowHeight);

  // 设置所有列的宽度为默认值
  sheet.setColumnWidths(1, lastCol, defaultColWidth);


  for (var i = 0; i < titles.length; i++) {
    sheet.getRange(1, i + 1).setValue(titles[i]);
  }

  sheet.setColumnWidth(titles.indexOf('租屋標題') + 1, 300);
  sheet.setColumnWidth(titles.indexOf('距離最近的地點') + 1, 300);
  sheet.setColumnWidth(titles.indexOf('網址') + 1, 250);
  try {
    range.createFilter();
  } catch {

  }
}
