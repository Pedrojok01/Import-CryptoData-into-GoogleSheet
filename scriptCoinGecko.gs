function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CoinGecko')
    .addItem('Update', 'CRYPTODATAJSON')
    .addToUi();
}

function CRYPTODATA(symbol, colName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheet = ss.getSheetByName("data");
  var data = sheet.getDataRange().getValues();
  var col = data[0].indexOf(colName);
  if (col != -1) {
    var row = -1;
    for (var i = 0; i < data.length; i++) {
      if (data[i][2] == symbol) {
        row = i;
        break;
      }
    }
    if (row != -1) {
      return data[row][col];
    } else {
      return "symbol not found";
    }
  } else {
    return "datatype not found";
  }
}

async function CRYPTODATAJSON() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheet = ss.getSheetByName("data");

  var responses = await UrlFetchApp.fetchAll(
    [
      {
        'url': 'https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=250&page=1&sparkline=false',
        'muteHttpExceptions' : true
      },
      {
        'url': 'https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=250&page=2&sparkline=false',
        'muteHttpExceptions' : true
      },
      {
        'url': 'https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=250&page=3&sparkline=false',
        'muteHttpExceptions' : true
      }
    ]
  );

for (let i = 0; i < responses.length; i++) {
    if (responses[i].getResponseCode() == 429)
      {
        ui.alert("too many requests");
        return "";
      }
    else if (responses[i].getResponseCode() != 200)
      {
        ui.alert("server error : http " + (responses.getResponseCode()));
        return "";
      }
  };
  
  var jsons = responses.map(res => JSON.parse(res.getContentText()));
  var dataSet = jsons[0].concat(jsons[1], jsons[2]);

  var rows = [],
    data;
  var nbCol = 0;
  rows.push(["id",
    "name",
    "symbol",
    "market_cap_rank",
    "current_price",
    "market_cap",
    "circulating_supply",
    "total_supply",
    "max_supply",
    "total_volume",
    "high_24h",
    "low_24h",
    "price_change_24h",
    "price_change_percentage_24h",
    "market_cap_change_24h",
    "market_cap_change_percentage_24h",
    "usd_percent_change_1y",
    "ath",
    "ath_change_percentage",
    "ath_date",
    "atl",
    "atl_change_percentage",
    "atl_date",
    "last_updated"
  ]);

  for (i = 0; i < dataSet.length; i++) {
    data = dataSet[i];
    rows.push([data.id,
      data.name,
      data.symbol,
      data.market_cap_rank,
      data.current_price,
      data.market_cap,
      data.circulating_supply,
      data.total_supply,
      data.max_supply,
      data.total_volume,
      data.high_24h,
      data.low_24h,
      data.price_change_24h,
      data.price_change_percentage_24h,
      data.market_cap_change_24h,
      data.market_cap_change_percentage_24h,
      data.usd_percent_change_1y,
      data.ath,
      data.ath_change_percentage,
      data.ath_date,
      data.atl,
      data.atl_change_percentage,
      data.atl_date,
      data.last_updated,
    ]);
  }

  dataRange = sheet.getRange(1, 1, rows.length, rows[0].length);
  dataRange.setValues(rows);
}