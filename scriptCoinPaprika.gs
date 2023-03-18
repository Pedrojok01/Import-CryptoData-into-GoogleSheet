function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("CoinPaprika")
    .addItem("Update", "CRYPTODATAJSON")
    .addToUi();
}

function CRYPTODATA(symbol, colName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data");
  var data = sheet.getDataRange().getValues();
  var colIndex = data[0].indexOf(colName);

  if (colIndex === -1) {
    return "datatype not found";
  }

  for (var i = 1; i < data.length; i++) {
    if (data[i][2] === symbol) {
      return data[i][colIndex];
    }
  }

  return "symbol not found";
}

function CRYPTODATAHISTORY(symbol, date, type, quote) {
  quote = quote || "usd";

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data");
  const data = sheet.getDataRange().getValues();
  const colIndex = data[0].indexOf("id");

  if (colIndex === -1) {
    return "datatype not found";
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === symbol) {
      const coinId = data[i][colIndex];
      const formattedDate = date instanceof Date ? date.toISOString() : date;
      const url = `https://api.coinpaprika.com/v1/tickers/${coinId}/historical?start=${formattedDate}&interval=5m&limit=1&quote=${quote}`;

      const responseData = fetchUrl(url);
      return responseData[0].hasOwnProperty(type)
        ? responseData[0][type]
        : "price not found";
    }
  }

  return "symbol not found";
}

function CRYPTODATACOINDETAILS(coin_id, colName) {
  const param = encodeURI(coin_id);

  if (colName.indexOf("/") !== -1) {
    const [quote, property] = colName.split("/");
    const url =
      "https://api.coinpaprika.com/v1/tickers/" + param + "?quotes=" + quote;
    const responseData = fetchUrl(url);

    if (
      responseData.hasOwnProperty("error") &&
      responseData["error"] === "id not found"
    ) {
      return "coin_id not found";
    }

    if (responseData["quotes"][quote].hasOwnProperty(property)) {
      return responseData["quotes"][quote][property];
    } else {
      return "property not found";
    }
  } else {
    const url = "https://api.coinpaprika.com/v1/coins/" + param;
    const responseData = fetchUrl(url);

    if (
      responseData.hasOwnProperty("error") &&
      responseData["error"] === "id not found"
    ) {
      return "coin_id not found";
    }

    if (responseData.hasOwnProperty(colName)) {
      return responseData[colName];
    } else {
      return "property not found";
    }
  }
}

function CRYPTODATAJSON() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("data");

  const url = "https://api.coinpaprika.com/v1/tickers?quotes=BTC,USD,ETH";
  const dataAll = fetchUrl(url);
  const dataSet = dataAll.sort((a, b) => (a.rank > b.rank ? 1 : -1));

  const rows = [headers];

  dataSet.forEach((data) => {
    const rowData = headers.map((header) => {
      const [quote, property] = header.split("_");
      if (["BTC", "USD", "ETH"].includes(quote) && data.quotes[quote]) {
        return data.quotes[quote][property] || data[header];
      }
      return data[header];
    });
    rows.push(rowData);
  });

  const dataRange = sheet.getRange(1, 1, rows.length, rows[0].length);
  dataRange.setValues(rows);
}

/* UTILS:
 ***********/

function fetchUrl(url) {
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  RESPONSECODE(response);
  return JSON.parse(response.getContentText());
}

function RESPONSECODE(v) {
  var responseCode = v.getResponseCode();

  if (responseCode === 429) {
    throw new Error("Too many requests");
  } else if (responseCode !== 200) {
    throw new Error("Server error");
  }
}

/* CRYPTODATAJSON() - DATA HEADERS
 **********************************/

const headers = [
  "id",
  "name",
  "symbol",
  "rank",
  "circulating_supply",
  "total_supply",
  "max_supply",
  "beta_value",
  "btc_price",
  "btc_volume_24h",
  "btc_volume_24h_change_24h",
  "btc_market_cap",
  "btc_market_cap_change_24h",
  "btc_percent_change_1h",
  "btc_percent_change_12h",
  "btc_percent_change_24h",
  "btc_percent_change_7d",
  "btc_percent_change_30d",
  "btc_percent_change_1y",
  "btc_ath_price",
  "btc_ath_date",
  "btc_percent_from_price_ath",
  "usd_price",
  "usd_volume_24h",
  "usd_volume_24h_change_24h",
  "usd_market_cap",
  "usd_market_cap_change_24h",
  "usd_percent_change_1h",
  "usd_percent_change_12h",
  "usd_percent_change_24h",
  "usd_percent_change_7d",
  "usd_percent_change_30d",
  "usd_percent_change_1y",
  "usd_ath_price",
  "usd_ath_date",
  "usd_percent_from_price_ath",
  "eth_price",
  "eth_volume_24h",
  "eth_volume_24h_change_24h",
  "eth_market_cap",
  "eth_market_cap_change_24h",
  "eth_percent_change_1h",
  "eth_percent_change_12h",
  "eth_percent_change_24h",
  "eth_percent_change_7d",
  "eth_percent_change_30d",
  "eth_percent_change_1y",
  "eth_ath_price",
  "eth_ath_date",
  "eth_percent_from_price_ath",
  "last_updated",
];
