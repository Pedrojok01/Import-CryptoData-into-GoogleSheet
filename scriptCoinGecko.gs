function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("CoinGecko")
    .addItem("Update", "CRYPTODATAJSON")
    .addToUi();
}

function CRYPTODATA(symbol, colName) {
  const sheetData = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("data")
    .getDataRange()
    .getValues();
  const col = sheetData[0].indexOf(colName);

  if (col === -1) return "datatype not found";

  const row = sheetData.findIndex((row, i) => i > 0 && row[2] === symbol);
  return row === -1 ? "symbol not found" : sheetData[row][col];
}

function CRYPTODATAJSON() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data");

  const urls = [
    "https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=250&page=1&sparkline=false",
    "https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=250&page=2&sparkline=false",
    "https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=250&page=3&sparkline=false",
  ];

  const fetchOptions = { muteHttpExceptions: true };
  const responses = UrlFetchApp.fetchAll(
    urls.map((url) => ({ url, ...fetchOptions }))
  );

  if (responses.some((res) => res.getResponseCode() === 429)) {
    ui.alert("too many requests");
    return;
  }

  if (responses.some((res) => res.getResponseCode() !== 200)) {
    ui.alert("server error : http " + res.getResponseCode());
    return;
  }

  const dataSet = responses.flatMap((res) => JSON.parse(res.getContentText()));

  const header = [
    "id",
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
    "last_updated",
  ];

  const rows = [
    header,
    ...dataSet.map((data) => header.map((col) => data[col])),
  ];
  const dataRange = sheet.getRange(1, 1, rows.length, rows[0].length);
  dataRange.setValues(rows);
}
