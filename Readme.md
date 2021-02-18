# Import-CryptoData-into-GoogleSheet
Fetch cryptodata from either [Coingecko](https://www.coingecko.com) or [Coinpaprika](https://coinpaprika.com/) into Google sheets for free. Enjoy!

## Looking to import crypto data into your google sheet?

You can use those free scripts to import data from either [CoinGecko API](https://www.coingecko.com/en/api) or [CoinPaprika API](https://coinpaprika.com/api/).

## Table of contents
* [General info](#general-info)
* [Which one do I need?](#Which-one-do-I-need?)
* [Setup](#setup)
* [Additional explanations](#Additional-explanations)
* [Update data](#Update-data)

# General info?

In order to go soft on those 2 great and free API, the script fetch all data at once in a dedicated tab in your spreadsheet. You can then simply import any data you need into your coinfolio/project by using this function: `CRYPTODATA("symbol"; "data")`.

# Which one do I need?

The CoinGecko API import the first 750 coins by marketcap. However, prices are only quoted in USD. (more coins but less data)
The CoinPaprika APi import all coins listed on coinpaprika.com, and prices are quoted in USD, BTC, or ETH. However, I found out that many coins are still missing from their great API. (more data but some coins missing)

>Notes:</br>
1/ As of now, you can only use one script at a time. So depending on what you need, you can either pick the Coingecko.gs or Coinpaprika.gs. </br>
2/ The performance isn't great and it takes a little time to refresh, but it's free to use and it goes soft on both API (which hopefully will stay free a bit longer).

# Setup

- Open your crypto spreadsheet, and create an additional empty tab named `data`;
- Go to the menu `Add-ons` > `Script Apps`;
- Paste the content of the file `scriptCoinGecko.gs` or `scriptCoinPaprika.gs` into the editor, then save the file;
- Go back to your spreadsheet and refresh the page;
- A new menu item called `CoinGecko` or `CoinPaprika` depending on the one you chose should appear. Click on it, then click on `Update`;
- All the data will be imported into your `data` tab. From now on, you can import any data with this simple: `=CRYPTODATA("coin_sympol", "data_needed")`

❗ Do not use the data sheet, it's used by the script ❗

# Additional explanations

### For the coingecko script:
`=CRYPTODATA("btc", "current_price")` => Will get btc price quoted in usd;</br>
Feel free to explore the data tab to find the data you need. Just copy the column title and paste it in the formula above.

### For the coinpaprika script:

#### Get data for a specific coin :
`=CRYPTODATA("ETH"; "btc_price")` => Will get eth price quoted in btc;

#### Get historical data for a specific coin :
`=CRYPTODATAHISTORY("<coin>"; "<date>"; "<type of data>"; "<quote>")` with params being :

- coin : a single coin ticket like "ETH" or "BTC" or "XMR"
- date : something like "2018-02-20"
- type of data : can be "price", "volume_24h" or "market_cap"
- quote : optional, defaults to usd, but can be set to "usd" or "btc"

#### Get global data :
`=CRYPTODATAGLOBAL("bitcoin_dominance_percentage")`

# Update data

- Click on the menu item, then on Update to refresh the data;</br>
- Then, you need to refresh your spreadsheet. One simple way is to add a new line at the top of your sheet, and to delete it right away. This small change will refresh all data automatically.

Any issues, improvements, forks are welcomed. It is certainly not optimized, but it's working :)

If you like it, a donation is always welcome!

BTC : 37wP5rdaFgtHrEQ44M5Tntyeb9nChd8jC4