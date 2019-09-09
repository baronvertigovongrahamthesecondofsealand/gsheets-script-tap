# gsheets-script-tap
Google Sheets script for calculating stock market trade adjusted price

Usage
-----

    =TRADEADJUSTEDPRICE(selected_symbol, max_row_idx, selected_postype)

* selected_symbol: The trade symbol you wish to calculate on, eg. "AMZN".
* max_row_idx: If you wish to calculate the adjusted price at every trade of your history, use this to tell the script to calculate only up to the current trade.
* selected_postype: Whether long ("L"), or short ("S").

For this script to work optimally, you should have a trade history sheet with symbol, position type, buy price, sell price, and quantity fields.

For example, for longs, a buy trade would have buy price be the trade price, and sell price zero. If you make a long sell trade, buy price will be zero, and sell price the trade price. Shorts are opposite.

You will need to set these named ranges in your history sheet:
* HistorySymbolRange
* HistoryQuantityRange
* HistoryBuypriceRange
* HistorySellpriceRange
* HistoryPositiontypeRange
