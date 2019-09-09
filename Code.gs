/**
 * A custom function that gets the adjusted purchase price of a trading sequence
 *
 * @param {String} selected_symbol The selected symbol.
 * @param {Number} max_row_idx The maximum row index to process.
 * @param {String} selected_postype Whether a long or short position (L|S).
 * @return {Number} The adjusted purchase price.
 */
function TRADEADJUSTEDPRICE(selected_symbol, max_row_idx, selected_postype) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var weighted_average_price = 0;
  var total_qty = 0;
  var total_value = 0;
  
  var range_symbol = ss.getRangeByName('HistorySymbolRange');
  var range_qty = ss.getRangeByName('HistoryQuantityRange');
  var range_buy = ss.getRangeByName('HistoryBuypriceRange');
  var range_sell = ss.getRangeByName('HistorySellpriceRange');
  var range_postype = ss.getRangeByName('HistoryPositiontypeRange');
  
  var values_symbol = range_symbol.getValues();
  var values_qty = range_qty.getValues();
  var values_buy = range_buy.getValues();
  var values_sell = range_sell.getValues();
  var values_postype = range_postype.getValues();
  
  var num_rows = values_symbol.length;
  
  for (i = 0; i < num_rows; i++) {
    // LOOP THROUGH ALL TRADES
    
    if (i > max_row_idx) {
      // bail out here because we dont want to calculate price for more trades than this cell represents
      break;
    }
    
    var symbol = values_symbol[i];
    var qty = Number(values_qty[i]);
    var buy_price = Number(values_buy[i]);
    var sell_price = Number(values_sell[i]);
    var postype = values_postype[i];
    
    if (symbol != selected_symbol) {
      // only process the selected symbol
      continue;
    }
        
    if (postype != selected_postype) {
      // only process the selected postype
      continue;
    }
    
    if (!isNumeric(qty) || !isNumeric(buy_price) || !isNumeric(sell_price)) {
      // skip row if trade data is shit
      continue;
    }
    
    // ONWARDS, TO PROCESS PRICE
    
    // LONG POSITIONS
    if (postype == "L") {
      if (sell_price > 0) {
        total_qty -= qty;
        total_value = total_qty *weighted_average_price;
      }
      
      if (total_qty == 0) {
        total_value = 0;
      }
      
      if (buy_price > 0) {
        total_qty += qty;
        total_value += (qty *buy_price);
        
        weighted_average_price = total_value /total_qty;
      }
    }
    
    // SHORT POSITIONS
    if (postype == "S") {
      if (buy_price > 0) {
        total_qty -= qty;
        total_value = total_qty *weighted_average_price;
      }
      
      if (total_qty == 0) {
        total_value = 0;
      }
      
      if (sell_price > 0) {
        total_qty += qty;
        total_value += (qty *sell_price);
        
        weighted_average_price = total_value /total_qty;
      }
    }
  }
  
  return weighted_average_price;
}

function isNumeric(num) {
  return !isNaN(parseFloat(num)) && isFinite(num);
}
