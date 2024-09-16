function onOpen() {
  // Adds a custom menu to the Google Sheet upon opening
  SpreadsheetApp.getUi()
    .createMenu('Price Fetcher')
    .addItem('Fetch Price for Selection', 'fetchPriceForSelectedRange')
    .addToUi();
}

function fetchPriceForSelectedRange() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRange();
  const selectedValues = selection.getValues();

  // Determine the header row (assuming it's the first row)
  const headerRow = 1;
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Find the "Price" column index
  const priceColIndex = headers.indexOf("Price") + 1;
  if (priceColIndex === 0) {
    SpreadsheetApp.getUi().alert('Price column not found. Please ensure there is a header named "Price".');
    return;
  }

  // Prepare an array to hold the updated prices
  const updatedPrices = [];

  // Iterate through each row in the selection
  for (let i = 0; i < selectedValues.length; i++) {
    const symbol = selectedValues[i][0].toString().trim();
    if (symbol === "") {
      updatedPrices.push([""]); // Leave empty if symbol is empty
      continue;
    }

    const url = `https://dps.psx.com.pk/company/${encodeURIComponent(symbol)}`;
    try {
      const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      if (response.getResponseCode() !== 200) {
        updatedPrices.push([`Error: ${response.getResponseCode()}`]);
        continue;
      }

      const html = response.getContentText();
      const price = extractPrice(html);
      
      if (price !== null) {
        updatedPrices.push([price]);
      } else {
        updatedPrices.push(["Price not found"]);
      }

    } catch (error) {
      updatedPrices.push(["Error fetching data"]);
      Logger.log(`Error fetching data for symbol ${symbol}: ${error}`);
    }
  }

  // Determine the starting row and column for the "Price" updates
  const startRow = selection.getRow();
  const priceRange = sheet.getRange(startRow, priceColIndex, updatedPrices.length, 1);
  
  // Set the updated prices in bulk
  priceRange.setValues(updatedPrices);

  SpreadsheetApp.getUi().alert('Price update completed.');
}

function extractPrice(html) {
  // This function extracts the price from the HTML using a regular expression
  // It looks for <div class="quote__close">Rs.182.48</div>
  const regex = /<div\s+class=["']quote__close["']>\s*Rs\.([\d,]+\.\d{2})\s*<\/div>/i;
  const match = html.match(regex);
  if (match && match[1]) {
    // Remove any commas in the number and convert to a number type
    return parseFloat(match[1].replace(/,/g, ''));
  }
  return null;
}
