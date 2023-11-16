function GET_HYPERLINK_TEXT(cellAddress) {
  //How to use: 
  //Add the formula =GET_HYPERLINK_TEXT("A1") - The cell address should be in quotes
  //You can use it this way =GET_HYPERLINK_TEXT(CONCAT("A",ROW())) to copy the formula across rows. Replace A with the correct column name.
  //Contact: www.zyxware.com
  // Check if the cell address is provided
  if (!cellAddress) return "";

  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(cellAddress);

  // Directly return if the range is not a single cell
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
    return "";
  }

  var richText = range.getRichTextValue();
  var runs = richText.getRuns();

  for (var i = 0, len = runs.length; i < len; i++) {
    var url = runs[i].getLinkUrl();
    if (url) {
      return url;
    }
  }

  return "";
}
