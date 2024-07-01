function onEdit(e) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  
  sheets.forEach(function(sheet) {
    processSheetForEmojis(sheet);
  });
}

function processSheetForEmojis(sheet) {
  if (!sheet) {
    return;
  }
  
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  values.forEach(function(row, rowIndex) {
    row.forEach(function(cellValue, columnIndex) {
      if (typeof cellValue === 'string' && containsEmoji(cellValue)) {
        var cleanedValue = removeEmojis(cellValue);
        if (cleanedValue !== undefined) {
          sheet.getRange(rowIndex + 1, columnIndex + 1).setValue(cleanedValue);
        }
      }
    });
  });
}

function containsEmoji(text) {
  var emojiRegex = /[\u{1F600}-\u{1F64F}\u{1F300}-\u{1F5FF}\u{1F680}-\u{1F6FF}\u{1F1E0}-\u{1F1FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}]/u;
  return emojiRegex.test(text);
}

function removeEmojis(text) {
  var emojiRegex = /[\u{1F600}-\u{1F64F}\u{1F300}-\u{1F5FF}\u{1F680}-\u{1F6FF}\u{1F1E0}-\u{1F1FF}\u{2600}-\u{26FF}\u{2700}-\u{27BF}]/ug;
  if (typeof text === 'string') {
    return text.replace(emojiRegex, '');
  } else {
    return text;
  }
}
