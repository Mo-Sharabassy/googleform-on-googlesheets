var mainWsName = "Form";
var optionsWsName = "Data";
var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mainWsName);
var wsOptions = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(optionsWsName);
var options = wsOptions.getRange(2, 1, wsOptions.getLastRow() - 1, 3).getValues();

function onEdit(e) {
  var activeCell = e.range;
  var val = activeCell.getValue();
  var r = activeCell.getRow();
  var c = activeCell.getColumn();
  var wsName = activeCell.getSheet().getName();

  if (wsName === mainWsName && c === 5) {
    if (r === 14) {
      applyFirstLevelValidation(val);
    } else if (r === 17) {
      applySecondLevelValidation(val);
    }
  }
}

function clearValidationAndContent(range) {
  range.clearContent();
  range.clearDataValidations();
}

function applyFirstLevelValidation(val) {
  clearValidationAndContent(ws.getRange(17, 5));
  clearValidationAndContent(ws.getRange(20, 5));

  if (val !== "") {
    var filteredOptions = options.filter(function (o) {
      return o[0] === val;
    });
    var listToApply = filteredOptions.map(function (o) {
      return o[1];
    });
    applyValidationToCell(listToApply, ws.getRange(17, 5));
  }
}

function applySecondLevelValidation(val) {
  clearValidationAndContent(ws.getRange(20, 5));
  var firstLevelColValue = ws.getRange(14, 5).getValue();
  var filteredOptions = options.filter(function (o) {
    return o[0] === firstLevelColValue && o[1] === val;
  });
  var listToApply = filteredOptions.map(function (o) {
    return o[2];
  });
  applyValidationToCell(listToApply, ws.getRange(20, 5));
}

function applyValidationToCell(list, cell) {
  var rule = SpreadsheetApp
    .newDataValidation()
    .requireValueInList(list)
    .setAllowInvalid(false)
    .build();

  cell.setDataValidation(rule);
}
