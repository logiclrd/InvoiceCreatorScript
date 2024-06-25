function ActivateSpreadsheetCreator()
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var invoiceNumber = sheet.getRange("F13").getValue();

  if (invoiceNumber == 1)
    StartNewInvoice();
  else
    AddInvoiceRow();
}

function StartNewInvoice()
{
  var invoicesFolder = DriveApp.getFolderById("188RjmsuYas0YFg0OOUSlXzfQzePQmMzM");

  var lastInvoiceNumber = -1;
  var invoiceCreatorFile = null;

  var allFiles = invoicesFolder.getFiles();

  while (allFiles.hasNext())
  {
    var file = allFiles.next();

    var fileName = file.getName();

    if (fileName.startsWith("Invoice #"))
    {
      var invoiceNumber = parseInt(fileName.substring(9));

      if (invoiceNumber > lastInvoiceNumber)
        lastInvoiceNumber = invoiceNumber;
    }
    else if (fileName == "Invoice Creator")
      invoiceCreatorFile = file;
  }

  var newInvoiceNumber = lastInvoiceNumber + 1;

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  spreadsheet.getRange("F13").setValue(newInvoiceNumber);
  spreadsheet.getRange("F16").setValue(new Date().toISOString().substring(0, 10));

  SpreadsheetApp.flush();

  var newInvoiceFileName = "Invoice #" + newInvoiceNumber;

  var newInvoice = invoiceCreatorFile.makeCopy(newInvoiceFileName, invoicesFolder);

  var newInvoiceURL = "https://docs.google.com/spreadsheets/d/" + newInvoice.getId();

  OpenURL(newInvoiceURL);

  spreadsheet.getRange("F13").setValue(1);
}

function OpenURL(url)
{
  var html =
  '<html>'
  + '<script>'
  +   'window.close = function()'
  +     '{'
  +       'window.setTimeout('
  +         'function()'
  +         '{'
  +           'google.script.host.close();'
  +         '}, 10);'
  +     '};'
  +   'var result = window.open("' + url + '");'
  +   'if (result)'
  +     'close();'
  + '</script>'
  // Offer URL as clickable link in case above code fails.
  + '<body style="word-break: break-word; font-family:sans-serif;">'
  +   'Failed to open automatically. '
  +   '<a href="' + url + '" target="_blank" onclick="window.close()">Click here to proceed</a>.'
  + '</body>'
  + '<script>'
  +   'google.script.host.setHeight(40);'
  +   'google.script.host.setWidth(410);'
  + '</script>' +
  '</html>';

  var htmlOutputObject = HtmlService.createHtmlOutput(html).setWidth(90).setHeight(1);

  SpreadsheetApp.getUi().showModalDialog(htmlOutputObject, "Opening...");
}
function AddInvoiceRow()
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var lastUsedRow = -1;

  var firstRow = 20;

  for (var row = firstRow; row < 2000; row++)
  {
    if (spreadsheet.getRange('F' + row).getValue() == 'Subtotal')
    {
      lastUsedRow = row - 1;
      break;
    }
  }

  if (lastUsedRow < 0)
  {
    Browser.msgBox("Couldn't find last used row");
    return;
  }

  spreadsheet.insertRowAfter(lastUsedRow);

  var newRow = lastUsedRow + 1;

  var range = spreadsheet.getRange('B' + newRow + ':G' + newRow);
  var descriptionCellAddress = 'B' + newRow;
  var rowQuantityCell = spreadsheet.getRange('E' + newRow);
  var rowTotalCell = spreadsheet.getRange('G' + newRow);
  var totalCell = spreadsheet.getRange('G' + (newRow + 1));

  range.setBorder(/* top */ false, /* left */ false, /* bottom */ true, /* right */ false, /* vertical */ false, /* horizontal */ false, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
  rowQuantityCell.setValue(1);
  rowTotalCell.setFormula('IF(ISBLANK(E' + newRow + '),1,E' + newRow + ')*F' + newRow + '');
  totalCell.setFormula('SUM(G' + firstRow + ':G' + newRow + ')');

  spreadsheet.setActiveSelection(descriptionCellAddress);
}

function MakePrintInvoice()
{
  var document = SpreadsheetApp.getActiveSpreadsheet();

  var allSheets = document.getSheets();

  var invoiceSheet;
  var printSheet;

  for (var i = 0; i < allSheets.length; i++)
  {
    let sheet = allSheets[i];

    let name = sheet.getName();

    if (name == 'Invoice')
      invoiceSheet = sheet;
    else if (name == 'Print')
      printSheet = sheet;
  }

  if (printSheet.getMaxRows() >= 15)
    printSheet.deleteRows(15, printSheet.getMaxRows() - 14);
  
  var firstHeaderRow = 12;
  var lastHeaderRow = 13;

  var firstItemRow = 14;
  var lastItemRow = 15;

  for (var row = firstHeaderRow; row < 1000; row++)
  {
    if (invoiceSheet.getRange(row, 2).getValue() == 'Description')
    {
      lastHeaderRow = row - 1;
      firstItemRow = row + 1;
      break;
    }
  }

  for (var row = firstItemRow + 1; row < 1000; row++)
  {
    if (invoiceSheet.getRange(row, 6).getValue() == 'Subtotal')
    {
      lastItemRow = row - 1;
      break;
    }
  }

  var headerFields = {};
  var headerName = null;
  var headerValue = [];

  headerFields["Date"] = invoiceSheet.getRange("B10").getValue();

  for (var column = 1; column <= 7; column++)
  {
    for (var row = firstHeaderRow; row <= lastHeaderRow; row++)
    {
      var cell = invoiceSheet.getRange(row, column);

      var value = cell.getValue();

      if (cell.getFontWeight() == 'bold')
      {
        if ((headerName != null) && (headerValue.length > 0))
          headerFields[headerName] = headerValue;

        headerName = value;
        headerValue = [];
      }
      else
      {
        if (value != '')
          headerValue.push(value);
      }
    }
  }

  if ((headerName != null) && (headerValue.length > 0))
    headerFields[headerName] = headerValue;

  var rowNumber = 14;

  function StartNewRow()
  {
    printSheet.insertRowAfter(rowNumber);

    rowNumber++;

    var dataArea = printSheet.getRange(rowNumber, 1, 1, 6);

    dataArea.clearFormat();
    dataArea.setFontFamily("Roboto Mono");
    dataArea.setFontSize(16);
  }

  StartNewRow();

  if ("Invoice #" in headerFields)
    printSheet.getRange(rowNumber, 1).setValue("Invoice #: " + headerFields["Invoice #"])
  
  if ("Date" in headerFields)
  {
    var dateRange = printSheet.getRange(rowNumber, 5, 1, 2);

    dateRange.merge();
    dateRange.setHorizontalAlignment("right");
    dateRange.setValue(headerFields["Date"].toISOString().substring(0, 10));
  }

  StartNewRow();

  if ("Invoice for" in headerFields)
  {
    var forRows = headerFields["Invoice for"];

    for (var i = 0; i < forRows.length; i++)
    {
      StartNewRow();
      printSheet.getRange(rowNumber, 1).setValue(forRows[i]);
    }
  }

  if ("Due date" in headerFields)
  {
    var dueDate = headerFields["Due date"];

    if (dueDate != '')
    {
      StartNewRow();
      StartNewRow();

      dueDate = new Date(dueDate).toISOString().substring(0, 10);

      printSheet.getRange(rowNumber, 1).setValue("Due: " + dueDate);

      StartNewRow();
    }
  }

  StartNewRow();

  printSheet.getRange(rowNumber, 1).setValue("Description");
  printSheet.getRange(rowNumber, 4).setValue("Qty");
  printSheet.getRange(rowNumber, 5).setValue("Price");
  printSheet.getRange(rowNumber, 6).setValue("Subtotal");

  printSheet.getRange(rowNumber, 1, 1, 6).setFontWeight("bold");

  const RowWidth = 45;
  const QtyColumnOffset = 21;

  for (var itemRow = firstItemRow; itemRow <= lastItemRow; itemRow++)
  {
    var description = invoiceSheet.getRange(itemRow, 2).getValue();
    var quantity = invoiceSheet.getRange(itemRow, 5).getValue();
    var unitPrice = invoiceSheet.getRange(itemRow, 6).getValue();
    var totalPrice = invoiceSheet.getRange(itemRow, 7).getValue();

    while (description.length > 45)
    {
      var splitPoint;

      for (splitPoint = 45; description[splitPoint] != ' '; splitPoint--)
      {
        if (splitPoint == 0)
        {
          splitPoint = 45;
          break;
        }
      }

      StartNewRow();

      printSheet.getRange(rowNumber, 1).setValue(description.substring(0, splitPoint));

      description = description.substring(splitPoint).trim();
    }

    StartNewRow();
    printSheet.getRange(rowNumber, 1).setValue(description);

    if (description.length > QtyColumnOffset)
      StartNewRow();

    var quantityCell = printSheet.getRange(rowNumber, 4);
    var unitPriceCell = printSheet.getRange(rowNumber, 5);
    var totalPriceCell = printSheet.getRange(rowNumber, 6);

    quantityCell.setValue(quantity + " @");
    quantityCell.setHorizontalAlignment("right");

    unitPriceCell.setValue(unitPrice);
    unitPriceCell.setNumberFormat("$#,##0.00");
    unitPriceCell.setHorizontalAlignment("right");

    totalPriceCell.setValue(totalPrice);
    totalPriceCell.setNumberFormat("$#,##0.00");
    totalPriceCell.setHorizontalAlignment("right");
  }

  StartNewRow();

  var summaryRow = lastItemRow + 1;

  while (summaryRow <= invoiceSheet.getLastRow())
  {
    var cell = invoiceSheet.getRange(summaryRow, 6);
    
    if (!cell.isBlank())
    {
      var label = cell.getValue();
      var value = invoiceSheet.getRange(summaryRow, 7).getValue();

      StartNewRow();

      if (label == "Total")
        StartNewRow();

      var labelCell = printSheet.getRange(rowNumber, 5);
      var valueCell = printSheet.getRange(rowNumber, 6);

      labelCell.setValue(label);

      valueCell.setValue(value);
      valueCell.setNumberFormat("$#,##0.00");
      valueCell.setHorizontalAlignment("right");

      if (label == "Total")
        StartNewRow();
    }

    summaryRow++;
  }
}