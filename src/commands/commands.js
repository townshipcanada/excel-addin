/**
 * Township Canada Excel Add-In - Ribbon Commands
 *
 * Handles ribbon button actions for batch conversion
 * of selected cells directly from the Excel ribbon.
 */

import { MAX_BATCH_SIZE, apiConvertBatch } from "../shared/config";

/**
 * Ribbon button handler: Convert selected cells.
 * Reads LLDs from selected range, batch-converts via API,
 * and writes Latitude/Longitude/Province to adjacent columns.
 */
async function convertSelectedCells(event) {
  try {
    await Excel.run(async function (context) {
      var range = context.workbook.getSelectedRange();
      range.load("values,rowIndex,columnIndex,rowCount,columnCount");
      await context.sync();

      var values = range.values;
      var queries = [];
      var rowIndices = [];

      for (var idx = 0; idx < values.length; idx++) {
        var val = String(values[idx][0]).trim();
        if (val && val !== "undefined" && val !== "null") {
          queries.push(val);
          rowIndices.push(idx);
        }
      }

      if (queries.length === 0) {
        event.completed();
        return;
      }

      // Process in batches
      var allResults = [];
      for (var batchIdx = 0; batchIdx < queries.length; batchIdx += MAX_BATCH_SIZE) {
        var chunk = queries.slice(batchIdx, batchIdx + MAX_BATCH_SIZE);
        var batchResponse = await apiConvertBatch(chunk);
        allResults = allResults.concat(batchResponse.data);
      }

      // Write headers
      var outputCol = range.columnIndex + range.columnCount;
      var startRow = range.rowIndex;
      var sheet = context.workbook.worksheets.getActiveWorksheet();

      if (startRow > 0) {
        var headerRange = sheet.getRangeByIndexes(startRow - 1, outputCol, 1, 3);
        headerRange.load("values");
        await context.sync();

        var existingHeaders = headerRange.values[0];
        if (!existingHeaders[0] && !existingHeaders[1] && !existingHeaders[2]) {
          headerRange.values = [["Latitude", "Longitude", "Province"]];
          headerRange.format.font.bold = true;
        }
      }

      // Write results
      var resultIndex = 0;
      for (var rowIdx = 0; rowIdx < rowIndices.length; rowIdx++) {
        if (resultIndex < allResults.length) {
          var r = allResults[resultIndex];
          var outputRange = sheet.getRangeByIndexes(startRow + rowIndices[rowIdx], outputCol, 1, 3);
          outputRange.values = [[r.latitude || "N/A", r.longitude || "N/A", r.province || "N/A"]];
          resultIndex++;
        }
      }

      await context.sync();
    });
  } catch (e) {
    // Errors from ribbon commands are not displayed to users easily,
    // so we log them. The user can use the task pane for detailed feedback.
    console.error("Township Canada conversion error:", e.message);
  }

  event.completed();
}

// Register the command function inside Office.onReady
Office.onReady(function () {
  Office.actions.associate("convertSelectedCells", convertSelectedCells);
});
