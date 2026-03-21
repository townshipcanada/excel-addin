/**
 * Township Canada Excel Add-In - Task Pane Logic
 *
 * Handles batch conversion, settings/API key management,
 * and usage tracking for the task pane sidebar.
 *
 * Requires a trial or paid API key.
 */

import {
  MAX_BATCH_SIZE,
  TRIAL_URL,
  getApiKey,
  setApiKey,
  removeApiKey,
  hasApiKey,
  apiConvertBatch,
  apiGetUsage
} from "../shared/config";

// ── Tab switching ──

function switchTab(tabName) {
  document.querySelectorAll(".tab").forEach(function (tab) {
    tab.classList.toggle("active", tab.dataset.tab === tabName);
  });

  document.querySelectorAll(".tab-content").forEach(function (content) {
    content.classList.toggle("active", content.id === "tab-" + tabName);
  });

  if (tabName === "settings") {
    checkCurrentKey();
  }
}

// ── Usage display ──

async function loadUsage() {
  try {
    var usage = await apiGetUsage();
    updateUsageDisplay(usage);
  } catch (e) {
    document.getElementById("usageCount").innerHTML = "<span>Unable to load</span>";
  }
}

function updateUsageDisplay(usage) {
  var usageBar = document.getElementById("usageBar");
  var countEl = document.getElementById("usageCount");
  var progressEl = document.getElementById("usageProgress");
  var setupBanner = document.getElementById("setupBanner");
  var trialInfo = document.getElementById("trialInfo");
  var usageLabel = document.getElementById("usageLabel");

  if (!usage.apiKeyValid) {
    // No API key connected
    setupBanner.style.display = "block";
    usageBar.style.display = "none";
    trialInfo.style.display = "none";
  } else if (usage.plan === "trial") {
    // Trial key
    setupBanner.style.display = "none";
    usageBar.style.display = "block";
    usageLabel.textContent = "Trial Usage";
    countEl.innerHTML = usage.used + "<span> / " + usage.limit + " calls</span>";
    var pct = Math.min(100, (usage.used / usage.limit) * 100);
    progressEl.style.width = pct + "%";

    if (pct >= 100) {
      progressEl.className = "progress-fill danger";
    } else if (pct >= 70) {
      progressEl.className = "progress-fill warning";
    } else {
      progressEl.className = "progress-fill";
    }

    trialInfo.style.display = "block";
    trialInfo.textContent = usage.daysLeft + " days remaining on trial. Upgrade at townshipcanada.com/pricing for unlimited access.";
  } else {
    // Paid key
    setupBanner.style.display = "none";
    usageBar.style.display = "block";
    usageLabel.textContent = "API Usage";
    countEl.innerHTML = "Unlimited <span>(paid API key)</span>";
    progressEl.className = "progress-fill unlimited";
    progressEl.style.width = "100%";
    trialInfo.style.display = "none";
  }
}

// ── Batch conversion ──

/**
 * Convert column letter(s) to 0-based column index.
 * e.g., "A" -> 0, "B" -> 1, "AA" -> 26
 */
function columnLetterToIndex(letter) {
  var col = 0;
  for (var i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col - 1;
}

function handleConversionError(e) {
  if (e.message === "NO_API_KEY") {
    showStatus("API key required. Connect a trial or paid key in Settings.", "error");
    document.getElementById("setupBanner").style.display = "block";
  } else if (e.message === "TRIAL_EXPIRED") {
    showStatus("Trial key has expired. Upgrade at townshipcanada.com/pricing for continued access.", "error");
  } else if (e.message === "TRIAL_LIMIT_REACHED") {
    showStatus("Trial usage limit reached. Upgrade at townshipcanada.com/pricing for continued access.", "error");
  } else if (e.message === "INVALID_API_KEY") {
    showStatus("Invalid API key. Check your key in Settings.", "error");
  } else {
    showStatus(e.message || "Conversion failed", "error");
  }
}

async function convertSelected() {
  var btn = document.getElementById("convertSelectedBtn");
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>Converting...';
  hideStatus();

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
        showStatus("No legal land descriptions found in the selected cells.", "error");
        return;
      }

      // Process in batches
      var allResults = [];
      for (var batchIdx = 0; batchIdx < queries.length; batchIdx += MAX_BATCH_SIZE) {
        var chunk = queries.slice(batchIdx, batchIdx + MAX_BATCH_SIZE);
        var batchResponse = await apiConvertBatch(chunk);
        allResults = allResults.concat(batchResponse.data);
      }

      // Write headers in the row above if possible
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

      var success = allResults.filter(function (r) {
        return r.latitude !== null;
      }).length;
      var failure = allResults.length - success;

      showStatus("Conversion complete!", "success");
      showResults(allResults.length, success, failure);
    });
  } catch (e) {
    handleConversionError(e);
  } finally {
    btn.disabled = false;
    btn.textContent = "Convert Selected Cells";
    loadUsage();
  }
}

async function convertColumn() {
  var col = document.getElementById("columnInput").value.trim().toUpperCase();
  if (!col || !/^[A-Z]{1,2}$/.test(col)) {
    showStatus("Please enter a valid column letter (e.g., A, B, AA).", "error");
    return;
  }

  var btn = document.getElementById("convertColumnBtn");
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>Converting...';
  hideStatus();

  try {
    await Excel.run(async function (context) {
      var sheet = context.workbook.worksheets.getActiveWorksheet();
      var usedRange = sheet.getUsedRange();
      usedRange.load("rowCount");
      await context.sync();

      var lastRow = usedRange.rowCount;
      if (lastRow < 2) {
        showStatus("No data found in the sheet.", "error");
        return;
      }

      var colIndex = columnLetterToIndex(col);
      var dataRange = sheet.getRangeByIndexes(1, colIndex, lastRow - 1, 1);
      dataRange.load("values");
      await context.sync();

      var values = dataRange.values;
      var queries = [];
      var rowIndices = [];

      for (var colIdx = 0; colIdx < values.length; colIdx++) {
        var val = String(values[colIdx][0]).trim();
        if (val && val !== "undefined" && val !== "null") {
          queries.push(val);
          rowIndices.push(colIdx);
        }
      }

      if (queries.length === 0) {
        showStatus("No legal land descriptions found in column " + col + ".", "error");
        return;
      }

      // Process in batches
      var allResults = [];
      for (var colBatchIdx = 0; colBatchIdx < queries.length; colBatchIdx += MAX_BATCH_SIZE) {
        var chunk = queries.slice(colBatchIdx, colBatchIdx + MAX_BATCH_SIZE);
        var colBatchResponse = await apiConvertBatch(chunk);
        allResults = allResults.concat(colBatchResponse.data);
      }

      // Write headers
      var outputCol = colIndex + 1;
      var headerRange = sheet.getRangeByIndexes(0, outputCol, 1, 3);
      headerRange.values = [["Latitude", "Longitude", "Province"]];
      headerRange.format.font.bold = true;

      // Write results
      var resultIndex = 0;
      for (var colRowIdx = 0; colRowIdx < rowIndices.length; colRowIdx++) {
        if (resultIndex < allResults.length) {
          var r = allResults[resultIndex];
          var outputRange = sheet.getRangeByIndexes(rowIndices[colRowIdx] + 1, outputCol, 1, 3);
          outputRange.values = [[r.latitude || "N/A", r.longitude || "N/A", r.province || "N/A"]];
          resultIndex++;
        }
      }

      await context.sync();

      var success = allResults.filter(function (r) {
        return r.latitude !== null;
      }).length;
      var failure = allResults.length - success;

      showStatus(
        "Done! Converted " + success + "/" + allResults.length + " descriptions.",
        "success"
      );
      showResults(allResults.length, success, failure);
    });
  } catch (e) {
    handleConversionError(e);
  } finally {
    btn.disabled = false;
    btn.textContent = "Convert Entire Column";
    loadUsage();
  }
}

// ── Status display helpers ──

function showStatus(message, type) {
  var el = document.getElementById("statusMessage");
  el.textContent = message;
  el.className = "status-message " + type;
}

function hideStatus() {
  document.getElementById("statusMessage").className = "status-message";
  document.getElementById("resultsSummary").style.display = "none";
}

function showResults(total, success, failure) {
  document.getElementById("resultsSummary").style.display = "block";
  document.getElementById("statTotal").textContent = total;
  document.getElementById("statSuccess").textContent = success;
  document.getElementById("statFailed").textContent = failure;
}

// ── Settings: API key management ──

async function checkCurrentKey() {
  var key = getApiKey();
  if (key) {
    document.getElementById("apiKeyInput").value = key;
    var usage = await apiGetUsage();
    if (usage.apiKeyValid) {
      if (usage.plan === "trial") {
        showTrial(usage);
      } else {
        showConnected();
      }
    } else {
      showDisconnected();
      showSettingsMessage("Stored API key is invalid. Please update it.", "error");
    }
  } else {
    showDisconnected();
  }
}

async function saveKey() {
  var key = document.getElementById("apiKeyInput").value.trim();
  if (!key) {
    showSettingsMessage("Please enter an API key.", "error");
    return;
  }

  if (!key.startsWith("tc_")) {
    showSettingsMessage("Invalid API key format. Keys start with 'tc_'.", "error");
    return;
  }

  setApiKey(key);

  var usage = await apiGetUsage();
  if (usage.apiKeyValid) {
    if (usage.plan === "trial") {
      showTrial(usage);
      showSettingsMessage("Trial key connected! " + usage.remaining + " calls remaining.", "success");
    } else {
      showConnected();
      showSettingsMessage("API key connected successfully!", "success");
    }
    loadUsage();
  } else {
    showDisconnected();
    showSettingsMessage("API key is invalid. Please check and try again.", "error");
    removeApiKey();
  }
}

function removeKeyHandler() {
  removeApiKey();
  document.getElementById("apiKeyInput").value = "";
  showDisconnected();
  showSettingsMessage("API key removed.", "success");
  loadUsage();
}

function showConnected() {
  var badge = document.getElementById("statusBadge");
  badge.className = "status-badge connected";
  document.getElementById("statusDot").className = "status-dot green";
  document.getElementById("statusText").textContent = "Connected (unlimited)";
  document.getElementById("removeBtn").style.display = "inline-block";
}

function showTrial(usage) {
  var badge = document.getElementById("statusBadge");
  badge.className = "status-badge trial";
  document.getElementById("statusDot").className = "status-dot blue";
  document.getElementById("statusText").textContent = "Trial (" + usage.remaining + " calls, " + usage.daysLeft + " days left)";
  document.getElementById("removeBtn").style.display = "inline-block";
}

function showDisconnected() {
  var badge = document.getElementById("statusBadge");
  badge.className = "status-badge disconnected";
  document.getElementById("statusDot").className = "status-dot gray";
  document.getElementById("statusText").textContent = "Not connected";
  document.getElementById("removeBtn").style.display = "none";
}

function showSettingsMessage(text, type) {
  var el = document.getElementById("settingsMessage");
  el.textContent = text;
  el.className = "message " + type;
  setTimeout(function () {
    el.className = "message";
  }, 5000);
}

// ── Expose functions to HTML onclick handlers ──

window.switchTab = switchTab;
window.convertSelected = convertSelected;
window.convertColumn = convertColumn;
window.saveKey = saveKey;
window.removeKey = removeKeyHandler;

// ── Initialize ──

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    loadUsage();
  }
});
