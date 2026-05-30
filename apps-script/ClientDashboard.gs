/**
 * ClientDashboard.gs
 * ------------------------------------------------------------------
 * Live client dashboard for a Google Sheet.
 *
 * What it does:
 *   - You and your team edit the "working" tabs (every tab EXCEPT the
 *     client-facing ones listed in CONFIG.clientTabs).
 *   - This script aggregates the "Status" column across all those
 *     working tabs and renders, on each client-facing tab:
 *         1. A summary table (count of items per status)
 *         2. A PIE chart of the status breakdown
 *         3. A PROGRESS BAR showing overall % complete
 *   - It re-runs automatically whenever anyone edits the sheet (via an
 *     installable onEdit trigger) and can also be run on a timer, so the
 *     client always sees the latest version.
 *
 * Setup (one time):
 *   1. Open the sheet -> Extensions -> Apps Script.
 *   2. Paste this file in (replace any default Code.gs content) and Save.
 *   3. Run the function `setup` once. Approve the permissions prompt.
 *   4. Reload the sheet. Use the new "Client Dashboard" menu to refresh
 *      manually any time.
 * ------------------------------------------------------------------
 */

var CONFIG = {
  // Tabs the client sees. Everything else is treated as editable/working.
  clientTabs: ['Clients', 'POV'],

  // Header names (case-insensitive) that identify the status column on a
  // working tab. The first matching header found is used.
  statusHeaderAliases: ['status', 'stage', 'progress', 'state'],

  // Row number that holds the headers on each working tab.
  headerRow: 1,

  // Known statuses, their pie-slice colors, and whether each counts as
  // "complete" for the progress bar. Any status found in the data that is
  // not listed here is still counted and shown (with a default color).
  statuses: [
    { name: 'Done',        color: '#34a853', complete: true  },
    { name: 'In Progress', color: '#fbbc04', complete: false },
    { name: 'Not Started', color: '#ea4335', complete: false },
    { name: 'Blocked',     color: '#9e9e9e', complete: false },
  ],

  // Default color for any unexpected status value.
  otherColor: '#4285f4',

  // Top-left cell of the dashboard block on each client tab.
  // Pick an area that does not overlap your existing client content.
  // Charts float over the grid; only the summary table + progress bar
  // write into cells, starting here.
  anchorCell: 'J1',
};

/* ============================ MENU / SETUP ============================ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Client Dashboard')
    .addItem('Refresh now', 'refreshClientViews')
    .addItem('Install auto-refresh (on edit)', 'setup')
    .addToUi();
}

/** Run once: installs the onEdit trigger and does a first refresh. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remove any existing triggers for this handler to avoid duplicates.
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onEditRefresh') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Installable onEdit trigger (works for edits by any editor).
  ScriptApp.newTrigger('onEditRefresh')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  refreshClientViews();
  SpreadsheetApp.getActive().toast('Client Dashboard installed and refreshed.');
}

/** Trigger target. Only refreshes when a working (non-client) tab changes. */
function onEditRefresh(e) {
  try {
    if (e && e.range) {
      var name = e.range.getSheet().getName();
      if (CONFIG.clientTabs.indexOf(name) !== -1) return; // ignore edits on client tabs
    }
    refreshClientViews();
  } catch (err) {
    // Swallow errors so editing is never blocked; check Executions log if needed.
    console.error(err);
  }
}

/* ========================== CORE AGGREGATION ========================== */

/** Counts statuses across every working tab. Returns ordered {name,count,color,complete}. */
function aggregateStatuses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var counts = {};   // name -> count
  var order = [];    // preserve discovery / config order

  // Seed with configured statuses so they always appear (even at 0).
  CONFIG.statuses.forEach(function (s) {
    counts[s.name] = 0;
    order.push(s.name);
  });

  ss.getSheets().forEach(function (sheet) {
    var name = sheet.getName();
    if (CONFIG.clientTabs.indexOf(name) !== -1) return; // skip client tabs
    if (name.charAt(0) === '_') return;                 // skip helper/hidden tabs

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow <= CONFIG.headerRow || lastCol < 1) return;

    var headers = sheet.getRange(CONFIG.headerRow, 1, 1, lastCol).getValues()[0];
    var statusCol = -1;
    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c]).trim().toLowerCase();
      if (CONFIG.statusHeaderAliases.indexOf(h) !== -1) { statusCol = c + 1; break; }
    }
    if (statusCol === -1) return; // no status column on this tab

    var n = lastRow - CONFIG.headerRow;
    var values = sheet.getRange(CONFIG.headerRow + 1, statusCol, n, 1).getValues();
    values.forEach(function (row) {
      var v = String(row[0]).trim();
      if (!v) return;
      var key = canonicalStatus(v);
      if (!(key in counts)) { counts[key] = 0; order.push(key); }
      counts[key]++;
    });
  });

  return order.map(function (name) {
    var cfg = findStatusConfig(name);
    return {
      name: name,
      count: counts[name],
      color: cfg ? cfg.color : CONFIG.otherColor,
      complete: cfg ? !!cfg.complete : false,
    };
  });
}

/** Maps a raw status value to a configured status name (case-insensitive). */
function canonicalStatus(value) {
  var lower = value.toLowerCase();
  for (var i = 0; i < CONFIG.statuses.length; i++) {
    if (CONFIG.statuses[i].name.toLowerCase() === lower) return CONFIG.statuses[i].name;
  }
  return value; // keep original label for unknown statuses
}

function findStatusConfig(name) {
  for (var i = 0; i < CONFIG.statuses.length; i++) {
    if (CONFIG.statuses[i].name === name) return CONFIG.statuses[i];
  }
  return null;
}

/* =========================== RENDERING =========================== */

/** Public entry point: refresh both client tabs. */
function refreshClientViews() {
  var data = aggregateStatuses();
  CONFIG.clientTabs.forEach(function (tabName) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
    if (sheet) renderDashboard(sheet, data);
  });
}

function renderDashboard(sheet, data) {
  var anchor = sheet.getRange(CONFIG.anchorCell);
  var r0 = anchor.getRow();
  var c0 = anchor.getColumn();

  var total = data.reduce(function (sum, d) { return sum + d.count; }, 0);
  var completed = data.reduce(function (sum, d) { return sum + (d.complete ? d.count : 0); }, 0);
  var pct = total > 0 ? completed / total : 0;

  // --- Clear the previous dashboard block (table area) ---
  sheet.getRange(r0, c0, data.length + 12, 3).clearContent();

  // --- Title / timestamp ---
  sheet.getRange(r0, c0).setValue('Project Dashboard (auto-updated)')
    .setFontWeight('bold').setFontSize(12);
  sheet.getRange(r0 + 1, c0).setValue('Last updated: ' + new Date());

  // --- Status table (source range for the pie chart) ---
  var tableTop = r0 + 3;
  sheet.getRange(tableTop, c0, 1, 2).setValues([['Status', 'Count']])
    .setFontWeight('bold');
  for (var i = 0; i < data.length; i++) {
    var rowIdx = tableTop + 1 + i;
    sheet.getRange(rowIdx, c0, 1, 2).setValues([[data[i].name, data[i].count]]);
    // Color-swatch the status cell so the table matches the pie.
    sheet.getRange(rowIdx, c0).setBackground(data[i].color)
      .setFontColor(pickTextColor(data[i].color));
  }
  var totalRow = tableTop + 1 + data.length;
  sheet.getRange(totalRow, c0, 1, 2).setValues([['Total', total]])
    .setFontWeight('bold').setBackground(null).setFontColor(null);

  // --- Progress bar (in-cell SPARKLINE) + percentage ---
  var pbLabelRow = totalRow + 2;
  sheet.getRange(pbLabelRow, c0).setValue('Overall % Complete').setFontWeight('bold');
  var pbRow = pbLabelRow + 1;
  // SPARKLINE bar from 0..1; color shifts from red -> amber -> green by progress.
  var barColor = pct >= 0.8 ? '#34a853' : (pct >= 0.4 ? '#fbbc04' : '#ea4335');
  sheet.getRange(pbRow, c0).setFormula(
    '=SPARKLINE(' + pct + ',{"charttype","bar";"max",1;"color1","' + barColor + '"})'
  );
  sheet.getRange(pbRow, c0 + 1).setValue(pct).setNumberFormat('0%').setFontWeight('bold');

  // --- Pie chart (rebuilt each refresh) ---
  rebuildPieChart(sheet, tableTop, c0, data.length, pbRow);
}

/** Removes our previous pie chart and inserts a fresh one. */
function rebuildPieChart(sheet, tableTop, c0, nStatuses, pbRow) {
  var TITLE = 'Status Breakdown';

  // Remove any existing chart we previously made (match by title).
  sheet.getCharts().forEach(function (ch) {
    if (ch.getOptions().get('title') === TITLE) sheet.removeChart(ch);
  });

  // Source: the Status/Count table (header + rows, 2 columns).
  var dataRange = sheet.getRange(tableTop, c0, nStatuses + 1, 2);

  var colors = [];
  for (var i = 0; i < nStatuses; i++) {
    colors.push(sheet.getRange(tableTop + 1 + i, c0).getBackground());
  }

  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(dataRange)
    .setNumHeaders(1)
    .setOption('title', TITLE)
    .setOption('pieHole', 0)               // set to 0.4 for a donut
    .setOption('colors', colors)
    .setOption('width', 360)
    .setOption('height', 260)
    // Anchor the chart a few rows below the progress bar, same column.
    .setPosition(pbRow + 2, c0, 0, 0)
    .build();

  sheet.insertChart(chart);
}

/* ============================ HELPERS ============================ */

/** Returns black or white text for best contrast on a hex background. */
function pickTextColor(hex) {
  var h = hex.replace('#', '');
  if (h.length !== 6) return '#000000';
  var r = parseInt(h.substr(0, 2), 16);
  var g = parseInt(h.substr(2, 2), 16);
  var b = parseInt(h.substr(4, 2), 16);
  var luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance > 0.6 ? '#000000' : '#ffffff';
}
