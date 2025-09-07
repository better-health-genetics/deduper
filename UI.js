
/** ===========================================================================
 * 
 *                        UI & MENU FUNCTIONS
 * 
 * ===========================================================================/

/**
 * Serves the main HTML file to the sidebar.
 */
function showDuplicateCheckerSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('BHG DeDuper')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Opens the web dashboard in a new tab.
 */
function ui_openDashboard_() {
  var url = ScriptApp.getService().getUrl();
  var html = HtmlService.createHtmlOutput('<script>window.open("' + url + '", "_blank");</script>')
    .setWidth(100)
    .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Dashboard...');
}

/**
 * Provides a UI prompt for the user to set their Gemini API key.
 * The key is stored securely in Script Properties.
 */
function ui_setApiKey_() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Set Gemini API Key',
    'Please enter your Google AI Gemini API key. This is stored securely in your script properties and is required for the AI Summary feature.',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    var apiKey = result.getResponseText().trim();
    if (apiKey) {
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', apiKey);
      SpreadsheetApp.getActive().toast('Gemini API Key saved successfully.', 'Success', 5);
    } else {
      ui.alert('API Key cannot be empty.');
    }
  }
}


/**
 * Returns the HTML content for the dashboard web app.
 * This version includes summary cards, a chart, and a detailed table.
 * @return {string} The complete HTML for the dashboard.
 */
function getDashboardHtml_() {
  return `
  <!DOCTYPE html>
  <html>
    <head>
      <base target="_top">
      <title>BHG Duplication Stats Dashboard</title>
      <script src="https://cdn.tailwindcss.com"></script>
      <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
      <style>
        body { background-color: #111827; color: #f3f4f6; font-family: system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', 'Helvetica Neue', sans-serif; }
        .stat-card { background-color: #1f2937; border-color: #374151; }
        .table-header { position: sticky; top: 0; background-color: #374151; }
      </style>
    </head>
    <body>
      <div class="container mx-auto p-4 md:p-8">
        <h1 class="text-3xl md:text-4xl font-bold text-cyan-400 mb-2">BHG DeDuper Dashboard</h1>
        <p id="subtitle" class="text-md text-gray-400 mb-8">Loading data for the current week...</p>
        
        <div id="loading" class="text-center py-16">
          <p class="text-2xl animate-pulse">Loading Dashboard Data...</p>
        </div>
        
        <div id="dashboard-content" style="display: none;">
          <!-- Overall Stat Cards -->
          <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-8">
            <div class="stat-card p-4 rounded-lg border">
              <p class="text-sm text-gray-400">Overall Duplicate Rate</p>
              <p id="overall-rate" class="text-3xl font-bold text-cyan-400">-</p>
            </div>
            <div class="stat-card p-4 rounded-lg border">
              <p class="text-sm text-gray-400">Total Duplicates Found</p>
              <p id="total-duplicates" class="text-3xl font-bold text-white">-</p>
            </div>
            <div class="stat-card p-4 rounded-lg border">
              <p class="text-sm text-gray-400">Total Records Checked</p>
              <p id="total-records" class="text-3xl font-bold text-white">-</p>
            </div>
          </div>

          <!-- Chart and Table -->
          <div class="grid grid-cols-1 lg:grid-cols-5 gap-8">
            <div class="lg:col-span-2 bg-gray-800 p-4 rounded-lg shadow-xl border border-gray-700">
              <h2 class="font-semibold mb-4 text-gray-200">Duplicates by Source (%)</h2>
              <div class="relative h-96">
                <canvas id="statsChart"></canvas>
              </div>
            </div>
            <div class="lg:col-span-3 bg-gray-800 p-4 rounded-lg shadow-xl border border-gray-700">
              <h2 class="font-semibold mb-4 text-gray-200">Detailed Source Stats</h2>
              <div class="overflow-y-auto max-h-96">
                <table class="w-full text-sm text-left text-gray-300">
                  <thead class="text-xs text-gray-400 uppercase">
                    <tr>
                      <th scope="col" class="table-header px-4 py-3">Source Name</th>
                      <th scope="col" class="table-header px-4 py-3 text-right">Duplicate Rate</th>
                      <th scope="col" class="table-header px-4 py-3 text-right">Duplicates</th>
                      <th scope="col" class="table-header px-4 py-3 text-right">Total Records</th>
                    </tr>
                  </thead>
                  <tbody id="stats-table-body">
                    <!-- Rows will be injected here by script -->
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
        
        <div id="error" class="text-center py-16 text-red-400 text-2xl" style="display: none;"></div>
      </div>

      <script>
        document.addEventListener('DOMContentLoaded', function() {
          const today = new Date();
          const day = today.getDay();
          const diff = today.getDate() - day + (day === 0 ? -6 : 1); // Adjust for Sunday
          const monday = new Date(new Date().setDate(diff));
          const startDate = monday.toISOString().split('T')[0];
          const endDate = new Date().toISOString().split('T')[0];

          google.script.run
            .withSuccessHandler(onDataLoaded)
            .withFailureHandler(onDataError)
            .getAdminDashboardData(startDate, endDate);
        });

        function onDataError(error) {
          document.getElementById('loading').style.display = 'none';
          var errorDiv = document.getElementById('error');
          errorDiv.style.display = 'block';
          errorDiv.textContent = 'Error loading data: ' + error.message;
        }

        function onDataLoaded(data) {
          document.getElementById('loading').style.display = 'none';
          
          if (!data || data.length === 0) {
              var errorDiv = document.getElementById('error');
              errorDiv.style.display = 'block';
              errorDiv.textContent = 'No data available to display for this period.';
              return;
          }

          document.getElementById('dashboard-content').style.display = 'block';

          // Sort data by percentage, descending for chart and table
          data.sort((a, b) => b.percentage - a.percentage);

          // --- 1. Calculate and Populate Overall Stats Cards ---
          const totalRecords = data.reduce((sum, item) => sum + item.totalChecked, 0);
          const totalDuplicates = data.reduce((sum, item) => sum + item.totalDuplicates, 0);
          const overallRate = totalRecords > 0 ? Math.round((totalDuplicates / totalRecords) * 100) : 0;
          
          document.getElementById('overall-rate').textContent = overallRate + '%';
          document.getElementById('total-duplicates').textContent = totalDuplicates.toLocaleString();
          document.getElementById('total-records').textContent = totalRecords.toLocaleString();
          document.getElementById('subtitle').textContent = 'Displaying stats for the current week (' + new Date().toLocaleDateString() + ').';

          // --- 2. Populate Detailed Table ---
          const tableBody = document.getElementById('stats-table-body');
          let tableHtml = '';
          data.forEach(item => {
            const healthColor = item.percentage <= 15 ? 'text-green-400' : (item.percentage < 30 ? 'text-yellow-400' : 'text-red-400');
            tableHtml += \`
              <tr class="border-b border-gray-700 hover:bg-gray-700/50">
                <td class="px-4 py-2 font-medium whitespace-nowrap">\${item.sourceName}</td>
                <td class="px-4 py-2 text-right font-bold \${healthColor}">\${item.percentage}%</td>
                <td class="px-4 py-2 text-right">\${item.totalDuplicates.toLocaleString()}</td>
                <td class="px-4 py-2 text-right">\${item.totalChecked.toLocaleString()}</td>
              </tr>
            \`;
          });
          tableBody.innerHTML = tableHtml;

          // --- 3. Render Chart ---
          const chartLabels = data.map(item => item.sourceName).slice(0, 15); // Show top 15 in chart
          const chartPercentages = data.map(item => item.percentage).slice(0, 15);
          
          const getHealthColor = (percentage, opacity = '0.7') => {
              if (percentage <= 15) return \`rgba(74, 222, 128, \${opacity})\`; // green-400
              if (percentage < 30) return \`rgba(250, 204, 21, \${opacity})\`; // yellow-400
              return \`rgba(248, 113, 113, \${opacity})\`; // red-400
          };

          const backgroundColors = chartPercentages.map(p => getHealthColor(p));
          
          const ctx = document.getElementById('statsChart').getContext('2d');
          new Chart(ctx, {
            type: 'bar',
            data: {
              labels: chartLabels,
              datasets: [{
                label: 'Duplicate %',
                data: chartPercentages,
                backgroundColor: backgroundColors,
                borderColor: backgroundColors.map(c => c.replace('0.7', '1')),
                borderWidth: 1
              }]
            },
            options: {
              indexAxis: 'y', // Horizontal bar chart
              responsive: true,
              maintainAspectRatio: false,
              scales: {
                x: { beginAtZero: true, max: 100, ticks: { color: '#9ca3af' }, grid: { color: 'rgba(255, 255, 255, 0.1)' } },
                y: { ticks: { color: '#d1d5db' }, grid: { display: false } }
              },
              plugins: {
                legend: { display: false },
                tooltip: {
                  callbacks: {
                      label: function(context) {
                          let label = context.dataset.label || '';
                          if (label) { label += ': '; }
                          if (context.parsed.x !== null) { label += context.parsed.x + '%'; }
                          // Find original full data item for tooltip
                          const sourceData = data.find(d => d.sourceName === context.label);
                          if (sourceData) {
                             label += ' (' + sourceData.totalDuplicates + ' / ' + sourceData.totalChecked + ' records)';
                          }
                          return label;
                      }
                  }
                }
              }
            }
          });
        }
      </script>
    </body>
  </html>
  `;
}

/** -------- Menu handlers (thin wrappers with confirmation & toasts) -------- */
function ui_rebuildMasterFullReset_() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Rebuild Master (Full Reset)',
    'This will CLEAR rows in Master, Log, and Duplicates (headers/formatting kept), ' +
    'reset cursors & duplicate logs, then reimport everything from sources. Proceed?',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;
  try {
    SpreadsheetApp.getActive().toast('Starting full reset & reimport…', 'Data Consolidator', 5);
    resetMasterAndReimportAll();
    SpreadsheetApp.getActive().toast('Full rebuild complete ✅', 'Data Consolidator', 5);
    SpreadsheetApp.flush();
  } catch (e) {
    ui.alert('Full Reset Failed', e.message, ui.ButtonSet.OK);
  }
}

function ui_reimportHistoryNoClear_() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Reimport History (Keep Master)',
    'This will reset incremental cursors (and duplicate-group memory) and re-run import without clearing Master. Continue?',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;
  try {
    SpreadsheetApp.getActive().toast('Reimporting history (keeping Master)…', 'Data Consolidator', 5);
    resetCursorsAndReimportWithoutClearing();
    SpreadsheetApp.getActive().toast('Reimport complete ✅', 'Data Consolidator', 5);
    SpreadsheetApp.flush();
  } catch (e) {
    ui.alert('Reimport Failed', e.message, ui.ButtonSet.OK);
  }
}

function ui_initSourceLastRows_() {
  try {
    SpreadsheetApp.getActive().toast('Seeding source last rows…', 'Data Consolidator', 3);
    initSourceLastRows();
    SpreadsheetApp.getActive().toast('Seeded source last rows ✅', 'Data Consolidator', 3);
  } catch (e) {
    SpreadsheetApp.getUi().alert('initSourceLastRows failed', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function ui_initConsolidateLastRows_() {
  try {
    SpreadsheetApp.getActive().toast('Seeding consolidation cursors…', 'Data Consolidator', 3);
    initConsolidateLastRows();
    SpreadsheetApp.getActive().toast('Seeded consolidation cursors ✅', 'Data Consolidator', 3);
  } catch (e) {
    SpreadsheetApp.getUi().alert('initConsolidateLastRows failed', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function ui_createTriggersForSources_() {
  try {
    SpreadsheetApp.getActive().toast('Creating onEdit triggers for sources…', 'Data Consolidator', 5);
    createTriggersForSources();
    SpreadsheetApp.getActive().toast('Source triggers created ✅', 'Data Consolidator', 5);
  } catch (e) {
    SpreadsheetApp.getUi().alert('createTriggersForSources failed', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function ui_createDailyConsolidationTrigger_() {
  try {
    SpreadsheetApp.getActive().toast('Creating daily 2am consolidation trigger…', 'Data Consolidator', 5);
    createDailyConsolidationTrigger();
    SpreadsheetApp.getActive().toast('Daily trigger created ✅', 'Data Consolidator', 5);
  } catch (e) {
    SpreadsheetApp.getUi().alert('createDailyConsolidationTrigger failed', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function ui_openLogSheet_() {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(LOG_SHEET_NAME) || ss.insertSheet(LOG_SHEET_NAME);
    ss.setActiveSheet(sh);
    SpreadsheetApp.getActive().toast('Opened Log sheet', 'Data Consolidator', 2);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Open Log failed', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function buildAlertsHtmlForSourceNice(alerts) {
  function escText(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function escAttr(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  var html = '<div style="font-family: Roboto, Arial, sans-serif; padding:18px; color:#1d1d1f;">';
  html += '<div style="display:flex;align-items:center;margin-bottom:12px;">';
  html += '<div style="width:8px;height:32px;background:#1e88e5;border-radius:4px;margin-right:12px;"></div>';
  html += '<h2 style="margin:0;font-size:18px;color:#1e88e5;">Potential duplicates found</h2>';
  html += '</div>';
  html += '<div style="font-size:13px;color:#3c3c3c;margin-bottom:10px;">Matches were found in <strong>Master</strong>. Click a link to open that Master row in a new tab.</div>';

  alerts.forEach(function(a) {
    html += '<div style="border:1px solid #e6f0fb;background:#f7fbff;padding:10px;border-radius:6px;margin-bottom:10px;">';
    html += '<div style="font-weight:600;margin-bottom:6px;">Source row: Row ' + escText(a.sourceRow) + ' — ' + escText(a.sourceSheetName) + '</div>';
    html += '<ul style="margin:0 0 6px 16px;padding:0;">';

    a.matches.forEach(function(m) {
      var label = escText(m.fileName || '') + ' (' + escText(m.sheetName || '') + ') Row ' + escText(m.row);
      var href = '';
      try { href = buildLinkForMatch(m) || (m.url || ''); } catch (e) { href = (m.url || ''); }

      if (href) {
        html += '<li style="margin-bottom:6px;"><a href="' + escAttr(href) + '" target="_blank" rel="noopener noreferrer" style="color:#1565c0;text-decoration:none;font-weight:500;">' + label + '</a></li>';
      } else {
        html += '<li style="margin-bottom:6px;color:#333;">' + label + '</li>';
      }
    });

    html += '</ul>';
    html += '</div>';
  });

  html += '<div style="display:flex;justify-content:flex-end;margin-top:8px;">';
  html += '<button onclick="google.script.host.close()" style="background:#1e88e5;color:white;border:none;padding:8px 12px;border-radius:6px;font-weight:600;cursor:pointer">Close</button>';
  html += '</div>';
  html += '</div>';
  return html;
}