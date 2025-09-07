
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
        #chart-container { height: 400px; }
      </style>
    </head>
    <body>
      <div class="container mx-auto p-4 md:p-8">
        <h1 class="text-3xl md:text-4xl font-bold text-cyan-400 mb-2">BHG DeDuper Dashboard</h1>
        <p class="text-md md:text-lg text-gray-400 mb-8">Weekly duplicate percentages for each source sheet.</p>
        
        <div id="loading" class="text-center py-16">
          <p class="text-2xl animate-pulse">Loading Dashboard Data...</p>
        </div>
        
        <div id="chart-container" class="bg-gray-800 p-2 md:p-6 rounded-lg shadow-xl" style="display: none; position: relative;">
          <canvas id="statsChart"></canvas>
        </div>
        
        <div id="error" class="text-center py-16 text-red-400 text-2xl" style="display: none;">
        </div>
      </div>

      <script>
        document.addEventListener('DOMContentLoaded', function() {
          const today = new Date();
          const day = today.getDay();
          const diff = today.getDate() - day + (day === 0 ? -6 : 1);
          const monday = new Date(today.setDate(diff));
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
              errorDiv.textContent = 'No data available to display.';
              return;
          }

          document.getElementById('chart-container').style.display = 'block';

          // Sort data by percentage, descending
          data.sort((a, b) => b.percentage - a.percentage);

          const labels = data.map(item => item.sourceName);
          const percentages = data.map(item => item.percentage);
          
          const getHealthColor = (percentage) => {
              if (percentage <= 15) return 'rgba(74, 222, 128, 0.7)'; // green-400
              if (percentage < 30) return 'rgba(250, 204, 21, 0.7)'; // yellow-400
              return 'rgba(248, 113, 113, 0.7)'; // red-400
          };

          const backgroundColors = percentages.map(p => getHealthColor(p));
          
          const ctx = document.getElementById('statsChart').getContext('2d');
          new Chart(ctx, {
            type: 'bar',
            data: {
              labels: labels,
              datasets: [{
                label: 'Duplicate %',
                data: percentages,
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
                x: {
                  beginAtZero: true,
                  max: 100,
                  ticks: { color: '#9ca3af', font: { size: 14 } },
                  grid: { color: 'rgba(255, 255, 255, 0.1)' }
                },
                y: {
                  ticks: { color: '#d1d5db', font: { size: 14 } },
                  grid: { display: false }
                }
              },
              plugins: {
                legend: { display: false },
                tooltip: {
                  callbacks: {
                      label: function(context) {
                          let label = context.dataset.label || '';
                          if (label) { label += ': '; }
                          if (context.parsed.x !== null) { label += context.parsed.x + '%'; }
                          const sourceData = data[context.dataIndex];
                          label += ' (' + sourceData.totalDuplicates + ' / ' + sourceData.totalChecked + ' records)';
                          return label;
                      }
                  }
                }
              }
            }
          });

          const chartContainer = document.getElementById('chart-container');
          const newHeight = Math.max(400, data.length * 40); // 40px per bar, min 400px
          chartContainer.style.height = newHeight + 'px';
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
