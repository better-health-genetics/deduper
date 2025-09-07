/** ===========================================================================
 * 
 *                        UI & MENU FUNCTIONS
 * 
 * ===========================================================================/

/**
 * Opens the web dashboard in a new tab using the configured URL.
 */
function ui_openDashboard_() {
  if (!WEB_APP_URL || WEB_APP_URL === 'PASTE_YOUR_WEB_APP_URL_HERE') {
    SpreadsheetApp.getUi().alert('Dashboard URL Not Configured', 'Please ask your administrator to set the WEB_APP_URL in the Config.js script file.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  var html = HtmlService.createHtmlOutput('<script>window.open("' + WEB_APP_URL + '", "_blank");</script>')
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
 * This version is now fully interactive with date pickers and an AI summary feature.
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
        .spinner { border: 2px solid #f3f4f5; border-top: 2px solid #3b82f6; border-radius: 50%; width: 1rem; height: 1rem; animation: spin 1s linear infinite; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
      </style>
    </head>
    <body>
      <div class="container mx-auto p-4 md:p-8">
        <div class="flex flex-wrap justify-between items-center mb-2">
          <h1 class="text-3xl md:text-4xl font-bold text-cyan-400">BHG DeDuper Dashboard</h1>
        </div>
        <p id="subtitle" class="text-md text-gray-400 mb-6">Loading data for the current week...</p>
        
        <!-- Controls -->
        <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 mb-8">
            <div class="md:col-span-2 lg:col-span-2 bg-gray-800 border border-gray-700 p-3 rounded-lg flex flex-col space-y-2">
                <div class="grid grid-cols-2 gap-2">
                    <input type="date" id="start-date" class="w-full text-sm bg-gray-700 border border-gray-600 rounded py-1 px-2 text-gray-200 focus:outline-none focus:ring-1 focus:ring-cyan-500">
                    <input type="date" id="end-date" class="w-full text-sm bg-gray-700 border border-gray-600 rounded py-1 px-2 text-gray-200 focus:outline-none focus:ring-1 focus:ring-cyan-500">
                </div>
                <div id="filter-buttons" class="flex justify-center items-center gap-2 pt-1">
                    <button data-filter="WTD" class="filter-btn px-3 py-1 text-xs rounded-full transition-colors">WTD</button>
                    <button data-filter="MTD" class="filter-btn px-3 py-1 text-xs rounded-full transition-colors">MTD</button>
                    <button data-filter="YTD" class="filter-btn px-3 py-1 text-xs rounded-full transition-colors">YTD</button>
                    <button data-filter="ALL" class="filter-btn px-3 py-1 text-xs rounded-full transition-colors">ALL</button>
                </div>
            </div>
            <div class="md:col-span-2 lg:col-span-2 bg-gray-800 border border-gray-700 rounded-lg p-3 space-y-2">
                <div class="flex items-center justify-between">
                    <h3 class="text-md font-semibold text-gray-200 flex items-center gap-2">
                        <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" viewBox="0 0 24 24" fill="currentColor"><path d="M12.5 5.515c1.334-2.134 4.133-2.68 6.267-.985c2.134 1.694 2.68 4.827.985 6.961l-7.252 8.016l-7.252-8.016c-1.695-2.134-.94-5.267 1.194-6.961c2.134-1.695 4.933-.94 6.267.985Z"/></svg>
                        AI Health Summary
                    </h3>
                    <button id="gemini-button" class="flex items-center gap-1.5 px-2.5 py-1 text-xs font-semibold text-white bg-gradient-to-r from-cyan-600 to-blue-600 rounded-md hover:opacity-90 transition-opacity disabled:opacity-50 disabled:cursor-not-allowed">
                        <span id="gemini-button-text">Generate Analysis</span>
                        <div id="gemini-spinner" class="spinner" style="display: none;"></div>
                    </button>
                </div>
                <div id="gemini-error" class="bg-red-900/50 text-red-300 text-sm font-medium text-center p-1.5 rounded-md w-full" style="display: none;"></div>
                <div id="gemini-summary-container" class="bg-gray-900/50 p-2.5 rounded-md border-l-4 border-cyan-500" style="display: none;">
                    <p id="gemini-summary" class="text-sm text-gray-300 italic"></p>
                </div>
            </div>
        </div>


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
                  <tbody id="stats-table-body"></tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
        
        <div id="error-container" class="text-center py-16 text-red-400 text-2xl" style="display: none;"></div>
      </div>

      <script>
        let chartInstance = null;
        let currentStats = [];
        let activeFilter = 'WTD';

        const loadingDiv = document.getElementById('loading');
        const contentDiv = document.getElementById('dashboard-content');
        const errorDiv = document.getElementById('error-container');
        const subtitle = document.getElementById('subtitle');
        const startDateInput = document.getElementById('start-date');
        const endDateInput = document.getElementById('end-date');
        const filterButtons = document.querySelectorAll('.filter-btn');
        
        const geminiButton = document.getElementById('gemini-button');
        const geminiButtonText = document.getElementById('gemini-button-text');
        const geminiSpinner = document.getElementById('gemini-spinner');
        const geminiError = document.getElementById('gemini-error');
        const geminiSummaryContainer = document.getElementById('gemini-summary-container');
        const geminiSummary = document.getElementById('gemini-summary');
        
        const getMonday = (d) => {
            d = new Date(d);
            const day = d.getDay(), diff = d.getDate() - day + (day === 0 ? -6 : 1);
            return new Date(d.setDate(diff));
        };
        const dateToInputValue = (date) => date ? date.toISOString().split('T')[0] : '';
        
        function updateFilterButtons() {
          filterButtons.forEach(btn => {
            if (btn.dataset.filter === activeFilter) {
              btn.classList.add('bg-cyan-600', 'text-white', 'font-semibold');
              btn.classList.remove('bg-gray-600', 'text-gray-300', 'hover:bg-gray-500');
            } else {
              btn.classList.remove('bg-cyan-600', 'text-white', 'font-semibold');
              btn.classList.add('bg-gray-600', 'text-gray-300', 'hover:bg-gray-500');
            }
          });
        }
        
        function setDateRange(filter) {
            const today = new Date();
            let newStart = new Date(today);
            let newEnd = new Date(today);

            if (filter === 'WTD') newStart = getMonday(today);
            else if (filter === 'MTD') newStart = new Date(today.getFullYear(), today.getMonth(), 1);
            else if (filter === 'YTD') newStart = new Date(today.getFullYear(), 0, 1);
            else if (filter === 'ALL') { newStart = null; newEnd = null; }
            
            activeFilter = filter;
            updateFilterButtons();
            
            startDateInput.value = dateToInputValue(newStart);
            endDateInput.value = dateToInputValue(newEnd);
            
            fetchData(newStart, newEnd);
        };
        
        function fetchData(start, end) {
            loadingDiv.style.display = 'block';
            contentDiv.style.display = 'none';
            errorDiv.style.display = 'none';
            geminiButton.disabled = true;

            const startDateStr = start ? dateToInputValue(start) : null;
            const endDateStr = end ? dateToInputValue(end) : null;
            
            google.script.run
                .withSuccessHandler((data) => onDataLoaded(data, start, end))
                .withFailureHandler(onDataError)
                .getAdminDashboardData(startDateStr, endDateStr);
        }

        function onDataError(error) {
          loadingDiv.style.display = 'none';
          errorDiv.style.display = 'block';
          errorDiv.textContent = 'Error loading data: ' + error.message;
        }

        function onDataLoaded(data, start, end) {
          loadingDiv.style.display = 'none';
          geminiButton.disabled = false;
          
          currentStats = data || []; // Store stats for AI summary

          // Update subtitle
          if (activeFilter === 'ALL') {
            subtitle.textContent = 'Displaying stats for all records.';
          } else if (start && end) {
            const options = { year: 'numeric', month: 'short', day: 'numeric' };
            subtitle.textContent = 'Displaying stats from ' + start.toLocaleDateString(undefined, options) + ' to ' + end.toLocaleDateString(undefined, options) + '.';
          }
          
          if (!data || data.length === 0) {
              errorDiv.style.display = 'block';
              errorDiv.textContent = 'No data available to display for this period.';
              return;
          }

          contentDiv.style.display = 'block';
          data.sort((a, b) => b.percentage - a.percentage);

          // Populate Cards
          const totalRecords = data.reduce((sum, item) => sum + item.totalChecked, 0);
          const totalDuplicates = data.reduce((sum, item) => sum + item.totalDuplicates, 0);
          const overallRate = totalRecords > 0 ? Math.round((totalDuplicates / totalRecords) * 100) : 0;
          
          document.getElementById('overall-rate').textContent = overallRate + '%';
          document.getElementById('total-duplicates').textContent = totalDuplicates.toLocaleString();
          document.getElementById('total-records').textContent = totalRecords.toLocaleString();

          // Populate Table
          const tableBody = document.getElementById('stats-table-body');
          tableBody.innerHTML = data.map(item => {
            const healthColor = item.percentage <= 15 ? 'text-green-400' : (item.percentage < 30 ? 'text-yellow-400' : 'text-red-400');
            return \`
              <tr class="border-b border-gray-700 hover:bg-gray-700/50">
                <td class="px-4 py-2 font-medium whitespace-nowrap">\${item.sourceName}</td>
                <td class="px-4 py-2 text-right font-bold \${healthColor}">\${item.percentage}%</td>
                <td class="px-4 py-2 text-right">\${item.totalDuplicates.toLocaleString()}</td>
                <td class="px-4 py-2 text-right">\${item.totalChecked.toLocaleString()}</td>
              </tr>
            \`;
          }).join('');

          // Render Chart
          const top15 = data.slice(0, 15);
          const chartLabels = top15.map(item => item.sourceName);
          const chartPercentages = top15.map(item => item.percentage);
          const getHealthColor = (p, op = '0.7') => p <= 15 ? \`rgba(74, 222, 128, \${op})\` : (p < 30 ? \`rgba(250, 204, 21, \${op})\` : \`rgba(248, 113, 113, \${op})\`);

          const chartData = {
              labels: chartLabels,
              datasets: [{
                label: 'Duplicate %',
                data: chartPercentages,
                backgroundColor: chartPercentages.map(p => getHealthColor(p)),
                borderColor: chartPercentages.map(p => getHealthColor(p, '1')),
                borderWidth: 1
              }]
            };

          if (chartInstance) {
              chartInstance.data = chartData;
              chartInstance.update();
          } else {
              const ctx = document.getElementById('statsChart').getContext('2d');
              chartInstance = new Chart(ctx, { type: 'bar', data: chartData, options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, scales: { x: { beginAtZero: true, max: 100, ticks: { color: '#9ca3af' }, grid: { color: 'rgba(255, 255, 255, 0.1)' } }, y: { ticks: { color: '#d1d5db' }, grid: { display: false } } }, plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => \`\${c.dataset.label || ''}: \${c.parsed.x}%\` } } } } });
          }
        }
        
        function handleGenerateSummary() {
            if (currentStats.length === 0) {
                geminiError.textContent = 'There is no data to analyze for the selected period.';
                geminiError.style.display = 'block';
                return;
            }
            geminiButton.disabled = true;
            geminiButtonText.style.display = 'none';
            geminiSpinner.style.display = 'block';
            geminiError.style.display = 'none';
            geminiSummaryContainer.style.display = 'none';

            google.script.run
                .withSuccessHandler(onSummarySuccess)
                .withFailureHandler(onSummaryError)
                .getGeminiHealthSummary(currentStats);
        }

        function onSummarySuccess(summary) {
            geminiButton.disabled = false;
            geminiButtonText.style.display = 'block';
            geminiSpinner.style.display = 'none';
            geminiSummary.textContent = summary;
            geminiSummaryContainer.style.display = 'block';
        }

        function onSummaryError(error) {
            geminiButton.disabled = false;
            geminiButtonText.style.display = 'block';
            geminiSpinner.style.display = 'none';
            geminiError.textContent = error.message || 'Failed to generate summary.';
            geminiError.style.display = 'block';
        }

        // --- Event Listeners ---
        document.addEventListener('DOMContentLoaded', () => setDateRange('WTD'));
        filterButtons.forEach(btn => btn.addEventListener('click', () => setDateRange(btn.dataset.filter)));
        startDateInput.addEventListener('change', () => { activeFilter = ''; updateFilterButtons(); fetchData(new Date(startDateInput.value+'T00:00:00'), new Date(endDateInput.value+'T00:00:00')); });
        endDateInput.addEventListener('change', () => { activeFilter = ''; updateFilterButtons(); fetchData(new Date(startDateInput.value+'T00:00:00'), new Date(endDateInput.value+'T00:00:00')); });
        geminiButton.addEventListener('click', handleGenerateSummary);

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