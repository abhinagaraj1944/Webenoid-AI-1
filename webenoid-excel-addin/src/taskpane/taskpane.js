/* global Office, Excel */

Office.onReady(() => {

  document.getElementById("runBtn").onclick = runWebenoidAI;

  document.getElementById("prompt").addEventListener("keydown", e => {
    if (e.key === "Enter") {
      e.preventDefault();
      runWebenoidAI();
    }
  });

  const powerBtn = document.getElementById("openPowerBI");
  if (powerBtn) powerBtn.onclick = exportToPowerBI;
});


/* =====================================================
   UI HELPERS
===================================================== */

function scrollToBottom() {
  const box = document.getElementById("resultBox");
  box.scrollTop = box.scrollHeight;
}

function addMessage(text, type = "ai") {
  const box = document.getElementById("resultBox");

  const div = document.createElement("div");
  div.className = `message ${type}`;
  div.innerText = text;

  box.appendChild(div);
  scrollToBottom();
}

function showLoader() {
  const box = document.getElementById("resultBox");

  const loader = document.createElement("div");
  loader.className = "loader-message";
  loader.id = "aiLoader";
  loader.innerText = "🔍 Analyzing Excel data...";

  box.appendChild(loader);
  scrollToBottom();
}

function removeLoader() {
  const loader = document.getElementById("aiLoader");
  if (loader) loader.remove();
}


/* =====================================================
   KPI CARD (Power BI Style)
===================================================== */

function addKpiCard(title, value, is_percentage = false) {

  const box = document.getElementById("resultBox");

  const card = document.createElement("div");
  card.className = "kpi-card";

  let formatted = value;
  if (!isNaN(value)) {
    if (is_percentage) {
      // Assume value is a decimal (e.g., 0.15 representing 15%) or whole number (e.g., 15)
      // If it's already a whole number representing %, we might just append %
      // Let's multiply by 100 if it's less than or equal to 1, or just append % if it seems it's already a scaled percentage.
      // For safety, let's just append % if the backend format is unknown, or rely on NumberFormat.
      // A standard approach: Assume backend returns 0.15 for 15% and format it.
      // We'll use Intl.NumberFormat
      formatted = new Intl.NumberFormat('en-US', { style: 'percent', minimumFractionDigits: 2 }).format(value);
    } else {
      formatted = Number(value).toLocaleString();
    }
  }

  card.innerHTML = `
    <div style="font-size:12px;opacity:0.7">${title}</div>
    <div style="font-size:26px;font-weight:bold;margin-top:5px">${formatted}</div>
  `;

  box.appendChild(card);
  scrollToBottom();
}


/* =====================================================
   RESULT RENDERING
===================================================== */

function addResultCard(title, result) {

  const box = document.getElementById("resultBox");

  const card = document.createElement("div");
  card.className = "result-card";

  const header = document.createElement("div");
  header.className = "result-title";
  header.innerText = title;
  card.appendChild(header);

  if (Array.isArray(result) && result.length > 0 && typeof result[0] === "object") {

    const wrapper = document.createElement("div");
    wrapper.className = "table-wrapper";

    const table = document.createElement("table");
    table.className = "result-table";

    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");

    Object.keys(result[0]).forEach(key => {
      const th = document.createElement("th");
      th.innerText = key;
      headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");

    result.forEach(rowObj => {
      const tr = document.createElement("tr");

      Object.values(rowObj).forEach(val => {
        const td = document.createElement("td");
        td.innerText = val ?? "";
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    wrapper.appendChild(table);
    card.appendChild(wrapper);
  }
  else if (Array.isArray(result)) {
    result.forEach((item, i) => {
      const row = document.createElement("div");
      row.className = "result-list-item";
      row.innerText = `${i + 1}. ${item}`;
      card.appendChild(row);
    });
  }
  else if (typeof result === "number") {
    addKpiCard(title, result);
    return;
  }
  else if (typeof result === "object" && result !== null) {
    Object.entries(result).forEach(([k, v]) => {
      const row = document.createElement("div");
      row.className = "result-step";
      row.innerHTML = `<strong>${k}:</strong> ${v ?? ""}`;
      card.appendChild(row);
    });
  }
  else {
    const value = document.createElement("div");
    value.className = "result-value";
    value.innerText = result ?? "";
    card.appendChild(value);
  }

  box.appendChild(card);
  scrollToBottom();
}


/* =====================================================
   READ ALL SHEETS
===================================================== */

async function readAllSheetsData() {
  return Excel.run(async context => {

    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const result = {};

    for (const sheet of sheets.items) {
      const range = sheet.getUsedRange();
      range.load("values");
      await context.sync();

      if (!range.values || range.values.length < 2) continue;

      const headers = range.values[0];

      result[sheet.name] = range.values.slice(1).map(row => {
        const obj = {};
        headers.forEach((h, i) => {
          obj[String(h).trim()] = row[i];
        });
        return obj;
      });
    }

    return result;
  });
}


/* =====================================================
   MAIN AI FUNCTION
===================================================== */

async function runWebenoidAI() {

  try {

    const input = document.getElementById("prompt");
    const question = input.value.trim();
    if (!question) return;

    addMessage(question, "user");
    input.value = "";

    showLoader();

    const sheets = await readAllSheetsData();

    const res = await fetch("http://localhost:8000/analyze", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        question: question,
        data: sheets
      })
    });

    removeLoader();

    if (!res.ok) {
      addMessage("❌ Server error occurred.", "ai");
      return;
    }

    const raw = await res.json();

    let data = null;

    if (raw.success !== undefined) {
      if (raw.success === false) {
        addMessage(raw.error || "❌ Failed to analyze data.", "ai");
        return;
      }
      data = raw;
    }
    else if (raw.status === "success") {
      data = raw.data;
    }
    else {
      addMessage(raw?.message || "❌ Failed to analyze data.", "ai");
      return;
    }

    if (!data) {
      addMessage("❌ No data returned from server.", "ai");
      return;
    }

    if (data.type === "chart") {
      await createChartFromBackend(data);
      addMessage("📊 AI Chart created successfully.", "ai");
      if (data.insight) addMessage("💡 Insight: " + data.insight, "ai");
      return;
    }

    if (data.operation === "list") {
      addResultCard("RESULT", data.values);
      return;
    }

    if (data.operation === "count") {
      addKpiCard(data.title || "Total Count", data.row_count, data.is_percentage);
      return;
    }

    if (data.operation && data.operation.includes("_by_group")) {
      addResultCard("RESULT", data.data);
      return;
    }

    if (["sum", "average", "max", "min"].includes(data.operation)) {
      addKpiCard(data.title || data.operation.toUpperCase(), Object.values(data.data)[0], data.is_percentage);
      return;
    }

    if (question.toLowerCase().includes("dashboard")) {
      await buildDashboard();
      return;
    }


    // =============================
    // DASHBOARD HANDLER
    // =============================

    if (data.dashboard) {

      addMessage("🚀 AI Dashboard Created Successfully!", "ai");

      const dashboard = data.dashboard;

      addKpiCard("Total Rows", dashboard.total_rows);

      if (dashboard.total_numeric_sum) {
        addKpiCard(
          `Total ${dashboard.primary_numeric_column}`,
          dashboard.total_numeric_sum
        );
      }

      if (dashboard.top_grouped_data?.length > 0) {
        addResultCard("Top Categories", dashboard.top_grouped_data);
      }

      if (dashboard.trend_data?.length > 0) {
        addResultCard("Yearly Trend", dashboard.trend_data);
      }

      return;
    }
    addResultCard("RESULT", data.data || data);

  } catch (err) {
    console.error(err);
    removeLoader();
    addMessage("❌ Failed to analyze data.", "ai");
  }
}


/* =====================================================
   SAFE CHART CREATION (FULLY PRODUCTION FIXED)
===================================================== */

async function createChartFromBackend(config) {

  await Excel.run(async context => {

    const workbook = context.workbook;
    const sheet = workbook.worksheets.getActiveWorksheet();

    const categoryCol = config.category_column;

    // ✅ FIX: support both value_column and value_columns
    let valueCol = null;

    if (config.value_columns && config.value_columns.length > 0) {
      valueCol = config.value_columns[0];
    } else if (config.value_column) {
      valueCol = config.value_column;
    }

    if (!categoryCol || !valueCol) {
      console.error("Invalid chart config:", config);
      return;
    }

    const rows = config.data || [];

    let chartSheet;

    try {
      chartSheet = workbook.worksheets.getItem("AI_Chart_Data");
      chartSheet.delete();
      await context.sync();
    } catch (e) { }

    chartSheet = workbook.worksheets.add("AI_Chart_Data");

    // ✅ Build chart data safely
    const chartData = [];
    chartData.push([categoryCol, valueCol]);

    rows.forEach(row => {

      const categoryValue = row[categoryCol];

      let numericValue = row[valueCol];

      if (numericValue === undefined || numericValue === null || numericValue === "") {
        numericValue = 0;
      } else {
        numericValue = Number(numericValue);
        if (isNaN(numericValue)) numericValue = 0;

        // If the AI gives us a percentage between 0 and 100, we need to convert it visually 
        // back to a decimal (like 0.20 instead of 20) SO that Excel's numberFormat "0.00%"
        // correctly renders it as 20% instead of 2000%.
        if (config.is_percentage && numericValue > 1) {
          numericValue = numericValue / 100;
        }
      }

      chartData.push([categoryValue, numericValue]);
    });

    // ✅ SAFER RANGE CREATION
    const rowCount = chartData.length;
    const colCount = chartData[0].length;

    const range = chartSheet.getRangeByIndexes(
      0, 0, rowCount, colCount
    );

    range.values = chartData;

    await context.sync();

    // =============================
    // CHART TYPE LOGIC
    // =============================

    let chartType = "ColumnClustered";

    if (config.chart_type?.toLowerCase().includes("pie"))
      chartType = "Pie";
    else if (config.chart_type?.toLowerCase().includes("bar"))
      chartType = "BarClustered";
    else if (config.chart_type?.toLowerCase().includes("line"))
      chartType = "Line";
    else if (config.chart_type?.toLowerCase().includes("area"))
      chartType = "Area";

    const chart = sheet.charts.add(chartType, range, "Auto");

    chart.title.text = "AI Generated Chart";
    chart.height = 350;
    chart.width = 450;

    const dataLabels = chart.dataLabels;

    if (chartType === "Pie") {
      dataLabels.showCategoryName = true;
      dataLabels.showLegendKey = false;
      // In Excel Pie Charts, if the user wants percentage natively handled, we enable it.
      dataLabels.showPercentage = true;
      dataLabels.showValue = true;
    } else {
      dataLabels.showValue = true;
      if (config.is_percentage) {
        dataLabels.numberFormat = "0.00%";
      }
    }

    await context.sync();
  });
}


/* =====================================================
   POWER BI EXPORT
===================================================== */

async function exportToPowerBI() {

  await Excel.run(async context => {

    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load("values");
    await context.sync();

    const csv = range.values.map(r => r.join(",")).join("\n");

    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "webenoid_export.csv";
    a.click();

    URL.revokeObjectURL(url);
  });

  addMessage("✅ Exported for Power BI.", "ai");
}

async function buildDashboard() {

  addMessage("🚀 Building AI Dashboard...", "ai");

  const sheets = await readAllSheetsData();

  const res = await fetch("http://localhost:8000/dashboard", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ question: "dashboard", data: sheets })
  });

  const raw = await res.json();

  if (!raw.success) {
    addMessage("❌ Dashboard creation failed.", "ai");
    return;
  }

  const data = raw.dashboard;

  await Excel.run(async context => {

    const workbook = context.workbook;

    let dashSheet;
    try {
      dashSheet = workbook.worksheets.getItem("AI_Dashboard");
      dashSheet.delete();
      await context.sync();
    } catch (e) { }

    dashSheet = workbook.worksheets.add("AI_Dashboard");

    // ========================
    // KPI CARDS
    // ========================

    dashSheet.getRange("A1").values = [["Total Employees"]];
    dashSheet.getRange("A2").values = [[data.total_employees]];

    dashSheet.getRange("C1").values = [["Total Salary"]];
    dashSheet.getRange("C2").values = [[data.total_salary]];

    await context.sync();

    // ========================
    // TOP DEPARTMENTS CHART
    // ========================

    const deptData = [["Department", "Salary"]];

    data.top_departments.forEach(r => {
      deptData.push([r.department, r.salary]);
    });

    const deptRange = dashSheet.getRange("A5")
      .getResizedRange(deptData.length - 1, 1);

    deptRange.values = deptData;

    const deptChart = dashSheet.charts.add(
      "ColumnClustered",
      deptRange,
      "Auto"
    );

    deptChart.top = 120;
    deptChart.left = 20;
    deptChart.height = 300;
    deptChart.width = 400;

    // ========================
    // STATE PIE CHART
    // ========================

    const stateData = [["State", "Count"]];

    data.state_distribution.forEach(r => {
      stateData.push([r.state, r.Count]);
    });

    const stateRange = dashSheet.getRange("H5")
      .getResizedRange(stateData.length - 1, 1);

    stateRange.values = stateData;

    const stateChart = dashSheet.charts.add(
      "Pie",
      stateRange,
      "Auto"
    );

    stateChart.top = 120;
    stateChart.left = 500;
    stateChart.height = 300;
    stateChart.width = 350;

    await context.sync();
  });

  addMessage("✅ AI Dashboard Created Successfully!", "ai");
}

