/* global Office, Excel */

// State for trial preview limit
let promptCount = 0;
const PROMPT_LIMIT = 4;
let isLoggedIn = false;
let userEmail = "";
let chatHistory = [];

Office.onReady(() => {
  document.getElementById("runBtn").onclick = runWebenoidAI;

  document.getElementById("prompt").addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      runWebenoidAI();
    }
  });



  // Suggestion chip click handlers
  document.querySelectorAll(".chip").forEach((chip) => {
    chip.addEventListener("click", () => {
      const input = document.getElementById("prompt");
      input.value = chip.innerText;
      runWebenoidAI();
    });
  });

  // ===== MODALS =====
  document.getElementById("closeModalBtn")?.addEventListener("click", hideLoginModal);
  document.getElementById("authSubmitBtn")?.addEventListener("click", handleAuthSubmit);
  document.getElementById("authToggleLink")?.addEventListener("click", toggleAuthMode);

  document.getElementById("closeHistoryBtn")?.addEventListener("click", hideHistoryModal);
  document.getElementById("clearHistoryBtn")?.addEventListener("click", clearHistory);

  // ===== THREE DOTS MENU =====
  const menuBtn = document.getElementById("menuBtn");
  const menuDropdown = document.getElementById("menuDropdown");

  if (menuBtn && menuDropdown) {
    // Toggle dropdown on click
    menuBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      menuDropdown.classList.toggle("hidden");
    });

    // Close dropdown when clicking anywhere else
    document.addEventListener("click", () => {
      menuDropdown.classList.add("hidden");
    });

    // Prevent dropdown from closing when clicking inside it
    menuDropdown.addEventListener("click", (e) => {
      e.stopPropagation();
    });

    // Handle menu item actions
    menuDropdown.querySelectorAll(".menu-item").forEach((item) => {
      item.addEventListener("click", () => {
        const action = item.getAttribute("data-action");
        menuDropdown.classList.add("hidden");

        switch (action) {
          case "login":
            showLoginModal();
            break;
          case "logout":
            isLoggedIn = false;
            userEmail = "";
            promptCount = 0; // Reset preview session
            document.getElementById("menuItemLogin")?.classList.remove("hidden");
            document.getElementById("menuItemProfile")?.classList.add("hidden");
            document.getElementById("menuItemLogout")?.classList.add("hidden");
            document.getElementById("menuDivider")?.classList.add("hidden");
            showToast("🚪 Signed out successfully.", "info");
            break;
          case "profile":
            showToast(`👤 Signed in as: ${userEmail}`, "info");
            break;
          case "export":
            exportToPowerBI();
            break;
          case "clear":
            clearChat();
            break;
          case "history":
            showHistoryModal();
            break;
          case "help":
            addMessage("💡 Ask questions about your Excel data like:\n• \"What are the products?\"\n• \"Total price amount?\"\n• \"Show me a chart of sales by region\"", "ai");
            break;
          case "about":
            addMessage("ℹ️ WebEnoid AI — Intelligent Excel Assistant\nPowered by AI to analyze your spreadsheet data.", "ai");
            break;
        }
      });
    });
  }
});

/* =====================================================
   UI HELPERS
===================================================== */

function scrollToBottom() {
  const box = document.getElementById("resultBox");
  box.scrollTop = box.scrollHeight;
}

function clearChat() {
  const box = document.getElementById("resultBox");
  box.innerHTML = `
    <div class="welcome-card">
      <div class="welcome-icon"></div>
      <div class="welcome-title">Welcome to WebEnoid</div>
      <div class="welcome-subtitle">Ask anything about your Excel data — counts, charts, details & more.</div>
    </div>
  `;
  // Also clear the local history array
  chatHistory = [];
}

function addMessage(text, type = "ai") {
  const box = document.getElementById("resultBox");

  // Remove welcome card on first message
  const welcome = box.querySelector(".welcome-card");
  if (welcome) welcome.remove();

  const div = document.createElement("div");
  div.className = `message ${type}`;
  div.innerText = text;

  box.appendChild(div);
  scrollToBottom();
}

function showToast(message, type = "info", duration = 3000) {
  let container = document.getElementById("toast-container");
  if (!container) {
    container = document.createElement("div");
    container.id = "toast-container";
    document.body.appendChild(container);
  }

  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.innerText = message;
  container.appendChild(toast);

  setTimeout(() => {
    toast.classList.add("hide");
    setTimeout(() => toast.remove(), 400);
  }, duration);
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
   LOGIN / SIGNUP MODAL UI
===================================================== */

let authMode = "login"; // "login" or "signup"

function showLoginModal() {
  document.getElementById("loginModal").classList.remove("hidden");
  document.getElementById("authErrorMsg").innerText = "";
}

function hideLoginModal() {
  document.getElementById("loginModal").classList.add("hidden");
  document.getElementById("authErrorMsg").innerText = "";
}

function toggleAuthMode(e) {
  e.preventDefault();

  const title = document.getElementById("authTitle");
  const subtitle = document.getElementById("authSubtitle");
  const signupFields = document.getElementById("signupFields");
  const submitBtn = document.getElementById("authSubmitBtn");
  const toggleText = document.getElementById("authToggleText");
  const toggleLink = document.getElementById("authToggleLink");
  const errorMsg = document.getElementById("authErrorMsg");

  errorMsg.innerText = "";

  if (authMode === "login") {
    authMode = "signup";
    title.innerText = "Create an Account";
    subtitle.innerText = "Sign up to start analyzing your data.";

    // Move common fields inside the body layout correctly (HTML already handles position but we unhide signup fields)
    signupFields.classList.remove("hidden");

    submitBtn.innerText = "Sign Up";
    toggleText.innerText = "Already have an account?";
    toggleLink.innerText = "Sign In";
  } else {
    authMode = "login";
    title.innerText = "Sign In Required";
    subtitle.innerText = "Sign in to continue analyzing your data.";

    signupFields.classList.add("hidden");

    submitBtn.innerText = "Sign In";
    toggleText.innerText = "Don't have an account?";
    toggleLink.innerText = "Sign Up";
  }
}

async function handleAuthSubmit() {
  const email = document.getElementById("authEmail").value.trim();
  const password = document.getElementById("authPassword").value.trim();
  const errorMsg = document.getElementById("authErrorMsg");

  errorMsg.innerText = "";

  if (!email || !password) {
    errorMsg.innerText = "⚠️ Please enter both email and password.";
    return;
  }

  const submitBtn = document.getElementById("authSubmitBtn");
  const originalText = submitBtn.innerText;
  submitBtn.innerText = "Please wait...";
  submitBtn.disabled = true;

  try {
    if (authMode === "signup") {
      const name = document.getElementById("signupName").value.trim();
      const phone = document.getElementById("signupPhone").value.trim();

      if (!name || !phone) {
        errorMsg.innerText = "⚠️ Please enter your full name and phone number.";
        submitBtn.innerText = originalText;
        submitBtn.disabled = false;
        return;
      }

      const res = await fetch("https://willfully-rubricated-tianna.ngrok-free.dev/auth/signup", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name, email, phone, password })
      });

      const data = await res.json();

      if (!res.ok) {
        errorMsg.innerText = data.detail || "❌ Failed to create account.";
        submitBtn.innerText = originalText;
        submitBtn.disabled = false;
        return;
      }

      showAuthSuccess(data.user.email, data.message);

    } else {
      // Login mode
      const res = await fetch("https://willfully-rubricated-tianna.ngrok-free.dev/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email, password })
      });

      const data = await res.json();

      if (!res.ok) {
        errorMsg.innerText = data.detail || "❌ Invalid email or password.";
        submitBtn.innerText = originalText;
        submitBtn.disabled = false;
        return;
      }

      showAuthSuccess(data.user.email, data.message);
    }
  } catch (err) {
    errorMsg.innerText = "❌ Network error. Please try again.";
  } finally {
    submitBtn.innerText = originalText;
    submitBtn.disabled = false;
  }
}

function showAuthSuccess(email, msg) {
  isLoggedIn = true;
  userEmail = email;
  hideLoginModal();

  document.getElementById("authEmail").value = "";
  document.getElementById("authPassword").value = "";
  document.getElementById("signupName").value = "";
  document.getElementById("signupPhone").value = "";

  // Swap menu items
  document.getElementById("menuItemLogin")?.classList.add("hidden");

  const profileItem = document.getElementById("menuItemProfile");
  if (profileItem) {
    profileItem.classList.remove("hidden");
    const emailTxt = document.getElementById("profileEmailTxt");
    if (emailTxt) emailTxt.innerText = email;
  }

  document.getElementById("menuItemLogout")?.classList.remove("hidden");
  document.getElementById("menuDivider")?.classList.remove("hidden");

  showToast(`✅ ${msg}`, "success");
}

/* =====================================================
   HISTORY MODAL UI
===================================================== */

function showHistoryModal() {
  const modal = document.getElementById("historyModal");
  const container = document.getElementById("historyListContainer");

  if (modal) modal.classList.remove("hidden");
  if (!container) return;

  container.innerHTML = "";

  if (chatHistory.length === 0) {
    container.innerHTML = `<div class="empty-history">You haven't asked any questions yet.</div>`;
    return;
  }

  // Show newest at the top
  const reversedHistory = [...chatHistory].reverse();

  reversedHistory.forEach((q) => {
    const item = document.createElement("div");
    item.className = "history-item";
    item.innerHTML = `<span class="history-item-icon">🕒</span> <span>${q}</span>`;

    // Clicking a history item runs that query again!
    item.addEventListener("click", () => {
      hideHistoryModal();
      document.getElementById("prompt").value = q;
      runWebenoidAI();
    });

    container.appendChild(item);
  });
}

function hideHistoryModal() {
  document.getElementById("historyModal")?.classList.add("hidden");
}

function clearHistory() {
  chatHistory = [];
  showHistoryModal(); // Refresh the modal view
  showToast("🕒 History cleared.", "info");
}

/* =====================================================
   KPI CARD (Power BI Style)
===================================================== */

function addKpiCard(title, value, is_percentage = false, label = null) {
  const box = document.getElementById("resultBox");

  // Remove welcome card
  const welcome = box.querySelector(".welcome-card");
  if (welcome) welcome.remove();

  const card = document.createElement("div");
  card.className = "kpi-card";

  let formatted = value;
  if (!isNaN(value)) {
    if (is_percentage) {
      formatted = new Intl.NumberFormat("en-US", {
        style: "percent",
        minimumFractionDigits: 2,
      }).format(value);
    } else {
      formatted = Number(value).toLocaleString();
    }
  }

  let labelHtml = "";
  if (label) {
    labelHtml = `<div style="font-size:13px;margin-top:6px;color:#86868b;font-weight:500"> ${label}</div>`;
  }

  card.innerHTML = `
    <div style="font-size:11px;text-transform:uppercase;letter-spacing:0.5px;color:#a1a1a6;font-weight:500">${title}</div>
    <div style="font-size:24px;font-weight:700;margin-top:6px;color:#f5f5f7;word-break:break-all;overflow-wrap:break-word">${formatted}</div>
    ${labelHtml}
  `;

  box.appendChild(card);
  scrollToBottom();
}

/* =====================================================
   COMPARISON CARD
===================================================== */

function addComparisonCard(data) {
  const box = document.getElementById("resultBox");

  // Remove welcome card
  const welcome = box.querySelector(".welcome-card");
  if (welcome) welcome.remove();

  const card = document.createElement("div");
  card.className = "comparison-card";

  const items = data.items || [];
  const higher = data.higher || "";
  const metric = data.metric || "";

  // Build items HTML
  let itemsHtml = "";
  items.forEach((item) => {
    const isWinner = item.name === higher;
    const formatted = !isNaN(item.value) ? Number(item.value).toLocaleString() : item.value;
    itemsHtml += `
      <div class="comparison-item ${isWinner ? "winner" : ""}">
        <div class="comparison-item-name">${item.name}</div>
        <div class="comparison-item-value">${formatted}</div>
      </div>
    `;
  });

  // Format difference
  const diff = data.difference || 0;
  const diffFormatted = !isNaN(diff) ? Number(diff).toLocaleString() : diff;
  const diffPercent = data.difference_percent || 0;

  // Build difference HTML
  let diffHtml = "";
  if (diff === 0 || !higher) {
    diffHtml = `<div class="comparison-equal">Both values are equal</div>`;
  } else {
    diffHtml = `
      <div class="comparison-diff">
        <div class="comparison-diff-text"><strong>${higher}</strong> leads by <strong>${diffFormatted}</strong></div>
        <div class="comparison-diff-percent">${diffPercent}% difference${metric ? " in " + metric : ""}</div>
      </div>
    `;
  }

  card.innerHTML = `
    <div class="comparison-title">${data.title || "Comparison"}</div>
    <div class="comparison-items">
      ${itemsHtml}
    </div>
    ${diffHtml}
  `;

  box.appendChild(card);
  scrollToBottom();
}

/* =====================================================
   FORMULA CARD
===================================================== */

function addFormulaCard(data) {
  const box = document.getElementById("resultBox");

  const welcome = box.querySelector(".welcome-card");
  if (welcome) welcome.remove();

  const card = document.createElement("div");
  card.className = "formula-code-box";
  card.style.marginBottom = "12px";
  card.style.marginTop = "8px";
  card.innerText = data.formula || "";

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

    Object.keys(result[0]).forEach((key) => {
      const th = document.createElement("th");
      th.innerText = key;
      headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");

    result.forEach((rowObj) => {
      const tr = document.createElement("tr");

      Object.values(rowObj).forEach((val) => {
        const td = document.createElement("td");
        td.innerText = val ?? "";
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    wrapper.appendChild(table);
    card.appendChild(wrapper);
  } else if (Array.isArray(result)) {
    result.forEach((item, i) => {
      const row = document.createElement("div");
      row.className = "result-list-item";
      row.innerText = `${i + 1}. ${item}`;
      card.appendChild(row);
    });
  } else if (typeof result === "number") {
    addKpiCard(title, result);
    return;
  } else if (typeof result === "object" && result !== null) {
    Object.entries(result).forEach(([k, v]) => {
      const row = document.createElement("div");
      row.className = "result-step";
      row.innerHTML = `<strong>${k}:</strong> ${v ?? ""}`;
      card.appendChild(row);
    });
  } else {
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
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    // Batch request data from all sheets at once to save time
    const sheetDataRequests = sheets.items.map((sheet) => {
      const range = sheet.getUsedRange();
      range.load("values");
      return { name: sheet.name, range: range };
    });

    // Single sync for all data! This fixes the "late response" slowness.
    await context.sync();

    const result = {};

    sheetDataRequests.forEach((item) => {
      const values = item.range.values;
      if (!values || values.length < 1) return;

      // 🔥 SMART HEADER DETECTION
      // Find the first row that has at least 2 non-empty cells
      let headerIndex = 0;
      for (let i = 0; i < values.length; i++) {
        const nonEmptyCount = values[i].filter(v => v !== null && v !== "").length;
        if (nonEmptyCount >= 2) {
          headerIndex = i;
          break;
        }
      }

      const headers = values[headerIndex].map(h => String(h || "").trim());
      const dataRows = values.slice(headerIndex + 1);

      if (dataRows.length === 0) return;

      result[item.name] = dataRows.map((row) => {
        const obj = {};
        headers.forEach((h, i) => {
          if (h) { // Only add if header is not empty
            obj[h] = row[i];
          }
        });
        return obj;
      });
    });

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

    // ✅ Limit Check: Prompt for Login after 4 requests
    if (!isLoggedIn && promptCount >= PROMPT_LIMIT) {
      showLoginModal();
      return;
    }

    addMessage(question, "user");
    input.value = "";
    promptCount++;
    chatHistory.push(question);

    showLoader();

    // ✅ Intercept dashboard requests BEFORE calling the AI
    if (question.toLowerCase().includes("dashboard")) {
      removeLoader();
      await buildDashboard();
      return;
    }

    const sheets = await readAllSheetsData();

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 60000);

    const res = await fetch("https://willfully-rubricated-tianna.ngrok-free.dev/analyze", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "ngrok-skip-browser-warning": "69420"
      },
      body: JSON.stringify({
        question: question,
        data: sheets,
        user_email: isLoggedIn ? userEmail : null,
        user_name: isLoggedIn && userEmail ? userEmail.split('@')[0] : null
      }),
      signal: controller.signal,
    });

    clearTimeout(timeoutId);

    removeLoader();

    if (!res.ok) {
      addMessage("❌ I'm having trouble connecting to the AI engine. Please try again later.", "ai");
      return;
    }

    const raw = await res.json();

    let data = null;

    if (raw.success !== undefined) {
      if (raw.success === false) {
        addMessage(raw.error || "❌ I couldn't analyze that. Please try rephrasing your question.", "ai");
        return;
      }
      data = raw;
    } else if (raw.status === "success") {
      data = raw.data;
    } else {
      addMessage(raw?.message || "❌ I'm sorry, I couldn't analyze that. Could you try rephrasing your question?", "ai");
      return;
    }

    if (!data) {
      addMessage("❌ No analysis result was found for your request.", "ai");
      return;
    }

    if (data.type === "chart") {
      await createChartFromBackend(data);
      addMessage("📊 AI Chart created successfully.", "ai");
      if (data.insight) addMessage("💡 Insight: " + data.insight, "ai");
      return;
    }

    if (data.operation === "conversation" || data.operation === "message") {
      addMessage(data.message || data.text || "👋 Hello! How can I help you with your Excel data?", "ai");
      return;
    }

    if (data.operation === "formula") {
      addFormulaCard(data);
      return;
    }

    if (data.operation === "comparison") {
      addComparisonCard(data);
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

    if (data.operation === "details") {
      const rows = data.data || [];
      if (rows.length === 0) {
        addMessage("ℹ️ No matching records found.", "ai");
      } else {
        addMessage(`📋 Found ${rows.length} record(s)`, "ai");
        addResultCard("Details", rows);
      }
      return;
    }

    if (data.operation && data.operation.includes("_by_group")) {
      addResultCard("RESULT", data.data);
      return;
    }

    if (["sum", "average", "max", "min"].includes(data.operation)) {
      addKpiCard(
        data.title || data.operation.toUpperCase(),
        data.data.value,
        data.is_percentage,
        data.data.label || null
      );
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
        addKpiCard(`Total ${dashboard.primary_numeric_column}`, dashboard.total_numeric_sum);
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
    if (err.name === "AbortError") {
      addMessage("⏱️ Request timed out. Please try again.", "ai");
    } else if (err.message?.includes("Failed to fetch") || err.message?.includes("NetworkError")) {
      addMessage("🔌 Connection failed. Please ensure your Webenoid AI engine is active.", "ai");
    } else {
      addMessage("❌ I couldn't analyze that. Please try rephrasing your question.", "ai");
    }
  }
}

/* =====================================================
   SAFE CHART CREATION (FULLY PRODUCTION FIXED)
===================================================== */

async function createChartFromBackend(config) {
  await Excel.run(async (context) => {
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

    // Compute total for percentage breakdown (used in chat message, not written to chart)
    let totalSum = 0;
    rows.forEach((row) => {
      const v = Number(row[valueCol]);
      if (!isNaN(v)) totalSum += v;
    });

    rows.forEach((row) => {
      const categoryValue = row[categoryCol];
      let numericValue = row[valueCol];

      if (numericValue === undefined || numericValue === null || numericValue === "") {
        numericValue = 0;
      } else {
        numericValue = Number(numericValue);
        if (isNaN(numericValue)) numericValue = 0;
        // ✅ Always write RAW values to Excel so bars are correctly sized.
        // Percentage context is shown in the companion chat message below.
      }

      chartData.push([categoryValue, numericValue]);
    });

    // ✅ SAFER RANGE CREATION
    const rowCount = chartData.length;
    const colCount = chartData[0].length;

    const range = chartSheet.getRangeByIndexes(0, 0, rowCount, colCount);

    range.values = chartData;

    await context.sync();

    // =============================
    // CHART TYPE LOGIC
    // =============================

    let chartType = "ColumnClustered";

    if (config.chart_type?.toLowerCase().includes("pie")) chartType = "Pie";
    else if (config.chart_type?.toLowerCase().includes("doughnut")) chartType = "Doughnut";
    else if (config.chart_type?.toLowerCase().includes("scatter")) chartType = "XYScatter";
    else if (config.chart_type?.toLowerCase().includes("waterfall")) chartType = "Waterfall";
    else if (config.chart_type?.toLowerCase().includes("funnel")) chartType = "Funnel";
    else if (config.chart_type?.toLowerCase().includes("stacked")) chartType = "ColumnStacked";
    else if (config.chart_type?.toLowerCase().includes("bar")) chartType = "BarClustered";
    else if (config.chart_type?.toLowerCase().includes("line")) chartType = "Line";
    else if (config.chart_type?.toLowerCase().includes("area")) chartType = "Area";

    const chart = sheet.charts.add(chartType, range, "Auto");

    chart.title.text = "AI Generated Chart";
    chart.height = 350;
    chart.width = 450;

    const dataLabels = chart.dataLabels;

    if (chartType === "Pie" || chartType === "Doughnut") {
      dataLabels.showCategoryName = true;
      dataLabels.showLegendKey = false;
      dataLabels.showPercentage = config.is_percentage || false;
      dataLabels.showValue = true;
    } else {
      dataLabels.showValue = true;
    }

    await context.sync();
  });

  // ✅ When percentage mode: show breakdown in chat with raw value + % share side by side
  // Use config.data directly — rows/valueCol were scoped inside Excel.run
  if (config.is_percentage) {
    const breakdownRows = config.data || [];
    const breakdownValueCol = (config.value_columns && config.value_columns[0]) || config.value_column;
    const breakdownCatCol = config.category_column;

    if (breakdownRows.length > 0 && breakdownValueCol && breakdownCatCol) {
      const total = breakdownRows.reduce((sum, r) => sum + (Number(r[breakdownValueCol]) || 0), 0);
      if (total > 0) {
        const lines = breakdownRows.map((r) => {
          const val = Number(r[breakdownValueCol]) || 0;
          const pct = ((val / total) * 100).toFixed(2);
          return `• ${r[breakdownCatCol]}: ${val.toLocaleString()} (${pct}%)`;
        });
        addMessage(`📊 Breakdown:\n${lines.join("\n")}`, "ai");
      }
    }
  }
}

/* =====================================================
   POWER BI EXPORT
===================================================== */

async function exportToPowerBI() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getUsedRange();
    range.load("values");
    await context.sync();

    const csv = range.values.map((r) => r.join(",")).join("\n");

    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "webenoid_export.csv";
    a.click();

    URL.revokeObjectURL(url);
  });

  addMessage("✅ CSV file downloaded successfully!", "ai");
  addMessage(
    "📋 How to open in Power BI:\n\n" +
    "1. Open Power BI Desktop\n" +
    "2. Click 'Get Data' → 'Text/CSV'\n" +
    "3. Select the downloaded 'webenoid_export.csv' file\n" +
    "4. Click 'Load' to import\n" +
    "5. Build your dashboards & reports!\n\n" +
    "💡 Tip: You can also drag the CSV file directly into Power BI.",
    "ai"
  );
}

async function buildDashboard() {
  addMessage("🚀 Building AI Dashboard...", "ai");

  const sheets = await readAllSheetsData();

  const res = await fetch("https://willfully-rubricated-tianna.ngrok-free.dev/dashboard", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "ngrok-skip-browser-warning": "69420"
    },
    body: JSON.stringify({ question: "dashboard", data: sheets }),
  });

  const raw = await res.json();

  if (!raw.success) {
    addMessage("❌ Dashboard creation failed.", "ai");
    return;
  }

  const data = raw.dashboard;

  await Excel.run(async (context) => {
    const workbook = context.workbook;

    let dashSheet;
    try {
      dashSheet = workbook.worksheets.getItem("AI_Dashboard");
      dashSheet.delete();
      await context.sync();
    } catch (e) { }

    dashSheet = workbook.worksheets.add("AI_Dashboard");

    // ========================
    // DYNAMIC KPI CARDS
    // ========================

    const kpis = data.kpis || [];
    const colLetters = ["A", "C", "E", "G"];

    kpis.forEach((kpi, i) => {
      if (i >= colLetters.length) return;
      const col = colLetters[i];
      dashSheet.getRange(`${col}1`).values = [[kpi.title]];
      dashSheet.getRange(`${col}2`).values = [[kpi.value]];
    });

    await context.sync();

    // ========================
    // PRIMARY CHART (Column/Bar)
    // ========================

    let dataStartRow = 5;

    if (data.primary_chart) {
      const chart = data.primary_chart;
      const catCol = chart.category_column;
      const valCol = chart.value_column;
      const rows = chart.data || [];

      const chartData = [[catCol, valCol]];
      rows.forEach((r) => chartData.push([r[catCol], r[valCol]]));

      const range = dashSheet.getRange("A5").getResizedRange(chartData.length - 1, 1);
      range.values = chartData;

      const excelChart = dashSheet.charts.add(chart.chart_type || "ColumnClustered", range, "Auto");
      excelChart.title.text = chart.title || "Primary Chart";
      excelChart.top = 120;
      excelChart.left = 10;
      excelChart.height = 300;
      excelChart.width = 420;

      dataStartRow = 5 + chartData.length + 2;
    }

    await context.sync();

    // ========================
    // SECONDARY CHART (Pie)
    // ========================

    if (data.secondary_chart) {
      const chart = data.secondary_chart;
      const catCol = chart.category_column;
      const valCol = chart.value_column;
      const rows = chart.data || [];

      const chartData = [[catCol, valCol]];
      rows.forEach((r) => chartData.push([r[catCol], r[valCol]]));

      const range = dashSheet.getRange("D5").getResizedRange(chartData.length - 1, 1);
      range.values = chartData;

      const excelChart = dashSheet.charts.add(chart.chart_type || "Pie", range, "Auto");
      excelChart.title.text = chart.title || "Distribution";
      excelChart.top = 120;
      excelChart.left = 450;
      excelChart.height = 300;
      excelChart.width = 350;
    }

    await context.sync();

    // ========================
    // TREND CHART (Line)
    // ========================

    if (data.trend_chart) {
      const chart = data.trend_chart;
      const catCol = chart.category_column;
      const valCol = chart.value_column;
      const rows = chart.data || [];

      const chartData = [[catCol, valCol]];
      rows.forEach((r) => chartData.push([r[catCol], r[valCol]]));

      const startCell = `A${dataStartRow}`;
      const range = dashSheet.getRange(startCell).getResizedRange(chartData.length - 1, 1);
      range.values = chartData;

      const excelChart = dashSheet.charts.add("Line", range, "Auto");
      excelChart.title.text = chart.title || "Trend Over Time";
      excelChart.top = 460;
      excelChart.left = 10;
      excelChart.height = 280;
      excelChart.width = 800;
    }

    await context.sync();
  });

  // Show KPI summary in the chat panel
  const kpis = data.kpis || [];
  if (kpis.length > 0) {
    const summaryLines = kpis.map((k) => `• ${k.title}: ${Number(k.value).toLocaleString()}`).join("\n");
    addMessage(`📊 Dashboard Summary:\n${summaryLines}`, "ai");
  }

  addMessage("✅ AI Dashboard created in the 'AI_Dashboard' sheet!", "ai");
}

