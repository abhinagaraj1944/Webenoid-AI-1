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

  setupMenu();
});

/* =========================
   MENU
========================= */
function setupMenu() {
  const menuBtn = document.getElementById("menuBtn");
  const menu = document.getElementById("menuDropdown");

  if (!menuBtn || !menu) return;

  menuBtn.onclick = e => {
    e.stopPropagation();
    menu.classList.toggle("hidden");
  };

  document.onclick = () => menu.classList.add("hidden");
}

/* =========================
   UI HELPERS
========================= */
function addMessage(text, type = "ai") {
  const box = document.getElementById("resultBox");
  const div = document.createElement("div");
  div.className = `message ${type}`;
  div.innerText = text;
  box.appendChild(div);
  box.scrollTop = box.scrollHeight;
}

function addResultCard(title, result) {
  const box = document.getElementById("resultBox");
  const card = document.createElement("div");
  card.className = "result-card";
  card.innerHTML = `<div class="result-title">${title}</div>`;

  if (Array.isArray(result)) {
    result.forEach((item, index) => {
      const row = document.createElement("div");
      row.className = "result-value";
      row.innerText = `${index + 1}. ${item}`;
      card.appendChild(row);
    });
  } else if (typeof result === "object" && result !== null) {
    Object.entries(result).forEach(([k, v]) => {
      const row = document.createElement("div");
      row.className = "result-value";
      row.innerText = `${k}: ${v}`;
      card.appendChild(row);
    });
  } else {
    const row = document.createElement("div");
    row.className = "result-value";
    row.innerText = result;
    card.appendChild(row);
  }

  box.appendChild(card);
  box.scrollTop = box.scrollHeight;
}

/* =========================
   EXPORT SELECTED DATA → POWER BI
========================= */
async function exportToPowerBI() {
  try {
    addMessage("📤 Exporting selected data for Power BI…", "ai");

    await Excel.run(async context => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      const values = range.values;

      if (!values || values.length < 2) {
        addMessage("⚠️ Please select a table with headers.", "ai");
        return;
      }

      const csv = values
        .map(row =>
          row.map(cell =>
            `"${String(cell ?? "").replace(/"/g, '""')}"`
          ).join(",")
        )
        .join("\n");

      const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);

      const a = document.createElement("a");
      a.href = url;
      a.download = "webenoid_powerbi_data.csv";
      a.click();

      URL.revokeObjectURL(url);
    });

    addMessage(
      "✅ Data exported.\nOpen Power BI Desktop → Get Data → Text/CSV",
      "ai"
    );

  } catch (err) {
    console.error(err);
    addMessage("❌ Failed to export data.", "ai");
  }
}

/* =========================
   READ ALL SHEETS
========================= */
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

/* =========================
   FULL SMART GRADING SYSTEM
========================= */
async function addFullGradingSystem() {
  try {
    await Excel.run(async context => {

      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRange();
      range.load("values");
      await context.sync();

      const values = range.values;
      if (!values || values.length < 2) {
        addMessage("⚠️ No data found.", "ai");
        return;
      }

      const headers = values[0];

      const subjectIndexes = [];

      for (let col = 0; col < headers.length; col++) {
        let isNumeric = true;
        for (let row = 1; row < values.length; row++) {
          if (values[row][col] !== null && values[row][col] !== "") {
            if (typeof values[row][col] !== "number") {
              isNumeric = false;
              break;
            }
          }
        }
        if (isNumeric) subjectIndexes.push(col);
      }

      if (subjectIndexes.length === 0) {
        addMessage("❌ No numeric subject columns detected.", "ai");
        return;
      }

      const totalCol = headers.length;
      const percentCol = headers.length + 1;
      const gradeCol = headers.length + 2;

      sheet.getCell(0, totalCol).values = [["Total Marks"]];
      sheet.getCell(0, percentCol).values = [["Percentage"]];
      sheet.getCell(0, gradeCol).values = [["Grade"]];

      const maxMarks = subjectIndexes.length * 100;

      for (let i = 1; i < values.length; i++) {

        const subjectCells = subjectIndexes.map(
          idx => sheet.getCell(i, idx).getAddress().replace(/\$/g, "")
        );

        sheet.getCell(i, totalCol).formulas = [
          [`=SUM(${subjectCells.join(",")})`]
        ];

        const totalCell = sheet.getCell(i, totalCol).getAddress().replace(/\$/g, "");

        sheet.getCell(i, percentCol).formulas = [
          [`=(${totalCell}/${maxMarks})*100`]
        ];

        const percentCell = sheet.getCell(i, percentCol).getAddress().replace(/\$/g, "");

        sheet.getCell(i, gradeCol).formulas = [[
          `=IF(${percentCell}>=90,"A+",
           IF(${percentCell}>=80,"A",
           IF(${percentCell}>=70,"B",
           IF(${percentCell}>=60,"C",
           IF(${percentCell}>=50,"D","F")))))`
        ]];
      }

      await context.sync();
    });

    addMessage("✅ Smart Grading System Added Successfully.", "ai");

  } catch (error) {
    console.error(error);
    addMessage("❌ Failed to apply grading system.", "ai");
  }
}

/* =========================
   MAIN AI FLOW
========================= */
async function runWebenoidAI() {
  try {
    const input = document.getElementById("prompt");
    const prompt = input.value.trim();
    if (!prompt) return;

    addMessage(prompt, "user");
    input.value = "";

    // 🔥 SMART GRADING TRIGGER
    if (prompt.toLowerCase().includes("grading")) {
      await addFullGradingSystem();
      return;
    }

    addMessage("🔍 Analyzing Excel data…", "ai");

    const sheets = await readAllSheetsData();

    const res = await fetch("http://localhost:8000/analyze", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ prompt, sheets })
    });

    if (!res.ok) {
      addMessage("❌ Server error occurred.", "ai");
      return;
    }

    const result = await res.json();

    if (result.type === "answer") {
      addResultCard(result.operation.toUpperCase(), result.result);
    } else {
      addMessage(result.message, "ai");
    }

  } catch (err) {
    console.error(err);
    addMessage("❌ Failed to analyze data.", "ai");
  }
}
