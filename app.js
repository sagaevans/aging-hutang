let finalRows = [];
let logLines = [];
let workbookName = "aging_hutang";

function processFile() {
  const fileInput = document.getElementById("fileInput");
  const output = document.getElementById("output");
  logLines = [];
  finalRows = [];

  if (!fileInput.files.length) {
    output.textContent = "ERROR: File belum dipilih";
    return;
  }

  const file = fileInput.files[0];
  workbookName = file.name.replace(".xlsx", "");

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const wb = XLSX.read(data, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const header = rows[0];
    const dataRows = rows.slice(1);

    output.textContent =
      `HEADER VALID ✔\n` +
      `Sheet: ${wb.SheetNames[0]}\n` +
      `Total baris data: ${dataRows.length}\n\n`;

    const IDX = {
      gl: 0,
      vendor: 1,
      assignment: 2,
      docNo: 3,
      docType: 4,
      date: 5,
      amount: 6,
      currency: 7,
      text: 8
    };

    // ===== GROUPING =====
    const groups = {};
    dataRows.forEach((r, i) => {
      const key = `${r[IDX.vendor]}||${r[IDX.assignment]}`;
      if (!groups[key]) groups[key] = [];
      groups[key].push({ row: r, idx: i + 2 });
    });

    output.textContent +=
      `STEP 2 – GROUPING RESULT\n` +
      `Total group (Vendor + Assignment): ${Object.keys(groups).length}\n\n`;

    let hasPositive = false;

    Object.entries(groups).forEach(([key, items]) => {
      let sum = 0;
      let reRow = null;

      items.forEach(o => {
        const val = Number(o.row[IDX.amount]) || 0;
        sum += val;
        if (o.row[IDX.docType] === "RE" && !reRow) {
          reRow = o.row;
        }
      });

      if (!reRow) {
        reRow = items[0].row;
        logLines.push(`[WARN] RE tidak ditemukan → pakai baris pertama (${key})`);
      }

      // ===== NORMALISASI =====
      reRow[IDX.amount] = sum;

      items.forEach(o => {
        if (o.row !== reRow) {
          o.row[IDX.amount] = 0;
        }
      });

      if (sum > 0) {
        hasPositive = true;
        logLines.push(`[ERROR] POSITIF > 0 | ${key} | ${sum}`);
      }
    });

    finalRows = [header, ...dataRows];

    if (hasPositive) {
      output.innerHTML += `\n<span class="error">ERROR: Masih ada nilai POSITIF (>0). Export DIBLOKIR.</span>`;
      logLines.push("EXPORT DIBLOKIR: nilai positif masih ada");
      return;
    }

    output.innerHTML += `\n<span class="ok">OK: Tidak ada nilai positif. Siap export.</span>`;
    logLines.push("SUCCESS: Semua nilai ≤ 0");
  };

  reader.readAsArrayBuffer(file);
}

// ===== EXPORT EXCEL =====
function downloadExcel() {
  if (!finalRows.length) {
    alert("Belum ada hasil");
    return;
  }

  const ws = XLSX.utils.aoa_to_sheet(finalRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "AGING_RESULT");
  XLSX.writeFile(wb, workbookName + "_RESULT.xlsx");
}

// ===== EXPORT LOG =====
function downloadLog() {
  if (!logLines.length) {
    alert("Log kosong");
    return;
  }

  const blob = new Blob([logLines.join("\n")], { type: "text/plain" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = workbookName + "_LOG.txt";
  a.click();
}
