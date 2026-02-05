// Aging Hutang - WEB
// STEP 1 + STEP 2
// READ EXCEL + HEADER VALIDATION + GROUPING
// SAFE VERSION (NO DATA MODIFICATION)

document.getElementById("btnProcess").addEventListener("click", () => {
    const fileInput = document.getElementById("fileInput");
    const status = document.getElementById("status");

    status.textContent = "";

    if (!fileInput.files || fileInput.files.length === 0) {
        status.textContent = "Status: belum ada file dipilih.";
        return;
    }

    const file = fileInput.files[0];
    status.textContent = "Membaca file Excel...\nFile: " + file.name;

    const reader = new FileReader();

    reader.onload = function (e) {
        try {
            // =====================
            // STEP 1 — READ EXCEL
            // =====================
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });

            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            const rows = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                raw: true,
                defval: ""
            });

            if (rows.length < 2) {
                status.textContent = "ERROR: Sheet kosong / tidak ada data.";
                return;
            }

            const header = rows[0];
            const rowCount = rows.length - 1;

            // =====================
            // HEADER VALIDATION
            // =====================
            const requiredHeaders = [
                "G/L Account",
                "Account",
                "Assignment",
                "Document Number",
                "Document Type",
                "Posting Date",
                "Amount in local currency",
                "Local Currency",
                "Text"
            ];

            const colIndex = {};
            requiredHeaders.forEach(h => {
                const idx = header.indexOf(h);
                if (idx === -1) {
                    throw new Error(`Header wajib tidak ditemukan: ${h}`);
                }
                colIndex[h] = idx;
            });

            let msg = "";
            msg += "Aging Hutang (Web – Step 1)\n";
            msg += file.name + "\n";
            msg += "HEADER VALID ✔\n";
            msg += "Sheet: " + sheetName + "\n";
            msg += "Total baris data: " + rowCount + "\n\n";
            msg += "Mapping kolom:\n";

            Object.keys(colIndex).forEach(k => {
                msg += `- ${k} → kolom index ${colIndex[k]}\n`;
            });

            // =====================
            // STEP 2 — GROUPING
            // =====================
            const groups = {};
            const dataRows = rows.slice(1); // tanpa header

            dataRows.forEach((r, idx) => {
                const vendor = String(r[colIndex["Account"]] || "").trim();
                const assignment = String(r[colIndex["Assignment"]] || "").trim();
                const docType = String(r[colIndex["Document Type"]] || "").trim();
                const amount = Number(r[colIndex["Amount in local currency"]]) || 0;

                if (!vendor || !assignment) return;

                const key = vendor + "||" + assignment;

                if (!groups[key]) {
                    groups[key] = {
                        key,
                        vendor,
                        assignment,
                        rows: [],
                        reRows: [],
                        otherRows: []
                    };
                }

                const rowObj = {
                    excelRow: idx + 2, // baris Excel asli
                    docType,
                    amount
                };

                groups[key].rows.push(rowObj);

                if (docType === "RE") {
                    groups[key].reRows.push(rowObj);
                } else {
                    groups[key].otherRows.push(rowObj);
                }
            });

            // =====================
            // GROUPING SUMMARY
            // =====================
            const totalGroups = Object.keys(groups).length;

            msg += "\nSTEP 2 – GROUPING RESULT\n";
            msg += "Total group (Vendor + Assignment): " + totalGroups + "\n";

            let shown = 0;
            for (const k in groups) {
                if (shown >= 3) break;
                const g = groups[k];
                msg += `\n[${g.key}]\n`;
                msg += `  Total Rows : ${g.rows.length}\n`;
                msg += `  RE Rows    : ${g.reRows.length}\n`;
                msg += `  Other Rows : ${g.otherRows.length}\n`;
                shown++;
            }

            status.textContent = msg;

        } catch (err) {
            status.textContent = "ERROR:\n" + err.message;
        }
    };

    reader.onerror = function () {
        status.textContent = "ERROR: gagal membaca file.";
    };

    reader.readAsArrayBuffer(file);
});
