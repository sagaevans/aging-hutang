// Aging Hutang - WEB
// STEP 1 + STEP 2 + STEP 3
// READ EXCEL + HEADER VALIDATION + GROUPING + NORMALISASI (NO POSITIVE)

document.getElementById("btnProcess").addEventListener("click", () => {
    const fileInput = document.getElementById("fileInput");
    const status = document.getElementById("status");

    status.textContent = "";

    if (!fileInput.files || fileInput.files.length === 0) {
        status.textContent = "Status: belum ada file dipilih.";
        return;
    }

    const file = fileInput.files[0];
    let msg = "";
    msg += "Aging Hutang (Web – Step 1)\n";
    msg += file.name + "\n";

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
            const dataRows = rows.slice(1);

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
                    excelRow: idx + 2,
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

            msg += "\nSTEP 2 – GROUPING RESULT\n";
            msg += "Total group (Vendor + Assignment): " + Object.keys(groups).length + "\n";

            // =====================
            // STEP 3 — NORMALISASI HUTANG
            // =====================
            let errorNoRE = [];
            let positiveTotal = [];

            for (const key in groups) {
                const g = groups[key];

                const total = g.rows.reduce((s, r) => s + r.amount, 0);

                if (g.reRows.length === 0) {
                    errorNoRE.push(key);
                    continue;
                }

                if (total > 0) {
                    positiveTotal.push(`${key} = ${total}`);
                }

                // simpan hasil (belum ditulis ke Excel)
                g.finalTotal = total;

                g.reRows.forEach(r => r.newAmount = total);
                g.otherRows.forEach(r => r.newAmount = 0);
            }

            msg += "\nSTEP 3 – NORMALISASI HUTANG\n";
            msg += "Group diproses : " + Object.keys(groups).length + "\n";
            msg += "Group tanpa RE : " + errorNoRE.length + "\n";
            msg += "Total positif  : " + positiveTotal.length + "\n";

            if (errorNoRE.length > 0) {
                msg += "\n❌ GROUP TANPA RE:\n";
                errorNoRE.slice(0, 5).forEach(e => msg += "- " + e + "\n");
            }

            if (positiveTotal.length > 0) {
                msg += "\n❌ TOTAL MASIH POSITIF (INVALID):\n";
                positiveTotal.slice(0, 5).forEach(e => msg += "- " + e + "\n");
            }

            status.textContent = msg;

            // STOP DI SINI (BELUM EXPORT)

        } catch (err) {
            status.textContent = "ERROR:\n" + err.message;
        }
    };

    reader.onerror = function () {
        status.textContent = "ERROR: gagal membaca file.";
    };

    reader.readAsArrayBuffer(file);
});
