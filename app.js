// Aging Hutang - WEB
// STEP 1–4
// Read + Group + Normalize (NO POSITIVE) + Export Excel + Export Log

document.getElementById("btnProcess").addEventListener("click", () => {
    const fileInput = document.getElementById("fileInput");
    const status = document.getElementById("status");

    status.textContent = "";

    if (!fileInput.files || fileInput.files.length === 0) {
        status.textContent = "Status: belum ada file dipilih.";
        return;
    }

    const file = fileInput.files[0];
    const baseName = file.name.replace(/\.xlsx$/i, "");

    let msg = "";
    msg += "Aging Hutang (Web)\n";
    msg += file.name + "\n\n";

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
                throw new Error("Sheet kosong / tidak ada data.");
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
                    amount,
                    originalRow: r
                };

                groups[key].rows.push(rowObj);

                if (docType === "RE") {
                    groups[key].reRows.push(rowObj);
                } else {
                    groups[key].otherRows.push(rowObj);
                }
            });

            // =====================
            // STEP 3 — NORMALISASI
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

                g.finalTotal = total;

                g.reRows.forEach(r => r.newAmount = total);
                g.otherRows.forEach(r => r.newAmount = 0);
            }

            // =====================
            // STEP 4 — WRITE RESULT
            // =====================
            const outputRows = [header];

            dataRows.forEach((r, idx) => {
                const vendor = String(r[colIndex["Account"]] || "").trim();
                const assignment = String(r[colIndex["Assignment"]] || "").trim();

                if (!vendor || !assignment) {
                    outputRows.push(r);
                    return;
                }

                const key = vendor + "||" + assignment;
                const g = groups[key];

                if (!g) {
                    outputRows.push(r);
                    return;
                }

                const rowObj = g.rows.find(x => x.excelRow === idx + 2);

                if (rowObj && typeof rowObj.newAmount === "number") {
                    r[colIndex["Amount in local currency"]] = rowObj.newAmount;
                }

                outputRows.push(r);
            });

            // =====================
            // EXPORT EXCEL
            // =====================
            const newWb = XLSX.utils.book_new();
            const newWs = XLSX.utils.aoa_to_sheet(outputRows);
            XLSX.utils.book_append_sheet(newWb, newWs, "AGING");

            XLSX.writeFile(newWb, baseName + "_AGING.xlsx");

            // =====================
            // EXPORT LOG
            // =====================
            let log = "";
            log += "AGING HUTANG LOG\n";
            log += file.name + "\n\n";
            log += "Total Group: " + Object.keys(groups).length + "\n";
            log += "Group tanpa RE: " + errorNoRE.length + "\n";
            log += "Group total positif: " + positiveTotal.length + "\n\n";

            if (errorNoRE.length > 0) {
                log += "=== GROUP TANPA RE ===\n";
                errorNoRE.forEach(e => log += e + "\n");
                log += "\n";
            }

            if (positiveTotal.length > 0) {
                log += "=== TOTAL MASIH POSITIF (ERROR) ===\n";
                positiveTotal.forEach(e => log += e + "\n");
                log += "\n";
            }

            const logBlob = new Blob([log], { type: "text/plain;charset=utf-8" });
            const logLink = document.createElement("a");
            logLink.href = URL.createObjectURL(logBlob);
            logLink.download = baseName + "_AGING_LOG.txt";
            logLink.click();

            // =====================
            // FINAL STATUS
            // =====================
            msg += "STEP 4 – EXPORT\n";
            msg += "Excel hasil : " + baseName + "_AGING.xlsx\n";
            msg += "Log file    : " + baseName + "_AGING_LOG.txt\n\n";

            if (positiveTotal.length > 0) {
                msg += "❌ ERROR: masih ada total positif. CEK LOG.\n";
            } else {
                msg += "✅ SELESAI: tidak ada nilai positif.\n";
            }

            status.textContent = msg;

        } catch (err) {
            status.textContent = "ERROR:\n" + err.message;
        }
    };

    reader.readAsArrayBuffer(file);
});
