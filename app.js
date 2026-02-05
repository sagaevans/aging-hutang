// STEP 6.3 - HEADER VALIDATION (SAFE MODE)

document.getElementById("btnProcess").addEventListener("click", () => {
    const fileInput = document.getElementById("fileInput");
    const status = document.getElementById("status");

    if (!fileInput.files || fileInput.files.length === 0) {
        status.textContent = "ERROR: belum ada file dipilih.";
        return;
    }

    const file = fileInput.files[0];
    status.textContent = "Membaca file: " + file.name;

    const reader = new FileReader();

    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });

            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            const rows = XLSX.utils.sheet_to_json(sheet, {
                header: 1,
                raw: true,
                defval: ""
            });

            if (rows.length < 2) {
                status.textContent = "ERROR: file tidak memiliki data.";
                return;
            }

            const headerRow = rows[0].map(h => h.toString().trim());

            // === KOLOM WAJIB ===
            const requiredColumns = [
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

            let missing = [];
            let colIndex = {};

            requiredColumns.forEach(col => {
                const idx = headerRow.findIndex(h => h.toLowerCase() === col.toLowerCase());
                if (idx === -1) {
                    missing.push(col);
                } else {
                    colIndex[col] = idx;
                }
            });

            if (missing.length > 0) {
                status.textContent =
                    "ERROR: Kolom wajib tidak ditemukan:\n" +
                    missing.join("\n");
                return;
            }

            let msg = "";
            msg += "HEADER VALID ✔\n";
            msg += "Sheet: " + sheetName + "\n";
            msg += "Total baris data: " + (rows.length - 1) + "\n\n";
            msg += "Mapping kolom:\n";

            Object.keys(colIndex).forEach(k => {
                msg += `- ${k} → kolom index ${colIndex[k]}\n`;
            });

            status.textContent = msg;

            // ⚠️ STOP DI SINI (BELUM ADA HITUNG)

        } catch (err) {
            status.textContent = "ERROR runtime:\n" + err.message;
        }
    };

    reader.onerror = function () {
        status.textContent = "ERROR: gagal membaca file.";
    };

    reader.readAsArrayBuffer(file);
});
