// STEP 6.2 - READ ONLY EXCEL (NO PROCESSING)

document.getElementById("btnProcess").addEventListener("click", () => {
    const fileInput = document.getElementById("fileInput");
    const status = document.getElementById("status");

    if (!fileInput.files || fileInput.files.length === 0) {
        status.textContent = "Status: belum ada file dipilih.";
        return;
    }

    const file = fileInput.files[0];
    status.textContent = "Status: membaca file Excel...\nFile: " + file.name;

    const reader = new FileReader();

    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });

            // Ambil sheet pertama
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            // Convert ke array (raw, tanpa parsing aneh)
            const rows = XLSX.utils.sheet_to_json(worksheet, {
                header: 1,
                raw: true,
                defval: ""
            });

            if (rows.length === 0) {
                status.textContent = "ERROR: Sheet kosong.";
                return;
            }

            const header = rows[0];
            const rowCount = rows.length - 1;

            let msg = "";
            msg += "Status: file berhasil dibaca ✔\n";
            msg += "Sheet: " + sheetName + "\n";
            msg += "Jumlah baris data: " + rowCount + "\n\n";
            msg += "Header terdeteksi:\n";

            header.forEach((h, i) => {
                msg += `  [${i}] ${h}\n`;
            });

            status.textContent = msg;

        } catch (err) {
            status.textContent = "ERROR saat membaca file:\n" + err.message;
        }
    };

    reader.onerror = function () {
        status.textContent = "ERROR: gagal membaca file.";
    };

    reader.readAsArrayBuffer(file);
});
