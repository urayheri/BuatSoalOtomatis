import { shuffleArray } from '../utils.js';

let originalPGQuestions = [];
let originalEssayQuestions = [];

export function initShuffler() {
    const excelInput = document.getElementById('excelInput');
    const fileStatus = document.getElementById('fileStatus');
    const processAcakBtn = document.getElementById('processAcakBtn');
    const previewSection = document.getElementById('previewSection');
    const soalContainer = document.getElementById('soalContainer');
    const badgeTotalSoal = document.getElementById('badgeTotalSoal');
    const downloadEBtn = document.getElementById('downloadTemplateEBtn');
    const downloadDBtn = document.getElementById('downloadTemplateDBtn');

    if (!excelInput) return;

    // --- A. Download Template Handler ---
    if (downloadEBtn) {
        downloadEBtn.addEventListener('click', () => downloadExcelTemplate(5));
    }
    if (downloadDBtn) {
        downloadDBtn.addEventListener('click', () => downloadExcelTemplate(4));
    }

    // --- B. Handle Upload File Excel ---
    excelInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        fileStatus.textContent = `Berkas Terpilih: ${file.name}`;
        fileStatus.className = "text-xs text-emerald-400 font-semibold block mt-1";

        const reader = new FileReader();

        reader.onload = function (event) {
            try {
                const data = new Uint8Array(event.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const ws = workbook.Sheets[firstSheetName];
                const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

                if (rows.length < 2) {
                    alert("Berkas Excel kosong atau tidak sesuai format.");
                    return;
                }

                originalPGQuestions = [];
                originalEssayQuestions = [];

                // Parsing baris Excel (mulai dari baris ke-2)
                for (let i = 1; i < rows.length; i++) {
                    const row = rows[i];
                    if (!row || !row[0]) continue;

                    const hasOptions = row[1] || row[2];
                    if (hasOptions) {
                        originalPGQuestions.push({
                            soal: row[0],
                            opsi: [
                                { teks: row[1] || "-", labelOriginal: "A" },
                                { teks: row[2] || "-", labelOriginal: "B" },
                                { teks: row[3] || "-", labelOriginal: "C" },
                                { teks: row[4] || "-", labelOriginal: "D" },
                                ...(row[5] ? [{ teks: row[5], labelOriginal: "E" }] : [])
                            ],
                            kunciOriginal: (row[6] || "A").toString().trim().toUpperCase()
                        });
                    } else {
                        originalEssayQuestions.push({
                            soal: row[0],
                            kunci: row[6] || "-"
                        });
                    }
                }

                // Render Preview
                previewSection.classList.remove('hidden');
                badgeTotalSoal.textContent = `${originalPGQuestions.length} PG / ${originalEssayQuestions.length} Essay`;
                
                soalContainer.innerHTML = `
                    <div class="p-3 bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 rounded-xl font-semibold">
                        ✓ Berhasil memuat ${originalPGQuestions.length} Soal PG dan ${originalEssayQuestions.length} Soal Uraian.
                    </div>
                `;

                processAcakBtn.disabled = false;

            } catch (err) {
                alert("Gagal membaca berkas Excel: " + err.message);
            }
        };

        reader.readAsArrayBuffer(file);
    });

    // --- C. Handle Proses Acak & Ekspor ZIP ---
    processAcakBtn.addEventListener('click', async () => {
        if (originalPGQuestions.length === 0 && originalEssayQuestions.length === 0) {
            alert("Belum ada data soal yang dimuat!");
            return;
        }

        const packetCount = parseInt(document.getElementById('packetCount').value) || 2;
        const zip = new JSZip();

        processAcakBtn.disabled = true;
        processAcakBtn.textContent = "⏳ Sedang Mengacak & Membuat Dokumen...";

        try {
            for (let p = 0; p < packetCount; p++) {
                const packetLetter = String.fromCharCode(65 + p);
                const shuffledPG = shuffleArray(originalPGQuestions);
                const shuffledEssay = shuffleArray(originalEssayQuestions);

                let docText = `=========================================\n`;
                docText += `    SOAL UJIAN PAKET ${packetLetter}\n`;
                docText += `=========================================\n\n`;

                docText += `--- A. SOAL PILIHAN GANDA ---\n\n`;
                shuffledPG.forEach((q, idx) => {
                    const shuffledOpsi = shuffleArray(q.opsi);
                    docText += `${idx + 1}. ${q.soal}\n`;
                    shuffledOpsi.forEach((opt, oIdx) => {
                        docText += `   ${String.fromCharCode(65 + oIdx)}. ${opt.teks}\n`;
                    });
                    docText += `\n`;
                });

                if (shuffledEssay.length > 0) {
                    docText += `\n--- B. SOAL URAIAN ---\n\n`;
                    shuffledEssay.forEach((q, idx) => {
                        docText += `${idx + 1}. ${q.soal}\n\n`;
                    });
                }

                zip.file(`Soal_Ujian_Paket_${packetLetter}.txt`, docText);
            }

            const content = await zip.generateAsync({ type: "blob" });
            const link = document.createElement("a");
            link.href = URL.createObjectURL(content);
            link.download = `Paket_Soal_Acak_Ujian.zip`;
            link.click();

        } catch (err) {
            alert("Terjadi kesalahan: " + err.message);
        } finally {
            processAcakBtn.disabled = false;
            processAcakBtn.textContent = "📝 Acak & Ekspor ke ZIP";
        }
    });
}

// Helper Generator Template Excel
function downloadExcelTemplate(optionCount) {
    const headers = ["Pertanyaan / Soal", "Opsi A", "Opsi B", "Opsi C", "Opsi D"];
    if (optionCount === 5) headers.push("Opsi E");
    headers.push("Kunci Jawaban");

    const sampleData = [
        headers,
        ["Siapakah penemu komputer?", "Charles Babbage", "Alan Turing", "Albert Einstein", "Nikola Tesla", "Thomas Edison", "A"].slice(0, optionCount + 2),
        ["Jelaskan fungsi dari RAM pada perangkat komputer!", "", "", "", "", "", "Kunci / Rubrik Essay"].slice(0, optionCount + 2)
    ];

    const ws = XLSX.utils.aoa_to_sheet(sampleData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "MasterSoal");
    XLSX.writeFile(wb, `Template_Master_Soal_${optionCount === 5 ? 'SMA_SMK' : 'SD_SMP'}.xlsx`);
}

// Auto Run
document.addEventListener('DOMContentLoaded', () => {
    initShuffler();
});