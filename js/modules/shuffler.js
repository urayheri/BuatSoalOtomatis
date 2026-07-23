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

                // Parsing baris Excel (mulai dari baris ke-2 / index 1)
                for (let i = 1; i < rows.length; i++) {
                    const row = rows[i];
                    if (!row || !row[0]) continue;

                    const hasOptions = row[1] || row[2];
                    if (hasOptions) {
                        originalPGQuestions.push({
                            soal: row[0],
                            gambar: row[7] || "", // Kolom H (Index 7) untuk URL/Base64 Gambar
                            opsi: [
                                { teks: String(row[1] || "-"), labelOriginal: "A" },
                                { teks: String(row[2] || "-"), labelOriginal: "B" },
                                { teks: String(row[3] || "-"), labelOriginal: "C" },
                                { teks: String(row[4] || "-"), labelOriginal: "D" },
                                ...(row[5] ? [{ teks: String(row[5]), labelOriginal: "E" }] : [])
                            ],
                            kunciOriginal: (row[6] || "A").toString().trim().toUpperCase()
                        });
                    } else {
                        originalEssayQuestions.push({
                            soal: row[0],
                            gambar: row[7] || "",
                            kunci: row[6] || "-"
                        });
                    }
                }

                // Render Preview Paket Utama
                previewSection.classList.remove('hidden');
                badgeTotalSoal.textContent = `${originalPGQuestions.length} PG / ${originalEssayQuestions.length} Essay`;
                
                renderPreviewSoal(originalPGQuestions, originalEssayQuestions, soalContainer);

                processAcakBtn.disabled = false;

            } catch (err) {
                alert("Gagal membaca berkas Excel: " + err.message);
            }
        };

        reader.readAsArrayBuffer(file);
    });

    // --- C. Handle Proses Acak & Ekspor Dokumen Rapi (Word / ZIP) ---
    processAcakBtn.addEventListener('click', async () => {
        if (originalPGQuestions.length === 0 && originalEssayQuestions.length === 0) {
            alert("Belum ada data soal yang dimuat!");
            return;
        }

        const packetCount = parseInt(document.getElementById('packetCount')?.value) || 2;
        const zip = new JSZip();

        processAcakBtn.disabled = true;
        processAcakBtn.textContent = "⏳ Sedang Mengacak & Membuat Dokumen Word...";

        try {
            for (let p = 0; p < packetCount; p++) {
                const packetLetter = String.fromCharCode(65 + p);
                const shuffledPG = shuffleArray(originalPGQuestions);
                const shuffledEssay = shuffleArray(originalEssayQuestions);

                // Buat Dokumen HTML Lengkap untuk Paket
                const docHTML = generateFullExamHTML(shuffledPG, shuffledEssay, packetLetter);

                // Konversi ke format .docx (Microsoft Word)
                const wordHeader = "<html xmlns:o='urn:schemas-microsoft-com:office:office' " +
                                   "xmlns:w='urn:schemas-microsoft-com:office:word' " +
                                   "xmlns='http://www.w3.org/TR/REC-html40'>" +
                                   "<head><meta charset='utf-8'><title>Soal Ujian</title></head><body>";
                const wordFooter = "</body></html>";
                const fullWordHTML = wordHeader + docHTML + wordFooter;

                // Masukkan ke ZIP
                zip.file(`Soal_Ujian_Paket_${packetLetter}.docx`, fullWordHTML);
            }

            const content = await zip.generateAsync({ type: "blob" });
            const link = document.createElement("a");
            link.href = URL.createObjectURL(content);
            link.download = `Paket_Soal_Acak_Ujian_Rapi.zip`;
            link.click();

        } catch (err) {
            alert("Terjadi kesalahan pembuatan dokumen: " + err.message);
        } finally {
            processAcakBtn.disabled = false;
            processAcakBtn.textContent = "📝 Acak & Ekspor ke Word (ZIP)";
        }
    });
}

// --- D. LOGIKA PENATAAN TATA LETAK & MATEMATIKA ---

// Helper menentukan layout opsi (1 Baris, 2 Baris, atau Vertikal)
function getOpsiLayoutHtml(opsiArray) {
    const maxLen = Math.max(...opsiArray.map(o => (o.teks || '').length));
    const labels = ["a", "b", "c", "d", "e"];

    if (maxLen <= 15) {
        // Layout 1 Baris (Pendek)
        let cols = opsiArray.map((opt, idx) => `<div style="display:table-cell; width:20%; vertical-align:top;">${labels[idx]}. ${opt.teks}</div>`).join('');
        return `<div style="display:table; width:100%; table-layout:fixed;">${cols}</div>`;
    } else if (maxLen <= 35) {
        // Layout 2 Baris (Sedang: Baris 1 -> a,c,e | Baris 2 -> b,d)
        let row1 = '', row2 = '';
        opsiArray.forEach((opt, idx) => {
            const cell = `<div style="display:table-cell; ${idx < 3 ? 'width:33.33%' : 'width:50%'}; vertical-align:top;">${labels[idx]}. ${opt.teks}</div>`;
            if (idx % 2 === 0) row1 += cell;
            else row2 += cell;
        });
        return `
            <div style="display:table; width:100%; margin-bottom:2px; table-layout:fixed;">${row1}</div>
            <div style="display:table; width:100%; table-layout:fixed;">${row2}</div>
        `;
    } else {
        // Layout Vertikal (Panjang)
        return opsiArray.map((opt, idx) => 
            `<div style="display:table; width:100%; margin-bottom:2px;">
                <div style="display:table-cell; width:100%; vertical-align:top;">${labels[idx]}. ${opt.teks}</div>
             </div>`
        ).join('');
    }
}

// Render Preview Soal Ke Layar HTML dengan KaTeX Matematika
function renderPreviewSoal(pgList, essayList, containerElement) {
    const fullHTML = generateFullExamHTML(pgList, essayList, "A (Preview)");
    containerElement.innerHTML = fullHTML;

    // Render Matematika KaTeX jika pustaka tersedia
    if (window.renderMathInElement) {
        renderMathInElement(containerElement, {
            delimiters: [
                {left: '$$', right: '$$', display: true},
                {left: '$', right: '$', display: false},
                {left: '\\(', right: '\\)', display: false},
                {left: '\\[', right: '\\]', display: true}
            ],
            throwOnError: false
        });
    }
}

// Generasi HTML Soal Lengkap + Kop Surat Rapi
function generateFullExamHTML(pgList, essayList, packetName) {
    let html = `
    <div style="font-family:'Times New Roman', Times, serif; font-size:11pt; line-height:1.25; color:#000;">
        
        <!-- KOP SURAT -->
        <div style="display:table; width:100%; border-bottom:3px double #000; padding-bottom:5px; margin-bottom:12px;">
            <div style="display:table-cell; width:15%; vertical-align:middle; text-align:center;">
                <!-- Logo Kiri Placeholder -->
                <div style="font-weight:bold; font-size:10pt;">[LOGO]</div>
            </div>
            <div style="display:table-cell; width:70%; vertical-align:middle; text-align:center;">
                <div style="font-size:11pt; font-weight:bold; text-transform:uppercase;">PEMERINTAH PROVINSI KALIMANTAN BARAT</div>
                <div style="font-size:12pt; font-weight:bold; text-transform:uppercase;">DINAS PENDIDIKAN DAN KEBUDAYAAN</div>
                <div style="font-size:14pt; font-weight:bold; text-transform:uppercase; margin:2px 0;">SMK NEGERI 1 SEMPARUK</div>
                <div style="font-size:8.5pt;">Jalan Pendidikan Nomor 19 Kecamatan Semparuk Kabupaten Sambas 79453</div>
                <div style="font-size:8.5pt; font-weight:bold;"><u>smkn1_semparuk@yahoo.co.id</u></div>
            </div>
            <div style="display:table-cell; width:15%; vertical-align:middle; text-align:center;">
                <div style="font-weight:bold; font-size:10pt;">[LOGO]</div>
            </div>
        </div>

        <!-- JUDUL UJIAN -->
        <div style="text-align:center; font-weight:bold; font-size:12pt; text-transform:uppercase; margin-bottom:12px;">
            <u>SOAL UJIAN SEMESTER (PAKET ${packetName})</u>
        </div>

        <div style="font-weight:bold; margin-bottom:8px;">Pilihlah jawaban berikut dengan benar!</div>

        <!-- LIST SOAL PILIHAN GANDA -->
        <div class="soal-container">
    `;

    // Render PG
    pgList.forEach((q, idx) => {
        const shuffledOpsi = shuffleArray(q.opsi);
        const opsiLayout = getOpsiLayoutHtml(shuffledOpsi);
        const gambarHtml = q.gambar ? `<div style="margin:5px 0 5px 28px;"><img src="${q.gambar}" style="max-width:220px; max-height:150px; border:1px solid #ccc;"/></div>` : '';

        html += `
        <div style="margin-bottom:12px; page-break-inside:avoid;">
            <div style="display:table; width:100%; margin-bottom:4px;">
                <div style="display:table-cell; width:28px; vertical-align:top;">${idx + 1}.</div>
                <div style="display:table-cell; vertical-align:top;">${q.soal}</div>
            </div>
            ${gambarHtml}
            <div style="margin-left:28px; width:calc(100% - 28px);">
                ${opsiLayout}
            </div>
        </div>`;
    });

    // Render Essay
    if (essayList.length > 0) {
        html += `<div style="font-weight:bold; margin:16px 0 8px 0;">Jawablah pertanyaan berikut dengan singkat dan tepat!</div>`;
        essayList.forEach((q, idx) => {
            const gambarHtml = q.gambar ? `<div style="margin:5px 0 5px 28px;"><img src="${q.gambar}" style="max-width:220px; max-height:150px; border:1px solid #ccc;"/></div>` : '';
            html += `
            <div style="margin-bottom:12px; page-break-inside:avoid;">
                <div style="display:table; width:100%; margin-bottom:4px;">
                    <div style="display:table-cell; width:28px; vertical-align:top;">${idx + 1}.</div>
                    <div style="display:table-cell; vertical-align:top;">${q.soal}</div>
                </div>
                ${gambarHtml}
            </div>`;
        });
    }

    html += `</div></div>`;
    return html;
}

// Generator Template Excel Master
function downloadExcelTemplate(optionCount) {
    const headers = ["Pertanyaan / Soal", "Opsi A", "Opsi B", "Opsi C", "Opsi D"];
    if (optionCount === 5) headers.push("Opsi E");
    headers.push("Kunci Jawaban", "URL / Base64 Gambar (Opsional)");

    const sampleData = [
        headers,
        ["Berapakah hasil dari persamaan $x^2 - 5x + 6 = 0$?", "$x=2$ atau $x=3$", "$x=-2$ atau $x=3$", "$x=1$ atau $x=6$", "$x=-1$", "$x=0$", "A", ""].slice(0, optionCount + 3),
        ["Jelaskan fungsi dari RAM pada komputer!", "", "", "", "", "", "Kunci / Rubrik Essay", ""].slice(0, optionCount + 3)
    ];

    const ws = XLSX.utils.aoa_to_sheet(sampleData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "MasterSoal");
    XLSX.writeFile(wb, `Template_Master_Soal_${optionCount === 5 ? 'SMA_SMK' : 'SD_SMP'}.xlsx`);
}

// Auto Init saat DOM siap
document.addEventListener('DOMContentLoaded', () => {
    initShuffler();
});
