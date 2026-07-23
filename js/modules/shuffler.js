import { shuffleArray } from '../utils.js';

let originalPGQuestions = [];
let originalEssayQuestions = [];
let kopGambarBase64 = ""; // Menampung Base64 dari Upload Kop Surat

export function initShuffler() {
    const excelInput = document.getElementById('excelInput');
    const fileStatus = document.getElementById('fileStatus');
    const processAcakBtn = document.getElementById('processAcakBtn');
    const previewSection = document.getElementById('previewSection');
    const soalContainer = document.getElementById('soalContainer');
    const badgeTotalSoal = document.getElementById('badgeTotalSoal');
    const downloadEBtn = document.getElementById('downloadTemplateEBtn');
    const downloadDBtn = document.getElementById('downloadTemplateDBtn');
    const kopGambarInput = document.getElementById('kopGambarInput');

    if (!excelInput) return;

    // Handler Upload Gambar Kop Surat
    // Handler Upload Gambar Kop Surat
if (kopGambarInput) {
    kopGambarInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (evt) => { 
                kopGambarBase64 = evt.target.result;
                
                // 1. Cari elemen kotak preview (dengan fallback selector yang aman)
                const previewBox = document.getElementById('kopPreviewContainer') || 
                                   document.querySelector('#kopGambarInput').parentElement.nextElementSibling;
                
                if (previewBox) {
                    previewBox.innerHTML = `
                        <div class="space-y-1">
                            <img src="${kopGambarBase64}" class="max-h-12 max-w-full mx-auto rounded object-contain border border-emerald-500/30 shadow-sm" />
                            <span class="text-[10px] text-emerald-400 font-semibold block">✓ Kop Terpasang</span>
                        </div>
                    `;
                }

                // 2. Jika soal Excel sudah ada di layar, perbarui pratinjau lembar soalnya
                if (typeof updatePreview === 'function') {
                    updatePreview(); 
                }
            };
            reader.readAsDataURL(file);
        }
    });
}

    // Handler Download Template Excel
    if (downloadEBtn) downloadEBtn.addEventListener('click', () => downloadExcelTemplate(5));
    if (downloadDBtn) downloadDBtn.addEventListener('click', () => downloadExcelTemplate(4));

    // Handle Upload Excel
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
                const ws = workbook.Sheets[workbook.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

                if (rows.length < 2) {
                    alert("Berkas Excel kosong atau tidak sesuai format.");
                    return;
                }

                originalPGQuestions = [];
                originalEssayQuestions = [];

                for (let i = 1; i < rows.length; i++) {
                    const row = rows[i];
                    if (!row || !row[0]) continue;

                    const hasOptions = row[1] || row[2];
                    if (hasOptions) {
                        originalPGQuestions.push({
                            soal: String(row[0]),
                            opsi: [
                                { teks: String(row[1] || "-"), labelOriginal: "A" },
                                { teks: String(row[2] || "-"), labelOriginal: "B" },
                                { teks: String(row[3] || "-"), labelOriginal: "C" },
                                { teks: String(row[4] || "-"), labelOriginal: "D" },
                                ...(row[5] ? [{ teks: String(row[5]), labelOriginal: "E" }] : [])
                            ],
                            kunciOriginal: (row[6] || "A").toString().trim().toUpperCase(),
                            gambar: row[7] || ""
                        });
                    } else {
                        originalEssayQuestions.push({
                            soal: String(row[0]),
                            kunci: String(row[6] || "-"),
                            gambar: row[7] || ""
                        });
                    }
                }

                previewSection.classList.remove('hidden');
                badgeTotalSoal.textContent = `${originalPGQuestions.length} PG / ${originalEssayQuestions.length} Essay`;
                updatePreview();
                processAcakBtn.disabled = false;

            } catch (err) {
                alert("Gagal membaca berkas Excel: " + err.message);
            }
        };
        reader.readAsArrayBuffer(file);
    });

    function updatePreview() {
        if (originalPGQuestions.length > 0 || originalEssayQuestions.length > 0) {
            renderPreviewSoal(originalPGQuestions, originalEssayQuestions, soalContainer);
        }
    }

    // Handle Proses Acak & Buat File Word (.docx)
    processAcakBtn.addEventListener('click', async () => {
        if (originalPGQuestions.length === 0 && originalEssayQuestions.length === 0) {
            alert("Belum ada data soal yang dimuat!");
            return;
        }

        const packetCount = parseInt(document.getElementById('packetCount')?.value) || 2;
        const zip = new JSZip();

        processAcakBtn.disabled = true;
        processAcakBtn.textContent = "⏳ Memproses Dokumen Word...";

        try {
            for (let p = 0; p < packetCount; p++) {
                const packetLetter = String.fromCharCode(65 + p);
                const limitPG = parseInt(document.getElementById('limitCountPG')?.value) || originalPGQuestions.length;
                const limitEssay = parseInt(document.getElementById('limitCountEssay')?.value) || originalEssayQuestions.length;

                const shuffledPG = shuffleArray(originalPGQuestions).slice(0, limitPG);
                const shuffledEssay = shuffleArray(originalEssayQuestions).slice(0, limitEssay);

                const htmlString = generateFullExamHTML(shuffledPG, shuffledEssay, packetLetter);

                let convertedDocx;
                if (window.htmlDocx) {
                    convertedDocx = window.htmlDocx.asBlob(htmlString, {
                        orientation: 'portrait',
                        margins: { top: 720, right: 720, bottom: 720, left: 720 }
                    });
                } else {
                    convertedDocx = new Blob([htmlString], { type: 'application/msword' });
                }

                zip.file(`Soal_Ujian_Paket_${packetLetter}.docx`, convertedDocx);
            }

            const content = await zip.generateAsync({ type: "blob" });
            const link = document.createElement("a");
            link.href = URL.createObjectURL(content);
            link.download = `Paket_Soal_Ujian_Word.zip`;
            link.click();

        } catch (err) {
            alert("Terjadi kesalahan pembuatan dokumen Word: " + err.message);
        } finally {
            processAcakBtn.disabled = false;
            processAcakBtn.textContent = "📝 Acak & Ekspor ke ZIP";
        }
    });
}

function getOpsiLayoutHtml(opsiArray) {
    const maxLen = Math.max(...opsiArray.map(o => (o.teks || '').length));
    const labels = ["a", "b", "c", "d", "e"];

    if (maxLen <= 15) {
        let cols = opsiArray.map((opt, idx) => `<td style="width:20%; vertical-align:top;">${labels[idx]}. ${opt.teks}</td>`).join('');
        return `<table style="width:100%; border-collapse:collapse;"><tr>${cols}</tr></table>`;
    } else if (maxLen <= 35) {
        let r1 = '', r2 = '';
        opsiArray.forEach((opt, idx) => {
            const cell = `<td style="width:33.33%; vertical-align:top;">${labels[idx]}. ${opt.teks}</td>`;
            if (idx % 2 === 0) r1 += cell; else r2 += cell;
        });
        return `
            <table style="width:100%; border-collapse:collapse; margin-bottom:2px;"><tr>${r1}</tr></table>
            <table style="width:100%; border-collapse:collapse;"><tr>${r2}</tr></table>`;
    } else {
        return opsiArray.map((opt, idx) => 
            `<div style="margin-bottom:2px;">${labels[idx]}. ${opt.teks}</div>`
        ).join('');
    }
}

function renderPreviewSoal(pgList, essayList, containerElement) {
    containerElement.innerHTML = generateFullExamHTML(pgList, essayList, "A (Pratinjau)");

    if (window.renderMathInElement) {
        renderMathInElement(containerElement, {
            delimiters: [
                {left: '$$', right: '$$', display: true},
                {left: '$', right: '$', display: false}
            ],
            throwOnError: false
        });
    }
}

// Generasi HTML Dokumen Menggunakan GAMBAR KOP UTUH
function generateFullExamHTML(pgList, essayList, packetName) {
    const judul = document.getElementById('metaJudul')?.value || "SOAL ULANGAN SEMESTER GENAP";
    const mapel = document.getElementById('metaMapel')?.value || "Pemrograman Web";
    const kelas = document.getElementById('metaKelas')?.value || "XI RPL 1, 2";
    const waktu = document.getElementById('metaWaktu')?.value || "-";

    // Elemen Kop Surat (Memuat Gambar Utama atau Placeholder jika belum diunggah)
    let kopSection = '';
    if (kopGambarBase64) {
        kopSection = `<div style="text-align:center; margin-bottom:10px;"><img src="${kopGambarBase64}" style="width:100%; max-width:680px; height:auto;"/></div>`;
    } else {
        kopSection = `
        <div style="border-bottom: 3px double #000; padding-bottom: 5px; margin-bottom: 10px; text-align: center;">
            <div style="font-size:12pt; font-weight:bold;">[ KOP SURAT SEKOLAH ]</div>
            <div style="font-size:9pt; color:#666;">(Unggah gambar Kop Surat pada form di atas)</div>
        </div>`;
    }

    let html = `
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body { font-family: 'Times New Roman', Times, serif; font-size: 11pt; color: #000; line-height: 1.2; }
            table { border-collapse: collapse; width: 100%; }
            .border-single { border-bottom: 2px solid #000; margin-bottom: 12px; }
            .text-center { text-align: center; }
            .font-bold { font-weight: bold; }
        </style>
    </head>
    <body>
        <!-- KOP SURAT BERBENTUK GAMBAR UTUH -->
        ${kopSection}

        <!-- JUDUL & METADATA -->
        <div class="text-center font-bold" style="font-size:12pt; margin-bottom:12px; text-decoration:underline;">
            ${judul} (PAKET ${packetName})
        </div>

        <table style="width:100%; font-size:10.5pt; margin-bottom:8px;">
            <tr>
                <td style="width:18%;">Program Diklat</td>
                <td style="width:2%;">:</td>
                <td style="width:40%; font-weight:bold;">${mapel}</td>
                <td style="width:18%;">Hari / Tanggal</td>
                <td style="width:2%;">:</td>
                <td style="width:20%;">${waktu}</td>
            </tr>
            <tr>
                <td>Kelas / Jurusan</td>
                <td>:</td>
                <td>${kelas}</td>
                <td>Waktu</td>
                <td>:</td>
                <td>-</td>
            </tr>
        </table>

        <div class="border-single"></div>

        <div style="font-weight:bold; margin-bottom:8px;">Pilihlah jawaban berikut dengan benar!</div>

        <!-- DAFTAR SOAL PG -->
        <div>
    `;

    pgList.forEach((q, idx) => {
        const shuffledOpsi = shuffleArray(q.opsi);
        const opsiLayout = getOpsiLayoutHtml(shuffledOpsi);
        const gambarHtml = q.gambar ? `<div style="margin:4px 0 4px 24px;"><img src="${q.gambar}" style="max-width:200px; height:auto;"/></div>` : '';

        html += `
        <div style="margin-bottom:10px; page-break-inside:avoid;">
            <table style="width:100%;">
                <tr>
                    <td style="width:24px; vertical-align:top; font-weight:bold;">${idx + 1}.</td>
                    <td style="vertical-align:top;">${q.soal}</td>
                </tr>
            </table>
            ${gambarHtml}
            <div style="margin-left:24px; margin-top:2px;">
                ${opsiLayout}
            </div>
        </div>`;
    });

    if (essayList.length > 0) {
        html += `<div style="font-weight:bold; margin:16px 0 8px 0;">Jawablah pertanyaan berikut dengan singkat dan jelas!</div>`;
        essayList.forEach((q, idx) => {
            const gambarHtml = q.gambar ? `<div style="margin:4px 0 4px 24px;"><img src="${q.gambar}" style="max-width:200px; height:auto;"/></div>` : '';
            html += `
            <div style="margin-bottom:10px; page-break-inside:avoid;">
                <table style="width:100%;">
                    <tr>
                        <td style="width:24px; vertical-align:top; font-weight:bold;">${idx + 1}.</td>
                        <td style="vertical-align:top;">${q.soal}</td>
                    </tr>
                </table>
                ${gambarHtml}
            </div>`;
        });
    }

    html += `</div></body></html>`;
    return html;
}

function downloadExcelTemplate(optionCount) {
    const headers = ["Pertanyaan / Soal", "Opsi A", "Opsi B", "Opsi C", "Opsi D"];
    if (optionCount === 5) headers.push("Opsi E");
    headers.push("Kunci Jawaban", "URL / Base64 Gambar (Opsional)");

    const sampleData = [
        headers,
        ["Hasil dari $\\sum_{i=1}^{3} i^2$ adalah...", "12", "14", "16", "18", "20", "B", ""].slice(0, optionCount + 3)
    ];

    const ws = XLSX.utils.aoa_to_sheet(sampleData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "MasterSoal");
    XLSX.writeFile(wb, `Template_Master_Soal.xlsx`);
}

document.addEventListener('DOMContentLoaded', () => {
    initShuffler();
});
