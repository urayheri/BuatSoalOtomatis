import { shuffleArray } from '../utils.js';

let originalPGQuestions = [];
let originalEssayQuestions = [];
let kopGambarBase64 = "";

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
    if (kopGambarInput) {
        kopGambarInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = (evt) => { 
                    kopGambarBase64 = evt.target.result;
                    
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

                    updatePreview(); 
                };
                reader.readAsDataURL(file);
            }
        });
    }

    if (downloadEBtn) downloadEBtn.addEventListener('click', () => downloadExcelTemplate(5));
    if (downloadDBtn) downloadDBtn.addEventListener('click', () => downloadExcelTemplate(4));

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
                                { teks: String(row[1] || "-") },
                                { teks: String(row[2] || "-") },
                                { teks: String(row[3] || "-") },
                                { teks: String(row[4] || "-") },
                                ...(row[5] ? [{ teks: String(row[5]) }] : [])
                            ],
                            gambar: row[7] || ""
                        });
                    } else {
                        originalEssayQuestions.push({
                            soal: String(row[0]),
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

// LOGIKA LAYOUT OPSI ADAPTIF (1 BARIS, 2 BARIS, ATAU KE BAWAH)
function renderOpsiAdaptif(opsiArray) {
    const labels = ["A", "B", "C", "D", "E"];
    
    // Hitung panjang teks opsi terpanjang
    const maxLen = Math.max(...opsiArray.map(o => (o.teks || '').length));

    // KONDISI 1: Sangat Pendek (misal: angka, kata pendek) -> TAMPIL 1 BARIS HORIZONTAL
    if (maxLen <= 12) {
        let items = opsiArray.map((opt, idx) => `
            <span style="display: inline-block; margin-right: 24pt; font-family: 'Times New Roman', Times, serif; font-size: 12pt;">
                <b>${labels[idx]}.</b> ${opt.teks}
            </span>
        `).join('');
        return `<div style="margin-left: 24pt; margin-top: 2pt; margin-bottom: 4pt;">${items}</div>`;
    } 
    // KONDISI 2: Sedang -> TAMPIL 2 BARIS / 2 KOLOM
    else if (maxLen <= 30) {
        let items = opsiArray.map((opt, idx) => `
            <div style="display: inline-block; width: 45%; vertical-align: top; margin-bottom: 2pt; font-family: 'Times New Roman', Times, serif; font-size: 12pt;">
                <b>${labels[idx]}.</b> ${opt.teks}
            </div>
        `).join('');
        return `<div style="margin-left: 24pt; margin-top: 2pt; margin-bottom: 4pt;">${items}</div>`;
    } 
    // KONDISI 3: Panjang -> TAMPIL MENURUN PER BARIS
    else {
        return opsiArray.map((opt, idx) => `
            <div style="margin-left: 24pt; margin-bottom: 2pt; text-align: justify; font-family: 'Times New Roman', Times, serif; font-size: 12pt;">
                <b>${labels[idx]}.</b> ${opt.teks}
            </div>
        `).join('');
    }
}

// GENERASI HTML DOKUMEN FULL TIMES NEW ROMAN 12PT & JUSTIFY
function generateFullExamHTML(pgList, essayList, packetName) {
    const judul = document.getElementById('metaJudul')?.value || "SOAL ULANGAN SEMESTER GENAP";
    const mapel = document.getElementById('metaMapel')?.value || "Pemrograman Web";
    const kelas = document.getElementById('metaKelas')?.value || "XI RPL 1, 2";
    const waktu = document.getElementById('metaWaktu')?.value || "-";

    // ELEMEN KOP SURAT
    let kopSection = '';
    if (kopGambarBase64) {
        kopSection = `
        <div style="text-align: center; margin-bottom: 12pt;">
            <img src="${kopGambarBase64}" style="width: 100%; max-width: 650px; height: auto; display: block; margin: 0 auto;" />
        </div>`;
    } else {
        kopSection = `
        <div style="border-bottom: 3px double #000; padding-bottom: 5px; margin-bottom: 12pt; text-align: center; font-family: 'Times New Roman', Times, serif;">
            <div style="font-size: 14pt; font-weight: bold;">[ KOP SURAT SEKOLAH ]</div>
            <div style="font-size: 10pt; color: #555;">(Unggah gambar Kop Surat pada menu di atas)</div>
        </div>`;
    }

    let html = `
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            @page {
                size: A4;
                margin: 2cm;
            }
            * {
                font-family: 'Times New Roman', Times, serif !important;
            }
            body { 
                font-family: 'Times New Roman', Times, serif !important; 
                font-size: 12pt !important; 
                color: #000; 
                line-height: 1.25; 
                text-align: justify;
            }
            p, div, td, span, b {
                font-family: 'Times New Roman', Times, serif !important;
                font-size: 12pt !important;
            }
            p {
                margin: 0 0 4pt 0;
                text-align: justify;
            }
            .soal-text {
                margin-left: 24pt;
                text-indent: -24pt;
                margin-bottom: 4pt;
                text-align: justify;
            }
            .header-info {
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 8pt;
                font-size: 12pt !important;
            }
            .header-info td {
                padding: 1pt 0;
                vertical-align: top;
                font-size: 12pt !important;
            }
            .border-line {
                border-bottom: 1.5pt solid #000;
                margin-bottom: 12pt;
            }
        </style>
    </head>
    <body>
        <!-- KOP SURAT -->
        ${kopSection}

        <!-- JUDUL & METADATA UJIAN -->
        <div style="text-align: center; font-weight: bold; font-size: 12pt !important; margin-bottom: 10pt; text-transform: uppercase; text-decoration: underline;">
            ${judul} (PAKET ${packetName})
        </div>

        <table class="header-info">
            <tr>
                <td style="width: 18%;">Program Diklat</td>
                <td style="width: 2%;">:</td>
                <td style="width: 40%; font-weight: bold;">${mapel}</td>
                <td style="width: 18%;">Hari / Tanggal</td>
                <td style="width: 2%;">:</td>
                <td style="width: 20%;">${waktu}</td>
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

        <div class="border-line"></div>

        <p style="font-weight: bold; margin-bottom: 10pt;">Pilihlah jawaban berikut dengan benar!</p>

        <!-- DAFTAR SOAL PG -->
        <div>
    `;

    pgList.forEach((q, idx) => {
        const shuffledOpsi = shuffleArray(q.opsi);
        const opsiContent = renderOpsiAdaptif(shuffledOpsi);
        const gambarHtml = q.gambar ? `<div style="margin: 4pt 0 4pt 24pt;"><img src="${q.gambar}" style="max-width: 250px; height: auto;"/></div>` : '';

        html += `
        <div style="margin-bottom: 8pt; page-break-inside: avoid;">
            <div class="soal-text">
                <b>${idx + 1}.</b> ${q.soal}
            </div>
            ${gambarHtml}
            ${opsiContent}
        </div>`;
    });

    if (essayList.length > 0) {
        html += `<p style="font-weight: bold; margin: 14pt 0 8pt 0;">Jawablah pertanyaan berikut dengan singkat dan jelas!</p>`;
        essayList.forEach((q, idx) => {
            const gambarHtml = q.gambar ? `<div style="margin: 4pt 0 4pt 24pt;"><img src="${q.gambar}" style="max-width: 250px; height: auto;"/></div>` : '';
            html += `
            <div style="margin-bottom: 8pt; page-break-inside: avoid;">
                <div class="soal-text">
                    <b>${idx + 1}.</b> ${q.soal}
                </div>
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
