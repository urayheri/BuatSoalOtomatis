import { CONFIG } from '../config.js';

let extractedTextFromDocument = "";
let generatedPGQuestions = [];
let generatedEssayQuestions = [];

export function initQuestionGen() {
    const materiFileInput = document.getElementById('materiFileInput');
    const fileMateriStatus = document.getElementById('fileMateriStatus');
    const materiInputText = document.getElementById('materiInputText');
    const processGenerateBtn = document.getElementById('processGenerateBtn');
    const generateCountPG = document.getElementById('generateCountPG');
    const generateCountEssay = document.getElementById('generateCountEssay');
    const previewSection = document.getElementById('previewSection');
    const soalContainer = document.getElementById('soalContainer');

    if (!processGenerateBtn) return;

    // Hitung Otomatis Total Soal PG
    const updateTotalPG = () => {
        const c1 = parseInt(document.getElementById('countC1')?.value) || 0;
        const c2 = parseInt(document.getElementById('countC2')?.value) || 0;
        const c3 = parseInt(document.getElementById('countC3')?.value) || 0;
        const c4 = parseInt(document.getElementById('countC4')?.value) || 0;
        const c5 = parseInt(document.getElementById('countC5')?.value) || 0;
        const c6 = parseInt(document.getElementById('countC6')?.value) || 0;
        if (generateCountPG) generateCountPG.value = c1 + c2 + c3 + c4 + c5 + c6;
    };

    // Hitung Otomatis Total Soal Uraian
    const updateTotalEssay = () => {
        const ec1 = parseInt(document.getElementById('countEssayC1')?.value) || 0;
        const ec2 = parseInt(document.getElementById('countEssayC2')?.value) || 0;
        const ec3 = parseInt(document.getElementById('countEssayC3')?.value) || 0;
        const ec4 = parseInt(document.getElementById('countEssayC4')?.value) || 0;
        const ec5 = parseInt(document.getElementById('countEssayC5')?.value) || 0;
        const ec6 = parseInt(document.getElementById('countEssayC6')?.value) || 0;
        if (generateCountEssay) generateCountEssay.value = ec1 + ec2 + ec3 + ec4 + ec5 + ec6;
    };

    document.querySelectorAll('.cog-input').forEach(i => i.addEventListener('input', updateTotalPG));
    document.querySelectorAll('.essay-cog-input').forEach(i => i.addEventListener('input', updateTotalEssay));

    // Ekstraksi Berkas Referensi (.docx atau .pdf)
    if (materiFileInput) {
        materiFileInput.addEventListener('change', async (e) => {
            const file = e.target.files[0];
            if (!file) return;

            fileMateriStatus.textContent = "Sedang memindai data dokumen...";
            fileMateriStatus.className = "text-[10px] text-emerald-600 font-medium block mt-0.5";
            extractedTextFromDocument = "";

            const reader = new FileReader();

            if (file.name.endsWith('.docx')) {
                reader.onload = (event) => {
                    try {
                        mammoth.extractRawText({ arrayBuffer: event.target.result }).then((result) => {
                            extractedTextFromDocument = result.value || "";
                            const wordCount = extractedTextFromDocument.split(/\s+/).filter(Boolean).length;
                            fileMateriStatus.textContent = `✓ Berkas Word berhasil dimuat (${wordCount} kata).`;
                            fileMateriStatus.className = "text-[10px] text-emerald-600 font-semibold block mt-0.5";
                        });
                    } catch (err) {
                        fileMateriStatus.textContent = "❌ Gagal membaca berkas Word.";
                        fileMateriStatus.className = "text-[10px] text-red-500 block";
                    }
                };
                reader.readAsArrayBuffer(file);
            } else if (file.name.endsWith('.pdf')) {
                reader.onload = async (event) => {
                    try {
                        const typedarray = new Uint8Array(event.target.result);
                        const pdf = await pdfjsLib.getDocument(typedarray).promise;
                        let textContent = "";
                        for (let i = 1; i <= pdf.numPages; i++) {
                            const page = await pdf.getPage(i);
                            const textObj = await page.getTextContent();
                            textContent += textObj.items.map(item => item.str).join(" ") + "\n";
                        }
                        extractedTextFromDocument = textContent || "";
                        const wordCount = extractedTextFromDocument.split(/\s+/).filter(Boolean).length;
                        fileMateriStatus.textContent = `✓ Berkas PDF berhasil dimuat (${wordCount} kata).`;
                        fileMateriStatus.className = "text-[10px] text-emerald-600 font-semibold block mt-0.5";
                    } catch (err) {
                        fileMateriStatus.textContent = "❌ Gagal membaca berkas PDF.";
                        fileMateriStatus.className = "text-[10px] text-red-500 block";
                    }
                };
                reader.readAsArrayBuffer(file);
            }
        });
    }

    // Proses Generasi Soal via API
    processGenerateBtn.addEventListener('click', async () => {
        const fileText = extractedTextFromDocument.trim();
        const manualText = materiInputText ? materiInputText.value.trim() : "";
        const materi = fileText !== "" ? fileText : manualText;

        if (!materi) {
            alert("Silakan unggah dokumen atau tempel ringkasan materi terlebih dahulu.");
            return;
        }

        const totalPG = parseInt(generateCountPG.value) || 0;
        const totalEssay = parseInt(generateCountEssay.value) || 0;

        if (totalPG === 0 && totalEssay === 0) {
            alert("Jumlah target pembuatan soal tidak boleh 0.");
            return;
        }

        previewSection.classList.remove('hidden');
        soalContainer.innerHTML = `
            <div class="flex flex-col items-center justify-center py-10 space-y-3">
                <div class="w-8 h-8 border-4 border-emerald-600 border-t-transparent rounded-full animate-spin"></div>
                <p class="text-xs font-bold text-slate-600">AI sedang memformulasikan soal HOTS, mohon tunggu...</p>
            </div>
        `;

        processGenerateBtn.disabled = true;
        processGenerateBtn.textContent = "AI Sedang Berpikir...";

        const promptText = `Bertindaklah sebagai Pakar Penyusun Soal Evaluasi Pendidikan.
Buatkan naskah bank soal berdasarkan referensi materi berikut: "${materi.substring(0, 4000)}".

Kriteria:
- Total Pilihan Ganda: ${totalPG} soal.
- Total Uraian: ${totalEssay} soal.

Format Output WAJIB JSON murni tanpa markdown codeblock:
{
  "pilihan_ganda": [
    {
      "soal": "Pertanyaan...",
      "opsi": [{"teks": "A"}, {"teks": "B"}, {"teks": "C"}, {"teks": "D"}],
      "kunciOriginal": "A",
      "kognitif": "C3"
    }
  ],
  "essay": [
    {
      "soal": "Pertanyaan uraian...",
      "kunci": "Jawaban/Rubrik...",
      "kognitif": "C4"
    }
  ]
}`;

        try {
            const response = await fetch(CONFIG.PROXY_URL, {
                method: "POST",
                headers: { "Content-Type": "text/plain;charset=utf-8" },
                body: JSON.stringify({ contents: [{ parts: [{ text: promptText }] }] })
            });

            const rawText = await response.text();
            let cleanJson = rawText;

            try {
                const parsedResponse = JSON.parse(rawText);
                if (parsedResponse.candidates?.[0]?.content?.parts?.[0]?.text) {
                    cleanJson = parsedResponse.candidates[0].content.parts[0].text;
                }
            } catch (e) {}

            cleanJson = cleanJson.replace(/```json/gi, '').replace(/```/g, '').trim();
            const data = JSON.parse(cleanJson);

            generatedPGQuestions = data.pilihan_ganda || [];
            generatedEssayQuestions = data.essay || [];

            renderPreview(soalContainer);

        } catch (err) {
            soalContainer.innerHTML = `<div class="p-4 bg-red-50 text-red-700 rounded-xl font-semibold">❌ Gagal memproses: ${err.message}</div>`;
        } finally {
            processGenerateBtn.disabled = false;
            processGenerateBtn.textContent = "✨ Buat Soal Otomatis";
        }
    });
}

function renderPreview(container) {
    container.innerHTML = "";

    if (generatedPGQuestions.length > 0) {
        container.innerHTML += `<h4 class="font-bold text-emerald-800 text-xs uppercase mb-2">A. Pilihan Ganda</h4>`;
        generatedPGQuestions.forEach((q, idx) => {
            let opsiHtml = (q.opsi || []).map((o, oIdx) => `
                <div class="bg-slate-50 p-1.5 rounded border border-slate-200">
                    <span class="font-bold text-emerald-700">${String.fromCharCode(65 + oIdx)}.</span> ${o.teks || o}
                </div>
            `).join('');

            container.innerHTML += `
                <div class="py-3 space-y-1.5">
                    <p class="font-semibold text-slate-800">${idx + 1}. ${q.soal} <span class="text-[10px] bg-slate-100 text-slate-600 px-2 py-0.5 rounded font-bold">${q.kognitif || ''}</span></p>
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-2 text-[11px] pl-4">${opsiHtml}</div>
                </div>
            `;
        });
    }

    if (generatedEssayQuestions.length > 0) {
        container.innerHTML += `<h4 class="font-bold text-emerald-800 text-xs uppercase mt-4 mb-2">B. Soal Uraian / Essay</h4>`;
        generatedEssayQuestions.forEach((q, idx) => {
            container.innerHTML += `
                <div class="py-3 space-y-1">
                    <p class="font-semibold text-slate-800">${idx + 1}. ${q.soal} <span class="text-[10px] bg-slate-100 text-slate-600 px-2 py-0.5 rounded font-bold">${q.kognitif || ''}</span></p>
                    <p class="text-[11px] text-slate-500 italic pl-4">Kunci: ${q.kunci}</p>
                </div>
            `;
        });
    }
}

// Auto-init saat dimuat mandiri
document.addEventListener('DOMContentLoaded', () => {
    initQuestionGen();
});