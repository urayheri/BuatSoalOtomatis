import { CONFIG } from '../config.js';

let modulAjarGeneratedContent = "";

export function initModulAjar() {
    const processModulBtn = document.getElementById('processModulBtn');
    const previewSection = document.getElementById('previewSection');
    const previewTitle = document.getElementById('previewTitle');
    const exportSingleWordBtn = document.getElementById('exportSingleWordBtn');
    const modulEditableArea = document.getElementById('modulEditableArea');

    if (!processModulBtn) return;

    processModulBtn.addEventListener('click', async () => {
        const mapel = document.getElementById('modulMapel')?.value.trim();
        const fase = document.getElementById('modulFase')?.value;
        const topik = document.getElementById('modulTopik')?.value.trim();
        const waktu = document.getElementById('modulWaktu')?.value.trim();
        const fasilitas = document.getElementById('modulFasilitas')?.value;
        const namaGuru = document.getElementById('inputNamaGuru')?.value;
        const nipGuru = document.getElementById('inputNipGuru')?.value;
        const namaKepsek = document.getElementById('inputNamaKepsek')?.value;
        const nipKepsek = document.getElementById('inputNipKepsek')?.value;
        const namaSekolah = document.getElementById('inputNamaSekolah')?.value;

        const checkboxes = document.querySelectorAll('#dimensiProfilGroup input[type="checkbox"]:checked');
        const profilTerpilih = Array.from(checkboxes).map(cb => cb.value).join(', ');

        if (!topik) {
            alert("Silakan masukkan Topik / Materi Pokok Pembahasan terlebih dahulu!");
            return;
        }

        if (checkboxes.length === 0) {
            alert("Silakan centang minimal satu Target Dimensi Profil Lulusan!");
            return;
        }

        processModulBtn.disabled = true;
        processModulBtn.textContent = "AI Sedang Merumuskan Modul Ajar...";

        const promptText = `Bertindaklah sebagai Konsultan Ahli Kurikulum Merdeka Kemendikbudristek RI. Buatkan naskah RENCANA PELAKSANAAN PEMBELAJARAN MENDALAM (DEEP LEARNING) yang SANGAT LENGKAP, PANJANG, MENDALAM, dan SANGAT DETAIL dalam bentuk paragraf naratif utuh.

📌 DATA ADMINISTRASI UTAMA:
- Nama Guru: ${namaGuru}
- NIP Guru: ${nipGuru}
- Satuan Pendidikan: ${namaSekolah}
- Nama Kepala Sekolah: ${namaKepsek}
- NIP Kepala Sekolah: ${nipKepsek}
- Mata Pelajaran: ${mapel}
- Fase / Kelas: ${fase}
- Topik / Materi: ${topik}
- Alokasi Waktu: ${waktu}
- Target Profil Lulusan: ${profilTerpilih}
- Kondisi Fasilitas: ${fasilitas}

📋 STRUKTUR DOKUMEN UTUH YANG WAJIB DIHASILKAN:
BAGIAN I: IDENTITAS & ANALISIS KONDISI PEMBELAJARAN
BAGIAN II: DESAIN INSTRUKSIONAL & CAPAIAN
BAGIAN III: PENGALAMAN BELAJAR MENDALAM (LANGKAH-LANGKAH PEMBELAJARAN)
BAGIAN IV: BAHAN AJAR / RINGKASAN MATERI PEMBELAJARAN
BAGIAN V: INSTRUMEN ASESMEN & EVALUASI LENGKAP
BAGIAN VI: LEMBAR REFLEKSI GURU
BAGIAN VII: LEMBAR PENGESAHAN / TANDA TANGAN`;

        try {
            const response = await fetch(CONFIG.PROXY_URL, {
                method: "POST",
                headers: { "Content-Type": "text/plain;charset=utf-8" },
                body: JSON.stringify({ contents: [{ parts: [{ text: promptText }] }] })
            });

            const rawText = await response.text();
            let aiResponseText = rawText;

            try {
                const parsedRaw = JSON.parse(rawText);
                aiResponseText = parsedRaw.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || rawText;
            } catch (e) {}

            modulAjarGeneratedContent = aiResponseText
                .replace(/^```(json|markdown)?/i, '')
                .replace(/```$/, '')
                .replace(/\*\*/g, '')
                .replace(/\*/g, '')
                .trim();

            previewSection.classList.remove('hidden');
            if (previewTitle) previewTitle.textContent = "Pratinjau Hasil Rumusan Modul Ajar Kurikulum Merdeka";
            if (modulEditableArea) modulEditableArea.value = modulAjarGeneratedContent;

            previewSection.scrollIntoView({ behavior: 'smooth' });

        } catch (err) {
            alert("Gagal merumuskan modul: " + err.message);
        } finally {
            processModulBtn.disabled = false;
            processModulBtn.textContent = "✨ Susun Perangkat Ajar Bersama AI";
        }
    });

    // Event Sync Textarea Manual Edit
    if (modulEditableArea) {
        modulEditableArea.addEventListener('input', function () {
            modulAjarGeneratedContent = this.value;
        });
    }

    // Ekspor ke Word (.docx)
    if (exportSingleWordBtn) {
        exportSingleWordBtn.addEventListener('click', async () => {
            exportSingleWordBtn.disabled = true;
            try {
                const { Document, Paragraph, TextRun, Packer } = window.docx;
                const lines = (modulAjarGeneratedContent || "").split('\n');
                
                const children = lines.map(line => {
                    const cleanLine = line.trim();
                    const isHeading = cleanLine.startsWith('BAGIAN') || cleanLine.startsWith('#');
                    return new Paragraph({
                        children: [
                            new TextRun({
                                text: cleanLine.replace(/^#+\s*/, ''),
                                bold: isHeading,
                                size: isHeading ? 26 : 22,
                                font: "Times New Roman"
                            })
                        ],
                        spacing: { before: isHeading ? 180 : 0, after: 120, line: 276 }
                    });
                });

                const doc = new Document({
                    sections: [{ children }]
                });

                const blob = await Packer.toBlob(doc);
                const link = document.createElement('a');
                link.href = URL.createObjectURL(blob);
                link.download = "Modul_Ajar_Kurikulum_Merdeka.docx";
                link.click();

            } catch (e) {
                alert("Gagal mengunduh berkas Word: " + e.message);
            } finally {
                exportSingleWordBtn.disabled = false;
            }
        });
    }
}

// Auto-init saat dimuat mandiri
document.addEventListener('DOMContentLoaded', () => {
    initModulAjar();
});