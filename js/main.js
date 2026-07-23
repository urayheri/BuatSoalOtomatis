import { CONFIG } from './config.js';
import { initShuffler } from './modules/shuffler.js';
import { initQuestionGen } from './modules/question-gen.js';
import { initModulAjar } from './modules/modul-ajar.js';

document.addEventListener('DOMContentLoaded', () => {
    const tabAcakBtn = document.getElementById('tabAcakBtn');
    const tabBuatBtn = document.getElementById('tabBuatBtn');
    const tabModulBtn = document.getElementById('tabModulBtn');

    const activeClass = "w-full py-3 px-4 rounded-xl bg-emerald-600 text-white font-bold text-xs flex items-center justify-center gap-2 transition-all cursor-pointer shadow-sm";
    const inactiveClass = "w-full py-3 px-4 rounded-xl text-slate-600 hover:bg-slate-100 font-semibold text-xs flex items-center justify-center gap-2 transition-all cursor-pointer";

    function setActiveTab(btnTerpilih) {
        [tabAcakBtn, tabBuatBtn, tabModulBtn].forEach(btn => {
            if (btn) btn.className = inactiveClass;
        });
        if (btnTerpilih) btnTerpilih.className = activeClass;
    }

    if (tabAcakBtn) {
        tabAcakBtn.addEventListener('click', () => {
            setActiveTab(tabAcakBtn);
            initShuffler();
        });
    }

    if (tabBuatBtn) {
        tabBuatBtn.addEventListener('click', () => {
            setActiveTab(tabBuatBtn);
            initQuestionGen();
        });
    }

    if (tabModulBtn) {
        tabModulBtn.addEventListener('click', () => {
            setActiveTab(tabModulBtn);
            initModulAjar();
        });
    }
    

    // Default Load Mode 1
    initShuffler();
});
