export function initModals() {
    // Modal Apresiasi / Sedekah
    const openSedekahBtn = document.getElementById('openSedekahBtn');
    const closeSedekahBtn = document.getElementById('closeSedekahBtn');
    const sedekahModal = document.getElementById('sedekahModal');

    if (openSedekahBtn && sedekahModal) {
        openSedekahBtn.addEventListener('click', () => sedekahModal.classList.remove('hidden'));
    }
    if (closeSedekahBtn && sedekahModal) {
        closeSedekahBtn.addEventListener('click', () => sedekahModal.classList.add('hidden'));
    }

    // Modal Statement Pengembang
    const openDevBtn = document.getElementById('openDevStatementBtn');
    const closeDevBtn = document.getElementById('closeDevStatementBtn');
    const devModal = document.getElementById('devStatementModal');

    if (openDevBtn && devModal) {
        openDevBtn.addEventListener('click', () => devModal.classList.remove('hidden'));
    }
    if (closeDevBtn && devModal) {
        closeDevBtn.addEventListener('click', () => devModal.classList.add('hidden'));
    }
}