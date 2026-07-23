// Fungsi Toast Notifikasi
export function showCustomToast(message, type = 'error', title = null) {
    const toastModal = document.getElementById('customToastModal');
    if (!toastModal) {
        alert(`${title ? title + ': ' : ''}${message}`);
        return;
    }
    // Logic modal toast...
}

// Helper Pengacak Array
export function shuffleArray(array) {
    let arr = JSON.parse(JSON.stringify(array));
    for (let i = arr.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [arr[i], arr[j]] = [arr[j], arr[i]];
    }
    return arr;
}