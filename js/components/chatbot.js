import { CONFIG } from '../config.js';

export function initChatbot() {
    const toggleChatbotBtn = document.getElementById('toggleChatbotBtn');
    const closeChatbotBtn = document.getElementById('closeChatbotBtn');
    const chatbotWindow = document.getElementById('chatbotWindow');
    const chatForm = document.getElementById('chatForm');
    const chatInput = document.getElementById('chatInput');
    const chatMessages = document.getElementById('chatMessages');

    if (!toggleChatbotBtn || !chatbotWindow) return;

    // Toggle Tampilan Chatbot
    toggleChatbotBtn.addEventListener('click', () => {
        if (chatbotWindow.classList.contains('hidden')) {
            chatbotWindow.classList.remove('hidden');
            setTimeout(() => chatbotWindow.classList.add('chat-popup-enter-active'), 10);
        } else {
            chatbotWindow.classList.remove('chat-popup-enter-active');
            setTimeout(() => chatbotWindow.classList.add('hidden'), 200);
        }
    });

    if (closeChatbotBtn) {
        closeChatbotBtn.addEventListener('click', () => {
            chatbotWindow.classList.remove('chat-popup-enter-active');
            setTimeout(() => chatbotWindow.classList.add('hidden'), 200);
        });
    }

    // Append Pesan User
    function appendUserMessage(msg) {
        const userDiv = document.createElement('div');
        userDiv.className = "flex gap-2 items-start justify-end";
        userDiv.innerHTML = `
            <div class="bg-emerald-600 text-white p-2.5 rounded-2xl rounded-tr-xs shadow-xs max-w-[85%] leading-relaxed text-[11px]">
                ${msg}
            </div>
        `;
        chatMessages.appendChild(userDiv);
        chatMessages.scrollTop = chatMessages.scrollHeight;
    }

    // Append Pesan AI
    function appendAiMessage(msg) {
        const aiDiv = document.createElement('div');
        aiDiv.className = "flex gap-2 items-start";
        aiDiv.innerHTML = `
            <div class="w-6 h-6 rounded-full bg-emerald-600 text-white flex items-center justify-center text-[10px] shrink-0 mt-0.5">🤖</div>
            <div class="bg-white border border-slate-200 p-2.5 rounded-2xl rounded-tl-xs text-slate-700 shadow-xs max-w-[85%] leading-relaxed text-[11px]">
                ${msg}
            </div>
        `;
        chatMessages.appendChild(aiDiv);
        chatMessages.scrollTop = chatMessages.scrollHeight;
    }

    // Handle Submit Form
    if (chatForm) {
        chatForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const query = chatInput.value.trim();
            if (!query) return;

            appendUserMessage(query);
            chatInput.value = "";

            const promptText = `Bertindaklah sebagai Konsultan Pendidikan AI yang ramah, ringkas, dan sangat membantu guru sekolah. Jawab pertanyaan berikut: ${query}`;

            try {
                const response = await fetch(CONFIG.PROXY_URL, {
                    method: "POST",
                    body: JSON.stringify({ contents: [{ parts: [{ text: promptText }] }] })
                });

                const rawText = await response.text();
                let reply = rawText;

                try {
                    const parsed = JSON.parse(rawText);
                    reply = parsed.candidates?.[0]?.content?.parts?.[0]?.text || rawText;
                } catch(e) {}

                appendAiMessage(reply.replace(/\*\*/g, "").replace(/\*/g, ""));

            } catch (err) {
                appendAiMessage("Maaf, koneksi server AI sedang padat. Silakan coba tanyakan kembali.");
            }
        });
    }
}