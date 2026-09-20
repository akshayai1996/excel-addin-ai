/**
 * Taskpane Frontend Logic
 */

document.addEventListener("DOMContentLoaded", async () => {
    const BUILD_TAG = "2r";
    console.log("[AI Assistant] sidebar build " + BUILD_TAG);
    const UI = {
        btnSettings: document.getElementById("btnSettings"),
        settingsModal: document.getElementById("settingsModal"),
        btnCloseSettings: document.getElementById("btnCloseSettings"),
        btnSaveSettings: document.getElementById("btnSaveSettings"),
        // Bug 49: txtGeminiKey removed — no Gemini key used in OpenCode flow.
        
        selectionBadge: document.getElementById("selectionBadge"),
        selectionText: document.getElementById("selectionText"),
        btnRefreshSelection: document.getElementById("btnRefreshSelection"),
        
        chatMessages: document.getElementById("chatMessages"),
        txtPrompt: document.getElementById("txtPrompt"),
        btnSend: document.getElementById("btnSend"),
        btnClearChat: document.getElementById("btnClearChat"),
        btnUndo: document.getElementById("btnUndo"),
        btnRedo: document.getElementById("btnRedo"),
        btnRestore: document.getElementById("btnRestore"),
        btnScrollLatest: document.getElementById("btnScrollLatest"),
        btnCollapse: document.getElementById("btnCollapse"),
        slimStatus: document.getElementById("slimStatus"),
        slimModel: document.getElementById("slimModel"),
        slimEffort: document.getElementById("slimEffort"),
        
        selectProvider: document.getElementById("selectProvider"),
        selectEffort: document.getElementById("selectEffort"),
        chkAutoApply: document.getElementById("chkAutoApply"),
        statusDot: document.getElementById("statusDot"),
        statusText: document.getElementById("statusText")
    };

    let activeSelectionContext = null;
    let modelVariants = {}; // provider value -> [variant ids]

    const EFFORT_LABELS = {
        minimal: "Minimal",
        low: "Low",
        medium: "Medium",
        high: "High",
        xhigh: "Max"
    };

    const READ_TIMEOUT_MS = 45000;

    function loadPref(key, fallback) {
        try {
            return localStorage.getItem(key) || fallback;
        } catch (_) {
            return fallback;
        }
    }

    function savePref(key, value) {
        try {
            localStorage.setItem(key, value);
        } catch (_) { /* storage unavailable */ }
    }

    function populateEffort(modelId) {
        const select = UI.selectEffort;
        const variants = modelVariants[modelId] || [];
        select.innerHTML = "";
        if (variants.length === 0) {
            const option = document.createElement("option");
            option.value = "";
            option.textContent = "Default";
            select.appendChild(option);
            select.disabled = true;
            return;
        }
        select.disabled = false;
        const saved = loadPref("excelAiEffort", "medium");
        variants.forEach(v => {
            const option = document.createElement("option");
            option.value = v;
            option.textContent = EFFORT_LABELS[v] || v;
            if (v === saved) option.selected = true;
            select.appendChild(option);
        });
        if (!select.value && select.options.length > 0) {
            const preferred = Array.from(select.options).find(o => o.value === "medium");
            select.selectedIndex = preferred ? preferred.index : 0;
        }
    }

    async function fetchModels() {
        try {
            const response = await fetch("http://localhost:3000/api/models");
            const data = await response.json();
            const select = UI.selectProvider;
            select.innerHTML = "";
            modelVariants = {};
            data.models.forEach(model => {
                const option = document.createElement("option");
                option.value = model.id;
                option.textContent = model.name;
                option.title = model.id + (model.context ? ` · ${(model.context / 1000).toFixed(0)}k context` : "");
                select.appendChild(option);
                modelVariants[model.id] = model.variants || [];
            });
            const savedModel = loadPref("excelAiModel", "");
            const defaultModel = "opencode:opencode/muse-spark-1.3-contributor-free";
            if (savedModel && modelVariants[savedModel]) {
                select.value = savedModel;
            } else if (modelVariants[defaultModel]) {
                select.value = defaultModel;
            } else if (select.options.length > 0) {
                select.selectedIndex = 0;
            }
            populateEffort(select.value);
            updateSlimStatus();
        } catch (e) {
            console.warn("Failed to fetch models dynamically:", e);
            // Fallback so chat still works when the model list can't load
            const select = UI.selectProvider;
            select.innerHTML = "";
            const option = document.createElement("option");
            option.value = "opencode:opencode/muse-spark-1.3-contributor-free";
            option.textContent = "muse-spark-1.3-contributor-free";
            select.appendChild(option);
            modelVariants[option.value] = ["minimal", "low", "medium", "high", "xhigh"];
            populateEffort(option.value);
            updateSlimStatus();
        }
    }
    fetchModels().then(restoreStartupSession);

    UI.selectProvider.addEventListener("change", () => {
        savePref("excelAiModel", UI.selectProvider.value);
        populateEffort(UI.selectProvider.value);
        updateSlimStatus();
        saveActiveSession();
    });
    UI.selectEffort.addEventListener("change", () => {
        savePref("excelAiEffort", UI.selectEffort.value);
        updateSlimStatus();
        saveActiveSession();
    });
    setCollapsed(loadPref("excelAiCollapsed", "0") === "1");

    // Load Settings
    UI.btnSettings.onclick = () => {
        renderDebugList();
        UI.settingsModal.classList.remove("hidden");
    };
    UI.btnCloseSettings.onclick = () => UI.settingsModal.classList.add("hidden");

    // Save a full .xlsx copy under any name
    const btnSaveCopy = document.getElementById("btnSaveCopy");
    if (btnSaveCopy) {
        btnSaveCopy.addEventListener("click", async () => {
            const nameInput = document.getElementById("txtCopyName");
            const resultEl = document.getElementById("saveCopyResult");
            btnSaveCopy.disabled = true;
            btnSaveCopy.innerText = "Saving…";
            try {
                const r = await window.ExcelBridge.saveCopyAs(nameInput ? nameInput.value : "workbook-copy");
                if (resultEl) resultEl.textContent = (r.success ? "✅ " : "❌ ") + r.message;
            } catch (err) {
                if (resultEl) resultEl.textContent = "❌ " + (err.message || err);
            }
            btnSaveCopy.disabled = false;
            btnSaveCopy.innerText = "Save copy";
        });
    }
    const buildTagEl = document.getElementById("buildTag");
    if (buildTagEl) buildTagEl.textContent = "Build " + BUILD_TAG + " — if I ship a fix, reopen the pane to load it.";
    UI.btnSaveSettings.onclick = () => {
        UI.settingsModal.classList.add("hidden");
    };

    // Auto-resize textarea
    UI.txtPrompt.addEventListener("input", function() {
        this.style.height = "auto";
        this.style.height = Math.min(this.scrollHeight, 150) + "px";
    });

    // Collapse / expand the model bar (choice persists)
    function updateSlimStatus() {
        const m = UI.selectProvider.selectedOptions && UI.selectProvider.selectedOptions[0];
        UI.slimModel.textContent = m ? m.textContent : "—";
        const e = UI.selectEffort.selectedOptions && UI.selectEffort.selectedOptions[0];
        UI.slimEffort.textContent = (!UI.selectEffort.disabled && e) ? e.textContent : "";
    }

    function setCollapsed(collapsed) {
        document.getElementById("app").classList.toggle("collapsed", collapsed);
        savePref("excelAiCollapsed", collapsed ? "1" : "0");
        UI.btnCollapse.title = collapsed ? "Show model bar" : "Hide model bar";
        UI.btnCollapse.setAttribute("aria-label", collapsed ? "Show model bar" : "Hide model bar");
        if (collapsed) updateSlimStatus();
    }

    UI.btnCollapse.addEventListener("click", () => {
        setCollapsed(!document.getElementById("app").classList.contains("collapsed"));
    });
    UI.slimStatus.addEventListener("click", () => setCollapsed(false));

    // "Jump to latest" pill — visible only when scrolled up, so chat is
    // navigable even where the OS hides scrollbars entirely.
    UI.chatMessages.addEventListener("scroll", () => {
        const gap = UI.chatMessages.scrollHeight - UI.chatMessages.scrollTop - UI.chatMessages.clientHeight;
        UI.btnScrollLatest.classList.toggle("show", gap > 120);
    });
    UI.btnScrollLatest.addEventListener("click", () => {
        const reduceMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
        UI.chatMessages.scrollTo({ top: UI.chatMessages.scrollHeight, behavior: reduceMotion ? "auto" : "smooth" });
    });

    // Clear Chat = archive current chat into history, start blank
    UI.btnClearChat.addEventListener("click", () => {
        newChatSession();
    });

    // History panel toggle + new chat
    const btnHistory = document.getElementById("btnHistory");
    const historyPanel = document.getElementById("historyPanel");
    if (btnHistory && historyPanel) {
        btnHistory.addEventListener("click", () => {
            historyPanel.hidden = !historyPanel.hidden;
            if (!historyPanel.hidden) renderHistoryList();
        });
    }
    const btnNewChat = document.getElementById("btnNewChat");
    if (btnNewChat) {
        btnNewChat.addEventListener("click", () => {
            newChatSession();
        });
    }

    // Undo last AI change (one entry per click; structural undos survive reloads)
    if (UI.btnUndo) {
        UI.btnUndo.addEventListener("click", async () => {
            const result = await window.ExcelBridge.undoLast();
            if (result.success) {
                const left = result.remaining > 0 ? ` (${result.remaining} left — one step per click)` : "";
                const redoHint = result.redoAvailable ? " Redo available." : "";
                appendMessage(`↩️ ${result.message}${left}${redoHint}`, "assistant");
            } else {
                appendMessage(`↩️ ${result.message} (Content undo clears on pane reload; structural undos + Restore survive.)`, "assistant");
            }
        });
    }

    // Redo last undone change (range/format only — not structural)
    if (UI.btnRedo) {
        UI.btnRedo.addEventListener("click", async () => {
            const result = await window.ExcelBridge.redoLast();
            if (result.success) {
                const left = result.remaining > 0 ? ` (${result.remaining} redo(s) left)` : "";
                appendMessage(`↪️ ${result.message}${left}`, "assistant");
            } else {
                appendMessage(`↪️ ${result.message}`, "assistant");
            }
        });
    }

    // Bug 6/7: Restore button — clarify what IS and is NOT restored.
    if (UI.btnRestore) {
        UI.btnRestore.title = "Restores values, formulas, formats, tables, and charts. PivotTables and comments are NOT restored.";
        UI.btnRestore.addEventListener("click", async () => {
            const result = await window.ExcelBridge.restoreLastBackup();
            const disclaimer = " Note: Restores values, formulas, formats, tables, charts. PivotTables and comments are not restored.";
            appendMessage(
                result.success
                    ? `💾 ${result.message}${disclaimer}`
                    : `${result.message}`,
                "assistant"
            );
        });
    }

    // Bug 1: Per-session conversationId (UUID) for server-side session isolation.
    // Sending this in every /api/chat request causes the Python bridge to maintain
    // a separate OpenCode session per chat — cross-workbook pollution is prevented.
    let _conversationId = _generateUUID();

    function _generateUUID() {
        return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, c => {
            const r = Math.random() * 16 | 0;
            return (c === "x" ? r : (r & 0x3 | 0x8)).toString(16);
        });
    }

    // Initialize Excel Bridge
    if (window.ExcelBridge) {
        const initialized = await window.ExcelBridge.init();
        if (initialized) {
            // Bug 46: Only add 'connected' class AFTER Office.onReady() confirms Excel host.
            UI.statusDot.classList.add("connected");
            UI.statusText.textContent = "Connected";

            // Listen for selection changes (debounced inside ExcelBridge)
            window.ExcelBridge.onSelectionChange((data) => {
                activeSelectionContext = data;
                updateSelectionUI(data);
            });

            // Initial fetch
            refreshSelection();
        } else {
            console.log("Running in DEV Mock mode (not inside Excel).");
            UI.statusText.textContent = "Preview mode";
            refreshSelection();
        }
    }

    function readWorkbookWithTimeout() {
        return Promise.race([
            window.ExcelBridge.getSelectedRangeData(),
            new Promise((_, reject) => setTimeout(() => reject(new Error(
                "Timed out after 45s reading the workbook. If the file was just " +
                "downloaded, click Enable Editing in Excel's yellow bar first. If it is " +
                "a very large workbook, try a smaller file to isolate the issue."
            )), READ_TIMEOUT_MS))
        ]);
    }

    async function refreshSelection() {
        UI.selectionText.textContent = "Fetching...";
        try {
            const data = await readWorkbookWithTimeout();
            activeSelectionContext = data;
            updateSelectionUI(data);
        } catch (err) {
            UI.selectionText.textContent = "⚠️ Could not read workbook";
            appendMessage(
                `⚠️ I couldn't read this workbook (${err.message || err}). ` +
                `If the file was just downloaded or received by email, Excel may have opened it in ` +
                `Protected View — click <b>Enable Editing</b> in the yellow bar at the top of Excel, ` +
                `then press ↻ to retry.`,
                "assistant"
            );
        }
    }

    UI.btnRefreshSelection.addEventListener("click", refreshSelection);

    function updateSelectionUI(data) {
        if (!data) return;
        if (data.readError) {
            UI.selectionText.textContent = "⚠️ Could not read workbook";
            return;
        }
        const sheetCount = data.sheets ? data.sheets.length : 1;
        const sheetInfo = sheetCount > 1 ? ` · ${sheetCount} sheets` : "";
        const sel = data.selection && data.selection.address ? ` · sel ${data.selection.address}` : "";
        const trunc = data.truncated ? " · truncated" : "";
        const warn = data.readWarnings ? " · ⚠️ unreadable sheet" : "";
        UI.selectionText.textContent = `📊 ${data.activeSheet || data.sheetName}${sheetInfo} (${data.rowCount}×${data.columnCount})${sel}${trunc}${warn}`;
    }

    // Chat Logic
    UI.btnSend.addEventListener("click", handleSend);
    UI.txtPrompt.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            handleSend();
        }
    });

    let sending = false;

    async function handleSend() {
        if (sending) return;
        const prompt = UI.txtPrompt.value.trim();
        if (!prompt) return;
        sending = true;
        UI.btnSend.disabled = true;
        UI.btnSend.style.opacity = "0.5";

        // Clear input
        UI.txtPrompt.value = "";
        UI.txtPrompt.style.height = "auto";

        // Dismiss the welcome card on first send to free chat space
        const welcome = UI.chatMessages.querySelector(".system-welcome");
        if (welcome) welcome.remove();

        // Add User Message
        appendMessage(prompt, "user");

        // Bug 23: Refresh active-sheet context right before packaging the request
        // to guarantee freshness — selection can change between user typing and sending.
        try {
            if (window.ExcelBridge && window.ExcelBridge.isOfficeInitialized) {
                const freshData = await window.ExcelBridge.getActiveSheetData();
                if (freshData) activeSelectionContext = freshData;
            }
        } catch (_) { /* non-fatal: use cached context */ }

        // Prepare request
        // Bug 1: Include conversationId so the server isolates sessions correctly.
        const requestPayload = {
            provider: UI.selectProvider.value,
            variant: UI.selectEffort && !UI.selectEffort.disabled ? (UI.selectEffort.value || null) : null,
            prompt: prompt,
            context: activeSelectionContext,
            conversationId: _conversationId
        };

        // Add loading assistant message
        const loadingId = appendMessage("Thinking...", "assistant");
        
        try {
            // NOTE: bridge_server.py is running on localhost:3000
            const controller = new AbortController();
            const timer = setTimeout(() => controller.abort(), 180000);
            let response;
            try {
                response = await fetch("http://localhost:3000/api/chat", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify(requestPayload),
                    signal: controller.signal
                });
            } finally {
                clearTimeout(timer);
            }

            if (!response.ok) {
                throw new Error(`Server responded with ${response.status}`);
            }

            const data = await response.json();
            
            // Update message with response text
            const msgBody = document.querySelector('#' + loadingId + ' .message-body');
            if (msgBody) {
                msgBody.innerHTML = renderAssistantMarkdown(data.response);
            }

            // Execute Excel Actions if any and auto-apply is on
            if (data.actions && data.actions.length > 0) {
                if (UI.chkAutoApply.checked) {
                    const result = await window.ExcelBridge.executeActions(data.actions);
                    console.log("[ExcelBridge] outcomes:", result.details);
                    recordDebugRun(result, data, requestPayload);
                    if (result.success) {
                        appendMessage(`✅ ${result.message}`, "assistant");
                    } else {
                        appendMessage(`❌ ${result.message} Open ⚙️ Settings → Debug logs for the per-action breakdown.`, "assistant");
                    }
                } else {
                    // Bug 26: Show action preview before the Apply button.
                    const previewHtml = window.ExcelBridge
                        ? window.ExcelBridge.buildActionPreviewHtml(data.actions) : "";

                    // Bug 25: Capture the active sheet name at preparation time.
                    const preparedOnSheet = (activeSelectionContext && activeSelectionContext.activeSheet) || null;

                    const applyBtnId = 'apply-' + Date.now();
                    const applyMsg = `I have prepared ${data.actions.length} update(s) for your sheet.${previewHtml ? "<br>" + previewHtml : ""}<br><button id="${applyBtnId}" class="chip" style="margin-top:8px;">Apply Changes</button>`;
                    appendMessage(applyMsg, "assistant");

                    document.getElementById(applyBtnId).addEventListener('click', async function() {
                        // Bug 25: Warn if active sheet changed since actions were prepared.
                        try {
                            if (preparedOnSheet && window.ExcelBridge && window.ExcelBridge.isOfficeInitialized) {
                                const current = await window.ExcelBridge.getActiveSheetData();
                                if (current && current.activeSheet && current.activeSheet !== preparedOnSheet) {
                                    const proceed = window.confirm(
                                        `⚠️ The active sheet has changed from "${preparedOnSheet}" to "${current.activeSheet}".\n\nApply changes to the NEW sheet instead?`
                                    );
                                    if (!proceed) { this.innerText = "Cancelled"; return; }
                                    activeSelectionContext = current;
                                }
                            }
                        } catch (_) { /* best-effort; proceed if check fails */ }

                        this.innerText = "Applying...";
                        this.disabled = true;
                        const result = await window.ExcelBridge.executeActions(data.actions);
                        console.log("[ExcelBridge] outcomes:", result.details);
                        recordDebugRun(result, data, requestPayload);
                        if (result.success) {
                            this.innerText = "Applied ✅";
                            this.style.background = "var(--accent-excel)";
                        } else {
                            this.innerText = "Error ❌";
                        }
                        appendMessage(
                            result.success
                                ? `✅ ${result.message}`
                                : `❌ ${result.message} Open ⚙️ Settings → Debug logs for the per-action breakdown.`,
                            "assistant"
                        );
                    });
                }
            }

            // Bug 49/51: Dead quickAction paths removed — all updates now go
            // through the excel-action block / executeActions() pipeline.

        } catch (err) {
            const msgBody = document.querySelector('#' + loadingId + ' .message-body');
            if (msgBody) {
                msgBody.textContent = (err && err.name === "AbortError")
                    ? "Timed out after 3 minutes with no reply. The free model may be busy or rate-limited — wait a minute and resend. No changes were made."
                    : `Error: Could not connect to local bridge server. Make sure bridge_server.py is running. (${err.message})`;
            }
        } finally {
            sending = false;
            UI.btnSend.disabled = false;
            UI.btnSend.style.opacity = "";
        }
    }

    // Debug runs log — lives ONLY in Settings, never inline in chat.
    // Each entry: { time, title, payload } with the last 10 kept.
    const debugRuns = [];

    // ---- Chat sessions: last 5 autosaved locally, survive pane reloads ----
    const MAX_SESSIONS = 5;
    const MAX_MSGS = 200;
    let sessions = [];
    let activeSessionId = null;
    try {
        sessions = JSON.parse(localStorage.getItem("excelAiSessions") || "[]");
        if (!Array.isArray(sessions)) sessions = [];
        activeSessionId = localStorage.getItem("excelAiActiveId") || null;
    } catch (_) { sessions = []; activeSessionId = null; }

    function getActiveSession() {
        return sessions.find(s => s.id === activeSessionId) || null;
    }

    function persistSessions() {
        try {
            const trimmed = sessions.slice(0, MAX_SESSIONS).map(s => ({
                id: s.id, title: s.title, updatedAt: s.updatedAt,
                provider: s.provider, variant: s.variant,
                messages: (s.messages || []).slice(-MAX_MSGS)
            }));
            localStorage.setItem("excelAiSessions", JSON.stringify(trimmed));
            if (activeSessionId) localStorage.setItem("excelAiActiveId", activeSessionId);
        } catch (_) {
            // Quota exceeded: drop the oldest session and retry once.
            try {
                sessions.pop();
                localStorage.setItem("excelAiSessions", JSON.stringify(sessions.slice(0, MAX_SESSIONS)));
            } catch (_) { /* give up silently */ }
        }
    }

    function snapshotDomToMessages() {
        const out = [];
        UI.chatMessages.querySelectorAll(".message").forEach(el => {
            if (el.classList.contains("system-welcome")) return;
            const role = el.classList.contains("user") ? "user" : "assistant";
            const body = el.querySelector(".message-body");
            if (!body) return;
            out.push(role === "user"
                ? { role: "user", text: body.textContent }
                : { role: "assistant", html: body.innerHTML });
        });
        return out.slice(-MAX_MSGS);
    }

    function saveActiveSession() {
        const msgs = snapshotDomToMessages();
        if (msgs.length === 0 && !getActiveSession()) return;
        let s = getActiveSession();
        if (!s) {
            s = { id: "sess-" + Date.now(), title: "New chat", updatedAt: Date.now() };
            sessions.unshift(s);
            activeSessionId = s.id;
        }
        s.messages = msgs;
        s.updatedAt = Date.now();
        s.provider = UI.selectProvider.value;
        s.variant = UI.selectEffort && !UI.selectEffort.disabled ? (UI.selectEffort.value || "") : "";
        const firstUser = msgs.find(m => m.role === "user");
        s.title = firstUser
            ? firstUser.text.slice(0, 42) + (firstUser.text.length > 42 ? "…" : "")
            : "Chat " + new Date(s.updatedAt).toLocaleTimeString();
        sessions = sessions.slice(0, MAX_SESSIONS);
        persistSessions();
    }

    function renderSessionMessages(s) {
        UI.chatMessages.innerHTML = "";
        (s.messages || []).forEach(m => {
            if (m.role === "user") appendMessage(m.text || "", "user");
            else appendMessage(m.html || "", "assistant");
        });
        UI.chatMessages.scrollTop = UI.chatMessages.scrollHeight;
    }

    function loadSession(id) {
        saveActiveSession();
        const s = sessions.find(x => x.id === id);
        if (!s) return;
        activeSessionId = id;
        if (s.provider && modelVariants[s.provider]) {
            UI.selectProvider.value = s.provider;
            populateEffort(s.provider);
            if (s.variant && Array.from(UI.selectEffort.options).some(o => o.value === s.variant)) {
                UI.selectEffort.value = s.variant;
            }
        }
        renderSessionMessages(s);
        persistSessions();
        renderHistoryList();
    }

    function newChatSession() {
        saveActiveSession();
        activeSessionId = null;
        UI.chatMessages.innerHTML = "";
        persistSessions();
        renderHistoryList();
        // Bug 1: Generate a new conversationId so the server creates a fresh
        // OpenCode session for this chat — avoids cross-chat context pollution.
        _conversationId = _generateUUID();
    }

    function renderHistoryList() {
        const box = document.getElementById("historyList");
        if (!box) return;
        box.innerHTML = "";
        if (sessions.length === 0) {
            box.innerHTML = '<div class="form-hint">No saved chats yet.</div>';
            return;
        }
        sessions.forEach(s => {
            const row = document.createElement("div");
            row.className = "history-row" + (s.id === activeSessionId ? " active" : "");
            const label = document.createElement("button");
            label.className = "history-label";
            label.title = s.title;
            label.innerHTML = `<span class="history-dot">${s.id === activeSessionId ? "● " : ""}</span><span class="history-title">${escapeHtml(s.title)}</span><span class="history-time">${escapeHtml(new Date(s.updatedAt).toLocaleString())}</span>`;
            label.addEventListener("click", () => loadSession(s.id));
            const del = document.createElement("button");
            del.className = "history-del";
            del.textContent = "×";
            del.title = "Delete this chat";
            del.setAttribute("aria-label", "Delete chat: " + s.title);
            del.addEventListener("click", (e) => {
                e.stopPropagation();
                sessions = sessions.filter(x => x.id !== s.id);
                if (activeSessionId === s.id) {
                    activeSessionId = null;
                    UI.chatMessages.innerHTML = "";
                }
                persistSessions();
                renderHistoryList();
            });
            row.appendChild(label);
            row.appendChild(del);
            box.appendChild(row);
        });
    }

    async function copyTextToClipboard(text) {
        // 1) Modern async clipboard (blocked in Excel WebView2).
        try {
            await navigator.clipboard.writeText(text);
            return true;
        } catch (_) { /* fall through to legacy copy */ }
        // 2) Legacy execCommand copy — sync in the click gesture.
        try {
            const ta = document.createElement("textarea");
            ta.value = text;
            ta.style.position = "fixed";
            ta.style.top = "0";
            ta.style.opacity = "0";
            document.body.appendChild(ta);
            ta.focus();
            ta.select();
            const ok = document.execCommand("copy");
            ta.remove();
            if (!ok) console.log("[Debug payload]", text);
            return ok;
        } catch (_) {
            console.log("[Debug payload]", text);
            return false;
        }
    }

    function recordDebugRun(result, data, requestPayload) {
        const details = (result && result.details) || [];
        const failed = details.filter(d => !d.ok).length;
        const title = failed === 0
            ? `✅ Applied ${details.length}`
            : `❌ ${details.length - failed} ok, ${failed} failed`;
        debugRuns.unshift({
            time: new Date(),
            title: title,
            payload: {
                prompt: requestPayload.prompt,
                provider: requestPayload.provider,
                variant: requestPayload.variant || null,
                actions: data.actions,
                outcomes: details
            }
        });
        if (debugRuns.length > 10) debugRuns.pop();
    }

    function renderDebugList() {
        const box = document.getElementById("debugList");
        if (!box) return;
        box.innerHTML = "";
        if (debugRuns.length === 0) {
            box.innerHTML = '<div class="form-hint">No apply runs yet this session.</div>';
            return;
        }
        debugRuns.forEach(run => {
            const row = document.createElement("div");
            row.className = "debug-run-row";
            const label = document.createElement("span");
            label.className = "debug-run-title";
            label.textContent = `${run.time.toLocaleTimeString()} — ${run.title}`;
            const btn = document.createElement("button");
            btn.className = "chip debug-copy";
            btn.textContent = "Copy";
            btn.addEventListener("click", async () => {
                const ok = await copyTextToClipboard(JSON.stringify(run.payload, null, 2));
                btn.innerText = ok ? "Copied ✅" : "Copy failed ❌";
                setTimeout(() => { btn.innerText = "Copy"; }, 2000);
            });
            row.appendChild(label);
            row.appendChild(btn);
            box.appendChild(row);
        });
    }

    function escapeHtml(s) {
        return String(s == null ? "" : s).replace(/[&<>"']/g, c => ({
            "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
        }[c]));
    }

    // Lightweight markdown rendering for AI replies:
    // - ```excel-action blocks are machine payload: stripped, replaced by
    //   a one-line note (results arrive as separate apply messages).
    // - other fenced blocks become styled <pre><code>.
    // - inline `code` and **bold** are honored; raw HTML is escaped.
    function renderAssistantMarkdown(src) {
        let text = String(src || "");
        const actionBlocks = (text.match(/```excel-action[\s\S]*?```/gi) || []).length;
        text = text.replace(/```excel-action[\s\S]*?```/gi, "");
        const fenced = [];
        text = text.replace(/```(\w*)\n?([\s\S]*?)```/g, (m, lang, code) => {
            fenced.push(`<pre><code>${escapeHtml(code.replace(/\n$/, ""))}</code></pre>`);
            return "\u0000" + (fenced.length - 1) + "\u0000";
        });
        text = escapeHtml(text);
        text = text.replace(/`([^`\n]+)`/g, "<code>$1</code>");
        text = text.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
        text = text.replace(/\n/g, "<br/>");
        text = text.replace(/\u0000(\d+)\u0000/g, (m, i) => fenced[+i]);
        if (actionBlocks > 0) {
            text = `<div class="action-note">⚙️ Spreadsheet update${actionBlocks > 1 ? "s" : ""} prepared — see results below.</div>` + text;
        }
        return text;
    }

    // Debug hook (used by automated screenshot tests + console probing).
    window.ExcelAssistant = window.ExcelAssistant || {};
    window.ExcelAssistant.render = renderAssistantMarkdown;

    /* in-chat debug rendering retired — logs live in Settings (renderDebugList) */

    function appendMessage(text, role) {
        const msgDiv = document.createElement("div");
        msgDiv.className = `message ${role}`;
        msgDiv.id = 'msg-' + Date.now();

        const label = document.createElement("span");
        label.className = "role-label";
        label.textContent = role === "user" ? "You" : "AI Assistant";
        msgDiv.appendChild(label);

        const body = document.createElement("div");
        body.className = "message-body";
        if (role === 'user') {
            body.textContent = text;
        } else {
            body.innerHTML = text; // allow basic html for assistant
        }
        msgDiv.appendChild(body);

        UI.chatMessages.appendChild(msgDiv);
        UI.chatMessages.scrollTop = UI.chatMessages.scrollHeight;
        saveActiveSession();
        return msgDiv.id;
    }

    function restoreStartupSession() {
        const s = getActiveSession();
        if (s && s.messages && s.messages.length > 0) {
            if (s.provider && modelVariants[s.provider]) {
                UI.selectProvider.value = s.provider;
                populateEffort(s.provider);
                if (s.variant && Array.from(UI.selectEffort.options).some(o => o.value === s.variant)) {
                    UI.selectEffort.value = s.variant;
                }
            }
            renderSessionMessages(s);
        }
        renderHistoryList();
    }
});

