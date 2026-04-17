// ==UserScript==
// @name         Flex 📁 Folder Copy & 📝 Quick Notes
// @namespace    fisher-flex-folder-notes
// @version      1.1.0
// @description  在 Flex Dashboard 每列案件旁加上「複製資料夾名稱」按鈕與「快速筆記」功能，支援依時間/專案瀏覽歷史筆記。
// @match        https://portal.ul.com/Dashboard*
// @grant        GM_addStyle
// @grant        GM_registerMenuCommand
// @run-at       document-end
// ==/UserScript==

(function() {
    'use strict';

    const GRID_CONTAINER_SELECTOR = '#projectDashboardGrid';
    const GRID_HEADER_SELECTOR = '.k-grid-header';
    const GRID_CONTENT_SELECTOR = '.k-grid-content';
    const GRID_ROW_SELECTOR = 'tr:not(.k-grouping-row):not(.k-filter-row):not(.k-grid-norecords)';

    const FOLDER_COLUMNS = ["File No", "Project Number", "Order Lines", "Company Name", "Date Project Created"];
    const NOTE_CONTEXT_COLUMNS = ["Project Number", "Project Name", "File No"];

    const STORAGE_KEY = 'flex_folder_notes_v1';
    const FLASH_MS = 300;
    const KENDO_RETRY_MS = 300;
    const GRID_OBSERVER_RETRY_MS = 500;

    let notesOverlayEl = null;
    let notesPanelEl = null;
    let gridObserver = null;
    let uuidCounter = 0;

    GM_addStyle(`
        /* ── 側邊收合抽屜 ── */
        .ffn-btn-drawer{
            position:absolute;
            left:0;
            top:50%;
            transform:translateY(-50%);
            width:5px;
            height:26px;
            background:#94a3b8;
            border-radius:0 4px 4px 0;
            overflow:hidden;
            display:flex;
            align-items:center;
            gap:2px;
            padding:0;
            transition:width .18s ease,background .18s ease,padding .18s ease,border-radius .18s ease;
            z-index:3;
            cursor:pointer;
            box-shadow:1px 0 4px rgba(0,0,0,.15);
        }
        .ffn-btn-drawer:hover{
            width:60px;
            background:#f1f5f9;
            border:1px solid #d1d5db;
            border-radius:0 10px 10px 0;
            padding:0 4px;
        }

        /* ── 列按鈕（放在抽屜內，平時隱藏） ── */
        .ffn-row-btn{
            position:static;
            transform:none;
            width:24px;
            height:24px;
            border-radius:50%;
            border:1px solid #d1d5db;
            background:#fff;
            cursor:pointer;
            box-shadow:0 1px 3px rgba(0,0,0,.12);
            display:flex;
            align-items:center;
            justify-content:center;
            line-height:1;
            padding:0;
            flex-shrink:0;
            opacity:0;
            pointer-events:none;
            transition:opacity .1s .05s;
        }
        .ffn-btn-drawer:hover .ffn-row-btn{
            opacity:1;
            pointer-events:auto;
        }
        .ffn-row-btn:hover{ background:#f3f4f6; }
        .ffn-copy-flash{ background:#d1fae5 !important; }

        .ffn-note-editor-backdrop,
        .ffn-notes-overlay{
            position:fixed;
            inset:0;
            background:rgba(0,0,0,.35);
            z-index:99997;
            display:flex;
            align-items:center;
            justify-content:center;
        }
        .ffn-note-editor,
        .ffn-notes-panel{
            background:#fff;
            border:1px solid #d1d5db;
            border-radius:8px;
            box-shadow:0 12px 24px rgba(0,0,0,.2);
            color:#111827;
            font-family:Segoe UI,Arial,sans-serif;
        }
        .ffn-note-editor{ width:520px; max-width:92vw; padding:12px; z-index:99998; }
        .ffn-note-editor h3{ margin:0 0 8px; font-size:16px; }
        .ffn-note-info{ font-size:12px; color:#374151; background:#f9fafb; border:1px solid #e5e7eb; border-radius:4px; padding:6px; margin-bottom:8px; }
        .ffn-note-editor textarea{ width:100%; min-height:120px; resize:vertical; box-sizing:border-box; border:1px solid #d1d5db; border-radius:6px; padding:8px; }
        .ffn-actions{ margin-top:8px; display:flex; gap:8px; justify-content:flex-end; }

        /* ── 筆記瀏覽面板：80% 螢幕寬 ── */
        .ffn-notes-panel{ width:80vw; max-height:90vh; display:flex; flex-direction:column; }
        .ffn-notes-header{ display:flex; align-items:center; justify-content:space-between; padding:10px 12px; border-bottom:1px solid #e5e7eb; }
        .ffn-notes-header h3{ margin:0; font-size:16px; }
        .ffn-notes-controls{ padding:10px 12px; border-bottom:1px solid #f3f4f6; display:grid; gap:8px; grid-template-columns:repeat(6,minmax(120px,1fr)); }
        .ffn-notes-controls input,
        .ffn-notes-controls select,
        .ffn-note-editor button,
        .ffn-notes-header button,
        .ffn-note-card button{ border:1px solid #d1d5db; border-radius:6px; background:#fff; padding:6px 8px; font-size:12px; }
        .ffn-notes-list{ overflow:auto; padding:12px; display:flex; flex-direction:column; gap:8px; }
        .ffn-note-card{ border:1px solid #e5e7eb; border-radius:8px; padding:8px; background:#fff; }
        .ffn-note-title{ font-size:13px; margin-bottom:6px; white-space:pre-wrap; }
        .ffn-note-meta{ font-size:12px; color:#374151; margin-bottom:6px; }
        .ffn-note-time{ font-size:12px; color:#6b7280; }
        .ffn-note-card .ffn-actions{ margin-top:6px; justify-content:flex-start; }
        .ffn-empty{ text-align:center; color:#6b7280; padding:16px; }
    `);

    function analyzeGridHeaders(columnNames = FOLDER_COLUMNS) {
        const gridContainer = document.querySelector(GRID_CONTAINER_SELECTOR);
        if (!gridContainer) return null;
        const headerDiv = gridContainer.querySelector(GRID_HEADER_SELECTOR);
        if (!headerDiv) return null;
        const headerRows = Array.from(headerDiv.querySelectorAll('tr'))
            .filter(tr => tr.querySelectorAll('th').length > 0);
        const headerTr = headerRows[headerRows.length - 1] || null;
        if (!headerTr) return null;
        const thElements = Array.from(headerTr.querySelectorAll('th'));
        const columnIndexMap = new Map();

        columnNames.forEach(name => {
            for (let i = 0; i < thElements.length; i++) {
                const th = thElements[i];
                const textContent = th.textContent.trim().toLowerCase();
                const titleAttribute = th.getAttribute('title');
                const dataTitleAttribute = th.getAttribute('data-title');
                const target = name.toLowerCase();
                if (
                    textContent.includes(target) ||
                    (titleAttribute && titleAttribute.toLowerCase().includes(target)) ||
                    (dataTitleAttribute && dataTitleAttribute.toLowerCase().includes(target))
                ) {
                    columnIndexMap.set(name, i);
                    break;
                }
            }
        });

        return columnIndexMap.size > 0 ? columnIndexMap : null;
    }

    function extractFromRow(rowElement, columnIndexMap) {
        const cells = rowElement.children;
        const data = {};
        columnIndexMap.forEach((colIndex, columnName) => {
            data[columnName] = (colIndex < cells.length) ? cells[colIndex].textContent.trim() : '';
        });
        return data;
    }

    function sanitizeFolderPart(value) {
        return (value || '').replace(/[<>:"/\\|?*]/g, '_').trim();
    }

    function normalizeDateToYMD(input) {
        if (!input) return '';
        const direct = new Date(input);
        if (!Number.isNaN(direct.getTime())) {
            return `${direct.getFullYear()}-${String(direct.getMonth() + 1).padStart(2, '0')}-${String(direct.getDate()).padStart(2, '0')}`;
        }

        const mdy = String(input).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (mdy) {
            const d = new Date(Number(mdy[3]), Number(mdy[1]) - 1, Number(mdy[2]));
            if (!Number.isNaN(d.getTime())) return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
        }

        const ymd = String(input).match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
        if (ymd) {
            return `${ymd[1]}-${String(ymd[2]).padStart(2, '0')}-${String(ymd[3]).padStart(2, '0')}`;
        }

        return sanitizeFolderPart(String(input));
    }

    function buildFolderString(rowData) {
        const fileNo = sanitizeFolderPart(rowData['File No']);
        const projectNumber = sanitizeFolderPart(rowData['Project Number']);
        const orderLines = sanitizeFolderPart(rowData['Order Lines']);
        const companyName = sanitizeFolderPart(rowData['Company Name']).slice(0, 10);
        const createdDate = normalizeDateToYMD(rowData['Date Project Created']);
        return `${fileNo}-${projectNumber}-${orderLines}-${companyName}-${createdDate}`;
    }

    async function copyText(text) {
        if (!text) return false;
        try {
            await navigator.clipboard.writeText(text);
            return true;
        } catch (_e) {
            const ta = document.createElement('textarea');
            ta.value = text;
            ta.style.position = 'fixed';
            ta.style.opacity = '0';
            document.body.appendChild(ta);
            ta.select();
            const ok = document.execCommand('copy');
            ta.remove();
            return ok;
        }
    }

    function loadStore() {
        try {
            const parsed = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
            const notes = Array.isArray(parsed.notes) ? parsed.notes : [];
            return { notes };
        } catch (_e) {
            return { notes: [] };
        }
    }

    function saveStore(store) {
        localStorage.setItem(STORAGE_KEY, JSON.stringify({ notes: store.notes || [] }));
    }

    function createUUID() {
        if (window.crypto && typeof window.crypto.randomUUID === 'function') return window.crypto.randomUUID();
        uuidCounter += 1;
        const randomChunk = window.crypto && typeof window.crypto.getRandomValues === 'function'
            ? Array.from(window.crypto.getRandomValues(new Uint32Array(2))).map(n => n.toString(16)).join('')
            : `${Math.random().toString(16).slice(2, 10)}${Math.random().toString(16).slice(2, 10)}`;
        return `ffn-${Date.now()}-${uuidCounter}-${randomChunk}`;
    }

    function formatDateTime(value) {
        const d = new Date(value);
        if (Number.isNaN(d.getTime())) return '';
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')} ${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
    }

    function getRowData(rowElement) {
        const targetColumns = [...new Set([...FOLDER_COLUMNS, ...NOTE_CONTEXT_COLUMNS])];
        const colMap = analyzeGridHeaders(targetColumns);
        if (!colMap) return null;
        return extractFromRow(rowElement, colMap);
    }

    function closeNoteEditor(editorRoot) {
        if (editorRoot && editorRoot.parentNode) editorRoot.parentNode.remove();
    }

    function openNoteEditor(options) {
        const existing = document.querySelector('.ffn-note-editor-backdrop');
        if (existing) existing.remove();

        const note = options.note || null;
        const rowData = options.rowData || {};
        const now = new Date().toISOString();

        const projectNumber = note ? note.projectNumber : (rowData['Project Number'] || '');
        const projectName = note ? note.projectName : (rowData['Project Name'] || '');
        const fileNo = note ? note.fileNo : (rowData['File No'] || '');

        const backdrop = document.createElement('div');
        backdrop.className = 'ffn-note-editor-backdrop';

        const box = document.createElement('div');
        box.className = 'ffn-note-editor';

        const title = document.createElement('h3');
        title.textContent = note ? '編輯筆記' : '新增快速筆記';

        const info = document.createElement('div');
        info.className = 'ffn-note-info';
        info.textContent = `📁 ${projectNumber || '(未指定專案)'} - ${fileNo || '-'} - ${projectName || '-'}`;

        const textarea = document.createElement('textarea');
        textarea.placeholder = '請輸入筆記內容...';
        textarea.value = note ? (note.text || '') : (options.initialText || '');

        const actions = document.createElement('div');
        actions.className = 'ffn-actions';

        const cancelBtn = document.createElement('button');
        cancelBtn.textContent = '取消';

        const saveBtn = document.createElement('button');
        saveBtn.textContent = '儲存';

        cancelBtn.addEventListener('click', () => closeNoteEditor(box));
        backdrop.addEventListener('click', (e) => {
            if (e.target === backdrop) closeNoteEditor(box);
        });

        saveBtn.addEventListener('click', () => {
            const text = textarea.value.trim();
            if (!text) {
                alert('請輸入筆記內容');
                return;
            }

            const store = loadStore();
            if (note) {
                const idx = store.notes.findIndex(n => n.id === note.id);
                if (idx >= 0) {
                    store.notes[idx] = {
                        ...store.notes[idx],
                        projectNumber,
                        projectName,
                        fileNo,
                        text,
                        updatedAt: now
                    };
                }
            } else {
                store.notes.push({
                    id: createUUID(),
                    projectNumber,
                    projectName,
                    fileNo,
                    text,
                    createdAt: now,
                    updatedAt: now
                });
            }
            saveStore(store);
            closeNoteEditor(box);
            renderNotesList();
        });

        actions.append(cancelBtn, saveBtn);
        box.append(title, info, textarea, actions);
        backdrop.appendChild(box);
        document.body.appendChild(backdrop);
    }

    function getFilteredSortedNotes() {
        const store = loadStore();
        const notes = [...store.notes];
        if (!notesPanelEl) return notes;

        const sortBy = notesPanelEl.querySelector('#ffn-sort-by').value;
        const order = notesPanelEl.querySelector('#ffn-order').value;
        const projectKeyword = notesPanelEl.querySelector('#ffn-project-filter').value.trim().toLowerCase();
        const fromDate = notesPanelEl.querySelector('#ffn-date-from').value;
        const toDate = notesPanelEl.querySelector('#ffn-date-to').value;
        const keyword = notesPanelEl.querySelector('#ffn-keyword').value.trim().toLowerCase();

        const filtered = notes.filter(note => {
            const created = new Date(note.createdAt);
            if (fromDate) {
                const from = new Date(`${fromDate}T00:00:00`);
                if (created < from) return false;
            }
            if (toDate) {
                const to = new Date(`${toDate}T23:59:59`);
                if (created > to) return false;
            }
            if (projectKeyword && !(note.projectNumber || '').toLowerCase().includes(projectKeyword)) return false;
            if (keyword) {
                const searchableText = `${note.text || ''} ${(note.projectName || '')} ${(note.projectNumber || '')}`.toLowerCase();
                if (!searchableText.includes(keyword)) return false;
            }
            return true;
        });

        filtered.sort((a, b) => {
            let va = '';
            let vb = '';
            if (sortBy === 'updatedAt') {
                va = a.updatedAt || '';
                vb = b.updatedAt || '';
            } else if (sortBy === 'createdAt') {
                va = a.createdAt || '';
                vb = b.createdAt || '';
            } else {
                va = a.projectName || '';
                vb = b.projectName || '';
            }
            if (va < vb) return -1;
            if (va > vb) return 1;
            return 0;
        });

        if (order === 'desc') filtered.reverse();
        return filtered;
    }

    function renderNotesList() {
        if (!notesPanelEl) return;
        const list = notesPanelEl.querySelector('.ffn-notes-list');
        list.innerHTML = '';

        const notes = getFilteredSortedNotes();
        if (!notes.length) {
            const empty = document.createElement('div');
            empty.className = 'ffn-empty';
            empty.textContent = '目前沒有符合條件的筆記';
            list.appendChild(empty);
            return;
        }

        notes.forEach(note => {
            const card = document.createElement('div');
            card.className = 'ffn-note-card';

            const preview = document.createElement('div');
            preview.className = 'ffn-note-title';
            const shortText = (note.text || '').length > 120 ? `${note.text.slice(0, 120)}...` : (note.text || '');
            preview.textContent = `📝 ${shortText}`;

            const project = document.createElement('div');
            project.className = 'ffn-note-meta';
            project.textContent = `📁 ${note.projectNumber || '-'} - ${note.fileNo || '-'} - ${note.projectName || '-'}`;

            const time = document.createElement('div');
            time.className = 'ffn-note-time';
            time.textContent = `🕐 建立: ${formatDateTime(note.createdAt)} | 更新: ${formatDateTime(note.updatedAt)}`;

            const actions = document.createElement('div');
            actions.className = 'ffn-actions';

            const editBtn = document.createElement('button');
            editBtn.textContent = '✏️ 編輯';
            editBtn.addEventListener('click', () => openNoteEditor({ note }));

            const delBtn = document.createElement('button');
            delBtn.textContent = '🗑️ 刪除';
            delBtn.addEventListener('click', () => {
                if (!confirm('確定要刪除此筆記？')) return;
                const store = loadStore();
                store.notes = store.notes.filter(n => n.id !== note.id);
                saveStore(store);
                renderNotesList();
            });

            actions.append(editBtn, delBtn);
            card.append(preview, project, time, actions);
            list.appendChild(card);
        });
    }

    function createNotesPanel() {
        if (notesOverlayEl) return notesOverlayEl;

        notesOverlayEl = document.createElement('div');
        notesOverlayEl.className = 'ffn-notes-overlay';

        notesPanelEl = document.createElement('div');
        notesPanelEl.className = 'ffn-notes-panel';

        notesPanelEl.innerHTML = `
            <div class="ffn-notes-header">
                <h3>📝 Quick Notes History</h3>
                <div class="ffn-actions" style="margin:0;">
                    <button type="button" id="ffn-today-log">📅 今日紀錄</button>
                    <button type="button" id="ffn-close-panel">關閉</button>
                </div>
            </div>
            <div class="ffn-notes-controls">
                <select id="ffn-sort-by">
                    <option value="updatedAt">最近更新</option>
                    <option value="createdAt">建立時間</option>
                    <option value="projectName">專案名稱</option>
                </select>
                <select id="ffn-order">
                    <option value="desc">新→舊</option>
                    <option value="asc">舊→新</option>
                </select>
                <input id="ffn-project-filter" type="text" placeholder="專案編號篩選">
                <input id="ffn-date-from" type="date" title="開始日期">
                <input id="ffn-date-to" type="date" title="結束日期">
                <input id="ffn-keyword" type="text" placeholder="關鍵字搜尋">
            </div>
            <div class="ffn-notes-list"></div>
        `;

        notesOverlayEl.appendChild(notesPanelEl);
        notesOverlayEl.addEventListener('click', (e) => {
            if (e.target === notesOverlayEl) notesOverlayEl.remove();
        });

        notesPanelEl.querySelector('#ffn-close-panel').addEventListener('click', () => notesOverlayEl.remove());
        notesPanelEl.querySelector('#ffn-today-log').addEventListener('click', () => {
            const dateStr = normalizeDateToYMD(new Date());
            openNoteEditor({ rowData: {}, initialText: `${dateStr} 每日紀錄` });
        });

        notesPanelEl.querySelectorAll('input,select').forEach(el => el.addEventListener('input', renderNotesList));

        return notesOverlayEl;
    }

    function openNotesPanel() {
        const panel = createNotesPanel();
        document.body.appendChild(panel);
        renderNotesList();
    }

    async function handleCopyButtonClick(button, rowElement) {
        const rowData = getRowData(rowElement);
        if (!rowData) {
            alert('找不到表頭資料，請確認 Grid 已載入');
            return;
        }
        const folderName = buildFolderString(rowData);
        const ok = await copyText(folderName);
        if (ok) {
            button.classList.add('ffn-copy-flash');
            setTimeout(() => button.classList.remove('ffn-copy-flash'), FLASH_MS);
        } else {
            alert('複製失敗，請稍後再試');
        }
    }

    function attachButtons() {
        const rows = document.querySelectorAll(`${GRID_CONTAINER_SELECTOR} ${GRID_CONTENT_SELECTOR} ${GRID_ROW_SELECTOR}`);
        rows.forEach(row => {
            const firstCell = row.querySelector('td');
            if (!firstCell) return;
            if (!firstCell.style.position) firstCell.style.position = 'relative';
            firstCell.style.overflow = 'visible';

            // 已存在就跳過
            if (row.querySelector('.ffn-btn-drawer')) return;

            // 建立收合抽屜
            const drawer = document.createElement('div');
            drawer.className = 'ffn-btn-drawer';
            drawer.title = '展開操作按鈕';

            const copyBtn = document.createElement('button');
            copyBtn.type = 'button';
            copyBtn.className = 'ffn-row-btn ffn-rowcopy';
            copyBtn.title = '複製資料夾名稱';
            copyBtn.textContent = '📁';
            copyBtn.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                handleCopyButtonClick(copyBtn, row);
            });

            const noteBtn = document.createElement('button');
            noteBtn.type = 'button';
            noteBtn.className = 'ffn-row-btn ffn-rownote';
            noteBtn.title = '快速筆記';
            noteBtn.textContent = '📝';
            noteBtn.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                const rowData = getRowData(row) || {};
                openNoteEditor({ rowData });
            });

            drawer.append(copyBtn, noteBtn);
            firstCell.appendChild(drawer);
        });
    }

    function observeGridRows() {
        const tbody = document.querySelector(`${GRID_CONTAINER_SELECTOR} ${GRID_CONTENT_SELECTOR} tbody`);
        if (!tbody) {
            setTimeout(observeGridRows, GRID_OBSERVER_RETRY_MS);
            return;
        }

        if (gridObserver) gridObserver.disconnect();
        gridObserver = new MutationObserver(() => attachButtons());
        gridObserver.observe(tbody, { childList: true, subtree: true });
        attachButtons();
    }

    (function bindKendo() {
        const $ = window.jQuery;
        if (!$) return setTimeout(bindKendo, KENDO_RETRY_MS);
        const grid = $(GRID_CONTAINER_SELECTOR).data('kendoGrid');
        if (!grid) return setTimeout(bindKendo, KENDO_RETRY_MS);
        grid.bind('dataBound', () => { attachButtons(); observeGridRows(); });
    })();

    GM_registerMenuCommand('📝 開啟快速筆記面板', openNotesPanel);

    observeGridRows();
    attachButtons();
})();
