// ==UserScript==
// @name         Flex 📁 Folder Copy & 📝 Quick Notes
// @namespace    fisher-flex-folder-notes
// @version      1.2.1
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
    const DRAFT_KEY_PREFIX = `${STORAGE_KEY}_draft__`;
    const DRAFT_DEBOUNCE_MS = 500;
    const DRAFT_KEY_EMPTY_PLACEHOLDER = '_';
    const MAX_NOTE_TAGS = 20;
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
            width:6px;
            height:28px;
            background:#d5deea;
            border:1px solid rgba(148,163,184,.45);
            border-left:none;
            border-radius:0 8px 8px 0;
            overflow:hidden;
            display:flex;
            align-items:center;
            gap:4px;
            padding:0;
            transition:width .24s ease,background .2s ease,padding .2s ease,border-radius .2s ease,box-shadow .2s ease,transform .2s ease,border-color .2s ease;
            z-index:3;
            cursor:pointer;
            box-shadow:0 4px 12px rgba(15,23,42,.15);
        }
        .ffn-btn-drawer:hover{
            width:68px;
            background:rgba(247,249,252,.96);
            border-color:rgba(203,213,225,.95);
            border-radius:0 12px 12px 0;
            padding:0 6px;
            transform:translateY(-50%) translateX(1px);
            box-shadow:0 10px 24px rgba(15,23,42,.16);
        }

        /* ── 列按鈕（放在抽屜內，平時隱藏） ── */
        .ffn-row-btn{
            position:static;
            transform:none;
            width:26px;
            height:26px;
            border-radius:50%;
            border:1px solid #d4dce5;
            background:rgba(255,255,255,.97);
            cursor:pointer;
            box-shadow:0 1px 4px rgba(15,23,42,.16);
            display:flex;
            align-items:center;
            justify-content:center;
            line-height:1;
            padding:0;
            flex-shrink:0;
            opacity:0;
            pointer-events:none;
            transform:translateY(2px) scale(.96);
            transition:opacity .16s .04s,transform .18s ease,background .15s,border-color .15s,box-shadow .15s;
        }
        .ffn-btn-drawer:hover .ffn-row-btn{
            opacity:1;
            pointer-events:auto;
            transform:translateY(0) scale(1);
        }
        .ffn-row-btn:hover{
            background:#f5f7fb;
            border-color:#bcc8d6;
            box-shadow:0 4px 10px rgba(15,23,42,.18);
        }
        .ffn-row-btn:focus-visible{
            outline:none;
            border-color:#0071e3;
            box-shadow:0 0 0 4px rgba(0,113,227,.2);
        }
        .ffn-copy-flash{ background:#d1fae5 !important; }

        /* ── 筆記編輯對話框 ── */
        .ffn-note-editor-backdrop{
            position:fixed;
            inset:0;
            background:rgba(20,26,34,.36);
            backdrop-filter:blur(12px) saturate(140%);
            -webkit-backdrop-filter:blur(12px) saturate(140%);
            z-index:99999;
            display:flex;
            align-items:center;
            justify-content:center;
            animation:ffn-fade-in .22s ease-out;
        }
        .ffn-note-editor{
            background:rgba(255,255,255,.88);
            border:1px solid rgba(255,255,255,.7);
            border-radius:16px;
            box-shadow:0 24px 56px rgba(15,23,42,.24);
            width:560px;
            max-width:94vw;
            padding:24px 26px;
            color:#1d1d1f;
            font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;
            animation:ffn-slide-up .24s cubic-bezier(.2,.8,.2,1);
        }
        .ffn-note-editor h3{
            margin:0 0 14px;
            font-size:20px;
            font-weight:600;
            letter-spacing:.01em;
            color:#1d1d1f;
        }
        .ffn-note-info{
            font-size:12px;
            color:#556170;
            background:rgba(246,248,251,.9);
            border:1px solid rgba(210,218,229,.86);
            border-radius:10px;
            padding:10px 12px;
            margin-bottom:14px;
            line-height:1.5;
        }
        .ffn-note-editor textarea{
            width:100%;
            min-height:152px;
            resize:vertical;
            box-sizing:border-box;
            border:1px solid #d4dce6;
            border-radius:12px;
            background:rgba(255,255,255,.94);
            padding:12px 14px;
            font-size:14px;
            line-height:1.7;
            color:#1f2937;
            outline:none;
            transition:border-color .15s,box-shadow .15s,background .15s;
            font-family:inherit;
        }
        .ffn-note-editor textarea:focus{
            border-color:#0071e3;
            box-shadow:0 0 0 4px rgba(0,113,227,.16);
            background:#fff;
        }
        .ffn-note-tags-input{
            margin-top:12px;
            width:100%;
            box-sizing:border-box;
            border:1px solid #d4dce6;
            border-radius:12px;
            background:rgba(255,255,255,.94);
            padding:10px 12px;
            font-size:13px;
            color:#1f2937;
            font-family:inherit;
            outline:none;
            transition:border-color .15s,box-shadow .15s,background .15s;
        }
        .ffn-note-tags-input:focus{
            border-color:#0071e3;
            box-shadow:0 0 0 4px rgba(0,113,227,.16);
            background:#fff;
        }
        .ffn-actions{
            margin-top:16px;
            display:flex;
            gap:10px;
            justify-content:flex-end;
        }
        .ffn-note-editor button{
            padding:9px 18px;
            border-radius:10px;
            border:1px solid #d3dbe5;
            background:#fff;
            cursor:pointer;
            font-size:13px;
            font-weight:600;
            color:#374151;
            transition:background .15s,border-color .15s,color .15s,box-shadow .15s,transform .15s;
        }
        .ffn-note-editor button:hover{
            background:#f5f8fc;
            border-color:#c2cddd;
            transform:translateY(-1px);
        }
        .ffn-note-editor button:focus-visible{
            outline:none;
            border-color:#0071e3;
            box-shadow:0 0 0 4px rgba(0,113,227,.16);
        }
        .ffn-note-editor button.ffn-primary{
            background:#0071e3;
            color:#fff;
            border-color:#0071e3;
            box-shadow:0 8px 18px rgba(0,113,227,.28);
        }
        .ffn-note-editor button.ffn-primary:hover{
            background:#0066cc;
            border-color:#0066cc;
            box-shadow:0 10px 20px rgba(0,102,204,.3);
        }

        /* ── Memos 風格筆記面板 ── */
        .ffn-notes-overlay{
            position:fixed;
            inset:0;
            background:rgba(20,26,34,.36);
            backdrop-filter:blur(12px) saturate(140%);
            -webkit-backdrop-filter:blur(12px) saturate(140%);
            z-index:99997;
            display:flex;
            align-items:center;
            justify-content:center;
            animation:ffn-fade-in .22s ease-out;
        }
        .ffn-notes-panel{
            width:92vw;
            max-width:1120px;
            height:86vh;
            display:flex;
            border-radius:16px;
            overflow:hidden;
            background:rgba(255,255,255,.9);
            border:1px solid rgba(255,255,255,.7);
            box-shadow:0 30px 72px rgba(15,23,42,.24);
            color:#1d1d1f;
            font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;
            animation:ffn-slide-up .24s cubic-bezier(.2,.8,.2,1);
        }

        /* Sidebar */
        .ffn-sidebar{
            width:252px;
            flex-shrink:0;
            background:linear-gradient(180deg,rgba(245,248,252,.97) 0%,rgba(240,244,249,.95) 100%);
            border-right:1px solid rgba(216,224,235,.95);
            display:flex;
            flex-direction:column;
            padding:24px 18px;
            gap:16px;
            overflow-y:auto;
            box-sizing:border-box;
        }
        .ffn-sidebar-title{
            font-size:18px;
            font-weight:600;
            color:#1d1d1f;
            margin:0;
            display:flex;
            align-items:center;
            gap:8px;
            letter-spacing:.01em;
        }
        .ffn-sidebar-new-btn{
            width:100%;
            padding:10px 14px;
            border-radius:10px;
            background:#0071e3;
            color:#fff;
            border:none;
            font-size:13px;
            font-weight:600;
            cursor:pointer;
            display:flex;
            align-items:center;
            gap:6px;
            justify-content:center;
            box-shadow:0 10px 20px rgba(0,113,227,.2);
            transition:background .15s,transform .15s,box-shadow .15s;
        }
        .ffn-sidebar-new-btn:hover{
            background:#0066cc;
            transform:translateY(-1px);
            box-shadow:0 12px 22px rgba(0,102,204,.24);
        }
        .ffn-sidebar-new-btn:focus-visible{
            outline:none;
            box-shadow:0 0 0 4px rgba(0,113,227,.22);
        }
        .ffn-sidebar-section{
            display:flex;
            flex-direction:column;
            gap:8px;
        }
        .ffn-sidebar-label{
            font-size:11px;
            font-weight:600;
            color:#7f8da1;
            text-transform:uppercase;
            letter-spacing:.08em;
        }
        .ffn-sidebar input,
        .ffn-sidebar select{
            width:100%;
            padding:9px 11px;
            border:1px solid #d1dbe6;
            border-radius:10px;
            background:rgba(255,255,255,.92);
            font-size:13px;
            color:#1f2937;
            box-sizing:border-box;
            outline:none;
            transition:border-color .15s,box-shadow .15s,background .15s;
        }
        .ffn-sidebar input:focus,
        .ffn-sidebar select:focus{
            border-color:#0071e3;
            box-shadow:0 0 0 4px rgba(0,113,227,.14);
            background:#fff;
        }

        /* Main feed */
        .ffn-main{
            flex:1;
            display:flex;
            flex-direction:column;
            overflow:hidden;
            background:rgba(252,253,255,.9);
        }
        .ffn-main-header{
            display:flex;
            align-items:center;
            justify-content:space-between;
            padding:16px 24px;
            border-bottom:1px solid rgba(221,228,238,.85);
            flex-shrink:0;
            background:rgba(255,255,255,.75);
        }
        .ffn-main-header-title{
            font-size:14px;
            font-weight:600;
            color:#5d6a79;
            letter-spacing:.01em;
        }
        .ffn-close-btn{
            width:32px;
            height:32px;
            border-radius:50%;
            border:none;
            background:#eef2f7;
            cursor:pointer;
            font-size:15px;
            display:flex;
            align-items:center;
            justify-content:center;
            color:#6b7787;
            transition:background .15s,color .15s,transform .15s,box-shadow .15s;
            line-height:1;
        }
        .ffn-close-btn:hover{
            background:#e2e8f0;
            color:#1d1d1f;
            transform:translateY(-1px);
        }
        .ffn-close-btn:focus-visible{
            outline:none;
            box-shadow:0 0 0 4px rgba(0,113,227,.16);
        }
        .ffn-notes-list{
            flex:1;
            overflow-y:auto;
            padding:16px 24px 24px;
            background:linear-gradient(180deg,rgba(255,255,255,.7) 0%,rgba(249,251,255,.62) 100%);
        }

        /* Date group */
        .ffn-date-group{
            font-size:11px;
            font-weight:600;
            color:#8b97a9;
            text-transform:uppercase;
            letter-spacing:.06em;
            margin:20px 0 10px;
            padding-left:2px;
        }
        .ffn-date-group:first-child{ margin-top:4px; }

        /* Memo card */
        .ffn-memo-card{
            border:1px solid rgba(214,223,234,.78);
            border-radius:14px;
            padding:16px 18px;
            margin-bottom:10px;
            background:rgba(255,255,255,.95);
            box-shadow:0 2px 8px rgba(15,23,42,.05);
            transition:box-shadow .2s,border-color .2s,transform .2s,background .2s;
        }
        .ffn-memo-card:hover{
            box-shadow:0 12px 28px rgba(15,23,42,.12);
            border-color:rgba(189,203,220,.9);
            background:#fff;
            transform:translateY(-1px);
        }
        .ffn-memo-content{
            font-size:14px;
            line-height:1.7;
            color:#1f2937;
            white-space:normal;
            word-break:break-word;
            margin-bottom:12px;
        }
        .ffn-memo-content > :first-child{ margin-top:0; }
        .ffn-memo-content > :last-child{ margin-bottom:0; }
        .ffn-memo-content p{ margin:.5em 0; }
        .ffn-memo-content h1,.ffn-memo-content h2,.ffn-memo-content h3,
        .ffn-memo-content h4,.ffn-memo-content h5,.ffn-memo-content h6{
            margin:.7em 0 .4em;
            line-height:1.35;
        }
        .ffn-memo-content blockquote{
            margin:.6em 0;
            padding:.1em .8em;
            border-left:3px solid #99c4f5;
            color:#3f4a59;
            background:#f4f8ff;
        }
        .ffn-memo-content ul,.ffn-memo-content ol{
            margin:.5em 0;
            padding-left:1.3em;
        }
        .ffn-memo-content code{
            font-family:ui-monospace,SFMono-Regular,Menlo,Monaco,Consolas,'Liberation Mono','Courier New',monospace;
            background:#f2f5fa;
            border-radius:6px;
            padding:2px 6px;
            font-size:12px;
        }
        .ffn-memo-content pre{
            margin:.6em 0;
            background:#111827;
            color:#f9fafb;
            border-radius:10px;
            padding:12px;
            overflow:auto;
        }
        .ffn-memo-content pre code{
            background:transparent;
            color:inherit;
            padding:0;
        }
        .ffn-memo-footer{
            display:flex;
            align-items:center;
            justify-content:space-between;
            gap:8px;
        }
        .ffn-memo-meta{
            display:flex;
            align-items:center;
            gap:8px;
            flex-wrap:wrap;
            flex:1;
            min-width:0;
        }
        .ffn-memo-tag{
            display:inline-flex;
            align-items:center;
            padding:4px 10px;
            border-radius:999px;
            border:1px solid #d8e4f2;
            background:#f3f8fe;
            font-size:11px;
            font-weight:600;
            color:#335b86;
            white-space:nowrap;
            max-width:180px;
            overflow:hidden;
            text-overflow:ellipsis;
        }
        .ffn-memo-tag.ffn-note-custom-tag{
            background:#e9f2ff;
            border-color:#cfe3ff;
            color:#155fb3;
        }
        .ffn-memo-tag.ffn-note-custom-tag:nth-child(4n+1){
            background:#e9f7ef;
            border-color:#caead7;
            color:#1f6a46;
        }
        .ffn-memo-tag.ffn-note-custom-tag:nth-child(4n+2){
            background:#fff4e8;
            border-color:#f8dec0;
            color:#7a460f;
        }
        .ffn-memo-tag.ffn-note-custom-tag:nth-child(4n+3){
            background:#f2eeff;
            border-color:#ddd2ff;
            color:#4931a1;
        }
        .ffn-memo-time{
            font-size:12px;
            color:#8b97a8;
            white-space:nowrap;
        }
        .ffn-memo-actions{
            display:flex;
            gap:4px;
            opacity:0;
            transform:translateY(2px);
            transition:opacity .18s,transform .18s;
            flex-shrink:0;
        }
        .ffn-memo-card:hover .ffn-memo-actions{
            opacity:1;
            transform:translateY(0);
        }
        .ffn-memo-action-btn{
            padding:5px 10px;
            border-radius:8px;
            border:none;
            background:transparent;
            cursor:pointer;
            font-size:12px;
            color:#5f6b7a;
            transition:background .1s,color .1s,box-shadow .1s;
        }
        .ffn-memo-action-btn:hover{
            background:#edf2f8;
            color:#1d1d1f;
        }
        .ffn-memo-action-btn:focus-visible{
            outline:none;
            box-shadow:0 0 0 3px rgba(0,113,227,.16);
        }
        .ffn-memo-action-btn.ffn-delete:hover{
            background:#ffecec;
            color:#d83b3b;
        }
        .ffn-empty-state{
            display:flex;
            flex-direction:column;
            align-items:center;
            justify-content:center;
            height:220px;
            color:#8f9bad;
            font-size:14px;
            gap:10px;
        }
        .ffn-empty-state-icon{
            font-size:42px;
            opacity:.8;
        }
        @keyframes ffn-fade-in{
            from{ opacity:0; }
            to{ opacity:1; }
        }
        @keyframes ffn-slide-up{
            from{
                opacity:0;
                transform:translateY(8px) scale(.99);
            }
            to{
                opacity:1;
                transform:translateY(0) scale(1);
            }
        }
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

    function parseTagsInput(value) {
        if (!value) return [];
        const source = Array.isArray(value) ? value.join(',') : String(value);
        const tags = source
            .split(/[,\n]/)
            .map(tag => tag.trim().replace(/^#/, ''))
            .filter(Boolean)
            .slice(0, MAX_NOTE_TAGS);
        const unique = [];
        const existed = new Set();
        tags.forEach(tag => {
            const key = tag.toLowerCase();
            if (existed.has(key)) return;
            existed.add(key);
            unique.push(tag);
        });
        return unique;
    }

    function formatTagsForInput(value) {
        return parseTagsInput(value).join(', ');
    }

    function encodeDraftPart(value) {
        return encodeURIComponent(String(value || DRAFT_KEY_EMPTY_PLACEHOLDER).trim());
    }

    function getDraftKey(note, projectNumber, fileNo) {
        if (note && note.id) return `${DRAFT_KEY_PREFIX}edit__${encodeDraftPart(note.id)}`;
        return `${DRAFT_KEY_PREFIX}new__${encodeDraftPart(fileNo || DRAFT_KEY_EMPTY_PLACEHOLDER)}`;
    }

    function saveDraftText(draftKey, value) {
        localStorage.setItem(draftKey, value || '');
    }

    function loadDraftText(draftKey) {
        return localStorage.getItem(draftKey);
    }

    function clearDraftText(draftKey) {
        localStorage.removeItem(draftKey);
    }

    function getInitialEditorText(note, options, draftKey) {
        const savedDraft = loadDraftText(draftKey);
        if (savedDraft !== null) return savedDraft;
        if (note) return note.text || '';
        return options.initialText || '';
    }

    function flushDraftSave(debouncedSaveDraft, draftKey, textarea) {
        debouncedSaveDraft.cancel();
        saveDraftText(draftKey, textarea.value);
    }

    function debounce(fn, delay) {
        let timer = null;
        const wrapped = (...args) => {
            if (timer) clearTimeout(timer);
            timer = window.setTimeout(() => {
                timer = null;
                fn(...args);
            }, delay);
        };
        wrapped.cancel = () => {
            if (!timer) return;
            clearTimeout(timer);
            timer = null;
        };
        return wrapped;
    }

    function escapeHtml(text) {
        return String(text || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    function renderMarkdownInline(text) {
        const escaped = escapeHtml(text);
        return escaped
            .replace(/`([^`]+?)`/g, '<code>$1</code>')
            .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
            .replace(/(?<!\*)\*([^*\n]+?)\*(?!\*)/g, '<em>$1</em>')
            .replace(/~~(.+?)~~/g, '<del>$1</del>');
    }

    function renderMarkdown(text) {
        const normalized = String(text || '').replace(/\r\n?/g, '\n');
        if (!normalized.trim()) return '';

        const lines = normalized.split('\n');
        const html = [];
        let inUl = false;
        let inOl = false;
        let inCode = false;
        let codeLines = [];

        const closeLists = () => {
            if (inUl) {
                html.push('</ul>');
                inUl = false;
            }
            if (inOl) {
                html.push('</ol>');
                inOl = false;
            }
        };

        lines.forEach(line => {
            if (/^```/.test(line)) {
                closeLists();
                if (inCode) {
                    html.push(`<pre><code>${codeLines.join('\n')}</code></pre>`);
                    codeLines = [];
                }
                inCode = !inCode;
                return;
            }

            if (inCode) {
                codeLines.push(escapeHtml(line));
                return;
            }

            const headingMatch = line.match(/^(#{1,6})\s+(.*)$/);
            if (headingMatch) {
                closeLists();
                const level = headingMatch[1].length;
                html.push(`<h${level}>${renderMarkdownInline(headingMatch[2])}</h${level}>`);
                return;
            }

            const ulMatch = line.match(/^[-*+]\s+(.*)$/);
            if (ulMatch) {
                if (inOl) {
                    html.push('</ol>');
                    inOl = false;
                }
                if (!inUl) {
                    html.push('<ul>');
                    inUl = true;
                }
                html.push(`<li>${renderMarkdownInline(ulMatch[1])}</li>`);
                return;
            }

            const olMatch = line.match(/^\d+\.\s+(.*)$/);
            if (olMatch) {
                if (inUl) {
                    html.push('</ul>');
                    inUl = false;
                }
                if (!inOl) {
                    html.push('<ol>');
                    inOl = true;
                }
                html.push(`<li>${renderMarkdownInline(olMatch[1])}</li>`);
                return;
            }

            const quoteMatch = line.match(/^>\s?(.*)$/);
            if (quoteMatch) {
                closeLists();
                html.push(`<blockquote>${renderMarkdownInline(quoteMatch[1])}</blockquote>`);
                return;
            }

            if (!line.trim()) {
                closeLists();
                return;
            }

            closeLists();
            html.push(`<p>${renderMarkdownInline(line)}</p>`);
        });

        if (inCode) html.push(`<pre><code>${codeLines.join('\n')}</code></pre>`);
        closeLists();
        return html.join('');
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
        const draftKey = getDraftKey(note, projectNumber, fileNo);

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
        textarea.value = getInitialEditorText(note, options, draftKey);

        const tagsInput = document.createElement('input');
        tagsInput.type = 'text';
        tagsInput.className = 'ffn-note-tags-input';
        tagsInput.placeholder = '標籤（以逗號分隔，例如：urgent, follow-up）';
        tagsInput.value = note ? formatTagsForInput(note.tags) : '';

        const debouncedSaveDraft = debounce(() => saveDraftText(draftKey, textarea.value), DRAFT_DEBOUNCE_MS);
        textarea.addEventListener('input', debouncedSaveDraft);

        const actions = document.createElement('div');
        actions.className = 'ffn-actions';

        const cancelBtn = document.createElement('button');
        cancelBtn.textContent = '取消';

        const saveBtn = document.createElement('button');
        saveBtn.className = 'ffn-primary';
        saveBtn.textContent = '儲存';

        cancelBtn.addEventListener('click', () => {
            flushDraftSave(debouncedSaveDraft, draftKey, textarea);
            closeNoteEditor(box);
        });
        backdrop.addEventListener('click', (e) => {
            if (e.target === backdrop) {
                flushDraftSave(debouncedSaveDraft, draftKey, textarea);
                closeNoteEditor(box);
            }
        });

        saveBtn.addEventListener('click', () => {
            const text = textarea.value.trim();
            const tags = parseTagsInput(tagsInput.value);
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
                        tags,
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
                    tags,
                    createdAt: now,
                    updatedAt: now
                });
            }
            saveStore(store);
            debouncedSaveDraft.cancel();
            clearDraftText(draftKey);
            closeNoteEditor(box);
            renderNotesList();
        });

        actions.append(cancelBtn, saveBtn);
        box.append(title, info, textarea, tagsInput, actions);
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
        const tagKeyword = notesPanelEl.querySelector('#ffn-tag-filter').value.trim().replace(/^#/, '').toLowerCase();

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
            if (tagKeyword) {
                const tags = parseTagsInput(note.tags);
                if (!tags.some(tag => tag.toLowerCase().includes(tagKeyword))) return false;
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

    function formatDateGroupLabel(isoString) {
        const d = new Date(isoString);
        if (Number.isNaN(d.getTime())) return '未知日期';
        const today = new Date();
        const yesterday = new Date(today);
        yesterday.setDate(today.getDate() - 1);
        const toLabel = (date) => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
        const dLabel = toLabel(d);
        if (dLabel === toLabel(today)) return '今天';
        if (dLabel === toLabel(yesterday)) return '昨天';
        return dLabel;
    }

    function renderNotesList() {
        if (!notesPanelEl) return;
        const list = notesPanelEl.querySelector('.ffn-notes-list');
        list.innerHTML = '';

        const notes = getFilteredSortedNotes();

        // Update header count
        const headerTitle = notesPanelEl.querySelector('.ffn-main-header-title');
        if (headerTitle) headerTitle.textContent = `所有筆記（${notes.length}）`;

        if (!notes.length) {
            const empty = document.createElement('div');
            empty.className = 'ffn-empty-state';
            empty.innerHTML = '<div class="ffn-empty-state-icon">📭</div><div>目前沒有符合條件的筆記</div>';
            list.appendChild(empty);
            return;
        }

        const sortBy = notesPanelEl.querySelector('#ffn-sort-by').value;
        const dateField = sortBy === 'updatedAt' ? 'updatedAt' : 'createdAt';

        let lastGroupLabel = null;
        notes.forEach(note => {
            const groupLabel = formatDateGroupLabel(note[dateField] || note.createdAt);
            if (groupLabel !== lastGroupLabel) {
                lastGroupLabel = groupLabel;
                const groupEl = document.createElement('div');
                groupEl.className = 'ffn-date-group';
                groupEl.textContent = groupLabel;
                list.appendChild(groupEl);
            }

            const card = document.createElement('div');
            card.className = 'ffn-memo-card';

            const content = document.createElement('div');
            content.className = 'ffn-memo-content';
            // renderMarkdown 內部會先 escape 使用者輸入，再套用受控 markdown 標記
            content.innerHTML = renderMarkdown(note.text || '');

            const footer = document.createElement('div');
            footer.className = 'ffn-memo-footer';

            const meta = document.createElement('div');
            meta.className = 'ffn-memo-meta';

            if (note.projectNumber || note.fileNo) {
                const tag = document.createElement('span');
                tag.className = 'ffn-memo-tag';
                const fileNoPart = note.fileNo ? ` · ${note.fileNo}` : '';
                tag.textContent = `📁 ${note.projectNumber || '-'}${fileNoPart}`;
                meta.appendChild(tag);
            }
            if (note.projectName) {
                const nameTag = document.createElement('span');
                nameTag.className = 'ffn-memo-tag';
                nameTag.textContent = note.projectName;
                meta.appendChild(nameTag);
            }
            parseTagsInput(note.tags).forEach(noteTag => {
                const customTag = document.createElement('span');
                customTag.className = 'ffn-memo-tag ffn-note-custom-tag';
                customTag.textContent = `#${noteTag}`;
                meta.appendChild(customTag);
            });

            const time = document.createElement('span');
            time.className = 'ffn-memo-time';
            time.title = `建立: ${formatDateTime(note.createdAt)}`;
            time.textContent = formatDateTime(note[dateField] || note.createdAt);

            const actions = document.createElement('div');
            actions.className = 'ffn-memo-actions';

            const editBtn = document.createElement('button');
            editBtn.className = 'ffn-memo-action-btn';
            editBtn.textContent = '✏️ 編輯';
            editBtn.addEventListener('click', () => openNoteEditor({ note }));

            const delBtn = document.createElement('button');
            delBtn.className = 'ffn-memo-action-btn ffn-delete';
            delBtn.textContent = '🗑️ 刪除';
            delBtn.addEventListener('click', () => {
                if (!confirm('確定要刪除此筆記？')) return;
                const store = loadStore();
                store.notes = store.notes.filter(n => n.id !== note.id);
                saveStore(store);
                renderNotesList();
            });

            actions.append(editBtn, delBtn);
            footer.append(meta, time, actions);
            card.append(content, footer);
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
            <div class="ffn-sidebar">
                <h2 class="ffn-sidebar-title">📝 Quick Notes</h2>
                <button type="button" id="ffn-today-log" class="ffn-sidebar-new-btn">＋ 新增今日紀錄</button>
                <div class="ffn-sidebar-section">
                    <div class="ffn-sidebar-label">排序方式</div>
                    <select id="ffn-sort-by">
                        <option value="updatedAt">最近更新</option>
                        <option value="createdAt">建立時間</option>
                        <option value="projectName">專案名稱</option>
                    </select>
                    <select id="ffn-order">
                        <option value="desc">新 → 舊</option>
                        <option value="asc">舊 → 新</option>
                    </select>
                </div>
                <div class="ffn-sidebar-section">
                    <div class="ffn-sidebar-label">篩選</div>
                    <input id="ffn-project-filter" type="text" placeholder="🔍 專案編號">
                    <input id="ffn-keyword" type="text" placeholder="🔍 關鍵字搜尋">
                    <input id="ffn-tag-filter" type="text" placeholder="🏷️ 標籤（例如：urgent）">
                </div>
                <div class="ffn-sidebar-section">
                    <div class="ffn-sidebar-label">日期範圍</div>
                    <input id="ffn-date-from" type="date" title="開始日期">
                    <input id="ffn-date-to" type="date" title="結束日期">
                </div>
            </div>
            <div class="ffn-main">
                <div class="ffn-main-header">
                    <span class="ffn-main-header-title">所有筆記</span>
                    <button type="button" id="ffn-close-panel" class="ffn-close-btn" title="關閉">✕</button>
                </div>
                <div class="ffn-notes-list"></div>
            </div>
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
