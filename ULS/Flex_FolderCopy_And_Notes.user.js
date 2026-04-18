// ==UserScript==
// @name         Flex 📁 Folder Copy & 📝 Quick Notes
// @namespace    fisher-flex-folder-notes
// @version      2.0.0
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
    const PREVIEW_PLACEHOLDER_HTML = '<p class="ffn-preview-placeholder">預覽區域</p>';
    const FLASH_MS = 300;
    const KENDO_RETRY_MS = 300;
    const GRID_OBSERVER_RETRY_MS = 500;

    let notesOverlayEl = null;
    let notesPanelEl = null;
    let gridObserver = null;
    let uuidCounter = 0;
    const inlineBoundContentEls = new WeakSet();
    const inlineCardState = new WeakMap();

    GM_addStyle(`
        /* ── 側邊收合抽屜 ── */
        .ffn-btn-drawer{
            position:absolute;
            left:0;
            top:50%;
            transform:translateY(-50%);
            width:5px;
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
            background:rgba(0,0,0,.2);
            z-index:99999;
            display:flex;
            align-items:center;
            justify-content:center;
            animation:ffn-fade-in .22s ease-out;
        }
        .ffn-note-editor{
            background:#fff;
            border:1px solid #e5e7eb;
            border-radius:8px;
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
            color:#4b5563;
            background:#f9fafb;
            border:1px solid #e5e7eb;
            border-radius:6px;
            padding:10px 12px;
            margin-bottom:14px;
            line-height:1.5;
        }
        .ffn-note-editor-tabs{
            display:flex;
            gap:8px;
            margin-bottom:10px;
        }
        .ffn-note-editor .ffn-note-editor-tab{
            border:none;
            background:transparent;
            border-radius:6px;
            padding:6px 10px;
            font-size:13px;
            color:#4b5563;
            cursor:pointer;
            transition:background .15s,color .15s;
        }
        .ffn-note-editor .ffn-note-editor-tab:hover{ background:#f3f4f6; }
        .ffn-note-editor .ffn-note-editor-tab.active{
            background:#eef2ff;
            color:#3730a3;
        }
        .ffn-note-editor-panel{ display:none; }
        .ffn-note-editor-panel.active{ display:block; }
        .ffn-note-editor textarea{
            width:100%;
            min-height:152px;
            resize:vertical;
            box-sizing:border-box;
            border:1px solid #d1d5db;
            border-radius:6px;
            background:transparent;
            padding:12px 14px;
            font-size:14px;
            line-height:1.7;
            color:#1f2937;
            outline:none;
            transition:border-color .15s,outline-color .15s;
            font-family:inherit;
        }
        .ffn-note-editor textarea:focus{
            border-color:#4f46e5;
            outline:2px solid #4f46e5;
            outline-offset:0;
        }
        .ffn-note-preview{
            min-height:152px;
            box-sizing:border-box;
            border:1px solid #d1d5db;
            border-radius:6px;
            padding:12px 14px;
            background:#fff;
            overflow:auto;
        }
        .ffn-note-mini-preview-wrap{
            margin-top:10px;
            border:1px solid #e5e7eb;
            border-radius:6px;
            background:#fff;
            overflow:hidden;
        }
        .ffn-note-mini-preview-wrap > summary{
            cursor:pointer;
            user-select:none;
            list-style:none;
            padding:8px 12px;
            font-size:12px;
            color:#6b7280;
            border-bottom:1px solid #f3f4f6;
        }
        .ffn-note-mini-preview-wrap > summary::-webkit-details-marker{ display:none; }
        .ffn-note-mini-preview{
            min-height:68px;
            max-height:180px;
            overflow:auto;
            padding:10px 12px;
            font-size:13px;
        }
        .ffn-preview-placeholder{
            opacity:.4;
        }
        .ffn-note-tags-input{
            margin-top:12px;
            width:100%;
            box-sizing:border-box;
            border:1px solid #d1d5db;
            border-radius:6px;
            background:transparent;
            padding:10px 12px;
            font-size:13px;
            color:#1f2937;
            font-family:inherit;
            outline:none;
            transition:border-color .15s,outline-color .15s;
        }
        .ffn-note-tags-input:focus{
            border-color:#4f46e5;
            outline:2px solid #4f46e5;
            outline-offset:0;
        }
        .ffn-actions{
            margin-top:16px;
            display:flex;
            gap:10px;
            justify-content:flex-end;
        }
        .ffn-note-editor button{
            padding:9px 18px;
            border-radius:6px;
            border:1px solid transparent;
            background:transparent;
            cursor:pointer;
            font-size:13px;
            font-weight:600;
            color:#374151;
            transition:background .15s,border-color .15s,color .15s;
        }
        .ffn-note-editor button:hover{
            background:#f3f4f6;
        }
        .ffn-note-editor button:focus-visible{
            outline:none;
            border-color:#4f46e5;
            outline:2px solid #4f46e5;
        }
        .ffn-note-editor button.ffn-primary{
            background:#4f46e5;
            color:#fff;
            border-color:#4f46e5;
        }
        .ffn-note-editor button.ffn-primary:hover{
            background:#4338ca;
            border-color:#4338ca;
        }

        /* ── Memos 風格筆記面板 ── */
        .ffn-notes-overlay{
            position:fixed;
            inset:0;
            background:rgba(0,0,0,.2);
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
            border-radius:8px;
            overflow:hidden;
            background:#fff;
            border:1px solid #e5e7eb;
            color:#1d1d1f;
            font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;
            animation:ffn-slide-up .24s cubic-bezier(.2,.8,.2,1);
        }

        /* Sidebar */
        .ffn-sidebar{
            width:252px;
            flex-shrink:0;
            background:#f9fafb;
            border-right:1px solid #e5e7eb;
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
            border-radius:6px;
            background:#4f46e5;
            color:#fff;
            border:none;
            font-size:13px;
            font-weight:600;
            cursor:pointer;
            display:flex;
            align-items:center;
            gap:6px;
            justify-content:center;
            transition:background .15s;
        }
        .ffn-sidebar-new-btn:hover{
            background:#4338ca;
        }
        .ffn-sidebar-new-btn:focus-visible{
            outline:none;
            outline:2px solid #4f46e5;
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
            border:1px solid #d1d5db;
            border-radius:6px;
            background:transparent;
            font-size:13px;
            color:#1f2937;
            box-sizing:border-box;
            outline:none;
            transition:border-color .15s,outline-color .15s;
        }
        .ffn-sidebar input:focus,
        .ffn-sidebar select:focus{
            border-color:#4f46e5;
            outline:2px solid #4f46e5;
            outline-offset:0;
        }

        /* Main feed */
        .ffn-main{
            flex:1;
            display:flex;
            flex-direction:column;
            overflow:hidden;
            background:#fff;
        }
        .ffn-main-header{
            display:flex;
            align-items:center;
            justify-content:space-between;
            padding:16px 24px;
            border-bottom:1px solid #e5e7eb;
            flex-shrink:0;
            background:#fff;
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
            border-radius:6px;
            border:none;
            background:transparent;
            cursor:pointer;
            font-size:15px;
            display:flex;
            align-items:center;
            justify-content:center;
            color:#6b7787;
            transition:background .15s,color .15s;
            line-height:1;
        }
        .ffn-close-btn:hover{
            background:#f3f4f6;
            color:#1d1d1f;
        }
        .ffn-close-btn:focus-visible{
            outline:none;
            outline:2px solid #4f46e5;
        }
        .ffn-notes-list{
            flex:1;
            overflow-y:auto;
            padding:16px 24px 24px;
            background:#fff;
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
            position:relative;
            display:flex;
            flex-direction:column;
            align-items:flex-start;
            gap:8px;
            border:1px solid #e5e7eb;
            border-radius:8px;
            padding:12px 16px;
            margin-bottom:8px;
            background:#fff;
            transition:background .15s,border-color .15s;
        }
        .ffn-memo-card:hover{
            background:#fafafa;
        }
        .ffn-memo-content{
            width:100%;
            font-size:14px;
            line-height:1.7;
            color:#1f2937;
            white-space:normal;
            word-break:break-word;
            margin-bottom:12px;
            cursor:text;
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
            border-left:3px solid #d1d5db;
            color:#374151;
            background:#f9fafb;
        }
        .ffn-memo-content ul,.ffn-memo-content ol{
            margin:.5em 0;
            padding-left:1.3em;
        }
        .ffn-memo-content code{
            font-family:ui-monospace,SFMono-Regular,Menlo,Monaco,Consolas,'Liberation Mono','Courier New',monospace;
            background:#f3f4f6;
            border-radius:6px;
            padding:2px 6px;
            font-size:12px;
        }
        .ffn-memo-content pre{
            margin:.6em 0;
            background:#111827;
            color:#f9fafb;
            border-radius:6px;
            padding:12px;
            overflow:auto;
        }
        .ffn-memo-content pre code{
            background:transparent;
            color:inherit;
            padding:0;
        }
        .ffn-memo-footer{
            width:100%;
            display:flex;
            align-items:center;
            justify-content:flex-start;
            gap:8px;
        }
        .ffn-memo-header{
            width:100%;
            display:flex;
            justify-content:space-between;
            align-items:center;
        }
        .ffn-memo-meta{
            width:100%;
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
            border:1px solid #e5e7eb;
            background:#f3f4f6;
            font-size:11px;
            font-weight:600;
            color:#374151;
            white-space:nowrap;
            max-width:180px;
            overflow:hidden;
            text-overflow:ellipsis;
        }
        .ffn-memo-time{
            font-size:12px;
            color:#6b7280;
            white-space:nowrap;
        }
        .ffn-memo-actions{
            display:flex;
            gap:4px;
            flex-shrink:0;
            opacity:0;
            transition:opacity .15s;
        }
        .ffn-memo-card:hover .ffn-memo-actions{
            opacity:1;
        }
        .ffn-memo-action-btn{
            padding:5px 10px;
            border-radius:6px;
            border:none;
            background:transparent;
            cursor:pointer;
            font-size:12px;
            color:#4b5563;
            transition:background .1s,color .1s;
        }
        .ffn-memo-action-btn:hover{
            background:#f3f4f6;
            color:#1d1d1f;
        }
        .ffn-memo-action-btn:focus-visible{
            outline:none;
            outline:2px solid #4f46e5;
        }
        .ffn-memo-action-btn.ffn-primary{
            background:#4f46e5;
            color:#fff;
        }
        .ffn-memo-action-btn.ffn-primary:hover{
            background:#4338ca;
            color:#fff;
        }
        .ffn-memo-action-btn.ffn-delete:hover{
            background:#fef2f2;
            color:#dc2626;
        }
        .ffn-memo-card.ffn-editing{
            border-color:#818cf8;
            box-shadow:0 0 0 2px rgba(99,102,241,.15);
        }
        .ffn-inline-editor{
            width:100%;
            display:grid;
            grid-template-columns:1fr 1fr;
            gap:12px;
            min-height:120px;
        }
        @media(max-width:700px){
            .ffn-inline-editor{
                grid-template-columns:1fr;
            }
        }
        .ffn-inline-textarea{
            width:100%;
            box-sizing:border-box;
            border:1px solid #d1d5db;
            border-radius:6px;
            background:#fafafa;
            padding:12px;
            font-size:14px;
            line-height:1.7;
            color:#1f2937;
            resize:none;
            overflow:hidden;
            outline:none;
            font-family:ui-monospace,SFMono-Regular,Menlo,Monaco,Consolas,monospace;
            transition:border-color .15s;
            min-height:120px;
        }
        .ffn-inline-textarea:focus{
            border-color:#4f46e5;
            background:#fff;
        }
        .ffn-inline-preview{
            border:1px solid #e5e7eb;
            border-radius:6px;
            padding:12px;
            background:#fff;
            overflow-y:auto;
            max-height:400px;
            min-height:120px;
        }
        .ffn-inline-toolbar{
            width:100%;
            display:flex;
            justify-content:space-between;
            align-items:center;
            margin-top:8px;
            padding-top:8px;
            border-top:1px solid #f3f4f6;
        }
        .ffn-md-shortcuts{
            display:flex;
            gap:2px;
        }
        .ffn-md-btn{
            width:28px;
            height:28px;
            display:flex;
            align-items:center;
            justify-content:center;
            border:none;
            background:transparent;
            border-radius:4px;
            font-size:12px;
            font-weight:600;
            color:#6b7280;
            cursor:pointer;
            font-family:ui-monospace,SFMono-Regular,Menlo,monospace;
            transition:background .1s,color .1s;
        }
        .ffn-md-btn:hover{
            background:#f3f4f6;
            color:#111827;
        }
        .ffn-inline-action-btns{
            display:flex;
            gap:6px;
        }
        .ffn-inline-cancel{
            padding:6px 14px;
            border:none;
            background:transparent;
            border-radius:6px;
            font-size:13px;
            font-weight:500;
            color:#6b7280;
            cursor:pointer;
            transition:background .1s;
        }
        .ffn-inline-cancel:hover{ background:#f3f4f6; }
        .ffn-inline-save{
            padding:6px 14px;
            border:none;
            background:#4f46e5;
            color:#fff;
            border-radius:6px;
            font-size:13px;
            font-weight:500;
            cursor:pointer;
            transition:background .1s;
        }
        .ffn-inline-save:hover{ background:#4338ca; }
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

    function autoGrowTextarea(textarea) {
        textarea.style.height = 'auto';
        textarea.style.height = `${textarea.scrollHeight}px`;
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

        const tabs = document.createElement('div');
        tabs.className = 'ffn-note-editor-tabs';

        const editTabBtn = document.createElement('button');
        editTabBtn.type = 'button';
        editTabBtn.className = 'ffn-note-editor-tab active';
        editTabBtn.textContent = '✏️ 編輯';

        const previewTabBtn = document.createElement('button');
        previewTabBtn.type = 'button';
        previewTabBtn.className = 'ffn-note-editor-tab';
        previewTabBtn.textContent = '👁️ 預覽';

        tabs.append(editTabBtn, previewTabBtn);

        const editPanel = document.createElement('div');
        editPanel.className = 'ffn-note-editor-panel active';
        editPanel.appendChild(textarea);

        const miniPreviewWrap = document.createElement('details');
        miniPreviewWrap.className = 'ffn-note-mini-preview-wrap';

        const miniPreviewSummary = document.createElement('summary');
        miniPreviewSummary.textContent = '即時預覽';

        const miniPreviewContent = document.createElement('div');
        miniPreviewContent.className = 'ffn-note-mini-preview ffn-memo-content';
        miniPreviewWrap.append(miniPreviewSummary, miniPreviewContent);
        editPanel.appendChild(miniPreviewWrap);
        miniPreviewWrap.open = true;

        const previewPanel = document.createElement('div');
        previewPanel.className = 'ffn-note-editor-panel';

        const previewContent = document.createElement('div');
        previewContent.className = 'ffn-note-preview ffn-memo-content';
        previewPanel.appendChild(previewContent);

        const tagsInput = document.createElement('input');
        tagsInput.type = 'text';
        tagsInput.className = 'ffn-note-tags-input';
        tagsInput.placeholder = '標籤（以逗號分隔，例如：urgent, follow-up）';
        tagsInput.value = note ? formatTagsForInput(note.tags) : '';

        const updatePreview = () => {
            const html = renderMarkdown(textarea.value || '');
            previewContent.innerHTML = html;
            miniPreviewContent.innerHTML = html || PREVIEW_PLACEHOLDER_HTML;
        };
        updatePreview();

        const switchEditorTab = (tab) => {
            const isEdit = tab === 'edit';
            editTabBtn.classList.toggle('active', isEdit);
            previewTabBtn.classList.toggle('active', !isEdit);
            editPanel.classList.toggle('active', isEdit);
            previewPanel.classList.toggle('active', !isEdit);
        };
        editTabBtn.addEventListener('click', () => switchEditorTab('edit'));
        previewTabBtn.addEventListener('click', () => switchEditorTab('preview'));

        const debouncedSaveDraft = debounce(() => saveDraftText(draftKey, textarea.value), DRAFT_DEBOUNCE_MS);
        textarea.addEventListener('input', () => {
            debouncedSaveDraft();
            updatePreview();
        });

        const actions = document.createElement('div');
        actions.className = 'ffn-actions';

        const cancelBtn = document.createElement('button');
        cancelBtn.textContent = '取消';

        const saveBtn = document.createElement('button');
        saveBtn.className = 'ffn-primary';
        saveBtn.textContent = '儲存';

        const saveNote = () => {
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
        };

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
        saveBtn.addEventListener('click', saveNote);

        const handleSaveShortcut = (e) => {
            if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
                e.preventDefault();
                saveNote();
            }
        };
        textarea.addEventListener('keydown', handleSaveShortcut);
        tagsInput.addEventListener('keydown', handleSaveShortcut);

        actions.append(cancelBtn, saveBtn);
        box.append(title, info, tabs, editPanel, previewPanel, tagsInput, actions);
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

    function bindInlineEditTrigger(card, note) {
        const contentEl = card.querySelector('.ffn-memo-content');
        if (!contentEl) return;
        if (inlineBoundContentEls.has(contentEl)) return;
        inlineBoundContentEls.add(contentEl);
        contentEl.addEventListener('dblclick', (e) => {
            e.preventDefault();
            enterInlineEdit(card, note);
        });
    }

    function enterInlineEdit(card, note) {
        if (!card || !note) return;
        if (card.classList.contains('ffn-editing')) return;
        const existing = document.querySelector('.ffn-memo-card.ffn-editing');
        if (existing && existing !== card) exitInlineEdit(existing);

        const contentEl = card.querySelector('.ffn-memo-content');
        if (!contentEl) return;

        card.classList.add('ffn-editing');
        inlineCardState.set(card, { originalHTML: contentEl.innerHTML, note: { ...note } });
        contentEl.innerHTML = '';

        const editorWrap = document.createElement('div');
        editorWrap.className = 'ffn-inline-editor';

        const textarea = document.createElement('textarea');
        textarea.className = 'ffn-inline-textarea';
        textarea.value = note.text || '';
        textarea.placeholder = '輸入筆記內容（支援 Markdown）...';

        const preview = document.createElement('div');
        preview.className = 'ffn-inline-preview ffn-memo-content';
        preview.innerHTML = renderMarkdown(note.text || '') || PREVIEW_PLACEHOLDER_HTML;

        const updatePreview = () => {
            preview.innerHTML = renderMarkdown(textarea.value || '') || PREVIEW_PLACEHOLDER_HTML;
        };

        textarea.addEventListener('input', () => {
            autoGrowTextarea(textarea);
            updatePreview();
        });

        const toolbar = document.createElement('div');
        toolbar.className = 'ffn-inline-toolbar';

        const mdBtns = document.createElement('div');
        mdBtns.className = 'ffn-md-shortcuts';
        const shortcuts = [
            { label: 'B', title: '粗體', prefix: '**', suffix: '**' },
            { label: 'I', title: '斜體', prefix: '*', suffix: '*' },
            { label: '~', title: '刪除線', prefix: '~~', suffix: '~~' },
            { label: '<>', title: '程式碼', prefix: '`', suffix: '`' },
            { label: '—', title: '分隔線', insert: '\n---\n' },
            { label: '•', title: '清單', insert: '\n- ' }
        ];

        shortcuts.forEach(s => {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'ffn-md-btn';
            btn.textContent = s.label;
            btn.title = s.title;
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                if (s.insert) {
                    const pos = textarea.selectionStart;
                    textarea.value = textarea.value.slice(0, pos) + s.insert + textarea.value.slice(pos);
                    textarea.selectionStart = textarea.selectionEnd = pos + s.insert.length;
                } else {
                    const start = textarea.selectionStart;
                    const end = textarea.selectionEnd;
                    const selected = textarea.value.slice(start, end);
                    textarea.value = textarea.value.slice(0, start) + s.prefix + selected + s.suffix + textarea.value.slice(end);
                    textarea.selectionStart = start + s.prefix.length;
                    textarea.selectionEnd = start + s.prefix.length + selected.length;
                }
                textarea.focus();
                autoGrowTextarea(textarea);
                updatePreview();
            });
            mdBtns.appendChild(btn);
        });

        const actionBtns = document.createElement('div');
        actionBtns.className = 'ffn-inline-action-btns';

        const cancelBtn = document.createElement('button');
        cancelBtn.type = 'button';
        cancelBtn.className = 'ffn-inline-cancel';
        cancelBtn.textContent = '取消';
        cancelBtn.addEventListener('click', () => exitInlineEdit(card));

        const saveBtn = document.createElement('button');
        saveBtn.type = 'button';
        saveBtn.className = 'ffn-inline-save';
        saveBtn.textContent = '儲存';
        saveBtn.addEventListener('click', () => commitInlineEdit(card, textarea.value));

        textarea.addEventListener('keydown', (e) => {
            if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
                e.preventDefault();
                commitInlineEdit(card, textarea.value);
            } else if (e.key === 'Escape') {
                e.preventDefault();
                exitInlineEdit(card);
            }
        });

        toolbar.append(mdBtns, actionBtns);
        actionBtns.append(cancelBtn, saveBtn);
        editorWrap.append(textarea, preview);
        contentEl.append(editorWrap, toolbar);

        requestAnimationFrame(() => {
            autoGrowTextarea(textarea);
            textarea.focus();
            if (typeof textarea.setSelectionRange === 'function') textarea.setSelectionRange(0, 0);
        });
    }

    function exitInlineEdit(card) {
        if (!card) return;
        const contentEl = card.querySelector('.ffn-memo-content');
        if (!contentEl) return;
        const state = inlineCardState.get(card) || {};
        card.classList.remove('ffn-editing');
        contentEl.innerHTML = state.originalHTML || '';
        bindInlineEditTrigger(card, state.note);
        inlineCardState.delete(card);
    }

    function commitInlineEdit(card, text) {
        if (!card) return;
        const trimmed = (text || '').trim();
        if (!trimmed) {
            alert('請輸入筆記內容');
            return;
        }

        const state = inlineCardState.get(card) || {};
        const note = state.note;
        if (!note) return;
        const store = loadStore();
        const idx = store.notes.findIndex(n => n.id === note.id);
        if (idx < 0) return;
        store.notes[idx].text = trimmed;
        store.notes[idx].updatedAt = new Date().toISOString();
        saveStore(store);
        const reboundNote = {
            ...note,
            text: trimmed,
            updatedAt: store.notes[idx].updatedAt
        };

        const contentEl = card.querySelector('.ffn-memo-content');
        if (!contentEl) return;
        card.classList.remove('ffn-editing');
        contentEl.innerHTML = renderMarkdown(trimmed);
        bindInlineEditTrigger(card, reboundNote);
        inlineCardState.delete(card);
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

            const time = document.createElement('span');
            time.className = 'ffn-memo-time';
            time.title = `建立: ${formatDateTime(note.createdAt)}`;
            time.textContent = formatDateTime(note[dateField] || note.createdAt);

            const actions = document.createElement('div');
            actions.className = 'ffn-memo-actions';

            const editBtn = document.createElement('button');
            editBtn.className = 'ffn-memo-action-btn';
            editBtn.textContent = '✏️';
            editBtn.title = '編輯';

            const delBtn = document.createElement('button');
            delBtn.className = 'ffn-memo-action-btn ffn-delete';
            delBtn.textContent = '🗑️';
            delBtn.title = '刪除';

            actions.append(editBtn, delBtn);

            const header = document.createElement('div');
            header.className = 'ffn-memo-header';
            header.append(time, actions);

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

            editBtn.addEventListener('click', (e) => {
                e.preventDefault();
                enterInlineEdit(card, note);
            });
            delBtn.addEventListener('click', () => {
                if (!confirm('確定要刪除此筆記？')) return;
                const store = loadStore();
                store.notes = store.notes.filter(n => n.id !== note.id);
                saveStore(store);
                renderNotesList();
            });

            footer.appendChild(meta);
            card.append(header, content, footer);
            bindInlineEditTrigger(card, note);
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
