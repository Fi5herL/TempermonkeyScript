// ==UserScript==
// @name         Flex 📁 Folder Copy & 📝 Quick Notes
// @namespace    fisher-flex-folder-notes
// @version      1.2.0
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

        /* ── 筆記編輯對話框 ── */
        .ffn-note-editor-backdrop{
            position:fixed;
            inset:0;
            background:rgba(0,0,0,.45);
            z-index:99999;
            display:flex;
            align-items:center;
            justify-content:center;
        }
        .ffn-note-editor{
            background:#fff;
            border-radius:12px;
            box-shadow:0 24px 64px rgba(0,0,0,.25);
            width:560px;
            max-width:94vw;
            padding:20px 24px;
            color:#111;
            font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;
        }
        .ffn-note-editor h3{
            margin:0 0 12px;
            font-size:16px;
            font-weight:600;
        }
        .ffn-note-info{
            font-size:12px;
            color:#6b7280;
            background:#f9fafb;
            border:1px solid #e5e7eb;
            border-radius:6px;
            padding:8px 10px;
            margin-bottom:12px;
        }
        .ffn-note-editor textarea{
            width:100%;
            min-height:140px;
            resize:vertical;
            box-sizing:border-box;
            border:1px solid #e5e7eb;
            border-radius:8px;
            padding:10px 12px;
            font-size:14px;
            line-height:1.6;
            color:#111;
            outline:none;
            transition:border-color .15s;
            font-family:inherit;
        }
        .ffn-note-editor textarea:focus{ border-color:#6366f1; }
        .ffn-note-tags-input{
            margin-top:10px;
            width:100%;
            box-sizing:border-box;
            border:1px solid #e5e7eb;
            border-radius:8px;
            padding:8px 10px;
            font-size:13px;
            color:#111;
            font-family:inherit;
            outline:none;
            transition:border-color .15s;
        }
        .ffn-note-tags-input:focus{ border-color:#6366f1; }
        .ffn-actions{
            margin-top:12px;
            display:flex;
            gap:8px;
            justify-content:flex-end;
        }
        .ffn-note-editor button{
            padding:8px 18px;
            border-radius:8px;
            border:1px solid #e5e7eb;
            background:#fff;
            cursor:pointer;
            font-size:13px;
            font-weight:500;
            color:#374151;
            transition:background .15s,border-color .15s;
        }
        .ffn-note-editor button:hover{ background:#f9fafb; }
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
            background:rgba(0,0,0,.45);
            z-index:99997;
            display:flex;
            align-items:center;
            justify-content:center;
        }
        .ffn-notes-panel{
            width:90vw;
            max-width:1100px;
            height:85vh;
            display:flex;
            border-radius:12px;
            overflow:hidden;
            background:#fff;
            box-shadow:0 24px 64px rgba(0,0,0,.25);
            color:#1a1a1a;
            font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;
        }

        /* Sidebar */
        .ffn-sidebar{
            width:252px;
            flex-shrink:0;
            background:#f7f7f8;
            border-right:1px solid #e8e8e8;
            display:flex;
            flex-direction:column;
            padding:20px 16px;
            gap:14px;
            overflow-y:auto;
            box-sizing:border-box;
        }
        .ffn-sidebar-title{
            font-size:17px;
            font-weight:600;
            color:#111;
            margin:0;
            display:flex;
            align-items:center;
            gap:7px;
        }
        .ffn-sidebar-new-btn{
            width:100%;
            padding:9px 14px;
            border-radius:8px;
            background:#4f46e5;
            color:#fff;
            border:none;
            font-size:13px;
            font-weight:500;
            cursor:pointer;
            display:flex;
            align-items:center;
            gap:6px;
            justify-content:center;
            transition:background .15s;
        }
        .ffn-sidebar-new-btn:hover{ background:#4338ca; }
        .ffn-sidebar-section{
            display:flex;
            flex-direction:column;
            gap:6px;
        }
        .ffn-sidebar-label{
            font-size:11px;
            font-weight:600;
            color:#9ca3af;
            text-transform:uppercase;
            letter-spacing:.05em;
        }
        .ffn-sidebar input,
        .ffn-sidebar select{
            width:100%;
            padding:7px 10px;
            border:1px solid #e0e0e0;
            border-radius:6px;
            background:#fff;
            font-size:13px;
            color:#111;
            box-sizing:border-box;
            outline:none;
            transition:border-color .15s;
        }
        .ffn-sidebar input:focus,
        .ffn-sidebar select:focus{ border-color:#6366f1; }

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
            padding:14px 20px;
            border-bottom:1px solid #f0f0f0;
            flex-shrink:0;
        }
        .ffn-main-header-title{
            font-size:14px;
            font-weight:500;
            color:#6b7280;
        }
        .ffn-close-btn{
            width:30px;
            height:30px;
            border-radius:50%;
            border:none;
            background:#f3f4f6;
            cursor:pointer;
            font-size:15px;
            display:flex;
            align-items:center;
            justify-content:center;
            color:#6b7280;
            transition:background .15s;
            line-height:1;
        }
        .ffn-close-btn:hover{ background:#e5e7eb; color:#111; }
        .ffn-notes-list{
            flex:1;
            overflow-y:auto;
            padding:12px 20px 20px;
        }

        /* Date group */
        .ffn-date-group{
            font-size:11px;
            font-weight:600;
            color:#9ca3af;
            text-transform:uppercase;
            letter-spacing:.06em;
            margin:18px 0 8px;
            padding-left:2px;
        }
        .ffn-date-group:first-child{ margin-top:4px; }

        /* Memo card */
        .ffn-memo-card{
            border:1px solid #f0f0f0;
            border-radius:10px;
            padding:14px 16px;
            margin-bottom:8px;
            background:#fff;
            transition:box-shadow .15s,border-color .15s;
        }
        .ffn-memo-card:hover{
            box-shadow:0 2px 14px rgba(0,0,0,.07);
            border-color:#e5e7eb;
        }
        .ffn-memo-content{
            font-size:14px;
            line-height:1.65;
            color:#1a1a1a;
            white-space:normal;
            word-break:break-word;
            margin-bottom:10px;
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
            color:#4b5563;
            background:#f9fafb;
        }
        .ffn-memo-content ul,.ffn-memo-content ol{
            margin:.5em 0;
            padding-left:1.3em;
        }
        .ffn-memo-content code{
            font-family:ui-monospace,SFMono-Regular,Menlo,Monaco,Consolas,'Liberation Mono','Courier New',monospace;
            background:#f3f4f6;
            border-radius:4px;
            padding:1px 5px;
            font-size:12px;
        }
        .ffn-memo-content pre{
            margin:.6em 0;
            background:#111827;
            color:#f9fafb;
            border-radius:8px;
            padding:10px;
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
            gap:6px;
            flex-wrap:wrap;
            flex:1;
            min-width:0;
        }
        .ffn-memo-tag{
            display:inline-flex;
            align-items:center;
            padding:2px 8px;
            border-radius:4px;
            background:#f3f4f6;
            font-size:12px;
            color:#374151;
            white-space:nowrap;
            max-width:180px;
            overflow:hidden;
            text-overflow:ellipsis;
        }
        .ffn-memo-tag.ffn-note-custom-tag{
            background:#eef2ff;
            color:#4338ca;
        }
        .ffn-memo-time{
            font-size:12px;
            color:#9ca3af;
            white-space:nowrap;
        }
        .ffn-memo-actions{
            display:flex;
            gap:2px;
            opacity:0;
            transition:opacity .15s;
            flex-shrink:0;
        }
        .ffn-memo-card:hover .ffn-memo-actions{ opacity:1; }
        .ffn-memo-action-btn{
            padding:4px 9px;
            border-radius:6px;
            border:none;
            background:transparent;
            cursor:pointer;
            font-size:12px;
            color:#6b7280;
            transition:background .1s,color .1s;
        }
        .ffn-memo-action-btn:hover{
            background:#f3f4f6;
            color:#111;
        }
        .ffn-memo-action-btn.ffn-delete:hover{
            background:#fee2e2;
            color:#ef4444;
        }
        .ffn-empty-state{
            display:flex;
            flex-direction:column;
            align-items:center;
            justify-content:center;
            height:220px;
            color:#9ca3af;
            font-size:14px;
            gap:10px;
        }
        .ffn-empty-state-icon{ font-size:42px; }
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

    function sanitizeDraftPart(value) {
        return encodeURIComponent(String(value || '_').trim());
    }

    function getDraftKey(note, projectNumber, fileNo) {
        if (note && note.id) return `${DRAFT_KEY_PREFIX}edit__${sanitizeDraftPart(note.id)}`;
        return `${DRAFT_KEY_PREFIX}new__${sanitizeDraftPart(fileNo || '_')}`;
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
            .replace(/`([^`]+)`/g, '<code>$1</code>')
            .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
            .replace(/\*([^*]+)\*/g, '<em>$1</em>')
            .replace(/~~([^~]+)~~/g, '<del>$1</del>');
    }

    function renderMarkdown(text) {
        const normalized = String(text || '').replace(/\r\n?/g, '\n');
        if (!normalized.trim()) return '';

        const lines = normalized.split('\n');
        const html = [];
        let inUl = false;
        let inOl = false;
        let inCode = false;

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
                html.push(inCode ? '</code></pre>' : '<pre><code>');
                inCode = !inCode;
                return;
            }

            if (inCode) {
                html.push(`${escapeHtml(line)}\n`);
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

        if (inCode) html.push('</code></pre>');
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
            debouncedSaveDraft.cancel();
            saveDraftText(draftKey, textarea.value);
            closeNoteEditor(box);
        });
        backdrop.addEventListener('click', (e) => {
            if (e.target === backdrop) {
                debouncedSaveDraft.cancel();
                saveDraftText(draftKey, textarea.value);
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
