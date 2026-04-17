// ==UserScript==
// @name         Flex Folder Copy Helper
// @namespace    fisher-flex-folder-copy-helper
// @version      1.0.0
// @description  在 Dashboard 每筆案件列新增 📁 按鈕，一鍵複製資料夾名稱格式。
// @match        https://portal.ul.com/Dashboard*
// @grant        none
// @run-at       document-end
// ==/UserScript==

(function () {
  'use strict';

  const GRID_ID = '#projectDashboardGrid';

  const GRID_HEADER_FOR = {
    projectNumber: 'Project Number',
    customer: 'Company Name',
  };

  const FOLDER_HEADER_FOR = {
    fileNumber: 'File Number',
    orderLines: 'Order Lines',
    dateCreated: 'Date Project Created',
  };

  const INVALID_FOLDER_CHARS_RE = /[<>:"/\\|?*]/g;
  const BTN_CLASS = 'feh-folder-copy-btn';

  function ensureStyle() {
    if (document.getElementById('feh-folder-copy-style')) return;
    const style = document.createElement('style');
    style.id = 'feh-folder-copy-style';
    style.textContent = `
      ${GRID_ID} td { overflow: visible; }
      ${GRID_ID} .${BTN_CLASS} {
        position: absolute;
        left: 32px;
        top: 50%;
        transform: translateY(-50%);
        width: 22px;
        height: 22px;
        border: 1px solid #cfd4dc;
        border-radius: 999px;
        background: #fff;
        color: #333;
        font-size: 12px;
        line-height: 1;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        transition: background-color .15s ease, border-color .15s ease, transform .15s ease;
        z-index: 2;
      }
      ${GRID_ID} .${BTN_CLASS}:hover {
        background: #eef3ff;
        border-color: #9ab0e4;
      }
    `;
    document.head.appendChild(style);
  }

  function buildHeaderIndex() {
    const ths = document.querySelectorAll(`${GRID_ID} .k-grid-header thead th`);
    const map = {};
    ths.forEach((th, i) => {
      const title = th.getAttribute('data-title') || th.textContent || '';
      const key = title.trim();
      if (key) map[key] = i;
    });
    return map;
  }

  function normalizeDateYYYYMMDD(input) {
    if (!input) return '';
    const text = String(input).trim();
    if (!text) return '';

    const ymd = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
    if (ymd) {
      const [, y, m, d] = ymd;
      return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
    }

    const mdy = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
    if (mdy) {
      const [, m, d, y] = mdy;
      return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
    }

    const date = new Date(text);
    if (!Number.isNaN(date.getTime())) {
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, '0');
      const d = String(date.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }

    return text;
  }

  function sanitizeFolderText(text) {
    return String(text || '').replace(INVALID_FOLDER_CHARS_RE, '_').trim();
  }

  function flash(el) {
    if (!el) return;
    const oldBg = el.style.backgroundColor;
    const oldBorder = el.style.borderColor;
    el.style.backgroundColor = '#b7f4c0';
    el.style.borderColor = '#4caf50';
    window.setTimeout(() => {
      el.style.backgroundColor = oldBg;
      el.style.borderColor = oldBorder;
    }, 400);
  }

  function cellText(cells, idx) {
    if (idx === null || idx === undefined || idx < 0 || idx >= cells.length) return '';
    return (cells[idx].textContent || '').trim();
  }

  function buildFolderNameFromRow(tr) {
    const headers = buildHeaderIndex();
    const cells = tr.querySelectorAll('td');

    const fileNumber = cellText(cells, headers[FOLDER_HEADER_FOR.fileNumber]);
    const projectNumber = cellText(cells, headers[GRID_HEADER_FOR.projectNumber]);
    const orderLines = cellText(cells, headers[FOLDER_HEADER_FOR.orderLines]);
    const company = cellText(cells, headers[GRID_HEADER_FOR.customer]);
    const dateCreatedRaw = cellText(cells, headers[FOLDER_HEADER_FOR.dateCreated]);

    const company10 = sanitizeFolderText(company.slice(0, 10));
    const dateCreated = normalizeDateYYYYMMDD(dateCreatedRaw);

    const folder = `${fileNumber}-${projectNumber}-${orderLines}-${company10}-${dateCreated}`;
    return sanitizeFolderText(folder);
  }

  async function onCopyFolderClick(tr, button) {
    const folderName = buildFolderNameFromRow(tr);
    if (!folderName || /^-+$/.test(folderName)) {
      console.warn('[Flex Folder Copy Helper] Empty folder name. Check configured headers.');
      return;
    }

    try {
      await navigator.clipboard.writeText(folderName);
      flash(button);
    } catch (err) {
      console.error('[Flex Folder Copy Helper] Failed to copy folder name:', err);
    }
  }

  function injectFolderButton(tr) {
    if (!(tr instanceof HTMLElement)) return;
    if (tr.querySelector(`.${BTN_CLASS}`)) return;

    const firstTd = tr.querySelector('td');
    if (!firstTd) return;
    firstTd.style.position = firstTd.style.position || 'relative';
    firstTd.style.overflow = 'visible';

    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = BTN_CLASS;
    btn.textContent = '📁';
    btn.title = 'Copy folder name';
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      onCopyFolderClick(tr, btn);
    });

    firstTd.appendChild(btn);
  }

  function injectButtonsInAllRows() {
    const rows = document.querySelectorAll(`${GRID_ID} .k-grid-content tbody tr[role="row"]`);
    rows.forEach(injectFolderButton);
  }

  let rafPending = false;
  function scheduleInject() {
    if (rafPending) return;
    rafPending = true;
    requestAnimationFrame(() => {
      rafPending = false;
      injectButtonsInAllRows();
    });
  }

  function bindGridDataBound() {
    const gridEl = document.querySelector(GRID_ID);
    if (!gridEl || !window.jQuery) return;
    const grid = window.jQuery(gridEl).data('kendoGrid');
    if (!grid || grid.__fehFolderCopyBound) return;
    grid.__fehFolderCopyBound = true;
    grid.bind('dataBound', scheduleInject);
  }

  function watchGridRows() {
    const tbody = document.querySelector(`${GRID_ID} .k-grid-content tbody`);
    if (!tbody || tbody.__fehFolderObserverAttached) return;

    const observer = new MutationObserver(scheduleInject);
    observer.observe(tbody, { childList: true, subtree: true });
    tbody.__fehFolderObserverAttached = true;
  }

  function enhance() {
    ensureStyle();
    bindGridDataBound();
    watchGridRows();
    scheduleInject();
  }

  function bootstrap() {
    enhance();

    const bodyObserver = new MutationObserver(() => {
      if (!document.querySelector(GRID_ID)) return;
      enhance();
    });
    bodyObserver.observe(document.body, { childList: true, subtree: true });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', bootstrap);
  } else {
    bootstrap();
  }
})();
