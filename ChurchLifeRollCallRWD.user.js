// ==UserScript==
// @name         ChurchLife 點名 RWD 助手 (Apple 風格)
// @namespace    http://tampermonkey.net/
// @version      1.0.0
// @description  在點名頁加入簡潔 RWD 面板，支援人名查詢、快速點名與即時人數統計（電腦/手機皆可用）。
// @author       Fisher Li & AI Assistant
// @match        https://www.chlife-stat.org/index.php*
// @match        https://www.chlife-stat.org/
// @grant        GM_addStyle
// ==/UserScript==

(function () {
    'use strict';

    const TABLE_SELECTOR = '#roll-call-panel table#table';
    const PANEL_ID = 'tm-rollcall-rwd-panel';
    const FIXED_COLUMNS_COUNT = 4;
    // Conservative cap to avoid long main-thread UI stalls on mobile when matched rows are very large.
    const LIST_LIMIT = 300;
    const RENDER_DEBOUNCE_DELAY = 120;
    const UNNAMED_TEXT = '(未命名)';
    const UI_TEXT = {
        EMPTY_RESULT: '查無符合資料',
        LIST_LIMIT_HINT: `僅顯示前 ${LIST_LIMIT} 筆，請再縮小搜尋條件`
    };

    const state = {
        query: '',
        filter: 'all',
        attendanceIndex: 0,
        collapsed: false
    };

    let refs = {};

    function normalize(value) {
        return String(value || '')
            .toLowerCase()
            .replace(/\s+/g, '')
            .trim();
    }

    function getTable() {
        return document.querySelector(TABLE_SELECTOR);
    }

    function getRows() {
        const rows = Array.from(document.querySelectorAll(`${TABLE_SELECTOR} tbody tr`));
        rows.shift();
        return rows.filter((row) => row.querySelectorAll('td').length > FIXED_COLUMNS_COUNT);
    }

    function getAttendanceHeaders(table) {
        const headerCells = Array.from(table.querySelectorAll('thead th'));
        return headerCells
            .slice(FIXED_COLUMNS_COUNT)
            .map((th, index) => th.textContent.trim() || `欄位 ${index + 1}`);
    }

    function getPeopleData() {
        return getRows().map((row) => {
            const cells = Array.from(row.querySelectorAll('td'));
            const checkboxes = cells
                .slice(FIXED_COLUMNS_COUNT)
                .map((cell) => cell.querySelector('input[type="checkbox"]'))
                .filter(Boolean);
            const targetCheckbox = checkboxes[state.attendanceIndex] || null;
            return {
                row,
                no: cells[0]?.textContent.trim() || '',
                distinction: cells[1]?.textContent.trim() || '',
                name: cells[2]?.textContent.trim() || '',
                sex: cells[3]?.textContent.trim() || '',
                checkboxes,
                targetCheckbox
            };
        });
    }

    function isMatched(person, query) {
        if (!query) return true;
        const pool = normalize(`${person.no} ${person.distinction} ${person.name} ${person.sex}`);
        return pool.includes(query);
    }

    function isSelected(person) {
        return !!person.targetCheckbox?.checked;
    }

    function setCheckboxChecked(checkbox, checked) {
        if (!checkbox || checkbox.checked === checked) return;
        checkbox.click();
    }

    function renderFilters() {
        refs.filterButtons.forEach((button) => {
            const isActive = button.dataset.filter === state.filter;
            button.classList.toggle('active', isActive);
        });
    }

    function renderHeaderOptions() {
        const table = getTable();
        if (!table) return;
        const headers = getAttendanceHeaders(table);
        if (!headers.length) return;

        if (state.attendanceIndex >= headers.length) {
            state.attendanceIndex = 0;
        }

        refs.attendanceSelect.replaceChildren();
        headers.forEach((header, index) => {
            const option = document.createElement('option');
            option.value = String(index);
            option.textContent = header;
            refs.attendanceSelect.appendChild(option);
        });
        refs.attendanceSelect.value = String(state.attendanceIndex);
    }

    function renderList(people) {
        if (!people.length) {
            const empty = document.createElement('div');
            empty.className = 'tm-rollcall-empty';
            empty.textContent = UI_TEXT.EMPTY_RESULT;
            refs.list.replaceChildren(empty);
            return;
        }

        const fragment = document.createDocumentFragment();
        // Apply render cap to keep DOM updates smooth on mobile/low-end devices.
        people.slice(0, LIST_LIMIT).forEach((person) => {
            const item = document.createElement('div');
            item.className = 'tm-rollcall-item';

            const info = document.createElement('div');
            info.className = 'tm-rollcall-item-info';
            const nameNode = document.createElement('div');
            nameNode.className = 'tm-rollcall-item-name';
            nameNode.textContent = person.name || UNNAMED_TEXT;
            const subNode = document.createElement('div');
            subNode.className = 'tm-rollcall-item-sub';
            subNode.textContent = `${person.no}｜${person.distinction}｜${person.sex}`;
            info.appendChild(nameNode);
            info.appendChild(subNode);

            const toggle = document.createElement('button');
            const selected = isSelected(person);
            toggle.className = `tm-rollcall-item-toggle ${selected ? 'selected' : ''}`;
            toggle.textContent = selected ? '已點名' : '點名';
            toggle.addEventListener('click', () => {
                setCheckboxChecked(person.targetCheckbox, !selected);
                scheduleRender();
            });

            item.appendChild(info);
            item.appendChild(toggle);
            fragment.appendChild(item);
        });

        if (people.length > LIST_LIMIT) {
            const tip = document.createElement('div');
            tip.className = 'tm-rollcall-empty';
            tip.textContent = UI_TEXT.LIST_LIMIT_HINT;
            fragment.appendChild(tip);
        }

        refs.list.replaceChildren(fragment);
    }

    function render() {
        const table = getTable();
        if (!table) {
            if (refs.panel) refs.panel.style.display = 'none';
            return;
        }
        refs.panel.style.display = '';

        renderHeaderOptions();
        renderFilters();

        const allPeople = getPeopleData();
        const normalizedQuery = normalize(state.query);
        let matched = allPeople.filter((person) => isMatched(person, normalizedQuery));

        if (state.filter === 'checked') {
            matched = matched.filter((person) => isSelected(person));
        } else if (state.filter === 'unchecked') {
            matched = matched.filter((person) => !isSelected(person));
        }

        const checkedCount = allPeople.filter((person) => isSelected(person)).length;
        refs.totalCount.textContent = String(allPeople.length);
        refs.filteredCount.textContent = String(matched.length);
        refs.checkedCount.textContent = String(checkedCount);

        renderList(matched);
    }

    function createMetric(label, valueRefName) {
        const box = document.createElement('div');
        box.className = 'tm-rollcall-metric';
        const value = document.createElement('div');
        value.className = 'tm-rollcall-metric-value';
        value.textContent = '0';
        const title = document.createElement('div');
        title.className = 'tm-rollcall-metric-label';
        title.textContent = label;
        box.appendChild(value);
        box.appendChild(title);
        refs[valueRefName] = value;
        return box;
    }

    function buildPanel() {
        const panel = document.createElement('section');
        panel.id = PANEL_ID;
        panel.innerHTML = `
            <div class="tm-rollcall-header">
                <div class="tm-rollcall-title-wrap">
                    <h2 class="tm-rollcall-title">點名助手</h2>
                    <p class="tm-rollcall-subtitle">RWD 查詢與快速點名</p>
                </div>
                <button class="tm-rollcall-collapse" type="button" aria-label="收合面板">－</button>
            </div>
            <div class="tm-rollcall-body">
                <div class="tm-rollcall-controls">
                    <input class="tm-rollcall-search" type="search" placeholder="搜尋姓名 / NO. / 區別 / 性別" aria-label="搜尋姓名、編號、區別或性別" />
                    <div class="tm-rollcall-select-wrap">
                        <label class="tm-rollcall-inline-label" for="tm-rollcall-select">點名欄位</label>
                        <select id="tm-rollcall-select" class="tm-rollcall-select" aria-label="選擇點名欄位"></select>
                    </div>
                </div>
                <div class="tm-rollcall-filter-row">
                     <button type="button" data-filter="all" class="active">全部</button>
                     <button type="button" data-filter="checked">已點名</button>
                     <button type="button" data-filter="unchecked">未點名</button>
                 </div>
                <div class="tm-rollcall-metrics"></div>
                 <div class="tm-rollcall-action-row" style="display:none">
                     <button type="button" class="tm-rollcall-bulk-check">符合全點名</button>
                     <button type="button" class="tm-rollcall-bulk-uncheck">符合全取消</button>
                     <button type="button" class="tm-rollcall-refresh">重新整理</button>
                 </div>
                <div class="tm-rollcall-list"></div>
            </div>
        `;


        document.body.appendChild(panel);

        refs.panel = panel;
        refs.body = panel.querySelector('.tm-rollcall-body');
        refs.search = panel.querySelector('.tm-rollcall-search');
        refs.attendanceSelect = panel.querySelector('.tm-rollcall-select');
        refs.filterButtons = Array.from(panel.querySelectorAll('.tm-rollcall-filter-row button'));
        refs.bulkCheck = panel.querySelector('.tm-rollcall-bulk-check');
        refs.bulkUncheck = panel.querySelector('.tm-rollcall-bulk-uncheck');
        refs.refresh = panel.querySelector('.tm-rollcall-refresh');
        refs.list = panel.querySelector('.tm-rollcall-list');
        refs.collapse = panel.querySelector('.tm-rollcall-collapse');

        const metrics = panel.querySelector('.tm-rollcall-metrics');
        metrics.appendChild(createMetric('總人數', 'totalCount'));
        metrics.appendChild(createMetric('符合搜尋', 'filteredCount'));
        metrics.appendChild(createMetric('已點名', 'checkedCount'));
    }

    function applyEvents() {
        refs.search.addEventListener('input', (event) => {
            state.query = event.target.value || '';
            scheduleRender();
        });

        refs.attendanceSelect.addEventListener('change', (event) => {
            state.attendanceIndex = Number(event.target.value) || 0;
            render();
        });

        refs.filterButtons.forEach((button) => {
            button.addEventListener('click', () => {
                state.filter = button.dataset.filter || 'all';
                render();
            });
        });

        refs.bulkCheck.addEventListener('click', () => {
            const normalizedQuery = normalize(state.query);
            getPeopleData()
                .filter((person) => isMatched(person, normalizedQuery))
                .forEach((person) => setCheckboxChecked(person.targetCheckbox, true));
            scheduleRender();
        });

        refs.bulkUncheck.addEventListener('click', () => {
            const normalizedQuery = normalize(state.query);
            getPeopleData()
                .filter((person) => isMatched(person, normalizedQuery))
                .forEach((person) => setCheckboxChecked(person.targetCheckbox, false));
            scheduleRender();
        });

        refs.refresh.addEventListener('click', () => {
            render();
        });

        refs.collapse.addEventListener('click', () => {
            state.collapsed = !state.collapsed;
            refs.panel.classList.toggle('collapsed', state.collapsed);
            refs.collapse.textContent = state.collapsed ? '＋' : '－';
            refs.collapse.setAttribute('aria-expanded', String(!state.collapsed));
            refs.collapse.setAttribute('aria-label', state.collapsed ? '展開面板' : '收合面板');
        });
    }

    let renderTimer = null;
    function scheduleRender() {
        clearTimeout(renderTimer);
        renderTimer = setTimeout(render, RENDER_DEBOUNCE_DELAY);
    }

    function observeTableChanges() {
        const observerTarget = document.querySelector('#roll-call-panel') || document.body;
        const isRelevantElement = (element) =>
            element.closest('#roll-call-panel') || element.closest('#pagination');

        const observer = new MutationObserver((mutations) => {
            const shouldRender = mutations.some((mutation) => {
                const target = mutation.target;
                if (!(target instanceof Element)) return false;
                if (isRelevantElement(target)) return true;
                if (mutation.type === 'childList') {
                    const nodes = [...mutation.addedNodes, ...mutation.removedNodes];
                    return nodes.some((node) => node instanceof Element && isRelevantElement(node));
                }
                return false;
            });
            if (shouldRender) scheduleRender();
        });
        observer.observe(observerTarget, {
            childList: true,
            subtree: true
        });
    }

    function addStyles() {
        GM_addStyle(`
            #${PANEL_ID} {
                position: fixed;
                right: 16px;
                bottom: 5px;
                z-index: 99999;
                width: min(500px, calc(100vw - 5px));
                max-height: calc(100vh - 10px);
                background: rgba(255, 255, 255, 0.85);
                backdrop-filter: blur(14px);
                -webkit-backdrop-filter: blur(14px);
                border: 1px solid rgba(255, 255, 255, 0.55);
                border-radius: 22px;
                box-shadow: 0 14px 30px rgba(15, 23, 42, 0.2);
                font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "PingFang TC", "Microsoft JhengHei", "Segoe UI", sans-serif;
                color: #1d1d1f;
                overflow: hidden;
            }

            #${PANEL_ID} * { box-sizing: border-box; }

            #${PANEL_ID} .tm-rollcall-header {
                display: flex;
                justify-content: space-between;
                gap: 8px;
                padding: 14px 14px 10px;
                border-bottom: 1px solid rgba(0, 0, 0, 0.07);
            }

            #${PANEL_ID} .tm-rollcall-title { margin: 0; font-size: 17px; font-weight: 700; letter-spacing: -0.01em; }
            #${PANEL_ID} .tm-rollcall-subtitle { margin: 2px 0 0; font-size: 12px; color: #6e6e73; }
            #${PANEL_ID} .tm-rollcall-collapse {
                width: 28px; height: 28px; border: 0; border-radius: 999px; cursor: pointer;
                background: rgba(0, 0, 0, 0.07); color: #1d1d1f; font-size: 18px; line-height: 1;
            }

            #${PANEL_ID} .tm-rollcall-body { padding: 10px 12px 12px; }
            #${PANEL_ID}.collapsed .tm-rollcall-body { display: none; }

            #${PANEL_ID} .tm-rollcall-controls { display: grid; grid-template-columns: 1fr 122px; gap: 8px; margin-bottom: 8px; }
            #${PANEL_ID} .tm-rollcall-select-wrap { display: flex; flex-direction: column; gap: 3px; }
            #${PANEL_ID} .tm-rollcall-inline-label { font-size: 11px; color: #6e6e73; line-height: 1; padding-left: 2px; }
            #${PANEL_ID} input, #${PANEL_ID} select, #${PANEL_ID} button {
                border: 1px solid rgba(60, 60, 67, 0.18);
                border-radius: 12px;
                background: rgba(255, 255, 255, 0.92);
                color: #1d1d1f;
                font-size: 13px;
            }

            #${PANEL_ID} .tm-rollcall-search, #${PANEL_ID} .tm-rollcall-select { padding: 8px 10px; }
            #${PANEL_ID} .tm-rollcall-filter-row, #${PANEL_ID} .tm-rollcall-action-row {
                display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; margin-bottom: 8px;
            }

            #${PANEL_ID} button {
                padding: 7px 8px;
                cursor: pointer;
                transition: all 0.15s ease;
            }

            #${PANEL_ID} .tm-rollcall-filter-row button.active,
            #${PANEL_ID} .tm-rollcall-item-toggle.selected,
            #${PANEL_ID} .tm-rollcall-bulk-check {
                background: #0071e3;
                color: #fff;
                border-color: #0071e3;
            }

            #${PANEL_ID} .tm-rollcall-metrics {
                display: grid;
                grid-template-columns: repeat(3, 1fr);
                gap: 8px;
                margin-bottom: 8px;
            }

            #${PANEL_ID} .tm-rollcall-metric {
                padding: 9px;
                border-radius: 14px;
                background: rgba(246, 246, 247, 0.9);
                border: 1px solid rgba(0, 0, 0, 0.04);
                text-align: center;
            }
            #${PANEL_ID} .tm-rollcall-metric-value { font-size: 18px; font-weight: 700; line-height: 1.1; }
            #${PANEL_ID} .tm-rollcall-metric-label { margin-top: 3px; font-size: 11px; color: #6e6e73; }

            #${PANEL_ID} .tm-rollcall-list {
                max-height: min(46vh, 420px);
                overflow: auto;
                display: flex;
                flex-direction: column;
                gap: 7px;
                padding-right: 2px;
            }

            #${PANEL_ID} .tm-rollcall-item {
                display: grid;
                grid-template-columns: 1fr auto;
                align-items: center;
                gap: 8px;
                border: 1px solid rgba(0, 0, 0, 0.06);
                border-radius: 12px;
                padding: 8px;
                background: rgba(255, 255, 255, 0.96);
            }
            #${PANEL_ID} .tm-rollcall-item-name { font-size: 14px; font-weight: 600; line-height: 1.2; }
            #${PANEL_ID} .tm-rollcall-item-sub { margin-top: 2px; color: #6e6e73; font-size: 11px; }
            #${PANEL_ID} .tm-rollcall-item-toggle { min-width: 70px; padding: 6px 10px; border-radius: 999px; }
            #${PANEL_ID} .tm-rollcall-empty {
                text-align: center;
                font-size: 12px;
                color: #6e6e73;
                padding: 10px 8px;
            }

            @media (max-width: 760px) {
                #${PANEL_ID} {
                    right: 10px;
                    left: 10px;
                    width: auto;
                    bottom: max(10px, env(safe-area-inset-bottom));
                    max-height: calc(100vh - 18px);
                }
                #${PANEL_ID} .tm-rollcall-controls { grid-template-columns: 1fr; }
                #${PANEL_ID} .tm-rollcall-action-row { grid-template-columns: 1fr; }
                #${PANEL_ID} .tm-rollcall-list { max-height: min(44vh, 360px); }
            }
        `);
    }

    function init() {
        addStyles();
        buildPanel();
        applyEvents();
        observeTableChanges();
        render();
        refs.collapse.setAttribute('aria-expanded', 'true');
        refs.collapse.setAttribute('aria-label', '收合面板');
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
