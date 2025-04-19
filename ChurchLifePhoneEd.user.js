// ==UserScript==
// @name         表格樣式切換器 (含RWD)
// @namespace    http://tampermonkey.net/
// @version      2025-04-19
// @description  try to take over the world!
// @author       You
// @match        https://www.chlife-stat.org/
// @match        https://www.chlife-stat.org/index.php
// @icon         https://www.google.com/s2/favicons?sz=64&domain=chlife-stat.org
// @grant        GM_addStyle
// @grant        GM_setValue
// @grant        GM_getValue
// ==/UserScript==

(function() {
    'use strict';

    const TABLE_ID = 'table';
    const MENU_ID = 'menu'; // 目標菜單欄 ID
    const FIXED_HEADER_TABLE_SELECTOR = 'table.fixed-header';
    const FIXED_HEADER2_TABLE_SELECTOR = 'table.fixed-header2';
    const FIXED_BODY_TABLE_SELECTOR = 'table.fixed-body';
    const TOGGLE_BUTTON_ID = 'table-style-toggle-button';
    const SIMPLIFIED_CLASS = 'simplified-table-view';
    const STORAGE_KEY = 'tableSimplifiedPreference';
    const LABEL_ADDED_CLASS = 'label-added';

    // 簡化標題的映射
    const headerSimplificationMap = {
        "主日": "主", "禱告": "禱", "家聚會出訪": "出", "家聚會受訪": "受",
        "小排": "排", "晨興": "興", "福音出訪": "福", "生命讀經": "生", "主日申言": "申"
    };
    let simplifiedHeaderLabels = [];

    // --- CSS 樣式 (強制覆蓋) ---
    GM_addStyle(`
        /* --- 屏蔽複製出來的固定元素 --- */
        ${FIXED_HEADER_TABLE_SELECTOR},
        ${FIXED_HEADER2_TABLE_SELECTOR},
        ${FIXED_BODY_TABLE_SELECTOR} {
            display: none !important;
        }

        /* --- 樣式切換按鈕 (放入菜單欄) --- */
        #${MENU_ID} #${TOGGLE_BUTTON_ID} { /* 限定在 #menu 內的按鈕 */
            display: inline-block !important; /* 與菜單項並排 */
            position: relative !important; /* 相對定位，覆蓋 fixed */
            top: auto !important; /* 移除 top 定位 */
            left: auto !important; /* 移除 left 定位 */
            transform: none !important; /* 移除 transform */
            margin: 2px 8px !important; /* 在菜單項之間添加邊距 */
            padding: 5px 10px !important; /* 調整內邊距使其接近菜單項 */
            background-color: #e0e0e0 !important; /* 淺灰色背景，更像按鈕 */
            color: #333 !important; /* 深色文字 */
            border: 1px solid #ccc !important; /* 添加邊框 */
            border-radius: 4px !important;
            cursor: pointer !important;
            z-index: 1 !important; /* 在菜單內，不需要太高 z-index */
            box-shadow: none !important; /* 移除陰影 */
            font-size: 13px !important; /* 匹配菜單字體大小 */
            font-family: sans-serif !important;
            white-space: nowrap !important;
            vertical-align: middle; /* 垂直居中對齊 */
            line-height: normal; /* 正常行高 */
        }
        #${MENU_ID} #${TOGGLE_BUTTON_ID}:hover {
            background-color: #d0d0d0 !important; /* 懸停時變深一點 */
            border-color: #bbb !important;
        }
        #${MENU_ID} #${TOGGLE_BUTTON_ID}:active { /* 點擊效果 */
             background-color: #c0c0c0 !important;
             border-color: #aaa !important;
        }

        /* --- 取消主表格 thead.normal 的固定 --- */
        #${TABLE_ID} thead.normal { position: static !important; top: auto !important; }

        /* --- 強制重設與基礎表格樣式 (針對 #table) --- */
        #${TABLE_ID}, #${TABLE_ID} *,
        #${TABLE_ID}::before, #${TABLE_ID}::after,
        #${TABLE_ID} *::before, #${TABLE_ID} *::after { all: revert !important; }

        #${TABLE_ID} {
            display: table !important; border-collapse: collapse !important; width: 100% !important;
            border: 1px solid #bbb !important; font-family: sans-serif !important; font-size: 13px !important;
            margin-top: 1em !important; margin-bottom: 1em !important; table-layout: fixed;
        }

        #${TABLE_ID} thead { display: table-header-group !important; }
        #${TABLE_ID} tbody { display: table-row-group !important; }
        #${TABLE_ID} tr { display: table-row !important; }
        #${TABLE_ID} th, #${TABLE_ID} td {
             display: table-cell !important; word-wrap: break-word;
        }

        #${TABLE_ID} thead.mobile { display: none !important; }

        #${TABLE_ID} th, #${TABLE_ID} td {
            border: 1px solid #bbb !important; padding: 4px 5px !important;
            text-align: left !important; vertical-align: middle !important; line-height: 1.3 !important;
        }

        #${TABLE_ID} thead.normal th {
            background-color: #e8e8e8 !important; font-weight: bold !important; white-space: nowrap !important;
        }

        /* 特定欄位樣式與對齊 (針對 #table) */
        #${TABLE_ID} thead.normal th:nth-child(1), #${TABLE_ID} tbody td:nth-child(1) { text-align: right !important; width: 3.5em !important; }
        #${TABLE_ID} thead.normal th:nth-child(2), #${TABLE_ID} tbody td:nth-child(2) { text-align: center !important; width: 4em !important; }
        #${TABLE_ID} thead.normal th:nth-child(3), #${TABLE_ID} tbody td:nth-child(3) { text-align: left !important; width: 5em !important; }
        #${TABLE_ID} thead.normal th:nth-child(4), #${TABLE_ID} tbody td:nth-child(4) { text-align: center !important; width: 3em !important; }

        #${TABLE_ID} td.check-cell, #${TABLE_ID} thead.normal th:has(input.check-all) {
             text-align: center !important; width: 3em !important; padding: 3px !important;
        }
         #${TABLE_ID} thead.normal th:has(input.check-all) { font-size: 0.9em !important; white-space: normal !important; line-height: 1.2 !important; word-break: keep-all; }
         #${TABLE_ID} tbody td.check-cell .checkbox-label {
             display: none; margin-right: 3px !important; font-size: 0.9em !important; color: #333 !important; font-weight: bold; vertical-align: middle !important;
         }
         #${TABLE_ID} tbody td.check-cell input[type="checkbox"] { margin: 0 !important; vertical-align: middle !important; }

        /* --- 簡化模式下的樣式 (針對 #table) --- */
        #${TABLE_ID}.${SIMPLIFIED_CLASS} thead.normal th:nth-child(1), #${TABLE_ID}.${SIMPLIFIED_CLASS} tbody td:nth-child(1) { display: none !important; }
        #${TABLE_ID}.${SIMPLIFIED_CLASS} thead.normal th:nth-child(4), #${TABLE_ID}.${SIMPLIFIED_CLASS} tbody td:nth-child(4) { display: none !important; }

        #${TABLE_ID}.${SIMPLIFIED_CLASS} thead.normal th:has(input.check-all) { width: 2.8em !important; white-space: nowrap !important; padding: 3px 2px !important; }
         #${TABLE_ID}.${SIMPLIFIED_CLASS} tbody td.check-cell { text-align: left !important; padding-left: 4px !important; width: auto !important; }
        #${TABLE_ID}.${SIMPLIFIED_CLASS} tbody td.check-cell .checkbox-label { display: inline-block !important; }
         #${TABLE_ID}.${SIMPLIFIED_CLASS} tbody td.check-cell input[type="checkbox"] { margin-left: 2px !important; }
    `);

    // --- JavaScript 功能函數 (與 v3.3 相同) ---
    // ... (省略重複的 JS 代碼，請確保包含) ...
    // <editor-fold desc="v3.3 JS Functions">
    let originalHeaderTexts = {};
    function updateHeaderTexts(tableElement, simplify) {
        if (!tableElement) return;
        const headerCells = tableElement.querySelectorAll('thead.normal th');
        simplifiedHeaderLabels = []; // 清空/初始化每次更新

        headerCells.forEach((th, index) => {
             let textNode = null;
             let originalText = '';
             let inputEl = th.querySelector('input.check-all');

             // 獲取原始文字並儲存
             if (!originalHeaderTexts[index]) {
                 let textNodeFound = false;
                 if (inputEl) { // checkbox 表頭
                     let currentNode = inputEl.nextSibling;
                     while(currentNode) {
                         if (currentNode.nodeType === Node.TEXT_NODE && currentNode.textContent.trim()) {
                             originalText = currentNode.textContent.trim(); textNodeFound = true; textNode = currentNode; break;
                         }
                          if (currentNode.nodeName === 'BR') {
                              let afterBrNode = currentNode.nextSibling;
                              if (afterBrNode && afterBrNode.nodeType === Node.TEXT_NODE && afterBrNode.textContent.trim()) {
                                  originalText = afterBrNode.textContent.trim(); textNodeFound = true; textNode = afterBrNode;
                              }
                              break;
                          }
                         currentNode = currentNode.nextSibling;
                     }
                 }
                 if (!textNodeFound) { // 非 checkbox 表頭或 checkbox 後無文字
                    for (const node of th.childNodes) {
                        if (node.nodeType === Node.TEXT_NODE && node.textContent.trim()) {
                            originalText = node.textContent.trim(); textNodeFound = true; textNode = node; break;
                        }
                    }
                 }
                 if (!textNodeFound) { // Fallback
                     originalText = th.textContent.trim().split(/[\n\r]/)[0].trim();
                 }
                 originalHeaderTexts[index] = originalText;
             } else {
                 originalText = originalHeaderTexts[index];
                 // 仍然需要找到 textNode 以便修改
                 if (inputEl) {
                    let currentNode = inputEl.nextSibling;
                     while(currentNode) {
                         if (currentNode.nodeType === Node.TEXT_NODE && currentNode.textContent.trim()) {
                             textNode = currentNode; break; }
                         if (currentNode.nodeName === 'BR') {
                             let afterBrNode = currentNode.nextSibling;
                             if (afterBrNode && afterBrNode.nodeType === Node.TEXT_NODE && afterBrNode.textContent.trim()) { textNode = afterBrNode; }
                             break;
                         }
                         currentNode = currentNode.nextSibling;
                     }
                 }
                 if (!textNode) {
                     for (const node of th.childNodes) {
                        if (node.nodeType === Node.TEXT_NODE && node.textContent.trim()) {
                            textNode = node; break; }
                     }
                 }
             }

             // 更新文字內容並記錄簡化標籤
             let targetText = originalText;
             let simplifiedLabel = null; // 用於 tbody

             if (simplify) {
                 let simplified = false;
                 for (const key in headerSimplificationMap) {
                     if (originalText.includes(key)) {
                         targetText = headerSimplificationMap[key];
                         simplifiedLabel = targetText; // 儲存簡化標籤
                         simplified = true;
                         break;
                     }
                 }
             } else {
                 targetText = originalText; // 還原
                 for (const key in headerSimplificationMap) {
                     if (originalText.includes(key)) {
                         simplifiedLabel = headerSimplificationMap[key]; break;
                     }
                 }
             }
             simplifiedHeaderLabels[index] = simplifiedLabel;

             // 更新表頭的顯示文字
             if (textNode) {
                 textNode.textContent = ' ' + targetText + ' ';
             }
        });
    }

    function addLabelsToCheckboxCells(tableElement) {
        if (!tableElement) return;
        const bodyRows = tableElement.querySelectorAll('tbody#members tr.member');

        if (simplifiedHeaderLabels.length === 0) {
             updateHeaderTexts(tableElement, false);
             if (simplifiedHeaderLabels.length === 0) { console.warn("無法生成表頭標籤。"); return; }
        }

        bodyRows.forEach(row => {
            const cells = row.querySelectorAll('td');
            cells.forEach((cell, cellIndex) => {
                if (!cell.classList.contains('check-cell')) return;
                if (cell.classList.contains(LABEL_ADDED_CLASS)) return;

                const checkbox = cell.querySelector('input[type="checkbox"]');
                if (!checkbox) return;

                const labelText = simplifiedHeaderLabels[cellIndex];

                if (labelText) {
                    const labelSpan = document.createElement('span');
                    labelSpan.className = 'checkbox-label';
                    labelSpan.textContent = labelText;
                    cell.insertBefore(labelSpan, checkbox);
                    cell.classList.add(LABEL_ADDED_CLASS);
                }
            });
        });
         console.log("Labels added to checkbox cells.");
    }

    function updateButtonText(button, isSimplified) {
        button.textContent = isSimplified ? '顯示完整表格' : '簡化表格顯示';
    }
    // </editor-fold>

    // --- 主邏輯 (修改按鈕添加位置) ---
    const table = document.getElementById(TABLE_ID);
    const menuDiv = document.getElementById(MENU_ID); // *** 獲取菜單欄元素 ***

    if (table && menuDiv) { // *** 確保表格和菜單欄都存在 ***
         const normalThead = table.querySelector('thead.normal');
         if (!normalThead) { console.error('找不到 #table thead.normal'); return; }

         // 創建按鈕
         const toggleButton = document.createElement('button');
         toggleButton.id = TOGGLE_BUTTON_ID;
         // *** 將按鈕添加到菜單欄末尾 ***
         menuDiv.appendChild(toggleButton);

         // 讀取偏好
         let isSimplified = GM_getValue(STORAGE_KEY, false);

         // 初始化
         updateHeaderTexts(table, false);
         addLabelsToCheckboxCells(table);

         if (isSimplified) {
             table.classList.add(SIMPLIFIED_CLASS);
             updateHeaderTexts(table, true);
         } else {
             updateHeaderTexts(table, false);
         }
         updateButtonText(toggleButton, isSimplified);

         // 按鈕事件
         toggleButton.addEventListener('click', () => {
             isSimplified = !isSimplified;
             table.classList.toggle(SIMPLIFIED_CLASS, isSimplified);
             updateHeaderTexts(table, isSimplified);
             GM_setValue(STORAGE_KEY, isSimplified);
             updateButtonText(toggleButton, isSimplified);
         });
    } else {
         if (!table) console.warn(`無法找到 ID 為 "${TABLE_ID}" 的主表格。`);
         if (!menuDiv) console.warn(`無法找到 ID 為 "${MENU_ID}" 的菜單欄。按鈕將不會被添加。`);
    }

})();
