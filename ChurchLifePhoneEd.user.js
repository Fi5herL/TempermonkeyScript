// ==UserScript==
// @name         點名系統手機版
// @namespace    http://tampermonkey.net/
// @version      2025-04-19
// @description  try to take over the world!
// @author       Fisher Li
// @match        https://www.chlife-stat.org/
// @icon         https://www.google.com/s2/favicons?sz=64&domain=chlife-stat.org
// @grant        GM_addStyle
// @grant        GM_setValue
// @grant        GM_getValue
// ==/UserScript==

(function() {
    'use strict';

    const TABLE_ID = 'table';
    const FIXED_HEADER_TABLE_SELECTOR = 'table.fixed-header';
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
        /* --- 屏蔽複製出來的固定表頭 --- */
        ${FIXED_HEADER_TABLE_SELECTOR} { display: none !important; }

        /* --- 浮動按鈕樣式 --- */
        #${TOGGLE_BUTTON_ID} {
            position: fixed !important; bottom: 20px !important; right: 20px !important;
            padding: 10px 15px !important; background-color: #007bff !important; color: white !important;
            border: none !important; border-radius: 5px !important; cursor: pointer !important;
            z-index: 9999 !important; box-shadow: 2px 2px 5px rgba(0,0,0,0.2) !important;
            font-size: 14px !important; font-family: sans-serif !important;
        }
        #${TOGGLE_BUTTON_ID}:hover { background-color: #0056b3 !important; }

        /* --- 取消主表格 thead.normal 的固定 --- */
        #${TABLE_ID} thead.normal { position: static !important; top: auto !important; }

        /* --- 強制重設與基礎表格樣式 (針對 #table) --- */
        #${TABLE_ID}, #${TABLE_ID} *,
        #${TABLE_ID}::before, #${TABLE_ID}::after,
        #${TABLE_ID} *::before, #${TABLE_ID} *::after { all: revert !important; }

        #${TABLE_ID} {
            display: table !important; border-collapse: collapse !important; width: 100% !important;
            border: 1px solid #bbb !important; font-family: sans-serif !important; font-size: 13px !important;
            margin-top: 1em !important; margin-bottom: 1em !important; table-layout: fixed; /* *** 新增：嘗試固定佈局 *** */
        }

        #${TABLE_ID} thead { display: table-header-group !important; }
        #${TABLE_ID} tbody { display: table-row-group !important; }
        #${TABLE_ID} tr { display: table-row !important; }
        #${TABLE_ID} th, #${TABLE_ID} td {
             display: table-cell !important; /* *** 確保是 table-cell *** */
             word-wrap: break-word; /* *** 新增：允許長內容換行 *** */
        }

        #${TABLE_ID} thead.mobile { display: none !important; }

        #${TABLE_ID} th, #${TABLE_ID} td {
            border: 1px solid #bbb !important; padding: 4px 5px !important; /* 稍微減少 padding */
            text-align: left !important; vertical-align: middle !important; line-height: 1.3 !important; /* 稍微減小行高 */
        }

        #${TABLE_ID} thead.normal th {
            background-color: #e8e8e8 !important; font-weight: bold !important; white-space: nowrap !important;
        }

        /* --- 特定欄位樣式與對齊 (針對 #table) --- */
        /* NO. 欄 */
        #${TABLE_ID} thead.normal th:nth-child(1), #${TABLE_ID} tbody td:nth-child(1) {
            text-align: right !important;
            width: 3.5em !important; /* 縮小一點 */
        }
        /* 區別欄 */
        #${TABLE_ID} thead.normal th:nth-child(2), #${TABLE_ID} tbody td:nth-child(2) {
            text-align: center !important;
            width: 4em !important; /* *** 縮小寬度 *** */
        }
        /* 姓名欄 */
        #${TABLE_ID} thead.normal th:nth-child(3), #${TABLE_ID} tbody td:nth-child(3) {
            text-align: left !important;
            width: 5em !important; /* *** 縮小寬度 *** */
        }
        /* 性別欄 */
        #${TABLE_ID} thead.normal th:nth-child(4), #${TABLE_ID} tbody td:nth-child(4) {
            text-align: center !important;
            width: 3em !important; /* 縮小一點 */
        }

        /* Checkbox 相關樣式調整 */
        #${TABLE_ID} td.check-cell, #${TABLE_ID} thead.normal th:has(input.check-all) {
             text-align: center !important;
             width: 3em !important; /* Checkbox 列也窄一點 */
             padding: 3px !important;
        }
         #${TABLE_ID} thead.normal th:has(input.check-all) {
             font-size: 0.9em !important;
             white-space: normal !important; /* 表頭允許換行 */
             line-height: 1.2 !important;
             word-break: keep-all; /* 避免表頭單字內斷行 */
         }
         /* *** 移除 td.check-cell 的 flex *** */
         /* #table tbody td.check-cell { display: table-cell !important; } */ /* 保證是 table-cell */

         /* 注入的標籤樣式 (預設隱藏) */
         #${TABLE_ID} tbody td.check-cell .checkbox-label {
             display: none; /* 預設不顯示 */
             margin-right: 3px !important; /* 標籤和 checkbox 間距 */
             font-size: 0.9em !important;
             color: #333 !important;
             font-weight: bold;
             vertical-align: middle !important; /* 確保垂直對齊 */
         }
          /* Checkbox 本身樣式 */
         #${TABLE_ID} tbody td.check-cell input[type="checkbox"] {
              margin: 0 !important;
              vertical-align: middle !important; /* 確保垂直對齊 */
         }


        /* --- 簡化模式下的樣式 (針對 #table) --- */
        /* 隱藏 NO. (第一欄) */
        #${TABLE_ID}.${SIMPLIFIED_CLASS} thead.normal th:nth-child(1), #${TABLE_ID}.${SIMPLIFIED_CLASS} tbody td:nth-child(1) { display: none !important; }
        /* 隱藏 性別 (第四欄) */
        #${TABLE_ID}.${SIMPLIFIED_CLASS} thead.normal th:nth-child(4), #${TABLE_ID}.${SIMPLIFIED_CLASS} tbody td:nth-child(4) { display: none !important; }

        /* 簡化模式下 Checkbox 表頭 & 單元格 */
        #${TABLE_ID}.${SIMPLIFIED_CLASS} thead.normal th:has(input.check-all) {
             width: 2.8em !important; /* 可以非常窄 */
             white-space: nowrap !important;
             padding: 3px 2px !important; /* 減少左右 padding */
        }
         #${TABLE_ID}.${SIMPLIFIED_CLASS} tbody td.check-cell {
             /* justify-content: flex-start !important; */ /* 移除 flex 屬性 */
             text-align: left !important; /* 簡化模式下內容靠左 */
             padding-left: 4px !important;
             width: auto !important; /* 讓寬度自適應一點，但受 th 影響 */
         }
        #${TABLE_ID}.${SIMPLIFIED_CLASS} tbody td.check-cell .checkbox-label {
             display: inline-block !important; /* 在簡化模式下顯示標籤 (inline-block 更好對齊) */
        }
        /* 簡化模式下，Checkbox 稍微右移一點點 */
         #${TABLE_ID}.${SIMPLIFIED_CLASS} tbody td.check-cell input[type="checkbox"] {
             margin-left: 2px !important;
         }

    `);

    // --- JavaScript 功能函數 (與 v3.2 基本相同) ---

    let originalHeaderTexts = {}; // 只儲存原始文字

    // 更新表頭文字
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

    // 為 tbody 的 checkbox 單元格添加標籤 (只執行一次)
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


    // 更新按鈕文字
    function updateButtonText(button, isSimplified) {
        button.textContent = isSimplified ? '顯示完整表格' : '簡化表格顯示';
    }

    // --- 主邏輯 (與 v3.2 相同) ---
    const table = document.getElementById(TABLE_ID);
    if (table) {
         const normalThead = table.querySelector('thead.normal');
         if (!normalThead) { console.error('找不到 #table thead.normal'); return; }

         const toggleButton = document.createElement('button');
         toggleButton.id = TOGGLE_BUTTON_ID;
         document.body.appendChild(toggleButton);

         let isSimplified = GM_getValue(STORAGE_KEY, false);

         updateHeaderTexts(table, false);
         addLabelsToCheckboxCells(table);

         if (isSimplified) {
             table.classList.add(SIMPLIFIED_CLASS);
             updateHeaderTexts(table, true);
         } else {
             updateHeaderTexts(table, false);
         }
         updateButtonText(toggleButton, isSimplified);

         toggleButton.addEventListener('click', () => {
             isSimplified = !isSimplified;
             table.classList.toggle(SIMPLIFIED_CLASS, isSimplified);
             updateHeaderTexts(table, isSimplified);
             GM_setValue(STORAGE_KEY, isSimplified);
             updateButtonText(toggleButton, isSimplified);
         });
    } else {
        console.warn(`無法找到 ID 為 "${TABLE_ID}" 的主表格。`);
    }
})();
