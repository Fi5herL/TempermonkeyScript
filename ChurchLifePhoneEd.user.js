// ==UserScript==
// @name         表格樣式切換器 (v4.0 - 隱藏表頭勾選 & 固定前三欄)
// @namespace    http://tampermonkey.net/
// @version      4.0
// @description  強制替換 #table 樣式，屏蔽固定元素，按鈕置頂中，固定前三欄，隱藏表頭勾選，並提供按鈕切換簡化顯示。
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
    const MENU_ID = 'menu';
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
        "小排": "排", "晨興": "興", "福音出訪": "福", "生命讀經": "讀", "主日申言": "申"
    };
    let simplifiedHeaderLabels = [];

    // --- CSS 樣式 ---
    GM_addStyle(`
        /* --- 屏蔽複製出來的固定元素 --- */
        ${FIXED_HEADER_TABLE_SELECTOR},
        ${FIXED_HEADER2_TABLE_SELECTOR},
        ${FIXED_BODY_TABLE_SELECTOR} {
            display: none !important;
        }

        /* --- 樣式切換按鈕 (放入菜單欄) --- */
        #${MENU_ID} #${TOGGLE_BUTTON_ID} {
            display: inline-block !important; position: relative !important; top: auto !important; left: auto !important;
            transform: none !important; margin: 2px 8px !important; padding: 5px 10px !important;
            background-color: #e0e0e0 !important; color: #333 !important; border: 1px solid #ccc !important;
            border-radius: 4px !important; cursor: pointer !important; z-index: 1 !important; box-shadow: none !important;
            font-size: 13px !important; font-family: sans-serif !important; white-space: nowrap !important;
            vertical-align: middle; line-height: normal;
        }
        #${MENU_ID} #${TOGGLE_BUTTON_ID}:hover { background-color: #d0d0d0 !important; border-color: #bbb !important; }
        #${MENU_ID} #${TOGGLE_BUTTON_ID}:active { background-color: #c0c0c0 !important; border-color: #aaa !important; }

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
            background-color: white !important; /* *** 基礎背景設為白色 *** */
        }

        /* 表頭樣式 & 垂直固定 */
        #${TABLE_ID} thead.normal {
            position: sticky !important; top: 0 !important; z-index: 4 !important;
        }
        #${TABLE_ID} thead.normal th {
            background-color: #e8e8e8 !important; font-weight: bold !important;
            white-space: nowrap !important; z-index: 3 !important;
        }
        /* *** 隱藏表頭的 Checkbox *** */
        #${TABLE_ID} thead.normal th input.check-all {
            display: none !important;
        }

        /* --- 固定欄位 (NO., 區別, 姓名) --- */
        /* 固定第 1 欄 (NO.) */
        #${TABLE_ID} thead.normal th:nth-child(1),
        #${TABLE_ID} tbody td:nth-child(1) {
            position: sticky !important;
            left: 0px !important; /* 固定在最左側 */
            width: 3.5em !important; /* 保持寬度定義 */
            text-align: right !important;
            z-index: 2 !important;
            background-color: white !important;
        }
         #${TABLE_ID} thead.normal th:nth-child(1) { background-color: #e8e8e8 !important; z-index: 3; }
         /* #tbody tr.r1 td:nth-child(1) { background-color: #f9f9f9 !important; } */
         /* #tbody tr.r2 td:nth-child(1) { background-color: white !important; } */

        /* 固定第 2 欄 (區別) */
        #${TABLE_ID} thead.normal th:nth-child(2),
        #${TABLE_ID} tbody td:nth-child(2) {
            position: sticky !important;
            left: 3.5em !important; /* *** 左側偏移量 = 第 1 欄的寬度 *** */
            width: 4em !important; /* 保持寬度定義 */
            text-align: center !important;
            z-index: 2 !important;
            background-color: white !important;
        }
         #${TABLE_ID} thead.normal th:nth-child(2) { background-color: #e8e8e8 !important; z-index: 3; }
         /* #tbody tr.r1 td:nth-child(2) { background-color: #f9f9f9 !important; } */
         /* #tbody tr.r2 td:nth-child(2) { background-color: white !important; } */

        /* 固定第 3 欄 (姓名) */
        #${TABLE_ID} thead.normal th:nth-child(3),
        #${TABLE_ID} tbody td:nth-child(3) {
            position: sticky !important;
            left: 7.5em !important; /* *** 左側偏移量 = 第 1 欄寬 (3.5em) + 第 2 欄寬 (4em) *** */
            width: 4em !important; /* 保持寬度定義 */
            text-align: left !important;
            z-index: 2 !important;
            background-color: white !important;
        }
        #${TABLE_ID} thead.normal th:nth-child(3) { background-color: #e8e8e8 !important; z-index: 3; }
        /* #tbody tr.r1 td:nth-child(3) { background-color: #f9f9f9 !important; } */
        /* #tbody tr.r2 td:nth-child(3) { background-color: white !important; } */


        /* --- 其他欄位樣式 (確保不衝突) --- */
        /* 性別欄 */
        #${TABLE_ID} thead.normal th:nth-child(4), #${TABLE_ID} tbody td:nth-child(4) { text-align: center !important; width: 3em !important; }

        /* Checkbox 相關 */
        #${TABLE_ID} td.check-cell, #${TABLE_ID} thead.normal th:has(input.check-all) {
             text-align: center !important; width: 3em !important; padding: 3px !important;
        }
         /* *** 移除表頭Checkbox後，調整表頭樣式（例如不需要換行了） *** */
         #${TABLE_ID} thead.normal th:has(input.check-all) {
             font-size: 0.9em !important;
             white-space: nowrap !important; /* 簡化後文字不換行 */
             line-height: 1.2 !important;
             word-break: keep-all;
         }
         #${TABLE_ID} tbody td.check-cell .checkbox-label { display: none; margin-right: 3px !important; font-size: 0.9em !important; color: #333 !important; font-weight: bold; vertical-align: middle !important; }
         #${TABLE_ID} tbody td.check-cell input[type="checkbox"] { margin: 0 !important; vertical-align: middle !important; }


        /* --- 簡化模式下的樣式 (針對 #table) --- */
        /* 隱藏 NO. (第一欄) - 但因為現在 NO. 固定了，所以不隱藏它，只隱藏性別 */
        /* #${TABLE_ID}.${SIMPLIFIED_CLASS} thead.normal th:nth-child(1), #${TABLE_ID}.${SIMPLIFIED_CLASS} tbody td:nth-child(1) { display: none !important; } */
        /* 隱藏 性別 (第四欄) */
        #${TABLE_ID}.${SIMPLIFIED_CLASS} thead.normal th:nth-child(4), #${TABLE_ID}.${SIMPLIFIED_CLASS} tbody td:nth-child(4) { display: none !important; }

        /* 簡化模式下，固定欄位的 left 值不變，因為它們是基於原始列計算的 */
        /* 簡化模式下 Checkbox 表頭 & 單元格 */
        #${TABLE_ID}.${SIMPLIFIED_CLASS} thead.normal th:has(input.check-all) { width: 2.8em !important; padding: 3px 2px !important; }
         #${TABLE_ID}.${SIMPLIFIED_CLASS} tbody td.check-cell { text-align: left !important; padding-left: 4px !important; width: auto !important; }
        #${TABLE_ID}.${SIMPLIFIED_CLASS} tbody td.check-cell .checkbox-label { display: inline-block !important; }
         #${TABLE_ID}.${SIMPLIFIED_CLASS} tbody td.check-cell input[type="checkbox"] { margin-left: 2px !important; }
    `);

    // --- JavaScript 功能函數 (與 v3.3 相同) ---
    // <editor-fold desc="v3.3 JS Functions">
    let originalHeaderTexts = {};
    function updateHeaderTexts(tableElement, simplify) {
        if (!tableElement) return;
        const headerCells = tableElement.querySelectorAll('thead.normal th');
        simplifiedHeaderLabels = []; // 清空/初始化每次更新

        headerCells.forEach((th, index) => {
             let textNode = null;
             let originalText = '';
             let inputEl = th.querySelector('input.check-all'); // 仍然需要查找，即使它被隱藏了

             // 獲取原始文字並儲存
             if (!originalHeaderTexts[index]) {
                 let textNodeFound = false;
                 if (inputEl) { // checkbox 表頭（即使 input 隱藏）
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
                     // 嘗試去除可能的 <div> 包裹的文字
                     let divText = th.querySelector('div:not(.icon)');
                     if (divText) {
                         originalText = divText.textContent.trim();
                     } else {
                         originalText = th.textContent.trim().split(/[\n\r]/)[0].trim();
                     }

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
                  // Fallback to find any text node if specific logic fails
                  if (!textNode) {
                      textNode = Array.from(th.childNodes).find(node => node.nodeType === Node.TEXT_NODE && node.textContent.trim());
                  }
             }

             // 更新文字內容並記錄簡化標籤
             let targetText = originalText;
             let simplifiedLabel = null; // 用於 tbody

             if (simplify) {
                 let simplified = false;
                 for (const key in headerSimplificationMap) {
                     // 使用 includes 或完全匹配，取決於原始標籤的複雜度
                     if (originalText.trim() === key || originalText.includes(key)) {
                         targetText = headerSimplificationMap[key];
                         simplifiedLabel = targetText; // 儲存簡化標籤
                         simplified = true;
                         break;
                     }
                 }
                 // 如果是 NO., 區別, 姓名, 性別，即使在簡化模式也顯示原始文字
                 if (index < 4) { // 假設前四列不需要簡化文字
                     targetText = originalText;
                     simplifiedLabel = null; // 這些列不需要 tbody 標籤
                 }

             } else {
                 targetText = originalText; // 還原
                 for (const key in headerSimplificationMap) {
                     if (originalText.includes(key)) {
                         simplifiedLabel = headerSimplificationMap[key]; break;
                     }
                 }
                  if (index < 4) { // 前四列不需要 tbody 標籤
                     simplifiedLabel = null;
                  }
             }
             simplifiedHeaderLabels[index] = simplifiedLabel;

             // 更新表頭的顯示文字
             if (textNode) {
                 textNode.textContent = ' ' + targetText + ' ';
             } else {
                 // 如果找不到 textNode，可能需要更複雜的 DOM 操作
                 // 檢查是否是前四列，它們可能有 <div>
                 if (index < 4) {
                    let targetDiv = th.querySelector('div:not(.icon)');
                    if (targetDiv) {
                        targetDiv.textContent = targetText; // 嘗試更新 div 內容
                    } else {
                        // 最終 fallback，直接設置 textContent，可能清除內部結構
                        // th.textContent = targetText;
                        console.warn(`無法為表頭 ${index} 找到合適的文字節點或容器來更新。`);
                    }
                 } else {
                      console.warn(`無法為表頭 ${index} 找到合適的文字節點來更新。`);
                 }
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

    // <editor-fold desc="v3.3 Main Logic">
    const table = document.getElementById(TABLE_ID);
    const menuDiv = document.getElementById(MENU_ID);

    if (table && menuDiv) {
         const normalThead = table.querySelector('thead.normal');
         if (!normalThead) { console.error('找不到 #table thead.normal'); return; }

         const toggleButton = document.createElement('button');
         toggleButton.id = TOGGLE_BUTTON_ID;
         menuDiv.appendChild(toggleButton);

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
         if (!table) console.warn(`無法找到 ID 為 "${TABLE_ID}" 的主表格。`);
         if (!menuDiv) console.warn(`無法找到 ID 為 "${MENU_ID}" 的菜單欄。按鈕將不會被添加。`);
    }
    // </editor-fold>

})();
