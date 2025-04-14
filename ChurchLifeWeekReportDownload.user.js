// ==UserScript==
// @name         人數統計下載按鈕 (菜單+日期檔名+多頁合併 V1.1.2)
// @namespace    http://tampermonkey.net/
// @version      1.1.2
// @description  在主菜單添加下載按鈕(含日期檔名),點擊後合併所有分頁數據下載CSV
// @author       You
// @match        https://www.chlife-stat.org/
// @match        https://www.chlife-stat.org/index.php
// @grant        GM_addStyle
// ==/UserScript==

(function() {
    'use strict';

    // --- 配置 ---
    const targetTableSelector = '#roll-call-panel table#table';
    const targetMenuSelector = '#menu';
    const dateDisplaySelector = '#show_year_week';
    const paginationContainerSelector = '#pagination.jPaginate';

    // --- 函數：將數據轉換為 CSV 字符串 ---
    function convertToCSV(data) {
        if (!data || !data.headers || !data.rows) return null;
        const headers = data.headers;
        const rows = data.rows;
        const escapeCSV = (field) => {
            if (field === null || field === undefined) return '';
            const stringField = String(field);
            if (stringField.includes(',') || stringField.includes('"') || stringField.includes('\n')) {
                return `"${stringField.replace(/"/g, '""')}"`;
            }
            return stringField;
        };
        const headerString = headers.map(escapeCSV).join(',');
        const rowStrings = rows.map(row => headers.map((_, index) => escapeCSV(row[index])).join(','));
        const BOM = '\uFEFF';
        return BOM + headerString + '\n' + rowStrings.join('\n');
    }

    // --- 函數：觸發 CSV 下載 ---
    function downloadCSV(csvString, filename) {
        if (!csvString) return;
        const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.setAttribute("href", url);
        link.setAttribute("download", filename);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        console.log(`CSV 檔案 "${filename}" 已觸發下載。`);
    }

    // --- 函數：獲取並格式化日期範圍用於檔名 ---
    function getDateRangeForFilename() {
        const dateDisplayElement = document.querySelector(dateDisplaySelector);
        if (!dateDisplayElement) { console.error(`找不到日期元素 (${dateDisplaySelector})`); return null; }
        let fullDateString = dateDisplayElement.textContent.trim().split('【')[0].trim();
        const dateRegex = /(\d{4})年(\d{1,2})月(\d{1,2})日/g;
        const matches = [...fullDateString.matchAll(dateRegex)];
        if (matches.length < 2) { console.error("無法提取日期:", fullDateString); return null; }
        const formatDate = (match) => `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
        const formattedStartDate = formatDate(matches[0]);
        const formattedEndDate = formatDate(matches[1]);
        return (formattedStartDate && formattedEndDate) ? `${formattedStartDate}_${formattedEndDate}` : null;
    }

    // --- 函數：從當前表格抓取數據行 (修改：跳過第一行數據) ---
    function scrapeCurrentPageRows(tableElement) {
        if (!tableElement) return [];

        // 1. 選擇 tbody 中所有非空行的 tr 元素
        const allRowElements = tableElement.querySelectorAll('tbody tr:not(.empty-row)');

        // 2. 將 NodeList 轉換為陣列，並 **跳過第一個元素 (第一行)**
        //    使用 slice(1) 從索引 1 開始取元素，達到跳過第一行的目的
        const rowsToProcess = Array.from(allRowElements).slice(1);

        // 3. 對剩餘的行進行處理和映射
        const dataRows = rowsToProcess.map(row => {
            const cells = row.querySelectorAll('td');
            // 如果找不到儲存格或只有一個 '無資料' 儲存格，則跳過此行 (理論上不應發生，因為已跳過第一行)
            if (cells.length === 0 || (cells.length === 1 && cells[0].textContent.trim() === '無資料')) return null;

            // 提取每一格的資料
            const rowData = Array.from(cells).map(td => {
                 if (td.classList.contains('check-cell')) {
                     // 處理勾選框儲存格
                     const input = td.querySelector('input');
                     return (input && input.checked) ? '1' : '0';
                 } else {
                     // 處理一般文字儲存格 (並去除前後空白)
                     return td.textContent.trim();
                 }
             });

            // 過濾邏輯：如果整行所有儲存格都是空的，也濾掉
            if (rowData.every(cell => cell === '')) return null;

            // 保留此行資料
            return rowData;
        }).filter(row => row !== null); // 移除所有被標記為 null 的行

        return dataRows;
    }

     // --- **新增** 函數：從分頁控件獲取總頁數 ---
    function getTotalPages(paginationContainer) {
        if (!paginationContainer) return 1; // 如果找不到控件，假設只有1頁

        // 選擇所有代表頁碼的元素 (<a> 和 <span>)
        const pageElements = paginationContainer.querySelectorAll('ul.jPag-pages li a, ul.jPag-pages li span.jPag-current');
        let maxPage = 0;

        pageElements.forEach(el => {
            const pageNum = parseInt(el.textContent.trim(), 10);
            // 確保是有效的數字，並更新最大值
            if (!isNaN(pageNum) && pageNum > maxPage) {
                maxPage = pageNum;
            }
        });

        // 如果找不到任何頁碼元素，預設為1頁，否則返回找到的最大頁碼
        return maxPage > 0 ? maxPage : 1;
    }


    // --- 函數：等待指定頁碼的數據載入完成 ---
    async function waitForPageLoad(expectedPageNum, paginationContainer) {
        console.log(`等待第 ${expectedPageNum} 頁載入...`);
        return new Promise((resolve, reject) => {
            const maxWaitTime = 45000;
            const checkInterval = 500;
            let elapsedTime = 0;

            const intervalId = setInterval(() => {
                elapsedTime += checkInterval;
                const activePageElement = paginationContainer?.querySelector('span.jPag-current');
                let currentPageNum = -1;
                if (activePageElement) {
                    currentPageNum = parseInt(activePageElement.textContent.trim());
                } else {
                     console.warn("waitForPageLoad: 未找到 span.jPag-current");
                }

                if (currentPageNum === expectedPageNum) {
                    console.log(`第 ${expectedPageNum} 頁似乎已載入。`);
                    clearInterval(intervalId);
                    setTimeout(resolve, 500); // 短暫延遲確保渲染
                } else if (elapsedTime >= maxWaitTime) {
                    clearInterval(intervalId);
                    console.error(`等待第 ${expectedPageNum} 頁載入超時。`);
                    reject(new Error(`Timeout waiting for page ${expectedPageNum}`));
                } else {
                    // console.log(`仍在等待第 ${expectedPageNum} 頁，當前: ${currentPageNum}`); // 可選的詳細日誌
                }
            }, checkInterval);
        });
    }

    // --- 函數：抓取所有分頁數據並觸發下載 ---
    async function scrapeAndDownloadAllPages() {
        const button = document.getElementById('menu-download-csv-button');
        if (button) {
            button.disabled = true;
            button.textContent = '初始化...';
        }

        let allRows = [];
        let headers = [];
        let currentPage = 1;
        let totalPages = 1; // 先預設為 1

        try {
            // --- **步驟 1: 獲取總頁數** ---
            const paginationContainerInitial = document.querySelector(paginationContainerSelector);
            if (paginationContainerInitial) {
                totalPages = getTotalPages(paginationContainerInitial);
                console.log(`檢測到總頁數：${totalPages}`);
                if (button) button.textContent = `處理中(1/${totalPages})...`;
            } else {
                console.log("未找到分頁控件，假設只有一頁。");
                if (button) button.textContent = '處理中(1/1)...';
            }

            // --- **步驟 2: 循環抓取每一頁** ---
            while (currentPage <= totalPages) {
                console.log(`正在處理第 ${currentPage} 頁...`);
                 if (button) button.textContent = `處理中(${currentPage}/${totalPages})...`;

                const tableElement = document.querySelector(targetTableSelector);
                 if (!tableElement) throw new Error(`找不到表格在第 ${currentPage} 頁`);

                // 只在第一頁獲取表頭
                if (currentPage === 1) {
                    const headerRow = tableElement.querySelector('thead tr');
                    if (!headerRow) throw new Error("找不到表頭");
                    headers = Array.from(headerRow.querySelectorAll('th')).map(th => th.textContent.trim());
                    if (headers.length === 0) throw new Error("表頭為空");
                }

                // 抓取當前頁數據
                const currentRows = scrapeCurrentPageRows(tableElement);
                allRows.push(...currentRows);
                console.log(`第 ${currentPage} 頁抓取到 ${currentRows.length} 行數據，總計 ${allRows.length} 行。`);

                // --- **步驟 3: 檢查是否需要翻頁** ---
                if (currentPage >= totalPages) {
                    console.log("已到達檢測到的最後一頁。");
                    break; // 跳出循環
                }

                // --- 如果不是最後一頁，則翻頁 ---
                const paginationContainer = document.querySelector(paginationContainerSelector); // 重新獲取以防萬一
                 if (!paginationContainer) {
                     console.warn("在循環中找不到分頁控件，無法翻頁。");
                     break;
                 }
                const nextPageButton = paginationContainer.querySelector('a.jPag-next');

                if (nextPageButton) {
                    const nextPageNum = currentPage + 1;
                    console.log(`準備點擊 '下一頁' 前往第 ${nextPageNum} 頁...`);
                    nextPageButton.click();
                    // 等待新頁面載入完成 (用 waitForPageLoad 檢查頁碼是否變為 nextPageNum)
                    await waitForPageLoad(nextPageNum, paginationContainer);
                    currentPage = nextPageNum; // **重要：在確認載入後才更新當前頁碼**
                } else {
                    // 理論上不應該執行到這裡，因為上面已經判斷過 currentPage 和 totalPages
                    console.warn("找不到 '下一頁' 按鈕，即使尚未到達預計的最後一頁。停止翻頁。");
                    break;
                }
            } // End of while loop

            // --- 所有頁面處理完成 ---
            console.log(`所有頁面處理完成，共抓取 ${allRows.length} 行數據。`);
            if (button) button.textContent = '產生CSV...';

            if (headers.length === 0 || allRows.length === 0) {
                 // 如果總頁數大於1但未抓到數據，可能抓取邏輯有問題
                 if (totalPages > 1) {
                    alert("未抓取到任何有效數據，請檢查腳本或網頁。");
                 } else {
                    // 如果只有1頁且無數據
                    alert("表格中沒有數據，無法生成 CSV。");
                 }
                throw new Error("未抓取到有效數據");
            }

            const data = { headers: headers, rows: allRows };
            const csvData = convertToCSV(data);

            if (csvData) {
                const dateRangePart = getDateRangeForFilename();
                let dynamicFilename = `人數統計(共${allRows.length}筆)_${dateRangePart || new Date().toISOString().slice(0, 10) + '_日期獲取失敗'}.csv`;
                if (!dateRangePart) console.warn("無法從頁面獲取日期範圍，使用備用檔名。");
                downloadCSV(csvData, dynamicFilename);
            } else {
                 alert("錯誤：無法將數據轉換為 CSV 格式。");
            }

        } catch (error) {
            console.error("處理多頁數據時出錯:", error);
            alert(`處理過程中發生錯誤：\n${error.message}`);
        } finally {
            if (button) {
                button.disabled = false;
                button.textContent = '下載統計';
            }
        }
    }

    // --- 函數：創建並添加到主菜單欄 ---
    function createMenuButton() {
        GM_addStyle(`
            #menu-download-csv-wrapper { float: left; margin-left: 2px; opacity: 1; }
            #menu-download-csv-button {
                display: block; padding: 8px 12px; background-color: #4CAF50;
                color: white; border: 1px solid #388E3C; border-radius: 5px;
                cursor: pointer; font-size: 14px; text-align: center;
                white-space: nowrap; transition: background-color 0.3s ease;
            }
            #menu-download-csv-button:hover { background-color: #388E3C; }
            #menu-download-csv-button:disabled { background-color: #cccccc; border-color: #aaaaaa; cursor: not-allowed; }
        `);
        const button = document.createElement('button');
        button.id = 'menu-download-csv-button';
        button.textContent = '下載統計';
        button.type = 'button';
        button.addEventListener('click', scrapeAndDownloadAllPages);

        const buttonWrapper = document.createElement('div');
        buttonWrapper.id = 'menu-download-csv-wrapper';
        buttonWrapper.className = 'menu_box';
        buttonWrapper.style.opacity = 1;
        buttonWrapper.appendChild(button);

        const menuContainer = document.querySelector(targetMenuSelector);
        if (menuContainer) {
            menuContainer.appendChild(buttonWrapper);
            console.log("多頁合併下載按鈕(V1.1.2)已添加到主菜單。");
        } else {
            console.error(`錯誤：找不到目標菜單欄 (${targetMenuSelector})，無法添加按鈕。`);
        }
    }

    // --- 主邏輯：等待頁面載入後創建按鈕 ---
    function runInitialization() {
        createMenuButton();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', runInitialization);
    } else {
        runInitialization();
    }

})();
