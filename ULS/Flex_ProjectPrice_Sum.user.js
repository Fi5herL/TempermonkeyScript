// ==UserScript==
// @name         網頁金額加總顯示器 (含案件資訊、日期過濾、手動刷新 - v1.2匯率邏輯)
// @namespace    http://tampermonkey.net/
// @version      1.6
// @description  在右下角顯示網頁中特定貨幣金額的加總(USD)，點擊可看詳情、換算公式、相關案件資訊，可依完成日期過濾，並可手動刷新數據。採用 v1.2 版匯率處理。
// @author       Fisher Li & AI Assistant
// @match        https://portal.ul.com/Dashboard*
// @grant        GM_xmlhttpRequest
// @connect      cdn.jsdelivr.net
// @connect      gh.fawazahmed0.com
// @run-at       document-idle
// @license      MIT
// ==/UserScript==

(function() {
    'use strict';

    // --- 常數與配置 ---
    const TARGET_CURRENCIES = ['USD', 'CNY', 'JPY', 'EUR', 'KRW', 'VND'];
    const API_URL = 'https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@latest/v1/currencies/usd.min.json';
    const SCRIPT_ID = 'currency-converter-aggregator-box-extreme-debug';
    const EXECUTION_DELAY = 3000;
    const LOCALE = 'en-US';
    const ZERO_DECIMAL_CURRENCIES = new Set(['USD', 'CNY', 'JPY', 'EUR', 'KRW', 'VND']);

    const PROJECT_INFO_COLUMNS = ["Project Number", "File No", "Project Name", "ECD", "Order Line Price", "Project Handler", "Completion Date", "Project Scope", "Status Note"];
    const GRID_CONTAINER_SELECTOR = '#projectDashboardGrid';
    const GRID_HEADER_SELECTOR = '.k-grid-header';
    const GRID_CONTENT_SELECTOR = '.k-grid-content.k-auto-scrollable';
    const GRID_ROW_SELECTOR = 'tr:not(.k-grouping-row):not(.k-filter-row):not(.k-grid-norecords)';

    let gridHeaderInfo = null;
    let allDetailedItemsGlobal = [];
    let currentFilterDate = null;
    let ratesDateGlobal = 'N/A';
    let isProcessing = false; // Flag to prevent multiple simultaneous processing

    // --- 輔助函數：解析 Title 字串 (與 v1.2 相同) ---
    function parseTitle(title) {
        if (!title) return null;
        const amountMatch = title.match(/\d{2,20}/);
        if (!amountMatch) return null;
        const amountString = amountMatch[0].replace(/,/g, '');
        const amount = parseFloat(amountString);
        if (isNaN(amount)) return null;
        let currency = null;
        const upperTitle = title.toUpperCase();
        for (const code of TARGET_CURRENCIES) {
             if (upperTitle.includes(code)) { currency = code; break; }
        }
        if (currency === null) return null;
        return { originalTitle: title, amount, currency };
    }

    // --- 輔助函數：格式化貨幣 (與 v1.2 相同) ---
    function formatCurrency(amount, currencyCode) {
        try {
            let minDigits = 2; let maxDigits = 2;
            const upperCode = currencyCode.toUpperCase();
            if (ZERO_DECIMAL_CURRENCIES.has(upperCode)) { minDigits = 0; maxDigits = 0; }
            return amount.toLocaleString(LOCALE, { style: 'currency', currency: upperCode, minimumFractionDigits: minDigits, maximumFractionDigits: maxDigits });
        } catch (e) {
            return `${currencyCode.toUpperCase()} ${amount.toFixed(ZERO_DECIMAL_CURRENCIES.has(currencyCode.toUpperCase()) ? 0 : 2)}`;
        }
    }

    // --- 異步函數：獲取匯率 (採用 v1.2 邏輯) ---
    async function getExchangeRates() {
        console.log("Tampermonkey: 正在獲取匯率...");
        return new Promise((resolve) => {
            GM_xmlhttpRequest({
                method: "GET",
                url: API_URL,
                timeout: 10000,
                onload: function(response) {
                    if (response.status >= 200 && response.status < 300) {
                        try {
                            const data = JSON.parse(response.responseText);
                            if (data && data.usd) {
                                const rates = { 'USD': 1 };
                                TARGET_CURRENCIES.forEach(code => {
                                    if (code === 'USD') return;
                                    const lowerCode = code.toLowerCase();
                                    if (data.usd.hasOwnProperty(lowerCode)) {
                                        const rateUsdToForeign = data.usd[lowerCode];
                                        if (rateUsdToForeign && typeof rateUsdToForeign === 'number' && rateUsdToForeign !== 0) {
                                            rates[code] = 1 / rateUsdToForeign;
                                        }
                                    }
                                });
                                ratesDateGlobal = data.date;
                                resolve({ rates, date: data.date });
                            } else {
                                ratesDateGlobal = 'N/A (API格式錯誤)';
                                resolve({ rates: getDefaultRates(), date: ratesDateGlobal });
                            }
                        } catch (e) {
                            ratesDateGlobal = 'N/A (解析失敗)';
                            resolve({ rates: getDefaultRates(), date: ratesDateGlobal });
                        }
                    } else {
                        ratesDateGlobal = 'N/A (請求失敗)';
                        resolve({ rates: getDefaultRates(), date: ratesDateGlobal });
                    }
                },
                onerror: function() {
                    ratesDateGlobal = 'N/A (網絡錯誤)';
                    resolve({ rates: getDefaultRates(), date: ratesDateGlobal });
                },
                ontimeout: function() {
                    ratesDateGlobal = 'N/A (超時)';
                    resolve({ rates: getDefaultRates(), date: ratesDateGlobal });
                }
            });
        });
    }

    // --- 輔助函數：提供預設/備用匯率 (採用 v1.2 邏輯: 1 Foreign = Y USD) ---
    function getDefaultRates() {
        console.warn("Tampermonkey: 使用預設/備用匯率。");
        return {
            'USD': 1, 'CNY': 0.14, 'JPY': 0.0067, 'EUR': 1.08, 'KRW': 0.00073, 'VND': 0.00004
        };
    }

    // --- 輔助函數：分析 Kendo UI Grid 的表頭 ---
    function analyzeGridHeaders() {
        const gridContainer = document.querySelector(GRID_CONTAINER_SELECTOR);
        if (!gridContainer) return null;
        const headerDiv = gridContainer.querySelector(GRID_HEADER_SELECTOR);
        if (!headerDiv) return null;
        const headerTr = headerDiv.querySelector('tr');
        if (!headerTr) return null;
        const thElements = Array.from(headerTr.querySelectorAll('th'));
        const columnIndexMap = new Map();
        const foundHeaderOrder = [];
        PROJECT_INFO_COLUMNS.forEach(name => {
            let thIndex = -1;
            for (let i = 0; i < thElements.length; i++) {
                const th = thElements[i];
                const textContent = th.textContent.trim().toLowerCase();
                const titleAttribute = th.getAttribute('title');
                if (textContent.includes(name.toLowerCase()) || (titleAttribute && titleAttribute.toLowerCase().includes(name.toLowerCase()))) {
                    thIndex = i; break;
                }
            }
            if (thIndex !== -1) {
                columnIndexMap.set(name, thIndex);
                foundHeaderOrder.push(name);
            }
        });
        return columnIndexMap.size > 0 ? { columnIndexMap, foundHeaderOrder } : null;
    }

    // --- 輔助函數：從特定 Grid 行提取數據 ---
    function extractProjectDataFromRow(rowElement, pGridHeaderInfo) {
        if (!rowElement || !pGridHeaderInfo || pGridHeaderInfo.columnIndexMap.size === 0) return null;
        const cellElements = rowElement.children;
        const rowData = {};
        pGridHeaderInfo.foundHeaderOrder.forEach(columnName => {
            const colIndex = pGridHeaderInfo.columnIndexMap.get(columnName);
            rowData[columnName] = (colIndex !== undefined && colIndex < cellElements.length) ? cellElements[colIndex].textContent.trim() : 'N/A';
        });
        return rowData;
    }

    // --- 輔助函數: 解析日期字串 (MM/DD/YYYY) ---
    function parseMMDDYYYY(dateString) {
        if (!dateString || typeof dateString !== 'string') return null;
        const parts = dateString.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (parts) {
            const date = new Date(parseInt(parts[3], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
            date.setHours(0, 0, 0, 0); return date;
        }
        const isoMatch = dateString.match(/(\d{4})-(\d{2})-(\d{2})/);
        if (isoMatch) {
            const date = new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10));
            date.setHours(0,0,0,0); return date;
        }
        return null;
    }

    // --- 主要異步執行函數 (採用 v1.2 匯率計算邏輯) ---
    async function processPage() {
        if (isProcessing) {
            console.log("Tampermonkey: 處理中，請稍候...");
            return;
        }
        isProcessing = true;
        console.log("Tampermonkey: 開始處理頁面 (v1.2 匯率邏輯)...");

        // Show loading indicator on refresh button if it exists
        const refreshBtn = document.getElementById('currency-refresh-button');
        if (refreshBtn) refreshBtn.textContent = '讀取中...';


        try {
            const { rates: exchangeRates } = await getExchangeRates();

            if (!gridHeaderInfo) { // Only analyze headers if not already done or explicitly reset
                gridHeaderInfo = analyzeGridHeaders();
            }

            const elements = document.querySelectorAll('*[title]');
            const parsedItems = [];

            for (const el of elements) {
                const currencyItem = parseTitle(el.getAttribute('title'));
                if (currencyItem) {
                    let projectInfo = null;
                    const gridRowElement = el.closest(`${GRID_CONTAINER_SELECTOR} ${GRID_CONTENT_SELECTOR} ${GRID_ROW_SELECTOR}`);
                    if (gridRowElement && gridHeaderInfo) {
                        projectInfo = extractProjectDataFromRow(gridRowElement, gridHeaderInfo);
                    }
                    parsedItems.push({ ...currencyItem, projectInfo, element: el });
                }
            }

            allDetailedItemsGlobal = [];

            if (parsedItems.length > 0) {
                parsedItems.forEach(item => {
                    const rate = exchangeRates[item.currency];
                    if (rate === undefined) return;
                    const amountInUSD = item.amount * rate;
                    allDetailedItemsGlobal.push({ ...item, rateUsed: rate, amountInUSD: amountInUSD });
                });
            }
        } catch (error) {
            console.error("Tampermonkey: processPage 發生錯誤:", error);
        } finally {
            isProcessing = false;
            if (refreshBtn) refreshBtn.textContent = '🔄'; // Reset button text
            createOrUpdateUI(); // Always update UI, even if there was an error (to clear loading state)
        }
    }

    // --- UI 創建與更新函數 ---
    function createOrUpdateUI() {
        let floatBox = document.getElementById(SCRIPT_ID);
        let totalDisplay, detailsDisplay, detailsList, infoDisplay, filterContainer, dateFilterInput, clearFilterButton, refreshButton;

        let itemsToDisplay = [...allDetailedItemsGlobal];
        if (currentFilterDate) {
            const filterDateObj = new Date(currentFilterDate);
            filterDateObj.setUTCHours(0,0,0,0);
            itemsToDisplay = allDetailedItemsGlobal.filter(item => {
                if (item.projectInfo && item.projectInfo["Completion Date"]) {
                    const itemDateObj = parseMMDDYYYY(item.projectInfo["Completion Date"]);
                    if (itemDateObj) return itemDateObj.getTime() >= filterDateObj.getTime();
                }
                return false;
            });
        }

        let currentTotalUSD = 0;
        itemsToDisplay.forEach(item => { currentTotalUSD += item.amountInUSD; });

        if (!floatBox) {
            floatBox = document.createElement('div');
            floatBox.id = SCRIPT_ID;
            // ... (floatBox styling)
            floatBox.style.position = 'fixed';
            floatBox.style.bottom = '15px';
            floatBox.style.right = '15px';
            floatBox.style.padding = '10px 15px';
            floatBox.style.backgroundColor = 'rgba(0, 0, 0, 0.75)';
            floatBox.style.color = 'white';
            floatBox.style.borderRadius = '8px';
            floatBox.style.zIndex = '9999';
            floatBox.style.fontFamily = 'Arial, sans-serif';
            floatBox.style.fontSize = '14px';
            floatBox.style.boxShadow = '0 2px 5px rgba(0,0,0,0.3)';
            floatBox.style.minWidth = '280px';
            floatBox.style.lineHeight = '1.4';
             // Add relative positioning for absolute positioned refresh button
            floatBox.style.position = 'fixed'; // Ensure it's fixed for the child to be absolute relative to it

            // --- Refresh Button ---
            refreshButton = document.createElement('button');
            refreshButton.id = 'currency-refresh-button';
            refreshButton.textContent = '🔄'; // Refresh icon
            refreshButton.title = '重新載入資料';
            refreshButton.style.position = 'absolute';
            refreshButton.style.top = '5px';
            refreshButton.style.right = '5px';
            refreshButton.style.background = 'none';
            refreshButton.style.border = 'none';
            refreshButton.style.color = 'white';
            refreshButton.style.fontSize = '16px';
            refreshButton.style.cursor = 'pointer';
            refreshButton.style.padding = '5px';
            refreshButton.style.lineHeight = '1';
            refreshButton.addEventListener('click', (e) => {
                e.stopPropagation(); // Prevent toggling details view
                if (!isProcessing) {
                    gridHeaderInfo = null; // Force re-analysis of headers on manual refresh if desired
                                           // or comment this out if headers are static.
                    processPage();
                }
            });
            floatBox.appendChild(refreshButton);
            // --- End Refresh Button ---

            totalDisplay = document.createElement('div');
            totalDisplay.id = 'currency-total-display';
            totalDisplay.style.fontWeight = 'bold';
            // Add some padding to total display if refresh button is present
            totalDisplay.style.paddingRight = '30px'; // Space for the refresh button
            floatBox.appendChild(totalDisplay);


            infoDisplay = document.createElement('div');
            infoDisplay.id = 'currency-info-display';
            // ... (infoDisplay styling)
            infoDisplay.style.fontSize = '10px';
            infoDisplay.style.color = '#bbb';
            infoDisplay.style.marginTop = '4px';
            floatBox.appendChild(infoDisplay);

            filterContainer = document.createElement('div');
            filterContainer.id = 'currency-filter-container';
            // ... (filterContainer styling)
            filterContainer.style.marginTop = '8px';
            filterContainer.style.paddingTop = '8px';
            filterContainer.style.borderTop = '1px solid rgba(255,255,255,0.3)';


            const filterLabel = document.createElement('label');
            filterLabel.textContent = '完成日期篩選 (之後): ';
            // ... (filterLabel styling)
            filterLabel.style.fontSize = '11px';
            filterLabel.style.marginRight = '5px';
            filterContainer.appendChild(filterLabel);

            dateFilterInput = document.createElement('input');
            dateFilterInput.type = 'date';
            dateFilterInput.id = 'completion-date-filter';
            // ... (dateFilterInput styling)
            dateFilterInput.style.fontSize = '11px';
            dateFilterInput.style.padding = '2px';
            dateFilterInput.style.backgroundColor = '#333';
            dateFilterInput.style.color = 'white';
            dateFilterInput.style.border = '1px solid #555';
            dateFilterInput.addEventListener('change', (e) => {
                currentFilterDate = e.target.value;
                createOrUpdateUI();
            });
            filterContainer.appendChild(dateFilterInput);

            clearFilterButton = document.createElement('button');
            clearFilterButton.textContent = '清除';
            // ... (clearFilterButton styling)
            clearFilterButton.style.fontSize = '11px';
            clearFilterButton.style.marginLeft = '5px';
            clearFilterButton.style.padding = '2px 5px';
            clearFilterButton.style.cursor = 'pointer';
            clearFilterButton.style.backgroundColor = '#555';
            clearFilterButton.style.color = 'white';
            clearFilterButton.style.border = '1px solid #777';
            clearFilterButton.addEventListener('click', () => {
                currentFilterDate = null;
                dateFilterInput.value = '';
                createOrUpdateUI();
            });
            filterContainer.appendChild(clearFilterButton);
            floatBox.appendChild(filterContainer);

            detailsDisplay = document.createElement('div');
            detailsDisplay.id = 'currency-details-display';
            // ... (detailsDisplay styling)
            detailsDisplay.style.display = 'none';
            detailsDisplay.style.marginTop = '10px';
            detailsDisplay.style.paddingTop = '10px';
            detailsDisplay.style.borderTop = '1px solid rgba(255, 255, 255, 0.5)';
            detailsDisplay.style.maxHeight = '350px';
            detailsDisplay.style.overflowY = 'auto';
            detailsDisplay.style.fontSize = '12px';

            detailsList = document.createElement('ul');
            detailsList.style.listStyle = 'none';
            detailsList.style.paddingLeft = '0';
            detailsList.style.margin = '0';
            detailsDisplay.appendChild(detailsList);
            floatBox.appendChild(detailsDisplay);

            floatBox.addEventListener('click', (event) => {
                if (event.target.closest('#currency-filter-container') || event.target.closest('#currency-refresh-button')) {
                    event.stopPropagation(); return;
                }
                if (event.target === floatBox || event.target === totalDisplay || event.target === infoDisplay) {
                    const isHidden = detailsDisplay.style.display === 'none';
                    detailsDisplay.style.display = isHidden ? 'block' : 'none';
                    floatBox.style.backgroundColor = isHidden ? 'rgba(0, 0, 0, 0.9)' : 'rgba(0, 0, 0, 0.75)';
                }
            });
            document.body.appendChild(floatBox);
        } else {
            totalDisplay = floatBox.querySelector('#currency-total-display');
            infoDisplay = floatBox.querySelector('#currency-info-display');
            detailsDisplay = floatBox.querySelector('#currency-details-display');
            detailsList = detailsDisplay.querySelector('ul');
            dateFilterInput = floatBox.querySelector('#completion-date-filter');
            refreshButton = floatBox.querySelector('#currency-refresh-button'); // Get existing refresh button
        }

        if (itemsToDisplay.length === 0 && allDetailedItemsGlobal.length > 0 && currentFilterDate) {
            totalDisplay.textContent = `總金額 ≈ ${formatCurrency(0, 'USD')} (無符合篩選條件的案件)`;
        } else if (itemsToDisplay.length === 0 && allDetailedItemsGlobal.length === 0 && !isProcessing) { // Don't overwrite if processing
             totalDisplay.textContent = `總金額 ≈ ${formatCurrency(0, 'USD')} (無案件資訊)`;
        } else if (!isProcessing) { // Don't overwrite if processing
            totalDisplay.textContent = `總金額 ≈ ${formatCurrency(currentTotalUSD, 'USD')}`;
        }

        infoDisplay.textContent = `基於 ${ratesDateGlobal} 匯率 (Locale: ${LOCALE})`;
        if (dateFilterInput) dateFilterInput.value = currentFilterDate || '';
        if (refreshButton && !isProcessing) refreshButton.textContent = '🔄'; // Ensure button text is reset if not processing

        detailsList.innerHTML = '';

        if (itemsToDisplay.length === 0 && !isProcessing) { // Don't show "no items" if still processing
            const noItemsLi = document.createElement('li');
            noItemsLi.textContent = currentFilterDate ? "沒有符合篩選條件的案件。" : "沒有可顯示的案件資訊。";
            // ... (noItemsLi styling)
            noItemsLi.style.padding = "10px";
            noItemsLi.style.textAlign = "center";
            noItemsLi.style.color = "#aaa";
            detailsList.appendChild(noItemsLi);
        } else if (itemsToDisplay.length > 0) { // Only populate if there are items and not processing
            const groupedItems = itemsToDisplay.reduce((acc, item) => {
                if (!acc[item.currency]) acc[item.currency] = [];
                acc[item.currency].push(item);
                return acc;
            }, {});

            TARGET_CURRENCIES.forEach(currency => {
                if (groupedItems[currency] && groupedItems[currency].length > 0) {
                    const currencyHeader = document.createElement('li');
                    currencyHeader.textContent = `--- ${currency} ---`;
                    // ... (currencyHeader styling)
                    currencyHeader.style.fontWeight = 'bold';
                    currencyHeader.style.marginTop = '10px';
                    currencyHeader.style.color = '#eee';
                    detailsList.appendChild(currencyHeader);

                    groupedItems[currency].forEach(item => {
                        const listItem = document.createElement('li');
                        // ... (listItem styling)
                        listItem.style.marginBottom = '10px';
                        listItem.style.paddingBottom = '8px';
                        listItem.style.borderBottom = '1px dotted rgba(255, 255, 255, 0.2)';

                        const originalSpan = document.createElement('span');
                        // ... (originalSpan styling and content)
                        originalSpan.style.display = 'block';
                        originalSpan.style.fontWeight = 'bold';
                        originalSpan.style.color = 'yellow';
                        originalSpan.textContent = `${item.amount} ${item.currency}`;
                        listItem.appendChild(originalSpan);

                        const conversionSpan = document.createElement('span');
                        // ... (conversionSpan styling and content)
                        conversionSpan.style.display = 'block';
                        conversionSpan.style.fontSize = '11px';
                        conversionSpan.style.color = '#ccc';
                        conversionSpan.style.paddingLeft = '10px';
                        const rateDisplay = item.rateUsed.toFixed(5);
                        const convertedFormatted = formatCurrency(item.amountInUSD, 'USD');
                        conversionSpan.innerHTML = `(1 ${item.currency} ≈ ${rateDisplay} USD) ≈ ${convertedFormatted}`;
                        listItem.appendChild(conversionSpan);


                        if (item.projectInfo && gridHeaderInfo && gridHeaderInfo.foundHeaderOrder.length > 0) {
                            const projectInfoDiv = document.createElement('div');
                            // ... (projectInfoDiv styling and content)
                            projectInfoDiv.style.marginTop = '5px';
                            projectInfoDiv.style.paddingLeft = '10px';
                            projectInfoDiv.style.fontSize = '11px';
                            projectInfoDiv.style.color = '#ddd';
                            gridHeaderInfo.foundHeaderOrder.forEach(colName => {
                                if (item.projectInfo[colName]) {
                                    const infoLine = document.createElement('div');
                                    infoLine.innerHTML = `<strong>${colName}:</strong> ${item.projectInfo[colName]}`;
                                    projectInfoDiv.appendChild(infoLine);
                                }
                            });
                            listItem.appendChild(projectInfoDiv);
                        }
                        detailsList.appendChild(listItem);
                    });
                }
            });
        }
    }

    // --- 延遲執行主程序 & DOM觀察 ---
    let executionScheduled = false;
    let mainExecuted = false;
    let observer = null;

    function scheduleMainExecution() {
        if (executionScheduled || mainExecuted) return;
        executionScheduled = true;
        setTimeout(() => {
            if (mainExecuted) return;
            mainExecuted = true;
            if (observer) observer.disconnect();
            processPage();
        }, EXECUTION_DELAY);
    }

    if (document.readyState === "complete" || document.readyState === "interactive") {
        scheduleMainExecution();
    } else {
        window.addEventListener('DOMContentLoaded', scheduleMainExecution);
    }

    const gridEl = document.querySelector(GRID_CONTAINER_SELECTOR);
    if (gridEl) {
        scheduleMainExecution();
    } else {
        observer = new MutationObserver((mutationsList, obs) => {
            if (document.querySelector(GRID_CONTAINER_SELECTOR)) {
                scheduleMainExecution();
                obs.disconnect(); observer = null;
            }
        });
        observer.observe(document.body, { childList: true, subtree: true });
        setTimeout(() => {
            if (observer) observer.disconnect();
            scheduleMainExecution();
        }, EXECUTION_DELAY + 2000);
    }
})();
