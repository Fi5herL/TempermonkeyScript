// ==UserScript==
// @name         網頁金額加總顯示器 (含案件資訊與日期過濾 - v1.2匯率邏輯)
// @namespace    http://tampermonkey.net/
// @version      1.5
// @description  在右下角顯示網頁中特定貨幣金額的加總(USD)，點擊可看詳情、換算公式、相關案件資訊，並可依完成日期過濾。採用 v1.2 版匯率處理。
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
    const GRID_CONTENT_SELECTOR = '.k-grid-content.k-auto-scrollable'; // Or just .k-grid-content
    const GRID_ROW_SELECTOR = 'tr:not(.k-grouping-row):not(.k-filter-row):not(.k-grid-norecords)';

    let gridHeaderInfo = null;
    let allDetailedItemsGlobal = [];
    let currentFilterDate = null;
    let ratesDateGlobal = 'N/A';

    // --- 輔助函數：解析 Title 字串 (與 v1.2 相同) ---
    function parseTitle(title) {
        if (!title) return null;
        const amountMatch = title.match(/\d{2,20}/); // v1.2 uses this
        if (!amountMatch) return null;
        const amountString = amountMatch[0].replace(/,/g, '');
        const amount = parseFloat(amountString);
        if (isNaN(amount)) return null;

        // console.log(`[DEBUG PARSE] Title: "${title}" | Parsed String: "${amountString}" | Parsed Amount: ${amount} (Type: ${typeof amount})`);

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
                                const rates = { 'USD': 1 }; // 始終包含 USD (1 USD = 1 USD)
                                // 遍歷目標貨幣列表 (排除 USD)
                                TARGET_CURRENCIES.forEach(code => {
                                    if (code === 'USD') return; // 跳過 USD 本身
                                    const lowerCode = code.toLowerCase();
                                    if (data.usd.hasOwnProperty(lowerCode)) {
                                        const rateUsdToForeign = data.usd[lowerCode]; // This is 1 USD = X Foreign
                                        if (rateUsdToForeign && typeof rateUsdToForeign === 'number' && rateUsdToForeign !== 0) {
                                            rates[code] = 1 / rateUsdToForeign; // 存儲 1 Foreign = Y USD
                                        } else {
                                             console.warn(`Tampermonkey: API 返回的 ${code} 匯率無效: ${rateUsdToForeign}`);
                                        }
                                    } else {
                                        console.warn(`Tampermonkey: API 未返回 ${code} 的匯率。`);
                                    }
                                });
                                console.log(`Tampermonkey: 成功處理目標貨幣匯率 (基於 ${data.date}):`, rates);
                                ratesDateGlobal = data.date;
                                resolve({ rates, date: data.date });
                            } else {
                                console.warn("Tampermonkey: API 回應格式不符，使用預設匯率。", data);
                                ratesDateGlobal = 'N/A (API格式錯誤)';
                                resolve({ rates: getDefaultRates(), date: ratesDateGlobal });
                            }
                        } catch (e) {
                            console.error("Tampermonkey: 解析匯率 API 回應失敗:", e);
                            ratesDateGlobal = 'N/A (解析失敗)';
                            resolve({ rates: getDefaultRates(), date: ratesDateGlobal });
                        }
                    } else {
                        console.error(`Tampermonkey: 獲取匯率失敗，狀態碼: ${response.status}`);
                        ratesDateGlobal = 'N/A (請求失敗)';
                        resolve({ rates: getDefaultRates(), date: ratesDateGlobal });
                    }
                },
                onerror: function() {
                    console.error("Tampermonkey: 獲取匯率請求錯誤:");
                    ratesDateGlobal = 'N/A (網絡錯誤)';
                    resolve({ rates: getDefaultRates(), date: ratesDateGlobal });
                },
                ontimeout: function() {
                    console.error("Tampermonkey: 獲取匯率請求超時。");
                    ratesDateGlobal = 'N/A (超時)';
                    resolve({ rates: getDefaultRates(), date: ratesDateGlobal });
                }
            });
        });
    }

    // --- 輔助函數：提供預設/備用匯率 (採用 v1.2 邏輯: 1 Foreign = Y USD) ---
    function getDefaultRates() {
        console.warn("Tampermonkey: 使用預設/備用匯率。");
        // These rates should represent: 1 ForeignCurrency = X USD
        return {
            'USD': 1,
            'CNY': 0.14,    // 1 CNY = 0.14 USD (approx)
            'JPY': 0.0067,  // 1 JPY = 0.0067 USD (approx)
            'EUR': 1.08,    // 1 EUR = 1.08 USD (approx)
            'KRW': 0.00073, // 1 KRW = 0.00073 USD (approx)
            'VND': 0.00004  // 1 VND = 0.00004 USD (approx)
        };
    }

    // --- 輔助函數：分析 Kendo UI Grid 的表頭 ---
    function analyzeGridHeaders() {
        // ... (no changes from v1.4)
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
        // ... (no changes from v1.4)
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
        // ... (no changes from v1.4)
        if (!dateString || typeof dateString !== 'string') return null;
        const parts = dateString.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
        if (parts) {
            const date = new Date(parseInt(parts[3], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
            date.setHours(0, 0, 0, 0);
            return date;
        }
        const isoMatch = dateString.match(/(\d{4})-(\d{2})-(\d{2})/);
        if (isoMatch) {
            const date = new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10));
            date.setHours(0,0,0,0);
            return date;
        }
        console.warn("Tampermonkey: 無法解析日期字串:", dateString);
        return null;
    }

    // --- 主要異步執行函數 (採用 v1.2 匯率計算邏輯) ---
    async function processPage() {
        console.log("Tampermonkey: 開始處理頁面 (v1.2 匯率邏輯)...");
        const { rates: exchangeRates } = await getExchangeRates(); // ratesDate is global

        if (!gridHeaderInfo) {
            gridHeaderInfo = analyzeGridHeaders();
        }

        const elements = document.querySelectorAll('*[title]');
        const parsedItems = [];

        for (const el of elements) {
            const currencyItem = parseTitle(el.getAttribute('title')); // Uses v1.2 parseTitle
            if (currencyItem) {
                let projectInfo = null;
                const gridRowElement = el.closest(`${GRID_CONTAINER_SELECTOR} ${GRID_CONTENT_SELECTOR} ${GRID_ROW_SELECTOR}`);
                if (gridRowElement && gridHeaderInfo) {
                    projectInfo = extractProjectDataFromRow(gridRowElement, gridHeaderInfo);
                }
                parsedItems.push({ ...currencyItem, projectInfo, element: el });
            }
        }

        console.log(`Tampermonkey: 解析完成，找到 ${parsedItems.length} 個目標貨幣金額。`);
        allDetailedItemsGlobal = [];

        if (parsedItems.length > 0) {
            parsedItems.forEach(item => {
                const rate = exchangeRates[item.currency]; // rate is now (1 Foreign = Y USD)
                if (rate === undefined) {
                    console.warn(`Tampermonkey: 找不到貨幣 ${item.currency} 的匯率，跳過:`, item.originalTitle);
                    return;
                }
                // Calculation now matches v1.2: item.amount (in Foreign) * rate (Foreign to USD)
                const amountInUSD = item.amount * rate;

                // console.log(`[DEBUG PROCESS] Storing item: Amount=${item.amount}, Currency=${item.currency}, Rate=${rate}, AmountInUSD=${amountInUSD}`);
                allDetailedItemsGlobal.push({ ...item, rateUsed: rate, amountInUSD: amountInUSD });
            });
        } else {
            console.log("Tampermonkey: 未找到任何目標貨幣的金額進行計算。");
        }
        createOrUpdateUI();
    }

    // --- UI 創建與更新函數 (與 v1.4 相同，但顯示的 rateUsed 和 amountInUSD 會基於新邏輯) ---
    function createOrUpdateUI() {
        let floatBox = document.getElementById(SCRIPT_ID);
        let totalDisplay, detailsDisplay, detailsList, infoDisplay, filterContainer, dateFilterInput, clearFilterButton;

        let itemsToDisplay = [...allDetailedItemsGlobal];
        if (currentFilterDate) {
            const filterDateObj = new Date(currentFilterDate);
            filterDateObj.setUTCHours(0,0,0,0);

            itemsToDisplay = allDetailedItemsGlobal.filter(item => {
                if (item.projectInfo && item.projectInfo["Completion Date"]) {
                    const itemDateObj = parseMMDDYYYY(item.projectInfo["Completion Date"]);
                    if (itemDateObj) {
                        return itemDateObj.getTime() >= filterDateObj.getTime();
                    }
                }
                return false;
            });
        }

        let currentTotalUSD = 0;
        itemsToDisplay.forEach(item => {
            currentTotalUSD += item.amountInUSD; // This amountInUSD is now calculated as per v1.2
        });


        if (!floatBox) {
            floatBox = document.createElement('div');
            floatBox.id = SCRIPT_ID;
            floatBox.style.position = 'fixed';
            floatBox.style.bottom = '15px';
            floatBox.style.right = '15px';
            floatBox.style.padding = '10px 15px';
            floatBox.style.backgroundColor = 'rgba(0, 0, 0, 0.75)';
            floatBox.style.color = 'white';
            floatBox.style.borderRadius = '8px';
            floatBox.style.zIndex = '9999';
            floatBox.style.cursor = 'pointer';
            floatBox.style.fontFamily = 'Arial, sans-serif';
            floatBox.style.fontSize = '14px';
            floatBox.style.boxShadow = '0 2px 5px rgba(0,0,0,0.3)';
            floatBox.style.minWidth = '280px';
            floatBox.style.lineHeight = '1.4';

            totalDisplay = document.createElement('div');
            totalDisplay.id = 'currency-total-display';
            totalDisplay.style.fontWeight = 'bold';
            floatBox.appendChild(totalDisplay);

            infoDisplay = document.createElement('div');
            infoDisplay.id = 'currency-info-display';
            infoDisplay.style.fontSize = '10px';
            infoDisplay.style.color = '#bbb';
            infoDisplay.style.marginTop = '4px';
            floatBox.appendChild(infoDisplay);

            filterContainer = document.createElement('div');
            filterContainer.id = 'currency-filter-container';
            filterContainer.style.marginTop = '8px';
            filterContainer.style.paddingTop = '8px';
            filterContainer.style.borderTop = '1px solid rgba(255,255,255,0.3)';

            const filterLabel = document.createElement('label');
            filterLabel.textContent = '完成日期篩選 (之後): ';
            filterLabel.style.fontSize = '11px';
            filterLabel.style.marginRight = '5px';
            filterContainer.appendChild(filterLabel);

            dateFilterInput = document.createElement('input');
            dateFilterInput.type = 'date';
            dateFilterInput.id = 'completion-date-filter';
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
                if (event.target.closest('#currency-filter-container')) {
                    event.stopPropagation();
                    return;
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
        }

        if (itemsToDisplay.length === 0 && allDetailedItemsGlobal.length > 0 && currentFilterDate) {
            totalDisplay.textContent = `總金額 ≈ ${formatCurrency(0, 'USD')} (無符合篩選條件的案件)`;
        } else if (itemsToDisplay.length === 0 && allDetailedItemsGlobal.length === 0) {
             totalDisplay.textContent = `總金額 ≈ ${formatCurrency(0, 'USD')} (無案件資訊)`;
        } else {
            totalDisplay.textContent = `總金額 ≈ ${formatCurrency(currentTotalUSD, 'USD')}`;
        }
        infoDisplay.textContent = `基於 ${ratesDateGlobal} 匯率 (Locale: ${LOCALE})`;
        if (dateFilterInput) dateFilterInput.value = currentFilterDate || '';

        detailsList.innerHTML = '';

        if (itemsToDisplay.length === 0) {
            const noItemsLi = document.createElement('li');
            noItemsLi.textContent = currentFilterDate ? "沒有符合篩選條件的案件。" : "沒有可顯示的案件資訊。";
            noItemsLi.style.padding = "10px";
            noItemsLi.style.textAlign = "center";
            noItemsLi.style.color = "#aaa";
            detailsList.appendChild(noItemsLi);
        } else {
            const groupedItems = itemsToDisplay.reduce((acc, item) => {
                if (!acc[item.currency]) acc[item.currency] = [];
                acc[item.currency].push(item);
                return acc;
            }, {});

            TARGET_CURRENCIES.forEach(currency => {
                if (groupedItems[currency] && groupedItems[currency].length > 0) {
                    const currencyHeader = document.createElement('li');
                    currencyHeader.textContent = `--- ${currency} ---`;
                    currencyHeader.style.fontWeight = 'bold';
                    currencyHeader.style.marginTop = '10px';
                    currencyHeader.style.color = '#eee';
                    detailsList.appendChild(currencyHeader);

                    groupedItems[currency].forEach(item => {
                        const listItem = document.createElement('li');
                        listItem.style.marginBottom = '10px';
                        listItem.style.paddingBottom = '8px';
                        listItem.style.borderBottom = '1px dotted rgba(255, 255, 255, 0.2)';

                        const originalSpan = document.createElement('span');
                        originalSpan.style.display = 'block';
                        originalSpan.style.fontWeight = 'bold';
                        originalSpan.style.color = 'yellow';
                        // Display original item.amount as per v1.2
                        originalSpan.textContent = `${item.amount} ${item.currency}`;
                        listItem.appendChild(originalSpan);

                        const conversionSpan = document.createElement('span');
                        conversionSpan.style.display = 'block';
                        conversionSpan.style.fontSize = '11px';
                        conversionSpan.style.color = '#ccc';
                        conversionSpan.style.paddingLeft = '10px';
                        // rateUsed is now (1 Foreign = Y USD)
                        // item.amountInUSD is item.amount * rateUsed
                        const rateDisplay = item.rateUsed.toFixed(5); // Potentially show more precision for inverse rate
                        const convertedFormatted = formatCurrency(item.amountInUSD, 'USD');
                        conversionSpan.innerHTML = `(1 ${item.currency} ≈ ${rateDisplay} USD) ≈ ${convertedFormatted}`;
                        listItem.appendChild(conversionSpan);

                        if (item.projectInfo && gridHeaderInfo && gridHeaderInfo.foundHeaderOrder.length > 0) {
                            const projectInfoDiv = document.createElement('div');
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
        // console.log("Tampermonkey: UI 內容已更新 (v1.2 匯率邏輯)。");
    }

    // --- 延遲執行主程序 & DOM觀察 (與 v1.4 相同) ---
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
                obs.disconnect();
                observer = null;
            }
        });
        observer.observe(document.body, { childList: true, subtree: true });
        setTimeout(() => {
            if (observer) observer.disconnect();
            scheduleMainExecution();
        }, EXECUTION_DELAY + 2000);
    }
})();
