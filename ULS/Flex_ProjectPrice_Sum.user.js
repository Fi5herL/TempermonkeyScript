// ==UserScript==
// @name         網頁金額加總顯示器
// @namespace    http://tampermonkey.net/
// @version      1.2
// @description  在右下角顯示網頁中特定貨幣金額的加總(USD)，點擊可看詳情與換算公式。
// @author       Fisher Li
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
    const SCRIPT_ID = 'currency-converter-aggregator-box-extreme-debug'; // UI 元素的 ID
    const EXECUTION_DELAY = 2000;
    const LOCALE = 'en-US';
    const ZERO_DECIMAL_CURRENCIES = new Set(['USD', 'CNY', 'JPY', 'EUR', 'KRW', 'VND']);

    // --- 輔助函數：解析 Title 字串 (增加類型和值檢查) ---
    function parseTitle(title) {
        if (!title) return null;
        const amountMatch = title.match(/\d{2,20}/);
        if (!amountMatch) return null;
        const amountString = amountMatch[0].replace(/,/g, '');
        const amount = parseFloat(amountString);
        if (isNaN(amount)) return null;

        // *** 極端調試點 1: 檢查解析後的類型和值 ***
        console.log(`[!!! DEBUG PARSE !!!] Title: "${title}" | Parsed String: "${amountString}" | Parsed Amount: ${amount} (Type: ${typeof amount})`);

        let currency = null;
        const upperTitle = title.toUpperCase();
        for (const code of TARGET_CURRENCIES) {
             if (upperTitle.includes(code)) { currency = code; break; }
        }
        if (currency === null) return null;

        return { originalTitle: title, amount, currency };
    }

    // --- 輔助函數：格式化貨幣 (用於總金額和換算後金額) ---
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

    // --- 異步函數：獲取匯率 (完整版) ---
    async function getExchangeRates() {
        console.log("Tampermonkey (v3.4): 正在獲取匯率...");
        return new Promise((resolve) => {
            GM_xmlhttpRequest({
                method: "GET",
                url: API_URL,
                timeout: 10000, // 10秒超時
                onload: function(response) {
                    if (response.status >= 200 && response.status < 300) {
                        try {
                            const data = JSON.parse(response.responseText);
                            if (data && data.usd) {
                                const rates = { 'USD': 1 }; // 始終包含 USD
                                // 遍歷目標貨幣列表 (排除 USD)
                                TARGET_CURRENCIES.forEach(code => {
                                    if (code === 'USD') return; // 跳過 USD 本身
                                    const lowerCode = code.toLowerCase();
                                    if (data.usd.hasOwnProperty(lowerCode)) {
                                        const rateUsdToForeign = data.usd[lowerCode];
                                        if (rateUsdToForeign && typeof rateUsdToForeign === 'number' && rateUsdToForeign !== 0) {
                                            rates[code] = 1 / rateUsdToForeign; // 存儲 1 Foreign = X USD
                                        } else {
                                             console.warn(`Tampermonkey (v3.4): API 返回的 ${code} 匯率無效: ${rateUsdToForeign}`);
                                        }
                                    } else {
                                        console.warn(`Tampermonkey (v3.4): API 未返回 ${code} 的匯率。`);
                                    }
                                });
                                console.log(`Tampermonkey (v3.4): 成功處理目標貨幣匯率 (基於 ${data.date}):`, rates);
                                resolve({ rates, date: data.date });
                            } else {
                                console.warn("Tampermonkey (v3.4): API 回應格式不符，使用預設匯率。 Response:", data);
                                resolve({ rates: getDefaultRates(), date: 'N/A (API格式錯誤)' });
                            }
                        } catch (e) {
                            console.error("Tampermonkey (v3.4): 解析匯率 API 回應失敗:", e, "Response Text:", response.responseText);
                            resolve({ rates: getDefaultRates(), date: 'N/A (解析失敗)' });
                        }
                    } else {
                        console.error(`Tampermonkey (v3.4): 獲取匯率失敗，狀態碼: ${response.status}`, response);
                        resolve({ rates: getDefaultRates(), date: 'N/A (請求失敗)' });
                    }
                },
                onerror: function(response) {
                    console.error("Tampermonkey (v3.4): 獲取匯率請求錯誤:", response);
                    resolve({ rates: getDefaultRates(), date: 'N/A (網絡錯誤)' });
                },
                ontimeout: function() {
                    console.error("Tampermonkey (v3.4): 獲取匯率請求超時。");
                    resolve({ rates: getDefaultRates(), date: 'N/A (超時)' });
                }
            });
        });
    }

    // --- 輔助函數：提供預設/備用匯率 (完整版) ---
    function getDefaultRates() {
        console.warn("Tampermonkey (v3.4): 使用預設/備用匯率。");
        return {
            'USD': 1,
            'CNY': 0.14, // 示例匯率，請按需更新
            'JPY': 0.0067, // 示例匯率
            'EUR': 1.08, // 示例匯率
            'KRW': 0.00073, // 示例匯率
            'VND': 0.00004 // 示例匯率
        };
    }

    // --- 主要異步執行函數 (完整版) ---
    async function processPage() {
        console.log("Tampermonkey (v3.4): 開始處理頁面...");
        const { rates: exchangeRates, date: ratesDate } = await getExchangeRates();

        console.log("Tampermonkey (v3.4): 正在查找和解析 title 屬性...");
        const elements = document.querySelectorAll('*');
        const parsedData = Array.from(elements)
            .filter(el => el.hasAttribute('title'))
            .map(el => parseTitle(el.getAttribute('title'))) // parseTitle 會打印 amount
            .filter(item => item !== null); // parseTitle 內部已確保是目標貨幣或返回 null

        console.log(`Tampermonkey (v3.4): 解析完成，找到 ${parsedData.length} 個目標貨幣金額。`);

        let totalUSD = 0;
        const detailedItems = [];

        if (parsedData.length > 0) {
            parsedData.forEach(item => {
                const rate = exchangeRates[item.currency];
                if (rate === undefined) {
                    // 理論上不應發生，但以防萬一
                    console.warn(`Tampermonkey (v3.4): 找不到貨幣 ${item.currency} 的匯率，跳過:`, item.originalTitle);
                    return;
                }
                const amountInUSD = item.amount * rate;
                totalUSD += amountInUSD;

                // *** 極端調試點 2: 檢查存儲到 detailedItems 的 amount ***
                console.log(`[!!! DEBUG PROCESS !!!] Storing item: Amount=${item.amount}, Currency=${item.currency}, Rate=${rate}, AmountInUSD=${amountInUSD}`);
                detailedItems.push({ ...item, rateUsed: rate, amountInUSD: amountInUSD });
            });
            console.log(`Tampermonkey (v3.4): 計算完成，總金額 (原始值): ${totalUSD} USD`);
        } else {
            console.log("Tampermonkey (v3.4): 未找到任何目標貨幣的金額進行計算。");
             const existingBox = document.getElementById(SCRIPT_ID);
             if (existingBox) existingBox.remove();
            return; // 結束執行
        }
        createOrUpdateUI(totalUSD, detailedItems, ratesDate);
    }

    // --- UI 創建與更新函數 (繞過 formatCurrency 顯示原始 amount) ---
    function createOrUpdateUI(totalUSD, detailedItems, ratesDate) {
        let floatBox = document.getElementById(SCRIPT_ID);
        let totalDisplay, detailsDisplay, detailsList, infoDisplay;

        if (!floatBox) {
            // --- 創建新浮動框 ---
            floatBox = document.createElement('div');
            floatBox.id = SCRIPT_ID;
            // --- 樣式設定 ---
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
            floatBox.style.minWidth = '200px';
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

            detailsDisplay = document.createElement('div');
            detailsDisplay.id = 'currency-details-display';
            detailsDisplay.style.display = 'none';
            detailsDisplay.style.marginTop = '10px';
            detailsDisplay.style.paddingTop = '10px';
            detailsDisplay.style.borderTop = '1px solid rgba(255, 255, 255, 0.5)';
            detailsDisplay.style.maxHeight = '300px';
            detailsDisplay.style.overflowY = 'auto';
            detailsDisplay.style.fontSize = '12px';

            detailsList = document.createElement('ul');
            detailsList.style.listStyle = 'none';
            detailsList.style.paddingLeft = '0';
            detailsList.style.margin = '0';
            detailsDisplay.appendChild(detailsList);
            floatBox.appendChild(detailsDisplay);

            // --- 點擊事件 ---
            floatBox.addEventListener('click', (event) => {
                if (event.target === floatBox || event.target === totalDisplay || event.target === infoDisplay) {
                    const isHidden = detailsDisplay.style.display === 'none';
                    detailsDisplay.style.display = isHidden ? 'block' : 'none';
                    floatBox.style.backgroundColor = isHidden ? 'rgba(0, 0, 0, 0.9)' : 'rgba(0, 0, 0, 0.75)';
                }
            });

            document.body.appendChild(floatBox);
            console.log("Tampermonkey (v3.4): UI 浮動框已創建。");

        } else {
            // --- 更新現有浮動框 ---
            totalDisplay = floatBox.querySelector('#currency-total-display');
            infoDisplay = floatBox.querySelector('#currency-info-display');
            detailsDisplay = floatBox.querySelector('#currency-details-display');
            detailsList = detailsDisplay.querySelector('ul');
            detailsList.innerHTML = ''; // 清空舊列表
            detailsDisplay.style.display = 'none'; // 確保詳情預設收起
            floatBox.style.backgroundColor = 'rgba(0, 0, 0, 0.75)'; // 重置背景色
            console.log("Tampermonkey (v3.4): UI 浮動框已找到，準備更新。");
        }

        // --- 更新顯示內容 ---
        totalDisplay.textContent = `總金額 ≈ ${formatCurrency(totalUSD, 'USD')}`; // 總金額仍然格式化
        infoDisplay.textContent = `基於 ${ratesDate} 匯率 (Locale: ${LOCALE})`;

        // 按目標貨幣分組顯示
        const groupedItems = detailedItems.reduce((acc, item) => {
            if (!acc[item.currency]) acc[item.currency] = [];
            acc[item.currency].push(item);
            return acc;
        }, {});

        // 保持 TARGET_CURRENCIES 的順序來顯示分組
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

                    // 1. *** 極端調試：直接顯示 item.amount，不格式化 ***
                    const originalSpan = document.createElement('span');
                    originalSpan.style.display = 'block';
                    originalSpan.style.fontWeight = 'bold'; // 加粗以示區別
                    originalSpan.style.color = 'yellow'; // 顯眼顏色
                    // *** 極端調試點 3: 檢查 UI 更新前 item.amount 的值 ***
                    console.log(`[!!! DEBUG UI !!!] Displaying original item.amount: ${item.amount} (Type: ${typeof item.amount}) for Currency: ${item.currency}`);
                    originalSpan.textContent = `${item.amount} ${item.currency}`; // 直接顯示數字和貨幣代碼
                    listItem.appendChild(originalSpan);

                    // 2. 顯示換算公式和結果 (仍然使用 formatCurrency)
                    const conversionSpan = document.createElement('span');
                    conversionSpan.style.display = 'block';
                    conversionSpan.style.fontSize = '11px';
                    conversionSpan.style.color = '#ccc';
                    conversionSpan.style.paddingLeft = '10px';
                    const rateDisplay = item.rateUsed.toFixed(2);
                    const convertedFormatted = formatCurrency(item.amountInUSD, 'USD');
                    conversionSpan.innerHTML = `(1 ${item.currency} ≈ ${rateDisplay} USD) ≈ ${convertedFormatted}`;
                    listItem.appendChild(conversionSpan);

                    // 3. 顯示原始 Title
//                     const titleSpan = document.createElement('span');
//                     titleSpan.style.display = 'block';
//                     titleSpan.style.fontSize = '10px';
//                     titleSpan.style.color = '#aaa';
//                     titleSpan.style.marginTop = '3px';
//                     titleSpan.style.paddingLeft = '10px';
//                     titleSpan.style.fontStyle = 'italic';
//                     titleSpan.style.wordBreak = 'break-all';
//                     titleSpan.textContent = `原始: "${item.originalTitle}"`;
//                     listItem.appendChild(titleSpan);

                    detailsList.appendChild(listItem);
                });
            }
        });
         console.log("Tampermonkey (v3.4): UI 內容已更新 (使用極端調試)。");
    }

    // --- 延遲執行主程序 ---
    console.log(`Tampermonkey (v3.4 Extreme Debug): 腳本已加載，將在 ${EXECUTION_DELAY} 毫秒後執行主程序。`);
    setTimeout(processPage, EXECUTION_DELAY);

})();
