// ==UserScript==
// @name         Parse and Link UL Standard from ProductId
// @namespace    http://tampermonkey.net/
// @version      4.0
// @description  Finds links with a productId (e.g., "UL60730-1"), parses it to get the catalog and standard number, and adds a corresponding "UL" button.
// @author       Your Name
// @match        https://iq.ulprospector.com/en/profile*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // --- 配置 ---
    const BASE_URL = "https://www.ulsestandards.org/uls-standardsdocs/StandardDocuments.aspx";
    const BUTTON_STYLES = 'margin-left: 5px; padding: 2px 5px; background-color: #007bff; color: white; text-decoration: none; border-radius: 3px; font-size: 0.8em; cursor: pointer; display: inline-block; vertical-align: middle;';

    /**
     * 掃描頁面並為符合條件的連結添加按鈕
     */
    function addUlButtons() {
        // 1. 尋找所有 href 中包含 "productId=" 的連結
        const links = document.querySelectorAll('a[href*="productId="]');

        for (const link of links) {
            // 防止重複添加按鈕
            if (link.nextElementSibling && link.nextElementSibling.classList.contains('ul-product-button')) {
                continue;
            }

            try {
                const url = new URL(link.href);
                const productId = url.searchParams.get('productId');

                if (!productId) {
                    continue; // 如果沒有 productId，跳過
                }

                // 2. *** 核心解析邏輯 ***
                //    從 productId 字串中分離出 Catalog 和 Standard Number
                let catalog = null;
                let standardNumber = null;

                // 為了不區分大小寫，我們先將 productId 轉換為大寫進行判斷
                const upperProductId = productId.toUpperCase();

                // 必須先檢查 "ULC" 和 "ULE"，因為它們也以 "UL" 開頭
                if (upperProductId.startsWith('ULC')) {
                    catalog = 'ULC';
                    standardNumber = productId.substring(3); // 移除 "ULC" (3個字符)
                } else if (upperProductId.startsWith('ULE')) {
                    catalog = 'ULE';
                    standardNumber = productId.substring(3); // 移除 "ULE" (3個字符)
                } else if (upperProductId.startsWith('UL')) {
                    catalog = 'UL';
                    standardNumber = productId.substring(2); // 移除 "UL" (2個字符)
                } else {
                    // 如果 productId 不是以任何已知前綴開頭，則跳過此連結
                    continue;
                }

                // 3. 如果成功解析，則創建按鈕
                if (catalog && standardNumber) {
                    const newHref = `${BASE_URL}?Catalog=${catalog}&Standard=${encodeURIComponent(standardNumber)}`;

                    const button = document.createElement('a');
                    button.href = newHref;
                    button.textContent = 'UL';
                    button.className = 'ul-product-button';
                    button.style.cssText = BUTTON_STYLES;
                    button.target = '_blank';
                    button.rel = 'noopener noreferrer';
                    button.title = `在新分頁中打開 ${catalog} ${standardNumber} 標準`;

                    link.parentNode.insertBefore(button, link.nextSibling);
                }
            } catch (e) {
                console.error("腳本處理連結時出錯:", link.href, e);
            }
        }
    }

    // --- 腳本初始化與監聽 ---

    let debounceTimer;
    const debouncedScan = () => {
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(addUlButtons, 300);
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', debouncedScan);
    } else {
        debouncedScan();
    }

    const observer = new MutationObserver(debouncedScan);
    observer.observe(document.body, {
        childList: true,
        subtree: true
    });

})();
