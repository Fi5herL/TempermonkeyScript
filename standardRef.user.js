// ==UserScript==
// @name         連結目標上下文預覽 (智能表格捕獲)
// @namespace    http://tampermonkey.net/
// @version      2025-05-23.3
// @description  滑鼠懸停於內部連結時，智能捕獲並完整顯示相關表格，保留HTML格式
// @author       You
// @match        https://www.ulsestandards.org/uls-standardsdocs/onlineviewer/*
// @include      https://www.ulsestandards.org/uls-standardsdocs/onlineviewer/*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=ulsestandards.org
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // --- 配置 ---
    const LINES_BEFORE = 0;
    const LINES_AFTER = 0;
    const DEBUG = true;
    const MAX_CONTENT_LENGTH_NON_TABLE = 3000; // 非表格內容的最大長度

    // --- 全局變量 ---
    const PREVIEW_DIV_ID = 'userscript-internal-link-preview-div';
    let previewDiv = null;
    let currentHoveredLink = null;
    let hidePreviewTimeout = null;

    function log(...args) {
        if (DEBUG) {
            let prefix = '[LinkPreviewSmartTable';
            if (window.self !== window.top) {
                prefix += ` (iframe: ${window.location.pathname.split('/').pop() || 'index'})`;
            }
            prefix += ']';
            console.log(prefix, ...args);
        }
    }

    log("腳本開始運行。當前頁面/iframe:", window.location.href);

    function createPreviewDiv() {
        // ... (與版本 2025-05-23.2 相同，為簡潔省略)
        if (document.getElementById(PREVIEW_DIV_ID)) {
            previewDiv = document.getElementById(PREVIEW_DIV_ID);
            return;
        }
        previewDiv = document.createElement('div');
        previewDiv.id = PREVIEW_DIV_ID;
        Object.assign(previewDiv.style, {
            position: 'fixed',
            border: '1px solid #888',
            backgroundColor: '#fff',
            color: '#333',
            padding: '12px',
            zIndex: '2147483647',
            maxWidth: '85vw', // 再次增大以適應複雜表格
            maxHeight: '80vh',
            overflow: 'auto',
            fontSize: '13px',
            fontFamily: 'Arial, sans-serif',
            lineHeight: '1.5',
            boxShadow: '0 5px 15px rgba(0,0,0,0.2)',
            display: 'none',
            textAlign: 'left'
        });
        document.body.appendChild(previewDiv);
        log("預覽框已創建並附加到 body of", window.location.href);

        previewDiv.addEventListener('mouseenter', () => {
            if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout);
            hidePreviewTimeout = null;
        });
        previewDiv.addEventListener('mouseleave', (event) => {
            if (event.relatedTarget !== currentHoveredLink) delayedHidePreview(100);
        });
    }

    function showPreview(htmlContent, event) {
        // ... (與版本 2025-05-23.2 相同，為簡潔省略)
        if (!previewDiv) createPreviewDiv();
        if (!previewDiv) {
            log("錯誤：預覽框創建失敗。");
            return;
        }
        if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout);
        hidePreviewTimeout = null;

        previewDiv.innerHTML = htmlContent;
        previewDiv.style.display = 'block';

        const pRect = previewDiv.getBoundingClientRect();
        let x = event.clientX + 20, y = event.clientY + 20;
        const vpWidth = window.innerWidth, vpHeight = window.innerHeight;

        if (x + pRect.width > vpWidth) x = Math.max(10, event.clientX - pRect.width - 20);
        if (y + pRect.height > vpHeight) y = Math.max(10, event.clientY - pRect.height - 20);
        if (x < 0) x = 10;
        if (y < 0) y = 10;

        previewDiv.style.left = x + 'px';
        previewDiv.style.top = y + 'px';
    }

    function hidePreview() {
        // ... (與版本 2025-05-23.2 相同，為簡潔省略)
        if (previewDiv && previewDiv.style.display !== 'none') {
            previewDiv.style.display = 'none';
        }
        currentHoveredLink = null;
        if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout);
        hidePreviewTimeout = null;
    }

    function delayedHidePreview(delay = 300) {
        // ... (與版本 2025-05-23.2 相同，為簡潔省略)
        if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout);
        hidePreviewTimeout = setTimeout(() => {
            if (previewDiv && previewDiv.matches(':hover')) return;
            hidePreview();
        }, delay);
    }

    /**
     * 嘗試找到與 targetElement 相關聯的完整表格。
     * 1. 如果 targetElement 本身是 table，返回它。
     * 2. 如果 targetElement 在 table 內部，返回最近的 table 祖先。
     * 3. 如果 targetElement 的直接子元素是 table (且只有一個 table)，返回該 table。
     * 4. 如果 targetElement 的下一個兄弟是 table，返回該 table。 (可選，看需求)
     * 5. 如果 targetElement 的前一個兄弟是 table 且 targetElement 是個簡單標題，返回該 table。 (可選)
     * @param {Element} targetElement - 連結指向的原始目標元素。
     * @returns {Element|null} - 找到的表格元素，或 null。
     */
    function findRelatedTable(targetElement) {
        if (!targetElement) return null;

        // 情況 1: targetElement 本身就是 table
        if (targetElement.tagName === 'TABLE') {
            log("findRelatedTable: targetElement is a TABLE.");
            return targetElement;
        }

        // 情況 2: targetElement 在 table 內部
        const closestTable = targetElement.closest('table');
        if (closestTable) {
            log("findRelatedTable: targetElement is inside a TABLE (closest).");
            return closestTable;
        }

        // 情況 3: targetElement 的直接子元素中包含 table
        // 我們特別關注 targetElement 是否是一個簡單的容器，其主要內容是表格
        const childTables = Array.from(targetElement.children).filter(child => child.tagName === 'TABLE');
        if (childTables.length === 1) {
            // 如果只有一個子表格，且父元素沒有太多其他複雜內容，可以考慮返回這個子表格
            // 這裡的判斷可以更複雜，例如檢查父元素是否除了這個表格外幾乎沒有其他可見內容
            log("findRelatedTable: targetElement contains one child TABLE.");
            return childTables[0];
        }
        if (childTables.length > 1) {
             log("findRelatedTable: targetElement contains multiple child TABLEs. Returning targetElement itself for now.");
             return targetElement; // 或者返回第一個表格 childTables[0]
        }


        // 情況 4: targetElement 的下一個兄弟元素是 table
        // 有時錨點可能在表格緊鄰的前一個元素上 (例如一個標題 <p id="foo"></p><table>...</table>)
        const nextSibling = targetElement.nextElementSibling;
        if (nextSibling && nextSibling.tagName === 'TABLE') {
            // 確保 targetElement 內容不多，否則可能不相關
            if (targetElement.textContent.trim().length < 100) { // 隨意設定一個閾值
                 log("findRelatedTable: nextElementSibling is a TABLE.");
                 return nextSibling;
            }
        }

        // 如果以上都沒找到特定關聯的表格，就返回原始 targetElement
        // getElementHtmlContent 會處理 targetElement 本身是否包含需要特殊處理的表格
        log("findRelatedTable: No clearly related single table found by structure, returning original targetElement.");
        return targetElement;
    }


    /**
     * 獲取元素的 HTML 內容。
     * 如果是表格，獲取 outerHTML 且不截斷。
     * 否則，獲取 innerHTML，清理並截斷。
     * @param {Element} element - 要處理的元素 (可能是原始 targetElement，也可能是 findRelatedTable 返回的表格)
     * @returns {string}
     */
    function getElementHtmlContent(element) {
        if (!element) return "";

        let html;
        let isConsideredTable = false; // 標記是否按表格方式處理 (即不截斷)

        if (element.tagName === 'TABLE') {
            isConsideredTable = true;
            html = element.outerHTML;
            log("getElementHtmlContent: Element is TABLE, using outerHTML.");
        } else {
            // 對於非 TABLE 元素，我們仍然檢查其內部是否包含表格
            // 但主要的截斷邏輯是針對非表格內容的
            html = element.innerHTML;
            log("getElementHtmlContent: Element is not TABLE, using innerHTML.");
        }

        // 清理 script, style, comments，這對所有內容都適用
        html = html.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "");
        html = html.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, "");
        html = html.replace(/<!--[\s\S]*?-->/g, "");

        // 僅當不是按表格方式處理時，才進行長度截斷
        if (!isConsideredTable && html.length > MAX_CONTENT_LENGTH_NON_TABLE) {
            log(`getElementHtmlContent: Content too long (${html.length}), truncating.`);
            let cutPoint = html.lastIndexOf('>', MAX_CONTENT_LENGTH_NON_TABLE);
            if (cutPoint === -1 || cutPoint < MAX_CONTENT_LENGTH_NON_TABLE / 2) {
                cutPoint = html.lastIndexOf(' ', MAX_CONTENT_LENGTH_NON_TABLE);
            }
            if (cutPoint === -1) cutPoint = MAX_CONTENT_LENGTH_NON_TABLE;
            html = html.substring(0, cutPoint) + '...';
        }
        return html.trim();
    }

    function getElementContext(initialTargetElement) {
        let contextParts = [];

        const styleSubHeading = "font-weight: bold; color: #555; margin-top: 8px; margin-bottom: 4px; display: block; border-bottom: 1px solid #eee; padding-bottom: 3px;";
        const styleTargetContainer = "border: 1px solid #ccc; padding: 10px; margin: 5px 0; display: block; overflow: auto; background-color: #f9f9f9;";

        // --- 智能查找要顯示的主要元素 (可能是表格) ---
        let displayElement = findRelatedTable(initialTargetElement);
        if (!displayElement) { // 如果 findRelatedTable 返回 null (理論上不應該，至少返回 initialTargetElement)
            log("Error: findRelatedTable returned null. Defaulting to initialTargetElement.");
            displayElement = initialTargetElement;
        }

        log(`getElementContext: Initial target is #${initialTargetElement.id}, displayElement is ${displayElement.tagName}#${displayElement.id || '(no id)'}`);

        // --- 獲取主要顯示元素的內容 ---
        const mainContentHtml = getElementHtmlContent(displayElement);

        contextParts.push(`<div style="${styleSubHeading}">--- 預覽內容 (ID: ${initialTargetElement.id || '無ID'}) ---</div>`);
        contextParts.push(`<div style="${styleTargetContainer}">${mainContentHtml || '(無可顯示內容)'}</div>`);

        // 前文和後文的邏輯 (LINES_BEFORE/AFTER 為 0 時不執行)
        // 如果需要顯示前文/後文，它們應該是相對於 initialTargetElement 的，而不是 displayElement
        let currentSibling;
        if (LINES_BEFORE > 0) {
            let beforeElementsHtml = [];
            currentSibling = initialTargetElement.previousElementSibling;
            for (let i = 0; i < LINES_BEFORE && currentSibling; i++) {
                // 確保前文/後文的元素不是我們已經顯示的 displayElement (如果 displayElement 是 initialTargetElement 的兄弟)
                if (currentSibling !== displayElement) {
                    const htmlSnippet = getElementHtmlContent(currentSibling); // 前後文也用 getElementHtmlContent
                    if (htmlSnippet.trim()) {
                        beforeElementsHtml.unshift(htmlSnippet);
                    }
                }
                currentSibling = currentSibling.previousElementSibling;
            }
            if (beforeElementsHtml.length > 0) {
                contextParts.unshift(`<div style="margin-top: 5px; padding-top: 5px; border-top: 1px dotted #ddd; overflow: auto;">${beforeElementsHtml.join('<hr style="border:none;border-top:1px dotted #ddd;margin:5px 0;">')}</div>`);
                contextParts.unshift(`<div style="${styleSubHeading}">--- 前文 (${beforeElementsHtml.length} 個兄弟元素) ---</div>`);
            }
        }

        if (LINES_AFTER > 0) {
            let afterElementsHtml = [];
            currentSibling = initialTargetElement.nextElementSibling;
            for (let i = 0; i < LINES_AFTER && currentSibling; i++) {
                if (currentSibling !== displayElement) {
                    const htmlSnippet = getElementHtmlContent(currentSibling);
                    if (htmlSnippet.trim()) {
                        afterElementsHtml.push(htmlSnippet);
                    }
                }
                currentSibling = currentSibling.nextElementSibling;
            }
            if (afterElementsHtml.length > 0) {
                contextParts.push(`<div style="${styleSubHeading}">--- 後文 (${afterElementsHtml.length} 個兄弟元素) ---</div>`);
                contextParts.push(`<div style="margin-bottom: 5px; padding-bottom: 5px; border-bottom: 1px dotted #ddd; overflow: auto;">${afterElementsHtml.join('<hr style="border:none;border-top:1px dotted #ddd;margin:5px 0;">')}</div>`);
            }
        }


        if (!mainContentHtml.trim() && contextParts.length <= 2) { // 檢查是否真的沒有任何有效內容
             return "連結目標及其上下文均無可顯示的內容。";
        }

        return contextParts.join('');
    }


    // --- 事件監聽器 ---
    document.addEventListener('mouseover', function(event) {
        // ... (與版本 2025-05-23.2 相同，為簡潔省略，僅修改日誌)
        const aTag = event.target.closest('a');

        if (aTag && aTag.href) {
            if (previewDiv && previewDiv.contains(event.target)) return;
            if (aTag === currentHoveredLink && previewDiv && previewDiv.style.display === 'block') {
                 if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout);
                 hidePreviewTimeout = null;
                 return;
            }

            try {
                const url = new URL(aTag.href, window.location.href);
                if (url.hash && url.origin === window.location.origin && url.pathname === window.location.pathname) {
                    const targetId = decodeURIComponent(url.hash.substring(1));
                    if (!targetId) return;

                    const targetElement = document.getElementById(targetId);
                    if (targetElement) {
                        log(`Mouseover on link to #${targetId}. Initial target element: ${targetElement.tagName}`);
                        currentHoveredLink = aTag;
                        const contextText = getElementContext(targetElement); // 傳入原始錨點元素
                        showPreview(contextText, event);
                    } else {
                         log("目標元素未找到: #" + targetId + " 在文檔:", window.location.href);
                    }
                }
            } catch (e) {
                log("處理連結 mouseover 錯誤:", e, "對於連結:", aTag.href);
            }
        }
    }, true);

    document.addEventListener('mouseout', function(event) {
        // ... (與版本 2025-05-23.2 相同，為簡潔省略)
        const aTag = event.target.closest('a');
        if (aTag && aTag === currentHoveredLink) {
            if (event.relatedTarget !== previewDiv && (!previewDiv || !previewDiv.contains(event.relatedTarget))) {
                delayedHidePreview();
            }
        }
    }, true);

    if (document.readyState === "complete" || document.readyState === "interactive") {
        createPreviewDiv();
    } else {
        document.addEventListener("DOMContentLoaded", createPreviewDiv);
    }

    log("腳本初始化完成。");

})();
