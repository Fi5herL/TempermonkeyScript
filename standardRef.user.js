// ==UserScript==
// @name         連結目標前後文預覽 (保留格式版)
// @namespace    http://tampermonkey.net/
// @version      2025-05-23.1
// @description  滑鼠懸停於內部連結時，顯示目標元素的前後文，並嘗試保留部分HTML格式
// @author       You
// @match        https://www.ulsestandards.org/uls-standardsdocs/onlineviewer/*
// @include      https://www.ulsestandards.org/uls-standardsdocs/onlineviewer/*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=ulsestandards.org
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // --- 配置 ---
    const LINES_BEFORE = 4; // 顯示目標元素之前的兄弟元素數量
    const LINES_AFTER = 4;  // 顯示目標元素之後的兄弟元素數量
    const DEBUG = true;     // 設置為 true 以在控制台輸出調試信息
    const MAX_CONTENT_LENGTH_PER_ELEMENT = 500; // 每個元素顯示內容的最大長度，防止過長

    // --- 全局變量 ---
    const PREVIEW_DIV_ID = 'userscript-internal-link-preview-div';
    let previewDiv = null;
    let currentHoveredLink = null;
    let hidePreviewTimeout = null;

    function log(...args) {
        if (DEBUG) {
            let prefix = '[LinkPreviewFormat';
            if (window.self !== window.top) {
                prefix += ` (iframe: ${window.location.pathname.split('/').pop() || 'index'})`;
            }
            prefix += ']';
            console.log(prefix, ...args);
        }
    }

    log("腳本開始運行。當前頁面/iframe:", window.location.href);

    function createPreviewDiv() {
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
            maxWidth: '600px',
            maxHeight: '450px',
            overflowY: 'auto',
            fontSize: '13px',
            fontFamily: 'Arial, sans-serif',
            lineHeight: '1.5',
            boxShadow: '0 5px 15px rgba(0,0,0,0.2)',
            // whiteSpace: 'pre-wrap', // 移除這個，讓HTML自己處理空白和換行
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
        if (!previewDiv) createPreviewDiv();
        if (!previewDiv) {
            log("錯誤：預覽框創建失敗。");
            return;
        }
        if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout);
        hidePreviewTimeout = null;

        previewDiv.innerHTML = htmlContent; // 直接使用innerHTML
        previewDiv.style.display = 'block';

        let x = event.clientX + 20, y = event.clientY + 20;
        const vpWidth = window.innerWidth, vpHeight = window.innerHeight;
        const pRect = previewDiv.getBoundingClientRect();

        if (x + pRect.width > vpWidth) x = Math.max(10, event.clientX - pRect.width - 20);
        if (y + pRect.height > vpHeight) y = Math.max(10, event.clientY - pRect.height - 20);

        previewDiv.style.left = x + 'px';
        previewDiv.style.top = y + 'px';
    }

    function hidePreview() {
        if (previewDiv && previewDiv.style.display !== 'none') {
            previewDiv.style.display = 'none';
        }
        currentHoveredLink = null;
        if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout);
        hidePreviewTimeout = null;
    }

    function delayedHidePreview(delay = 300) {
        if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout);
        hidePreviewTimeout = setTimeout(() => {
            if (previewDiv && previewDiv.matches(':hover')) return;
            hidePreview();
        }, delay);
    }

    /**
     * 獲取元素的 HTML 內容，並進行清理和截斷
     * @param {Element} element
     * @returns {string}
     */
    function getElementHtmlSnippet(element) {
        if (!element) return "";
        let html = element.innerHTML;
        // 簡單的清理：移除 script 和 style 標籤，防止意外執行或樣式衝突
        html = html.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "");
        html = html.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, "");

        // 移除 HTML 註釋
        html = html.replace(/<!--[\s\S]*?-->/g, "");

        // 截斷（可選，如果內容過長）
        if (html.length > MAX_CONTENT_LENGTH_PER_ELEMENT) {
            // 嘗試在標籤結束處截斷，以避免產生無效HTML
            let cutPoint = html.lastIndexOf('>', MAX_CONTENT_LENGTH_PER_ELEMENT);
            if (cutPoint === -1 || cutPoint < MAX_CONTENT_LENGTH_PER_ELEMENT / 2) { // 如果找不到合適的 > 或太靠前
                cutPoint = html.lastIndexOf(' ', MAX_CONTENT_LENGTH_PER_ELEMENT); // 嘗試在空格處截斷
            }
            if (cutPoint === -1) cutPoint = MAX_CONTENT_LENGTH_PER_ELEMENT;

            html = html.substring(0, cutPoint) + '...';
        }
        return html.trim();
    }

    /**
     * 獲取目標元素的上下文
     * @param {Element} targetElement - 目標錨點對應的元素
     * @returns {string} - HTML 格式的上下文內容
     */
    function getElementContext(targetElement) {
        let contextParts = [];
        let currentElement;

        const styleSubHeading = "font-weight: bold; color: #555; margin-top: 8px; margin-bottom: 4px; display: block; border-bottom: 1px solid #eee; padding-bottom: 3px;";
        const styleTargetHighlightContainer = "background-color: #e6f7ff; border: 1px dashed #91d5ff; padding: 5px; margin: 2px 0; display: block;"; // 改為 display: block

        // --- 前文 ---
        let beforeElementsHtml = [];
        currentElement = targetElement.previousElementSibling;
        for (let i = 0; i < LINES_BEFORE && currentElement; i++) {
            const htmlSnippet = getElementHtmlSnippet(currentElement);
            if (htmlSnippet) {
                beforeElementsHtml.unshift(htmlSnippet); // unshift 添加到數組開頭，保持順序
            }
            currentElement = currentElement.previousElementSibling;
        }
        if (beforeElementsHtml.length > 0) {
            contextParts.push(`<div style="${styleSubHeading}">--- 前文 (${beforeElementsHtml.length} 個兄弟元素) ---</div>`);
            // 每個兄弟元素用 div 包裹，以便更好地控制間距和塊級顯示
            beforeElementsHtml.forEach(html => {
                contextParts.push(`<div style="margin-bottom: 5px; padding-bottom: 5px; border-bottom: 1px dotted #ddd;">${html}</div>`);
            });
        }

        // --- 目標元素 ---
        const targetHtml = getElementHtmlSnippet(targetElement);
        contextParts.push(`<div style="${styleSubHeading}">--- 目標元素 (ID: ${targetElement.id || '無ID'}) ---</div>`);
        contextParts.push(`<div style="${styleTargetHighlightContainer}">${targetHtml || '(目標元素無內容)'}</div>`);


        // --- 後文 ---
        let afterElementsHtml = [];
        currentElement = targetElement.nextElementSibling;
        for (let i = 0; i < LINES_AFTER && currentElement; i++) {
            const htmlSnippet = getElementHtmlSnippet(currentElement);
            if (htmlSnippet) {
                afterElementsHtml.push(htmlSnippet);
            }
            currentElement = currentElement.nextElementSibling;
        }
        if (afterElementsHtml.length > 0) {
            contextParts.push(`<div style="${styleSubHeading}">--- 後文 (${afterElementsHtml.length} 個兄弟元素) ---</div>`);
            afterElementsHtml.forEach(html => {
                contextParts.push(`<div style="margin-top: 5px; padding-top: 5px; border-top: 1px dotted #ddd;">${html}</div>`);
            });
        }

        if (beforeElementsHtml.length === 0 && afterElementsHtml.length === 0 && !targetHtml) {
            return "目標元素及其相鄰兄弟元素均無可顯示的內容。";
        }

        return contextParts.join(''); // 因為我們已經用 div 包裹了，所以不再需要 <br> 來分隔主要部分
    }

    // `escapeHtml` 函數在這裡不再需要，因為我們希望保留HTML。
    // 如果需要對特定輸入進行轉義，可以保留它，但不要用在 getElementHtmlSnippet 的輸出上。

    // --- 事件監聽器 ---
    document.addEventListener('mouseover', function(event) {
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
                        currentHoveredLink = aTag;
                        const contextText = getElementContext(targetElement);
                        showPreview(contextText, event);
                    } else {
                        // log("目標元素未找到: #" + targetId + " 在文檔:", window.location.href);
                    }
                }
            } catch (e) {
                log("處理連結 mouseover 錯誤:", e, "對於連結:", aTag.href);
            }
        }
    }, true);

    document.addEventListener('mouseout', function(event) {
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
