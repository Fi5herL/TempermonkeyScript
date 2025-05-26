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
    const DELAY_BEFORE_SHOWING_PREVIEW = 300; // 滑鼠懸停後延遲顯示預覽的時間 (毫秒)

    // --- 全局變量 ---
    const PREVIEW_DIV_ID = 'userscript-internal-link-preview-div';
    let previewDiv = null;
    let currentHoveredLink = null;
    let hidePreviewTimeout = null;
    let showPreviewTimeout = null; // 新增：用於延遲顯示預覽的計時器

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
            maxWidth: '85vw',
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
        if (!previewDiv) createPreviewDiv();
        if (!previewDiv) {
            log("錯誤：預覽框創建失敗。");
            return;
        }
        if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout); // 清除可能存在的隱藏計時器
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
        if (previewDiv && previewDiv.style.display !== 'none') {
            previewDiv.style.display = 'none';
        }
        currentHoveredLink = null; // 重置當前懸停連結
        if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout);
        hidePreviewTimeout = null;
        if (showPreviewTimeout) { // 如果有正在等待顯示的計時器，也一併清除
            clearTimeout(showPreviewTimeout);
            showPreviewTimeout = null;
        }
    }

    function delayedHidePreview(delay = 300) {
        if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout);
        hidePreviewTimeout = setTimeout(() => {
            if (previewDiv && previewDiv.matches(':hover')) return; // 如果滑鼠在預覽框內，則不隱藏
            hidePreview();
        }, delay);
    }

    function findRelatedTable(targetElement) {
        if (!targetElement) return null;
        if (targetElement.tagName === 'TABLE') {
            log("findRelatedTable: targetElement is a TABLE.");
            return targetElement;
        }
        const closestTable = targetElement.closest('table');
        if (closestTable) {
            log("findRelatedTable: targetElement is inside a TABLE (closest).");
            return closestTable;
        }
        const childTables = Array.from(targetElement.children).filter(child => child.tagName === 'TABLE');
        if (childTables.length === 1) {
            log("findRelatedTable: targetElement contains one child TABLE.");
            return childTables[0];
        }
        if (childTables.length > 1) {
             log("findRelatedTable: targetElement contains multiple child TABLEs. Returning targetElement itself for now.");
             return targetElement;
        }
        const nextSibling = targetElement.nextElementSibling;
        if (nextSibling && nextSibling.tagName === 'TABLE') {
            if (targetElement.textContent.trim().length < 100) {
                 log("findRelatedTable: nextElementSibling is a TABLE.");
                 return nextSibling;
            }
        }
        log("findRelatedTable: No clearly related single table found by structure, returning original targetElement.");
        return targetElement;
    }

    function getElementHtmlContent(element) {
        if (!element) return "";
        let html;
        let isConsideredTable = false;
        if (element.tagName === 'TABLE') {
            isConsideredTable = true;
            html = element.outerHTML;
            log("getElementHtmlContent: Element is TABLE, using outerHTML.");
        } else {
            html = element.innerHTML;
            log("getElementHtmlContent: Element is not TABLE, using innerHTML.");
        }
        html = html.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "");
        html = html.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, "");
        html = html.replace(/<!--[\s\S]*?-->/g, "");
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
        let displayElement = findRelatedTable(initialTargetElement);
        if (!displayElement) {
            log("Error: findRelatedTable returned null. Defaulting to initialTargetElement.");
            displayElement = initialTargetElement;
        }
        log(`getElementContext: Initial target is #${initialTargetElement.id}, displayElement is ${displayElement.tagName}#${displayElement.id || '(no id)'}`);
        const mainContentHtml = getElementHtmlContent(displayElement);
        contextParts.push(`<div style="${styleSubHeading}">--- 預覽內容 (ID: ${initialTargetElement.id || '無ID'}) ---</div>`);
        contextParts.push(`<div style="${styleTargetContainer}">${mainContentHtml || '(無可顯示內容)'}</div>`);
        let currentSibling;
        if (LINES_BEFORE > 0) {
            let beforeElementsHtml = [];
            currentSibling = initialTargetElement.previousElementSibling;
            for (let i = 0; i < LINES_BEFORE && currentSibling; i++) {
                if (currentSibling !== displayElement) {
                    const htmlSnippet = getElementHtmlContent(currentSibling);
                    if (htmlSnippet.trim()) beforeElementsHtml.unshift(htmlSnippet);
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
                    if (htmlSnippet.trim()) afterElementsHtml.push(htmlSnippet);
                }
                currentSibling = currentSibling.nextElementSibling;
            }
            if (afterElementsHtml.length > 0) {
                contextParts.push(`<div style="${styleSubHeading}">--- 後文 (${afterElementsHtml.length} 個兄弟元素) ---</div>`);
                contextParts.push(`<div style="margin-bottom: 5px; padding-bottom: 5px; border-bottom: 1px dotted #ddd; overflow: auto;">${afterElementsHtml.join('<hr style="border:none;border-top:1px dotted #ddd;margin:5px 0;">')}</div>`);
            }
        }
        if (!mainContentHtml.trim() && contextParts.length <= 2) {
             return "連結目標及其上下文均無可顯示的內容。";
        }
        return contextParts.join('');
    }

    // --- 事件監聽器 ---
    document.addEventListener('mouseover', function(event) {
        const aTag = event.target.closest('a');

        if (aTag && aTag.href) {
            // 如果滑鼠在預覽框內，則不處理連結懸停
            if (previewDiv && previewDiv.contains(event.target)) return;

            // 如果滑鼠仍在當前已顯示預覽的連結上，則清除隱藏計時器並返回
            if (aTag === currentHoveredLink && previewDiv && previewDiv.style.display === 'block') {
                 if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout);
                 hidePreviewTimeout = null;
                 return;
            }

            // 清除之前可能存在的顯示計時器 (處理快速移動到新連結的情況)
            if (showPreviewTimeout) {
                clearTimeout(showPreviewTimeout);
                showPreviewTimeout = null;
            }

            try {
                const url = new URL(aTag.href, window.location.href);
                // 僅處理頁內錨點連結
                if (url.hash && url.origin === window.location.origin && url.pathname === window.location.pathname) {
                    const targetId = decodeURIComponent(url.hash.substring(1));
                    if (!targetId) return;

                    const targetElement = document.getElementById(targetId);
                    if (targetElement) {
                        log(`Mouseover on link to #${targetId}. Scheduling preview in ${DELAY_BEFORE_SHOWING_PREVIEW}ms.`);
                        currentHoveredLink = aTag; // 立即更新當前懸停的連結

                        // 捕獲事件對象，以便在計時器回調中使用正確的滑鼠位置
                        const capturedEvent = event;

                        // 設置延遲顯示
                        showPreviewTimeout = setTimeout(() => {
                            // 在計時器觸發時，再次確認滑鼠是否還在目標連結上
                            // （雖然 mouseout 應該已經清除了計時器，但這是一個額外的保險）
                            if (currentHoveredLink !== aTag) {
                                log(`Timer for #${targetId} expired, but mouse is no longer on the link. Aborting.`);
                                return;
                            }
                            log(`Timer expired for #${targetId}. Fetching content and showing preview.`);
                            const contextText = getElementContext(targetElement);
                            showPreview(contextText, capturedEvent); // 使用捕獲的事件對象
                        }, DELAY_BEFORE_SHOWING_PREVIEW);

                    } else {
                         log("目標元素未找到: #" + targetId + " 在文檔:", window.location.href);
                         // 如果目標元素未找到，我們可能需要清除 currentHoveredLink，如果它是這個無效的 aTag
                         if (currentHoveredLink === aTag) currentHoveredLink = null;
                    }
                } else {
                    // 如果不是頁內連結，且之前 currentHoveredLink 指向它，現在移開了，則清空
                     if (currentHoveredLink === aTag) currentHoveredLink = null;
                }
            } catch (e) {
                log("處理連結 mouseover 錯誤:", e, "對於連結:", aTag.href);
                if (currentHoveredLink === aTag) currentHoveredLink = null;
            }
        }
    }, true);

    document.addEventListener('mouseout', function(event) {
        const aTag = event.target.closest('a');

        // 只處理當滑鼠移出的是當前我們正在追蹤的連結
        if (aTag && aTag === currentHoveredLink) {
            // 清除延遲顯示的計時器 (如果存在且尚未觸發)
            if (showPreviewTimeout) {
                clearTimeout(showPreviewTimeout);
                showPreviewTimeout = null;
                log(`Mouseout from link #${aTag.hash ? decodeURIComponent(aTag.hash.substring(1)) : aTag.href}, cleared showPreviewTimeout.`);
            }

            // 如果預覽框已顯示，且滑鼠不是移到預覽框本身，則延遲隱藏預覽框
            if (previewDiv && previewDiv.style.display !== 'none') { // 確保 previewDiv 存在且可見
                 if (event.relatedTarget !== previewDiv && (!previewDiv.contains(event.relatedTarget))) {
                    delayedHidePreview();
                }
            }
            // 注意：currentHoveredLink 的重置主要由 hidePreview 或 delayedHidePreview 裡的 hidePreview 處理
            // 或者在 mouseover 到一個新連結時被覆蓋。
            // 如果只是移出，計時器被取消，但預覽還沒顯示，currentHoveredLink 可能暫時保留，
            // 直到 mouseover 新連結或 preview 被正式隱藏。
        }
    }, true);

    if (document.readyState === "complete" || document.readyState === "interactive") {
        createPreviewDiv();
    } else {
        document.addEventListener("DOMContentLoaded", createPreviewDiv);
    }

    log("腳本初始化完成。");

})();
