// ==UserScript==
// @name         連結目標上下文預覽 (智能表格捕獲 & UL標準連結預覽)
// @namespace    http://tampermonkey.net/
// @version      2025-05-24.6
// @description  懸停內部連結預覽上下文；自動替換標準號為連結，並懸停時在新浮窗中預覽目標網頁（支持內部連結點擊）。
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
    const MAX_CONTENT_LENGTH_NON_TABLE = 3000;
    const DELAY_BEFORE_SHOWING_PREVIEW = 500;

    // --- 全局變量 ---
    const PREVIEW_DIV_ID = 'userscript-internal-link-preview-div';
    let previewDiv = null;
    let currentHoveredLink = null;
    let hidePreviewTimeout = null;
    let showPreviewTimeout = null;

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

    // ===================================================================
    // === 掃描並將標準文本替換為連結的功能 ===
    // ===================================================================
    function createStandardLink(catalog, standard, fullMatchText) {
        const url = `https://www.ulsestandards.org/uls-standardsdocs/StandardDocuments.aspx?Catalog=${catalog.toUpperCase()}&Standard=${encodeURIComponent(standard)}`;
        const link = document.createElement('a');
        link.href = url;
        link.target = '_blank';
        link.rel = 'noopener noreferrer';
        link.textContent = `[${fullMatchText}]`;
        Object.assign(link.style, {
            color: '#005596',
            textDecoration: 'underline',
            fontWeight: 'normal',
            cursor: 'pointer',
            backgroundColor: 'transparent',
            border: 'none',
            padding: '0',
            margin: '0',
        });
        return link;
    }

    function addULStandardLinks() {
        log("開始掃描並替換 UL/ULC/ULE 標準...");
        const rootElement = document.body;
        if (!rootElement) return;

        const standardRegex = /\b(ULC?|ULE)\s+(\d[a-zA-Z0-9.-]*)\b/g;

        const walker = document.createTreeWalker(rootElement, NodeFilter.SHOW_TEXT, {
            acceptNode: function(node) {
                const parent = node.parentElement;
                if (parent.closest('script, style, a[target="_blank"]')) {
                    return NodeFilter.FILTER_REJECT;
                }
                return NodeFilter.FILTER_ACCEPT;
            }
        }, false);

        const nodesToProcess = [];
        let currentNode;
        while ((currentNode = walker.nextNode())) {
            if (standardRegex.test(currentNode.nodeValue)) {
                nodesToProcess.push(currentNode);
                standardRegex.lastIndex = 0;
            }
        }

        log(`找到 ${nodesToProcess.length} 個包含標準的文字節點進行替換。`);

        nodesToProcess.reverse().forEach(node => {
            const parent = node.parentNode;
            if (!parent) return;

            standardRegex.lastIndex = 0;
            const matches = Array.from(node.nodeValue.matchAll(standardRegex));
            if (matches.length === 0) return;

            const fragment = document.createDocumentFragment();
            let lastIndex = 0;

            matches.forEach(match => {
                fragment.appendChild(document.createTextNode(node.nodeValue.substring(lastIndex, match.index)));
                const [fullMatchText, catalog, standard] = match;
                const link = createStandardLink(catalog, standard, fullMatchText);
                fragment.appendChild(link);
                lastIndex = match.index + fullMatchText.length;
            });

            fragment.appendChild(document.createTextNode(node.nodeValue.substring(lastIndex)));
            parent.replaceChild(fragment, node);
        });

        log("UL 標準連結替換完成。");
    }

    // ===================================================================
    // === 預覽功能核心 ===
    // ===================================================================
    function createPreviewDiv() {
        if (document.getElementById(PREVIEW_DIV_ID)) {
            previewDiv = document.getElementById(PREVIEW_DIV_ID);
            return;
        }
        previewDiv = document.createElement('div');
        previewDiv.id = PREVIEW_DIV_ID;
        Object.assign(previewDiv.style, {
            position: 'fixed',
            zIndex: '2147483647',
            border: '1px solid #888',
            backgroundColor: '#fff',
            boxShadow: '0 5px 15px rgba(0,0,0,0.2)',
            display: 'none',
            overflow: 'hidden',
            resize: 'both',
        });
        document.body.appendChild(previewDiv);

        previewDiv.addEventListener('mouseenter', () => {
            if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout);
            hidePreviewTimeout = null;
        });
        previewDiv.addEventListener('mouseleave', (event) => {
            if (event.relatedTarget !== currentHoveredLink) delayedHidePreview(100);
        });
    }

    function showPreview(htmlContent, event, isIframe = false) {
        if (!previewDiv) createPreviewDiv();
        if (!previewDiv) return;
        if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout);
        hidePreviewTimeout = null;

        if (isIframe) {
            Object.assign(previewDiv.style, {
                padding: '0px', width: '80vw', height: '75vh',
                maxWidth: '95vw', maxHeight: '90vh', overflow: 'hidden'
            });
        } else {
            Object.assign(previewDiv.style, {
                padding: '12px', width: 'auto', height: 'auto',
                maxWidth: '85vw', maxHeight: '80vh', overflow: 'auto'
            });
        }

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
            previewDiv.innerHTML = '';
            Object.assign(previewDiv.style, {
                width: 'auto', height: 'auto', padding: '12px', overflow: 'hidden'
            });
        }
        currentHoveredLink = null;
        if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout);
        hidePreviewTimeout = null;
        if (showPreviewTimeout) {
            clearTimeout(showPreviewTimeout);
            showPreviewTimeout = null;
        }
    }

    function delayedHidePreview(delay = 300) {
        if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout);
        hidePreviewTimeout = setTimeout(() => {
            if (previewDiv && previewDiv.matches(':hover')) return;
            hidePreview();
        }, delay);
    }

    function getElementContext(initialTargetElement) {
        let contextParts = [];
        const styleSubHeading = "font-weight: bold; color: #555; margin-top: 8px; margin-bottom: 4px; display: block; border-bottom: 1px solid #eee; padding-bottom: 3px;";
        const styleTargetContainer = "border: 1px solid #ccc; padding: 10px; margin: 5px 0; display: block; overflow: auto; background-color: #f9f9f9;";
        let displayElement = initialTargetElement.closest('table') || initialTargetElement;
        const mainContentHtml = displayElement.outerHTML;
        contextParts.push(`<div style="${styleSubHeading}">--- 預覽內容 (ID: ${initialTargetElement.id || '無ID'}) ---</div>`);
        contextParts.push(`<div style="${styleTargetContainer}">${mainContentHtml || '(無可顯示內容)'}</div>`);
        return contextParts.join('');
    }

    // --- 事件監聽器 ---
    document.addEventListener('mouseover', function(event) {
        const aTag = event.target.closest('a');
        if (!aTag || !aTag.href) return;
        if (previewDiv && previewDiv.contains(event.target)) return;
        if (aTag === currentHoveredLink && previewDiv && previewDiv.style.display === 'block') {
             if (hidePreviewTimeout) clearTimeout(hidePreviewTimeout);
             hidePreviewTimeout = null;
             return;
        }

        if (showPreviewTimeout) clearTimeout(showPreviewTimeout);

        currentHoveredLink = aTag;
        const capturedEvent = event;

        showPreviewTimeout = setTimeout(() => {
            if (currentHoveredLink !== aTag) return;

            try {
                if (aTag.target === '_blank' && aTag.href.includes('StandardDocuments.aspx')) {
                    log(`懸停於標準連結: ${aTag.href}。準備 iframe 預覽。`);
                    // --- 核心修改處 ---
                    // 新增了 allow-popups 來允許 iframe 內的連結打開新視窗
                    const iframeContent = `<iframe src="${aTag.href}" style="width: 100%; height: 100%; border: none;" sandbox="allow-forms allow-scripts allow-same-origin allow-popups"></iframe>`;
                    showPreview(iframeContent, capturedEvent, true);
                }
                else {
                    const url = new URL(aTag.href, window.location.href);
                    if (url.hash && url.origin === window.location.origin && url.pathname === window.location.pathname) {
                        const targetId = decodeURIComponent(url.hash.substring(1));
                        if (!targetId) return;
                        const targetElement = document.getElementById(targetId);
                        if (targetElement) {
                            log(`懸停於內部連結 #${targetId}。準備上下文預覽。`);
                            const contextText = getElementContext(targetElement);
                            showPreview(contextText, capturedEvent, false);
                        }
                    }
                }
            } catch (e) {
                log("處理連結 mouseover 錯誤:", e, "對於連結:", aTag.href);
            }
        }, DELAY_BEFORE_SHOWING_PREVIEW);

    }, true);

    document.addEventListener('mouseout', function(event) {
        const aTag = event.target.closest('a');
        if (aTag && aTag === currentHoveredLink) {
            if (showPreviewTimeout) {
                clearTimeout(showPreviewTimeout);
                showPreviewTimeout = null;
            }
            if (previewDiv && previewDiv.style.display !== 'none') {
                 if (event.relatedTarget !== previewDiv && (!previewDiv.contains(event.relatedTarget))) {
                    delayedHidePreview();
                }
            }
        }
    }, true);

    // --- 腳本初始化 ---
    function initializeScript() {
        createPreviewDiv();
        setTimeout(addULStandardLinks, 500);
    }

    if (document.readyState === "complete" || document.readyState === "interactive") {
        initializeScript();
    } else {
        document.addEventListener("DOMContentLoaded", initializeScript);
    }

    log("腳本初始化完成。");

})();
