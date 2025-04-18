// ==UserScript==
// @name         召會生活 自動單選匯出基本資料 (v4.2 - CSV增檔名)
// @namespace    http://tampermonkey.net/
// @version      4.2
// @description  無確認框。先取消全選，再依序單選(點擊<a>)並匯出，按鈕在匯出鈕旁。CSV增加原始檔名欄。
// @author       Fisher Li
// @match        https://www.chlife-stat.org/io.php
// @grant        GM_download
// @grant        GM_addStyle
// @require      https://code.jquery.com/jquery-1.4.2.min.js
// @require      https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js // 用於備用下載
// ==/UserScript==

/* global $j, jQuery, saveAs */

(function() {
    'use strict';

    // --- 設定區 ---
    const clubInputSelector = '#export-church li.level1 input[type="checkbox"][name="church"]';
    const exportButtonSelector = '#export';
    const clickDelay = 200;
    const exportDelay = 3000;
    const initialUncheckDelay = 100;
    const postUncheckDelay = 100;
    const initialLoadDelay = 300;
    // ** 新增：原始檔名設定 **
    const BASE_FILENAME = "_export"; // 基礎檔名
    const FILE_EXTENSION = "";       // 副檔名

    // --- 全域變數 ---
    let isRunning = false;
    let exportLog = [];

    // --- 輔助函數 ---
    const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
    function log(message) { console.log(`[AutoExport v4.2] ${message}`); }
    function errorLog(message, error) { console.error(`[AutoExport v4.2] ${message}`, error); }
    function getPageJquery() { if (typeof $j !== 'undefined') return $j; if (typeof jQuery !== 'undefined') return jQuery; errorLog("頁面上找不到 $j 或 jQuery！", null); return null; }
    function getClubName(checkboxInputElement) { try { const linkElement = checkboxInputElement.closest('a'); if (!linkElement) return `未知會所(無a)_${checkboxInputElement.value || Date.now()}`; const tempLink = linkElement.cloneNode(true); tempLink.querySelectorAll('ins, input').forEach(el => el.remove()); let name = tempLink.textContent.trim(); if (!name && linkElement.textContent) name = linkElement.textContent.replace(/[\s\u00A0]+/g, ' ').trim(); name = name.replace(/^[^a-zA-Z0-9\u4e00-\u9fa5]+/, '').trim(); return name || `未知會所(空名)_${checkboxInputElement.value || Date.now()}`; } catch (err) { errorLog('獲取會所名稱時出錯:', err); return `錯誤名稱_${Date.now()}`; } }

    // --- CSV 下載 (修改: 增加檔名欄) ---
    function downloadCSVLog() {
        if (exportLog.length === 0) { log("沒有匯出紀錄可供下載。"); return; }
        // ** 修改：增加 CSV Header **
        const csvHeader = "匯出時間,會所名稱,匯出順序,原始檔名";
        // ** 修改：增加 originalFilename 到輸出行 **
        const csvRows = exportLog.map(log =>
            `"${log.timestamp}","${log.clubName.replace(/"/g, '""')}","${log.sequence}","${log.originalFilename.replace(/"/g, '""')}"`
        );
        const csvContent = "\uFEFF" + csvHeader + "\n" + csvRows.join("\n");
        log("準備下載 CSV 紀錄檔..."); console.log(csvContent);
        try {
            GM_download({
                url: 'data:text/csv;charset=utf-8,' + encodeURIComponent(csvContent),
                name: `會所匯出紀錄_${new Date().toISOString().slice(0,10)}.csv`, saveAs: true,
                onload: () => { log(`CSV 紀錄檔下載請求已發送 (GM_download)。共 ${exportLog.length} 筆紀錄。`); },
                onerror: (error) => { errorLog("GM_download 失敗:", error); alert(`CSV 下載失敗 (GM_download): ${error.error}\n嘗試備用方法。`); fallbackDownload(csvContent); }
            });
        } catch (e) { errorLog("GM_download 不可用，嘗試 FileSaver:", e); fallbackDownload(csvContent); }
    }
    function fallbackDownload(csvContent) {
        try { if (typeof saveAs === 'function') { const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' }); saveAs(blob, `會所匯出紀錄_${new Date().toISOString().slice(0,10)}.csv`); log(`CSV 紀錄檔下載請求已發送 (FileSaver)。共 ${exportLog.length} 筆紀錄。`); } else { errorLog("FileSaver.js 未載入"); alert("無法觸發CSV下載(備用)。"); console.log("--- CSV:\n", csvContent); }
        } catch(err) { errorLog("FileSaver 下載失敗:", err); alert("CSV下載失敗(備用)。"); console.log("--- CSV:\n", csvContent); }
    }

    // --- 主要處理流程 (修改: 增加計算原始檔名並記錄) ---
    async function processAllExports() {
        if (isRunning) { return; }
        const jq = getPageJquery();
        if (!jq) { alert("錯誤：無法找到 jQuery ($j)，無法繼續執行。"); return; }

        log('開始自動匯出流程 (v4.2)...');
        isRunning = true;
        exportLog = [];
        updateButtonState(true, '初始化...');

        // --- 初始取消勾選 ---
        log("執行初始取消勾選...");
        let allInputs = [];
        try { /* ...與v4.1相同... */
            await delay(500); allInputs = Array.from(document.querySelectorAll(clubInputSelector)); log(`找到 ${allInputs.length} 個會所 input 進行初始檢查。`); if (allInputs.length === 0) throw new Error("未找到任何會所 Input 元素！");
            let uncheckedCount = 0;
            for (const input of allInputs) { if (input.checked) { const anchorToUncheck = input.closest('a'); const clubName = getClubName(input); if (anchorToUncheck) { log(`初始狀態: ${clubName} 已勾選，取消中...`); jq(anchorToUncheck).trigger('click'); uncheckedCount++; await delay(initialUncheckDelay); } else { log(`警告：找不到 ${clubName} 的 <a> 進行初始取消。`); } } }
            if (uncheckedCount > 0) { log(`初始取消完成，共處理 ${uncheckedCount} 個。`); await delay(postUncheckDelay); } else { log("初始狀態無需取消。"); }
        } catch (error) { errorLog("初始取消勾選失敗:", error); alert(`錯誤：無法完成初始取消勾選。\n${error.message}\n腳本停止。`); isRunning = false; updateButtonState(false); return; }
        log("初始取消勾選完成。");

        // --- 循環處理 ---
        const exportButton = document.querySelector(exportButtonSelector);
        if (!exportButton) { errorLog("找不到匯出按鈕!", null); alert("錯誤：找不到匯出按鈕！腳本停止。"); isRunning = false; updateButtonState(false); return; }

        for (let i = 0; i < allInputs.length; i++) {
            const targetInput = allInputs[i];
            const targetAnchor = targetInput.closest('a');
            const sequence = i + 1;
            const clubName = getClubName(targetInput);
            log(`--- [${sequence}/${allInputs.length}] 處理: ${clubName} ---`);
            updateButtonState(true, `處理中 (${sequence}/${allInputs.length})...`);
            if (!targetAnchor) { errorLog(`找不到 ${clubName} 的 <a> 標籤，跳過此項。`, null); continue; }

            try {
                // 1. 確保單選
                // ...與v4.1相同...
                log("確保只有目標被選中...");
                for (const otherInput of allInputs) { if (otherInput !== targetInput && otherInput.checked) { const otherAnchor = otherInput.closest('a'); if(otherAnchor){ log(`取消非目標項: ${getClubName(otherInput)}`); jq(otherAnchor).trigger('click'); await delay(clickDelay); } } }
                // 2. 勾選目標
                // ...與v4.1相同...
                if (!targetInput.checked) { log(`勾選目標: ${clubName}`); jq(targetAnchor).trigger('click'); await delay(clickDelay + 100); if (!targetInput.checked) log(`警告: 點擊後 ${clubName} 似乎仍未勾選。`); else log(`${clubName} 已勾選。`); } else { log(`${clubName} 已是勾選狀態。`); }

                // ** 新增：計算原始檔名 **
                let originalFilename = '';
                if (sequence === 1) {
                    originalFilename = `${BASE_FILENAME}${FILE_EXTENSION}`;
                } else if (sequence > 1) {
                    originalFilename = `${BASE_FILENAME} (${sequence - 1})${FILE_EXTENSION}`;
                } else {
                    originalFilename = '計算錯誤'; // 理論上不會發生
                }
                log(`對應原始檔名: ${originalFilename}`);

                // 3. 點擊匯出
                log("點擊匯出按鈕..."); exportButton.click();
                // 4. 等待
                log(`等待 ${exportDelay / 1000} 秒讓下載處理...`); await delay(exportDelay);
                // 5. 記錄 (修改: 增加原始檔名)
                const timestamp = new Date().toLocaleString('zh-TW', { hour12: false });
                exportLog.push({ timestamp, clubName, sequence, originalFilename }); // <-- 增加原始檔名
                log(`紀錄已保存: ${timestamp}, ${clubName}, ${sequence}, ${originalFilename}`); // <-- 更新日誌

            } catch (error) {
                errorLog(`處理 ${clubName} 時發生錯誤:`, error);
                alert(`處理 ${clubName} 時發生錯誤，腳本將停止。\n錯誤: ${error.message}\n請檢查控制台。`);
                log("因錯誤停止流程。");
                isRunning = false; updateButtonState(false); return;
            }
        }

        log('所有會所處理循環完成。');
        isRunning = false;
        updateButtonState(false);

        // --- 下載記錄檔 ---
        downloadCSVLog();
    }

    // --- UI (與 v4.1 相同) ---
    function createControlButton() {
        const originalExportButton = document.querySelector(exportButtonSelector);
        if (!originalExportButton) { errorLog(`找不到原始匯出按鈕 (${exportButtonSelector})，無法定位腳本按鈕。`, null); return; }
        const button = document.createElement('button'); button.id = 'autoExportToggleButton_v42'; button.type = 'button';
        originalExportButton.insertAdjacentElement('afterend', button);
        GM_addStyle(`
            #autoExportToggleButton_v42 { /* Use New ID */
                padding: 3px 12px; font-size: inherit; border: 1px solid #a9a9a9; border-radius: 3px; cursor: pointer;
                background-color: #f0f0f0; color: #333; box-shadow: 1px 1px 2px rgba(0,0,0,0.1); transition: background-color 0.2s;
                min-width: 100px; text-align: center; display: inline-block; vertical-align: middle; margin-left: 8px;
            }
            #autoExportToggleButton_v42:hover { background-color: #e0e0e0; border-color: #888888; }
            #autoExportToggleButton_v42.running { background-color: #ffe0b2; border-color: #ffcc80; color: #a0522d; cursor: not-allowed; opacity: 0.8; }
            #autoExportToggleButton_v42.running.initializing { background-color: #fffacd; border-color: #ffe0b2; }
        `);
        log('自動匯出控制按鈕已添加 (v4.2 - CSV增檔名)。');
    }
    function updateButtonState(running, statusText = null) { /* ... 與 v4.1 相同 ... */
        const button = document.getElementById('autoExportToggleButton_v42'); if (!button) return; button.classList.remove('initializing');
        if (running) { let text = statusText || `處理中...`; button.textContent = text; button.classList.add('running'); if (statusText === '初始化...') button.classList.add('initializing'); button.disabled = true; button.onclick = null; }
        else { button.textContent = '開始自動匯出'; button.classList.remove('running'); button.disabled = false; button.onclick = processAllExports; }
    }

    // --- 腳本入口 (與 v4.1 相同) ---
    window.addEventListener('load', async () => {
        log(`頁面載入完成，初始化腳本 v4.2... (等待 ${initialLoadDelay}ms)`); await delay(initialLoadDelay);
        const jq = getPageJquery(); if (!jq) { alert("錯誤：無法找到 jQuery ($j)，腳本無法運行。"); return; } log(`jQuery 版本: ${jq.fn.jquery}`);
        createControlButton(); updateButtonState(false); log('腳本準備就緒。等待用戶點擊按鈕。');
        const initialInputs = document.querySelectorAll(clubInputSelector); if (initialInputs.length === 0) log("警告：似乎未找到任何會所 input 元素。"); else log(`檢測到 ${initialInputs.length} 個初始 input。`);
    });

})();
