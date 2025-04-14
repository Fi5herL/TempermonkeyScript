// ==UserScript==
// @name         報表頁面 - 匯出全年聚會資料
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  在聚會資料報表頁面添加按鈕，自動匯出選定年份內所有聚會類型的1-12月資料。
// @author       You
// @match        https://www.chlife-stat.org/attendace_report.php
// @icon         https://www.google.com/s2/favicons?sz=64&domain=chlife-stat.org
// @grant        GM_addStyle
// ==/UserScript==

(function() {
    'use strict';

    // --- 配置 ---
    const meetingSelectSelector = '#meeting';
    const yearFromSelectSelector = '#year_from';
    const monthFromSelectSelector = '#month_from';
    const yearToSelectSelector = '#year_to';
    const monthToSelectSelector = '#month_to';
    const originalExportButtonSelector = '#export';
    const buttonContainerSelector = '#export';
    const delayBetweenExportsMs = 5000;

    // --- 輔助函數：延遲 ---
    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    // --- 輔助函數：生成匯出紀錄檔 CSV 內容 ---
    function generateLogCsv(logEntries) {
        if (!logEntries || logEntries.length === 0) return null;

        const headers = ['匯出時間', '順序', '聚會類型'];
        const rows = logEntries.map(entry => [
            entry.timestamp,
            entry.sequence,
            entry.meetingName
        ]);

        // 處理 CSV 特殊字符 (簡單版本)
        const escapeCSV = (field) => {
            if (field === null || field === undefined) return '';
            const stringField = String(field);
            if (stringField.includes(',') || stringField.includes('"') || stringField.includes('\n')) {
                return `"${stringField.replace(/"/g, '""')}"`;
            }
            return stringField;
        };

        const headerString = headers.map(escapeCSV).join(',');
        const rowStrings = rows.map(row => row.map(escapeCSV).join(','));

        const BOM = '\uFEFF'; // UTF-8 BOM for Excel
        return BOM + headerString + '\n' + rowStrings.join('\n');
    }

     // --- 輔助函數：觸發 CSV 下載 (通用) ---
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

    // --- 主要執行函數 ---
    async function exportAllMeetingsForYear() {
        const newButton = document.getElementById('export-all-meetings-button');
        if (newButton) {
            newButton.disabled = true;
            newButton.textContent = '處理中...';
        }

        // --- 獲取元素 ---
        const meetingSelect = document.querySelector(meetingSelectSelector);
        const yearFromSelect = document.querySelector(yearFromSelectSelector);
        const monthFromSelect = document.querySelector(monthFromSelectSelector);
        const yearToSelect = document.querySelector(yearToSelectSelector);
        const monthToSelect = document.querySelector(monthToSelectSelector);
        const originalExportButton = document.querySelector(originalExportButtonSelector);

        if (!meetingSelect || !yearFromSelect || !monthFromSelect || !yearToSelect || !monthToSelect || !originalExportButton) {
            alert("錯誤：找不到必要的頁面元素，腳本無法執行。");
            if (newButton) { newButton.disabled = false; newButton.textContent = '匯出全年聚會資料'; }
            return;
        }

        const targetYear = yearFromSelect.value;
        if (!targetYear) {
             alert("錯誤：無法讀取目標年份。");
             if (newButton) { newButton.disabled = false; newButton.textContent = '匯出全年聚會資料'; }
             return;
        }
        console.log(`目標年份：${targetYear}`);

        const meetingOptions = Array.from(meetingSelect.options);
        const totalMeetings = meetingOptions.length;
        console.log(`找到 ${totalMeetings} 個聚會類型。`);

        // **新增**: 用於存儲匯出紀錄的陣列
        const exportLogEntries = [];

        // --- 循環處理每個聚會 ---
        try {
            for (let i = 0; i < totalMeetings; i++) {
                const option = meetingOptions[i];
                const meetingValue = option.value;
                const meetingName = option.text.trim(); // 去除可能的空白
                const currentStep = i + 1;

                console.log(`--- 處理第 ${currentStep}/${totalMeetings} 個聚會：${meetingName} (Value: ${meetingValue}) ---`);
                if (newButton) newButton.textContent = `匯出 ${meetingName} (${currentStep}/${totalMeetings})...`;

                // 1. 選擇聚會類型, 年份, 月份
                meetingSelect.value = meetingValue;
                yearFromSelect.value = targetYear;
                yearToSelect.value = targetYear;
                monthFromSelect.value = '1';
                monthToSelect.value = '12';

                // 等待一小段時間
                await sleep(300);

                // 2. 點擊原生的匯出按鈕
                console.log(`-------> 即將匯出：[${meetingName}] (${targetYear}年 1-12月) <-------`);
                originalExportButton.click();

                // **新增**: 記錄本次匯出信息
                const timestamp = new Date().toLocaleString(); // 獲取當前時間（本地格式）
                exportLogEntries.push({
                    timestamp: timestamp,
                    sequence: currentStep,
                    meetingName: meetingName
                });
                console.log(`已記錄匯出：順序 ${currentStep}, 聚會 ${meetingName}, 時間 ${timestamp}`);

                // 3. 等待設定的延遲時間
                console.log(`等待 ${delayBetweenExportsMs / 1000} 秒...`);
                await sleep(delayBetweenExportsMs);
            }

            console.log("所有聚會類型匯出請求已發送。");

            // --- **新增**: 生成並下載匯出紀錄檔 ---
            console.log("準備生成匯出紀錄檔...");
            const logCsvData = generateLogCsv(exportLogEntries);
            if (logCsvData) {
                const logFilename = `匯出紀錄_${targetYear}年_${new Date().toISOString().slice(0, 10)}.csv`;
                downloadCSV(logCsvData, logFilename);
                alert(`已完成 ${totalMeetings} 個聚會類型（${targetYear}年1-12月）的匯出請求。\n同時已下載匯出紀錄檔：${logFilename}\n請檢查瀏覽器下載。`);
            } else {
                alert(`已完成 ${totalMeetings} 個聚會類型（${targetYear}年1-12月）的匯出請求。\n但生成匯出紀錄檔失敗。\n請檢查瀏覽器下載。`);
            }

        } catch (error) {
            console.error("自動匯出過程中發生錯誤:", error);
            alert(`匯出過程中發生錯誤：\n${error.message}\n請查看控制台獲取更多資訊。`);
        } finally {
            if (newButton) {
                newButton.disabled = false;
                newButton.textContent = '匯出全年聚會資料';
            }
        }
    }

    // --- 函數：創建新的匯出按鈕 ---
    function createExportAllButton() {
        const referenceButton = document.querySelector(buttonContainerSelector);
        if (!referenceButton) {
            console.error("找不到原匯出按鈕，無法添加新按鈕。");
            return;
        }
        GM_addStyle(`
            #export-all-meetings-button {
                margin-left: 10px; padding: 5px 10px; background-color: #ffc107;
                color: black; border: 1px solid #d39e00; border-radius: 4px;
                cursor: pointer; font-size: 14px; vertical-align: middle;
            }
            #export-all-meetings-button:disabled {
                background-color: #cccccc; border-color: #aaaaaa; cursor: not-allowed; color: #666666;
            }
        `);
        const button = document.createElement('button');
        button.id = 'export-all-meetings-button';
        button.textContent = '匯出全年聚會資料';
        button.type = 'button';
        button.addEventListener('click', exportAllMeetingsForYear);
        referenceButton.parentNode.insertBefore(button, referenceButton.nextSibling);
        console.log("全年匯出按鈕已添加。");
    }

    // --- 主邏輯 ---
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', createExportAllButton);
    } else {
        createExportAllButton();
    }

})();
