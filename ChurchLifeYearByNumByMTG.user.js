// ==UserScript==
// @name         匯出各會所指定聚會資料
// @namespace    http://tampermonkey.net/
// @version      2.4
// @description  自動逐一勾選單一會所，匯出選定年份的指定項目資料(可由使用者設定，預設勾選特定項目)，遍歷所有會所，並生成匯出紀錄CSV。
// @author       Fisher Li
// @match        https://www.chlife-stat.org/attendace_report.php
// @icon         https://www.google.com/s2/favicons?sz=64&domain=chlife-stat.org
// @grant        GM_addStyle
// @grant        GM_log
// @grant        GM_setValue
// @grant        GM_getValue
// ==/UserScript==

(function() {
    'use strict';

    // --- Configuration ---
    const treeContainerSelector = '#church';
    const meetingSelectSelector = '#meeting';
    const yearFromSelectSelector = '#year_from';
    const monthFromSelectSelector = '#month_from';
    const yearToSelectSelector = '#year_to';
    const monthToSelectSelector = '#month_to';
    const originalExportButtonSelector = '#export';
    const buttonContainerSelector = '#export';
    const delayBetweenExportsMs = 5000;
    const delayAfterCheckboxClickMs = 1000;
    const selectedMeetingsStorageKey = 'chlife_stat_selected_meetings_v2'; // Key remains the same

    // *** Define the new default selected meetings ***
    const defaultSelectedMeetings = ["主日", "晨興", "禱告", "小排"];

    // --- Global Variable for Selected Meetings ---
    let selectedMeetings = [];

    // --- Helper Functions (clickSpecificCheckbox, isCheckboxChecked, generateLogCsv, downloadCSV - unchanged from v2.3) ---
    function sleep(ms) {
        /* ... */
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    function clickSpecificCheckbox(hallId) {
        /* ... */
        return new Promise(async (resolve, reject) => {
            let escapedHallId;
            try {
                escapedHallId = CSS.escape(hallId);
            } catch (e) {
                if (/^\d/.test(hallId)) {
                    escapedHallId = '\\3' + hallId.charAt(0) + ' ' + hallId.slice(1);
                } else {
                    escapedHallId = hallId;
                }
            }
            const selectorsToTry = [`${treeContainerSelector} li#${escapedHallId} > a > ins.jstree-checkbox`, `${treeContainerSelector} li#${escapedHallId} > ins.jstree-checkbox`];
            let targetCheckbox = null;
            for (const selector of selectorsToTry) {
                targetCheckbox = document.querySelector(selector);
                if (targetCheckbox) break;
            }
            if (targetCheckbox) {
                const clickEvent = new MouseEvent('click', {
                    bubbles: true,
                    cancelable: true
                });
                targetCheckbox.dispatchEvent(clickEvent);
                GM_log(`Dispatched click for Hall ID ${hallId}.`);
                await sleep(delayAfterCheckboxClickMs);
                resolve(true);
            } else {
                GM_log(`Error: Could not find checkbox for Hall ID ${hallId}.`);
                reject(new Error(`無法找到 ID 為 ${hallId} 的會所核取方塊。`));
            }
        });
    }

    function isCheckboxChecked(hallId) {
        /* ... */
        let escapedHallId;
        try {
            escapedHallId = CSS.escape(hallId);
        } catch {
            escapedHallId = hallId;
        }
        const listItem = document.querySelector(`${treeContainerSelector} li#${escapedHallId}`);
        if (listItem && listItem.classList.contains('jstree-checked')) return true;
        const anchor = document.querySelector(`${treeContainerSelector} li#${escapedHallId} > a`);
        if (anchor && anchor.classList.contains('jstree-checked')) return true;
        const checkbox = document.querySelector(`${treeContainerSelector} li#${escapedHallId} ins.jstree-checkbox`);
        if (checkbox) {
            const parentAnchor = checkbox.closest('a');
            if (parentAnchor && parentAnchor.classList.contains('jstree-checked')) return true;
            const parentLi = checkbox.closest('li');
            if (parentLi && parentLi.classList.contains('jstree-checked')) return true;
        }
        return false;
    }

    function generateLogCsv(logEntries) {
        /* ... */
        if (!logEntries || logEntries.length === 0) return null;
        const headers = ['匯出時間', '會所名稱', '項目名稱', '原始檔案名稱'];
        const rows = logEntries.map(entry => [entry.timestamp, entry.locationName, entry.meetingName, entry.originalFilename]);
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
        const BOM = '\uFEFF';
        return BOM + headerString + '\n' + rowStrings.join('\n');
    }

    function downloadCSV(csvString, filename) {
        /* ... */
        if (!csvString) return;
        const blob = new Blob([csvString], {
            type: 'text/csv;charset=utf-8;'
        });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.setAttribute("href", url);
        link.setAttribute("download", filename);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        GM_log(`CSV 檔案 "${filename}" 已觸發下載。`);
    }
    // --- Settings Modal Functions (showSelectSettingsModal, hideSelectSettingsModal, saveSelectSettings - unchanged from v2.3) ---
    function showSelectSettingsModal() {
        /* ... */
        const modal = document.getElementById('select-settings-modal');
        const overlay = document.getElementById('select-settings-overlay');
        const listContainer = document.getElementById('select-settings-list');
        if (!modal || !overlay || !listContainer) {
            GM_log("Error: Settings modal elements not found!");
            return;
        }
        const meetingSelect = document.querySelector(meetingSelectSelector);
        if (!meetingSelect) {
            alert("錯誤：找不到聚會項目下拉選單，無法載入設定。");
            return;
        }
        const meetingOptions = Array.from(meetingSelect.options).filter(opt => opt.value && opt.text.trim()).map(opt => opt.text.trim());
        listContainer.innerHTML = '';
        meetingOptions.forEach(name => {
            const listItem = document.createElement('div');
            listItem.style.marginBottom = '5px';
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `select-setting-${name.replace(/\s+/g, '-')}`;
            checkbox.value = name;
            checkbox.checked = selectedMeetings.includes(name);
            checkbox.style.marginRight = '8px';
            const label = document.createElement('label');
            label.htmlFor = checkbox.id;
            label.textContent = name;
            label.style.cursor = 'pointer';
            listItem.appendChild(checkbox);
            listItem.appendChild(label);
            listContainer.appendChild(listItem);
        });
        overlay.style.display = 'block';
        modal.style.display = 'block';
        GM_log("顯示'要下載項目'設定視窗。");
    }

    function hideSelectSettingsModal() {
        /* ... */
        const modal = document.getElementById('select-settings-modal');
        const overlay = document.getElementById('select-settings-overlay');
        if (modal && overlay) {
            modal.style.display = 'none';
            overlay.style.display = 'none';
            GM_log("隱藏'要下載項目'設定視窗。");
        }
    }

    function saveSelectSettings() {
        /* ... */
        const listContainer = document.getElementById('select-settings-list');
        if (!listContainer) return;
        const checkboxes = listContainer.querySelectorAll('input[type="checkbox"]');
        const newSelectedList = [];
        checkboxes.forEach(cb => {
            if (cb.checked) {
                newSelectedList.push(cb.value);
            }
        });
        selectedMeetings = newSelectedList;
        try {
            GM_setValue(selectedMeetingsStorageKey, JSON.stringify(selectedMeetings));
            GM_log("已儲存新的'要下載項目'列表:", selectedMeetings);
            // alert("'要下載項目'設定已儲存！");
        } catch (e) {
            GM_log("儲存'要下載項目'設定時發生錯誤:", e);
            alert("儲存設定失敗，請檢查控制台。");
        }
        hideSelectSettingsModal();
    }
    // --- Main Execution Function (exportAllLocationsAndMeetings - unchanged from v2.3) ---
    async function exportAllLocationsAndMeetings() {
        /* ... */
        const newButton = document.getElementById('export-single-location-items-button');
        if (newButton) {
            newButton.disabled = true;
            newButton.textContent = '準備中...';
        }
        const treeContainer = document.querySelector(treeContainerSelector);
        const meetingSelect = document.querySelector(meetingSelectSelector);
        const yearFromSelect = document.querySelector(yearFromSelectSelector);
        const monthFromSelect = document.querySelector(monthFromSelectSelector);
        const yearToSelect = document.querySelector(yearToSelectSelector);
        const monthToSelect = document.querySelector(monthToSelectSelector);
        const originalExportButton = document.querySelector(originalExportButtonSelector);
        if (!treeContainer || !meetingSelect || !yearFromSelect || !monthFromSelect || !yearToSelect || !monthToSelect || !originalExportButton) {
            alert("錯誤：找不到必要的頁面元素，無法開始匯出。");
            if (newButton) {
                newButton.disabled = false;
                newButton.textContent = '匯出各會所項目(逐一)';
            }
            return;
        }
        const targetYear = yearFromSelect.value;
        if (!targetYear) {
            alert("錯誤：無法讀取目標年份。");
            return;
        }
        GM_log(`目標年份：${targetYear}`);
        GM_log(`目前設定要下載的項目: ${selectedMeetings.join(', ') || '(無)'}`);
        if (selectedMeetings.length === 0) {
            alert("警告：您尚未在設定中勾選任何要下載的項目。請點擊 ⚙️ 按鈕設定。");
            if (newButton) {
                newButton.disabled = false;
                newButton.textContent = '匯出各會所項目(逐一)';
            }
            return;
        }
        const hallListItems = Array.from(treeContainer.querySelectorAll('li.level1[id]'));
        if (hallListItems.length === 0) {
            hallListItems.push(...treeContainer.querySelectorAll('li[id]:not(#\\30 )'));
        }
        const hallInfoArray = hallListItems.map(li => {
            const id = li.id;
            const anchor = li.querySelector('a');
            const name = anchor ? anchor.textContent.replace(/[\n\r\t]/g, '').trim() : `會所 ID ${id}`;
            return {
                id,
                name
            };
        }).filter(hall => hall.id && hall.name);
        if (hallInfoArray.length === 0) {
            alert("錯誤：在 jsTree 中找不到任何會所項目。");
            return;
        }
        GM_log(`找到 ${hallInfoArray.length} 個會所`);
        const meetingOptions = Array.from(meetingSelect.options).filter(opt => opt.value && opt.text.trim()).map(opt => ({
            value: opt.value,
            name: opt.text.trim()
        }));
        if (meetingOptions.length === 0) {
            alert("錯誤：找不到任何聚會項目。");
            return;
        }
        GM_log(`找到 ${meetingOptions.length} 個聚會項目`);
        const exportLogEntries = [];
        let downloadCounter = 1;
        try {
            GM_log("確保所有會所都取消勾選...");
            if (newButton) newButton.textContent = '取消全選...';
            for (const hall of hallInfoArray) {
                if (isCheckboxChecked(hall.id)) {
                    GM_log(`  - 取消勾選: ${hall.name}`);
                    await clickSpecificCheckbox(hall.id);
                }
            }
            GM_log("所有會所應已取消勾選。");
            await sleep(500);
            for (let i = 0; i < hallInfoArray.length; i++) {
                const currentHall = hallInfoArray[i];
                const currentHallNum = i + 1;
                GM_log(`--- [${currentHallNum}/${hallInfoArray.length}] 處理會所: ${currentHall.name} ---`);
                if (newButton) newButton.textContent = `處理中: ${currentHall.name} (${currentHallNum}/${hallInfoArray.length})...`;
                GM_log(`  勾選會所: ${currentHall.name}`);
                if (!isCheckboxChecked(currentHall.id)) {
                    await clickSpecificCheckbox(currentHall.id);
                    if (!isCheckboxChecked(currentHall.id)) {
                        GM_log(`  警告：勾選 ${currentHall.name} 失敗。跳過此會所。`);
                        continue;
                    }
                } else {
                    GM_log(`  會所 ${currentHall.name} 已勾選。`);
                }
                let exportedMeetingCount = 0;
                for (const meeting of meetingOptions) {
                    if (!selectedMeetings.includes(meeting.name)) {
                        continue;
                    }
                    exportedMeetingCount++;
                    GM_log(`    --- [項目 ${exportedMeetingCount}] 準備匯出: ${meeting.name} ---`);
                    if (newButton) newButton.textContent = `匯出 ${currentHall.name}: ${meeting.name}...`;
                    meetingSelect.value = meeting.value;
                    yearFromSelect.value = targetYear;
                    yearToSelect.value = targetYear;
                    monthFromSelect.value = '1';
                    monthToSelect.value = '12';
                    await sleep(300);
                    GM_log(`      ------> 匯出: [${currentHall.name}] - [${meeting.name}] (${targetYear}年) <-------`);
                    originalExportButton.click();
                    const timestamp = new Date().toLocaleString('zh-TW', {
                        hour12: false
                    });
                    let originalFilename;
                    if (downloadCounter === 1) {
                        originalFilename = 'export';
                    } else {
                        originalFilename = `export (${downloadCounter - 1})`;
                    }
                    exportLogEntries.push({
                        timestamp: timestamp,
                        locationName: currentHall.name,
                        meetingName: meeting.name,
                        originalFilename: originalFilename
                    });
                    GM_log(`      已記錄: ${currentHall.name}, ${meeting.name}, ${timestamp}, 檔名: ${originalFilename}`);
                    downloadCounter++;
                    GM_log(`      等待 ${delayBetweenExportsMs / 1000} 秒...`);
                    await sleep(delayBetweenExportsMs);
                }
                GM_log(`  處理完成，取消勾選會所: ${currentHall.name}`);
                if (isCheckboxChecked(currentHall.id)) {
                    await clickSpecificCheckbox(currentHall.id);
                }
                GM_log(`--- 會所 ${currentHall.name} 處理完畢 ---`);
            }
            GM_log("所有會所和選定項目的匯出請求已完成。");
            GM_log("準備生成匯出紀錄檔...");
            const logCsvData = generateLogCsv(exportLogEntries);
            if (logCsvData) {
                const logFilename = `聚會資料_${targetYear}年_${new Date().toISOString().slice(0, 10)}.csv`;
                downloadCSV(logCsvData, logFilename);
                alert(`匯出完成！\n共觸發 ${downloadCounter - 1} 個檔案下載。\n紀錄檔 "${logFilename}" 已下載。\n請檢查瀏覽器下載。`);
            } else {
                alert(`匯出請求已完成，但無任何項目被匯出或紀錄檔生成失敗。\n共觸發 ${downloadCounter - 1} 個檔案下載。\n請檢查瀏覽器下載。`);
            }
        } catch (error) {
            console.error("自動匯出過程中發生錯誤:", error);
            GM_log(`自動匯出錯誤: ${error.message}\n${error.stack}`);
            alert(`匯出過程中發生錯誤：\n${error.message}\n請查看瀏覽器控制台(F12)獲取詳細資訊。`);
        } finally {
            if (newButton) {
                newButton.disabled = false;
                newButton.textContent = '匯出各會所項目(逐一)';
            }
            GM_log("腳本執行完畢。");
        }
    }

    // --- Load Settings and Create UI ---

    function loadSelectedMeetings() {
        const storedSettings = GM_getValue(selectedMeetingsStorageKey, null);
        if (storedSettings) {
            try {
                selectedMeetings = JSON.parse(storedSettings);
                GM_log("成功從儲存空間載入'要下載項目':", selectedMeetings);
                // Basic validation: ensure it's an array
                if (!Array.isArray(selectedMeetings)) {
                    throw new Error("Loaded data is not an array.");
                }
            } catch (e) {
                GM_log("解析儲存的'要下載項目'設定時發生錯誤或格式不符，將使用預設值:", e);
                // *** USE NEW DEFAULT LIST ***
                selectedMeetings = [...defaultSelectedMeetings]; // Use the specified default list
                GM_setValue(selectedMeetingsStorageKey, JSON.stringify(selectedMeetings));
            }
        } else {
            // *** USE NEW DEFAULT LIST ***
            GM_log("未找到儲存的'要下載項目'設定，使用預設值:", defaultSelectedMeetings);
            selectedMeetings = [...defaultSelectedMeetings]; // Use the specified default list
            GM_setValue(selectedMeetingsStorageKey, JSON.stringify(selectedMeetings)); // Store the default
        }
        // Final check to ensure it's an array (redundant if catch block is correct, but safe)
        if (!Array.isArray(selectedMeetings)) {
            GM_log("最終檢查發現 selectedMeetings 不是陣列，重設為預設值。");
            selectedMeetings = [...defaultSelectedMeetings];
            GM_setValue(selectedMeetingsStorageKey, JSON.stringify(selectedMeetings));
        }
    }

    // Helper to get all meeting names (still needed for modal population, but not for default setting)
    function getAllMeetingNames() {
        const meetingSelect = document.querySelector(meetingSelectSelector);
        if (!meetingSelect) return [];
        return Array.from(meetingSelect.options)
            .filter(opt => opt.value && opt.text.trim())
            .map(opt => opt.text.trim());
    }

    // --- Create UI (createUI function unchanged from v2.3) ---
    function createUI() {
        /* ... */
        const referenceButton = document.querySelector(buttonContainerSelector);
        if (!referenceButton) {
            GM_log("找不到原匯出按鈕，無法添加UI。");
            return;
        }
        if (!document.getElementById('export-single-location-items-button')) {
            const exportButton = document.createElement('button');
            exportButton.id = 'export-single-location-items-button';
            exportButton.textContent = '匯出各會所項目(逐一)';
            exportButton.type = 'button';
            exportButton.addEventListener('click', exportAllLocationsAndMeetings);
            referenceButton.parentNode.insertBefore(exportButton, referenceButton.nextSibling);
            GM_log("自動匯出按鈕已添加。");
        }
        if (!document.getElementById('select-settings-button')) {
            const settingsButton = document.createElement('button');
            settingsButton.id = 'select-settings-button';
            settingsButton.textContent = '⚙️';
            settingsButton.title = '設定要下載的聚會項目';
            settingsButton.type = 'button';
            settingsButton.style.marginLeft = '5px';
            settingsButton.style.padding = '5px 8px';
            settingsButton.style.verticalAlign = 'middle';
            settingsButton.addEventListener('click', showSelectSettingsModal);
            const mainExportBtn = document.getElementById('export-single-location-items-button');
            if (mainExportBtn) {
                mainExportBtn.parentNode.insertBefore(settingsButton, mainExportBtn.nextSibling);
            } else {
                referenceButton.parentNode.insertBefore(settingsButton, referenceButton.nextSibling);
            }
            GM_log("設定按鈕已添加。");
        }
        if (!document.getElementById('select-settings-modal')) {
            const overlay = document.createElement('div');
            overlay.id = 'select-settings-overlay';
            overlay.style.display = 'none';
            overlay.style.position = 'fixed';
            overlay.style.top = '0';
            overlay.style.left = '0';
            overlay.style.width = '100%';
            overlay.style.height = '100%';
            overlay.style.backgroundColor = 'rgba(0, 0, 0, 0.5)';
            overlay.style.zIndex = '9998';
            overlay.addEventListener('click', hideSelectSettingsModal);
            document.body.appendChild(overlay);
            const modal = document.createElement('div');
            modal.id = 'select-settings-modal';
            modal.style.display = 'none';
            modal.style.position = 'fixed';
            modal.style.top = '50%';
            modal.style.left = '50%';
            modal.style.transform = 'translate(-50%, -50%)';
            modal.style.backgroundColor = 'white';
            modal.style.padding = '20px';
            modal.style.border = '1px solid #ccc';
            modal.style.borderRadius = '5px';
            modal.style.boxShadow = '0 2px 10px rgba(0, 0, 0, 0.1)';
            modal.style.zIndex = '9999';
            modal.style.minWidth = '300px';
            modal.style.maxHeight = '80vh';
            modal.style.overflowY = 'auto';
            modal.innerHTML = ` <h3 style="margin-top: 0; margin-bottom: 15px; border-bottom: 1px solid #eee; padding-bottom: 10px;">設定要下載的聚會項目</h3> <div id="select-settings-list" style="margin-bottom: 20px;"> 載入中... </div> <div style="text-align: right;"> <button id="select-settings-save" type="button" style="padding: 8px 15px; background-color: #28a745; color: white; border: none; border-radius: 3px; cursor: pointer; margin-right: 10px;">儲存設定</button> <button id="select-settings-cancel" type="button" style="padding: 8px 15px; background-color: #6c757d; color: white; border: none; border-radius: 3px; cursor: pointer;">取消</button> </div> `;
            document.body.appendChild(modal);
            document.getElementById('select-settings-save').addEventListener('click', saveSelectSettings);
            document.getElementById('select-settings-cancel').addEventListener('click', hideSelectSettingsModal);
            GM_log("設定視窗HTML已建立。");
        }
        GM_addStyle(` #export-single-location-items-button { margin-left: 10px; padding: 5px 10px; background-color: #17a2b8; color: white; border: 1px solid #117a8b; border-radius: 4px; cursor: pointer; font-size: 14px; vertical-align: middle; } #export-single-location-items-button:disabled { background-color: #cccccc; border-color: #aaaaaa; cursor: not-allowed; color: #666666; } #select-settings-button { background-color: #6c757d; color: white; border: 1px solid #5a6268; border-radius: 4px; } #select-settings-button:hover { background-color: #5a6268; } `);
    }

    // --- Run Script ---
    loadSelectedMeetings(); // Load settings first

    if (document.readyState === 'complete' || document.readyState === 'interactive') {
        createUI();
    } else {
        window.addEventListener('load', createUI);
    }

})();
