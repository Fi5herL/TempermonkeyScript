// ==UserScript==
// @name         自動匯出各會所聚會資料 v2.5 (指定預設項目+狀態)
// @namespace    http://tampermonkey.net/
// @version      2.5
// @description  自動逐一勾選單一會所，匯出選定年份的指定項目資料，並遍歷所有人員狀態，生成匯出紀錄CSV。
// @author       Fisher Li (Modified by Assistant)
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
    // *** NEW: Selector for Member Status Dropdown ***
    const memberStatusSelectSelector = '#member_status';
    const yearFromSelectSelector = '#year_from';
    const monthFromSelectSelector = '#month_from';
    const yearToSelectSelector = '#year_to';
    const monthToSelectSelector = '#month_to';
    const originalExportButtonSelector = '#export';
    const buttonContainerSelector = '#export'; // Container where buttons are added
    const delayBetweenExportsMs = 5000; // Delay between each export click
    const delayAfterCheckboxClickMs = 1000; // Delay after clicking a hall checkbox
    const delayAfterDropdownChangeMs = 500; // *** NEW: Delay after changing dropdowns (meeting/status) ***
    const selectedMeetingsStorageKey = 'chlife_stat_selected_meetings_v2';

    // Define the default selected meetings
    const defaultSelectedMeetings = ["主日", "晨興", "禱告", "小排"];

    // --- Global Variable for Selected Meetings ---
    let selectedMeetings = [];

    // --- Helper Functions (sleep, clickSpecificCheckbox, isCheckboxChecked - unchanged) ---
    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    function clickSpecificCheckbox(hallId) {
        return new Promise(async (resolve, reject) => {
            let escapedHallId;
            try {
                escapedHallId = CSS.escape(hallId);
            } catch (e) {
                if (/^\d/.test(hallId)) {
                    // Handle numeric IDs starting with a digit for CSS selectors
                    escapedHallId = '\\3' + hallId.charAt(0) + ' ' + hallId.slice(1);
                } else {
                    escapedHallId = hallId; // Use original if no escape needed or simple case
                }
            }
            // Try different selectors as before
            const selectorsToTry = [
                `${treeContainerSelector} li#${escapedHallId} > a > ins.jstree-checkbox`, // Standard checkbox within anchor
                `${treeContainerSelector} li#${escapedHallId} > ins.jstree-checkbox`      // Checkbox directly under li (fallback)
            ];

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
                await sleep(delayAfterCheckboxClickMs); // Wait after click
                resolve(true); // Indicate success
            } else {
                GM_log(`Error: Could not find checkbox for Hall ID ${hallId}. Tried selectors: ${selectorsToTry.join(', ')}`);
                reject(new Error(`無法找到 ID 為 ${hallId} 的會所核取方塊。`));
            }
        });
    }


    function isCheckboxChecked(hallId) {
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
        const listItem = document.querySelector(`${treeContainerSelector} li#${escapedHallId}`);

        // Check if the li element itself has the 'jstree-checked' class
        if (listItem && listItem.classList.contains('jstree-checked')) {
            // GM_log(`Checkbox for ${hallId} is checked (li has jstree-checked).`);
            return true;
        }

        // Check if the anchor tag within the li has the 'jstree-checked' class
        const anchor = document.querySelector(`${treeContainerSelector} li#${escapedHallId} > a`);
        if (anchor && anchor.classList.contains('jstree-checked')) {
            // GM_log(`Checkbox for ${hallId} is checked (anchor has jstree-checked).`);
            return true;
        }

        // Fallback: Check the checkbox element's parent anchor or li (less reliable but sometimes needed)
        const checkbox = document.querySelector(`${treeContainerSelector} li#${escapedHallId} ins.jstree-checkbox`);
        if (checkbox) {
             const parentAnchor = checkbox.closest('a');
             if (parentAnchor && parentAnchor.classList.contains('jstree-checked')) {
                 // GM_log(`Checkbox for ${hallId} is checked (checkbox's parent anchor has jstree-checked).`);
                 return true;
             }
             const parentLi = checkbox.closest('li');
             if (parentLi && parentLi.classList.contains('jstree-checked')) {
                // GM_log(`Checkbox for ${hallId} is checked (checkbox's parent li has jstree-checked).`);
                 return true;
             }
        }

        // GM_log(`Checkbox for ${hallId} is NOT checked.`);
        return false;
    }


    // *** MODIFIED: generateLogCsv includes Status Name ***
    function generateLogCsv(logEntries) {
        if (!logEntries || logEntries.length === 0) return null;
        const headers = ['匯出時間', '會所名稱', '項目名稱', '狀態名稱', '原始檔案名稱']; // Added '狀態名稱'
        const rows = logEntries.map(entry => [
            entry.timestamp,
            entry.locationName,
            entry.meetingName,
            entry.statusName, // Added status name field
            entry.originalFilename
        ]);

        const escapeCSV = (field) => {
            if (field === null || field === undefined) return '';
            const stringField = String(field);
            // Escape double quotes and wrap in double quotes if field contains comma, double quote, or newline
            if (stringField.includes(',') || stringField.includes('"') || stringField.includes('\n')) {
                return `"${stringField.replace(/"/g, '""')}"`;
            }
            return stringField;
        };

        const headerString = headers.map(escapeCSV).join(',');
        const rowStrings = rows.map(row => row.map(escapeCSV).join(','));

        const BOM = '\uFEFF'; // UTF-8 Byte Order Mark for Excel compatibility
        return BOM + headerString + '\n' + rowStrings.join('\n');
    }

    // --- downloadCSV function (unchanged) ---
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
         URL.revokeObjectURL(url); // Clean up the object URL
         GM_log(`CSV 檔案 "${filename}" 已觸發下載。`);
    }

    // --- Settings Modal Functions (showSelectSettingsModal, hideSelectSettingsModal, saveSelectSettings - unchanged) ---
    function showSelectSettingsModal() {
        const modal = document.getElementById('select-settings-modal');
        const overlay = document.getElementById('select-settings-overlay');
        const listContainer = document.getElementById('select-settings-list');
        if (!modal || !overlay || !listContainer) {
            GM_log("Error: Settings modal elements not found!");
            return;
        }

        // Get meeting options from the page's dropdown
        const meetingSelect = document.querySelector(meetingSelectSelector);
        if (!meetingSelect) {
            alert("錯誤：找不到聚會項目下拉選單，無法載入設定。");
            return;
        }
        const meetingOptions = Array.from(meetingSelect.options)
                                   .filter(opt => opt.value && opt.text.trim()) // Filter out empty options
                                   .map(opt => opt.text.trim()); // Get only the text names

        listContainer.innerHTML = ''; // Clear previous list

        // Populate modal with checkboxes for each meeting option
        meetingOptions.forEach(name => {
            const listItem = document.createElement('div');
            listItem.style.marginBottom = '5px';

            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `select-setting-${name.replace(/\s+/g, '-')}`; // Create a unique ID
            checkbox.value = name;
            checkbox.checked = selectedMeetings.includes(name); // Check if it's in the saved list
            checkbox.style.marginRight = '8px';

            const label = document.createElement('label');
            label.htmlFor = checkbox.id;
            label.textContent = name;
            label.style.cursor = 'pointer';

            listItem.appendChild(checkbox);
            listItem.appendChild(label);
            listContainer.appendChild(listItem);
        });

        // Show the modal and overlay
        overlay.style.display = 'block';
        modal.style.display = 'block';
        GM_log("顯示'要下載項目'設定視窗。");
    }

    function hideSelectSettingsModal() {
        const modal = document.getElementById('select-settings-modal');
        const overlay = document.getElementById('select-settings-overlay');
        if (modal && overlay) {
            modal.style.display = 'none';
            overlay.style.display = 'none';
            GM_log("隱藏'要下載項目'設定視窗。");
        }
    }

    function saveSelectSettings() {
        const listContainer = document.getElementById('select-settings-list');
        if (!listContainer) return;

        const checkboxes = listContainer.querySelectorAll('input[type="checkbox"]');
        const newSelectedList = [];
        checkboxes.forEach(cb => {
            if (cb.checked) {
                newSelectedList.push(cb.value);
            }
        });

        selectedMeetings = newSelectedList; // Update global variable

        try {
            // Save the updated list to GM storage
            GM_setValue(selectedMeetingsStorageKey, JSON.stringify(selectedMeetings));
            GM_log("已儲存新的'要下載項目'列表:", selectedMeetings);
            // alert("'要下載項目'設定已儲存！"); // Optional: Provide user feedback
        } catch (e) {
            GM_log("儲存'要下載項目'設定時發生錯誤:", e);
            alert("儲存設定失敗，請檢查控制台。");
        }

        hideSelectSettingsModal(); // Close the modal
    }


    // *** MAJOR MODIFICATION: exportAllLocationsAndMeetings includes Status Loop ***
    async function exportAllLocationsAndMeetings() {
        const newButton = document.getElementById('export-single-location-items-button');
        if (newButton) {
            newButton.disabled = true;
            newButton.textContent = '準備中...';
        }

        // Get references to essential page elements
        const treeContainer = document.querySelector(treeContainerSelector);
        const meetingSelect = document.querySelector(meetingSelectSelector);
        const memberStatusSelect = document.querySelector(memberStatusSelectSelector); // *** Get status dropdown
        const yearFromSelect = document.querySelector(yearFromSelectSelector);
        const monthFromSelect = document.querySelector(monthFromSelectSelector);
        const yearToSelect = document.querySelector(yearToSelectSelector);
        const monthToSelect = document.querySelector(monthToSelectSelector);
        const originalExportButton = document.querySelector(originalExportButtonSelector);

        // Check if all required elements are found
        if (!treeContainer || !meetingSelect || !memberStatusSelect || !yearFromSelect || !monthFromSelect || !yearToSelect || !monthToSelect || !originalExportButton) {
            alert("錯誤：找不到必要的頁面元素 (包含 #member_status)，無法開始匯出。");
            if (newButton) {
                newButton.disabled = false;
                newButton.textContent = '匯出各會所項目(逐一+狀態)';
            }
            return;
        }

        const targetYear = yearFromSelect.value;
        if (!targetYear) {
            alert("錯誤：無法讀取目標年份。");
             if (newButton) {
                 newButton.disabled = false;
                 newButton.textContent = '匯出各會所項目(逐一+狀態)';
             }
            return;
        }
        GM_log(`目標年份：${targetYear}`);

        GM_log(`目前設定要下載的項目: ${selectedMeetings.join(', ') || '(無)'}`);
        if (selectedMeetings.length === 0) {
            alert("警告：您尚未在設定中勾選任何要下載的項目。請點擊 ⚙️ 按鈕設定。");
            if (newButton) {
                newButton.disabled = false;
                newButton.textContent = '匯出各會所項目(逐一+狀態)';
            }
            return;
        }

        // Get Hall Information
        const hallListItems = Array.from(treeContainer.querySelectorAll('li.level1[id]'));
        // Fallback if level1 structure isn't found/used
        if (hallListItems.length === 0) {
            GM_log("未找到 li.level1，嘗試查找所有非根節點的 li[id]");
            hallListItems.push(...treeContainer.querySelectorAll('li[id]:not(#\\30 )')); // Exclude root node '0'
        }
        const hallInfoArray = hallListItems.map(li => {
            const id = li.id;
            const anchor = li.querySelector('a');
            // Extract name, removing potential line breaks/tabs
            const name = anchor ? anchor.textContent.replace(/[\n\r\t]/g, '').trim() : `會所 ID ${id}`;
            return { id, name };
        }).filter(hall => hall.id && hall.name); // Ensure valid id and name

        if (hallInfoArray.length === 0) {
            alert("錯誤：在 jsTree 中找不到任何會所項目。");
            if (newButton) {
                newButton.disabled = false;
                newButton.textContent = '匯出各會所項目(逐一+狀態)';
            }
            return;
        }
        GM_log(`找到 ${hallInfoArray.length} 個會所`);

        // Get Meeting Options
        const meetingOptions = Array.from(meetingSelect.options)
                                .filter(opt => opt.value && opt.text.trim()) // Valid value and non-empty text
                                .map(opt => ({ value: opt.value, name: opt.text.trim() }));
        if (meetingOptions.length === 0) {
            alert("錯誤：找不到任何聚會項目。");
             if (newButton) {
                 newButton.disabled = false;
                 newButton.textContent = '匯出各會所項目(逐一+狀態)';
             }
            return;
        }
        GM_log(`找到 ${meetingOptions.length} 個聚會項目`);

        // *** NEW: Get Member Status Options ***
        const statusOptions = Array.from(memberStatusSelect.options)
                                .filter(opt => opt.value && opt.value !== "") // Filter out the default "選擇狀態" (value="")
                                .map(opt => ({ value: opt.value, name: opt.text.trim() }));
        if (statusOptions.length === 0) {
            alert("錯誤：找不到任何有效的人員狀態選項 (已排除 '選擇狀態')。");
             if (newButton) {
                 newButton.disabled = false;
                 newButton.textContent = '匯出各會所項目(逐一+狀態)';
             }
            return;
        }
        GM_log(`找到 ${statusOptions.length} 個狀態選項: ${statusOptions.map(s => s.name).join(', ')}`);

        const exportLogEntries = []; // Array to store log data
        let downloadCounter = 1; // To track original filenames (e.g., export, export (1), etc.)

        try {
            // --- Ensure all halls are initially unchecked ---
            GM_log("確保所有會所都取消勾選...");
            if (newButton) newButton.textContent = '取消全選...';
            let anyInitiallyChecked = false;
            for (const hall of hallInfoArray) {
                if (isCheckboxChecked(hall.id)) {
                    anyInitiallyChecked = true;
                    GM_log(`  - 取消勾選: ${hall.name} (ID: ${hall.id})`);
                    await clickSpecificCheckbox(hall.id);
                }
            }
            // Brief pause after potentially unchecking many items
            if (anyInitiallyChecked) {
                 GM_log("等待取消勾選完成...");
                 await sleep(1500);
            } else {
                 GM_log("所有會所已是取消勾選狀態。");
            }

            // --- Main Loop: Iterate through Halls ---
            for (let i = 0; i < hallInfoArray.length; i++) {
                const currentHall = hallInfoArray[i];
                const currentHallNum = i + 1;
                GM_log(`--- [${currentHallNum}/${hallInfoArray.length}] 處理會所: ${currentHall.name} (ID: ${currentHall.id}) ---`);
                if (newButton) newButton.textContent = `處理中: ${currentHall.name} (${currentHallNum}/${hallInfoArray.length})...`;

                // --- Select the current hall ---
                GM_log(`  勾選會所: ${currentHall.name}`);
                if (!isCheckboxChecked(currentHall.id)) {
                    await clickSpecificCheckbox(currentHall.id);
                    // Verify checkbox is actually checked after click
                    if (!isCheckboxChecked(currentHall.id)) {
                        GM_log(`  警告：勾選 ${currentHall.name} 失敗。跳過此會所。`);
                        continue; // Skip to the next hall if checkbox failed
                    }
                } else {
                     GM_log(`  會所 ${currentHall.name} 已勾選 (可能由前次操作遺留)。`);
                }

                // --- Loop through selected Meeting Types ---
                let exportedMeetingCount = 0;
                for (const meeting of meetingOptions) {
                    // Skip if this meeting is not in the user's selected list
                    if (!selectedMeetings.includes(meeting.name)) {
                        continue;
                    }
                    exportedMeetingCount++;
                    GM_log(`    --- [項目 ${exportedMeetingCount}] ${meeting.name} ---`);

                    // --- *** NEW: Loop through Member Statuses *** ---
                    let exportedStatusCount = 0;
                    for (const status of statusOptions) {
                        exportedStatusCount++;
                        GM_log(`        --- [狀態 ${exportedStatusCount}/${statusOptions.length}] ${status.name} ---`);
                        if (newButton) newButton.textContent = `匯出 ${currentHall.name}: ${meeting.name} - ${status.name}... (${downloadCounter})`;

                        // --- Set Meeting Dropdown ---
                        meetingSelect.value = meeting.value;
                        // --- Set Status Dropdown ---
                        memberStatusSelect.value = status.value;
                        // --- Set Date Range ---
                        yearFromSelect.value = targetYear;
                        yearToSelect.value = targetYear;
                        monthFromSelect.value = '1'; // January
                        monthToSelect.value = '12'; // December

                        // Wait a bit after setting dropdowns before clicking export
                        await sleep(delayAfterDropdownChangeMs);

                        // --- Trigger Export ---
                        GM_log(`          ------> 匯出: [${currentHall.name}] - [${meeting.name}] - [${status.name}] (${targetYear}年) <-------`);
                        originalExportButton.click(); // Click the original export button

                        // --- Log the Export Action ---
                        const timestamp = new Date().toLocaleString('zh-TW', { hour12: false });
                        // Determine the likely original filename based on browser behavior
                        let originalFilename = (downloadCounter === 1) ? 'export' : `export (${downloadCounter - 1})`;

                        exportLogEntries.push({
                            timestamp: timestamp,
                            locationName: currentHall.name,
                            meetingName: meeting.name,
                            statusName: status.name, // Add status name to the log
                            originalFilename: originalFilename
                        });
                        GM_log(`          已記錄: ${currentHall.name}, ${meeting.name}, ${status.name}, ${timestamp}, 檔名: ${originalFilename}`);

                        downloadCounter++; // Increment for the next potential download

                        // Wait before the next export action
                        GM_log(`          等待 ${delayBetweenExportsMs / 1000} 秒...`);
                        await sleep(delayBetweenExportsMs);

                    } // --- End of Status Loop ---

                    // Optional: Reset status dropdown after iterating through all statuses for this meeting? Good practice.
                    memberStatusSelect.value = ""; // Reset to "選擇狀態"
                    await sleep(100); // Tiny delay

                } // --- End of Meeting Loop ---

                // --- Deselect the current hall ---
                GM_log(`  處理完成，取消勾選會所: ${currentHall.name}`);
                if (isCheckboxChecked(currentHall.id)) {
                    await clickSpecificCheckbox(currentHall.id);
                     // Optional: Verify it got unchecked
                     await sleep(200); // Small delay after uncheck
                     if (isCheckboxChecked(currentHall.id)) {
                          GM_log(`  警告：取消勾選 ${currentHall.name} 後似乎仍處於勾選狀態。`);
                     }
                }
                 // Optional: Reset meeting and status again just to be safe before next hall
                 meetingSelect.value = meetingOptions[0]?.value || ""; // Reset to first or empty
                 memberStatusSelect.value = "";
                 await sleep(100);

                GM_log(`--- 會所 ${currentHall.name} 處理完畢 ---`);

            } // --- End of Hall Loop ---

            GM_log("所有會所、選定項目和狀態的匯出請求已完成。");

            // --- Generate and Download Log File ---
            if (exportLogEntries.length > 0) {
                GM_log("準備生成匯出紀錄檔...");
                const logCsvData = generateLogCsv(exportLogEntries);
                if (logCsvData) {
                    const logFilename = `聚會資料_${targetYear}年_含狀態_${new Date().toISOString().slice(0, 10)}.csv`;
                    downloadCSV(logCsvData, logFilename);
                    alert(`匯出完成！\n共觸發 ${downloadCounter - 1} 個檔案下載。\n紀錄檔 "${logFilename}" 已下載。\n請檢查瀏覽器下載設定，可能需要手動允許多個檔案下載。`);
                } else {
                     GM_log("紀錄檔 CSV 字串生成失敗。");
                    alert(`匯出請求已完成，共觸發 ${downloadCounter - 1} 個檔案下載，但紀錄檔生成失敗。`);
                }
            } else {
                 GM_log("沒有觸發任何匯出操作。");
                 alert("未執行任何匯出操作 (可能是因為沒有選取任何聚會項目或找不到會所)。");
            }

        } catch (error) {
            console.error("自動匯出過程中發生錯誤:", error);
            GM_log(`自動匯出錯誤: ${error.message}\n${error.stack}`);
            alert(`匯出過程中發生錯誤：\n${error.message}\n請查看瀏覽器控制台(F12)獲取詳細資訊。`);
        } finally {
            // --- Re-enable the button ---
            if (newButton) {
                newButton.disabled = false;
                newButton.textContent = '匯出各會所項目(逐一+狀態)'; // Update button text
            }
             // Reset status dropdown one last time
             if(memberStatusSelect) memberStatusSelect.value = "";
            GM_log("腳本執行完畢。");
        }
    }

    // --- Load Settings ---
    function loadSelectedMeetings() {
        const storedSettings = GM_getValue(selectedMeetingsStorageKey, null);
        if (storedSettings) {
            try {
                selectedMeetings = JSON.parse(storedSettings);
                GM_log("成功從儲存空間載入'要下載項目':", selectedMeetings);
                // Basic validation
                if (!Array.isArray(selectedMeetings)) {
                    throw new Error("Loaded data is not an array.");
                }
            } catch (e) {
                GM_log("解析儲存的'要下載項目'設定時發生錯誤或格式不符，將使用預設值:", e);
                selectedMeetings = [...defaultSelectedMeetings]; // Use default
                GM_setValue(selectedMeetingsStorageKey, JSON.stringify(selectedMeetings)); // Save default
            }
        } else {
            GM_log("未找到儲存的'要下載項目'設定，使用預設值:", defaultSelectedMeetings);
            selectedMeetings = [...defaultSelectedMeetings]; // Use default
            GM_setValue(selectedMeetingsStorageKey, JSON.stringify(selectedMeetings)); // Store the default
        }
        // Final safety check
        if (!Array.isArray(selectedMeetings)) {
            GM_log("最終檢查發現 selectedMeetings 不是陣列，重設為預設值。");
            selectedMeetings = [...defaultSelectedMeetings];
            GM_setValue(selectedMeetingsStorageKey, JSON.stringify(selectedMeetings));
        }
    }

    // --- Create UI (Modal HTML and Buttons) ---
    function createUI() {
        const referenceButton = document.querySelector(buttonContainerSelector);
        if (!referenceButton) {
            GM_log("找不到原匯出按鈕 (#export)，無法添加UI。");
            return;
        }

        // --- Add Main Export Button ---
        if (!document.getElementById('export-single-location-items-button')) {
            const exportButton = document.createElement('button');
            exportButton.id = 'export-single-location-items-button';
            // *** MODIFIED: Button Text reflects new functionality ***
            exportButton.textContent = '匯出各會所項目(逐一+狀態)';
            exportButton.type = 'button'; // Important for forms
            exportButton.addEventListener('click', exportAllLocationsAndMeetings);

            // Insert after the original export button
            referenceButton.parentNode.insertBefore(exportButton, referenceButton.nextSibling);
            GM_log("自動匯出(含狀態)按鈕已添加。");
        }

        // --- Add Settings Button (Gear Icon) ---
        if (!document.getElementById('select-settings-button')) {
            const settingsButton = document.createElement('button');
            settingsButton.id = 'select-settings-button';
            settingsButton.textContent = '⚙️'; // Gear icon
            settingsButton.title = '設定要下載的聚會項目';
            settingsButton.type = 'button';
            settingsButton.style.marginLeft = '5px';
            settingsButton.style.padding = '5px 8px';
            settingsButton.style.verticalAlign = 'middle'; // Align with other buttons
            settingsButton.addEventListener('click', showSelectSettingsModal);

            // Insert after the new export button
            const mainExportBtn = document.getElementById('export-single-location-items-button');
            if (mainExportBtn) {
                mainExportBtn.parentNode.insertBefore(settingsButton, mainExportBtn.nextSibling);
            } else {
                // Fallback: insert after the original button if the new one failed to add
                referenceButton.parentNode.insertBefore(settingsButton, referenceButton.nextSibling);
            }
            GM_log("設定按鈕已添加。");
        }

        // --- Create Settings Modal HTML (if it doesn't exist) ---
        if (!document.getElementById('select-settings-modal')) {
            // Overlay Div
            const overlay = document.createElement('div');
            overlay.id = 'select-settings-overlay';
            Object.assign(overlay.style, {
                display: 'none', position: 'fixed', top: '0', left: '0',
                width: '100%', height: '100%', backgroundColor: 'rgba(0, 0, 0, 0.5)',
                zIndex: '9998'
            });
            overlay.addEventListener('click', hideSelectSettingsModal); // Click overlay to close
            document.body.appendChild(overlay);

            // Modal Div
            const modal = document.createElement('div');
            modal.id = 'select-settings-modal';
            Object.assign(modal.style, {
                display: 'none', position: 'fixed', top: '50%', left: '50%',
                transform: 'translate(-50%, -50%)', backgroundColor: 'white', padding: '20px',
                border: '1px solid #ccc', borderRadius: '5px', boxShadow: '0 2px 10px rgba(0, 0, 0, 0.1)',
                zIndex: '9999', minWidth: '300px', maxHeight: '80vh', overflowY: 'auto'
            });
            modal.innerHTML = `
                <h3 style="margin-top: 0; margin-bottom: 15px; border-bottom: 1px solid #eee; padding-bottom: 10px;">設定要下載的聚會項目</h3>
                <div id="select-settings-list" style="margin-bottom: 20px;">
                    載入中...
                </div>
                <div style="text-align: right;">
                    <button id="select-settings-save" type="button" style="padding: 8px 15px; background-color: #28a745; color: white; border: none; border-radius: 3px; cursor: pointer; margin-right: 10px;">儲存設定</button>
                    <button id="select-settings-cancel" type="button" style="padding: 8px 15px; background-color: #6c757d; color: white; border: none; border-radius: 3px; cursor: pointer;">取消</button>
                </div>
            `;
            document.body.appendChild(modal);

            // Add event listeners for modal buttons
            document.getElementById('select-settings-save').addEventListener('click', saveSelectSettings);
            document.getElementById('select-settings-cancel').addEventListener('click', hideSelectSettingsModal);
            GM_log("設定視窗HTML已建立。");
        }

        // --- Add CSS Styles ---
        // *** MODIFIED: Updated styles for clarity ***
        GM_addStyle(`
            #export-single-location-items-button {
                margin-left: 10px; /* Spacing from original button */
                padding: 5px 10px;
                background-color: #007bff; /* Blue color */
                color: white;
                border: 1px solid #0056b3;
                border-radius: 4px;
                cursor: pointer;
                font-size: 14px;
                vertical-align: middle; /* Align with other elements */
                transition: background-color 0.2s ease; /* Smooth hover effect */
            }
            #export-single-location-items-button:hover {
                 background-color: #0056b3;
            }
            #export-single-location-items-button:disabled {
                background-color: #cccccc;
                border-color: #aaaaaa;
                cursor: not-allowed;
                color: #666666;
            }
            #select-settings-button { /* Style for the gear button */
                background-color: #6c757d; /* Grey color */
                color: white;
                border: 1px solid #5a6268;
                border-radius: 4px;
                transition: background-color 0.2s ease;
            }
            #select-settings-button:hover {
                background-color: #5a6268;
            }
            /* Optional: Style for the modal checkboxes/labels for better spacing */
            #select-settings-list div {
                display: block; /* Ensure each item is on a new line */
            }
        `);
    }

    // --- Run Script ---
    loadSelectedMeetings(); // Load settings first

    // Create UI elements after the page is loaded or interactive
    if (document.readyState === 'complete' || document.readyState === 'interactive') {
        createUI();
    } else {
        window.addEventListener('load', createUI);
    }

})();
