// ==UserScript==
// @name         上傳CSV自動點名
// @namespace    http://tampermonkey.net/
// @version      2.0
// @description  上傳 CSV 檔案，自動遍歷所有分頁，根據檔案內容勾選點名項目。
// @author       AI Assistant & Fisher Li
// @match        https://www.chlife-stat.org/index.php
// @match        https://www.chlife-stat.org/
// @require      https://code.jquery.com/jquery-3.6.0.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.3.0/papaparse.min.js
// @grant        GM_addStyle
// @grant        unsafeWindow
// ==/UserScript==

/* globals jQuery, $, Papa */

(function() {
    'use strict';

    // --- Selectors (from pagination script & original) ---
    const targetTableSelector = '#roll-call-panel table#table'; // For checking rows within
    const paginationContainerSelector = '#pagination.jPaginate'; // Pagination controls

    // --- 1. Define the API function (Modified to return counts) ---
    /**
     * Updates attendance checkboxes for multiple people based on provided data *for the current page*.
     * @param {Array<object>} attendanceRecords - Full array of objects from CSV.
     * @returns {object} - { updatedCount: number, notFoundNames: string[] } for the current page.
     */
    unsafeWindow.updateAttendanceBatch_v2 = function(attendanceRecords) {
        if (!Array.isArray(attendanceRecords)) {
            console.error('[CSV Helper] Invalid input: attendanceRecords must be an array.');
            return { updatedCount: 0, notFoundNames: [] };
        }
        // console.log('[CSV Helper] Updating current page with batch:', attendanceRecords.length, 'total records');

        let pageUpdatedCount = 0;
        const pageNotFoundNames = [];
        const rowsOnCurrentPage = Array.from(document.querySelectorAll(targetTableSelector + ' tbody tr')); // Get only current page rows

        // Create a quick lookup map for rows on the current page for efficiency
        const rowMap = new Map();
        rowsOnCurrentPage.forEach(row => {
            const cells = Array.from(row.querySelectorAll('td'));
            if (cells.length > 3) {
                const key = `${cells[0]?.textContent.trim()}_${cells[1]?.textContent.trim()}_${cells[2]?.textContent.trim()}_${cells[3]?.textContent.trim()}`;
                if (key && !key.includes('undefined')) { // Ensure key is valid
                    rowMap.set(key, row);
                }
            }
        });

        for (const data of attendanceRecords) {
            if (!data || !data.no || !data.distinction || !data.name || !data.sex) {
                // console.warn('[CSV Helper] Skipping invalid data record in batch:', data);
                continue;
            }

            const { no, distinction, name, sex, sunday, prayer, homeVisitOut, homeVisitIn, smallGroup, morningRevival, gospelVisit, lifeStudy, sundayProphecy } = data;
            const searchKey = `${String(no).trim()}_${String(distinction).trim()}_${String(name).trim()}_${String(sex).trim()}`;

            const targetRow = rowMap.get(searchKey); // Check if this person is on the current page

            if (targetRow) {
                // Found the row on this specific page, proceed with checking
                const cells = Array.from(targetRow.querySelectorAll('td'));
                const attendance = [sunday, prayer, homeVisitOut, homeVisitIn, smallGroup, morningRevival, gospelVisit, lifeStudy, sundayProphecy];
                let rowChanged = false;

                attendance.forEach((attendedValue, index) => {
                    const cellIndex = index + 4;
                    const checkbox = cells[cellIndex]?.querySelector('input[type="checkbox"]');
                    if (checkbox) {
                        const shouldBeChecked = attendedValue === 1;
                        if (checkbox.checked !== shouldBeChecked) {
                            checkbox.click();
                            rowChanged = true;
                        }
                    }
                });

                if (rowChanged) {
                    pageUpdatedCount++;
                }
                 // Remove the row from the map once processed to avoid re-processing if CSV has duplicates
                rowMap.delete(searchKey);

            } else {
                // Person from CSV not found on *this* specific page.
                // We'll collect all names at the end.
                // pageNotFoundNames.push(name); // Don't add here, we check globally later
            }
        }

        // console.log(`[CSV Helper] Page Update Complete: ${pageUpdatedCount} records updated on this page.`);
        // Note: pageNotFoundNames is not accurate here, as it only reflects misses *on this page*.
        // Global "not found" calculation happens after iterating all pages.
        return { updatedCount: pageUpdatedCount, // Return count for this page
                 processedKeys: Array.from(rowMap.keys()) // Return keys processed on this page (optional, maybe useful later)
               };
    };


    // --- 2. Header Mapping ---
    const headerMapping = {
        "主日申言": "sundayProphecy", "生命讀經": "lifeStudy", "家聚會出訪": "homeVisitOut",
        "家聚會受訪": "homeVisitIn", "福音出訪": "gospelVisit", "晨興": "morningRevival",
        "禱告": "prayer", "小排": "smallGroup", "主日": "sunday",
        "NO.": "no", "區別": "distinction", "姓名": "name", "性別": "sex"
    };

    function findHeaderKey(csvHeaders, targetPhrase) {
        if (!csvHeaders) return null;
        const cleanedPhrase = targetPhrase.replace(/\s*0人$/, '').trim();
        return csvHeaders.find(header => header && header.trim().toUpperCase().includes(cleanedPhrase.toUpperCase()));
    }

    // --- 3. Pagination Helper Functions (from Download script) ---

    function getTotalPages(paginationContainer) {
        if (!paginationContainer) return 1;
        const pageElements = paginationContainer.querySelectorAll('ul.jPag-pages li a, ul.jPag-pages li span.jPag-current');
        let maxPage = 0;
        pageElements.forEach(el => {
            const pageNum = parseInt(el.textContent.trim(), 10);
            if (!isNaN(pageNum) && pageNum > maxPage) maxPage = pageNum;
        });
        // If jPaginate structure is different (e.g., only prev/next), might need adjustment
        if (maxPage === 0) {
             // Fallback: Check if there's a 'next' button that's not disabled
             const nextButton = paginationContainer.querySelector('a.jPag-next');
             if (nextButton) return 2; // At least 2 pages if next exists
             return 1; // Default to 1 otherwise
        }
        return maxPage > 0 ? maxPage : 1;
    }

    async function waitForPageLoad(expectedPageNum, paginationContainer) {
        console.log(`[CSV Helper] Waiting for page ${expectedPageNum} to load...`);
        return new Promise((resolve, reject) => {
            const maxWaitTime = 20000; // 20 seconds timeout
            const checkInterval = 300;
            let elapsedTime = 0;

            const intervalId = setInterval(() => {
                elapsedTime += checkInterval;
                // Re-query the container and current page element each time
                const currentPaginationContainer = document.querySelector(paginationContainerSelector);
                const activePageElement = currentPaginationContainer?.querySelector('span.jPag-current');
                let currentPageNum = -1;

                if (activePageElement) {
                    currentPageNum = parseInt(activePageElement.textContent.trim());
                } else {
                    // If no 'current' span, maybe it's page 1 and only 'next' exists?
                    // Or maybe the page structure changed. Assume page 1 if no span and expected is 1.
                    if (expectedPageNum === 1 && !currentPaginationContainer?.querySelector('a.jPag-prev')) {
                       currentPageNum = 1;
                    }
                   // console.warn("waitForPageLoad: Could not find span.jPag-current");
                }

                if (currentPageNum === expectedPageNum) {
                    console.log(`[CSV Helper] Page ${expectedPageNum} loaded.`);
                    clearInterval(intervalId);
                    // Short delay to allow table rendering after pagination state update
                    setTimeout(resolve, 350);
                } else if (elapsedTime >= maxWaitTime) {
                    clearInterval(intervalId);
                    console.error(`[CSV Helper] Timeout waiting for page ${expectedPageNum}. Current page reported as: ${currentPageNum}`);
                    reject(new Error(`Timeout waiting for page ${expectedPageNum}`));
                }
            }, checkInterval);
        });
    }


    // --- 4. Create Floating Buttons ---
    const buttonContainer = $('<div id="csv-helper-buttons"></div>');
    const uploadButton = $('<button id="upload-csv-btn" title="上傳CSV檔，將遍歷所有分頁進行勾選">上傳 CSV 勾選 (多頁)</button>');
    // const apiButton = $('<button id="call-api-btn" title="使用內建範例資料勾選當前頁面">勾選範例資料 (當前頁)</button>'); //測試範例
    const fileInput = $('<input type="file" id="csv-file-input" accept=".csv" style="display: none;">');

    //buttonContainer.append(uploadButton).append(apiButton).append(fileInput); //測試範例
    buttonContainer.append(uploadButton).append(fileInput);
    $('body').append(buttonContainer);

    // --- 5. Add Styles ---
    GM_addStyle(`
         #csv-helper-buttons {
            position: fixed;
            bottom: 15px;
            right: 15px;
            z-index: 9999;
            background-color: rgba(250, 250, 250, 0.95);
            border: 1px solid #ddd;
            padding: 8px;
            border-radius: 6px;
            box-shadow: 0 3px 10px rgba(0,0,0,0.1);
            display: flex;
            flex-direction: column;
            gap: 6px;
            width: 150px;
        }

        #csv-helper-buttons button {
            padding: 6px 10px;
            cursor: pointer;
            border: 1px solid #ccc;
            background-color: #f5f5f5;
            border-radius: 4px;
            font-size: 13px;
            transition: all 0.2s ease;
            width: 100%;
            text-align: center;
            color: #444;
        }

        #csv-helper-buttons button:hover {
            background-color: #e8e8e8;
            border-color: #aaa;
            transform: translateY(-1px);
            box-shadow: 0 2px 3px rgba(0,0,0,0.05);
        }

        #csv-helper-buttons button:disabled {
            background-color: #f0f0f0;
            color: #aaa;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }
    `);

    // --- 6. Event Handlers ---
    uploadButton.on('click', function() {
        fileInput.click();
    });

    fileInput.on('change', function(event) {
        const file = event.target.files[0];
        if (file) {
            console.log('[CSV Helper] File selected:', file.name);
            processCsvAcrossPages(file); // Call the multi-page processing function
            $(this).val(''); // Clear input
        }
    });

//     apiButton.on('click', function() { // Sample API button only acts on the current page 測試範例
//         console.log('[CSV Helper] Calling updateAttendanceBatch_v2 with sample data on CURRENT page...');
//         const sampleData = [
//              // Use names likely present in your sample CSV or actual data for testing
//             { no: "1", distinction: "2-4小區", name: "任萬德", sex: "男", sunday: 1, prayer: 1, homeVisitOut: 0, homeVisitIn: 0, smallGroup: 0, morningRevival: 1, gospelVisit: 0, lifeStudy: 0, sundayProphecy: 0},
//             { no: "1", distinction: "2-4小區", name: "李福軒", sex: "男", sunday: 0, prayer: 1, homeVisitOut: 0, homeVisitIn: 0, smallGroup: 0, morningRevival: 0, gospelVisit: 0, lifeStudy: 0, sundayProphecy: 0},
//             { no: "1", distinction: "2-4小區", name: "潘仰信", sex: "男", sunday: 1, prayer: 0, homeVisitOut: 0, homeVisitIn: 0, smallGroup: 0, morningRevival: 0, gospelVisit: 0, lifeStudy: 0, sundayProphecy: 0},
//             // Add a non-existent one for testing not-found on current page
//             { no: "99", distinction: "X區", name: "測試查無此人", sex: "男", sunday: 1, prayer: 1, homeVisitOut: 1, homeVisitIn: 1, smallGroup: 1, morningRevival: 1, gospelVisit: 1, lifeStudy: 1, sundayProphecy: 1}
//         ];

//         if (typeof unsafeWindow.updateAttendanceBatch_v2 === 'function') {
//             try {
//                  const result = unsafeWindow.updateAttendanceBatch_v2(sampleData);
//                  alert(`範例資料勾選完成 (僅當前頁)。\n本次更新 ${result.updatedCount} 筆記錄。`);
//             } catch (error) {
//                  console.error('[CSV Helper] Error executing sample data update:', error);
//                  alert('執行範例資料勾選時發生錯誤 (當前頁)，詳情請查看控制台。');
//             }
//         } else {
//             alert('錯誤: updateAttendanceBatch_v2 函數未找到。');
//         }
//     });


    // --- 7. Core Multi-Page Processing Logic ---
        // --- 7. Core Multi-Page Processing Logic ---
    async function processCsvAcrossPages(file) {
        uploadButton.prop('disabled', true).text('讀取 CSV...');
        let parsedData = [];
        // Declare csvKeys and foundKeys here, in the function's main scope
        let csvKeys = new Set();
        let foundKeys = new Set();

        // Step 1: Parse the CSV
        try {
            parsedData = await parseCsvFilePromise(file);
            if (!parsedData || parsedData.length === 0) {
                alert('CSV 文件為空或未能解析出有效數據行。');
                uploadButton.prop('disabled', false).text('上傳 CSV 勾選 (多頁)');
                return;
            }
            console.log(`[CSV Helper] CSV Parsed. ${parsedData.length} records found.`);
             // Initialize csvKeys *after* parsedData is successfully loaded
             csvKeys = new Set(parsedData.map(data => {
                 // Ensure data exists and has necessary props before creating key
                 if (data && data.no !== undefined && data.distinction !== undefined && data.name !== undefined && data.sex !== undefined) {
                    return `${String(data.no).trim()}_${String(data.distinction).trim()}_${String(data.name).trim()}_${String(data.sex).trim()}`;
                 }
                 return null; // Return null for invalid data rows
             }).filter(key => key !== null)); // Filter out null keys

             // foundKeys is already initialized as an empty Set above

        } catch (error) {
            alert(`CSV 解析失敗: ${error.message}`);
            console.error("[CSV Helper] CSV Parsing Failed:", error);
            uploadButton.prop('disabled', false).text('上傳 CSV 勾選 (多頁)');
            return;
        }

        // Step 2: Initialize Pagination
        let currentPage = 1;
        let totalPages = 1;
        let totalUpdatedCount = 0;
        const paginationContainerInitial = document.querySelector(paginationContainerSelector);

        // This try block covers Steps 2, 3, and 4
        try {
             if (paginationContainerInitial) {
                 totalPages = getTotalPages(paginationContainerInitial);
                 console.log(`[CSV Helper] Detected ${totalPages} pages.`);
             } else {
                 console.log("[CSV Helper] No pagination found, assuming 1 page.");
             }
             uploadButton.text(`處理中 (頁 1/${totalPages})...`);

             // Step 3: Loop through pages
             while (currentPage <= totalPages) {
                 console.log(`[CSV Helper] --- Processing Page ${currentPage} ---`);
                 uploadButton.text(`處理中 (頁 ${currentPage}/${totalPages})...`);

                 if (currentPage > 1) {
                      await new Promise(resolve => setTimeout(resolve, 150));
                 }

                 // Step 3a: Apply checks to the current page
                 if (typeof unsafeWindow.updateAttendanceBatch_v2 === 'function') {
                     const pageResult = unsafeWindow.updateAttendanceBatch_v2(parsedData);
                     totalUpdatedCount += pageResult.updatedCount;
                     console.log(`[CSV Helper] Page ${currentPage}: ${pageResult.updatedCount} records updated.`);

                      // Track which CSV keys were found on this page
                      // Now csvKeys and foundKeys are accessible here
                      const rowsOnCurrentPage = Array.from(document.querySelectorAll(targetTableSelector + ' tbody tr'));
                      rowsOnCurrentPage.forEach(row => {
                         const cells = Array.from(row.querySelectorAll('td'));
                         if (cells.length > 3) {
                             const key = `${cells[0]?.textContent.trim()}_${cells[1]?.textContent.trim()}_${cells[2]?.textContent.trim()}_${cells[3]?.textContent.trim()}`;
                              // Check against the correctly scoped csvKeys Set
                              if (key && csvKeys.has(key)) {
                                  foundKeys.add(key); // Add to the correctly scoped foundKeys Set
                              }
                         }
                      });

                 } else {
                     throw new Error("updateAttendanceBatch_v2 function is not defined.");
                 }

                 // Step 3b: Navigate to the next page if needed
                 if (currentPage >= totalPages) {
                     console.log("[CSV Helper] Reached the last page.");
                     break;
                 }

                 const paginationContainerCurrent = document.querySelector(paginationContainerSelector);
                 const nextPageButton = paginationContainerCurrent?.querySelector('a.jPag-next');

                 if (nextPageButton && !nextPageButton.classList.contains('jPag-disabled')) {
                     const nextPageNum = currentPage + 1;
                     console.log(`[CSV Helper] Clicking 'Next' to go to page ${nextPageNum}`);
                     nextPageButton.click();
                     await waitForPageLoad(nextPageNum, paginationContainerCurrent);
                     currentPage = nextPageNum;
                 } else {
                     console.warn(`[CSV Helper] Could not find 'Next' button or it's disabled on page ${currentPage}. Stopping.`);
                     // (Optional: Add more detailed checks as before)
                     break;
                 }
             } // End while loop

             // Step 4: Final Report
             console.log(`[CSV Helper] --- Multi-Page Processing Complete ---`);
             console.log(`[CSV Helper] Total records updated across all pages: ${totalUpdatedCount}`);

              // Calculate not found names accurately (using the correctly scoped foundKeys)
             const notFoundNamesList = [];
             parsedData.forEach(data => {
                 // Reconstruct the key consistently
                  if (data && data.no !== undefined && data.distinction !== undefined && data.name !== undefined && data.sex !== undefined) {
                    const key = `${String(data.no).trim()}_${String(data.distinction).trim()}_${String(data.name).trim()}_${String(data.sex).trim()}`;
                    if (!foundKeys.has(key)) {
                        notFoundNamesList.push(data.name || `(Row with NO ${data.no}, Dist ${data.distinction})`);
                    }
                 } else {
                     // Handle cases where parsed data might be incomplete (though filtered earlier)
                     console.warn("[CSV Helper] Incomplete record found during final check:", data);
                 }
             });

             // ... (rest of the final message generation using notFoundNamesList) ...
              let finalMessage = `跨頁勾選完成！\n共 ${totalPages} 個頁面。\n總共更新了 ${totalUpdatedCount} 筆記錄。`;
              if (notFoundNamesList.length > 0) {
                  finalMessage += `\n\n在所有頁面中，CSV 內的以下 ${notFoundNamesList.length} 筆資料似乎未找到對應的網頁行：\n${notFoundNamesList.slice(0, 10).join(', ')}`; // Show first 10
                  if (notFoundNamesList.length > 10) finalMessage += '... 等等。';
                  finalMessage += '\n(請檢查 CSV 中的 NO., 區別, 姓名, 性別 是否與網頁完全相符)';
                  console.warn(`[CSV Helper] Names/Records from CSV not found on any page (${notFoundNamesList.length}):`, notFoundNamesList);
              } else {
                  finalMessage += "\n所有 CSV 內的記錄似乎都在網頁中找到了對應行。";
              }
              // alert(finalMessage); //執行最後彈出提醒視窗


        } catch (error) {
            console.error("[CSV Helper] Error during multi-page processing:", error); // Log the detailed error
            alert(`處理分頁時發生錯誤：\n${error.message}\n請檢查控制台獲取詳細資訊。`);
        } finally {
             // Go back to page 1 (Optional)
             try {
                 // Use a more robust selector for page 1 link if 'first' doesn't exist
                 const firstPageLink = document.querySelector('#pagination.jPaginate li a[page="1"]') || document.querySelector('#pagination.jPaginate a.jPag-first');
                  if (firstPageLink && currentPage !== 1) { // Check if currentPage could be determined
                      console.log("[CSV Helper] Returning to page 1...");
                      firstPageLink.click();
                      // Wait briefly for navigation to initiate
                      await new Promise(resolve => setTimeout(resolve, 500));
                      // Optionally wait for page 1 load confirmation, but might be overkill here
                      // await waitForPageLoad(1, document.querySelector(paginationContainerSelector));
                  }
             } catch (navError) {
                 console.warn("[CSV Helper] Could not automatically return to page 1:", navError);
             }
             uploadButton.prop('disabled', false).text('上傳 CSV 勾選 (多頁)');
        }
    }


    // --- 8. CSV Parsing Logic (Promisified) ---
    function parseCsvFilePromise(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = function(e) {
                let csvContent = e.target.result;
                if (csvContent.startsWith('\uFEFF')) {
                    csvContent = csvContent.substring(1);
                }

                Papa.parse(csvContent, {
                    header: true,
                    skipEmptyLines: true,
                    complete: function(results) {
                        if (results.errors.length > 0) {
                            console.error("[CSV Helper] CSV Parsing Errors:", results.errors);
                            return reject(new Error("CSV 解析錯誤: " + results.errors[0].message));
                        }
                        if (!results.data || results.data.length === 0) {
                             return resolve([]); // Resolve with empty array for empty data
                        }

                        const csvHeaders = results.meta.fields;
                        const attendanceData = results.data.map((row, rowIndex) => {
                            const mappedRow = {};
                            let hasEssentialId = true; // Simple check for name presence

                            for (const [chinesePhrase, englishKey] of Object.entries(headerMapping)) {
                                const actualHeader = findHeaderKey(csvHeaders, chinesePhrase);
                                if (actualHeader) {
                                    const rawValue = row[actualHeader]?.trim();
                                    if (["sunday", "prayer", "homeVisitOut", "homeVisitIn", "smallGroup", "morningRevival", "gospelVisit", "lifeStudy", "sundayProphecy"].includes(englishKey)) {
                                        mappedRow[englishKey] = parseInt(rawValue || '0', 10) === 1 ? 1 : 0;
                                    } else {
                                        mappedRow[englishKey] = rawValue || '';
                                    }
                                    if (englishKey === 'name' && !mappedRow[englishKey]) {
                                        hasEssentialId = false;
                                    }
                                } else {
                                    // Header not found in CSV
                                     mappedRow[englishKey] = ["no", "distinction", "name", "sex"].includes(englishKey) ? '' : 0;
                                     if (englishKey === 'name') hasEssentialId = false; // Name header must exist
                                }
                            }

                             // Skip row if name is missing entirely
                            if (!hasEssentialId) {
                                console.warn(`[CSV Helper] Skipping CSV row ${rowIndex + 2} due to missing name.`);
                                return null;
                            }
                            return mappedRow;

                        }).filter(row => row !== null);

                        resolve(attendanceData);
                    },
                    error: function(error) {
                        console.error("[CSV Helper] CSV Parsing Failed:", error);
                        reject(new Error("CSV 文件解析失敗: " + error.message));
                    }
                });
            };

            reader.onerror = function() {
                console.error("[CSV Helper] File reading error:", reader.error);
                reject(new Error('無法讀取文件。'));
            };

            reader.readAsText(file, 'UTF-8'); // Assume UTF-8
        });
    }


    console.log('[CSV Helper] Multi-Page Script (v2.0) loaded.');

})();
