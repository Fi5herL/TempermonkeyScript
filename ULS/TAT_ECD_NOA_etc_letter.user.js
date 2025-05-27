// ==UserScript==
// @name         TAT ECD NOA Letter 浮動按鈕
// @name:zh-CN   TAT ECD NOA Letter 浮動按鈕
// @namespace    http://tampermonkey.net/
// @version      1.1
// @description  在網頁上添加一個浮動按鈕"ECD Letter"，點擊後根據當前頁面的ProjectID跳轉到指定URL。
// @description:zh-CN 在網頁上添加一個浮動按鈕"ECD Letter"，點擊後根據當前頁面的ProjectID跳轉到指定URL。
// @author       Your Name (可替換成你的名字)
// @match       https://portal.ul.com/Project/Details*
// @grant        GM_addStyle
// @grant        GM_openInTab
// ==/UserScript==

(function() {
    'use strict';

    // --- !!! 請修改這裡 !!! ---
    const targetAnchorSelector = '.section-crumbs-li a'; // 替換成您目標<a>標籤的CSS選擇器
    // --- 請修改這裡結束 ---

    const ARROW_CHAR = ' ›';
    const BASE_REPORT_URL = 'https://epic.ul.com/Report';

    function getFormattedDate() {
        const today = new Date();
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    function createReportButton(buttonConfig, bottomPosition) {
        const button = document.createElement('button');
        button.id = buttonConfig.id;
        button.className = 'reportFloatingButton';
        button.textContent = buttonConfig.text + ARROW_CHAR;
        document.body.appendChild(button);
        button.style.bottom = bottomPosition;

        button.addEventListener('click', function() {
            let hrefValueFromAnchor = null;
            const targetAnchorElement = document.querySelector(targetAnchorSelector);

            if (targetAnchorElement) {
                hrefValueFromAnchor = targetAnchorElement.getAttribute('href');
                if (hrefValueFromAnchor === null || hrefValueFromAnchor.trim() === "") {
                    alert(`錯誤：目標<a>標籤 (選擇器: ${targetAnchorSelector}) 的 href 屬性為空或不存在。`);
                    return;
                }
            } else {
                alert(`錯誤：未在頁面上找到目標<a>標籤。\n請檢查CSS選擇器 '${targetAnchorSelector}' 是否正確。`);
                return;
            }

            const reportParams = new URLSearchParams();
            reportParams.append('TemplateUNID', buttonConfig.params.templateUNID);
            reportParams.append('SelectedOutputType', buttonConfig.params.selectedOutputType);
            reportParams.append('ProjectID', encodeURIComponent(hrefValueFromAnchor));
            reportParams.append('isWorkbench', 'False');

            if (buttonConfig.params.addFRDate) {
                reportParams.append('FRDate', getFormattedDate());
            }

            const targetUrl = `${BASE_REPORT_URL}?${reportParams.toString()}`;

            console.log(`按鈕 "${buttonConfig.text}" 被點擊。`);
            console.log("將要打開的目標網址: " + targetUrl);
            console.log("用於 ProjectID 的原始 href 值: ", hrefValueFromAnchor);
            if (buttonConfig.params.addFRDate) {
                console.log("FRDate: ", getFormattedDate());
            }
            GM_openInTab(targetUrl, { active: true });
        });
        return button;
    }

    const buttonsData = [
        {
            id: 'ecdLetterFloatingButton',
            text: 'ECD Letter',
            params: {
                templateUNID: 'AHL ECD Letter',
                selectedOutputType: '.eml',
                addFRDate: false
            }
        },
        {
            id: 'tatLetterFloatingButton',
            text: 'TAT Letter',
            params: {
                templateUNID: 'AHL TAT Letter',
                selectedOutputType: '.eml',
                addFRDate: false
            }
        },
        {
            id: 'noaLetterFloatingButton',
            text: 'NOA Letter',
            params: {
                templateUNID: 'Notice of Authorization or Completion Letter',
                selectedOutputType: '.default',
                addFRDate: true,
                // UL_check 不知道怎麼打勾
            }
        }
    ];

    const initialBottom = 20;
    const buttonHeightEstimate = 38;
    const buttonSpacing = 5;

    buttonsData.forEach((config, index) => {
        const bottomPosition = initialBottom + (index * (buttonHeightEstimate + buttonSpacing));
        createReportButton(config, `${bottomPosition}px`);
    });

    GM_addStyle(`
        .reportFloatingButton {
            position: fixed;
            right: 20px;
            z-index: 9999;
            display: block;
            font-weight: bold;
            font-size: 1.1em;
            line-height: 1.4em;
            padding: 5px 10px 5px 8px;
            border: 1px solid #aeaeae;
            color: #333333;
            text-decoration: none;
            background-color: #f0f0f0;
            cursor: pointer;
            outline: 0 !important;
            border-radius: 3px;
            box-shadow: 1px 1px 3px rgba(0,0,0,0.1);
            transition: background-color 0.2s ease, border-color 0.2s ease;
            min-width: 160px;
            text-align: left;
        }

        .reportFloatingButton:hover {
            background-color: #e0e0e0;
            border-color: #999999;
            color: #000000;
        }
    `);

    if (!document.querySelector(targetAnchorSelector)) {
        console.warn(`腳本提示: 頁面加載時未找到符合選擇器 '${targetAnchorSelector}' 的元素。請確保此選擇器正確無誤。`);
    }
})();
