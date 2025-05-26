// ==UserScript==
// @name         Inactive Letter Generator (Ultra-Compact, Styled, Click to Copy)
// @namespace    http://tampermonkey.net/
// @version      3.2
// @description  Automated generation of Inactive Letter emails. Ultra-compact, styled, click to copy.
// @author       Your Name
// @match        https://portal.ul.com/Project/Details/*
// @grant        GM_addStyle
// ==/UserScript==

(function() {
    'use strict';

    // --- Helper Functions (保持不變) ---
    function extractFieldByLabel(labelText) {
        const xpath = `//div[@class='display-label-row' and normalize-space(.)='${labelText}']/following-sibling::div[@class='display-field-row'][1]`;
        const fieldElement = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        return fieldElement ? fieldElement.textContent.trim() : '';
    }

    function extractOdrNum() {
        const xpath = "//dt[normalize-space(.)='Order Number:']/following-sibling::dd[1]//span";
        const element = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        return element ? element.textContent.trim() : '';
    }

    // ... (其他 Helper Functions 保持不變) ...
    function extractRecipientEmail() {
        const labelText = "Customer Company Contact";
        const xpath = `//div[@class='display-label-row row-border-bottom' and normalize-space(.)='${labelText}']/following-sibling::div[contains(normalize-space(.), 'Email:')][1]`;
        const fieldElement = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        if (fieldElement) {
            const text = fieldElement.textContent;
            const emailMatch = text.match(/Email:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i);
            if (emailMatch && emailMatch[1]) {
                return emailMatch[1].trim();
            }
        }
        return '';
    }

    function extractProjectHandlerEmail() {
        const xpath = "//dt[normalize-space(.)='Project Handler:']/following-sibling::dd[1]";
        const element = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        return element ? (element.getAttribute('title') || element.textContent.trim()) : '';
    }

    function formatToROCYearMonth(dateStringMDY) {
        if (!dateStringMDY) return "未知日期";
        const parts = dateStringMDY.split('/');
        if (parts.length === 3) {
            const month = parseInt(parts[0], 10);
            const day = parseInt(parts[1], 10);
            const year = parseInt(parts[2], 10);
            if (!isNaN(month) && !isNaN(day) && !isNaN(year)) {
                const rocYear = year - 1911;
                return `中華民國${rocYear}年${month}月`;
            }
        }
        return "未知日期 (格式錯誤)";
    }

    function getFutureROCDate(daysOffset) {
        const futureDate = new Date();
        futureDate.setDate(futureDate.getDate() + daysOffset);
        const rocYear = futureDate.getFullYear() - 1911;
        const month = futureDate.getMonth() + 1;
        const day = futureDate.getDate();
        return `中華民國${rocYear}年${month}月${day}日`;
    }


    // --- Email Templates & Button Labels (保持不變) ---
    const buttonLabels = {
        'notice': 'Notice',
        'inactive1': '1',
        'inactive2': '2',
        'inactive3': '3',
        'final': 'Final'
    };

    const emailTemplates = [
        // ... (郵件模板內容保持不變, 為了簡潔此處省略)
        {
            id: 'notice',
            name: 'Notice Inactive Letter',
            titleTemplate: "Project Inactive Letter–Project #PjNum#/關於UL項目#PjNum#暫停通知書",
            contentTemplate: `尊敬的客戶：

感謝您及貴公司對UL服務的信任與支持。

關於貴公司xxx年xx月提交的UL認證項目，認證項目編號 #PjNum# ，服務訂單編號 #OdrNum# ，認證申請描述 #PjScope#，我們不得不書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。即日起，項目進度變更為暫停狀態。

為了繼續推進項目進程，我們需要貴公司盡快提供如下信息
#Project Hold Reason#

請知悉，在UL收到完整並正確的如上信息後，您的項目方可重新啟動。任何不明確之處，歡迎您隨時與我們聯繫，我們也將與貴公司保持積極的互動以盡快重啟您的認證項目。

順祝
商祺

Project Handler /Email:
Engineer Manager /Email:`
        },
        {
            id: 'inactive1',
            name: 'Inactive Letter 1',
            titleTemplate: "The 1st project inactive follow up letter–Project #PjNum#/第一次項目暫停跟進通知書",
            contentTemplate: `尊敬的客戶：

感謝您及貴公司對UL服務的信任與支持。

關於貴公司xxx年xx月提交的UL認證項目，認證項目編號#PjNum#，服務訂單編號#OdrNum#，認證申請描述#PjScope#，我們不得不書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。即日起，項目進度變更為暫停狀態。

為了繼續推進項目進程，我們需要貴公司盡快提供如下信息
#Project Hold Reason#

到目前為止，我們尚未從貴公司收到完整並正確的上述信息，項目仍處於暫停狀態。

請知悉，在UL收到完整並正確的如上信息後，您的項目方可重新啟動。任何不明確之處，歡迎您隨時與我們聯繫，我們也將與貴公司保持積極的互動以盡快重啟您的認證項目。

順祝
商祺

Project Handler /Email:`
        },
        {
            id: 'inactive2',
            name: 'Inactive Letter 2',
            titleTemplate: "The 2nd project inactive follow up letter–Project #PjNum#/第二次項目暫停跟進通知書",
            contentTemplate: `尊敬的客戶：

感謝您及貴公司對UL服務的信任與支持。

關於貴公司xxx年xx月提交的UL認證項目，認證項目編號#PjNum#，服務訂單編號#OdrNum#，認證申請描述#PjScope#，我們不得不書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。即日起，項目進度變更為暫停狀態。

為了繼續推進項目進程，我們需要貴公司盡快提供如下信息
#Project Hold Reason#

一個月前，我們發出第一次項目暫停跟進通知書，但到目前為止，我們仍未從貴公司收到完整並正確的上述信息，項目仍處於暫停狀態。

請知悉，在UL收到完整並正確的如上信息後，您的項目方可重新啟動。任何不明確之處，歡迎您隨時與我們聯繫，我們也將與貴公司保持積極的互動以盡快重啟您的認證項目。

順祝
商祺

Project Handler /Email:`
        },
        {
            id: 'inactive3',
            name: 'Inactive Letter 3',
            titleTemplate: "The 3nd project inactive follow up letter–Project #PjNum#/第三次項目暫停跟進通知書",
            contentTemplate: `尊敬的客戶：

感謝您及貴公司對UL服務的信任與支持。

關於貴公司xxx年xx月提交的UL認證項目，認證項目編號#PjNum#，服務訂單編號#OdrNum#，認證申請描述#PjScope#，我們不得不書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。即日起，項目進度變更為暫停狀態。

為了繼續推進項目進程，我們需要貴公司盡快提供如下信息
#Project Hold Reason#

二個月前，我們發出第一次項目暫停跟進通知書，並且在一個月前，向貴司發出第二次項目暫停跟進通知書，但是到目前為止，我們仍未從貴公司收到完整併正確的上述信息，項目依舊處於暫停狀態。

請知悉，在UL收到完整並正確的如上信息後，您的項目方可重新啟動。任何不明確之處，歡迎您隨時與我們聯繫，我們也將與貴公司保持積極的互動以盡快重啟您的認證項目。

順祝
商祺

Project Handler /Email:
Field Sales /Email:`
        },
        {
            id: 'final',
            name: 'Inactive Letter Final',
            titleTemplate: "The final notice before project close by letter – Project #PjNum#/項目終止前最後提醒",
            contentTemplate: `尊敬的客戶：
感謝您及貴公司對UL服務的信任與支持。
關於貴公司xxx年xx月提交的UL認證項目，認證項目編號#PjNum#，服務訂單編號： #OdrNum#, 認證申請描述 #PjScope#。約四個月前，我們曾書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。項目進度變更為暫停狀態。
為了繼續推進項目進程，我們需要貴公司盡快提供如下信息，或就該等信息和材料的提交提出明確的時間表。
#Project Hold Reason#

由於未能收到有效反饋，約三個月前，我們發出第一次項目暫停跟進通知書，並且在隨後二個月，接連向貴司發出第二次以及第三次項目暫停跟進通知書。但是到目前為止，我們仍未從貴公司收到完整併正確的上述信息，項目始終處於暫停狀態。

根據過去四個月的項目進展狀況，我們在此最後一次向您發出項目提醒函，請您務必引起重視，如果該項目在未來兩週內，也就是#DeadlineDate#前您仍不能提供完整併正確的上述必要信息/樣品，我們不得不遺憾的通知您，我們將終止貴公司【#OdrNum#】號服務訂單及其項下之#PjNum#認證項目。項目終止後，我們將就我司已經提供的服務向您收取相應的費用；項目採用預付款方式支付的，我們將在扣除必要費用後，退還您的剩餘款項。如以上服務需求在未來需要再次啟動，您可以向我們索取一份新的正式報價（報價有效期為三個月）。我們將重新核定貴公司的服務需求，並向您發出新的報價。

順祝
商祺

Field Sales /Email:                                       Project Handler /Email:
Sales Manager /Email:                                Engineer Manager /Email:`
        }
    ];


    // --- Data Gathering (保持不變) ---
    function gatherAllPageData() {
        const pjNum = extractFieldByLabel('Oracle Project Number');
        const dateBookedRaw = extractFieldByLabel('Date Booked');
        const pjScope = extractFieldByLabel('Project Scope');
        const projectHoldReason = extractFieldByLabel('Project Hold Reason');
        const odrNum = extractOdrNum();
        const rocDateBooked = formatToROCYearMonth(dateBookedRaw);
        const deadlineDate = getFutureROCDate(14);
        const projectHandlerEmail = extractProjectHandlerEmail();

        return {
            pjNum: pjNum || "N/A",
            odrNum: odrNum || "N/A",
            rocDateBooked: rocDateBooked || "未知日期",
            pjScope: pjScope || "N/A",
            projectHoldReason: projectHoldReason || "N/A",
            deadlineDate: deadlineDate,
            projectHandlerEmail: projectHandlerEmail || "N/A"
        };
    }

    // --- UI Creation and Logic ---
    function createUI() {
        const panel = document.createElement('div');
        panel.id = 'inactiveLetterPanel';
        panel.innerHTML = `
            <div id="ilg-header">
                <h3 id="ilg-header-title">Letters</h3>
                <span id="ilg-close-btn">×</span>
            </div>
            <div id="ilg-buttons-container"></div>
            <div id="ilg-output-area" style="display:none;">
                <h4 style="margin-top:0;">Generated Email:</h4>
                <label for="ilg-title-output">Title (click to copy):</label>
                <textarea id="ilg-title-output" rows="2" style="width:100%; box-sizing: border-box; cursor: pointer;" readonly title="Click to copy title"></textarea>

                <label for="ilg-content-output">Content (click to copy):</label>
                <textarea id="ilg-content-output" rows="5" style="width:100%; box-sizing: border-box; cursor: pointer;" readonly title="Click to copy content"></textarea>
            </div>
             <div id="ilg-global-status" style="text-align: center; min-height: 1em;"></div>
        `;
        document.body.appendChild(panel);

        const buttonsContainer = panel.querySelector('#ilg-buttons-container');
        emailTemplates.forEach(template => {
            const button = document.createElement('button');
            // MODIFIED: No arrow in button text for horizontal layout, to save space
            button.textContent = buttonLabels[template.id] || template.id;
            button.className = 'ilg-action-button';
            button.title = `Generate ${template.name}`; // Tooltip still shows full name
            button.addEventListener('click', () => generateAndDisplay(template.id));
            buttonsContainer.appendChild(button);
        });

        document.getElementById('ilg-close-btn').addEventListener('click', () => {
            panel.style.display = 'none';
            document.getElementById('ilg-output-area').style.display = 'none';
            document.getElementById('ilg-global-status').textContent = '';
        });

        const titleTextarea = document.getElementById('ilg-title-output');
        const contentTextarea = document.getElementById('ilg-content-output');

        function copyTextareaContent(textareaElement, fieldName) {
            if (!textareaElement.value) {
                showGlobalStatus(`${fieldName} is empty. Nothing to copy.`, true, 2000);
                return;
            }
            navigator.clipboard.writeText(textareaElement.value)
                .then(() => {
                    showGlobalStatus(`${fieldName} copied to clipboard!`, false, 2000);
                })
                .catch(err => {
                    console.error(`Could not copy ${fieldName}: `, err);
                    showGlobalStatus(`Failed to copy ${fieldName}. See console.`, true, 3000);
                });
        }

        titleTextarea.addEventListener('click', () => copyTextareaContent(titleTextarea, 'Title'));
        contentTextarea.addEventListener('click', () => copyTextareaContent(contentTextarea, 'Content'));

        dragElement(panel);
    }

    function showGlobalStatus(message, isError = false, duration = null) {
        const statusEl = document.getElementById('ilg-global-status');
        if (statusEl) {
            statusEl.textContent = message;
            statusEl.style.color = isError ? 'red' : '#006400';

            if (statusEl.timeoutId) {
                clearTimeout(statusEl.timeoutId);
            }

            const timeoutDuration = duration !== null ? duration : (isError ? 5000 : 3000);
            statusEl.timeoutId = setTimeout(() => {
                if (statusEl.textContent === message) {
                    statusEl.textContent = '';
                }
                statusEl.timeoutId = null;
            }, timeoutDuration);
        }
    }


    function generateAndDisplay(templateId) {
        const pageData = gatherAllPageData();
        const template = emailTemplates.find(t => t.id === templateId);

        if (!template) {
            console.error('錯誤：找不到郵件模板！ID:', templateId);
            showGlobalStatus('錯誤：找不到郵件模板！', true);
            document.getElementById('ilg-output-area').style.display = 'none';
            return;
        }

        let title = template.titleTemplate;
        let content = template.contentTemplate;

        title = title.replace(/#PjNum#/g, pageData.pjNum);
        content = content.replace(/#PjNum#/g, pageData.pjNum);
        content = content.replace(/#OdrNum#/g, pageData.odrNum);
        content = content.replace(/xxx年xx月/g, pageData.rocDateBooked);
        content = content.replace(/#PjScope#/g, pageData.pjScope);
        content = content.replace(/#Project Hold Reason#/g, pageData.projectHoldReason);
        content = content.replace(/^(Project Handler \/Email:)\s*$/gm, `$1 ${pageData.projectHandlerEmail}`);
        content = content.replace(/(Project Handler \/Email:)(?!.*\S)/gm, `$1 ${pageData.projectHandlerEmail}`);


        if (template.id === 'final') {
            content = content.replace(/#DeadlineDate#/g, pageData.deadlineDate);
        }

        document.getElementById('ilg-title-output').value = title;
        document.getElementById('ilg-content-output').value = content;
        document.getElementById('ilg-output-area').style.display = 'block';
        showGlobalStatus(`${buttonLabels[template.id] || template.name} 已生成。`);
    }

    // --- Draggable Panel Functionality (保持不變) ---
    function dragElement(elmnt) {
        var pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
        const header = document.getElementById("ilg-header"); // Target the header for dragging
        if (header) {
            header.onmousedown = dragMouseDown;
        } else {
            elmnt.onmousedown = dragMouseDown; // Fallback if header not found
        }
        function dragMouseDown(e) {
            e = e || window.event;
            e.preventDefault();
            pos3 = e.clientX;
            pos4 = e.clientY;
            document.onmouseup = closeDragElement;
            document.onmousemove = elementDrag;
        }
        function elementDrag(e) {
            e = e || window.event;
            e.preventDefault();
            pos1 = pos3 - e.clientX;
            pos2 = pos4 - e.clientY;
            pos3 = e.clientX;
            pos4 = e.clientY;

            let newTop = elmnt.offsetTop - pos2;
            let newLeft = elmnt.offsetLeft - pos1;

            // Boundary checks
            const maxTop = window.innerHeight - elmnt.offsetHeight;
            const maxLeft = window.innerWidth - elmnt.offsetWidth;
            newTop = Math.max(0, Math.min(newTop, maxTop));
            newLeft = Math.max(0, Math.min(newLeft, maxLeft));

            elmnt.style.top = newTop + "px";
            elmnt.style.left = newLeft + "px";
        }
        function closeDragElement() {
            document.onmouseup = null;
            document.onmousemove = null;
        }
    }

    // --- Styling (MODIFIED) ---
    GM_addStyle(`
        #inactiveLetterPanel {
            position: fixed;
            bottom: 10px;
            right: 10px;
            background-color: #f0f0f0;
            border: 1px solid #bababa;
            padding: 0;
            z-index: 10000;
            width: 220px; /* MODIFIED: Further reduced width for horizontal buttons */
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            font-family: Segoe UI, Arial, sans-serif;
            font-size: 11px; /* MODIFIED: Base font size further reduced */
            border-radius: 2px;
            display: none;
        }
        #ilg-header {
            padding: 2px 6px; /* MODIFIED: Drastically reduced padding for smaller header */
            cursor: move;
            background-color: #e0e0e0;
            color: #222; /* Darker text for contrast */
            border-bottom: 1px solid #c5c5c5;
            border-top-left-radius: 2px;
            border-top-right-radius: 2px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            height: 18px; /* MODIFIED: Explicit small height */
            box-sizing: border-box;
        }
        #ilg-header-title { /* Target title specifically */
            margin: 0;
            font-size: 1.0em; /* Approx 11px */
            font-weight: bold;
            line-height: 1; /* Adjust for tight fit */
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        #ilg-close-btn {
            font-size: 14px; /* MODIFIED */
            font-weight: bold;
            color: #555;
            cursor: pointer;
            padding: 0 2px;
            line-height: 1; /* Adjust for tight fit */
        }
        #ilg-close-btn:hover { color: #000; }

        #ilg-buttons-container {
            padding: 5px 6px; /* MODIFIED */
            display: flex;
            flex-direction: row;   /* MODIFIED: Horizontal buttons */
            flex-wrap: nowrap;     /* MODIFIED: Try to keep on one line */
            justify-content: space-around; /* Distribute space, or use space-between */
            gap: 3px;              /* MODIFIED: Reduced gap */
        }
        .ilg-action-button {
            /* display: inline-block; /* Let flexbox handle it */
            font-weight: bold;
            font-size: 0.9em; /* Approx 10px, adjust as needed */
            line-height: 1.3em;
            padding: 3px 5px;    /* MODIFIED: Tighter padding */
            border: 1px solid #aeaeae;
            color: #333333;
            background-color: #f9f9f9;
            text-decoration: none;
            cursor: pointer;
            border-radius: 2px;
            transition: background-color 0.2s, border-color 0.2s;
            text-align: center; /* Center text in button */
            flex-grow: 1;       /* Allow buttons to grow to fill space */
            flex-shrink: 1;
            flex-basis: 0;      /* Distribute space evenly */
            white-space: nowrap; /* Prevent text wrapping in button */
            overflow: hidden;
            text-overflow: ellipsis; /* If text is too long */
        }
        .ilg-action-button:hover {
            background-color: #e8e8e8;
            border-color: #999999;
        }
        /* Removed .ilg-arrow styles as arrow is removed from button text for now */

        #ilg-output-area {
            padding: 6px 8px 8px 8px; /* MODIFIED */
        }
        #ilg-output-area h4 {
             font-size: 1.0em;
             margin-bottom: 4px;
             font-weight: bold;
             color: #333;
        }
        #ilg-output-area label {
             font-size: 0.95em;
             font-weight: bold;
             display: block;
             margin-bottom: 2px;
             color: #444;
        }
        #ilg-output-area textarea {
            width: 100%;
            box-sizing: border-box;
            margin-bottom: 6px;
            padding: 3px;
            border: 1px solid #c0c0c0;
            border-radius: 2px;
            font-family: Consolas, monospace;
            font-size: 9px; /* MODIFIED: Smaller font for textareas */
            background-color: #fff;
        }
        #ilg-output-area textarea#ilg-content-output {
             margin-bottom: 4px;
        }
        #ilg-global-status {
            padding: 0 8px 5px 8px;
            font-style: italic;
            color: #555;
            font-size: 0.9em;
            min-height: 1em;
        }
    `);

    // --- Initialize ---
    function init() {
        if (document.querySelector('dt') || document.querySelector('.display-label-row')) {
            if (!document.getElementById('inactiveLetterPanel')) {
                 createUI();
            }
            const panel = document.getElementById('inactiveLetterPanel');
            if (panel) {
                panel.style.display = 'block';
            }
            console.log("Inactive Letter Generator (Ultra-Compact): UI created and shown.");
        } else {
            console.log("Inactive Letter Generator (Ultra-Compact): Required page structure not found. UI not added.");
        }
    }
    init();

})();
