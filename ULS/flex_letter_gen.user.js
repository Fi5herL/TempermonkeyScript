// ==UserScript==
// @name         Combined Letter & Report Generator (Fixed, Collapsible)
// @name:zh-CN   合併信件與報告生成器 (固定,可縮放)
// @namespace    http://tampermonkey.net/
// @version      4.1
// @description  Merges Inactive Letter Generator and TAT/ECD/NOA Letter buttons into a single, fixed, collapsible floating panel.
// @description:zh-CN 將 Inactive Letter 生成器與 TAT/ECD/NOA Letter 按鈕合併到單個固定的、可縮放的浮動面板中。
// @author       Your Name (Merged & Modified by AI)
// @match        https://portal.ul.com/Project/Details/*
// @grant        GM_addStyle
// @grant        GM_openInTab
// ==/UserScript==

(function() {
    'use strict';

    // --- Constants for Script 2 (Report Button Generator) ---
    const RBG_TARGET_ANCHOR_SELECTOR = '.section-crumbs-li a';
    const RBG_ARROW_CHAR = ' ›';
    const RBG_BASE_REPORT_URL = 'https://epic.ul.com/Report';

    // --- Helper Functions (from Script 1 - ILG) ---
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

    // --- Helper Function (from Script 2 - RBG) ---
    function rbgGetFormattedDate() {
        const today = new Date();
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    // --- Email Templates & Button Labels (from Script 1 - ILG) ---
    const ilgButtonLabels = {
        'notice': 'Notice', 'inactive1': '1', 'inactive2': '2', 'inactive3': '3', 'final': 'Final'
    };
    const ilgEmailTemplates = [
        { id: 'notice', name: 'Notice Inactive Letter', titleTemplate: "Project Inactive Letter–Project #PjNum#/關於UL項目#PjNum#暫停通知書", contentTemplate: `尊敬的客戶：\n\n感謝您及貴公司對UL服務的信任與支持。\n\n關於貴公司xxx年xx月提交的UL認證項目，認證項目編號 #PjNum# ，服務訂單編號 #OdrNum# ，認證申請描述 #PjScope#，我們不得不書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。即日起，項目進度變更為暫停狀態。\n\n為了繼續推進項目進程，我們需要貴公司盡快提供如下信息\n#Project Hold Reason#\n\n請知悉，在UL收到完整並正確的如上信息後，您的項目方可重新啟動。任何不明確之處，歡迎您隨時與我們聯繫，我們也將與貴公司保持積極的互動以盡快重啟您的認證項目。\n\n順祝\n商祺\n\nProject Handler /Email:\nEngineer Manager /Email:` },
        { id: 'inactive1', name: 'Inactive Letter 1', titleTemplate: "The 1st project inactive follow up letter–Project #PjNum#/第一次項目暫停跟進通知書", contentTemplate: `尊敬的客戶：\n\n感謝您及貴公司對UL服務的信任與支持。\n\n關於貴公司xxx年xx月提交的UL認證項目，認證項目編號#PjNum#，服務訂單編號#OdrNum#，認證申請描述#PjScope#，我們不得不書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。即日起，項目進度變更為暫停狀態。\n\n為了繼續推進項目進程，我們需要貴公司盡快提供如下信息\n#Project Hold Reason#\n\n到目前為止，我們尚未從貴公司收到完整並正確的上述信息，項目仍處於暫停狀態。\n\n請知悉，在UL收到完整並正確的如上信息後，您的項目方可重新啟動。任何不明確之處，歡迎您隨時與我們聯繫，我們也將與貴公司保持積極的互動以盡快重啟您的認證項目。\n\n順祝\n商祺\n\nProject Handler /Email:` },
        { id: 'inactive2', name: 'Inactive Letter 2', titleTemplate: "The 2nd project inactive follow up letter–Project #PjNum#/第二次項目暫停跟進通知書", contentTemplate: `尊敬的客戶：\n\n感謝您及貴公司對UL服務的信任與支持。\n\n關於貴公司xxx年xx月提交的UL認證項目，認證項目編號#PjNum#，服務訂單編號#OdrNum#，認證申請描述#PjScope#，我們不得不書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。即日起，項目進度變更為暫停狀態。\n\n為了繼續推進項目進程，我們需要貴公司盡快提供如下信息\n#Project Hold Reason#\n\n一個月前，我們發出第一次項目暫停跟進通知書，但到目前為止，我們仍未從貴公司收到完整並正確的上述信息，項目仍處於暫停狀態。\n\n請知悉，在UL收到完整並正確的如上信息後，您的項目方可重新啟動。任何不明確之處，歡迎您隨時與我們聯繫，我們也將與貴公司保持積極的互動以盡快重啟您的認證項目。\n\n順祝\n商祺\n\nProject Handler /Email:` },
        { id: 'inactive3', name: 'Inactive Letter 3', titleTemplate: "The 3nd project inactive follow up letter–Project #PjNum#/第三次項目暫停跟進通知書", contentTemplate: `尊敬的客戶：\n\n感謝您及貴公司對UL服務的信任與支持。\n\n關於貴公司xxx年xx月提交的UL認證項目，認證項目編號#PjNum#，服務訂單編號#OdrNum#，認證申請描述#PjScope#，我們不得不書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。即日起，項目進度變更為暫停狀態。\n\n為了繼續推進項目進程，我們需要貴公司盡快提供如下信息\n#Project Hold Reason#\n\n二個月前，我們發出第一次項目暫停跟進通知書，並且在一個月前，向貴司發出第二次項目暫停跟進通知書，但是到目前為止，我們仍未從貴公司收到完整併正確的上述信息，項目依舊處於暫停狀態。\n\n請知悉，在UL收到完整並正確的如上信息後，您的項目方可重新啟動。任何不明確之處，歡迎您隨時與我們聯繫，我們也將與貴公司保持積極的互動以盡快重啟您的認證項目。\n\n順祝\n商祺\n\nProject Handler /Email:\nField Sales /Email:` },
        { id: 'final', name: 'Inactive Letter Final', titleTemplate: "The final notice before project close by letter – Project #PjNum#/項目終止前最後提醒", contentTemplate: `尊敬的客戶：\n感謝您及貴公司對UL服務的信任與支持。\n關於貴公司xxx年xx月提交的UL認證項目，認證項目編號#PjNum#，服務訂單編號： #OdrNum#, 認證申請描述 #PjScope#。約四個月前，我們曾書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。項目進度變更為暫停狀態。\n為了繼續推進項目進程，我們需要貴公司盡快提供如下信息，或就該等信息和材料的提交提出明確的時間表。\n#Project Hold Reason#\n\n由於未能收到有效反饋，約三個月前，我們發出第一次項目暫停跟進通知書，並且在隨後二個月，接連向貴司發出第二次以及第三次項目暫停跟進通知書。但是到目前為止，我們仍未從貴公司收到完整併正確的上述信息，項目始終處於暫停狀態。\n\n根據過去四個月的項目進展狀況，我們在此最後一次向您發出項目提醒函，請您務必引起重視，如果該項目在未來兩週内，也就是#DeadlineDate#前您仍不能提供完整併正確的上述必要信息/樣品，我們不得不遺憾的通知您，我們將終止貴公司【#OdrNum#】號服務訂單及其項下之#PjNum#認證項目。項目終止後，我們將就我司已經提供的服務向您收取相應的費用；項目採用預付款方式支付的，我們將在扣除必要費用後，退還您的剩餘款項。如以上服務需求在未來需要再次啟動，您可以向我們索取一份新的正式報價（報價有效期為三個月）。我們將重新核定貴公司的服務需求，並向您發出新的報價。\n\n順祝\n商祺\n\nField Sales /Email:       Project Handler /Email:\nSales Manager /Email:       Engineer Manager /Email:` }
    ];

    // --- Data for Script 2 (RBG) buttons ---
    const rbgButtonsData = [
        { id: 'ecdLetterFloatingButton', text: 'ECD Letter', params: { templateUNID: 'AHL ECD Letter', selectedOutputType: '.eml', addFRDate: false }},
        { id: 'tatLetterFloatingButton', text: 'TAT Letter', params: { templateUNID: 'AHL TAT Letter', selectedOutputType: '.eml', addFRDate: false }},
        { id: 'noaLetterFloatingButton', text: 'NOA Letter', params: { templateUNID: 'Notice of Authorization or Completion Letter', selectedOutputType: '.default', addFRDate: true }}
    ];

    // --- Data Gathering (from Script 1 - ILG) ---
    function ilgGatherAllPageData() {
        const pjNum = extractFieldByLabel('Oracle Project Number');
        const dateBookedRaw = extractFieldByLabel('Date Booked');
        const pjScope = extractFieldByLabel('Project Scope');
        const projectHoldReason = extractFieldByLabel('Project Hold Reason');
        const odrNum = extractOdrNum();
        const rocDateBooked = formatToROCYearMonth(dateBookedRaw);
        const deadlineDate = getFutureROCDate(14);
        const projectHandlerEmail = extractProjectHandlerEmail();
        return { pjNum: pjNum || "N/A", odrNum: odrNum || "N/A", rocDateBooked: rocDateBooked || "未知日期", pjScope: pjScope || "N/A", projectHoldReason: projectHoldReason || "N/A", deadlineDate: deadlineDate, projectHandlerEmail: projectHandlerEmail || "N/A" };
    }

    // --- UI Creation and Logic ---
    function createUI() {
        const panel = document.createElement('div');
        panel.id = 'mergedUtilityPanel';
        panel.innerHTML = `
            <div id="mrg-header">
                <h3 id="mrg-header-title">Project Utilities</h3>
                <span id="mrg-close-btn">×</span>
            </div>
            <div id="mrg-content-wrapper">
                <div class="mrg-section">
                    <h4 class="mrg-section-title">Inactive Letters</h4>
                    <div id="ilg-buttons-container"></div>
                    <div id="ilg-output-area" style="display:none;">
                        <h5 style="margin-top:0;">Generated Email:</h5>
                        <label for="ilg-title-output">Title (click to copy):</label>
                        <textarea id="ilg-title-output" rows="2" readonly title="Click to copy title"></textarea>
                        <label for="ilg-content-output">Content (click to copy):</label>
                        <textarea id="ilg-content-output" rows="5" readonly title="Click to copy content"></textarea>
                    </div>
                </div>
                <hr class="mrg-divider">
                <div class="mrg-section">
                    <h4 class="mrg-section-title">Report Generators</h4>
                    <div id="rbg-buttons-container"></div>
                </div>
                <div id="mrg-global-status"></div>
            </div>
        `;
        document.body.appendChild(panel);

        const header = document.getElementById('mrg-header');
        const closeButton = document.getElementById('mrg-close-btn');
        // const contentWrapper = document.getElementById('mrg-content-wrapper'); // Not directly needed for toggle if using class on panel

        // Restore collapsed state
        if (localStorage.getItem('mergedPanelCollapsed') === 'true') {
            panel.classList.add('mrg-collapsed');
        }

        header.addEventListener('click', function(event) {
            if (event.target.id === 'mrg-close-btn') {
                // Allow close button's own event listener to handle this
                return;
            }
            panel.classList.toggle('mrg-collapsed');
            if (panel.classList.contains('mrg-collapsed')) {
                localStorage.setItem('mergedPanelCollapsed', 'true');
            } else {
                localStorage.removeItem('mergedPanelCollapsed');
            }
        });

        closeButton.addEventListener('click', () => {
            panel.style.display = 'none';
            // When closing, if output area was visible, hide it for next time
            document.getElementById('ilg-output-area').style.display = 'none';
            document.getElementById('mrg-global-status').textContent = '';
        });


        // --- ILG (Script 1) Button Creation ---
        const ilgButtonsContainer = panel.querySelector('#ilg-buttons-container');
        ilgEmailTemplates.forEach(template => {
            const button = document.createElement('button');
            button.textContent = ilgButtonLabels[template.id] || template.id;
            button.className = 'mrg-action-button ilg-button';
            button.title = `Generate ${template.name}`;
            button.addEventListener('click', () => ilgGenerateAndDisplay(template.id));
            ilgButtonsContainer.appendChild(button);
        });

        const titleTextarea = document.getElementById('ilg-title-output');
        const contentTextarea = document.getElementById('ilg-content-output');

        function ilgCopyTextareaContent(textareaElement, fieldName) {
            if (!textareaElement.value) {
                showGlobalStatus(`${fieldName} is empty. Nothing to copy.`, true, 2000);
                return;
            }
            navigator.clipboard.writeText(textareaElement.value)
                .then(() => { showGlobalStatus(`${fieldName} copied to clipboard!`, false, 2000); })
                .catch(err => { console.error(`Could not copy ${fieldName}: `, err); showGlobalStatus(`Failed to copy ${fieldName}. See console.`, true, 3000); });
        }

        titleTextarea.addEventListener('click', () => ilgCopyTextareaContent(titleTextarea, 'Title'));
        contentTextarea.addEventListener('click', () => ilgCopyTextareaContent(contentTextarea, 'Content'));

        // --- RBG (Script 2) Button Creation ---
        const rbgButtonsContainer = panel.querySelector('#rbg-buttons-container');
        rbgButtonsData.forEach((config) => {
            rbgCreateReportButton(config, rbgButtonsContainer);
        });

        if (!document.querySelector(RBG_TARGET_ANCHOR_SELECTOR)) {
             console.warn(`Merged Script Tip: Target anchor for Report Generators ('${RBG_TARGET_ANCHOR_SELECTOR}') not found on page load.`);
             showGlobalStatus(`Warning: Report Generator target ('${RBG_TARGET_ANCHOR_SELECTOR}') not found.`, true, 5000);
        }
        // REMOVED: dragElement(panel);
    }

    function showGlobalStatus(message, isError = false, duration = null) {
        const statusEl = document.getElementById('mrg-global-status');
        if (statusEl) {
            statusEl.textContent = message;
            statusEl.style.color = isError ? 'red' : '#006400';
            if (statusEl.timeoutId) clearTimeout(statusEl.timeoutId);
            const timeoutDuration = duration !== null ? duration : (isError ? 5000 : 3000);
            statusEl.timeoutId = setTimeout(() => {
                if (statusEl.textContent === message) statusEl.textContent = '';
                statusEl.timeoutId = null;
            }, timeoutDuration);
        }
    }

    function ilgGenerateAndDisplay(templateId) {
        const pageData = ilgGatherAllPageData();
        const template = ilgEmailTemplates.find(t => t.id === templateId);
        if (!template) {
            console.error('ILG Error: Template not found! ID:', templateId);
            showGlobalStatus('ILG Error: Template not found!', true);
            document.getElementById('ilg-output-area').style.display = 'none';
            return;
        }
        let title = template.titleTemplate.replace(/#PjNum#/g, pageData.pjNum);
        let content = template.contentTemplate
            .replace(/#PjNum#/g, pageData.pjNum)
            .replace(/#OdrNum#/g, pageData.odrNum)
            .replace(/xxx年xx月/g, pageData.rocDateBooked)
            .replace(/#PjScope#/g, pageData.pjScope)
            .replace(/#Project Hold Reason#/g, pageData.projectHoldReason)
            .replace(/^(Project Handler \/Email:)\s*$/gm, `$1 ${pageData.projectHandlerEmail}`)
            .replace(/(Project Handler \/Email:)(?!.*\S)/gm, `$1 ${pageData.projectHandlerEmail}`);
        if (template.id === 'final') content = content.replace(/#DeadlineDate#/g, pageData.deadlineDate);
        document.getElementById('ilg-title-output').value = title;
        document.getElementById('ilg-content-output').value = content;
        document.getElementById('ilg-output-area').style.display = 'block';
        showGlobalStatus(`${ilgButtonLabels[template.id] || template.name} generated.`);
    }

    function rbgCreateReportButton(buttonConfig, container) {
        const button = document.createElement('button');
        button.id = buttonConfig.id;
        button.className = 'mrg-action-button rbg-button';
        button.textContent = buttonConfig.text + RBG_ARROW_CHAR;
        container.appendChild(button);
        button.addEventListener('click', function() {
            let hrefValueFromAnchor = null;
            const targetAnchorElement = document.querySelector(RBG_TARGET_ANCHOR_SELECTOR);
            if (targetAnchorElement) {
                hrefValueFromAnchor = targetAnchorElement.getAttribute('href');
                if (hrefValueFromAnchor === null || hrefValueFromAnchor.trim() === "") {
                    alert(`RBG Error: Target <a> tag (selector: ${RBG_TARGET_ANCHOR_SELECTOR}) 'href' is empty or missing.`);
                    showGlobalStatus(`RBG Error: Target href empty/missing.`, true); return;
                }
            } else {
                alert(`RBG Error: Target <a> tag not found on page.\nCheck selector '${RBG_TARGET_ANCHOR_SELECTOR}'.`);
                showGlobalStatus(`RBG Error: Target anchor not found.`, true); return;
            }
            const reportParams = new URLSearchParams();
            reportParams.append('TemplateUNID', buttonConfig.params.templateUNID);
            reportParams.append('SelectedOutputType', buttonConfig.params.selectedOutputType);
            reportParams.append('ProjectID', encodeURIComponent(hrefValueFromAnchor));
            reportParams.append('isWorkbench', 'False');
            if (buttonConfig.params.addFRDate) reportParams.append('FRDate', rbgGetFormattedDate());
            const targetUrl = `${RBG_BASE_REPORT_URL}?${reportParams.toString()}`;
            console.log(`RBG Button "${buttonConfig.text}" clicked. URL: ${targetUrl}`);
            GM_openInTab(targetUrl, { active: true });
            showGlobalStatus(`Opening ${buttonConfig.text}...`, false, 1500);
        });
        return button;
    }

    // --- Styling (Merged and Adjusted) ---
    GM_addStyle(`
        #mergedUtilityPanel {
            position: fixed;
            bottom: 10px;
            right: 10px;
            background-color: #f0f0f0;
            border: 1px solid #bababa;
            padding: 0;
            z-index: 10000;
            width: 280px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            font-family: Segoe UI, Arial, sans-serif;
            font-size: 11px;
            border-radius: 3px;
            display: none; /* Initially hidden, shown by init() */
        }
        #mrg-header {
            padding: 4px 8px;
            cursor: pointer; /* Changed from move to pointer for collapse/expand */
            background-color: #e0e0e0;
            color: #222;
            border-bottom: 1px solid #c5c5c5;
            border-top-left-radius: 3px;
            border-top-right-radius: 3px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            height: 22px;
            box-sizing: border-box;
        }
        #mrg-header-title {
            margin: 0;
            font-size: 1.1em;
            font-weight: bold;
            line-height: 1;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            user-select: none; /* Prevent text selection on header click */
        }
        #mrg-close-btn {
            font-size: 16px;
            font-weight: bold;
            color: #555;
            cursor: pointer; /* Explicitly pointer for close button */
            padding: 0 3px;
            line-height: 1;
            user-select: none;
        }
        #mrg-close-btn:hover { color: #000; }

        #mrg-content-wrapper {
            /* Styles for the content area that can be hidden */
        }
        #mergedUtilityPanel.mrg-collapsed #mrg-content-wrapper {
            display: none;
        }
        #mergedUtilityPanel.mrg-collapsed #mrg-header {
            border-bottom: none; /* No border when collapsed */
            border-bottom-left-radius: 3px; /* Make it fully rounded when collapsed */
            border-bottom-right-radius: 3px;
        }


        .mrg-section { padding: 8px; }
        .mrg-section-title { font-size: 1.05em; font-weight: bold; color: #333; margin: 0 0 6px 0; padding-bottom: 3px; border-bottom: 1px solid #d0d0d0; }
        .mrg-divider { border: none; border-top: 1px dashed #c5c5c5; margin: 0px 8px; }

        #ilg-buttons-container { display: flex; flex-direction: row; flex-wrap: nowrap; justify-content: space-around; gap: 4px; margin-bottom: 8px; }
        .ilg-button { flex-grow: 1; flex-shrink: 1; flex-basis: 0; }
        #ilg-output-area { padding: 6px; border: 1px solid #d0d0d0; background-color: #f9f9f9; border-radius: 2px; }
        #ilg-output-area h5 { font-size: 1.0em; margin-top:0; margin-bottom: 4px; font-weight: bold; color: #333; }
        #ilg-output-area label { font-size: 0.95em; font-weight: bold; display: block; margin-bottom: 2px; color: #444; }
        #ilg-output-area textarea { width: 100%; box-sizing: border-box; margin-bottom: 6px; padding: 3px; border: 1px solid #c0c0c0; border-radius: 2px; font-family: Consolas, monospace; font-size: 9px; background-color: #fff; cursor: pointer; }
        #ilg-output-area textarea#ilg-content-output { margin-bottom: 4px; }

        #rbg-buttons-container { display: flex; flex-direction: column; gap: 5px; }
        .rbg-button { text-align: left; width: 100%; box-sizing: border-box; }

        .mrg-action-button { font-weight: bold; font-size: 0.95em; line-height: 1.3em; padding: 4px 8px; border: 1px solid #aeaeae; color: #333333; background-color: #f9f9f9; text-decoration: none; cursor: pointer; border-radius: 2px; transition: background-color 0.2s, border-color 0.2s; text-align: center; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
        .mrg-action-button:hover { background-color: #e8e8e8; border-color: #999999; }

        #mrg-global-status { padding: 5px 8px; font-style: italic; color: #555; font-size: 0.9em; min-height: 1em; border-top: 1px solid #e0e0e0; margin-top: 5px; }
        /* Hide status bar when panel is collapsed and content is hidden */
        #mergedUtilityPanel.mrg-collapsed #mrg-global-status {
             display: none; /* Should be handled by content wrapper, but good fallback */
        }

    `);

    // --- Initialize ---
    function init() {
        if (document.querySelector('dt') || document.querySelector('.display-label-row')) {
            if (!document.getElementById('mergedUtilityPanel')) {
                 createUI();
            }
            const panel = document.getElementById('mergedUtilityPanel');
            if (panel) {
                panel.style.display = 'block';
            }
            console.log("Merged Utility Panel (Fixed, Collapsible): UI created and shown.");
        } else {
            console.log("Merged Utility Panel: Required page structure for ILG not found. UI not fully functional or not added.");
        }
    }
    init();

})();
