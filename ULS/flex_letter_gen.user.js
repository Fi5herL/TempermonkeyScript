// ==UserScript==
// @name         Letter Generator (Inactive, ECD, TAT, NOA, Travel, Close)
// @name:zh-CN   信件生成器 (固定,可縮放,差旅審批,結案信)
// @namespace    http://tampermonkey.net/
// @version      4.7.3 // NEW: Incremented version
// @description  Merges Inactive Letter Generator, TAT/ECD/NOA/PI/Close Letter buttons, and adds a Travel Approval form into a single, fixed, collapsible floating panel.
// @description:zh-CN 將 Inactive Letter 生成器、TAT/ECD/NOA/PI/結案信 Letter 按鈕以及差旅申請表單合併到單個固定的、可縮放的浮動面板中。
// @author       Your Name
// @match        https://portal.ul.com/Project/Details/*
// @grant        GM_addStyle
// @grant        GM_openInTab
// ==/UserScript==

(function() {
    'use strict';

    // --- Constants ---
    const RBG_TARGET_ANCHOR_SELECTOR = 'a[href*="/Project/Index/"]';
    const RBG_ARROW_CHAR = ' ›';
    const RBG_BASE_REPORT_URL = 'https://epic.ul.com/Report';

    // --- Helper Functions ---
    function extractFieldByLabel(labelText) {
        const xpath = `//div[@class='display-label-row' and normalize-space(.)='${labelText}']/following-sibling::div[@class='display-field-row'][1]`;
        const fieldElement = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        return fieldElement ? fieldElement.textContent.trim() : '';
    }

    function extractProjectName() {
        const xpath = "//div[contains(@class, 'display-label-row') and contains(@class, 'customer-flag') and normalize-space(.)='Project Name']/following-sibling::div[contains(@class, 'display-field-row') and contains(@class, 'ellipsis-ctrl')][1]";
        const fieldElement = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        return fieldElement ? fieldElement.textContent.trim() : '';
    }

    function extractOdrNum() {
        const xpath = "//dt[normalize-space(.)='Order Number:']/following-sibling::dd[1]//span";
        const element = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        return element ? element.textContent.trim() : '';
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
            const month = parseInt(parts[0], 10), day = parseInt(parts[1], 10), year = parseInt(parts[2], 10);
            if (!isNaN(month) && !isNaN(day) && !isNaN(year)) return `中華民國${year - 1911}年${month}月`;
        }
        return "未知日期 (格式錯誤)";
    }
    function getFutureROCDate(daysOffset) {
        const d = new Date(); d.setDate(d.getDate() + daysOffset);
        return `中華民國${d.getFullYear() - 1911}年${d.getMonth() + 1}月${d.getDate()}日`;
    }
    function rbgGetFormattedDate() {
        const d = new Date();
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    }

    const ilgButtonLabels = { 'notice': 'Notice', 'inactive1': '1', 'inactive2': '2', 'inactive3': '3', 'final': 'Final' };
    const ilgEmailTemplates = [ // Templates as in your v4.7.1 (full content assumed here)
        { id: 'notice', name: 'Notice Inactive Letter', titleTemplate: "Project Inactive Letter–Project #PjNum#/關於UL項目#PjNum#暫停通知書", contentTemplate: `尊敬的客戶：\n\n感謝您及貴公司對UL服務的信任與支持。\n\n關於貴公司xxx年xx月提交的UL認證項目，認證項目編號 #PjNum# ，服務訂單編號 #OdrNum# ，認證申請描述 #PjScope#，我們不得不書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。即日起，項目進度變更為暫停狀態。\n\n為了繼續推進項目進程，我們需要貴公司盡快提供如下信息\n#Project Hold Reason#\n\n請知悉，在UL收到完整並正確的如上信息後，您的項目方可重新啟動。任何不明確之處，歡迎您隨時與我們聯繫，我們也將與貴公司保持積極的互動以盡快重啟您的認證項目。\n\n順祝\n商祺\n\nProject Handler /Email:\nEngineer Manager /Email:` },
        { id: 'inactive1', name: 'Inactive Letter 1', titleTemplate: "The 1st project inactive follow up letter–Project #PjNum#/第一次項目暫停跟進通知書", contentTemplate: `尊敬的客戶：\n\n感謝您及貴公司對UL服務的信任與支持。\n\n關於貴公司xxx年xx月提交的UL認證項目，認證項目編號#PjNum#，服務訂單編號#OdrNum#，認證申請描述#PjScope#，我們不得不書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。即日起，項目進度變更為暫停狀態。\n\n為了繼續推進項目進程，我們需要貴公司盡快提供如下信息\n#Project Hold Reason#\n\n到目前為止，我們尚未從貴公司收到完整並正確的上述信息，項目仍處於暫停狀態。\n\n請知悉，在UL收到完整並正確的如上信息後，您的項目方可重新啟動。任何不明確之處，歡迎您隨時與我們聯繫，我們也將與貴公司保持積極的互動以盡快重啟您的認證項目。\n\n順祝\n商祺\n\nProject Handler /Email:` },
        { id: 'inactive2', name: 'Inactive Letter 2', titleTemplate: "The 2nd project inactive follow up letter–Project #PjNum#/第二次項目暫停跟進通知書", contentTemplate: `尊敬的客戶：\n\n感謝您及貴公司對UL服務的信任與支持。\n\n關於貴公司xxx年xx月提交的UL認證項目，認證項目編號#PjNum#，服務訂單編號#OdrNum#，認證申請描述#PjScope#，我們不得不書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。即日起，項目進度變更為暫停狀態。\n\n為了繼續推進項目進程，我們需要貴公司盡快提供如下信息\n#Project Hold Reason#\n\n一個月前，我們發出第一次項目暫停跟進通知書，但到目前為止，我們仍未從貴公司收到完整並正確的上述信息，項目仍處於暫停狀態。\n\n請知悉，在UL收到完整並正確的如上信息後，您的項目方可重新啟動。任何不明確之處，歡迎您隨時與我們聯繫，我們也將與貴公司保持積極的互動以盡快重啟您的認證項目。\n\n順祝\n商祺\n\nProject Handler /Email:` },
        { id: 'inactive3', name: 'Inactive Letter 3', titleTemplate: "The 3nd project inactive follow up letter–Project #PjNum#/第三次項目暫停跟進通知書", contentTemplate: `尊敬的客戶：\n\n感謝您及貴公司對UL服務的信任與支持。\n\n關於貴公司xxx年xx月提交的UL認證項目，認證項目編號#PjNum#，服務訂單編號#OdrNum#，認證申請描述#PjScope#，我們不得不書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。即日起，項目進度變更為暫停狀態。\n\n為了繼續推進項目進程，我們需要貴公司盡快提供如下信息\n#Project Hold Reason#\n\n二個月前，我們發出第一次項目暫停跟進通知書，並且在一個月前，向貴司發出第二次項目暫停跟進通知書，但是到目前為止，我們仍未從貴公司收到完整併正確的上述信息，項目依舊處於暫停狀態。\n\n請知悉，在UL收到完整並正確的如上信息後，您的項目方可重新啟動。任何不明確之處，歡迎您隨時與我們聯繫，我們也將與貴公司保持積極的互動以盡快重啟您的認證項目。\n\n順祝\n商祺\n\nProject Handler /Email:\nField Sales /Email:` },
        { id: 'final', name: 'Inactive Letter Final', titleTemplate: "The final notice before project close by letter – Project #PjNum#/項目終止前最後提醒", contentTemplate: `尊敬的客戶：\n感謝您及貴公司對UL服務的信任與支持。\n關於貴公司xxx年xx月提交的UL認證項目，認證項目編號#PjNum#，服務訂單編號： #OdrNum#, 認證申請描述 #PjScope#。約四個月前，我們曾書面通知您及貴公司，由於未提供下述信息該項目已不能正常進行產品認證審核。項目進度變更為暫停狀態。\n為了繼續推進項目進程，我們需要貴公司盡快提供如下信息，或就該等信息和材料的提交提出明確的時間表。\n#Project Hold Reason#\n\n由於未能收到有效反饋，約三個月前，我們發出第一次項目暫停跟進通知書，並且在隨後二個月，接連向貴司發出第二次以及第三次項目暫停跟進通知書。但是到目前為止，我們仍未從貴公司收到完整併正確的上述信息，項目始終處於暫停狀態。\n\n根據過去四個月的項目進展狀況，我們在此最後一次向您發出項目提醒函，請您務必引起重視，如果該項目在未來兩週内，也就是#DeadlineDate#前您仍不能提供完整併正確的上述必要信息/樣品，我們不得不遺憾的通知您，我們將終止貴公司【#OdrNum#】號服務訂單及其項下之#PjNum#認證項目。項目終止後，我們將就我司已經提供的服務向您收取相應的費用；項目採用預付款方式支付的，我們將在扣除必要費用後，退還您的剩餘款項。如以上服務需求在未來需要再次啟動，您可以向我們索取一份新的正式報價（報價有效期為三個月）。我們將重新核定貴公司的服務需求，並向您發出新的報價。\n\n順祝\n商祺\n\nField Sales /Email:       Project Handler /Email:\nSales Manager /Email:       Engineer Manager /Email:` }
    ];
    // MODIFIED: Added closeLetterFloatingButton
    const rbgButtonsData = [
        { id: 'ecdLetterFloatingButton', text: 'ECD Letter', params: { templateUNID: 'AHL ECD Letter', selectedOutputType: '.eml', addFRDate: false }},
        { id: 'tatLetterFloatingButton', text: 'TAT Letter', params: { templateUNID: 'AHL TAT Letter', selectedOutputType: '.eml', addFRDate: false }},
        { id: 'noaLetterFloatingButton', text: 'NOA Letter', params: { templateUNID: 'Notice of Authorization or Completion Letter', selectedOutputType: '.default', addFRDate: true }},
        { id: 'piLetterFloatingButton', text: 'PI Letter', params: { templateUNID: 'AHL Preliminary Evaluation', selectedOutputType: '.default', addFRDate: false }},
        { id: 'closeLetterFloatingButton', text: 'Close Letter', params: {} }, // NEW: Close Letter button definition
        { id: 'travelApprovalFloatingButton', text: 'Travel Approval', params: {}}
    ];

    let currentPageDataForTravelModal = { pjNum: "N/A", pjName: "N/A", pjScope: "N/A", projectHandlerEmail: "N/A"};

    function ilgGatherAllPageData() {
        const pjNum = extractFieldByLabel('Oracle Project Number') || "N/A";
        const pjName = extractProjectName() || "N/A";
        const pjScope = extractFieldByLabel('Project Scope') || "N/A";
        const { clientName, clientEmail } = extractClientInfo();
        const projectHandlerEmail = extractProjectHandlerEmail() || "N/A";
        currentPageDataForTravelModal = { pjNum, pjName, pjScope, projectHandlerEmail };
        return {
            pjNum, pjName, pjScope, projectHandlerEmail,
            odrNum: extractOdrNum() || "N/A",
            rocDateBooked: formatToROCYearMonth(extractFieldByLabel('Date Booked')) || "未知日期",
            projectHoldReason: extractFieldByLabel('Project Hold Reason') || "N/A",
            deadlineDate: getFutureROCDate(14),
            // projectHandlerEmail: extractProjectHandlerEmail() || "N/A",
            clientName,
            clientEmail
        };
    }

    function createUI() {
        ilgGatherAllPageData();
        const panel = document.createElement('div'); panel.id = 'mergedUtilityPanel';
        panel.innerHTML = `
            <div id="mrg-header"> <h3 id="mrg-header-title">Project Utilities</h3> <span id="mrg-close-btn">×</span> </div>
            <div id="mrg-content-wrapper">
                <div class="mrg-section"> <h4 class="mrg-section-title">Inactive Letters</h4> <div id="ilg-buttons-container"></div>
                    <div id="ilg-output-area" style="display:none;">
                        <h5 style="margin-top:0;">Generated Email:</h5>
                        <label for="ilg-title-output">Title (click to copy):</label> <textarea id="ilg-title-output" rows="2" readonly title="Click to copy title"></textarea>
                        <label for="ilg-content-output">Content (click to copy):</label> <textarea id="ilg-content-output" rows="5" readonly title="Click to copy content"></textarea>
                    </div>
                </div> <hr class="mrg-divider">
                <div class="mrg-section"> <h4 class="mrg-section-title">Letter Links & Forms</h4> <div id="rbg-buttons-container"></div> </div>
                <div id="mrg-global-status"></div>
            </div>`;
        document.body.appendChild(panel);
        const header = panel.querySelector('#mrg-header'), closeBtn = panel.querySelector('#mrg-close-btn');
        if (localStorage.getItem('mergedPanelCollapsed') === 'true') panel.classList.add('mrg-collapsed');
        header.onclick = (e) => { if (e.target !== closeBtn) { panel.classList.toggle('mrg-collapsed'); localStorage.setItem('mergedPanelCollapsed', panel.classList.contains('mrg-collapsed')); }};
        closeBtn.onclick = () => { panel.style.display = 'none'; panel.querySelector('#ilg-output-area').style.display = 'none'; panel.querySelector('#mrg-global-status').textContent = ''; };
        const ilgBC = panel.querySelector('#ilg-buttons-container');
        ilgEmailTemplates.forEach(t => { const b = document.createElement('button'); b.textContent = ilgButtonLabels[t.id]||t.id; b.className='mrg-action-button ilg-button'; b.title=`Generate ${t.name}`; b.onclick=()=>ilgGenerateAndDisplay(t.id); ilgBC.appendChild(b); });
        const titleTA = panel.querySelector('#ilg-title-output'), contentTA = panel.querySelector('#ilg-content-output');
        const copyTA=(el,n)=>{if(!el.value){showGlobalStatus(`${n} empty.`,1,2e3);return;}navigator.clipboard.writeText(el.value).then(()=>showGlobalStatus(`${n} copied!`,0,2e3)).catch(err=>{console.error(`Copy ${n} err:`,err);showGlobalStatus(`Fail copy ${n}.`,1,3e3);});};
        titleTA.onclick=()=>copyTA(titleTA,'Title'); contentTA.onclick=()=>copyTA(contentTA,'Content');
        panel.querySelector('#rbg-buttons-container').append(...rbgButtonsData.map(c => rbgCreateReportButton(c)));
        if (rbgButtonsData.some(b => b.id !== 'travelApprovalFloatingButton' && b.id !== 'closeLetterFloatingButton') && !document.querySelector(RBG_TARGET_ANCHOR_SELECTOR)) { // MODIFIED: Also exclude closeLetterFloatingButton from this check
            console.warn(`RBG Target '${RBG_TARGET_ANCHOR_SELECTOR}' not found.`); showGlobalStatus(`Warn: RBG target (for ProjectID) missing.`,1,7000);
        }
    }

    function showGlobalStatus(msg, isErr=0, dur=null) { const el=document.getElementById('mrg-global-status'); if(el){el.textContent=msg; el.style.color=isErr?'red':'#006400'; if(el.timeoutId)clearTimeout(el.timeoutId); el.timeoutId=setTimeout(()=>{if(el.textContent===msg)el.textContent='';el.timeoutId=null;},dur??(isErr?5e3:3e3));}}

    function ilgGenerateAndDisplay(templateId) {
        const d = ilgGatherAllPageData(), t = ilgEmailTemplates.find(x=>x.id===templateId);
        if (!t) { console.error('ILG Err: Tpl not found:',templateId); showGlobalStatus('ILG Err: Tpl miss!',1); return; }
        let title = t.titleTemplate.replace(/#PjNum#/g,d.pjNum);
        let content = t.contentTemplate.replace(/#PjNum#/g,d.pjNum).replace(/#OdrNum#/g,d.odrNum).replace(/xxx年xx月/g,d.rocDateBooked).replace(/#PjScope#/g,d.pjScope).replace(/#Project Hold Reason#/g,d.projectHoldReason).replace(/^(Project Handler \/Email:)\s*$/gm,`$1 ${d.projectHandlerEmail}`).replace(/(Project Handler \/Email:)(?!.*\S)/gm,`$1 ${d.projectHandlerEmail}`);
        if (t.id==='final') content=content.replace(/#DeadlineDate#/g,d.deadlineDate);
        document.getElementById('ilg-title-output').value=title; document.getElementById('ilg-content-output').value=content;
        document.getElementById('ilg-output-area').style.display='block'; showGlobalStatus(`${ilgButtonLabels[t.id]||t.name} gen.`);
    }

    let travelApprovalModalCreated = false;
    // MODIFIED: Added logic for closeLetterFloatingButton
    function rbgCreateReportButton(buttonConfig) {
        const button = document.createElement('button'); button.id = buttonConfig.id; button.className = 'mrg-action-button rbg-button';

        // Only add arrow for buttons that open reports
        button.textContent = buttonConfig.text + (
            buttonConfig.id !== 'travelApprovalFloatingButton' &&
            buttonConfig.id !== 'closeLetterFloatingButton' // NEW: Exclude close letter from arrow
            ? RBG_ARROW_CHAR : ''
        );

        if (buttonConfig.id === 'travelApprovalFloatingButton') {
            button.onclick = function() {
                ilgGatherAllPageData(); // Ensure data is fresh
                if (!travelApprovalModalCreated) { createTravelApprovalModal(); travelApprovalModalCreated = true; }
                document.getElementById('travelApprovalModalContainer').style.display = 'flex';
                initializeTravelApprovalModalLogic(currentPageDataForTravelModal.pjNum, currentPageDataForTravelModal.pjName, currentPageDataForTravelModal.pjScope, currentPageDataForTravelModal.projectHandlerEmail);
                showGlobalStatus('Travel Approval open.',0,1500);
            };
        } else if (buttonConfig.id === 'closeLetterFloatingButton') { // NEW: Logic for Close Letter Button
            button.onclick = function() {
                const d = ilgGatherAllPageData();

                if (!d.clientEmail || d.clientEmail === "N/A") {
                    alert("Client email not found on the page. Cannot generate Close Letter email.");
                    showGlobalStatus("Client email missing for Close Letter.", 1, 5000);
                    return;
                }
                if (!d.pjNum || d.pjNum === "N/A") {
                    alert("Project Number not found on the page. Cannot generate Close Letter email.");
                    showGlobalStatus("Project Number missing for Close Letter.", 1, 5000);
                    return;
                }


                // Get current date for MMDD placeholder
                const today = new Date();
                const mm = String(today.getMonth() + 1).padStart(2, '0'); // Months are 0-indexed
                const ddDate = String(today.getDate()).padStart(2, '0'); // Renamed from 'dd' to 'ddDate' to avoid conflict with 'd' variable
                const MMDD_placeholder_value = `${mm}${ddDate}`; // e.g., "0726"

                const title = `Close Letter for Project No ${d.pjNum}`;
                let content = `Dear #clientName#,\n\n` +
                              `It's pleasured to meet with you on MM/DD, #pjScope# is carried out.\n` +
                              `Please also agree to use this letter as the basis for closing the project of #pjNum#.\n\n` +
                              `Thank you.\n\n` +
                              `BR,\n` +
                              `#projectHandlerEmail#`;

                // Replace placeholders
                content = content.replace(/#clientName#/g, d.clientName || "Valued Customer") // Fallback for client name
                                 .replace(/#pjNum#/g, d.pjNum)
                                 .replace(/#pjScope#/g, d.pjScope)
                                 .replace(/MMDD/g, MMDD_placeholder_value)
                                 .replace(/#projectHandlerEmail#/g, d.projectHandlerEmail || "Your Name/Email"); // Fallback for handler email

                const mailtoLink = `mailto:${d.clientEmail}?subject=${encodeURIComponent(title)}&body=${encodeURIComponent(content)}`;
                window.open(mailtoLink, '_blank');
                showGlobalStatus('Opening Close Letter email...',0,1500);
            };
        } else { // Existing logic for other report buttons (ECD, TAT, NOA, PI)
            button.onclick = function() {
                const anchor = document.querySelector(RBG_TARGET_ANCHOR_SELECTOR);
                if (!anchor || !anchor.href) { alert(`RBG Err: Target link for ProjectID not found (selector: ${RBG_TARGET_ANCHOR_SELECTOR}).`); showGlobalStatus(`RBG Err: Target link missing.`,1); return; }
                const p = new URLSearchParams({ TemplateUNID:buttonConfig.params.templateUNID, SelectedOutputType:buttonConfig.params.selectedOutputType, ProjectID:anchor.href, isWorkbench:'False' });
                if (buttonConfig.params.addFRDate) p.append('FRDate',rbgGetFormattedDate());
                GM_openInTab(`${RBG_BASE_REPORT_URL}?${p.toString()}`,{active:true}); showGlobalStatus(`Opening ${buttonConfig.text}...`,0,1500);

                if (buttonConfig.id === 'noaLetterFloatingButton') {
                    const d = ilgGatherAllPageData();
                    if (!d.clientEmail || d.clientEmail === "N/A") { // Added check for NOA mailto
                        showGlobalStatus("Client email missing for NOA follow-up.", 1, 4000);
                        return; // Don't open mailto if no email
                    }
                    const title = `NOA Letter for ${d.pjNum}`;
                    const content = `Dear ${d.clientName || "Valued Customer"},\n\n` +
                        `Congratulations! UL's investigation of your product has been completed and the products were determined to comply with the applicable requirements.\n` +
                        `\nThe attached is a Notice of Authorization for your reference.\n\n` +
                        `If there is any other way in which I can help, do not hesitate to contact me.\n\n` +
                        `${d.projectHandlerEmail || "Your Name/Email"}`;

                    const mailtoLink = `mailto:${d.clientEmail}?subject=${encodeURIComponent(title)}&body=${encodeURIComponent(content)}`;
                    window.open(mailtoLink, '_blank');
                }
            };
        }
        return button;
    }

    const travelApprovalModalHTMLBodyContent = `
        <div class="ta-container">
            <span id="travelApprovalModalCloseButton" class="ta-custom-modal-close-btn">×</span>
            <h2 class="ta-header-title">Travel Approval Generator</h2>
            <div class="ta-main-layout-wrapper">
                <div class="ta-main-content">
                    <div> <label for="ta_tripReason">Trip Reason:</label> <textarea id="ta_tripReason"></textarea> </div>
                    <div> <label for="ta_chargeable">Travel category & Account Allocation:</label> <select id="ta_chargeable"><option value="None">None</option><option value="Non-Billable">Non-Billable</option><option value="Billable-Inv">Billable to project and Invoiceable</option><option value="Billable-NonInv">Billable to project and Non-Invoiceable</option></select> </div>
                    <label>Select Dates (Click to select/deselect):</label>
                    <div class="ta-calendar-container"> <div class="ta-calendar-header"> <button id="ta_prevMonthBtn" class="ta-cal-nav-btn">◀</button> <h3 id="ta_monthYearDisplay"></h3> <button id="ta_nextMonthBtn" class="ta-cal-nav-btn">▶</button> </div> <div class="ta-calendar-grid" id="ta_calendarGrid"></div> </div>
                    <div class="ta-links-above-legs"> <a href="https://www.uber.com/global/zh-tw/price-estimate/" target="_blank" rel="noopener noreferrer">[Uber]</a> <a href="https://www.thsrc.com.tw/ArticleContent/a3b630bb-1066-4352-a1ef-58c7b4e8ef7c" target="_blank" rel="noopener noreferrer">[HSR]</a> <a href="https://www.metro.taipei/cp.aspx?n=ECEADC266D7120A7" target="_blank" rel="noopener noreferrer">[MRT]</a> <a href="https://www.mtaxi.com.tw/taxi-fare-estimate/" target="_blank" rel="noopener noreferrer">[mtaxi]</a> </div>
                    <div class="ta-trip-leg-embedded-section">
                        <label for="ta_newLocation">Add Custom Location (for From/To dropdowns):</label>
                        <div class="ta-add-location-container"> <input type="text" id="ta_newLocation" placeholder="Enter new location name"> <button id="ta_addCustomLocationBtn" class="ta-action-button-small">Add Location</button> <button id="ta_resetLocationsBtn" class="ta-action-button-secondary ta-action-button-small">Reset Locations</button> </div>
                        <h3 class="ta-section-title-inner">Add/Modify Trip Legs</h3>
                        <div class="ta-location-pair-selection"> <div> <label for="ta_fromLocationSelect">From:</label> <select id="ta_fromLocationSelect"></select> </div> <div> <label for="ta_toLocationSelect">To:</label> <select id="ta_toLocationSelect"></select> </div> <button id="ta_addLegToListBtn" class="ta-action-button-primary ta-action-button-small">Add Leg</button> </div>
                        <label>Current Legs:</label> <ul id="ta_currentLegsList" class="ta-list"></ul>
                    </div>
                    <div class="ta-action-buttons"> <button id="ta_generateEmailBtn" class="ta-action-button-primary ta-action-button-small"><b>Generate & Open Email</b></button> <button id="ta_downloadTextFileBtn" class="ta-action-button-secondary ta-action-button-small">Download Details</button> </div>
                </div>
            </div>
            <div id="ta_emailPreviewContainer" class="ta-preview-box" style="display: none;"> <h3 class="ta-section-title-inner">Email Preview (Plain Text):</h3> <pre id="ta_emailPreviewContent"></pre> </div>
            <div id="ta_emailLinkContainer" class="ta-preview-box" style="display: none;"> <h3 class="ta-section-title-inner">Email Link:</h3> <p id="ta_emailLinkText"></p> <a id="ta_mailtoLink" href="#" target="_blank">Click here to open email manually</a> </div>
        </div>
    `;

    function createTravelApprovalModal() {
        const modalContainer = document.createElement('div'); modalContainer.id = 'travelApprovalModalContainer';
        modalContainer.innerHTML = travelApprovalModalHTMLBodyContent; document.body.appendChild(modalContainer);
        modalContainer.querySelector('#travelApprovalModalCloseButton').onclick = () => { modalContainer.style.display = 'none'; };
        modalContainer.onclick = (e) => { if (e.target === modalContainer) modalContainer.style.display = 'none'; };
    }

    function initializeTravelApprovalModalLogic(pagePjNum = "N/A", pagePjName = "N/A", pagePjScope = "N/A", pageProjectHandlerEmail = "N/A") {
        const modalContainer = document.getElementById('travelApprovalModalContainer'); if (!modalContainer) return;
        let selectedDates = [], defaultLocations = ["台北", "關渡賓士大樓", "群通大樓", "新北", "桃園", "新竹", "台中", "台南", "高雄"], allLocations = [], selectedTripLegs = [];
        const fromSelect = modalContainer.querySelector('#ta_fromLocationSelect'), toSelect = modalContainer.querySelector('#ta_toLocationSelect'), legsUL = modalContainer.querySelector('#ta_currentLegsList'),
              tripReasonIn = modalContainer.querySelector('#ta_tripReason'), chargeSelect = modalContainer.querySelector('#ta_chargeable'), newLocIn = modalContainer.querySelector('#ta_newLocation'),
              previewBox = modalContainer.querySelector('#ta_emailPreviewContainer'), previewContent = modalContainer.querySelector('#ta_emailPreviewContent'),
              linkBox = modalContainer.querySelector('#ta_emailLinkContainer'), linkTextP = modalContainer.querySelector('#ta_emailLinkText'), mailtoA = modalContainer.querySelector('#ta_mailtoLink'),
              calGrid = modalContainer.querySelector('#ta_calendarGrid'), monthYearDisp = modalContainer.querySelector('#ta_monthYearDisplay');
        tripReasonIn.value = `Travel for project: ${pagePjNum}\nProject Name: ${pagePjName}\nProject Scope: ${pagePjScope}`;
        let curDate = new Date(), curMonth = curDate.getMonth(), curYear = curDate.getFullYear();
        const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"], storageKey = 'customTripAllLocations_v4_7_2'; // Ensure storageKey is unique if structure changes
        function loadLocs(){const s=localStorage.getItem(storageKey);allLocations=s?JSON.parse(s):[...defaultLocations];if(!allLocations.includes("北投")){allLocations.unshift("北投");allLocations=[...new Set(allLocations)];}allLocations.sort((a,b)=>(a==="北投")?-1:(b==="北投")?1:a.localeCompare(b));}
        function saveLocs(){localStorage.setItem(storageKey,JSON.stringify(allLocations));}
        function addCustomLoc(){const v=newLocIn.value.trim();if(!v){alert('Enter loc.');return;}if(allLocations.some(l=>l.toLowerCase()===v.toLowerCase())){alert(`"${v}" exists.`);return;}allLocations.push(v);loadLocs();saveLocs();newLocIn.value='';alert(`"${v}" added.`);popLocDd(fromSelect,"北投");popLocDd(toSelect);}
        function resetLocs(){if(confirm("Reset locs & clear legs?")){allLocations=[...defaultLocations];saveLocs();selectedTripLegs=[];renderLegs();popLocDd(fromSelect,"北投");popLocDd(toSelect);alert("Locs reset.");}}
        function popLocDd(sel,def=""){sel.innerHTML='';if(!def||!allLocations.includes(def))sel.add(new Option("-- Select --",""));allLocations.forEach(l=>sel.add(new Option(l,l)));if(def&&allLocations.includes(def))sel.value=def;}
        function addLeg(){const f=fromSelect.value,t=toSelect.value;if(!f||!t){alert("Select From & To.");return;}if(f===t){alert("From & To same.");return;}selectedTripLegs.push({from:f,to:t});renderLegs();fromSelect.value="北投";toSelect.value="";}
        function remLeg(idx){selectedTripLegs.splice(idx,1);renderLegs();}
        function renderLegs(){legsUL.innerHTML='';if(!selectedTripLegs.length){legsUL.innerHTML='<li>No legs.</li>';return;}selectedTripLegs.forEach((lg,i)=>{const li=document.createElement('li');li.textContent=`${lg.from} → ${lg.to}`;const b=document.createElement('button');b.textContent='X';b.className='ta-remove-leg-btn';b.onclick=()=>remLeg(i);li.appendChild(b);legsUL.appendChild(li);});}
        function renderCal(m,y){calGrid.innerHTML='';monthYearDisp.textContent=`${monthNames[m]} ${y}`;const fD=new Date(y,m,1).getDay(),dIM=new Date(y,m+1,0).getDate();const tD=new Date(),ty=tD.getFullYear(),tm=tD.getMonth(),td=tD.getDate();['S','M','T','W','T','F','S'].forEach(n=>{const c=document.createElement('div');c.className='ta-day-name';c.textContent=n;calGrid.appendChild(c);});for(let i=0;i<fD;i++){const c=document.createElement('div');c.className='ta-empty-day';calGrid.appendChild(c);}for(let d=1;d<=dIM;d++){const cl=document.createElement('div');cl.className='ta-calendar-day';cl.textContent=d;const dS=`${y}-${String(m+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;cl.dataset.date=dS;if(selectedDates.includes(dS))cl.classList.add('ta-selected-day');if(y===ty&&m===tm&&d===td)cl.classList.add('ta-today');cl.onclick=()=>toggleDate(dS,cl);calGrid.appendChild(cl);}}
        function toggleDate(dS,cl){const ix=selectedDates.indexOf(dS);if(ix>-1){selectedDates.splice(ix,1);cl.classList.remove('ta-selected-day');}else{selectedDates.push(dS);cl.classList.add('ta-selected-day');}selectedDates.sort();}
        modalContainer.querySelector('#ta_prevMonthBtn').onclick=()=>{curMonth--;if(curMonth<0){curMonth=11;curYear--;}renderCal(curMonth,curYear);};
        modalContainer.querySelector('#ta_nextMonthBtn').onclick=()=>{curMonth++;if(curMonth>11){curMonth=0;curYear++;}renderCal(curMonth,curYear);};
        modalContainer.querySelector('#ta_addCustomLocationBtn').onclick=addCustomLoc; modalContainer.querySelector('#ta_resetLocationsBtn').onclick=resetLocs; modalContainer.querySelector('#ta_addLegToListBtn').onclick=addLeg;
        function genEmailTxt(){const r=tripReasonIn.value.trim(),ch=chargeSelect.options[chargeSelect.selectedIndex];let sP=[];(pagePjNum&&pagePjNum!=="N/A")?sP.push(`Travel Approval for pj.${pagePjNum}`):sP.push("Business Trip Request");if(selectedDates.length){let dD;if(selectedDates.length===1)dD=selectedDates[0];else if(selectedDates.length<=3)dD=selectedDates.join(", ");else dD=selectedDates[0]+" etc.";sP.push(dD);}if(selectedTripLegs.length){if(selectedTripLegs.length===1)sP.push(`From ${selectedTripLegs[0].from} to ${selectedTripLegs[0].to}`);else sP.push("Multiple Legs");}const subj=sP.join(" - ");let body=`Dear Manager,\n\nI would like to apply for a business trip. Details:\n\nReason:\n${r}\n\nDates:\n`;selectedDates.length?selectedDates.forEach(d=>body+=`- ${d}\n`):body+="- (No dates)\n";body+="\nLocations/Legs:\n";selectedTripLegs.length?selectedTripLegs.forEach(l=>body+=`From: ${l.from}. To:   ${l.to}\n`):body+="- (No legs)\n";let accT="Charge & Account Allocation:";if(ch.value==="Billable-Inv")accT+=" (Billable Cust Exp):";else if(ch.value==="Billable-NonInv")accT+=" (Other Cust Exp):";body+=`\n${accT}\n- ${ch.text}\n\nPlease approve.\n\nThank you!\n${pageProjectHandlerEmail}`;return{subject:subj,body:body};}
        modalContainer.querySelector('#ta_generateEmailBtn').onclick=()=>{if(!selectedDates.length){alert('Select date(s)!');return;}if(!selectedTripLegs.length){alert('Add leg(s)!');return;}if(!tripReasonIn.value.trim()){alert('Enter reason!');return;}const e=genEmailTxt();previewContent.textContent=`Subject: ${e.subject}\n------------------\nBody:\n${e.body}`;previewBox.style.display='block';const m=`mailto:?subject=${encodeURIComponent(e.subject)}&body=${encodeURIComponent(e.body)}`;linkTextP.textContent=m;mailtoA.href=m;linkBox.style.display='block';window.open(m,'_blank');};
        modalContainer.querySelector('#ta_downloadTextFileBtn').onclick=()=>{if(!selectedDates.length&&!selectedTripLegs.length&&!tripReasonIn.value.trim()){alert('Enter details!');return;}const e=genEmailTxt(),r=tripReasonIn.value.trim(),ch=chargeSelect.options[chargeSelect.selectedIndex];let c=`Business Trip Info:\n\nSubject: ${e.subject}\n\nReason: ${r||'(N/A)'}\n\nDates:\n`;selectedDates.length?selectedDates.forEach(d=>c+=`- ${d}\n`):c+="- (None)\n";c+="\nLocations/Legs:\n";selectedTripLegs.length?selectedTripLegs.forEach(l=>c+=`- From: ${l.from}\n  To:   ${l.to}\n`):c+="- (None)\n";let accT="Charge & Account:";if(ch.value==="Billable-Inv")accT+=" (Billable Cust Exp):";else if(ch.value==="Billable-NonInv")accT+=" (Other Cust Exp):";c+=`\n${accT}\n- ${ch.text}\n`;const blb=new Blob([c],{type:'text/plain;charset=utf-8'}),lnk=document.createElement('a');lnk.href=URL.createObjectURL(blb);lnk.download='Trip_Info.txt';document.body.appendChild(lnk);lnk.click();document.body.removeChild(lnk);URL.revokeObjectURL(lnk.href);};
        loadLocs();popLocDd(fromSelect,"北投");popLocDd(toSelect);renderCal(curMonth,curYear);renderLegs();
    }

    function extractClientInfo() {
        const elements = document.querySelectorAll('.div-product-attribute');
        let targetElement = null;

        for (const el of elements) {
            if (el.textContent.includes('Customer Company Contact')) {
                targetElement = el;
                break;
            }
        }

        const data = {
            found: !!targetElement,
            clientName: null,
            clientEmail: null,
            clientPhone: null
        };

        if (targetElement) {
            const nameElement = targetElement.querySelector('.display-field-row');
            if (nameElement) {
                data.clientName = nameElement.textContent.trim();
            }

            const emailElement = Array.from(targetElement.querySelectorAll('.display-field-row')).find(el => el.textContent.includes('Email:'));
            if (emailElement) {
                data.clientEmail = emailElement.textContent.replace('Email:', '').trim();
            }

            const phoneElement = Array.from(targetElement.querySelectorAll('.display-field-row')).find(el => el.textContent.includes('Phone:'));
            if (phoneElement) {
                data.clientPhone = phoneElement.textContent.replace('Phone:', '').trim();
            }
        }
        // Return "N/A" explicitly if null or empty, consistent with other extractions
        return {
            clientName: data.clientName || "N/A",
            clientEmail: data.clientEmail || "N/A"
        };
    }

    GM_addStyle(`
        /* Merged Utility Panel Styles */
        #mergedUtilityPanel {
            position: fixed; bottom: 10px; right: 10px; background-color: #f0f0f0;
            border: 1px solid #bababa; padding: 0; z-index: 9999;
            width: 280px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            font-family: Segoe UI, Arial, sans-serif; font-size: 11px;
            border-radius: 3px; display: none;
        }
        #mrg-header {
            padding: 4px 8px; cursor: pointer; background-color: #e0e0e0; color: #222;
            border-bottom: 1px solid #c5c5c5; border-top-left-radius: 3px; border-top-right-radius: 3px;
            display: flex; justify-content: space-between; align-items: center; height: 22px; box-sizing: border-box;
        }
        #mrg-header-title { margin: 0; font-size: 1.1em; font-weight: bold; line-height: 1; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; user-select: none; }
        #mrg-close-btn { font-size: 16px; font-weight: bold; color: #555; cursor: pointer; padding: 0 3px; line-height: 1; user-select: none; }
        #mrg-close-btn:hover { color: #000; }
        #mergedUtilityPanel.mrg-collapsed #mrg-content-wrapper { display: none; }
        #mergedUtilityPanel.mrg-collapsed #mrg-header { border-bottom: none; border-bottom-left-radius: 3px; border-bottom-right-radius: 3px; }
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
        #mergedUtilityPanel.mrg-collapsed #mrg-global-status { display: none; }

        /* Travel Approval Modal Styles */
        #travelApprovalModalContainer {
            display: none; position: fixed; z-index: 10000;
            left: 0; top: 0; width: 100%; height: 100%;
            background-color: rgba(0,0,0,0.4);
            justify-content: center; align-items: center;
            font-family: Segoe UI, Arial, sans-serif;
            padding: 20px; box-sizing: border-box;
        }
        #travelApprovalModalContainer .ta-container {
            max-width: 550px; width: 100%;
            background: #f0f0f0; border: 1px solid #bababa; border-radius: 3px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.15);
            padding: 0; overflow-y: auto; max-height: calc(100vh - 40px);
            position: relative; color: #333; font-size: 12px;
        }
        #travelApprovalModalContainer .ta-custom-modal-close-btn {
            position: absolute; top: 5px; right: 8px;
            font-size: 18px; font-weight: bold; color: #555;
            cursor: pointer; line-height: 1; z-index: 10;
        }
        #travelApprovalModalContainer .ta-custom-modal-close-btn:hover { color: #000; }
        #travelApprovalModalContainer .ta-header-title {
            font-size: 1.1em; font-weight: bold; color: #222;
            background-color: #e0e0e0; padding: 6px 10px; margin: 0;
            border-bottom: 1px solid #c5c5c5; position: sticky; top: 0; z-index: 5;
        }
        #travelApprovalModalContainer .ta-main-layout-wrapper { padding: 10px; }
        #travelApprovalModalContainer .ta-main-content { width: 100%; }
        #travelApprovalModalContainer label { display: block; margin-bottom: 4px; font-weight: bold; font-size: 1em; color: #444; }
        #travelApprovalModalContainer input[type="text"],
        #travelApprovalModalContainer textarea,
        #travelApprovalModalContainer select {
            background-color: #fff; border: 1px solid #c0c0c0; border-radius: 2px;
            padding: 4px 6px; font-size: 12px; color: #333;
            width: 100%; box-sizing: border-box; margin-bottom: 10px;
        }
        #travelApprovalModalContainer input[type="text"]:focus,
        #travelApprovalModalContainer textarea:focus,
        #travelApprovalModalContainer select:focus { border-color: #673ab7; box-shadow: 0 0 0 1px #673ab7; outline: none; }
        #travelApprovalModalContainer #ta_tripReason { min-height: 120px; resize: vertical; }
        #travelApprovalModalContainer .ta-action-button-small,
        #travelApprovalModalContainer .ta-cal-nav-btn {
            font-weight: bold; font-size: 1em; line-height: 1.2em;
            padding: 3px 7px; border: 1px solid #aeaeae; color: #333333;
            background-color: #f9f9f9; cursor: pointer; border-radius: 2px;
            text-align: center; white-space: nowrap; margin-bottom: 5px;
        }
        #travelApprovalModalContainer .ta-action-button-small:hover,
        #travelApprovalModalContainer .ta-cal-nav-btn:hover { background-color: #e8e8e8; border-color: #999999; }
        #travelApprovalModalContainer .ta-action-button-primary { background-color: #673ab7; color: white; border-color: #5e35b1; }
        #travelApprovalModalContainer .ta-action-button-primary:hover { background-color: #5e35b1; }
        #travelApprovalModalContainer .ta-action-buttons button { width: 97%; } /* As per your v4.6 */
        #travelApprovalModalContainer .ta-add-location-container { display: flex; gap: 5px; align-items: center; margin-bottom: 10px; }
        #travelApprovalModalContainer .ta-add-location-container input { flex-grow: 1; margin-bottom: 0; }
        #travelApprovalModalContainer .ta-add-location-container button { width: auto; margin-left: 0px; }
        #travelApprovalModalContainer .ta-location-pair-selection button { width: auto; margin-left: 5px; }
        #travelApprovalModalContainer .ta-list { list-style: none; padding: 5px; margin: 5px 0 10px 0; overflow-y: auto; border: 1px solid #d0d0d0; border-radius: 2px; background-color: #f9f9f9; max-height: 100px; }
        #travelApprovalModalContainer .ta-list li { background: #fff; border-bottom: 1px solid #e0e0e0; padding: 3px 5px; margin-bottom: 3px; border-radius: 2px; display: flex; justify-content: space-between; align-items: center; font-size: 1em; }
        #travelApprovalModalContainer .ta-list li:last-child { margin-bottom: 0; border-bottom: none; }
        #travelApprovalModalContainer .ta-remove-leg-btn { background-color: transparent; color: #cc0000; border: none; font-size: 1em; font-weight: bold; padding: 0 4px; margin-left: 5px; cursor: pointer; border-radius: 2px; line-height: 1; }
        #travelApprovalModalContainer .ta-remove-leg-btn:hover { background-color: #f0f0f0; }
        #travelApprovalModalContainer .ta-trip-leg-embedded-section { padding: 10px; border: 1px solid #d0d0d0; border-radius: 2px; margin-top: 15px; margin-bottom: 15px; background-color: #e9e9e9; }
        #travelApprovalModalContainer .ta-section-title-inner { margin-top: 0; margin-bottom: 8px; font-size: 1em; font-weight: bold; color: #333; padding-bottom: 2px; border-bottom: 1px solid #c5c5c5; }
        #travelApprovalModalContainer .ta-links-above-legs { margin-bottom: 8px; }
        #travelApprovalModalContainer .ta-links-above-legs a { margin-right: 5px; font-size: 1.2em; }
        #travelApprovalModalContainer .ta-location-pair-selection { display: flex; gap: 10px; margin-bottom: 10px; align-items: flex-end; }
        #travelApprovalModalContainer .ta-location-pair-selection > div { flex: 1; }
        #travelApprovalModalContainer .ta-preview-box { margin-top: 15px; background-color: #e9e9e9; border: 1px solid #d0d0d0; border-radius: 2px; padding: 8px; }
        #travelApprovalModalContainer #ta_emailPreviewContent { background-color: #fff; border: 1px solid #c0c0c0; padding: 5px; white-space: pre-wrap; word-wrap: break-word; font-family: Consolas, monospace; font-size: 10px; max-height: 150px; overflow-y: auto; border-radius: 2px; }
        #travelApprovalModalContainer #ta_emailLinkContainer a { color: #0066cc; text-decoration: none; }
        #travelApprovalModalContainer #ta_emailLinkContainer a:hover { text-decoration: underline; }
        #travelApprovalModalContainer .ta-calendar-container { margin-bottom: 15px; border: 1px solid #d0d0d0; border-radius: 2px; padding: 8px; background-color: #fff; }
        #travelApprovalModalContainer .ta-calendar-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; }
        #travelApprovalModalContainer .ta-calendar-header h3 { font-size: 1em; font-weight: bold; color: #333; margin: 0; }
        #travelApprovalModalContainer .ta-cal-nav-btn { padding: 2px 5px; margin: 0 3px; }
        #travelApprovalModalContainer .ta-calendar-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 2px; text-align: center; }
        #travelApprovalModalContainer .ta-calendar-grid div { padding: 3px; border: 1px solid transparent; border-radius: 2px; cursor: pointer; font-size: 0.9em; line-height: 1.5em; display: flex; align-items: center; justify-content: center; box-sizing: border-box; }
        #travelApprovalModalContainer .ta-day-name { font-weight: bold; background-color: transparent; color: #555; cursor: default; border-radius: 0; padding: 2px;}
        #travelApprovalModalContainer .ta-empty-day { background-color: transparent; cursor: default; border-color: transparent; }
        #travelApprovalModalContainer .ta-calendar-day:hover { background-color: #e8e8e8; border-color: #ccc; }
        #travelApprovalModalContainer .ta-selected-day { background-color: #673ab7; color: white; font-weight: bold; border-color: #5e35b1; }
        #travelApprovalModalContainer .ta-selected-day:hover { background-color: #5e35b1; border-color: #502d9a; }
        #travelApprovalModalContainer .ta-calendar-day.ta-today:not(.ta-selected-day) { color: #cc0000; font-weight: bold; border: 1px dashed #cc0000; }
    `);

    function init() {
        if (document.querySelector('dt') || document.querySelector('.display-label-row')) {
            if (!document.getElementById('mergedUtilityPanel')) createUI();
            const panel = document.getElementById('mergedUtilityPanel');
            if (panel) panel.style.display = 'block';
            console.log("Merged Utility Panel (v4.7.3): UI created."); // MODIFIED: Updated log version
        } else {
            console.log("Merged Utility Panel: Required page structure not found.");
        }
    }
    // Ensure the script waits for the page to be sufficiently loaded.
    // Using a simple timeout or DOMContentLoaded might be more robust if init() runs too early.
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
