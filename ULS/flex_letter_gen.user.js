// ==UserScript==
// @name         Letter Generator (Inactive, ECD, TAT, NOA, Travel)
// @name:zh-CN   信件生成器 (固定,可縮放,差旅審批)
// @namespace    http://tampermonkey.net/
// @version      4.7.1
// @description  Merges Inactive Letter Generator, TAT/ECD/NOA Letter buttons, and adds a Travel Approval form into a single, fixed, collapsible floating panel.
// @description:zh-CN 將 Inactive Letter 生成器、TAT/ECD/NOA Letter 按鈕以及差旅申請表單合併到單個固定的、可縮放的浮動面板中。
// @author       Your Name
// @match        https://portal.ul.com/Project/Details/*
// @grant        GM_addStyle
// @grant        GM_openInTab
// ==/UserScript==

(function() {
    'use strict';

    // --- Constants ---
    const RBG_TARGET_ANCHOR_SELECTOR = '.section-crumbs-li a';
    const RBG_ARROW_CHAR = ' ›';
    const RBG_BASE_REPORT_URL = 'https://epic.ul.com/Report';

    // --- Helper Functions ---
    function extractFieldByLabel(labelText) { // This remains for other fields
        const xpath = `//div[@class='display-label-row' and normalize-space(.)='${labelText}']/following-sibling::div[@class='display-field-row'][1]`;
        const fieldElement = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
        return fieldElement ? fieldElement.textContent.trim() : '';
    }

    // NEW specific function for Project Name using provided classes
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
    const ilgEmailTemplates = [ // Templates as in your v4.6 (shortened for this example)
        { id: 'notice', name: 'Notice Inactive Letter', titleTemplate: "Project Inactive Letter–Project #PjNum#/關於UL項目#PjNum#暫停通知書", contentTemplate: `尊敬的客戶：\n\n... #PjNum# ... #OdrNum# ... #PjScope# ... #Project Hold Reason# ...` },
        { id: 'inactive1', name: 'Inactive Letter 1', titleTemplate: "The 1st project inactive follow up letter–Project #PjNum#/第一次項目暫停跟進通知書", contentTemplate: `尊敬的客戶：\n\n... #PjNum# ... #OdrNum# ... #PjScope# ... #Project Hold Reason# ...` },
        { id: 'inactive2', name: 'Inactive Letter 2', titleTemplate: "The 2nd project inactive follow up letter–Project #PjNum#/第二次項目暫停跟進通知書", contentTemplate: `尊敬的客戶：\n\n... #PjNum# ... #OdrNum# ... #PjScope# ... #Project Hold Reason# ...` },
        { id: 'inactive3', name: 'Inactive Letter 3', titleTemplate: "The 3nd project inactive follow up letter–Project #PjNum#/第三次項目暫停跟進通知書", contentTemplate: `尊敬的客戶：\n\n... #PjNum# ... #OdrNum# ... #PjScope# ... #Project Hold Reason# ...` },
        { id: 'final', name: 'Inactive Letter Final', titleTemplate: "The final notice before project close by letter – Project #PjNum#/項目終止前最後提醒", contentTemplate: `尊敬的客戶：\n... #PjNum# ... #OdrNum# ... #PjScope# ... #Project Hold Reason# ... #DeadlineDate# ...` }
    ];
    const rbgButtonsData = [
        { id: 'ecdLetterFloatingButton', text: 'ECD Letter', params: { templateUNID: 'AHL ECD Letter', selectedOutputType: '.eml', addFRDate: false }},
        { id: 'tatLetterFloatingButton', text: 'TAT Letter', params: { templateUNID: 'AHL TAT Letter', selectedOutputType: '.eml', addFRDate: false }},
        { id: 'noaLetterFloatingButton', text: 'NOA Letter', params: { templateUNID: 'Notice of Authorization or Completion Letter', selectedOutputType: '.default', addFRDate: true }},
        { id: 'travelApprovalFloatingButton', text: 'Travel Approval', params: {}}
    ];

    let currentPageDataForTravelModal = { pjNum: "N/A", pjName: "N/A" };

    function ilgGatherAllPageData() {
        const pjNum = extractFieldByLabel('Oracle Project Number') || "N/A";
        const pjName = extractProjectName() || "N/A"; // Use the new specific function
        const pjScope = extractFieldByLabel('Project Scope') || "N/A";

        currentPageDataForTravelModal = { pjNum, pjName };

        return {
            pjNum,
            pjName,
            pjScope,
            odrNum: extractOdrNum() || "N/A",
            rocDateBooked: formatToROCYearMonth(extractFieldByLabel('Date Booked')) || "未知日期",
            projectHoldReason: extractFieldByLabel('Project Hold Reason') || "N/A",
            deadlineDate: getFutureROCDate(14),
            projectHandlerEmail: extractProjectHandlerEmail() || "N/A"
        };
    }

    // ... createUI function remains the same as your v4.6 ...
    function createUI() {
        ilgGatherAllPageData();
        const panel = document.createElement('div');
        panel.id = 'mergedUtilityPanel';
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
        if (!document.querySelector(RBG_TARGET_ANCHOR_SELECTOR) && rbgButtonsData.some(b => b.id !== 'travelApprovalFloatingButton')) { console.warn(`RBG Target '${RBG_TARGET_ANCHOR_SELECTOR}' not found.`); showGlobalStatus(`Warn: RBG target missing.`,1,5e3); }
    }


    // ... showGlobalStatus function remains the same ...
    function showGlobalStatus(msg, isErr=0, dur=null) { const el=document.getElementById('mrg-global-status'); if(el){el.textContent=msg; el.style.color=isErr?'red':'#006400'; if(el.timeoutId)clearTimeout(el.timeoutId); el.timeoutId=setTimeout(()=>{if(el.textContent===msg)el.textContent='';el.timeoutId=null;},dur??(isErr?5e3:3e3));}}

    // ... ilgGenerateAndDisplay function remains the same (it uses pjScope from pageData) ...
    function ilgGenerateAndDisplay(templateId) {
        const d = ilgGatherAllPageData(), t = ilgEmailTemplates.find(x=>x.id===templateId);
        if (!t) { console.error('ILG Err: Tpl not found:',templateId); showGlobalStatus('ILG Err: Tpl miss!',1); return; }
        let title = t.titleTemplate.replace(/#PjNum#/g,d.pjNum);
        let content = t.contentTemplate.replace(/#PjNum#/g,d.pjNum).replace(/#OdrNum#/g,d.odrNum).replace(/xxx年xx月/g,d.rocDateBooked).replace(/#PjScope#/g,d.pjScope).replace(/#Project Hold Reason#/g,d.projectHoldReason).replace(/^(Project Handler \/Email:)\s*$/gm,`$1 ${d.projectHandlerEmail}`).replace(/(Project Handler \/Email:)(?!.*\S)/gm,`$1 ${d.projectHandlerEmail}`);
        if (t.id==='final') content=content.replace(/#DeadlineDate#/g,d.deadlineDate);
        document.getElementById('ilg-title-output').value=title; document.getElementById('ilg-content-output').value=content;
        document.getElementById('ilg-output-area').style.display='block'; showGlobalStatus(`${ilgButtonLabels[t.id]||t.name} gen.`);
    }


    // ... rbgCreateReportButton function remains the same (passes pjNum and pjName to initializeTravelApprovalModalLogic) ...
    let travelApprovalModalCreated = false;
    function rbgCreateReportButton(buttonConfig) {
        const button = document.createElement('button'); button.id = buttonConfig.id; button.className = 'mrg-action-button rbg-button';
        button.textContent = buttonConfig.text + (buttonConfig.id !== 'travelApprovalFloatingButton' ? RBG_ARROW_CHAR : '');
        if (buttonConfig.id === 'travelApprovalFloatingButton') {
            button.onclick = function() {
                ilgGatherAllPageData();
                if (!travelApprovalModalCreated) { createTravelApprovalModal(); travelApprovalModalCreated = true; }
                document.getElementById('travelApprovalModalContainer').style.display = 'flex';
                initializeTravelApprovalModalLogic(currentPageDataForTravelModal.pjNum, currentPageDataForTravelModal.pjName); // Pass pjName
                showGlobalStatus('Travel Approval open.',0,1500);
            };
        } else {
            button.onclick = function() {
                const anchor = document.querySelector(RBG_TARGET_ANCHOR_SELECTOR);
                if (!anchor || !anchor.href) { alert(`RBG Err: Target anchor/href missing (${RBG_TARGET_ANCHOR_SELECTOR}).`); showGlobalStatus(`RBG Err: Target missing.`,1); return; }
                // const p = new URLSearchParams({ TemplateUNID:buttonConfig.params.templateUNID, SelectedOutputType:buttonConfig.params.selectedOutputType, ProjectID:encodeURIComponent(anchor.href), isWorkbench:'False' });
                const p = new URLSearchParams({ TemplateUNID:buttonConfig.params.templateUNID, SelectedOutputType:buttonConfig.params.selectedOutputType, ProjectID:anchor.href, isWorkbench:'False' });
                if (buttonConfig.params.addFRDate) p.append('FRDate',rbgGetFormattedDate());
                GM_openInTab(`${RBG_BASE_REPORT_URL}?${p.toString()}`,{active:true}); showGlobalStatus(`Opening ${buttonConfig.text}...`,0,1500);
            };
        }
        return button;
    }


    // ... travelApprovalModalHTMLBodyContent remains the same as your v4.6 ...
    const travelApprovalModalHTMLBodyContent = `
        <div class="ta-container">
            <span id="travelApprovalModalCloseButton" class="ta-custom-modal-close-btn">×</span>
            <h2 class="ta-header-title">Travel Approval Generator</h2>
            <div class="ta-main-layout-wrapper">
                <div class="ta-main-content">
                    <div>
                        <label for="ta_tripReason">Trip Reason:</label>
                        <textarea id="ta_tripReason"></textarea>
                    </div>
                    <div>
                        <label for="ta_chargeable">Travel category & Account Allocation:</label>
                        <select id="ta_chargeable">
                            <option value="None">None</option>
                            <option value="Non-Billable">Non-Billable</option>
                            <option value="Billable-Inv">Billable to project and Invoiceable</option>
                            <option value="Billable-NonInv">Billable to project and Non-Invoiceable</option>
                        </select>
                    </div>
                    <label>Select Dates (Click to select/deselect):</label>
                    <div class="ta-calendar-container">
                        <div class="ta-calendar-header">
                            <button id="ta_prevMonthBtn" class="ta-cal-nav-btn">◀</button>
                            <h3 id="ta_monthYearDisplay"></h3>
                            <button id="ta_nextMonthBtn" class="ta-cal-nav-btn">▶</button>
                        </div>
                        <div class="ta-calendar-grid" id="ta_calendarGrid"></div>
                    </div>

                    <div class="ta-links-above-legs">
                             <a href="https://www.uber.com/global/zh-tw/price-estimate/" target="_blank" rel="noopener noreferrer">[Uber]</a>
                             <a href="https://www.thsrc.com.tw/ArticleContent/a3b630bb-1066-4352-a1ef-58c7b4e8ef7c" target="_blank" rel="noopener noreferrer">[HSR]</a>
                             <a href="https://www.metro.taipei/cp.aspx?n=ECEADC266D7120A7" target="_blank" rel="noopener noreferrer">[MRT]</a>
                             <a href="https://www.mtaxi.com.tw/taxi-fare-estimate/" target="_blank" rel="noopener noreferrer">[mtaxi]</a>
                    </div>
                    <div class="ta-trip-leg-embedded-section">
                        <label for="ta_newLocation">Add Custom Location (for From/To dropdowns):</label>
                        <div class="ta-add-location-container">
                            <input type="text" id="ta_newLocation" placeholder="Enter new location name">
                            <button id="ta_addCustomLocationBtn" class="ta-action-button-small">Add Location</button>
                            <button id="ta_resetLocationsBtn" class="ta-action-button-secondary ta-action-button-small">Reset Locations</button>
                        </div>
                        <h3 class="ta-section-title-inner">Add/Modify Trip Legs</h3>
                        <div class="ta-location-pair-selection">
                            <div>
                                <label for="ta_fromLocationSelect">From:</label>
                                <select id="ta_fromLocationSelect"></select>
                            </div>
                            <div>
                                <label for="ta_toLocationSelect">To:</label>
                                <select id="ta_toLocationSelect"></select>
                            </div>
                            <button id="ta_addLegToListBtn" class="ta-action-button-primary ta-action-button-small">Add Leg</button>
                        </div>
                        <label>Current Legs:</label>
                        <ul id="ta_currentLegsList" class="ta-list"></ul>
                    </div>

                    <div class="ta-action-buttons">
                        <button id="ta_generateEmailBtn" class="ta-action-button-primary ta-action-button-small"><b>Generate & Open Email</b></button>
                        <button id="ta_downloadTextFileBtn" class="ta-action-button-secondary ta-action-button-small">Download Details</button>
                    </div>
                </div>
            </div>
            <div id="ta_emailPreviewContainer" class="ta-preview-box" style="display: none;">
                <h3 class="ta-section-title-inner">Email Preview (Plain Text):</h3>
                <pre id="ta_emailPreviewContent"></pre>
            </div>
            <div id="ta_emailLinkContainer" class="ta-preview-box" style="display: none;">
                <h3 class="ta-section-title-inner">Email Link:</h3>
                <p id="ta_emailLinkText"></p>
                <a id="ta_mailtoLink" href="#" target="_blank">Click here to open email manually</a>
            </div>
        </div>
    `;

    // ... createTravelApprovalModal function remains the same ...
    function createTravelApprovalModal() {
        const modalContainer = document.createElement('div');
        modalContainer.id = 'travelApprovalModalContainer';
        modalContainer.innerHTML = travelApprovalModalHTMLBodyContent;
        document.body.appendChild(modalContainer);
        modalContainer.querySelector('#travelApprovalModalCloseButton').onclick = () => { modalContainer.style.display = 'none'; };
        modalContainer.onclick = (e) => { if (e.target === modalContainer) modalContainer.style.display = 'none'; };
    }


    // --- initializeTravelApprovalModalLogic: Updated to use pagePjName for trip reason ---
    function initializeTravelApprovalModalLogic(pagePjNum = "N/A", pagePjName = "N/A") { // Changed second param to pagePjName
        const modalContainer = document.getElementById('travelApprovalModalContainer');
        if (!modalContainer) return;

        let selectedDates = [];
        const defaultLocations = ["北投", "關渡賓士大樓", "群通大樓", "新店", "新莊", "桃園", "新竹", "台中", "高雄"];
        let allLocations = [];
        let selectedTripLegs = [];

        const fromLocationSelect = modalContainer.querySelector('#ta_fromLocationSelect');
        const toLocationSelect = modalContainer.querySelector('#ta_toLocationSelect');
        const currentLegsListUL = modalContainer.querySelector('#ta_currentLegsList');
        const tripReasonInput = modalContainer.querySelector('#ta_tripReason');
        const chargeableSelect = modalContainer.querySelector('#ta_chargeable');
        const newLocationInput = modalContainer.querySelector('#ta_newLocation');
        const emailPreviewContainer = modalContainer.querySelector('#ta_emailPreviewContainer');
        const emailPreviewContent = modalContainer.querySelector('#ta_emailPreviewContent');
        const emailLinkContainer = modalContainer.querySelector('#ta_emailLinkContainer');
        const emailLinkTextP = modalContainer.querySelector('#ta_emailLinkText');
        const mailtoLinkA = modalContainer.querySelector('#ta_mailtoLink');
        const calendarGrid = modalContainer.querySelector('#ta_calendarGrid');
        const monthYearDisplay = modalContainer.querySelector('#ta_monthYearDisplay');

        // Set default Trip Reason using Project Name
        tripReasonInput.value = `Travel for project: ${pagePjNum}\nProject Name: ${pagePjName}`;

        let currentDate = new Date(), currentMonth = currentDate.getMonth(), currentYear = currentDate.getFullYear();
        const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        const localStorageLocationsKey = 'customTripAllLocations_v4_7_1'; // New key for this version

        function loadAllLocations() {
            const saved = localStorage.getItem(localStorageLocationsKey);
            allLocations = saved ? JSON.parse(saved) : [...defaultLocations];
            if (!allLocations.includes("北投")) { allLocations.unshift("北投"); allLocations = [...new Set(allLocations)]; }
            allLocations.sort((a,b) => (a==="北投")?-1:(b==="北投")?1:a.localeCompare(b));
        }
        function saveAllLocations() { localStorage.setItem(localStorageLocationsKey, JSON.stringify(allLocations)); }
        function addCustomLocation() {
            const val = newLocationInput.value.trim();
            if (!val) { alert('Enter location.'); return; }
            if (allLocations.some(l=>l.toLowerCase()===val.toLowerCase())) { alert(`"${val}" exists.`); return; }
            allLocations.push(val); loadAllLocations(); saveAllLocations(); newLocationInput.value=''; alert(`"${val}" added.`);
            populateLocationDropdown(fromLocationSelect, "北投"); populateLocationDropdown(toLocationSelect);
        }
        function resetLocations() {
           if (confirm("Reset locations and clear legs?")) {
               allLocations=[...defaultLocations]; saveAllLocations(); selectedTripLegs=[];
               renderCurrentLegsList();
               populateLocationDropdown(fromLocationSelect,"北投"); populateLocationDropdown(toLocationSelect);
               alert("Locations reset.");
           }
        }
        function populateLocationDropdown(sel, defVal="") {
            sel.innerHTML=''; if(!defVal||!allLocations.includes(defVal)) sel.add(new Option("-- Select --",""));
            allLocations.forEach(l=>sel.add(new Option(l,l)));
            if(defVal&&allLocations.includes(defVal)) sel.value=defVal;
        }
        function addLegToList() {
            const from=fromLocationSelect.value, to=toLocationSelect.value;
            if(!from||!to){alert("Select From & To.");return;} if(from===to){alert("From & To same.");return;}
            selectedTripLegs.push({from,to}); renderCurrentLegsList();
            fromLocationSelect.value="北投"; toLocationSelect.value="";
        }
        function removeLegFromList(idx) { selectedTripLegs.splice(idx,1); renderCurrentLegsList(); }
        function renderCurrentLegsList() {
            currentLegsListUL.innerHTML = '';
            if(!selectedTripLegs.length){currentLegsListUL.innerHTML='<li>No legs.</li>';return;}
            selectedTripLegs.forEach((leg,idx)=>{
                const li=document.createElement('li');li.textContent=`${leg.from} → ${leg.to}`;
                const btn=document.createElement('button');btn.textContent='X';btn.className='ta-remove-leg-btn';
                btn.onclick=()=>removeLegFromList(idx); li.appendChild(btn); currentLegsListUL.appendChild(li);
            });
        }
        function renderCalendar(month, year) {
             calendarGrid.innerHTML=''; monthYearDisplay.textContent=`${monthNames[month]} ${year}`;
             const firstD=new Date(year,month,1).getDay(), daysInM=new Date(year,month+1,0).getDate();
             const today=new Date(), ty=today.getFullYear(), tm=today.getMonth(), td=today.getDate();
             ['S','M','T','W','T','F','S'].forEach(n=>{const c=document.createElement('div');c.className='ta-day-name';c.textContent=n;calendarGrid.appendChild(c);});
             for(let i=0;i<firstD;i++){const c=document.createElement('div');c.className='ta-empty-day';calendarGrid.appendChild(c);}
             for(let d=1;d<=daysInM;d++){
                 const cell=document.createElement('div');cell.className='ta-calendar-day';cell.textContent=d;
                 const ds=`${year}-${String(month+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`; cell.dataset.date=ds;
                 if(selectedDates.includes(ds))cell.classList.add('ta-selected-day');
                 if(year===ty&&month===tm&&d===td)cell.classList.add('ta-today');
                 cell.onclick=()=>toggleDateSelection(ds,cell); calendarGrid.appendChild(cell);
             }
        }
        function toggleDateSelection(dateString, cell) {
             const idx=selectedDates.indexOf(dateString);
             if(idx>-1){selectedDates.splice(idx,1);cell.classList.remove('ta-selected-day');}
             else{selectedDates.push(dateString);cell.classList.add('ta-selected-day');}
             selectedDates.sort();
        }

        modalContainer.querySelector('#ta_prevMonthBtn').onclick = () => { currentMonth--; if(currentMonth<0){currentMonth=11;currentYear--;} renderCalendar(currentMonth,currentYear); };
        modalContainer.querySelector('#ta_nextMonthBtn').onclick = () => { currentMonth++; if(currentMonth>11){currentMonth=0;currentYear++;} renderCalendar(currentMonth,currentYear); };
        modalContainer.querySelector('#ta_addCustomLocationBtn').onclick = addCustomLocation;
        modalContainer.querySelector('#ta_resetLocationsBtn').onclick = resetLocations;
        modalContainer.querySelector('#ta_addLegToListBtn').onclick = addLegToList;

        function generateEmailText() {
             const reason=tripReasonInput.value.trim(); const chargeOpt=chargeableSelect.options[chargeableSelect.selectedIndex];
             let subjParts=[]; (pagePjNum&&pagePjNum!=="N/A")?subjParts.push(`Travel Approval for pj.${pagePjNum}`):subjParts.push("Business Trip Request");
             if(selectedDates.length){let dDisp; if(selectedDates.length===1)dDisp=selectedDates[0]; else if(selectedDates.length<=3)dDisp=selectedDates.join(", "); else dDisp=selectedDates[0]+" etc."; subjParts.push(dDisp);}
             if(selectedTripLegs.length){if(selectedTripLegs.length===1)subjParts.push(`From ${selectedTripLegs[0].from} to ${selectedTripLegs[0].to}`); else subjParts.push("Multiple Legs");}
             const subject=subjParts.join(" - ");
             let body=`Dear Manager,\n\nI would like to apply for a business trip. Details:\n\nReason:\n${reason}\n\nDates:\n`;
             selectedDates.length?selectedDates.forEach(d=>body+=`- ${d}\n`):body+="- (No dates)\n"; body+="\nLocations/Legs:\n";
             selectedTripLegs.length?selectedTripLegs.forEach(l=>body+=`- From: ${l.from}\n  To:   ${l.to}\n`):body+="- (No legs)\n";
             let accTitle="Charge & Account Allocation:"; if(chargeOpt.value==="Billable-Inv")accTitle+=" (Billable Cust Exp):"; else if(chargeOpt.value==="Billable-NonInv")accTitle+=" (Other Cust Exp):";
             body+=`\n${accTitle}\n- ${chargeOpt.text}\n\nPlease approve.\n\nThank you!`;
             return {subject,body};
        }
        modalContainer.querySelector('#ta_generateEmailBtn').onclick = () => {
            if(!selectedDates.length){alert('Select trip date(s)!');return;} if(!selectedTripLegs.length){alert('Add trip leg(s)!');return;}
            if(!tripReasonInput.value.trim()){alert('Enter trip reason!');return;}
            const email=generateEmailText();
            emailPreviewContent.textContent=`Subject: ${email.subject}\n------------------\nBody:\n${email.body}`; emailPreviewContainer.style.display='block';
            const mailto=`mailto:?subject=${encodeURIComponent(email.subject)}&body=${encodeURIComponent(email.body)}`;
            emailLinkTextP.textContent=mailto; mailtoLinkA.href=mailto; emailLinkContainer.style.display='block'; window.open(mailto,'_blank');
        };
        modalContainer.querySelector('#ta_downloadTextFileBtn').onclick = () => {
            if(!selectedDates.length&&!selectedTripLegs.length&&!tripReasonInput.value.trim()){alert('Enter details!');return;}
            const email=generateEmailText(), reason=tripReasonInput.value.trim(), chargeOpt=chargeableSelect.options[chargeableSelect.selectedIndex];
            let content=`Business Trip Info:\n\nSubject: ${email.subject}\n\nReason: ${reason||'(N/A)'}\n\nDates:\n`;
            selectedDates.length?selectedDates.forEach(d=>content+=`- ${d}\n`):content+="- (None)\n"; content+="\nLocations/Legs:\n";
            selectedTripLegs.length?selectedTripLegs.forEach(l=>content+=`- From: ${l.from}\n  To:   ${l.to}\n`):content+="- (None)\n";
            let accTitle="Charge & Account:"; if(chargeOpt.value==="Billable-Inv")accTitle+=" (Billable Cust Exp):"; else if(chargeOpt.value==="Billable-NonInv")accTitle+=" (Other Cust Exp):";
            content+=`\n${accTitle}\n- ${chargeOpt.text}\n`;
            const blob=new Blob([content],{type:'text/plain;charset=utf-8'}), link=document.createElement('a');
            link.href=URL.createObjectURL(blob); link.download='Trip_Info.txt'; document.body.appendChild(link); link.click(); document.body.removeChild(link); URL.revokeObjectURL(link.href);
        };

        loadAllLocations();
        populateLocationDropdown(fromLocationSelect,"北投"); populateLocationDropdown(toLocationSelect);
        renderCalendar(currentMonth,currentYear); renderCurrentLegsList();
    }

    // ... GM_addStyle section remains the same as your v4.6 ...
    GM_addStyle(`
        /* Merged Utility Panel Styles (Unchanged) */
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

        /* Travel Approval Modal Styles (Unchanged from your v4.6 - compact, no sidebar) */
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
            console.log("Merged Utility Panel (v4.7.1): UI created.");
        } else {
            console.log("Merged Utility Panel: Required page structure not found.");
        }
    }
    init();
})();
