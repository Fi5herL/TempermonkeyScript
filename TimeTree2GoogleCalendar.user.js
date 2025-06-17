// ==UserScript==
// @name         TimeTree to Google Calendar Sync (v3 with Color)
// @namespace    http://tampermonkey.net/
// @version      3.0
// @description  Adds a button to TimeTree to fully sync (add/delete/color) events with a Google Calendar for the current month.
// @author       YourName
// @match        https://timetreeapp.com/calendars/*
// @grant        GM_xmlhttpRequest
// @grant        GM_addStyle
// @connect      script.google.com
// ==/UserScript==

(function() {
    'use strict';

    // --- CONFIGURATION ---
    // Make sure your Google Apps Script URL is still here
    const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbx4fZd5YRFSTkwEg4geL-phTbU0_AJ8DlDorB3nJ3bTyg492IOxzCql_kcZIHyA6zBzUQ/exec';

    // --- MAIN LOGIC ---
    function addSyncButton() {
        const targetContainer = document.querySelector('div.css-1bpni0w');
        if (!targetContainer) {
            setTimeout(addSyncButton, 1000);
            return;
        }

        const syncButton = document.createElement('button');
        syncButton.id = 'sync-to-gcal-button';
        syncButton.textContent = 'Full Sync to Google Calendar';
        syncButton.addEventListener('click', handleSync);

        syncButton.title = 'Warning: This will add, delete, and update colors on Google Calendar to match this page!';
        targetContainer.appendChild(syncButton);
    }

    function handleSync() {
        if (!confirm('This action will add, delete, and update event colors on your Google Calendar to match the current TimeTree month.\n\nAre you sure you want to continue?')) {
            return;
        }

        const button = document.getElementById('sync-to-gcal-button');
        button.textContent = 'Analyzing...';
        button.disabled = true;

        const syncData = scrapeCalendarData();

        if (syncData.events.length === 0 && !syncData.viewStartDate) {
            button.textContent = 'No events found!';
            setTimeout(() => {
                button.textContent = 'Full Sync to Google Calendar';
                button.disabled = false;
            }, 3000);
            return;
        }

        button.textContent = `Sending ${syncData.events.length} events...`;

        GM_xmlhttpRequest({
            method: 'POST',
            url: APPS_SCRIPT_URL,
            data: JSON.stringify(syncData),
            headers: { 'Content-Type': 'application/json' },
            onload: function(response) {
                try {
                    const result = JSON.parse(response.responseText);
                    button.textContent = result.status === 'success' ? `✅ ${result.message}` : `❌ Error: ${result.message}`;
                } catch (e) {
                    button.textContent = '❌ Response parse error';
                }
                setTimeout(() => {
                    button.textContent = 'Full Sync to Google Calendar';
                    button.disabled = false;
                }, 10000);
            },
            onerror: function(error) {
                console.error('Sync Script Error:', error);
                button.textContent = '❌ Network Error!';
                setTimeout(() => {
                    button.textContent = 'Full Sync to Google Calendar';
                    button.disabled = false;
                }, 5000);
            }
        });
    }

    function scrapeCalendarData() {
        // --- Scrape the date range (no changes here) ---
        const dateCells = document.querySelectorAll('div[role="gridcell"]');
        if (dateCells.length < 2) return { events: [] };

        const firstDateElement = dateCells[0].querySelector('.css-g51b5d, .css-q2isom');
        const lastDateElement = dateCells[dateCells.length - 1].querySelector('.css-g51b5d, .css-q2isom');
        const timeElement = document.querySelector('time.css-e1a69x');
        if (!timeElement || !firstDateElement || !lastDateElement) return { events: [] };

        const [monthName, yearStr] = timeElement.textContent.split(', ');
        const year = parseInt(yearStr, 10);
        const month = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"].findIndex(m => m.toLowerCase() === monthName.toLowerCase());
        const firstDay = parseInt(firstDateElement.textContent, 10);
        const lastDay = parseInt(lastDateElement.textContent, 10);
        const startMonth = (firstDay > 20) ? month - 1 : month;
        const endMonth = (lastDay < 15) ? month + 1 : month;
        const viewStartDate = new Date(year, startMonth, firstDay);
        const viewEndDate = new Date(year, endMonth, lastDay);
        viewEndDate.setHours(23, 59, 59, 999);

        // --- Scrape event data (with color scraping added) ---
        const dateMapByIndex = new Map();
        let currentMonthOffset = (firstDay > 20) ? -1 : 0;
        let lastDateNum = 0;
        dateCells.forEach((cell, index) => {
            const dateNumElement = cell.querySelector('.css-g51b5d, .css-q2isom');
            if (!dateNumElement) return;
            const dateNum = parseInt(dateNumElement.textContent, 10);
            if (index > 0 && dateNum < lastDateNum) currentMonthOffset++;
            const date = new Date(year, month + currentMonthOffset, dateNum);
            dateMapByIndex.set(index, date);
            lastDateNum = dateNum;
        });

        const eventElements = document.querySelectorAll('div.lndlxo5');
        const events = [];
        const colorRegex = /#([a-fA-F0-9]{6}|[a-fA-F0-9]{3})/; // Regex to find a hex color

        eventElements.forEach(eventEl => {
            const style = getComputedStyle(eventEl);
            const eventDisplayRow = parseInt(style.getPropertyValue('--lndlxo3').trim(), 10);
            const startCol = parseInt(style.getPropertyValue('--lndlxo2').trim(), 10);
            const titleEl = eventEl.querySelector('span.lndlxo7');
            if (!titleEl || !eventDisplayRow || !startCol) return;

            const weekRowIndex = Math.floor((eventDisplayRow - 3) / 6);
            const dateCellIndex = (weekRowIndex * 7) + (startCol - 1);
            const startDate = dateMapByIndex.get(dateCellIndex);

            if (startDate) {
                // ** SCRAPE THE COLOR **
                const buttonEl = eventEl.querySelector('button');
                const styleAttr = buttonEl ? buttonEl.getAttribute('style') : '';
                const colorMatch = styleAttr.match(colorRegex);
                const color = colorMatch ? colorMatch[0] : null; // Get the full hex code (e.g., #e73b3b)

                const title = titleEl.textContent.trim();
                const timeEl = eventEl.querySelector('._1r1c5vl9, ._1bf4eeq8');
                const time = timeEl ? timeEl.textContent.trim() : '全天';

                // Add the color to the payload
                events.push({ '任務標題': title, '開始日期': startDate.toISOString(), '時間': time, 'color': color });
            }
        });

        return { events, viewStartDate: viewStartDate.toISOString(), viewEndDate: viewEndDate.toISOString() };
    }

    // --- Styling (no changes here) ---
    GM_addStyle(`
        #sync-to-gcal-button { background-color: #D32F2F; color: white; border: none; padding: 0 16px; margin-left: 12px; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: bold; height: 36px; line-height: 36px; transition: background-color 0.3s; }
        #sync-to-gcal-button:hover { background-color: #B71C1C; }
        #sync-to-gcal-button:disabled { background-color: #9E9E9E; cursor: not-allowed; }
    `);

    window.addEventListener('load', addSyncButton, false);
})();
