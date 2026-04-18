// ==UserScript==
// @name         TimeTree → Google Calendar Full Sync (range + stable waits + include current month)
// @namespace    http://tampermonkey.net/
// @version      3.7.3
// @description  selector update for 2026-04 frontend change; scan previous 1 to next 3 months with robust, slower waits; ALWAYS include current month first; sync to GAS; with range inspectors & detailed logs.
// @match        https://timetreeapp.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_addStyle
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  // ========= CONFIG =========
  // 你的 GAS Web App /exec URL（保持你現在在用的那個）
  const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbw1HOoFJOPALcRkUcgF8qrvUNK_cC_azAdhCo7cbnLGCiVcNpVLVjL9gLuvY4qJmEXXgg/exec';

  // 等待參數（可視情況加大）
  let NAVIGATION_STEP_DELAY_MS   = 400;   // 月份切換後先等這麼久，再進「穩定等待」
  let REQUIRED_STABLE_WINDOW_MS  = 700;   // 事件列數量需連續穩定這麼久
  let STABLE_SAMPLE_INTERVAL_MS  = 150;   // 事件列穩定觀察取樣間隔
  let MAX_STABLE_WAIT_MS         = 7000;  // 單月最多等多久
  let POST_STABLE_EXTRA_DELAY_MS = 220;   // 判穩後額外緩衝
  const RETRY_IF_ZERO_MAX        = 1;     // 0 筆時補等重抓次數
  const AUTO_SLOWDOWN_FACTOR     = 1.2;   // 若曾重試，本次巡覽自動放慢倍率

  // ========= UI =========
  function ensureSyncButton() {
    if (document.getElementById('sync-to-gcal-button')) return;
    const btn = document.createElement('button');
    btn.id = 'sync-to-gcal-button';
    btn.type = 'button';
    btn.textContent = 'Full Sync (-1 ~ +3 months)';
    btn.title = 'Add/delete/update colors on Google Calendar to match range.';
    btn.addEventListener('click', () => handleSyncRange(-1, 3));
    document.body.appendChild(btn);
  }

  function ensureInspectMonthButton() {
    if (document.getElementById('inspect-events-button')) return;
    const btn = document.createElement('button');
    btn.id = 'inspect-events-button';
    btn.type = 'button';
    btn.textContent = 'Inspect (this month)';
    btn.title = 'Scan current month → Console only';
    btn.addEventListener('click', async () => {
      try {
        console.log('[Inspector:month] waiting stable...');
        await waitForCalendarStable_();
        const parsed = await parseWithRetryIfZero_(scrapeCalendarDataOneView);
        logSummary_('[Inspector:month] parsed', parsed);
        if (parsed.events?.length) console.table(parsed.events);
      } catch (e) {
        console.error('[Inspector:month] Failed:', e);
      }
    });
    document.body.appendChild(btn);
  }

  function ensureInspectRangeButton() {
    if (document.getElementById('inspect-range-button')) return;
    const btn = document.createElement('button');
    btn.id = 'inspect-range-button';
    btn.type = 'button';
    btn.textContent = 'Inspect Range (-1 ~ +3)';
    btn.title = 'Scan range (includes current month first) → Console only';
    btn.addEventListener('click', () => handleInspectRange(-1, 3));
    document.body.appendChild(btn);
  }

  const mo = new MutationObserver(() => {
    ensureSyncButton();
    ensureInspectMonthButton();
    ensureInspectRangeButton();
  });
  mo.observe(document.documentElement, { childList: true, subtree: true });

  // ========= Helpers =========
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));
  const getTimeNode = () => document.querySelector('time[datetime]'); // e.g. <time datetime="2026-03">

  function readYearMonthFromTimeNode() {
    const t = getTimeNode();
    if (!t) return null;
    const dt = t.getAttribute('datetime'); // "YYYY-MM"
    if (dt && /^\d{4}-\d{2}$/.test(dt)) {
      const [yy, mm] = dt.split('-').map(v => parseInt(v, 10));
      return { year: yy, monthIndex: mm - 1, label: t.textContent.trim() };
    }
    return null;
  }

  function waitFor_(predicate, timeoutMs = 5000) {
    return new Promise((resolve, reject) => {
      const start = Date.now();
      (function loop() {
        try { if (predicate()) return resolve(true); } catch (_) {}
        if (Date.now() - start > timeoutMs) return reject(new Error('waitFor_ timeout'));
        requestAnimationFrame(loop);
      })();
    });
  }

  // 事件列穩定觀察窗：>0 筆時須連續穩定 REQUIRED_STABLE_WINDOW_MS；若始終 0 筆，至 MAX_STABLE_WAIT_MS 視為穩定
  async function waitForStableEventsCount_() {
    const start = Date.now();
    let lastCount = -1;
    let stableSince = null;

    while (Date.now() - start < MAX_STABLE_WAIT_MS) {
      const rawCount = document.querySelectorAll('div.lndlxo3, div.lndlxo4, [data-testid="eventBar"], .event-bar').length;

      if (rawCount !== lastCount) {
        lastCount = rawCount;
        stableSince = Date.now();
      }

      if (rawCount > 0 && Date.now() - stableSince >= REQUIRED_STABLE_WINDOW_MS) {
        await sleep(POST_STABLE_EXTRA_DELAY_MS);
        return;
      }

      if (rawCount === 0 && Date.now() - start >= MAX_STABLE_WAIT_MS) {
        await sleep(POST_STABLE_EXTRA_DELAY_MS);
        return;
      }

      await sleep(STABLE_SAMPLE_INTERVAL_MS);
    }
  }

  async function waitForCalendarStable_() {
    await waitFor_(() => !!document.querySelector('time[datetime]'), 5000);
    await waitFor_(() => document.querySelectorAll('div[role="gridcell"]').length >= 7, 5000);
    await waitForStableEventsCount_();
  }

  // 0 筆時補等重抓
  async function parseWithRetryIfZero_(scrapeFn) {
    let parsed = scrapeFn();
    if ((parsed.events?.length || 0) === 0) {
      console.log('[Parser] 0 events → extra wait & retry once...');
      await sleep(300);
      await waitForStableEventsCount_();
      parsed = scrapeFn();
    }
    return parsed;
  }

  function logSummary_(prefix, data) {
    const total = data.events?.length || 0;
    console.log(`${prefix}: ${total} events; ${data.viewStartDate} → ${data.viewEndDate}`);
  }

  function buildOffsets_(prevMonths, nextMonths) {
    const arr = [];
    for (let i = -prevMonths; i <= nextMonths; i++) {
      if (i !== 0) arr.push(i);
    }
    return arr;
  }

  // ========= Scraper：只抓目前畫面這一個月 =========
  function scrapeCalendarDataOneView() {
    const dateCells = document.querySelectorAll('div[role="gridcell"]');
    if (dateCells.length < 2) return { events: [] };

    const pickDateNum = el =>
      el.querySelector('._2u4y7t6, ._2u4y7t5, .css-g51b5d, .css-q2isom, [data-testid="dayNumber"], .day-number');

    const firstDateElement = pickDateNum(dateCells[0]);
    const lastDateElement  = pickDateNum(dateCells[dateCells.length - 1]);
    const tnode = getTimeNode();

    if (!tnode || !firstDateElement || !lastDateElement) return { events: [] };

    const dt = tnode.getAttribute('datetime'); // "YYYY-MM"
    const [yy, mm] = dt.split('-').map(v => parseInt(v, 10));
    if (!yy || !mm) return { events: [] };
    const year = yy, month = mm - 1;

    const firstDay = parseInt(firstDateElement.textContent, 10);
    const lastDay  = parseInt(lastDateElement.textContent, 10);

    const startMonth = (firstDay > 20) ? month - 1 : month;
    const endMonth   = (lastDay  < 15) ? month + 1 : month;

    const viewStartDate = new Date(year, startMonth, firstDay);
    const viewEndDate   = new Date(year, endMonth,  lastDay);
    viewEndDate.setHours(23, 59, 59, 999);

    const dateMapByIndex = new Map();
    let currentMonthOffset = (firstDay > 20) ? -1 : 0;
    let lastDateNum = 0;

    dateCells.forEach((cell, index) => {
      const dateNumElement = pickDateNum(cell);
      if (!dateNumElement) return;
      const dateNum = parseInt(dateNumElement.textContent, 10);
      if (index > 0 && dateNum < lastDateNum) currentMonthOffset++;
      const date = new Date(year, month + currentMonthOffset, dateNum);
      dateMapByIndex.set(index, date);
      lastDateNum = dateNum;
    });

    const eventElements = document.querySelectorAll('div.lndlxo3, div.lndlxo4, [data-testid="eventBar"], .event-bar');
    const events = [];
    const colorRegex = /#([a-fA-F0-9]{6}|[a-fA-F0-9]{3})/;

    eventElements.forEach(eventEl => {
      const cs = getComputedStyle(eventEl);
      const styleAttr2 = eventEl.getAttribute('style') || '';
      const rowMatch = styleAttr2.match(/--lndlxo0\s*:\s*(\d+)/);
      const colMatch = styleAttr2.match(/--lndlxo1\s*:\s*(\d+)/);

      const eventDisplayRow = rowMatch
        ? parseInt(rowMatch[1], 10)
        : parseInt(String(cs.getPropertyValue('--lndlxo3') || cs.getPropertyValue('--row')).trim(), 10);
      const startCol = colMatch
        ? parseInt(colMatch[1], 10)
        : parseInt(String(cs.getPropertyValue('--lndlxo2') || cs.getPropertyValue('--col')).trim(), 10);
      const titleEl = eventEl.querySelector('span.lndlxo6, span.lndlxob, span.lndlxo7, [data-testid="eventTitle"], .event-title');
      if (!titleEl || !eventDisplayRow || !startCol) return;

      const weekRowIndex = Math.floor((eventDisplayRow - 3) / 7);
      const dateCellIndex = (weekRowIndex * 7) + (startCol - 1);
      const startDate = dateMapByIndex.get(dateCellIndex);
      if (!startDate) return;

      const buttonEl = eventEl.querySelector('button');
      const styleAttr = buttonEl ? buttonEl.getAttribute('style') : '';
      const colorMatch = styleAttr && colorRegex.test(styleAttr) ? styleAttr.match(colorRegex) : null;
      const color = colorMatch ? colorMatch[0] : null;

      const title = titleEl.textContent.trim();
      const timeEl = eventEl.querySelector('._1r1c5vl9, ._1bf4eeq8, [data-testid="eventTime"], .event-time');
      const time = timeEl ? timeEl.textContent.trim() : '全天';

      events.push({ '任務標題': title, '開始日期': startDate.toISOString(), '時間': time, 'color': color });
    });

    // 偵錯輸出
    try {
      const ymLabel = tnode.textContent?.trim();
      console.groupCollapsed(`[TimeTree] ${dt} (${ymLabel}) parsed=${events.length} cells=${dateCells.length} raw=${eventElements.length}`);
      console.info('Date range:', viewStartDate.toISOString(), '→', viewEndDate.toISOString());
      if (events.length) {
        console.info('First event:', events[0]);
        console.table(events.slice(0, 100));
      } else {
        console.warn('No parsed events (maybe none or still loading).');
      }
      eventElements.forEach((el, i) => {
        const cs = getComputedStyle(el);
        const styleAttr2 = el.getAttribute('style') || '';
        const rowMatch = styleAttr2.match(/--lndlxo0\s*:\s*(\d+)/);
        const colMatch = styleAttr2.match(/--lndlxo1\s*:\s*(\d+)/);
        const rowVar = rowMatch
          ? rowMatch[1]
          : String(cs.getPropertyValue('--lndlxo3') || cs.getPropertyValue('--row')).trim();
        const colVar = colMatch
          ? colMatch[1]
          : String(cs.getPropertyValue('--lndlxo2') || cs.getPropertyValue('--col')).trim();
        const titleEl = el.querySelector('span.lndlxo6, [data-testid="eventTitle"], .event-title');
        const btn = el.querySelector('button');
        const styleAttr = btn ? btn.getAttribute('style') : '';
        console.debug(`[#${i}] row=${rowVar} col=${colVar} title="${titleEl?.textContent?.trim()}" style="${styleAttr}"`);
      });
      console.groupEnd();
    } catch (e) {
      console.error('[TimeTree] debug logging failed:', e);
    }

    return {
      events,
      viewStartDate: viewStartDate.toISOString(),
      viewEndDate: viewEndDate.toISOString(),
      rawEventCount: eventElements.length
    };
  }

  // ========= Navigation（已修正） =========
  async function goToOffset_(currentOffset, targetOffset) {
    if (targetOffset === currentOffset) return currentOffset;
    const step = targetOffset - currentOffset;
    const dir = step > 0 ? 'next' : 'previous';
    for (let s = 0; s < Math.abs(step); s++) {
      const curLabel = document.querySelector('time[datetime]')?.getAttribute('datetime');
      const [yy, mm] = (curLabel || '').split('-').map(n => parseInt(n, 10));
      const target = new Date(yy, (mm || 1) - 1 + (dir === 'next' ? 1 : -1), 1);
      (
        document.querySelector(`button[value="${dir}"]`) ||
        document.querySelector(`button[aria-label*="${dir}" i]`) ||
        document.querySelector(`[data-test-id="${dir}-button"]`) ||
        document.querySelector(`[data-test-id="navigate-${dir}"]`)
      )?.click();

      // ① 等月份文字切換
      await waitFor_(() => {
        const now = document.querySelector('time[datetime]')?.getAttribute('datetime');
        if (!now) return false;
        const [cy, cm] = now.split('-').map(n => parseInt(n, 10));
        return cy === target.getFullYear() && (cm - 1) === target.getMonth();
      }, 6000);

      // ② 導航後緩衝
      await sleep(NAVIGATION_STEP_DELAY_MS);

      // ③ 進入穩定等待
      await waitForCalendarStable_();
    }
    return targetOffset;
  }

  // ========= Inspect Range（先掃 0 → 再掃其餘 offsets） =========
  async function handleInspectRange(prevMonths = 1, nextMonths = 3) {
    try {
      await waitForCalendarStable_();
      const origin = readYearMonthFromTimeNode();
      if (!origin) throw new Error('無法讀取目前月份');

      let currentOffset = 0;
      let all = [];
      let globalStart = null, globalEnd = null;
      let hitRetry = false;

      // 先掃「當前頁面月份(0)」
      const parsed0 = await parseWithRetryIfZero_(scrapeCalendarDataOneView);
      logSummary_('[InspectRange] offset 0', parsed0);
      if ((parsed0.events?.length || 0) === 0 && parsed0.rawEventCount > 0) hitRetry = true;
      if (parsed0.events?.length) {
        all = all.concat(parsed0.events);
        const vs = new Date(parsed0.viewStartDate);
        const ve = new Date(parsed0.viewEndDate);
        globalStart = (!globalStart || vs < globalStart) ? vs : globalStart;
        globalEnd   = (!globalEnd   || ve > globalEnd)   ? ve : globalEnd;
      }

      // 再掃 -1, +1, +2, +3（不含 0）
      const offsets = buildOffsets_(prevMonths, nextMonths);
      for (const targetOffset of offsets) {
        currentOffset = await goToOffset_(currentOffset, targetOffset);
        const parsed = await parseWithRetryIfZero_(scrapeCalendarDataOneView);
        logSummary_(`[InspectRange] offset ${targetOffset}`, parsed);

        if ((parsed.events?.length || 0) === 0 && parsed.rawEventCount > 0) hitRetry = true;

        if (parsed.events?.length) {
          all = all.concat(parsed.events);
          const vs = new Date(parsed.viewStartDate);
          const ve = new Date(parsed.viewEndDate);
          globalStart = (!globalStart || vs < globalStart) ? vs : globalStart;
          globalEnd   = (!globalEnd   || ve > globalEnd)   ? ve : globalEnd;
        }
      }

      // 有重試 → 本次巡覽自動放慢
      if (hitRetry) {
        NAVIGATION_STEP_DELAY_MS       = Math.round(NAVIGATION_STEP_DELAY_MS       * AUTO_SLOWDOWN_FACTOR);
        REQUIRED_STABLE_WINDOW_MS      = Math.round(REQUIRED_STABLE_WINDOW_MS      * AUTO_SLOWDOWN_FACTOR);
        MAX_STABLE_WAIT_MS             = Math.round(MAX_STABLE_WAIT_MS             * AUTO_SLOWDOWN_FACTOR);
        POST_STABLE_EXTRA_DELAY_MS     = Math.round(POST_STABLE_EXTRA_DELAY_MS     * AUTO_SLOWDOWN_FACTOR);
        console.warn('[InspectRange] Hit retry → auto slowdown applied for this session.');
      }

      // 回原月份
      currentOffset = await goToOffset_(currentOffset, 0);

      // 匯總
      console.group('[InspectRange] Summary (-1 ~ +3 months)');
      console.log('Total events:', all.length);
      if (all.length) {
        console.log('Global range:', globalStart?.toISOString(), '→', globalEnd?.toISOString());
        console.table(all.slice(0, 200));
      }
      console.groupEnd();

    } catch (e) {
      console.error('[InspectRange] Failed:', e);
    }
  }

  // ========= Full Sync（先掃 0 → 再掃其餘 offsets → 送 GAS） =========
  async function handleSyncRange(prevMonths = 1, nextMonths = 3) {
    const btn = document.getElementById('sync-to-gcal-button');
    if (!confirm(`將同步目前畫面所在月份的前 ${prevMonths} 個月，到後 ${nextMonths} 個月的事件（新增/刪除/顏色）。\n\n是否繼續？`)) return;

    try {
      if (btn) { btn.disabled = true; btn.textContent = 'Analyzing...'; }

      const origin = readYearMonthFromTimeNode();
      if (!origin) throw new Error('無法讀取目前月份');

      let currentOffset = 0;
      let allEvents = [];
      let globalStart = null, globalEnd = null;
      let hitRetry = false;

      await waitForCalendarStable_();

      // 先掃「當前頁面月份(0)」
      let parsed0 = await parseWithRetryIfZero_(scrapeCalendarDataOneView);
      logSummary_('[FullSync] offset 0', parsed0);
      if ((parsed0.events?.length || 0) === 0 && parsed0.rawEventCount > 0) hitRetry = true;
      if (Array.isArray(parsed0.events) && parsed0.events.length > 0) {
        allEvents = allEvents.concat(parsed0.events);
        const vs = new Date(parsed0.viewStartDate);
        const ve = new Date(parsed0.viewEndDate);
        globalStart = (!globalStart || vs < globalStart) ? vs : globalStart;
        globalEnd   = (!globalEnd   || ve > globalEnd)   ? ve : globalEnd;
      }

      // 再掃其餘 offsets（不含 0）
      const offsets = buildOffsets_(prevMonths, nextMonths);
      for (const targetOffset of offsets) {
        currentOffset = await goToOffset_(currentOffset, targetOffset);

        if (btn) btn.textContent = `Scraping month ${targetOffset >= 0 ? '+' + targetOffset : targetOffset}...`;

        const parsed = await parseWithRetryIfZero_(scrapeCalendarDataOneView);
        logSummary_(`[FullSync] offset ${targetOffset}`, parsed);

        if ((parsed.events?.length || 0) === 0 && parsed.rawEventCount > 0) hitRetry = true;

        if (Array.isArray(parsed.events) && parsed.events.length > 0) {
          allEvents = allEvents.concat(parsed.events);
          const vs = new Date(parsed.viewStartDate);
          const ve = new Date(parsed.viewEndDate);
          globalStart = (!globalStart || vs < globalStart) ? vs : globalStart;
          globalEnd   = (!globalEnd   || ve > globalEnd)   ? ve : globalEnd;
        }
      }

      if (hitRetry) {
        NAVIGATION_STEP_DELAY_MS       = Math.round(NAVIGATION_STEP_DELAY_MS       * AUTO_SLOWDOWN_FACTOR);
        REQUIRED_STABLE_WINDOW_MS      = Math.round(REQUIRED_STABLE_WINDOW_MS      * AUTO_SLOWDOWN_FACTOR);
        MAX_STABLE_WAIT_MS             = Math.round(MAX_STABLE_WAIT_MS             * AUTO_SLOWDOWN_FACTOR);
        POST_STABLE_EXTRA_DELAY_MS     = Math.round(POST_STABLE_EXTRA_DELAY_MS     * AUTO_SLOWDOWN_FACTOR);
        console.warn('[FullSync] Hit retry → auto slowdown applied for this session.');
      }

      // 回原月份
      currentOffset = await goToOffset_(currentOffset, 0);

      if (!globalStart || !globalEnd || allEvents.length === 0) {
        if (btn) { btn.textContent = 'No events found in the selected range'; setTimeout(() => { btn.textContent = 'Full Sync (-1 ~ +3 months)'; btn.disabled = false; }, 3000); }
        return;
      }

      if (btn) btn.textContent = `Sending ${allEvents.length} events...`;

      const payload = {
        events: allEvents,
        viewStartDate: globalStart.toISOString(),
        viewEndDate: globalEnd.toISOString()
      };

      GM_xmlhttpRequest({
        method: 'POST',
        url: APPS_SCRIPT_URL,
        data: JSON.stringify(payload),
        headers: { 'Content-Type': 'application/json' },
        onload: function (response) {
          try {
            const result = JSON.parse(response.responseText);
            if (btn) btn.textContent = result.status === 'success'
              ? `✅ ${result.message}`
              : `❌ Error: ${result.message}`;
          } catch (_) {
            if (btn) btn.textContent = '❌ Response parse error';
          }
          if (btn) setTimeout(() => { btn.textContent = 'Full Sync (-1 ~ +3 months)'; btn.disabled = false; }, 10000);
        },
        onerror: function (error) {
          console.error('Sync Script Error:', error);
          if (btn) {
            btn.textContent = '❌ Network Error!';
            setTimeout(() => { btn.textContent = 'Full Sync (-1 ~ +3 months)'; btn.disabled = false; }, 5000);
          }
        }
      });

    } catch (err) {
      console.error(err);
      const btn = document.getElementById('sync-to-gcal-button');
      if (btn) { btn.textContent = `❌ ${err.message || 'Unexpected error'}`; setTimeout(() => { btn.textContent = 'Full Sync (-1 ~ +3 months)'; btn.disabled = false; }, 5000); }
    }
  }

  // ========= Styles =========
  GM_addStyle(`
    #sync-to-gcal-button {
      position: fixed; right: 16px; bottom: 16px; z-index: 2147483647;
      background-color: #D32F2F; color: #fff; border: none;
      padding: 0 16px; border-radius: 20px; cursor: pointer;
      font-size: 14px; font-weight: bold; height: 40px; line-height: 40px;
      box-shadow: 0 4px 10px rgba(0,0,0,0.2);
      transition: background-color .2s, transform .1s;
    }
    #sync-to-gcal-button:hover { background-color: #B71C1C; }
    #sync-to-gcal-button:active { transform: translateY(1px); }
    #sync-to-gcal-button:disabled { background-color: #9E9E9E; cursor: not-allowed; }

    #inspect-events-button {
      position: fixed; right: 16px; bottom: 64px; z-index: 2147483647;
      background-color: #455A64; color: #fff; border: none;
      padding: 0 16px; border-radius: 20px; cursor: pointer;
      font-size: 13px; font-weight: 600; height: 36px; line-height: 36px;
      box-shadow: 0 4px 10px rgba(0,0,0,0.2);
      transition: background-color .2s, transform .1s;
    }
    #inspect-events-button:hover { background-color: #37474F; }
    #inspect-events-button:active { transform: translateY(1px); }

    #inspect-range-button {
      position: fixed; right: 16px; bottom: 106px; z-index: 2147483647;
      background-color: #2E7D32; color: #fff; border: none;
      padding: 0 16px; border-radius: 20px; cursor: pointer;
      font-size: 13px; font-weight: 600; height: 36px; line-height: 36px;
      box-shadow: 0 4px 10px rgba(0,0,0,0.2);
      transition: background-color .2s, transform .1s;
    }
    #inspect-range-button:hover { background-color: #1B5E20; }
    #inspect-range-button:active { transform: translateY(1px); }
  `);

  // 初始化按鈕
  ensureSyncButton();
  ensureInspectMonthButton();
  ensureInspectRangeButton();
})();
