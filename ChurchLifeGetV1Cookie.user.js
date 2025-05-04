// ==UserScript==
// @name         自動複製 Cookie 並顯示確認 ✅
// @namespace    http://tampermonkey.net/
// @version      2025-04-27
// @description  try to take over the world!
// @author       You
// @match        https://www.chlife-stat.org/index.php
// @icon         https://www.google.com/s2/favicons?sz=64&domain=chlife-stat.org
// @grant        GM_setClipboard
// @grant        GM_addStyle
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

    // 獲取當前網域的所有 Cookie
    const cookies = document.cookie;

    // 檢查是否有 Cookie 可供複製
    if (cookies && cookies.length > 0) {
        // 複製 Cookie 到剪貼簿
        // 第一個參數是 要複製的文字
        // 第二個參數是 可選的，複製的內容類型，預設是 'text'
        GM_setClipboard(cookies, 'text');

        // --- 顯示成功提示 ---

        // 定義提示框的 CSS 樣式 (使用 GM_addStyle 更佳)
        GM_addStyle(`
            .cookie-copy-success-toast-checkmark {
                position: fixed;
                top: 25px;
                right: 25px;
                padding: 15px 30px;
                background-color: rgba(40, 167, 69, 0.9); /* 深綠色，稍微透明 */
                color: white;
                border-radius: 10px;
                z-index: 99999; /* 確保在最上層 */
                font-size: 32px; /* 放大圖示 */
                box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
                opacity: 1;
                transition: opacity 0.5s ease-out; /* 淡出效果 */
                cursor: default; /* 避免顯示文字輸入游標 */
                line-height: 1; /* 確保圖示垂直居中 */
            }
        `);

        // 創建提示框 DIV 元素
        const notificationDiv = document.createElement('div');
        notificationDiv.classList.add('cookie-copy-success-toast-checkmark');
        notificationDiv.textContent = '✅'; // 設定顯示的圖示

        // 將提示框添加到頁面 body
        document.body.appendChild(notificationDiv);

        // 設置計時器，在 2.5 秒後開始淡出，3 秒後移除提示框
        setTimeout(() => {
            notificationDiv.style.opacity = '0'; // 開始淡出
            setTimeout(() => {
                // 再次檢查元素是否存在，以防萬一它已被移除
                if (notificationDiv && notificationDiv.parentNode) {
                   notificationDiv.remove();
                }
            }, 500); // 等待淡出動畫完成 (0.5秒)
        }, 2500); // 提示顯示 2.5 秒後開始淡出

        console.log('Cookies 已成功複製到剪貼簿！');

    } else {
        console.log('Cookie Copy Script: 在此網域未找到 Cookie 或 Cookie 為空。');
        // 如果沒有 cookie，可以選擇不顯示提示，或者顯示不同的提示
    }
})();
