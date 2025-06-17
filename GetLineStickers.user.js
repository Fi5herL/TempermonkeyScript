// ==UserScript==
// @name         貼圖拿過來
// @namespace    http://tampermonkey.net/
// @version      1.3
// @description  Click on a LINE sticker or emoji preview on store.line.me to copy its original image to the clipboard.
// @description:zh-TW 點擊 store.line.me 上的 LINE 貼圖或表情貼預覽，將其原始圖片複製到剪貼簿。
// @author       Your Name (或 AI)
// @match        https://store.line.me/stickershop/product/*
// @match        https://store.line.me/emojishop/product/*
// @grant        GM_addStyle
// @grant        GM_setClipboard
// @icon         https://www.google.com/s2/favicons?sz=64&domain=line.me
// ==/UserScript==

(function() {
    'use_strict';

    // --- 樣式 ---
    GM_addStyle(`
        .sticker-copier-notification {
            position: fixed;
            bottom: 20px;
            right: 20px;
            background-color: #4CAF50;
            color: white;
            padding: 15px;
            border-radius: 5px;
            z-index: 9999;
            font-size: 14px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.2);
            opacity: 0;
            transition: opacity 0.5s ease-in-out;
            text-align: center;
        }
        .sticker-copier-notification.show {
            opacity: 1;
        }
        .sticker-copier-notification.error {
            background-color: #F44336; /* Red for errors */
        }
    `);

    // --- 通知函數 ---
    function showNotification(message, isError = false) {
        let notification = document.querySelector('.sticker-copier-notification');
        if (!notification) {
            notification = document.createElement('div');
            notification.className = 'sticker-copier-notification';
            document.body.appendChild(notification);
        }

        notification.textContent = message;
        notification.classList.remove('error');
        if (isError) {
            notification.classList.add('error');
        }
        notification.classList.add('show');

        setTimeout(() => {
            notification.classList.remove('show');
            if (notification.parentNode) { // 確保元素還在DOM中
                // 可選：如果3秒後沒有新的通知，則移除DOM元素以保持清潔
                // setTimeout(() => {
                // if (!notification.classList.contains('show')) {
                // notification.remove();
                // }
                // }, 500); // 額外延遲後檢查並移除
            }
        }, 3000); // 3 秒後消失
    }

    // --- 輔助函數：從 preview data 獲取圖片 URL ---
    function getImageUrlFromPreview(previewData) {
        let imageUrl = '';
        // 優先順序:
        // 1. popupUrl (通常是貼圖的最高畫質，動態貼圖的 popupUrl 也是動態的)
        // 2. animationUrl (動態貼圖/表情貼的動態版本)
        // 3. staticUrl (靜態版本)
        // 4. fallbackStaticUrl (備用靜態版本，在列表項中較常見)

        if (previewData.popupUrl) {
            imageUrl = previewData.popupUrl;
        } else if (previewData.animationUrl) {
            imageUrl = previewData.animationUrl;
        } else if (previewData.staticUrl) {
            imageUrl = previewData.staticUrl;
        } else if (previewData.fallbackStaticUrl) {
            imageUrl = previewData.fallbackStaticUrl;
        }
        return imageUrl;
    }

    // --- 輔助函數：複製圖片到剪貼簿 ---
    async function copyImageToClipboard(imageUrl, successMessage, itemType = "項目") {
        if (!imageUrl) {
            console.warn(`No image URL provided for ${itemType}.`);
            showNotification(`找不到${itemType}的有效圖片網址`, true);
            return;
        }

        console.log(`Attempting to copy ${itemType}:`, imageUrl);
        try {
            const response = await fetch(imageUrl);
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status} for ${imageUrl}`);
            }
            const blob = await response.blob();

            let effectiveType = blob.type;
            // LINE APNG/動態表情貼 的 type 可能是 'image/png'
            if (!effectiveType || !effectiveType.startsWith('image/')) {
                console.warn(`Blob type "${effectiveType}" for ${imageUrl} might not be ideal. Trying to infer from URL.`);
                if (imageUrl.includes('_animation.png') || imageUrl.includes('/sticker.png') || imageUrl.includes('/main.png') || imageUrl.endsWith('.png') || imageUrl.includes('/sticon/')) {
                   effectiveType = 'image/png'; // APNGs are often served as image/png
                } else if (imageUrl.endsWith('.gif')) {
                   effectiveType = 'image/gif';
                } else {
                   effectiveType = 'image/png'; // Default
                }
                console.log('Inferred type:', effectiveType);
            }

            await navigator.clipboard.write([
                new ClipboardItem({
                    [effectiveType]: blob
                })
            ]);
            showNotification(successMessage);
            console.log(`${itemType} image copied to clipboard!`);
        } catch (err) {
            console.error(`Failed to copy ${itemType} image:`, err);
            try {
                GM_setClipboard(imageUrl);
                showNotification(`${itemType}圖片複製失敗，已複製圖片網址。`, true);
                console.log(`${itemType} image copy failed, URL copied to clipboard as fallback.`);
            } catch (gmErr) {
                showNotification(`${itemType}複製失敗: ` + (err.message || err), true);
                console.error(`Fallback GM_setClipboard for ${itemType} also failed:`, gmErr);
            }
        }
    }

    // --- 主要邏輯：列表中的項目 ---
    const itemListContainer = document.querySelector('div.mdCMN09ImgList'); // 通用父容器

    if (itemListContainer) {
        itemListContainer.addEventListener('click', async function(event) {
            const itemLi = event.target.closest('li.FnStickerPreviewItem'); // 貼圖和表情貼列表項通用 class

            if (itemLi) {
                event.preventDefault();
                event.stopPropagation();

                const previewDataAttr = itemLi.dataset.preview;
                if (!previewDataAttr) {
                    console.warn('Item data-preview attribute not found.');
                    showNotification('找不到項目資料', true);
                    return;
                }

                try {
                    const previewData = JSON.parse(previewDataAttr);
                    const imageUrl = getImageUrlFromPreview(previewData);
                    const itemType = window.location.pathname.includes('/stickershop/') ? '貼圖' : '表情貼';
                    await copyImageToClipboard(imageUrl, `${itemType}已複製到剪貼簿！`, itemType);

                } catch (e) {
                    console.error('Error processing item click:', e);
                    showNotification('處理點擊時發生錯誤', true);
                }
            }
        });
        console.log('LINE Item Copier activated for list items.');
    } else {
        console.warn('Item list container (div.mdCMN09ImgList) not found. Script might not work correctly for list items.');
    }

    // --- 主要邏輯：頁面頂部的主預覽圖 ---
    const mainImageContainer = document.querySelector('div.mdCMN38Img[data-preview]');
    if (mainImageContainer) {
        mainImageContainer.addEventListener('click', async function(event) {
            event.preventDefault();
            event.stopPropagation();

            const previewDataAttr = mainImageContainer.dataset.preview;
             if (!previewDataAttr) {
                console.warn('Main item data-preview attribute not found.');
                showNotification('找不到主要項目資料', true);
                return;
            }
            try {
                const previewData = JSON.parse(previewDataAttr);
                const imageUrl = getImageUrlFromPreview(previewData);
                const itemType = window.location.pathname.includes('/stickershop/') ? '主要貼圖' : '主要表情貼';
                await copyImageToClipboard(imageUrl, `${itemType}已複製！`, itemType);

            } catch (e) {
                console.error('Error processing main item click:', e);
                showNotification('處理主要項目點擊時發生錯誤', true);
            }
        });
        console.log('LINE Item Copier activated for main preview image.');
    } else {
         console.warn('Main preview image container (div.mdCMN38Img[data-preview]) not found.');
    }

})();
