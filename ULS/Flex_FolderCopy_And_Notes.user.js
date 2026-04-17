// ==UserScript==
// @name         Flex Folder Copy and Notes
// @namespace    http://tampermonkey.net/
// @version      1.2.0
// @description  Copies folders and adds notes with enhanced UI
// @author       Fi5herL
// @match        *://*/*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // Notes Panel UI styles update
    const panelStyles = `
    .ffn-notes-panel {
        background-color: #f0f0f0;  /* Updated background color */
        border-radius: 8px;  /* Rounded corners */
        padding: 16px;  /* Spacing */
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    .ffn-notes-header {
        font-size: 15px;
        font-weight: 600;
    }
    .ffn-notes-controls,
    input,
    select,
    button {
        font-size: 16px;
        margin: 4px;
    }
    .ffn-notes-list {
        margin-top: 10px;
    }
    .ffn-note-card:hover {
        transform: translateY(-2px);  /* Slight lift effect */
    }
    .ffn-note-title {
        font-size: 16px;
        /* other styles */
    }
    .ffn-note-meta,
    .ffn-note-time {
        font-size: 12.5px;
        /* other styles */
    }
    `;

    const styleSheet = document.createElement("style");
    styleSheet.type = "text/css";
    styleSheet.innerText = panelStyles;
    document.head.appendChild(styleSheet);

    // Modify note preview without emoji
    const notes = document.querySelectorAll('.ffn-note-card');
    notes.forEach(note => {
        const preview = note.querySelector('.note-preview');
        const shortText = preview.textContent.substring(3);  // Remove emoji prefix
        preview.textContent = shortText;
    });

})();
