// GM_addStyle for a memos-like UI for the notes overlay/panel
GM_addStyle(`
  .ffn-notes-overlay {
    background-color: lightgray;
    width: 80vw;
    max-height: 90vh;
  }
  .ffn-notes-panel {
    background-color: white;
    border-radius: 12px;
    box-shadow: 0 2px 5px rgba(0, 0, 0, 0.3);
    padding: 16px;
  }
  .ffn-notes-header {
    font-size: 16px;
    line-height: 1.6;
  }
  .ffn-notes-controls {
    margin-bottom: 10px;
  }
  .ffn-notes-list {
    margin-top: 10px;
  }
  .ffn-note-card {
    background-color: white;
    border-radius: 12px;
    box-shadow: 0 2px 5px rgba(0, 0, 0, 0.1);
    margin: 8px 0;
    padding: 10px;
  }
  .ffn-note-title {
    font-size: 15px;
    font-weight: 600;
    margin-bottom: 4px;
  }
  .ffn-note-meta {
    font-size: 12.5px;
    color: gray;
  }
  .ffn-note-time {
    font-size: 12.5px;
    color: darkgray;
  }
  input, button {
    font-size: 16px;
  }
  .ffn-note-editor textarea {
    font-size: 16px;
  }
`); 

// Removing emoji prefix from note preview text
const notePreviewText = shortText.replace(/^📝 /, '');

// Logic to show plain text without emoji prefix
// ... (existing logic to display note preview)