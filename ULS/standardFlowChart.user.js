// ==UserScript==
// @name         UL Standards Reference Extractor & Visualizer (On-Demand)
// @namespace    http://tampermonkey.net/
// @version      1.7.0
// @description  Extracts and visualizes UL standards on-demand via Ctrl+Shift+Z shortcut, with full pagination and export features.
// @author       Standards Navigator (with modifications)
// @match        https://www.ulsestandards.org/uls-standardsdocs/onlineviewer/*
// @match        https://ulstandards.ul.com/uls-standardsdocs/onlineviewer/*
// @grant        GM_addStyle
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_xmlhttpRequest
// @require      https://cdn.jsdelivr.net/npm/mermaid@10.6.1/dist/mermaid.min.js
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

    // Configuration
    const CONFIG = {
        version: '1.7.0',
        debug: GM_getValue('debug', true),
        diagramTheme: GM_getValue('diagramTheme', 'default'),
        panelPosition: GM_getValue('panelPosition', 'right'),
    };

    // Logger utility
    const Logger = {
        log: (...args) => CONFIG.debug && console.log('[UL Standards]', ...args),
        warn: (...args) => console.warn('[UL Standards]', ...args),
        error: (...args) => console.error('[UL Standards]', ...args)
    };

    // Standards reference patterns
    const PATTERNS = {
        ul: /\b(?:UL|ul)\s*[-\s]?\s*(\d+(?:[A-Z]+)?(?:-\d+)?(?:-\d{4})?)\b/gi,
        iec: /\b(?:IEC|iec)\s*[-\s]?\s*(\d+(?:-\d+)?(?:-\d+)?(?:-\d{4})?)\b/gi,
        ansi: /\b(?:ANSI|ansi)(?:\/\w+)?\s*[-\s]?\s*([A-Z]?\d+(?:\.\d+)?(?:-\d+)?)\b/gi,
        csa: /\b(?:CSA|csa)\s*[-\s]?\s*([A-Z]?\d+(?:\.\d+)?(?:-\d+)?)\b/gi,
        en: /\b(?:EN|en)\s*[-\s]?\s*(\d+(?:-\d+)?(?:-\d+)?)\b/gi,
        nfpa: /\b(?:NFPA|nfpa)\s*[-\s]?\s*(\d+(?:[A-Z]+)?)\b/gi
    };

    // CSV processor class
    class CSVProcessor {
        constructor() { this.uploadedStandards = new Set(); }
        parseCSVContent(csvText) {
            const lines = csvText.split('\n').filter(line => line.trim());
            const standards = new Set();
            lines.forEach(line => {
                const columns = line.split(',');
                if (columns.length > 0) {
                    const firstColumn = columns[0].trim().replace(/^"|"$/g, '');
                    const dashIndex = firstColumn.indexOf(' - ');
                    let standardText = dashIndex !== -1 ? firstColumn.substring(0, dashIndex) : firstColumn;
                    if (this.looksLikeStandard(standardText.trim())) {
                        standards.add(standardText.trim());
                    }
                }
            });
            return Array.from(standards);
        }
        looksLikeStandard(text) { return text && text.length > 2 && /[A-Za-z]/.test(text) && /\d/.test(text); }
        setUploadedStandards(standards) { this.uploadedStandards = new Set(standards); }
    }

    // Standards extractor class
    class StandardsExtractor {
        constructor() {
            this.standards = new Map();
            this.csvProcessor = new CSVProcessor();
        }

        extractStandards() {
            Logger.log('Starting on-demand standards extraction...');
            this.standards.clear();
            const textNodes = this.getTextNodes(document.body);
            let totalFound = 0;

            textNodes.forEach(node => {
                const text = node.textContent;
                Object.entries(PATTERNS).forEach(([type, pattern]) => {
                    let match;
                    const regex = new RegExp(pattern.source, 'gi');
                    while ((match = regex.exec(text)) !== null) {
                        const standardId = `${type.toUpperCase()} ${match[1]}`;
                        if (!this.standards.has(standardId)) {
                            this.standards.set(standardId, { id: standardId, type: type.toUpperCase(), occurrences: [], source: 'pattern' });
                        }
                        this.standards.get(standardId).occurrences.push({ node, context: this.getNodeContext(node) });
                        totalFound++;
                    }
                });
            });

            Logger.log(`Extracted ${totalFound} standard references (${this.standards.size} unique)`);
            this.findCrossReferences();
            return Array.from(this.standards.values());
        }

        getTextNodes(rootElement) {
            const textNodes = [];
            try {
                const walker = document.createTreeWalker(rootElement, NodeFilter.SHOW_TEXT, {
                    acceptNode: n => n.parentElement && !['SCRIPT', 'STYLE', 'NOSCRIPT'].includes(n.parentElement.tagName) && n.textContent.trim() ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT
                });
                let node; while (node = walker.nextNode()) textNodes.push(node);
                rootElement.querySelectorAll('iframe').forEach(iframe => {
                    try {
                        const iDoc = iframe.contentDocument;
                        if (iDoc && iDoc.body) textNodes.push(...this.getTextNodes(iDoc.body));
                    } catch (e) {}
                });
            } catch (e) { Logger.error("Error walking text nodes:", e); }
            return textNodes;
        }

        getNodeContext(node) {
            const context = { xTitleText: null };
            let searchElement = node.parentElement;
            if (!searchElement) return context;
            let titleFound = false;
            while (searchElement && searchElement !== document.body && !titleFound) {
                let sibling = searchElement.previousElementSibling;
                while (sibling) {
                    if (sibling.matches('div[class*="x-title"]')) {
                        context.xTitleText = sibling.textContent.trim();
                        titleFound = true; break;
                    }
                    const innerTitle = sibling.querySelector('div[class*="x-title"]');
                    if (innerTitle) {
                        context.xTitleText = innerTitle.textContent.trim();
                        titleFound = true; break;
                    }
                    sibling = sibling.previousElementSibling;
                }
                if (titleFound) break;
                searchElement = searchElement.parentElement;
                if (searchElement && searchElement.matches('div[class*="x-clause-"]')) {
                    const boundaryTitle = searchElement.querySelector('div[class*="x-title"]');
                    if (boundaryTitle && (boundaryTitle.compareDocumentPosition(node) & Node.DOCUMENT_POSITION_FOLLOWING)) {
                        context.xTitleText = boundaryTitle.textContent.trim();
                    }
                    break;
                }
            }
            return context;
        }
        findCrossReferences() {
            const allIds = new Set(this.standards.keys());
            this.standards.forEach(standard => {
                const crossRefs = new Set();
                const contextText = standard.occurrences.map(o => o.node.parentElement.textContent).join(' ').substring(0, 2000);
                allIds.forEach(id => {
                    if (id !== standard.id && contextText.includes(id.split(' ')[1])) {
                         crossRefs.add(id);
                    }
                });
                standard.crossReferences = Array.from(crossRefs);
            });
        }
    }

    // DiagramGenerator with Pagination Logic
    class DiagramGenerator {
        constructor() {
            this.EDGE_LIMIT_PER_PAGE = 50;
            this.mermaidInitialized = false;
            this.initializeMermaid();
        }
        initializeMermaid() {
            if (this.mermaidInitialized) return;
            try {
                mermaid.initialize({ startOnLoad: false, theme: CONFIG.diagramTheme, securityLevel: 'loose', flowchart: { useMaxWidth: true, htmlLabels: true, curve: 'basis' } });
                this.mermaidInitialized = true;
            } catch (e) { Logger.error('Failed to initialize Mermaid:', e); }
        }
        _cleanText(text) { return String(text).replace(/[^a-zA-Z0-9\s-]/g, '_'); }
        sanitizeId(id) { return this._cleanText(id).replace(/\s/g, '_'); }
        escapeLabel(label) { return this._cleanText(label).replace(/"/g, '#quot;'); }
        generateFlowchart(standards, filter = null) {
            const filteredStandards = filter ? standards.filter(s => s.id.toLowerCase().includes(filter.toLowerCase())) : standards;
            if (filteredStandards.length === 0) return ['graph LR\n    A["No standards found"]'];

            const titlesToStandardsMap = new Map();
            const allStandardNodes = new Map();
            const allTitleNodes = new Map();
            let titleCounter = 0;

            filteredStandards.forEach(standard => {
                allStandardNodes.set(standard.id, standard);
                if (standard.occurrences.length > 0) {
                    const uniqueTitlesForStandard = new Set();
                    standard.occurrences.forEach(occ => {
                        if (occ.context && occ.context.xTitleText) uniqueTitlesForStandard.add(occ.context.xTitleText.replace(/\s+/g, ' ').trim());
                    });
                    uniqueTitlesForStandard.forEach(titleText => {
                        if (!allTitleNodes.has(titleText)) allTitleNodes.set(titleText, { id: `title_${titleCounter++}` });
                        if (!titlesToStandardsMap.has(titleText)) titlesToStandardsMap.set(titleText, new Set());
                        titlesToStandardsMap.get(titleText).add(standard.id);
                    });
                }
            });

            const diagramChunks = [];
            let currentChunk = this._createNewChunk();
            const sortedTitles = Array.from(titlesToStandardsMap.keys()).sort();

            for (const titleText of sortedTitles) {
                const standardsInTitle = titlesToStandardsMap.get(titleText);
                const edgesForThisTitle = standardsInTitle.size;
                if (currentChunk.edgeCount > 0 && currentChunk.edgeCount + edgesForThisTitle > this.EDGE_LIMIT_PER_PAGE) {
                    diagramChunks.push(currentChunk);
                    currentChunk = this._createNewChunk();
                }
                const titleInfo = allTitleNodes.get(titleText);
                currentChunk.titleDefs.set(titleText, titleInfo);
                standardsInTitle.forEach(stdId => {
                    currentChunk.standardDefs.set(stdId, allStandardNodes.get(stdId));
                    currentChunk.hierarchyEdges.push({ from: titleInfo.id, to: this.sanitizeId(stdId) });
                });
                currentChunk.edgeCount += edgesForThisTitle;
            }
            if (currentChunk.edgeCount > 0) diagramChunks.push(currentChunk);
            if (diagramChunks.length === 0 && allStandardNodes.size > 0) {
                 allStandardNodes.forEach((std, id) => currentChunk.standardDefs.set(id, std));
                 diagramChunks.push(currentChunk);
            }

            diagramChunks.forEach(chunk => {
                const definedStandards = new Set(chunk.standardDefs.keys());
                chunk.standardDefs.forEach(standard => {
                    if (standard.crossReferences) {
                        standard.crossReferences.forEach(refId => {
                            if (definedStandards.has(refId)) chunk.crossRefEdges.push({ from: this.sanitizeId(standard.id), to: this.sanitizeId(refId) });
                        });
                    }
                });
            });
            return diagramChunks.map(chunk => this.buildDiagramString(chunk));
        }
        _createNewChunk() { return { edgeCount: 0, titleDefs: new Map(), standardDefs: new Map(), hierarchyEdges: [], crossRefEdges: [] }; }
        buildDiagramString(chunk) {
            let diagram = 'graph LR\n';
            chunk.titleDefs.forEach((info, text) => diagram += `    ${info.id}([${this.escapeLabel(text.substring(0, 70))}])\n`);
            chunk.standardDefs.forEach((std, id) => {
                const nodeLabel = `"${(std.source === 'csv' ? '📁' : '🔍')} <b>${this.escapeLabel(id)}</b><br/>${std.occurrences.length} occurrences"`;
                diagram += `    ${this.sanitizeId(id)}[${nodeLabel}]\n`;
            });
            chunk.hierarchyEdges.forEach(edge => diagram += `    ${edge.from} --> ${edge.to}\n`);
            chunk.crossRefEdges.forEach(edge => diagram += `    ${edge.from} -.-> ${edge.to}\n`);
            diagram += '    classDef titleNode fill:#f9f9f9,stroke:#333,stroke-width:2px,color:#000,font-weight:bold\n';
            chunk.titleDefs.forEach(info => diagram += `    class ${info.id} titleNode\n`);
            diagram += '    classDef ulClass fill:#e1f5fe,stroke:#01579b,stroke-width:2px,color:#000\n';
            diagram += '    classDef iecClass fill:#f3e5f5,stroke:#4a148c,stroke-width:2px,color:#000\n';
            diagram += '    classDef ansiClass fill:#e8f5e8,stroke:#1b5e20,stroke-width:2px,color:#000\n';
            diagram += '    classDef csaClass fill:#ffcdd2,stroke:#b71c1c,stroke-width:2px,color:#000\n';
            diagram += '    classDef enClass fill:#fff9c4,stroke:#f57f17,stroke-width:2px,color:#000\n';
            diagram += '    classDef nfpaClass fill:#d1c4e9,stroke:#311b92,stroke-width:2px,color:#000\n';
            diagram += '    classDef otherClass fill:#f5f5f5,stroke:#616161,stroke-width:2px,color:#000\n';
            diagram += '    classDef csvClass fill:#fff3cd,stroke:#ff8f00,stroke-width:3px,color:#000\n';
            const styleClasses = {ul: 'ulClass', iec: 'iecClass', ansi: 'ansiClass', csa: 'csaClass', en: 'enClass', nfpa: 'nfpaClass'};
            chunk.standardDefs.forEach(std => {
                const nodeId = this.sanitizeId(std.id);
                const style = std.source === 'csv' ? 'csvClass' : styleClasses[std.type.toLowerCase()] || 'otherClass';
                diagram += `    class ${nodeId} ${style}\n`;
            });
            return diagram;
        }
        async renderDiagram(diagramCode, elementId) {
            const element = document.getElementById(elementId);
            if (!element) return;
            element.innerHTML = `<div class="mermaid">${diagramCode}</div>`;
            await mermaid.run({ nodes: [element.firstChild] });
            this.addDiagramClickHandlers(element);
        }
        addDiagramClickHandlers(diagramElement) {
            const allStandards = Array.from(window.ulStandardsExtractor.standards.values());
            diagramElement.querySelectorAll('.node').forEach(node => {
                const matchingStandard = allStandards.find(std => this.sanitizeId(std.id) === node.id);
                if (matchingStandard) {
                    node.style.cursor = 'pointer';
                    node.addEventListener('click', e => {
                        e.preventDefault(); e.stopPropagation();
                        if (matchingStandard.occurrences[0]) {
                            matchingStandard.occurrences[0].node.parentElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
                            this.highlightElement(matchingStandard.occurrences[0].node.parentElement);
                        }
                    });
                }
            });
        }
        highlightElement(element) {
            element.style.transition = 'background-color 0.3s';
            element.style.backgroundColor = 'yellow';
            setTimeout(() => element.style.backgroundColor = '', 2000);
        }
    }

    // UIManager with On-Demand Logic
    class UIManager {
        constructor() {
            this.panel = null;
            this.isVisible = false; // Start hidden
            this.currentFilter = '';
            this.diagramChunks = [];
            this.currentPage = 0;
            this.initializeStyles();
        }

        initializeStyles() {
            GM_addStyle(`
                #ul-standards-panel { position: fixed; top: 20px; ${CONFIG.panelPosition}: 20px; width: 450px; max-height: 80vh; background: #ffffff; border: 2px solid #1976d2; border-radius: 8px; box-shadow: 0 8px 32px rgba(0,0,0,0.2); z-index: 999999; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; display: flex; flex-direction: column; resize: both; overflow: hidden; }
                #ul-standards-panel.collapsed { height: 40px !important; overflow: hidden; }
                #ul-standards-panel.hidden { display: none; }
                .ul-panel-header { background: linear-gradient(135deg, #1976d2, #1565c0); color: white; padding: 8px 12px; font-weight: 600; display: flex; justify-content: space-between; align-items: center; cursor: move; user-select: none; }
                .ul-panel-title { font-size: 14px; }
                .ul-panel-controls { display: flex; gap: 5px; }
                .ul-panel-btn { background: rgba(255,255,255,0.2); border: none; color: white; padding: 4px 8px; border-radius: 3px; cursor: pointer; font-size: 11px; transition: background 0.2s; }
                .ul-panel-btn:hover { background: rgba(255,255,255,0.3); }
                .ul-panel-content { flex: 1; overflow: hidden; display: flex; flex-direction: column; }
                .ul-panel-toolbar { padding: 8px; border-bottom: 1px solid #e0e0e0; display: flex; gap: 8px; align-items: center; flex-wrap: wrap; background: #f8f9fa; }
                .ul-toolbar-input { padding: 4px 6px; border: 1px solid #ddd; border-radius: 3px; }
                .ul-toolbar-button { padding: 4px 8px; border: 1px solid #ddd; border-radius: 3px; background: white; cursor: pointer; }
                .ul-diagram-container { flex: 1; overflow: auto; padding: 10px; background: #fafafa; }
                .ul-pagination { display: flex; justify-content: center; align-items: center; padding: 4px; border-top: 1px solid #e0e0e0; background: #f8f9fa; user-select: none; }
                .ul-pagination-btn { background: white; border: 1px solid #ddd; border-radius: 3px; padding: 4px 8px; margin: 0 4px; cursor: pointer; font-size: 12px; }
                .ul-pagination-btn:disabled { cursor: not-allowed; opacity: 0.5; }
                .ul-pagination-info { font-size: 12px; margin: 0 10px; font-weight: bold; }
                .ul-panel-status { padding: 6px 12px; background: #f5f5f5; border-top: 1px solid #e0e0e0; font-size: 11px; }
                .ul-loading { text-align: center; padding: 20px; color: #666; }
                .ul-error { color: red; padding: 20px; white-space: pre-wrap; }
            `);
        }

        createPanel() {
            if (this.panel) return;
            this.panel = document.createElement('div');
            this.panel.id = 'ul-standards-panel';
            this.panel.classList.add('hidden'); // Start hidden
            this.panel.innerHTML = `
                <div class="ul-panel-header">
                    <div class="ul-panel-title">UL Standards Navigator v${CONFIG.version}</div>
                    <div class="ul-panel-controls">
                        <button class="ul-panel-btn" id="ul-refresh-btn" title="Refresh scan">↻</button>
                        <button class="ul-panel-btn" id="ul-settings-btn" title="Settings">⚙</button>
                        <button class="ul-panel-btn" id="ul-collapse-btn" title="Collapse/Expand">−</button>
                        <button class="ul-panel-btn" id="ul-close-btn" title="Hide panel">×</button>
                    </div>
                </div>
                <div class="ul-panel-content">
                    <div class="ul-panel-toolbar">
                        <input type="text" class="ul-toolbar-input" id="ul-filter-input" placeholder="Filter...">
                        <button class="ul-toolbar-button" id="ul-upload-csv-btn">Upload CSV</button>
                        <button class="ul-toolbar-button" id="ul-export-btn">Export Page</button>
                    </div>
                    <div class="ul-diagram-container"><div id="ul-diagram" class="ul-loading"><span>Press Ctrl+Shift+Z to activate.</span></div></div>
                    <div class="ul-pagination" id="ul-pagination-controls" style="display: none;">
                        <button id="ul-page-first" class="ul-pagination-btn"><<</button>
                        <button id="ul-page-prev" class="ul-pagination-btn"><</button>
                        <span id="ul-page-info" class="ul-pagination-info">1 / 1</span>
                        <button id="ul-page-next" class="ul-pagination-btn">></button>
                        <button id="ul-page-last" class="ul-pagination-btn">>></button>
                    </div>
                </div>
                <div class="ul-panel-status"><span id="ul-status-text">Inactive</span></div>`;
            document.body.appendChild(this.panel);
            this.attachEventListeners();
            this.makeDraggable();
        }

        attachEventListeners() {
            document.getElementById('ul-refresh-btn').addEventListener('click', () => this.refreshScan());
            document.getElementById('ul-export-btn').addEventListener('click', () => this.exportDiagram());
            document.getElementById('ul-filter-input').addEventListener('input', e => { this.currentFilter = e.target.value; this.refreshScan(); });
            document.getElementById('ul-close-btn').addEventListener('click', () => this.toggleVisibility(false));
            document.getElementById('ul-collapse-btn').addEventListener('click', () => this.toggleCollapse());
            // Settings and CSV upload buttons would be attached here too
            this.attachPaginationListeners();
        }

        attachPaginationListeners() {
            document.getElementById('ul-page-first').addEventListener('click', () => this.goToPage(0));
            document.getElementById('ul-page-prev').addEventListener('click', () => this.goToPage(this.currentPage - 1));
            document.getElementById('ul-page-next').addEventListener('click', () => this.goToPage(this.currentPage + 1));
            document.getElementById('ul-page-last').addEventListener('click', () => this.goToPage(this.diagramChunks.length - 1));
        }

        async activateAndScan() {
            if (!this.isVisible) {
                this.toggleVisibility(true);
            }
            this.refreshScan();
        }

        async updateDiagram() {
            this.showLoading('Generating diagram...');
            try {
                const standards = Array.from(window.ulStandardsExtractor.standards.values());
                if (standards.length === 0) {
                    this.showError("No standards found in this document.");
                    document.getElementById('ul-pagination-controls').style.display = 'none';
                    return;
                }

                this.diagramChunks = window.ulDiagramGenerator.generateFlowchart(standards, this.currentFilter);
                this.currentPage = 0;
                await this.renderCurrentPage();

            } catch (error) {
                Logger.error('Error updating diagram:', error);
                this.showError(`Diagram generation failed: ${error.message}<br><pre>${error.stack}</pre>`);
            }
        }

        async renderCurrentPage() {
            const paginationControls = document.getElementById('ul-pagination-controls');
            if (this.diagramChunks.length > 1) {
                paginationControls.style.display = 'flex';
                document.getElementById('ul-page-info').textContent = `${this.currentPage + 1} / ${this.diagramChunks.length}`;
                document.getElementById('ul-page-first').disabled = this.currentPage === 0;
                document.getElementById('ul-page-prev').disabled = this.currentPage === 0;
                document.getElementById('ul-page-next').disabled = this.currentPage >= this.diagramChunks.length - 1;
                document.getElementById('ul-page-last').disabled = this.currentPage >= this.diagramChunks.length - 1;
            } else {
                paginationControls.style.display = 'none';
            }
            const diagramCode = this.diagramChunks[this.currentPage];
            if (diagramCode) {
                this.updateStatus(`Rendering page ${this.currentPage + 1}...`);
                await window.ulDiagramGenerator.renderDiagram(diagramCode, 'ul-diagram');
                this.updateStatus(`Page ${this.currentPage + 1} of ${this.diagramChunks.length} loaded`);
            } else {
                 this.showError("No data to display for the current filter.");
                 paginationControls.style.display = 'none';
            }
        }

        goToPage(pageNumber) {
            if (pageNumber >= 0 && pageNumber < this.diagramChunks.length) {
                this.currentPage = pageNumber;
                this.renderCurrentPage();
            }
        }

        refreshScan() {
            this.updateStatus('Scanning document...');
            this.showLoading('Scanning... Please wait.');
            setTimeout(() => {
                window.ulStandardsExtractor.extractStandards();
                this.updateDiagram();
            }, 100);
        }

        exportDiagram() {
            const svgElement = document.querySelector('#ul-diagram .mermaid > svg');
            if (!svgElement) { alert('No diagram to export.'); return; }
            const svgData = new XMLSerializer().serializeToString(svgElement);
            const blob = new Blob([svgData], { type: 'image/svg+xml;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = `ul-standards-diagram-page-${this.currentPage + 1}.svg`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
        }

        toggleVisibility(newState) {
            this.isVisible = newState;
            this.panel.classList.toggle('hidden', !this.isVisible);
        }
        toggleCollapse() { this.panel.classList.toggle('collapsed'); }
        showLoading(message) { document.getElementById('ul-diagram').innerHTML = `<div class="ul-loading">${message}</div>`; }
        showError(message) { document.getElementById('ul-diagram').innerHTML = `<div class="ul-error">${message}</div>`; }
        updateStatus(text) { document.getElementById('ul-status-text').textContent = text; }
        makeDraggable() {
            const header = this.panel.querySelector('.ul-panel-header');
            let isDragging = false, offset = { x: 0, y: 0 };
            header.addEventListener('mousedown', e => {
                isDragging = true;
                offset = { x: e.clientX - this.panel.offsetLeft, y: e.clientY - this.panel.offsetTop };
            });
            document.addEventListener('mousemove', e => {
                if (!isDragging) return;
                this.panel.style.left = `${e.clientX - offset.x}px`;
                this.panel.style.top = `${e.clientY - offset.y}px`;
            });
            document.addEventListener('mouseup', () => isDragging = false);
        }
    }

    // Main application class with On-Demand initialization
    class StandardsApp {
        constructor() {
            window.ulStandardsExtractor = new StandardsExtractor();
            window.ulDiagramGenerator = new DiagramGenerator();
            window.ulUIManager = new UIManager();
        }

        initialize() {
            Logger.log('UL Standards Navigator (On-Demand) Initialized.');
            Logger.log('Press Ctrl+Shift+Z to activate and scan.');

            window.ulUIManager.createPanel();
            this.setupActivationListener();
        }

        setupActivationListener() {
            document.addEventListener('keydown', (e) => {
                if (e.ctrlKey && e.shiftKey && (e.key === 'Z' || e.key === 'z')) {
                    e.preventDefault();
                    window.ulUIManager.activateAndScan();
                }
            });
        }
    }

    // Start the application
    if (window.location.href.includes('onlineviewer/')) {
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', () => new StandardsApp().initialize());
        } else {
            new StandardsApp().initialize();
        }
    }
})();
