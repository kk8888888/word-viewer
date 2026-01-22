// Microsoft Word Clone - Interactive Document Viewer
// With Translate and Read Aloud Functions

class WordApp {
    constructor() {
        // DOM Elements
        this.uploadArea = document.getElementById('uploadArea');
        this.documentViewer = document.getElementById('documentViewer');
        this.documentContent = document.getElementById('documentContent');
        this.fileInput = document.getElementById('fileInput');
        this.uploadBtn = document.getElementById('uploadBtn');
        this.browseBtn = document.getElementById('browseBtn');
        this.uploadBox = document.getElementById('uploadBox');
        this.progressText = document.getElementById('progressText');
        this.progressFill = document.getElementById('progressFill');
        this.wordCount = document.getElementById('wordCount');
        this.loadingOverlay = document.getElementById('loadingOverlay');

        // Menu tabs
        this.menuTabs = document.querySelectorAll('.menu-tab');

        // Ribbons map
        this.ribbons = {
            home: document.getElementById('homeRibbon'),
            insert: document.getElementById('insertRibbon'),
            draw: document.getElementById('drawRibbon'),
            design: document.getElementById('designRibbon'),
            layout: document.getElementById('layoutRibbon'),
            review: document.getElementById('reviewRibbon'),
            file: document.getElementById('fileRibbon')
        };

        // File ribbon buttons
        this.exportPdfBtn = document.getElementById('exportPdfBtn');
        this.exportDocxBtn = document.getElementById('exportDocxBtn');
        this.openFileBtn = document.getElementById('openFileBtn');

        // Review ribbon buttons
        this.openDocBtn = document.getElementById('openDocBtn');
        this.resetDocBtn = document.getElementById('resetDocBtn');
        this.translateDocBtn = document.getElementById('translateDocBtn');
        this.readAloudBtn = document.getElementById('readAloudBtn');

        // Selection popup
        this.selectionPopup = document.getElementById('selectionPopup');
        this.toChineseBtn = document.getElementById('toChineseBtn');
        this.toEnglishBtn = document.getElementById('toEnglishBtn');

        // Translate panel
        this.translatePanel = document.getElementById('translatePanel');
        this.closeTranslatePanel = document.getElementById('closeTranslatePanel');
        this.originalText = document.getElementById('originalText');
        this.translatedText = document.getElementById('translatedText');
        this.targetLanguage = document.getElementById('targetLanguage');
        this.sourceLanguage = document.getElementById('sourceLanguage');
        this.doTranslate = document.getElementById('doTranslate');
        this.readTranslationBtn = document.getElementById('readTranslationBtn');

        // Read panel
        this.readPanel = document.getElementById('readPanel');
        this.closeReadPanel = document.getElementById('closeReadPanel');
        this.playPauseBtn = document.getElementById('playPauseBtn');
        this.stopBtn = document.getElementById('stopBtn');
        this.voiceSelect = document.getElementById('voiceSelect');
        this.speedSlider = document.getElementById('speedSlider');
        this.speedValue = document.getElementById('speedValue');
        this.readingText = document.getElementById('readingText');

        // Dictionary panel
        this.dictionaryBtn = document.getElementById('dictionaryBtn');
        this.dictionaryPanel = document.getElementById('dictionaryPanel');
        this.closeDictionaryPanel = document.getElementById('closeDictionaryPanel');
        this.dictionaryContent = document.getElementById('dictionaryContent');
        this.quickSaveBtn = document.getElementById('quickSaveBtn');

        // Clipboard buttons
        this.cutBtn = document.getElementById('cutBtn');
        this.copyBtn = document.getElementById('copyBtn');
        this.pasteBtn = document.getElementById('pasteBtn');

        // Font formatting buttons
        this.boldBtn = document.getElementById('boldBtn');
        this.italicBtn = document.getElementById('italicBtn');
        this.underlineBtn = document.getElementById('underlineBtn');
        this.strikeBtn = document.getElementById('strikeBtn');
        this.fontSelect = document.getElementById('fontSelect');
        this.fontSizeSelect = document.getElementById('fontSizeSelect');
        this.fontSizeUpBtn = document.getElementById('fontSizeUpBtn');
        this.fontSizeDownBtn = document.getElementById('fontSizeDownBtn');
        this.highlightBtn = document.getElementById('highlightBtn');
        this.fontColorBtn = document.getElementById('fontColorBtn');

        // Insert Ribbon
        this.insertPageBtn = document.getElementById('insertPageBtn');
        this.insertBreakBtn = document.getElementById('insertBreakBtn');
        this.insertTableBtn = document.getElementById('insertTableBtn');
        this.insertPicBtn = document.getElementById('insertPicBtn');
        this.imageInput = document.getElementById('imageInput');
        this.finalizeEditBtn = document.getElementById('finalizeEditBtn');
        this.insertLinkBtn = document.getElementById('insertLinkBtn');
        this.insertHeaderBtn = document.getElementById('insertHeaderBtn');
        this.insertFooterBtn = document.getElementById('insertFooterBtn');
        this.insertPageNumBtn = document.getElementById('insertPageNumBtn');

        // Draw Ribbon
        this.drawPenBlackBtn = document.getElementById('drawPenBlackBtn');
        this.drawPenRedBtn = document.getElementById('drawPenRedBtn');
        this.drawHighlighterBtn = document.getElementById('drawHighlighterBtn');

        // Design Ribbon
        this.designPageColorBtn = document.getElementById('designPageColorBtn');
        this.colorInput = document.getElementById('colorInput');

        // Layout Ribbon
        this.layoutMarginsBtn = document.getElementById('layoutMarginsBtn');
        this.layoutOrientBtn = document.getElementById('layoutOrientBtn');
        this.layoutSizeBtn = document.getElementById('layoutSizeBtn');
        this.layoutColsBtn = document.getElementById('layoutColsBtn');

        // Paragraph buttons
        this.alignLeftBtn = document.getElementById('alignLeftBtn');
        this.alignCenterBtn = document.getElementById('alignCenterBtn');
        this.alignRightBtn = document.getElementById('alignRightBtn');
        this.alignJustifyBtn = document.getElementById('alignJustifyBtn');

        // State
        this.documentText = '';
        this.revealedChars = 0;
        this.totalChars = 0;
        this.isDocumentLoaded = false;
        this.selectedText = '';
        this.speechSynthesis = window.speechSynthesis;
        this.currentUtterance = null;
        this.isPlaying = false;

        // Initialize
        this.checkDependencies();
        this.initPdfJs();
        this.bindEvents();
        this.loadVoices();
    }

    initPdfJs() {
        if (typeof pdfjsLib !== 'undefined') {
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        }
    }

    checkDependencies() {
        if (typeof mammoth === 'undefined') {
            alert('Warning: Mammoth.js library not loaded. .docx files will not work. Please check your internet connection.');
        }
        if (typeof pdfjsLib === 'undefined') {
            alert('Warning: PDF.js library not loaded. .pdf files will not work. Please check your internet connection.');
        }
    }

    bindEvents() {
        // File upload
        this.browseBtn?.addEventListener('click', () => this.fileInput.click());
        this.uploadBtn?.addEventListener('click', () => this.fileInput.click());
        this.openDocBtn?.addEventListener('click', () => this.fileInput.click());
        this.fileInput?.addEventListener('change', (e) => this.handleFileSelect(e));

        // Drag and drop
        this.uploadBox?.addEventListener('dragover', (e) => this.handleDragOver(e));
        this.uploadBox?.addEventListener('dragleave', (e) => this.handleDragLeave(e));
        this.uploadBox?.addEventListener('drop', (e) => this.handleDrop(e));

        // Reset
        this.resetDocBtn?.addEventListener('click', () => this.resetDocument());

        // Menu tabs
        this.menuTabs.forEach(tab => {
            tab.addEventListener('click', (e) => this.switchTab(e.target));
        });

        // Insert events
        this.insertPageBtn?.addEventListener('click', () => this.insertBlankPage());
        this.insertBreakBtn?.addEventListener('click', () => this.insertPageBreak());
        this.insertTableBtn?.addEventListener('click', () => this.insertTable());
        this.insertPicBtn?.addEventListener('click', () => this.imageInput.click());
        this.imageInput?.addEventListener('change', (e) => this.handleImageUpload(e));
        this.insertLinkBtn?.addEventListener('click', () => this.insertLink());

        // Draw events
        this.drawPenBlackBtn?.addEventListener('click', () => this.execCmd('foreColor', '#000000'));
        this.drawPenRedBtn?.addEventListener('click', () => this.execCmd('foreColor', '#FF0000'));
        // Highlighter maps to hiliteColor
        this.drawHighlighterBtn?.addEventListener('click', () => this.execCmd('hiliteColor', 'yellow'));

        // Design events
        this.designPageColorBtn?.addEventListener('click', () => this.colorInput.click());
        this.colorInput?.addEventListener('change', (e) => this.handlePageColor(e));

        // Layout events
        this.layoutMarginsBtn?.addEventListener('click', () => this.toggleMargins());
        this.layoutOrientBtn?.addEventListener('click', () => this.toggleOrientation());

        // Keyboard events
        document.addEventListener('keydown', (e) => this.handleKeyPress(e));

        // Text selection
        document.addEventListener('mouseup', (e) => this.handleTextSelection(e));
        document.addEventListener('mousedown', (e) => {
            // Don't hide popup if clicking on the popup itself or translation result
            const popup = this.selectionPopup;
            const result = document.querySelector('.translation-result');
            if (popup && popup.contains(e.target)) return;
            if (result && result.contains(e.target)) return;
            this.hideSelectionPopup();
        });

        // Translate Document
        this.translateDocBtn?.addEventListener('click', () => {
            const lang = prompt(
                "Translate entire document:\n\nEnter target language code:\nzh - Chinese\nms - Malay\nes - Spanish\nfr - French\nde - German\nja - Japanese\nko - Korean",
                "zh"
            );
            if (lang) {
                this.translateWholeDocument(lang.toLowerCase());
            }
        });

        // File ribbon events
        this.exportPdfBtn?.addEventListener('click', () => this.exportPDF());
        this.exportDocxBtn?.addEventListener('click', () => this.exportDOCX());
        this.openFileBtn?.addEventListener('click', () => this.fileInput.click());

        // Selection popup - Direct translation
        this.toChineseBtn?.addEventListener('click', (e) => {
            e.stopPropagation();
            this.translateDirect('zh');
        });
        this.toEnglishBtn?.addEventListener('click', (e) => {
            e.stopPropagation();
            this.translateDirect('en');
        });

        // Save
        this.quickSaveBtn?.addEventListener('click', () => this.saveDocument());

        // Translate panel
        this.closeTranslatePanel?.addEventListener('click', () => this.closePanel('translate'));
        this.doTranslate?.addEventListener('click', () => this.translateText());
        this.translateDocBtn?.addEventListener('click', () => this.openTranslatePanel());
        this.readTranslationBtn?.addEventListener('click', () => this.readTranslatedText());

        // Read panel
        this.closeReadPanel?.addEventListener('click', () => this.closePanel('read'));
        this.closeDictionaryPanel?.addEventListener('click', () => this.closePanel('dictionary'));
        this.readAloudBtn?.addEventListener('click', () => this.openReadPanel());
        this.playPauseBtn?.addEventListener('click', () => this.togglePlayPause());
        this.stopBtn?.addEventListener('click', () => this.stopSpeaking());
        this.speedSlider?.addEventListener('input', (e) => this.updateSpeed(e));
        this.finalizeEditBtn?.addEventListener('click', () => this.finalizeDocument());
        this.targetLanguage?.addEventListener('change', () => this.autoSelectVoiceForLang());

        // Document content changes
        this.documentContent?.addEventListener('input', () => this.updateWordCount());

        // Clipboard functions
        this.cutBtn?.addEventListener('click', () => this.cutText());
        this.copyBtn?.addEventListener('click', () => this.copyText());
        this.pasteBtn?.addEventListener('click', () => this.pasteText());

        // Font formatting
        this.boldBtn?.addEventListener('click', () => this.formatText('bold'));
        this.italicBtn?.addEventListener('click', () => this.formatText('italic'));
        this.underlineBtn?.addEventListener('click', () => this.formatText('underline'));
        this.strikeBtn?.addEventListener('click', () => this.formatText('strikeThrough'));
        this.fontSelect?.addEventListener('change', (e) => this.changeFont(e.target.value));
        this.fontSizeSelect?.addEventListener('change', (e) => this.changeFontSize(e.target.value));
        this.fontSizeUpBtn?.addEventListener('click', () => this.changeFontSizeRelative(2));
        this.fontSizeDownBtn?.addEventListener('click', () => this.changeFontSizeRelative(-2));
        this.highlightBtn?.addEventListener('click', () => this.highlightText());
        this.fontColorBtn?.addEventListener('click', () => this.changeFontColor());

        // Paragraph alignment
        this.alignLeftBtn?.addEventListener('click', () => this.alignText('left'));
        this.alignCenterBtn?.addEventListener('click', () => this.alignText('center'));
        this.alignRightBtn?.addEventListener('click', () => this.alignText('right'));
        this.alignJustifyBtn?.addEventListener('click', () => this.alignText('justify'));

        // Keyboard shortcuts
        document.addEventListener('keydown', (e) => {
            if (e.ctrlKey || e.metaKey) {
                switch (e.key.toLowerCase()) {
                    case 'b':
                        e.preventDefault();
                        this.formatText('bold');
                        break;
                    case 'i':
                        e.preventDefault();
                        this.formatText('italic');
                        break;
                    case 'u':
                        e.preventDefault();
                        this.formatText('underline');
                        break;
                    case 'c':
                        // Allow default copy
                        break;
                    case 'v':
                        // Allow default paste
                        break;
                    case 'x':
                        // Allow default cut
                        break;
                }
            }
        });
    }

    loadVoices() {
        const loadVoiceList = () => {
            const voices = this.speechSynthesis.getVoices();
            if (this.voiceSelect && voices.length > 0) {
                this.voiceSelect.innerHTML = '';
                voices.forEach((voice, index) => {
                    const option = document.createElement('option');
                    option.value = index;
                    option.textContent = `${voice.name} (${voice.lang})`;
                    if (voice.default) option.selected = true;
                    this.voiceSelect.appendChild(option);
                });
            }
        };

        loadVoiceList();
        if (this.speechSynthesis.onvoiceschanged !== undefined) {
            this.speechSynthesis.onvoiceschanged = loadVoiceList;
        }
    }

    // --- New Ribbon Functions ---

    insertBlankPage() {
        this.documentContent.focus();
        // Insert a visual page break
        const pageHtml = '<div class="page-break" contenteditable="false"></div><p><br></p>';
        document.execCommand('insertHTML', false, pageHtml);
    }

    insertPageBreak() {
        this.documentContent.focus();
        const breakHtml = '<div style="page-break-after: always; border-bottom: 2px dashed #999; margin: 20px 0; text-align: center; color: #999; font-size: 12px;">--- Page Break ---</div><p><br></p>';
        document.execCommand('insertHTML', false, breakHtml);
    }

    insertTable() {
        this.documentContent.focus();
        // Basic 3x3 table
        const tableHtml = `
            <table style="width: 100%; border-collapse: collapse; margin: 10px 0; border: 1px solid #000;">
                <tr>
                    <td style="border: 1px solid #ccc; padding: 8px;">Cell 1</td>
                    <td style="border: 1px solid #ccc; padding: 8px;">Cell 2</td>
                    <td style="border: 1px solid #ccc; padding: 8px;">Cell 3</td>
                </tr>
                <tr>
                    <td style="border: 1px solid #ccc; padding: 8px;">&nbsp;</td>
                    <td style="border: 1px solid #ccc; padding: 8px;">&nbsp;</td>
                    <td style="border: 1px solid #ccc; padding: 8px;">&nbsp;</td>
                </tr>
                <tr>
                    <td style="border: 1px solid #ccc; padding: 8px;">&nbsp;</td>
                    <td style="border: 1px solid #ccc; padding: 8px;">&nbsp;</td>
                    <td style="border: 1px solid #ccc; padding: 8px;">&nbsp;</td>
                </tr>
            </table><p><br></p>
        `;
        document.execCommand('insertHTML', false, tableHtml);
    }

    handleImageUpload(e) {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = (event) => {
                this.documentContent.focus();
                // Resize for sanity
                const imgHtml = `<img src="${event.target.result}" style="max-width: 100%; height: auto; margin: 10px 0;">`;
                document.execCommand('insertHTML', false, imgHtml);
            };
            reader.readAsDataURL(file);
        }
        // Reset input
        e.target.value = '';
    }

    insertLink() {
        const url = prompt('Enter the Link URL:', 'https://');
        if (url) {
            this.documentContent.focus();
            document.execCommand('createLink', false, url);
        }
    }

    handlePageColor(e) {
        const color = e.target.value;
        const page = document.getElementById('documentPage');
        if (page) {
            page.style.backgroundColor = color;
        }
    }

    toggleMargins() {
        const page = document.getElementById('documentPage');
        if (page) {
            page.classList.toggle('narrow-margins');
            this.showNotification(page.classList.contains('narrow-margins') ? 'Margins: Narrow' : 'Margins: Normal');
        }
    }

    toggleOrientation() {
        const page = document.getElementById('documentPage');
        if (page) {
            page.classList.toggle('landscape');
            this.showNotification(page.classList.contains('landscape') ? 'Orientation: Landscape' : 'Orientation: Portrait');
        }
    }

    switchTab(tab) {
        const tabText = tab.textContent.trim().toLowerCase();

        // Deactivate all tabs
        this.menuTabs.forEach(t => t.classList.remove('active'));
        // Activate clicked tab
        tab.classList.add('active');

        // Hide all ribbons
        Object.values(this.ribbons).forEach(ribbon => {
            ribbon?.classList.add('hidden');
        });

        // Show selected ribbon
        if (this.ribbons[tabText]) {
            this.ribbons[tabText].classList.remove('hidden');
        } else {
            // Default to home if ribbon not found
            this.ribbons.home?.classList.remove('hidden');
        }
    }

    exportPDF() {
        if (!this.documentContent) return;

        this.showLoading('Generating PDF...');

        const opt = {
            margin: [20, 15, 20, 15],
            filename: 'translated_document.pdf',
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: {
                scale: 2,
                useCORS: true,
                logging: false,
                letterRendering: true,
                allowTaint: true
            },
            jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
        };

        const cleanElement = this.getCleanContent();
        document.body.appendChild(cleanElement);

        if (typeof html2pdf !== 'undefined') {
            html2pdf().set(opt).from(cleanElement).save().then(() => {
                this.hideLoading();
                this.showNotification('PDF exported successfully!');
                cleanElement.remove();
            }).catch(err => {
                console.error(err);
                this.hideLoading();
                this.showNotification('Failed to export PDF');
                cleanElement.remove();
            });
        } else {
            this.hideLoading();
            alert('PDF library not loaded.');
            cleanElement.remove();
        }
    }

    getCleanContent() {
        // Creates a high-fidelity clone focused on pure text for engine compatibility
        const clone = this.documentContent.cloneNode(true);

        // CRITICAL FIX: Collapse spans back into text nodes for the export engine
        // html2canvas struggles with thousands of individual character spans
        this.collapseChars(clone);

        // Force layout and visibility
        clone.style.display = 'block';
        clone.style.visibility = 'visible';
        clone.style.opacity = '1';
        clone.style.whiteSpace = 'pre-wrap';
        clone.style.wordWrap = 'break-word';
        clone.style.color = '#000000';
        clone.style.backgroundColor = '#ffffff';
        clone.style.fontFamily = "'Times New Roman', Times, serif";
        clone.style.fontSize = '12pt';
        clone.style.lineHeight = '1.6';
        clone.style.padding = '40px';
        clone.style.width = '700px';

        // Use absolute positioning highly offset to ensure it doesn't affect viewport
        // but is still "scannable" by html2canvas
        clone.style.position = 'absolute';
        clone.style.top = '-10000px';
        clone.style.left = '0';
        clone.style.zIndex = '-9999';

        return clone;
    }

    handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();
        this.uploadBox.style.borderColor = '#2b579a';
        this.uploadBox.style.background = '#f0f7ff';
    }

    handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();
        this.uploadBox.style.borderColor = '';
        this.uploadBox.style.background = '';
    }

    handleDrop(e) {
        e.preventDefault();
        e.stopPropagation();
        this.uploadBox.style.borderColor = '';
        this.uploadBox.style.background = '';

        const files = e.dataTransfer.files;
        if (files.length > 0) {
            this.processFile(files[0]);
        }
    }

    handleFileSelect(e) {
        // alert('File selected!'); // Debug
        const file = e.target.files[0];
        if (file) {
            this.processFile(file);
        }
    }

    async processFile(file) {
        const fileName = file.name.toLowerCase();

        if (!fileName.endsWith('.pdf') && !fileName.endsWith('.docx') && !fileName.endsWith('.doc')) {
            alert('Please upload a PDF or Word document (.pdf, .docx, .doc)');
            return;
        }

        this.showLoading();

        try {
            let text = '';

            if (fileName.endsWith('.pdf')) {
                text = await this.extractPdfText(file);
                // PDF is text-only for now
                await this.loadDocument(text, false);
            } else if (fileName.endsWith('.docx')) {
                const html = await this.extractDocxHtml(file);
                await this.loadDocument(html, true);
            } else if (fileName.endsWith('.doc')) {
                text = 'Note: Legacy .doc files have limited support. Please use .docx or .pdf format for best results.';
                await this.loadDocument(text, false);
            }

            // Title update moved inside loadDocument or here
            document.querySelector('.doc-title').textContent = file.name + ' - Word';
            document.title = file.name + ' - Word';

        } catch (error) {
            console.error('Error processing file:', error);
            alert('Error processing file: ' + error.message);
            this.hideLoading(); // Ensure loading hides on error
        } finally {
            // this.hideLoading(); // Moved inside to prevent premature hiding if async
        }
    }

    async extractPdfText(file) {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        let fullText = '';

        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            const pageText = textContent.items.map(item => item.str).join(' ');
            fullText += pageText + '\n\n';
        }

        return fullText;
    }

    async extractDocxHtml(file) {
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.convertToHtml({ arrayBuffer });
        return result.value;
    }

    async loadDocument(content, isHtml = false) {
        this.isDocumentLoaded = true;

        // Show loading state
        this.showLoading();

        if (isHtml) {
            // fast load for HTML content (DOCX)
            this.documentContent.innerHTML = content;

            // OPTIMIZED: Chunked wrapping to avoid freezing UI
            this.showLoading('Preparing document reveal...');
            this.wrapCharsInElementAsync(this.documentContent, () => {
                this.documentText = this.documentContent.innerText;
                this.totalChars = this.documentContent.querySelectorAll('.char').length;
                this.revealedChars = 0;

                this.updateWordCount();
                this.hideLoading();
                this.uploadArea.classList.add('hidden');
                this.documentViewer.classList.add('active');
            });
        } else {
            // "Type to reveal" mode for plain text / PDF
            this.documentText = content;
            this.totalChars = content.length;
            this.revealedChars = 0;
            this.documentContent.innerHTML = '';

            // Process in chunks to avoid freezing UI
            const chunkSize = 2000;
            let currentIndex = 0;

            const processChunk = () => {
                return new Promise((resolve) => {
                    const fragment = document.createDocumentFragment();
                    const end = Math.min(currentIndex + chunkSize, content.length);

                    for (let i = currentIndex; i < end; i++) {
                        const char = content[i];
                        const span = document.createElement('span');
                        span.className = 'char';
                        // Preserve line breaks
                        if (char === '\n') {
                            span.textContent = '';
                            span.classList.add('newline');
                            fragment.appendChild(span);
                            fragment.appendChild(document.createElement('br'));
                        } else {
                            span.textContent = char;
                            fragment.appendChild(span);
                        }
                    }
                    this.documentContent.appendChild(fragment);
                    resolve();
                });
            };

            const run = async () => {
                while (currentIndex < content.length) {
                    await processChunk();
                    currentIndex += chunkSize;
                    // Yield to main thread
                    await new Promise(r => setTimeout(r, 0));

                    // Update progress
                    const percent = Math.min(100, Math.round((currentIndex / content.length) * 100));
                    if (this.progressText) this.progressText.textContent = percent + '%';
                    if (this.progressFill) this.progressFill.style.width = percent + '%';
                }
                this.hideLoading();
                this.updateWordCount();
                this.uploadArea.classList.add('hidden');
                this.documentViewer.classList.add('active');
            };

            run();
        }
    }

    handleKeyPress(e) {
        // Ignore if typing in input fields
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.tagName === 'SELECT') {
            return;
        }

        // Ignore modifier keys
        if (e.ctrlKey || e.altKey || e.metaKey) return;
        if (e.key.length > 1 && e.key !== 'Backspace') return;

        // REVEAL MODE: If document is loaded and has hidden chars
        if (this.isDocumentLoaded && this.revealedChars < this.totalChars) {
            e.preventDefault(); // Prevent inserting the typed char

            if (e.key === 'Backspace') {
                if (this.revealedChars > 0) {
                    this.revealedChars--;
                    const chars = this.documentContent.querySelectorAll('.char');
                    if (chars[this.revealedChars]) {
                        chars[this.revealedChars].classList.remove('revealed');
                    }
                    this.updateProgress();
                }
            } else {
                this.revealNextChars(1);
            }
        }
        // EDIT MODE: If text is fully revealed or just editing normally
        else {
            // Allow default typing behavior
            // We'll update word count on 'input' event listener instead
        }
    }

    revealNextChars(count) {
        const chars = this.documentContent.querySelectorAll('.char');

        for (let i = 0; i < count && this.revealedChars < this.totalChars; i++) {
            if (chars[this.revealedChars]) {
                chars[this.revealedChars].classList.add('revealed');
                this.revealedChars++;

                // Auto-reveal whitespace
                while (this.revealedChars < this.totalChars) {
                    const nextChar = this.documentText[this.revealedChars];
                    if (nextChar === ' ' || nextChar === '\n' || nextChar === '\r' || nextChar === '\t') {
                        chars[this.revealedChars].classList.add('revealed');
                        this.revealedChars++;
                    } else {
                        break;
                    }
                }
            }
        }

        this.updateProgress();
    }

    updateProgress() {
        const percentage = this.totalChars > 0 ? Math.round((this.revealedChars / this.totalChars) * 100) : 100;
        if (this.progressText) this.progressText.textContent = percentage + '%';
        if (this.progressFill) this.progressFill.style.width = percentage + '%';
    }

    updateWordCount() {
        // Use innerText to count words of what's actually in buffer
        const text = this.documentContent.innerText || '';
        const words = text.split(/\s+/).filter(w => w.length > 0).length;
        if (this.wordCount) {
            this.wordCount.textContent = words + ' words';
        }
    }

    resetDocument() {
        if (!this.isDocumentLoaded) return;

        this.revealedChars = 0;
        const chars = this.documentContent.querySelectorAll('.char');
        chars.forEach(char => char.classList.remove('revealed'));
        this.updateProgress();
    }

    handleTextSelection(e) {
        // Allow selection even before document is loaded for testing
        setTimeout(() => {
            const selection = window.getSelection();
            const selectedText = selection.toString().trim();

            if (selectedText.length > 0) {
                this.selectedText = selectedText;
                this.showSelectionPopup(e);
            }
        }, 50);
    }

    showSelectionPopup(e) {
        const selection = window.getSelection();
        if (selection.rangeCount === 0) return;

        const range = selection.getRangeAt(0);
        const rect = range.getBoundingClientRect();

        // Position popup above the selection
        const popupWidth = 280;
        const popupX = rect.left + (rect.width / 2) - (popupWidth / 2);
        const popupY = rect.top - 50;

        this.selectionPopup.style.left = Math.max(10, popupX) + 'px';
        this.selectionPopup.style.top = Math.max(10, popupY) + 'px';
        this.selectionPopup.classList.add('active');

        console.log('Popup shown for text:', this.selectedText);
    }

    hideSelectionPopup() {
        this.selectionPopup?.classList.remove('active');
        // Remove any existing translation result
        const existingResult = document.querySelector('.translation-result');
        if (existingResult) existingResult.remove();
    }

    async translateDirect(targetLang) {
        const text = this.selectedText;
        if (!text) {
            alert('Please select text first');
            return;
        }

        this.showTranslationResult('Translating...');

        try {
            const result = await this.performTranslation(text, 'auto', targetLang);
            this.showTranslationResult(result);
        } catch (error) {
            console.error('Translation error:', error);
            const errorMsg = error.message.includes('429') ? 'Rate limit exceeded. Please wait a minute.' : 'Translation failed. Try again.';
            this.showTranslationResult(errorMsg);
        }
    }

    async performTranslation(text, source, target) {
        // Prepare text in chunks if it's very long (Google limit is ~2000 chars, MyMemory ~500)
        const CHUNK_LIMIT = 450;
        if (text.length > CHUNK_LIMIT) {
            const sentences = text.match(/[^.!?]+[.!?]+/g) || [text];
            let results = [];
            let currentChunk = "";

            for (let sentence of sentences) {
                if ((currentChunk + sentence).length > CHUNK_LIMIT) {
                    results.push(await this.callTranslationAPIs(currentChunk, source, target));
                    currentChunk = sentence;
                } else {
                    currentChunk += sentence;
                }
            }
            if (currentChunk) results.push(await this.callTranslationAPIs(currentChunk, source, target));
            return results.join(" ");
        }

        return await this.callTranslationAPIs(text, source, target);
    }

    async callTranslationAPIs(text, source, target) {
        const sourceCode = source === 'auto' ? 'auto' : source;

        // Service 1: Google Translate (Unofficial GTX) - Most Reliable
        try {
            const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${sourceCode}&tl=${target}&dt=t&q=${encodeURIComponent(text)}`;
            const response = await fetch(url);
            if (response.ok) {
                const data = await response.json();
                if (data && data[0]) {
                    return data[0].map(x => x[0]).join('');
                }
            }
        } catch (e) {
            console.warn('Google Translate failed...', e);
        }

        // Service 2: MyMemory (Primary)
        try {
            const langPair = source === 'auto' ? `auto|${target}` : `${source}|${target}`;
            const response = await fetch(`https://api.mymemory.translated.net/get?q=${encodeURIComponent(text)}&langpair=${langPair}`);
            if (response.ok) {
                const data = await response.json();
                if (data.responseStatus === 200 && data.responseData) {
                    return data.responseData.translatedText;
                }
            }
        } catch (e) {
            console.warn('MyMemory failed...', e);
        }

        // Service 3: Lingva (Fallback)
        try {
            const response = await fetch(`https://lingva.ml/api/v1/${sourceCode}/${target}/${encodeURIComponent(text)}`);
            if (response.ok) {
                const data = await response.json();
                if (data.translation) return data.translation;
            }
        } catch (e) {
            console.warn('Lingva failed...', e);
        }

        throw new Error('All translation services are currently unavailable.');
    }

    showTranslationResult(text) {
        // Remove existing result
        const existingResult = document.querySelector('.translation-result');
        if (existingResult) existingResult.remove();

        // Create result popup
        const resultDiv = document.createElement('div');
        resultDiv.className = 'translation-result';
        resultDiv.innerHTML = `
            <div class="result-header">
                <span>Translation</span>
                <button class="result-close">×</button>
            </div>
            <div class="result-text">${text}</div>
            <div class="result-actions">
                <button class="result-btn copy-btn">📋 Copy</button>
                <button class="result-btn read-btn">🔊 Read</button>
            </div>
        `;

        // Position near selection popup
        const popup = this.selectionPopup;
        if (popup) {
            resultDiv.style.left = popup.style.left;
            resultDiv.style.top = (parseInt(popup.style.top) + 45) + 'px';
        }

        document.body.appendChild(resultDiv);

        // Close button
        resultDiv.querySelector('.result-close').addEventListener('click', () => {
            resultDiv.remove();
        });

        // Copy button
        resultDiv.querySelector('.copy-btn').addEventListener('click', () => {
            navigator.clipboard.writeText(text);
            const btn = resultDiv.querySelector('.copy-btn');
            const originalText = btn.textContent;
            btn.textContent = '✓ Copied!';
            setTimeout(() => {
                btn.textContent = originalText;
            }, 2000);
        });

        // Read button
        resultDiv.querySelector('.read-btn').addEventListener('click', () => {
            // Stop any current speech
            this.stopSpeaking();

            // Speak the translated text
            const utterance = new SpeechSynthesisUtterance(text);
            utterance.rate = 1;

            // Try to set correct voice language based on text content
            // Simple heuristic: if Chinese chars, use 'zh', else 'en'
            const isChinese = /[\u4e00-\u9fa5]/.test(text);
            const langPrefix = isChinese ? 'zh' : 'en';

            const voices = this.speechSynthesis.getVoices();
            const voice = voices.find(v => v.lang.startsWith(langPrefix));
            if (voice) {
                utterance.voice = voice;
            }

            this.speechSynthesis.speak(utterance);
        });

        // Auto hide after 10 seconds
        setTimeout(() => {
            if (resultDiv.parentNode) resultDiv.remove();
        }, 10000);
    }

    openTranslatePanel() {
        this.hideSelectionPopup();
        if (this.originalText) {
            this.originalText.textContent = this.selectedText || 'Select text to translate';
        }
        if (this.translatedText) {
            this.translatedText.textContent = '';
        }
        this.translatePanel?.classList.add('active');
        this.readPanel?.classList.remove('active');
    }

    openReadPanel() {
        this.hideSelectionPopup();

        // Use selected text or fall back to full document content
        const fullText = this.documentContent?.innerText || '';
        const textToRead = this.selectedText || fullText;

        if (this.readingText) {
            if (this.selectedText) {
                this.readingText.textContent = this.selectedText;
            } else if (fullText.trim().length > 0) {
                // For "Read Aloud" on full doc, we read what's revealed
                this.readingText.textContent = fullText;
            } else {
                this.readingText.textContent = 'No text found to read.';
            }
        }

        this.readPanel?.classList.add('active');
        this.translatePanel?.classList.remove('active');

        // Auto-start reading when panel opens for convenience
        setTimeout(() => this.startSpeaking(), 300);
    }

    closePanel(panel) {
        if (panel === 'translate') {
            this.translatePanel?.classList.remove('active');
        } else if (panel === 'read') {
            this.readPanel?.classList.remove('active');
            this.stopSpeaking();
        } else if (panel === 'dictionary') {
            this.dictionaryPanel?.classList.remove('active');
        }
    }

    async translateText() {
        const text = this.originalText?.textContent || this.selectedText;
        let sourceLang = this.sourceLanguage?.value || 'auto';
        const targetLang = this.targetLanguage?.value || 'ms';

        if (!text || text === 'Select text to translate') {
            alert('Please select text to translate first.');
            return;
        }

        if (this.translatedText) {
            this.translatedText.textContent = 'Translating...';
        }

        try {
            const result = await this.performTranslation(text, sourceLang, targetLang);
            if (this.translatedText) {
                this.translatedText.textContent = result;
            }
        } catch (error) {
            console.error('Translation error:', error);
            if (this.translatedText) {
                const errorMsg = error.message.includes('429') ? 'Rate limit exceeded. Please wait a minute.' : 'Translation service busy. Please try again in 10 seconds.';
                this.translatedText.textContent = errorMsg;
            }
        }
    }

    readTranslatedText() {
        const text = this.translatedText?.textContent;
        if (!text || text === 'Translation service temporarily unavailable. Please try again.' || text === 'Translating...') {
            alert('Please translate some text first.');
            return;
        }

        // Stop any current speech
        this.stopSpeaking();

        // Speak
        const utterance = new SpeechSynthesisUtterance(text);
        utterance.rate = parseFloat(this.speedSlider?.value || 1);

        // Auto-select voice for target language
        const targetLang = this.targetLanguage?.value;
        const voices = this.speechSynthesis.getVoices();

        if (targetLang && voices.length > 0) {
            // Find best voice: priority 1: Exact match (e.g. zh-CN), priority 2: Starts with (e.g. zh)
            let voice = voices.find(v => v.lang.toLowerCase() === targetLang.toLowerCase());
            if (!voice) voice = voices.find(v => v.lang.toLowerCase().startsWith(targetLang.toLowerCase()));

            if (voice) {
                utterance.voice = voice;
                // Update the dropdown UI for visual feedback
                const voiceIndex = voices.indexOf(voice);
                if (this.voiceSelect) this.voiceSelect.value = voiceIndex;
            }
        }

        this.currentUtterance = utterance;
        this.currentUtterance.onend = () => {
            this.isPlaying = false;
        };

        this.speechSynthesis.speak(utterance);
        this.isPlaying = true;
    }

    autoSelectVoiceForLang() {
        const targetLang = this.targetLanguage?.value;
        const voices = this.speechSynthesis.getVoices();
        if (!targetLang || voices.length === 0) return;

        let voice = voices.find(v => v.lang.toLowerCase() === targetLang.toLowerCase());
        if (!voice) voice = voices.find(v => v.lang.toLowerCase().startsWith(targetLang.toLowerCase()));

        if (voice && this.voiceSelect) {
            const voiceIndex = voices.indexOf(voice);
            this.voiceSelect.value = voiceIndex;
        }
    }

    togglePlayPause() {
        if (this.isPlaying) {
            this.pauseSpeaking();
        } else {
            this.startSpeaking();
        }
    }

    startSpeaking() {
        // Use the text currently visible in the reading panel
        let text = this.readingText?.textContent;

        // If the panel is empty or placeholder, try to get from selection or document
        if (!text || text === 'Select text to read aloud' || text === 'No text found to read.') {
            text = this.selectedText || this.documentContent?.innerText || '';
        }

        if (!text || text.trim().length === 0) {
            alert('No text found to read. Please reveal some content or select text first.');
            return;
        }

        this.speechSynthesis.cancel();

        this.currentUtterance = new SpeechSynthesisUtterance(text);
        this.currentUtterance.rate = parseFloat(this.speedSlider?.value || 1);

        const voices = this.speechSynthesis.getVoices();
        const selectedIndex = this.voiceSelect?.value;
        if (selectedIndex && voices[selectedIndex]) {
            this.currentUtterance.voice = voices[selectedIndex];
        }

        this.currentUtterance.onend = () => {
            this.isPlaying = false;
            if (this.playPauseBtn) this.playPauseBtn.textContent = '▶ Play';
        };

        this.speechSynthesis.speak(this.currentUtterance);
        this.isPlaying = true;
        if (this.playPauseBtn) this.playPauseBtn.textContent = '⏸ Pause';
    }

    pauseSpeaking() {
        if (this.speechSynthesis.speaking) {
            this.speechSynthesis.pause();
            this.isPlaying = false;
            if (this.playPauseBtn) this.playPauseBtn.textContent = '▶ Play';
        }
    }

    stopSpeaking() {
        this.speechSynthesis.cancel();
        this.isPlaying = false;
        if (this.playPauseBtn) this.playPauseBtn.textContent = '▶ Play';
    }

    updateSpeed(e) {
        const speed = parseFloat(e.target.value);
        if (this.speedValue) {
            this.speedValue.textContent = speed.toFixed(1) + 'x';
        }
    }

    showLoading(message = 'Opening document...') {
        if (this.loadingOverlay) {
            const p = this.loadingOverlay.querySelector('p');
            if (p) p.textContent = message;
            this.loadingOverlay.classList.add('active');
        }
    }

    hideLoading() {
        this.loadingOverlay?.classList.remove('active');
    }

    // ========== CLIPBOARD FUNCTIONS ==========

    cutText() {
        const selection = window.getSelection();
        const selectedText = selection.toString();
        if (selectedText) {
            navigator.clipboard.writeText(selectedText).then(() => {
                // Remove selected text
                document.execCommand('delete');
                this.showNotification('Text cut to clipboard');
            });
        } else {
            this.showNotification('Please select text to cut');
        }
    }

    copyText() {
        const selection = window.getSelection();
        const selectedText = selection.toString();
        if (selectedText) {
            navigator.clipboard.writeText(selectedText).then(() => {
                this.showNotification('Text copied to clipboard');
            });
        } else {
            this.showNotification('Please select text to copy');
        }
    }

    async pasteText() {
        try {
            const text = await navigator.clipboard.readText();
            document.execCommand('insertText', false, text);
            this.showNotification('Text pasted');
        } catch (err) {
            this.showNotification('Unable to paste. Please use Ctrl+V');
        }
    }

    // ========== FONT FORMATTING FUNCTIONS ==========

    formatText(command) {
        document.execCommand(command, false, null);
        this.updateFormatButtons();
    }

    updateFormatButtons() {
        // Update button active states based on current selection
        if (this.boldBtn) {
            this.boldBtn.classList.toggle('active', document.queryCommandState('bold'));
        }
        if (this.italicBtn) {
            this.italicBtn.classList.toggle('active', document.queryCommandState('italic'));
        }
        if (this.underlineBtn) {
            this.underlineBtn.classList.toggle('active', document.queryCommandState('underline'));
        }
        if (this.strikeBtn) {
            this.strikeBtn.classList.toggle('active', document.queryCommandState('strikeThrough'));
        }
    }

    changeFont(fontName) {
        document.execCommand('fontName', false, fontName);
    }

    changeFontSize(size) {
        // execCommand fontSize only accepts 1-7, so we use CSS instead
        const selection = window.getSelection();
        if (selection.rangeCount > 0) {
            const range = selection.getRangeAt(0);
            if (!range.collapsed) {
                const span = document.createElement('span');
                span.style.fontSize = size + 'pt';
                range.surroundContents(span);
            }
        }
    }

    changeFontSizeRelative(delta) {
        const currentSize = parseInt(this.fontSizeSelect?.value || 12);
        const newSize = Math.max(8, Math.min(72, currentSize + delta));
        if (this.fontSizeSelect) {
            this.fontSizeSelect.value = newSize;
        }
        this.changeFontSize(newSize);
    }

    highlightText() {
        const colors = ['yellow', 'lime', 'cyan', 'pink', 'orange'];
        const currentColor = this.currentHighlightColor || 0;
        const color = colors[currentColor % colors.length];
        this.currentHighlightColor = currentColor + 1;

        document.execCommand('hiliteColor', false, color);
        this.showNotification('Highlight: ' + color);
    }

    changeFontColor() {
        const colors = ['#000000', '#ff0000', '#0000ff', '#008000', '#800080', '#ff6600'];
        const currentColor = this.currentFontColor || 0;
        const color = colors[currentColor % colors.length];
        this.currentFontColor = currentColor + 1;

        document.execCommand('foreColor', false, color);
        this.showNotification('Font color changed');
    }

    // ========== PARAGRAPH ALIGNMENT FUNCTIONS ==========

    alignText(alignment) {
        switch (alignment) {
            case 'left':
                document.execCommand('justifyLeft', false, null);
                break;
            case 'center':
                document.execCommand('justifyCenter', false, null);
                break;
            case 'right':
                document.execCommand('justifyRight', false, null);
                break;
            case 'justify':
                document.execCommand('justifyFull', false, null);
                break;
        }
        this.updateAlignButtons(alignment);
    }

    updateAlignButtons(active) {
        const buttons = {
            'left': this.alignLeftBtn,
            'center': this.alignCenterBtn,
            'right': this.alignRightBtn,
            'justify': this.alignJustifyBtn
        };

        Object.keys(buttons).forEach(key => {
            if (buttons[key]) {
                buttons[key].classList.toggle('active', key === active);
            }
        });
    }

    async translateWholeDocument(targetLang) {
        if (!this.documentContent.textContent.trim()) {
            this.showNotification('Document is empty.');
            return;
        }

        this.showLoading('Optimizing document for translation...');

        // CRITICAL PERFORMANCE FIX: Collapse chars before translating
        // This converts ~10,000 text nodes back into ~50 text nodes
        this.collapseChars(this.documentContent);

        // 1. Get all text nodes (now much fewer)
        const textNodes = [];
        const walk = document.createTreeWalker(this.documentContent, NodeFilter.SHOW_TEXT, null, false);
        let node;
        while (node = walk.nextNode()) {
            if (node.textContent.trim().length > 0) {
                textNodes.push(node);
            }
        }

        const total = textNodes.length;
        this.showLoading(`Translating ${total} segments...`);

        try {
            for (let i = 0; i < total; i++) {
                const textNode = textNodes[i];
                const originalText = textNode.textContent;

                try {
                    const result = await this.performTranslation(originalText, 'auto', targetLang);
                    textNode.textContent = result;
                } catch (e) {
                    console.error('Translation segment failed', e);
                }

                this.showLoading(`Translating: ${Math.round(((i + 1) / total) * 100)}% (${i + 1}/${total})`);
                await new Promise(r => setTimeout(r, 200)); // Slightly faster delay
            }

            this.showLoading('Restoring reveal logic...');
            this.wrapCharsInElementAsync(this.documentContent, () => {
                this.documentText = this.documentContent.innerText;
                this.totalChars = this.documentContent.querySelectorAll('.char').length;
                this.revealedChars = 0; // Reset reveal for new language
                this.updateWordCount();
                this.updateProgress();
                this.hideLoading();

                this.showNotification('Translation complete! You can use the File menu to save your PDF or Word document.');
            });

        } catch (error) {
            console.error(error);
            this.hideLoading();
            const errorMsg = error.message.includes('429') ? 'Rate limit exceeded. Try again in 1 minute.' : 'Document translation failed.';
            this.showNotification(errorMsg);
        }
    }

    exportDOCX() {
        if (!this.documentContent) return;

        const cleanElement = this.getCleanContent();

        const header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' " +
            "xmlns:w='urn:schemas-microsoft-com:office:word' " +
            "xmlns='http://www.w3.org/TR/REC-html40'>" +
            "<head><meta charset='utf-8'><title>Translated Document</title></head><body>";
        const footer = "</body></html>";
        const sourceHTML = header + cleanElement.innerHTML + footer;

        const source = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(sourceHTML);
        const fileDownload = document.createElement("a");
        document.body.appendChild(fileDownload);
        fileDownload.href = source;
        fileDownload.download = 'translated_document.doc';
        fileDownload.click();
        document.body.removeChild(fileDownload);
        cleanElement.remove();

        this.showNotification('Document saved as Word (.doc)');
    }
    showNotification(message) {
        // Remove existing notification
        const existing = document.querySelector('.app-notification');
        if (existing) existing.remove();

        const notification = document.createElement('div');
        notification.className = 'app-notification';
        notification.textContent = message;
        document.body.appendChild(notification);

        setTimeout(() => {
            notification.classList.add('show');
        }, 10);

        setTimeout(() => {
            notification.classList.remove('show');
            setTimeout(() => notification.remove(), 300);
        }, 2000);
    }

    // ========== DICTIONARY & SAVE ==========

    openDictionaryPanel() {
        this.hideSelectionPopup();
        this.translatePanel?.classList.remove('active');
        this.readPanel?.classList.remove('active');
        this.dictionaryPanel?.classList.add('active');
    }

    async lookupDefinition() {
        const text = this.selectedText;
        if (!text) {
            alert('Please select a word to define');
            return;
        }

        // Clean text (remove punctuation, take first word if multiple)
        const word = text.trim().split(/\s+/)[0].replace(/[.,!?;:()"]/g, '');

        this.openDictionaryPanel();

        if (this.dictionaryContent) {
            this.dictionaryContent.innerHTML = `<div class="dict-placeholder">Searching definition for "<b>${word}</b>"...</div>`;
        }

        try {
            const response = await fetch(`https://api.dictionaryapi.dev/api/v2/entries/en/${word}`);

            if (!response.ok) {
                throw new Error('Word not found');
            }

            const data = await response.json();
            this.displayDefinition(data[0]);
        } catch (error) {
            console.error('Dictionary error:', error);
            if (this.dictionaryContent) {
                this.dictionaryContent.innerHTML = `
                        <div class="dict-placeholder">
                            <div style="font-size: 40px; margin-bottom: 20px;">😕</div>
                            No definition found for "<b>${word}</b>".<br>
                            <small style="display:block; margin-top:10px;">Try selecting a simpler word or check the spelling.</small>
                        </div>`;
            }
        }
    }

    displayDefinition(entry) {
        if (!this.dictionaryContent) return;

        let html = `<div class="dict-word">${entry.word}</div>`;

        if (entry.phonetic) {
            html += `<div class="dict-phonetic">${entry.phonetic}</div>`;
        }

        entry.meanings.forEach(meaning => {
            html += `
                    <div class="dict-meaning">
                        <div class="dict-pos">${meaning.partOfSpeech}</div>
                        ${meaning.definitions.slice(0, 3).map(def => `
                            <div class="dict-def">
                                ${def.definition}
                                ${def.example ? `<div class="dict-example">"${def.example}"</div>` : ''}
                            </div>
                        `).join('')}
                    </div>
                `;
        });

        this.dictionaryContent.innerHTML = html;
    }

    saveDocument() {
        const content = this.documentContent?.innerHTML || '';
        const title = document.querySelector('.doc-title')?.textContent || 'Document';
        const filename = title.replace(/[^a-z0-9]/gi, '_').toLowerCase() + '.html';

        const html = `<!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>${title}</title>
        <style>
            body { font-family: Calibri, sans-serif; line-height: 1.15; padding: 40px; max-width: 816px; margin: 0 auto; white-space: pre-wrap; word-wrap: break-word; }
        </style>
    </head>
    <body>
        ${content}
    </body>
    </html>`;

        const blob = new Blob([html], { type: 'text/html' });
        const url = URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        this.showNotification('Document saved as ' + filename);
    }

    wrapCharsInElement(element) {
        // Legacy synchronous version - kept for small updates
        const textNodes = [];
        const walk = document.createTreeWalker(element, NodeFilter.SHOW_TEXT, null, false);
        let node;
        while (node = walk.nextNode()) textNodes.push(node);

        textNodes.forEach(textNode => {
            const text = textNode.textContent;
            if (text.trim().length === 0 && !text.includes('\n')) return;
            const fragment = document.createDocumentFragment();
            for (let i = 0; i < text.length; i++) {
                const char = text[i];
                const span = document.createElement('span');
                span.className = 'char';
                if (char === '\n') {
                    span.classList.add('newline');
                    fragment.appendChild(span);
                    const br = document.createElement('br');
                    br.classList.add('reveal-br');
                    fragment.appendChild(br);
                } else {
                    span.textContent = char;
                    fragment.appendChild(span);
                }
            }
            textNode.parentNode.replaceChild(fragment, textNode);
        });
    }

    async wrapCharsInElementAsync(element, onComplete) {
        const textNodes = [];
        const walk = document.createTreeWalker(element, NodeFilter.SHOW_TEXT, null, false);
        let node;
        while (node = walk.nextNode()) textNodes.push(node);

        let index = 0;
        const CHUNK_SIZE = 20; // 20 text nodes per frame

        const processBatch = () => {
            const end = Math.min(index + CHUNK_SIZE, textNodes.length);
            for (let i = index; i < end; i++) {
                const textNode = textNodes[index]; // Note: index updates below
                const text = textNode.textContent;
                if (text.trim().length === 0 && !text.includes('\n')) {
                    index++;
                    continue;
                }

                const fragment = document.createDocumentFragment();
                for (let j = 0; j < text.length; j++) {
                    const char = text[j];
                    const span = document.createElement('span');
                    span.className = 'char';
                    if (char === '\n') {
                        span.classList.add('newline');
                        fragment.appendChild(span);
                        fragment.appendChild(document.createElement('br'));
                    } else {
                        span.textContent = char;
                        fragment.appendChild(span);
                    }
                }
                if (textNode.parentNode) {
                    textNode.parentNode.replaceChild(fragment, textNode);
                }
                index++;
            }

            if (index < textNodes.length) {
                requestAnimationFrame(processBatch);
            } else if (onComplete) {
                onComplete();
            }
        };

        if (textNodes.length > 0) processBatch();
        else if (onComplete) onComplete();
    }

    finalizeDocument() {
        if (!this.documentContent) return;

        // 1. Stop reveal logic
        this.revealedChars = this.totalChars;

        // 2. Collapse all spans in the live document for easy editing
        this.collapseChars(this.documentContent);

        // 3. Update UI
        this.updateProgress();
        this.updateWordCount();
        this.showNotification('Document finalized. You can now edit freely!');

        // 4. Focus the editor
        this.documentContent.focus();

        // 5. Hide the finalize button since it's no longer needed
        if (this.finalizeEditBtn) this.finalizeEditBtn.style.display = 'none';

        // 6. Update hint text
        const hint = document.querySelector('.typing-hint');
        if (hint) hint.textContent = '✏️ Editing Mode';
    }

    collapseChars(element) {
        const spans = element.querySelectorAll('.char');
        spans.forEach(span => {
            // Keep the text, discard the span wrapper
            const text = span.textContent || (span.classList.contains('newline') ? '\n' : '');
            const textNode = document.createTextNode(text);
            span.parentNode.replaceChild(textNode, span);
        });

        // Remove only the <br> elements added during wrapping to avoid double-breaks
        const brs = element.querySelectorAll('br.reveal-br');
        brs.forEach(br => br.remove());

        element.normalize(); // Merge adjacent text nodes
    }
}

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    window.wordApp = new WordApp();
});
