// FlexiMapper Client Logic

// State management
let state = {
    source: {
        file: null,
        fileId: null,
        sheets: [],
        selectedSheet: '',
        headerRow: 1,
        columns: []
    },
    target: {
        file: null,
        fileId: null,
        sheets: [],
        selectedSheet: '',
        headerRow: 1,
        columns: []
    },
    mergeMode: 'fill' // 'fill' or 'append'
};

// UI Selectors
const dropzoneSource = document.getElementById('dropzone-source');
const dropzoneTarget = document.getElementById('dropzone-target');
const fileSourceInput = document.getElementById('file-source-input');
const fileTargetInput = document.getElementById('file-target-input');

const infoSource = document.getElementById('info-source');
const infoTarget = document.getElementById('info-target');
const filenameSource = document.getElementById('filename-source');
const filenameTarget = document.getElementById('filename-target');
const filesizeSource = document.getElementById('filesize-source');
const filesizeTarget = document.getElementById('filesize-target');

const btnClearSource = document.getElementById('btn-clear-source');
const btnClearTarget = document.getElementById('btn-clear-target');

const containerSheetsSource = document.getElementById('container-sheets-source');
const containerSheetsTarget = document.getElementById('container-sheets-target');
const selectSheetSource = document.getElementById('select-sheet-source');
const selectSheetTarget = document.getElementById('select-sheet-target');
const inputHeaderSource = document.getElementById('input-header-source');
const inputHeaderTarget = document.getElementById('input-header-target');

const parsingLoader = document.getElementById('parsing-loader');
const sectionMapping = document.getElementById('section-mapping');
const mappingTableBody = document.getElementById('mapping-table-body');

const btnSmartMatch = document.getElementById('btn-smart-match');
const btnProcess = document.getElementById('btn-process');
const btnDownload = document.getElementById('btn-download');
const btnReset = document.getElementById('btn-reset');

const processingLoader = document.getElementById('processing-loader');
const sectionSuccess = document.getElementById('section-success');

const labelModeFill = document.getElementById('label-mode-fill');
const labelModeAppend = document.getElementById('label-mode-append');
const radioFillDot = document.getElementById('radio-fill-dot');
const radioAppendDot = document.getElementById('radio-append-dot');

let mergedBlobUrl = null;
let downloadFilename = "FlexiMapper_Output.xlsx";

// Initialize Lucide Icons
document.addEventListener('DOMContentLoaded', () => {
    lucide.createIcons();
    setupDropzone(dropzoneSource, fileSourceInput, 'source');
    setupDropzone(dropzoneTarget, fileTargetInput, 'target');
    setupMergeModeToggle();
    setupEventListeners();
});

// Setup drag and drop events
function setupDropzone(dropzone, input, type) {
    // Click to select
    dropzone.addEventListener('click', () => input.click());

    // Input changed
    input.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileSelect(e.target.files[0], type);
        }
    });

    // Drag events
    ['dragenter', 'dragover'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropzone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.remove('dragover');
        }, false);
    });

    dropzone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files.length > 0) {
            handleFileSelect(files[0], type);
        }
    }, false);
}

// Handle selected file
async function handleFileSelect(file, type) {
    // Validate file extension
    const extension = file.name.split('.').pop().toLowerCase();
    if (extension !== 'xlsx' && extension !== 'xlsm') {
        alert('Invalid file format. Please upload an Excel sheet (.xlsx, .xlsm).');
        return;
    }

    state[type].file = file;

    // Update UI
    const dropzone = type === 'source' ? dropzoneSource : dropzoneTarget;
    const infoContainer = type === 'source' ? infoSource : infoTarget;
    const filenameLabel = type === 'source' ? filenameSource : filenameTarget;
    const filesizeLabel = type === 'source' ? filesizeSource : filesizeTarget;

    dropzone.classList.add('hidden');
    infoContainer.classList.remove('hidden');
    filenameLabel.textContent = file.name;
    filesizeLabel.textContent = formatBytes(file.size);

    // If both files are loaded, trigger automatic upload
    if (state.source.file && state.target.file) {
        await uploadAndScanFiles();
    }
}

// Format bytes to readable string
function formatBytes(bytes, decimals = 2) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

// Clear files logic
function setupEventListeners() {
    btnClearSource.addEventListener('click', (e) => {
        e.stopPropagation();
        resetFileType('source');
    });

    btnClearTarget.addEventListener('click', (e) => {
        e.stopPropagation();
        resetFileType('target');
    });

    selectSheetSource.addEventListener('change', async (e) => {
        state.source.selectedSheet = e.target.value;
        await fetchColumns('source');
    });

    selectSheetTarget.addEventListener('change', async (e) => {
        state.target.selectedSheet = e.target.value;
        await fetchColumns('target');
    });

    inputHeaderSource.addEventListener('change', async (e) => {
        let val = parseInt(e.target.value) || 1;
        if (val < 1) val = 1;
        e.target.value = val;
        state.source.headerRow = val;
        await fetchColumns('source');
    });

    inputHeaderTarget.addEventListener('change', async (e) => {
        let val = parseInt(e.target.value) || 1;
        if (val < 1) val = 1;
        e.target.value = val;
        state.target.headerRow = val;
        await fetchColumns('target');
    });

    btnSmartMatch.addEventListener('click', runSmartMatch);

    btnProcess.addEventListener('click', processAndMerge);

    btnReset.addEventListener('click', resetApp);
}

function resetFileType(type) {
    state[type].file = null;
    state[type].fileId = null;
    state[type].sheets = [];
    state[type].selectedSheet = '';
    state[type].headerRow = 1;
    state[type].columns = [];

    const dropzone = type === 'source' ? dropzoneSource : dropzoneTarget;
    const infoContainer = type === 'source' ? infoSource : infoTarget;
    const selectContainer = type === 'source' ? containerSheetsSource : containerSheetsTarget;
    const input = type === 'source' ? fileSourceInput : fileTargetInput;
    const headerInput = type === 'source' ? inputHeaderSource : inputHeaderTarget;

    input.value = '';
    if (headerInput) {
        headerInput.value = 1;
    }
    dropzone.classList.remove('hidden');
    infoContainer.classList.add('hidden');
    selectContainer.classList.add('hidden');

    hideMappingAndResults();
}

function hideMappingAndResults() {
    sectionMapping.classList.add('hidden');
    sectionSuccess.classList.add('hidden');
}

// Upload both sheets to FastAPI
async function uploadAndScanFiles() {
    parsingLoader.classList.remove('hidden');
    hideMappingAndResults();

    const formData = new FormData();
    formData.append('file_source', state.source.file);
    formData.append('file_target', state.target.file);

    try {
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            const err = await response.json();
            throw new Error(err.detail || 'Failed to scan files.');
        }

        const data = await response.json();
        
        state.source.fileId = data.source_id;
        state.target.fileId = data.target_id;
        state.source.sheets = data.source_sheets;
        state.target.sheets = data.target_sheets;

        // Render sheets dropdown
        populateSheetsDropdown('source', selectSheetSource, containerSheetsSource);
        populateSheetsDropdown('target', selectSheetTarget, containerSheetsTarget);

        // Auto fetch columns for first sheets
        state.source.selectedSheet = data.source_sheets[0];
        state.target.selectedSheet = data.target_sheets[0];
        
        await Promise.all([
            fetchColumns('source'),
            fetchColumns('target')
        ]);

    } catch (error) {
        alert('Upload failed: ' + error.message);
        resetApp();
    } finally {
        parsingLoader.classList.add('hidden');
    }
}

function populateSheetsDropdown(type, selectEl, containerEl) {
    selectEl.innerHTML = '';
    const sheets = state[type].sheets;
    
    sheets.forEach(sheet => {
        const opt = document.createElement('option');
        opt.value = sheet;
        opt.textContent = sheet;
        selectEl.appendChild(opt);
    });

    containerEl.classList.remove('hidden');
}

// Fetch columns for a specific file sheet
async function fetchColumns(type) {
    const fileId = state[type].fileId;
    const sheetName = state[type].selectedSheet;

    if (!fileId || !sheetName) return;

    try {
        const response = await fetch('/columns', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ 
                file_id: fileId, 
                sheet_name: sheetName,
                header_row: state[type].headerRow
            })
        });

        if (!response.ok) {
            const err = await response.json();
            throw new Error(err.detail || 'Failed to read columns.');
        }

        const data = await response.json();
        state[type].columns = data.columns;

        // If columns for both source and target are retrieved, render mapping table
        if (state.source.columns.length > 0 && state.target.columns.length > 0) {
            renderMappingTable();
        }

    } catch (error) {
        alert(`Error loading ${type} columns: ` + error.message);
    }
}

// Update status badge and row background styling for a mapping row dynamically
function updateRowStatus(selectEl) {
    const tr = selectEl.closest('tr');
    if (!tr) return;
    const statusCell = tr.querySelector('.status-cell');
    if (!statusCell) return;

    const sourceCol = selectEl.dataset.sourceColumn;
    const targetCol = selectEl.value;

    // Reset background styles
    tr.classList.remove('bg-emerald-500/5', 'bg-brand-500/5');

    if (!targetCol) {
        statusCell.innerHTML = '<i data-lucide="arrow-right" class="w-4 h-4 mx-auto text-gray-600"></i>';
    } else if (targetCol.trim().toLowerCase() === sourceCol.trim().toLowerCase()) {
        statusCell.innerHTML = `
            <span class="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-emerald-500/10 border border-emerald-500/20 text-emerald-400 text-xs font-semibold">
                <i data-lucide="check-circle" class="w-3.5 h-3.5"></i> Auto-Matched
            </span>
        `;
        tr.classList.add('bg-emerald-500/5');
    } else {
        statusCell.innerHTML = `
            <span class="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-indigo-500/10 border border-indigo-500/20 text-indigo-400 text-xs font-semibold">
                <i data-lucide="link" class="w-3.5 h-3.5"></i> Mapped
            </span>
        `;
        tr.classList.add('bg-brand-500/5');
    }
}

// Render dynamic mapping table
function renderMappingTable() {
    mappingTableBody.innerHTML = '';
    const targetCols = state.target.columns;
    const sourceCols = state.source.columns;

    sourceCols.forEach((sourceCol, index) => {
        const tr = document.createElement('tr');
        tr.className = 'border-b border-gray-800/40 hover:bg-gray-900/10 transition-colors';

        // Source name cell (LHS)
        const tdSource = document.createElement('td');
        tdSource.className = 'px-6 py-4 text-sm font-semibold text-gray-200';
        tdSource.textContent = sourceCol;

        // Icon indicator cell
        const tdArrow = document.createElement('td');
        tdArrow.className = 'px-6 py-4 text-center text-gray-500 status-cell';
        tdArrow.innerHTML = '<i data-lucide="arrow-right" class="w-4 h-4 mx-auto"></i>';

        // Target Dropdown Selector cell (RHS)
        const tdSelect = document.createElement('td');
        tdSelect.className = 'px-6 py-4';
        
        const selectContainer = document.createElement('div');
        selectContainer.className = 'relative max-w-xs';

        const select = document.createElement('select');
        select.className = 'mapping-select w-full bg-gray-900/80 border border-gray-800 text-gray-300 rounded-xl px-4 py-2 text-xs focus:outline-none focus:border-brand-500 appearance-none';
        select.dataset.sourceColumn = sourceCol;

        // Default 'Don't Map' option
        const defOpt = document.createElement('option');
        defOpt.value = '';
        defOpt.textContent = '— Skip / Don\'t Map —';
        select.appendChild(defOpt);

        // Check for exact matching target column case-insensitively
        let matchedTarget = '';
        const cleanSource = sourceCol.trim().toLowerCase();
        for (let tCol of targetCols) {
            if (tCol.trim().toLowerCase() === cleanSource) {
                matchedTarget = tCol;
                break;
            }
        }

        // Add target options
        targetCols.forEach(targetCol => {
            const opt = document.createElement('option');
            opt.value = targetCol;
            opt.textContent = targetCol;
            if (targetCol === matchedTarget) {
                opt.selected = true;
            }
            select.appendChild(opt);
        });

        const chevron = document.createElement('div');
        chevron.className = 'absolute inset-y-0 right-0 flex items-center pr-3 pointer-events-none text-gray-500';
        chevron.innerHTML = '<i data-lucide="chevron-down" class="w-3.5 h-3.5"></i>';

        selectContainer.appendChild(select);
        selectContainer.appendChild(chevron);
        tdSelect.appendChild(selectContainer);

        tr.appendChild(tdSource);
        tr.appendChild(tdArrow);
        tr.appendChild(tdSelect);
        mappingTableBody.appendChild(tr);

        // Update row status visual representation immediately
        updateRowStatus(select);

        // Dynamic change listener
        select.addEventListener('change', () => {
            updateRowStatus(select);
            lucide.createIcons();
        });
    });

    // Reinitialize icons in new elements
    lucide.createIcons();

    // Show Mapping Section
    sectionMapping.classList.remove('hidden');
    // Scroll into view gently
    sectionMapping.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// Setup merge strategy radio selection styling
function setupMergeModeToggle() {
    labelModeFill.addEventListener('click', () => {
        state.mergeMode = 'fill';
        
        // Update styling
        labelModeFill.className = 'relative flex p-4 rounded-xl border border-brand-500 bg-brand-500/5 cursor-pointer select-none focus:outline-none transition-all duration-300 hover:bg-brand-500/10';
        labelModeAppend.className = 'relative flex p-4 rounded-xl border border-gray-800 bg-gray-900/20 cursor-pointer select-none focus:outline-none transition-all duration-300 hover:bg-gray-900/40';
        
        radioFillDot.innerHTML = '<div class="w-2 h-2 rounded-full bg-brand-500"></div>';
        radioFillDot.className = 'w-4.5 h-4.5 rounded-full border-2 border-brand-500 flex items-center justify-center';
        
        radioAppendDot.innerHTML = '';
        radioAppendDot.className = 'w-4.5 h-4.5 rounded-full border-2 border-gray-600 flex items-center justify-center';
    });

    labelModeAppend.addEventListener('click', () => {
        state.mergeMode = 'append';
        
        // Update styling
        labelModeFill.className = 'relative flex p-4 rounded-xl border border-gray-800 bg-gray-900/20 cursor-pointer select-none focus:outline-none transition-all duration-300 hover:bg-gray-900/40';
        labelModeAppend.className = 'relative flex p-4 rounded-xl border border-brand-500 bg-brand-500/5 cursor-pointer select-none focus:outline-none transition-all duration-300 hover:bg-brand-500/10';
        
        radioFillDot.innerHTML = '';
        radioFillDot.className = 'w-4.5 h-4.5 rounded-full border-2 border-gray-600 flex items-center justify-center';
        
        radioAppendDot.innerHTML = '<div class="w-2 h-2 rounded-full bg-brand-500"></div>';
        radioAppendDot.className = 'w-4.5 h-4.5 rounded-full border-2 border-brand-500 flex items-center justify-center';
    });
}

// String similarity calculations for Smart Matching
function cleanString(str) {
    return str.toLowerCase().replace(/[^a-z0-9]/g, '');
}

// Levenshtein similarity score (0.0 to 1.0)
function getSimilarityScore(s1, s2) {
    const cleanS1 = cleanString(s1);
    const cleanS2 = cleanString(s2);

    if (cleanS1 === cleanS2) return 1.0; // Perfect match

    const len1 = cleanS1.length;
    const len2 = cleanS2.length;
    const maxLen = Math.max(len1, len2);
    if (maxLen === 0) return 1.0;

    const matrix = [];
    for (let i = 0; i <= len1; i++) {
        matrix[i] = [i];
    }
    for (let j = 0; j <= len2; j++) {
        matrix[0][j] = j;
    }

    for (let i = 1; i <= len1; i++) {
        for (let j = 1; j <= len2; j++) {
            if (cleanS1.charAt(i - 1) === cleanS2.charAt(j - 1)) {
                matrix[i][j] = matrix[i - 1][j - 1];
            } else {
                matrix[i][j] = Math.min(
                    matrix[i - 1][j - 1] + 1, // substitution
                    Math.min(
                        matrix[i][j - 1] + 1, // insertion
                        matrix[i - 1][j] + 1  // deletion
                    )
                );
            }
        }
    }

    const distance = matrix[len1][len2];
    return (maxLen - distance) / parseFloat(maxLen);
}

// Smart Match click handler
function runSmartMatch() {
    const selects = document.querySelectorAll('.mapping-select');
    const targetCols = state.target.columns;
    let matchCount = 0;

    selects.forEach(select => {
        // Skip columns that are already mapped (either auto-mapped on load or manually mapped)
        if (select.value !== '') {
            return;
        }

        const sourceCol = select.dataset.sourceColumn;
        let bestMatch = '';
        let highestScore = 0;

        targetCols.forEach(targetCol => {
            const score = getSimilarityScore(sourceCol, targetCol);
            
            // Substring checks to increase weight
            const cleanS = cleanString(sourceCol);
            const cleanT = cleanString(targetCol);
            let multiplier = 1.0;
            if (cleanS.includes(cleanT) || cleanT.includes(cleanS)) {
                multiplier = 1.15; // Give boost if one string completely wraps the other
            }

            const finalScore = score * multiplier;

            if (finalScore > highestScore) {
                highestScore = finalScore;
                bestMatch = targetCol;
            }
        });

        // Set threshold (e.g. 0.6 or 60% similarity after adjustment)
        if (highestScore >= 0.6) {
            select.value = bestMatch;
            matchCount++;
            updateRowStatus(select);
        }
    });

    // Reinitialize icons in newly updated elements
    lucide.createIcons();

    if (matchCount > 0) {
        // Simple micro feedback toast
        const feedback = document.createElement('div');
        feedback.className = 'fixed bottom-6 right-6 bg-brand-600 text-white text-xs px-4 py-2.5 rounded-xl shadow-lg border border-brand-500 z-50 animate-slideUp';
        feedback.textContent = `Auto-mapped ${matchCount} columns based on name similarity!`;
        document.body.appendChild(feedback);
        setTimeout(() => {
            feedback.classList.add('hidden');
            feedback.remove();
        }, 3000);
    } else {
        alert('No new highly similar columns found between the sheets. Please select mappings manually.');
    }
}

// Post mapping configuration and download merged file
async function processAndMerge() {
    const selects = document.querySelectorAll('.mapping-select');
    const mappings = {};
    let mappedCount = 0;

    selects.forEach(select => {
        const sourceCol = select.dataset.sourceColumn;
        const targetCol = select.value;
        if (targetCol) {
            mappings[sourceCol] = targetCol;
            mappedCount++;
        }
    });

    if (mappedCount === 0) {
        alert('Please map at least one column before processing.');
        return;
    }

    // Hide mapping controls and show processing loader
    sectionMapping.classList.add('hidden');
    processingLoader.classList.remove('hidden');
    processingLoader.scrollIntoView({ behavior: 'smooth', block: 'center' });

    const payload = {
        source_id: state.source.fileId,
        target_id: state.target.fileId,
        source_sheet: state.source.selectedSheet,
        target_sheet: state.target.selectedSheet,
        mappings: mappings,
        merge_mode: state.mergeMode,
        source_header_row: state.source.headerRow,
        target_header_row: state.target.headerRow
    };

    try {
        const response = await fetch('/merge', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            const err = await response.json();
            throw new Error(err.detail || 'An error occurred during sheet merging.');
        }

        // Get file name from response header if available
        const contentDisposition = response.headers.get('content-disposition');
        if (contentDisposition) {
            const matches = /filename="([^"]+)"/.exec(contentDisposition);
            if (matches && matches[1]) {
                downloadFilename = matches[1];
            }
        } else {
            downloadFilename = `FlexiMapper_${state.target.file.name}`;
        }

        // Read binary Excel blob response
        const blob = await response.blob();
        if (mergedBlobUrl) {
            window.URL.revokeObjectURL(mergedBlobUrl);
        }
        mergedBlobUrl = window.URL.createObjectURL(blob);

        // Show Success panel
        processingLoader.classList.add('hidden');
        sectionSuccess.classList.remove('hidden');
        sectionSuccess.scrollIntoView({ behavior: 'smooth', block: 'center' });

        // Auto trigger download
        triggerDownload();

    } catch (error) {
        alert('Processing failed: ' + error.message);
        // Show mapping controls back to allow retry
        processingLoader.classList.add('hidden');
        sectionMapping.classList.remove('hidden');
    }
}

// Download triggering
function triggerDownload() {
    if (!mergedBlobUrl) return;
    const a = document.createElement('a');
    a.href = mergedBlobUrl;
    a.download = downloadFilename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

// Bind download button
btnDownload.addEventListener('click', triggerDownload);

// Reset application to starting state
function resetApp() {
    if (mergedBlobUrl) {
        window.URL.revokeObjectURL(mergedBlobUrl);
        mergedBlobUrl = null;
    }
    
    resetFileType('source');
    resetFileType('target');

    hideMappingAndResults();
    
    processingLoader.classList.add('hidden');
    
    // Scroll to top
    window.scrollTo({ top: 0, behavior: 'smooth' });
}
